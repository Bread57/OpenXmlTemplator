namespace OpenXmlTemplator.Docx
{
    using OpenXmlTemplator.Docx.Models;
    using System.Collections.ObjectModel;
    using System.IO.Compression;
    using System.Xml.Linq;

    /// <summary>
    /// Варианты создания документов
    /// </summary>
    public class CreatingDocument
    {
        private const string _docxExtension = ".docx";

        /// <summary>
        /// Создание объеденных документов, т.е. много документов в одном docx файле
        /// </summary>
        /// <returns></returns>
        public byte[] MergedDocuments(DocxDocumentModel docxTemplatorModels)
        {
            using (MemoryStream docxStream = new())//Поток для итогового документа
            {
                docxTemplatorModels.InStream.CopyTo(docxStream);

                //Docx файл, в котором создано несколько документов документы
                using (ZipArchive docx = new(stream: docxStream, mode: ZipArchiveMode.Update))
                {
                    ZipArchiveEntry documentXmlEntry = docx.GetEntry(@"word/document.xml") ?? throw new FileNotFoundException("В архиве отсутствует XML файл с содержимым документа.");

                    //Получем поток document.xml, для загрузки Xdocument
                    using (Stream documentXmlStream = documentXmlEntry.Open())
                    {
                        XDocument document = XDocument.Load(documentXmlStream, LoadOptions.None);

                        if (document.Root is null)
                        {
                            throw new FileNotFoundException("В XML файле отсутствует корневой элемент <document>.");
                        }

                        XElement body = document.Root!.Element(DocxXNames.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                        int count = docxTemplatorModels.Documents.Count();

                        foreach ((string documentName, KeyWordsHandlerModel keyWords) in docxTemplatorModels.Documents)
                        {
                            //Что бы шаблон всегда был по рукой - создаем копию newBody
                            XElement newBody = new(body);

                            ICollection<XElement> toRemove = new Collection<XElement>();

                            //Рекурсивный поиск по document.xml
                            foreach (var element in newBody.Elements())
                            {
                                SearchAndReplace.RecursiveSearch(
                                    element: element,
                                    keyWordsHandler: keyWords,
                                    search: new SearchKeyWord(startingKeys: docxTemplatorModels.StartingKeys, endingKeys: docxTemplatorModels.EndingKeys, keyWordParamsSeparator: docxTemplatorModels.KeyWordParamsSeparator),
                                    toRemove: toRemove
                                    );
                            }

                            //Удаляем лишние элементы
                            foreach (XElement element in toRemove)
                            {
                                element.Remove();
                            }

                            body.AddBeforeSelf(newBody);

                            //Не добавляем разрыв за последним документов
                            if (--count > 0)
                            {
                                newBody.AddAfterSelf(DocxCommonElements.PageBreak);
                            }
                        }

                        //после записи всех newBody - удаляем шаблонный body из файла
                        body.Remove();

                        documentXmlStream.Seek(0, SeekOrigin.Begin);//Что бы документ заменялся, а не просто сохранился в конце имеющегося
                        document.Save(documentXmlStream);
                    }
                }
                return docxStream.ToArray();
            }
        }

        /// <summary>
        /// Создание раздельных документов, т.е. на каждого студента свои docx файл
        /// </summary>
        /// <returns></returns>
        public byte[] SeparateDocuments(DocxDocumentModel docxTemplatorModels)
        {
            using (MemoryStream zipStream = new())//Поток для итогового документа
            {
                using (ZipArchive zip = new(stream: zipStream, mode: ZipArchiveMode.Update))//собираем все docx файлы в архив
                {
                    foreach ((string documentName, KeyWordsHandlerModel keyWords) in docxTemplatorModels.Documents)
                    {
                        using (MemoryStream docxStream = new())//поток для записи docx файла в rar архив
                        {
                            //Копируем шаблон в поток
                            docxTemplatorModels.InStream.CopyTo(docxStream);
                            docxTemplatorModels.InStream.Seek(0, SeekOrigin.Begin);//Важно!!! После копирования, указатель будет в конце поток, переносим в начало

                            //создаем изменяемый архив на основе шаблона
                            using (ZipArchive docx = new(stream: docxStream, mode: ZipArchiveMode.Update, leaveOpen: true))
                            {
                                ZipArchiveEntry documentXmlEntry = docx.GetEntry(@"word/document.xml") ?? throw new FileNotFoundException("В архиве отсутствует XML файл с содержимым документа.");

                                //Получем поток document.xml, для загрузки Xdocument
                                using (Stream documentXmlStream = documentXmlEntry.Open())
                                {
                                    XDocument document = XDocument.Load(documentXmlStream, LoadOptions.None);

                                    if (document.Root is null)
                                    {
                                        throw new FileNotFoundException("В XML файле отсутствует корневой элемент <document>.");
                                    }

                                    XElement body = document.Root!.Element(DocxXNames.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                                    ICollection<XElement> toRemove = new Collection<XElement>();

                                    //Рекурсивный поиск по document.xml
                                    foreach (var element in body.Elements())
                                    {
                                        SearchAndReplace.RecursiveSearch(
                                            element: element,
                                            keyWordsHandler: keyWords,
                                            search: new SearchKeyWord(startingKeys: docxTemplatorModels.StartingKeys, endingKeys: docxTemplatorModels.EndingKeys, keyWordParamsSeparator: docxTemplatorModels.KeyWordParamsSeparator),
                                            toRemove: toRemove
                                            );
                                    }

                                    //Удаляем лишние элементы
                                    foreach (XElement element in toRemove)
                                    {
                                        element.Remove();
                                    }

                                    documentXmlStream.Seek(0, SeekOrigin.Begin);//Что бы документ заменялся, а не просто сохранился в конце имеющегося
                                    document.Save(documentXmlStream);
                                }
                            }

                            //Создаем в rar новое вхождение(файла)
                            ZipArchiveEntry zip_entry =
                                zip.CreateEntry(
                                    $"{documentName}{_docxExtension}", CompressionLevel.Optimal);


                            //Переносим указатель потока
                            docxStream.Seek(0, SeekOrigin.Begin);

                            //Копируем поток в новый фалй в архиве
                            docxStream.CopyTo(zip_entry.Open());
                        }
                    }
                }
                return zipStream.ToArray();//Важно вернуть после закрытия ZipArchive, т.к. только после этого, архив запишетсяя в поток zipStream
            }
        }
    }
}
