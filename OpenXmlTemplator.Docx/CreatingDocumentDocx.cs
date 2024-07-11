namespace OpenXmlTemplator.Docx
{
    using OpenXmlTemplator.Docx.Auxiliary;
    using OpenXmlTemplator.Docx.Models.InnerModels;
    using OpenXmlTemplator.Docx.Models.OuterModels;
    using System.Collections.ObjectModel;
    using System.IO.Compression;
    using System.Xml.Linq;

    /// <summary>
    /// Варианты создания документов
    /// </summary>
    public static class CreatingDocumentDocx
    {
        private const string _docxExtension = ".docx";

        /// <summary>
        /// Создание объеденных документов, т.е. много документов в одном docx файле
        /// </summary>
        /// <returns></returns>
        public static byte[] MergedDocuments(DocumentModelDocx docxTemplatorModels)
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

                        XElement body = document.Root!.Element(XNamesDocx.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                        int count = docxTemplatorModels.Documents.Count();

                        foreach ((string documentName, KeyWordsHandlerModelDocx keyWords) in docxTemplatorModels.Documents)
                        {
                            //Что бы шаблон всегда был по рукой - создаем копию newBody
                            XElement newBody = new(body);

                            ICollection<XElement> toDelayedRemove = new Collection<XElement>();

                            //Рекурсивный поиск по document.xml
                            foreach (var element in newBody.Elements())
                            {
                                SearchAndReplaceDocx.RecursiveSearch(
                                    element: element,
                                    keyWordsHandler: keyWords,
                                    search: new SearchingKeyWordModelDocx(docxTemplatorModels.SearchModel),
                                    builtInKeyWordsHandlers: docxTemplatorModels.BuiltInKeyWordsHandlers,
                                    toDelayedRemove: toDelayedRemove
                                    );
                            }

                            //Удаляем лишние элементы
                            foreach (XElement element in toDelayedRemove)
                            {
                                element.Remove();
                            }

                            body.AddBeforeSelf(newBody);

                            //Не добавляем разрыв за последним документов
                            if (--count > 0)
                            {
                                newBody.AddAfterSelf(CommonElementsDocx.PageBreak);
                            }
                        }

                        //после записи всех newBody - удаляем шаблонный body из файла
                        body.Remove();

                        DocumentSave(documentXmlStream: documentXmlStream, document: document);
                    }
                }
                return docxStream.ToArray();
            }
        }

        /// <summary>
        /// Создание раздельных документов, т.е. на каждого студента свои docx файл
        /// </summary>
        /// <returns></returns>
        public static byte[] SeparateDocuments(DocumentModelDocx docxTemplatorModels)
        {
            using (MemoryStream zipStream = new())//Поток для итогового документа
            {
                using (ZipArchive zip = new(stream: zipStream, mode: ZipArchiveMode.Update))//собираем все docx файлы в архив
                {
                    foreach ((string documentName, KeyWordsHandlerModelDocx keyWords) in docxTemplatorModels.Documents)
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

                                    XElement body = document.Root!.Element(XNamesDocx.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                                    ICollection<XElement> toDelayedRemove = new Collection<XElement>();

                                    //Рекурсивный поиск по document.xml
                                    foreach (var element in body.Elements())
                                    {
                                        SearchAndReplaceDocx.RecursiveSearch(
                                            element: element,
                                            keyWordsHandler: keyWords,
                                            search: new SearchingKeyWordModelDocx(docxTemplatorModels.SearchModel),
                                            builtInKeyWordsHandlers: docxTemplatorModels.BuiltInKeyWordsHandlers,
                                            toDelayedRemove: toDelayedRemove
                                            );
                                    }

                                    //Удаляем лишние элементы
                                    foreach (XElement element in toDelayedRemove)
                                    {
                                        element.Remove();
                                    }

                                    DocumentSave(documentXmlStream: documentXmlStream, document: document);
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

        /// <summary>
        /// Сохранение документа в поток
        /// </summary>
        /// <param name="documentXmlStream">Поток</param>
        /// <param name="document">Документ</param>
        private static void DocumentSave(Stream documentXmlStream, XDocument document)
        {
            long oldStreamLength = documentXmlStream.Length;//Размера шаблона

            documentXmlStream.Seek(0, SeekOrigin.Begin);//Что бы документ заменялся, а не просто сохранился в конце имеющегося
            document.Save(documentXmlStream);

            //Если позиция в потоке после сохранения файла меньше изначальной длины потока - поток нужно урезать, что бы в итоговом файле не попали элементы шаблона
            if (oldStreamLength > documentXmlStream.Position)
            {
                documentXmlStream.SetLength(documentXmlStream.Position);
            }
        }
    }
}
