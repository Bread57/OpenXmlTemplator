using System.IO.Compression;
using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
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
        public static async Task<byte[]> MergedDocuments(DocumentModelDocx docxTemplatorModels, CancellationToken token = default)
        {
            await using (MemoryStream docxStream = new())//Поток для итогового документа
            {
                await docxTemplatorModels.InStream.CopyToAsync(docxStream, token);

                //Docx файл, в котором создано несколько документов документы
                await using (ZipArchive docx = new(stream: docxStream, mode: ZipArchiveMode.Update))
                {
                    ZipArchiveEntry documentXmlEntry = docx.GetEntry(@"word/document.xml") ?? throw new FileNotFoundException("В архиве отсутствует XML файл с содержимым документа.");

                    //Получем поток document.xml, для загрузки Xdocument
                    await using (Stream documentXmlStream = await documentXmlEntry.OpenAsync(token))
                    {
                        XDocument document = await XDocument.LoadAsync(documentXmlStream, LoadOptions.None, token);

                        if (document.Root is null)
                        {
                            throw new FileNotFoundException("В XML файле отсутствует корневой элемент <document>.");
                        }

                        XElement body = document.Root!.Element(XNamesDocx.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                        int count = docxTemplatorModels.Documents.Count();

                        foreach ((string documentName, KeyWordsHandlerModelDocx keyWords) in docxTemplatorModels.Documents)
                        {
                            token.ThrowIfCancellationRequested();

                            //Что бы шаблон всегда был по рукой - создаем копию newBody
                            XElement newBody = new(body);

                            await TemplateReplace(newBody, keyWords, docxTemplatorModels, docx, token);

                            body.AddBeforeSelf(newBody);

                            //Не добавляем разрыв за последним документов
                            if (--count > 0)
                            {
                                newBody.AddAfterSelf(CommonElementsDocx.PageBreak);
                            }
                        }

                        //после записи всех newBody - удаляем шаблонный body из файла
                        body.Remove();

                        await DocumentSave(documentXmlStream: documentXmlStream, document: document, token);
                    }
                }

                return docxStream.ToArray();
            }
        }

        /// <summary>
        /// Создание раздельных документов, т.е. на каждого студента свой docx файл
        /// </summary>
        /// <returns></returns>
        public static async Task<byte[]> SeparateDocuments(DocumentModelDocx docxTemplatorModels, CancellationToken token = default)
        {
            await using (MemoryStream zipStream = new())//Поток для итогового документа
            {
                await using (ZipArchive zip = new(stream: zipStream, mode: ZipArchiveMode.Update))//собираем все docx файлы в архив
                {
                    foreach ((string documentName, KeyWordsHandlerModelDocx keyWords) in docxTemplatorModels.Documents)
                    {
                        await using (MemoryStream docxStream = new())//поток для записи docx файла в rar архив
                        {
                            //Копируем шаблон в поток
                            await docxTemplatorModels.InStream.CopyToAsync(docxStream, token);
                            docxTemplatorModels.InStream.Seek(0, SeekOrigin.Begin);//Важно!!! После копирования, указатель будет в конце поток, переносим в начало

                            //создаем изменяемый архив на основе шаблона
                            await using (ZipArchive docx = new(stream: docxStream, mode: ZipArchiveMode.Update, leaveOpen: true))
                            {
                                ZipArchiveEntry documentXmlEntry = docx.GetEntry(@"word/document.xml") ?? throw new FileNotFoundException("В архиве отсутствует XML файл с содержимым документа.");

                                //Получем поток document.xml, для загрузки Xdocument
                                await using (Stream documentXmlStream = await documentXmlEntry.OpenAsync(token))
                                {
                                    XDocument document = await XDocument.LoadAsync(documentXmlStream, LoadOptions.None, token);

                                    if (document.Root is null)
                                    {
                                        throw new FileNotFoundException("В XML файле отсутствует корневой элемент <document>.");
                                    }

                                    XElement body = document.Root!.Element(XNamesDocx.Body) ?? throw new FileNotFoundException("В XML файле отсутствует элемент <body>.");

                                    await TemplateReplace(body, keyWords, docxTemplatorModels, docx, token);

                                    await DocumentSave(documentXmlStream: documentXmlStream, document: document, token);
                                }
                            }

                            token.ThrowIfCancellationRequested();

                            //Создаем в rar новое вхождение(файла)
                            ZipArchiveEntry zip_entry =
                                zip.CreateEntry(
                                    $"{documentName}{_docxExtension}", CompressionLevel.Optimal);

                            //Переносим указатель потока
                            docxStream.Seek(0, SeekOrigin.Begin);

                            //Копируем поток в новый файл в архиве
                            await docxStream.CopyToAsync(await zip_entry.OpenAsync(token), token);
                        }
                    }
                }
                return zipStream.ToArray();//Важно вернуть после закрытия ZipArchive, т.к. только после этого, архив запишетсяя в поток zipStream
            }
        }

        /// <summary>
        /// Метод редактирования xml body документа
        /// </summary>
        /// <param name="body">тело документа</param>
        /// <param name="keyWords">обработчики ключевых слов</param>
        /// <param name="docxTemplatorModels"> Модель для формирования docx документа по шаблону</param>
        private static async Task TemplateReplace(
            XElement body,
            KeyWordsHandlerModelDocx keyWords,
            DocumentModelDocx docxTemplatorModels,
            ZipArchive docx,
            CancellationToken token
            )
        {
            await AddFiles(docx, keyWords, token);

            ICollection<XElement> toDelayedRemove = [];

            //Рекурсивный поиск по document.xml
            foreach (var element in body.Elements())
            {
                token.ThrowIfCancellationRequested();

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

            await RemoveUnnecessaryRelationShipFromRelsFile(docx, keyWords, token);
        }

        #region Работа с файлами внутри документа
        /// <summary>
        /// Папка Word
        /// </summary>
        private const string WordPath = "word";
        /// <summary>
        /// Полный путь до document.xml.rels
        /// </summary>
        /// <value></value>
        private const string _RelsPath = $"{WordPath}/_rels/document.xml.rels";

        /// <summary>
        /// Полный путь до [Content_Types].xml
        /// </summary>
        private const string Content_TypesPath = $"[Content_Types].xml";

        /// <summary>
        /// Часть Id всех файлов, далее к нему будет просто номер добавляться
        /// </summary>
        private const string FileIdPart = "_file";

        /// <summary>
        /// Удаляем неиспользуемые ссылки-файлов из document.xml.rels
        /// </summary>
        /// <param name="docx">Архив word</param>
        /// <param name="keyWords">Обработчик ключевых слов</param>
        /// <param name="token">Токен</param>
        /// <returns></returns>
        private static async Task RemoveUnnecessaryRelationShipFromRelsFile(ZipArchive docx, KeyWordsHandlerModelDocx keyWords, CancellationToken token)
        {
            if (keyWords.KeyWordsToFileReplace.Count == 0)
            {
                return;
            }

            //Выгружаем вхождение document.xml.rels
            ZipArchiveEntry relsEntry = docx.GetEntry(_RelsPath) ?? throw new FileNotFoundException($"Не найден файл {_RelsPath}");

            await using (Stream relsStream = await relsEntry.OpenAsync(token))
            {
                XDocument document = await XDocument.LoadAsync(relsStream, LoadOptions.None, token);

                //Берем корневой элемент  <Relationships>
                XElement relationships = document!.Root ?? throw new FileNotFoundException("В XML файле отсутствует элемент <Relationships>.");

                //Проходимся по его элементам
                foreach (XElement element in relationships.Elements())
                {
                    token.ThrowIfCancellationRequested();

                    //У каждого проверяем аттрибут
                    XAttribute? id = element.Attributes().FirstOrDefault(a => a.Name == XNamesDocx.RelationshipId);

                    if (id?.Value is not null
                    && keyWords.FileReplaceIds.Contains(id.Value))//И если аттрибут у нас помечен как не используемый в документе(при замене его туда записали)
                    {
                        //Удаляем ссылку
                        element.Remove();
                    }
                }

                //Сохраняем измененный  document.xml.rels файл
                await DocumentSave(relsStream, document, token);
            }
        }

        private static async Task AddFiles(ZipArchive docx, KeyWordsHandlerModelDocx keyWords, CancellationToken token)
        {
            if (keyWords.KeyWordsToFileReplace.Count == 0)
            {
                return;
            }

            //Редактирование файла document.xml.rels
            ZipArchiveEntry relsEntry = docx.GetEntry(_RelsPath) ?? throw new FileNotFoundException($"Не найден файл {_RelsPath}");
            await using (Stream relsStream = await relsEntry.OpenAsync(token))
            {
                XDocument relsDocument = await XDocument.LoadAsync(relsStream, LoadOptions.None, token);
                XElement relationships = relsDocument!.Root ?? throw new FileNotFoundException($"В {_RelsPath} файле отсутствует элемент <Relationships>.");

                //Редактирование файла [Content_Types].xml
                ZipArchiveEntry Content_TypesEntry = docx.GetEntry(Content_TypesPath) ?? throw new FileNotFoundException($"Не найден файл {Content_TypesPath}");
                await using (Stream Content_TypesStream = await Content_TypesEntry.OpenAsync(token))
                {
                    XDocument Content_TypesDocument = await XDocument.LoadAsync(Content_TypesStream, LoadOptions.None, token);
                    XElement types = Content_TypesDocument!.Root ?? throw new FileNotFoundException($"В {Content_TypesPath} файле отсутствует элемент <Types>.");

                    //Получаем все расширения, уже находящиеся в файле [Content_Types].xml
                    //Нужно во избежание дубликатов, ибо они приводят к ошибке
                    HashSet<string> existsExtensions = [.. types.Elements()
                    .Where(e => e.Name == XNamesDocx.TypeDefault && e.HasAttributes)
                    .Select(e => e.Attribute(XNamesDocx.TypeDefaultExtension)?.Value ?? throw new ArgumentException($"В файле {Content_TypesPath} у одного из элементов Default отсутствует аттрибут {XNamesDocx.TypeDefaultExtension}"))];

                    //базовый счетчик для файлов
                    int count = 1;

                    //перебираем все файлы пользователя,
                    //заполняем document.xml.rels, [Content_Types].xml и добавляем файл в архив
                    foreach (KeyValuePair<string, File> pair in keyWords.KeyWordsToFileReplace)
                    {
                        token.ThrowIfCancellationRequested();

                        //Проверяем наличие файла
                        if (pair.Value is File file)
                        {
                            string fileId = $"{FileIdPart}{count++}.{file.FileExtension.Extension}";

                            //Относительный путь файла, без "word" папки в начале
                            string relativeFilePath = $"{file.FileExtension.FolderPath}/{fileId}";

                            //Добавляем файл в архив
                            {
                                ZipArchiveEntry fileEntry = docx.CreateEntry($"{WordPath}/{relativeFilePath}");

                                await using (Stream fileEntryStream = await fileEntry.OpenAsync(token))
                                {
                                    await file.Stream.CopyToAsync(fileEntryStream, token);
                                }
                            }

                            //создаем новую строку <Relationship>
                            XElement relationship = CreateRelationship(fileId, relativeFilePath, file.FileExtension.RelationshipTypeScheme);
                            relationships.Add(relationship);

                            //Если тип файла не был ранее записан в [Content_Types].xml
                            if (!existsExtensions.Contains(file.FileExtension.Extension))
                            {
                                //Создаем и добавляем новый MIME тип, элемент <Default>
                                XElement type = CreateTypeDefault(file.FileExtension.Extension, file.FileExtension.ContentType);
                                types.AddFirst(type);
                                //Важно проверять наличие существующих, потому что дублирование строк <Default> приведет к ошибке открытия
                            }

                            //Добавляем ключ-слово и fileId в словарь для подмены embed аттрибутов
                            keyWords.KeyWordsToFileWordId.Add(pair.Key, fileId);
                        }
                    }

                    //Сохраняем измененный [Content_Types].xml файл
                    await DocumentSave(Content_TypesStream, Content_TypesDocument, token);
                }

                //Сохраняем измененный document.xml.rels файл
                await DocumentSave(relsStream, relsDocument, token);
            }

            /// <summary>
            /// Создание новой строки-элемента <Relationship>
            /// </summary>
            /// <param name="id">Value для аттрибута Id, в файле document.xml будут ссылки на эту строку для подрузки файла в документ, например Id="file1"</param>
            /// <param name="link">относительная ссылка на файл, например Target="media/_file2.gif"</param>
            /// <param name="type">Тип файла, например Type="image/gif"</param>
            /// <returns></returns>
            static XElement CreateRelationship(string id, string link, XNamespace type)
            {
                return new XElement(XNamesDocx.Relationship,
                            new XAttribute(XNamesDocx.RelationshipId, id),
                            new XAttribute(XNamesDocx.RelationshipType, type),
                            new XAttribute(XNamesDocx.RelationshipTarget, link)
                            );
            }

            /// <summary>
            /// Создание новой строки-элемента <Default>
            /// </summary>
            /// <param name="extension">расширение файла, Extension="jpeg"</param>
            /// <param name="contentType">Тип файла MIME, ContentType="image/jpeg"</param>
            /// <returns></returns>
            static XElement CreateTypeDefault(string extension, string contentType)
            {
                return new XElement(XNamesDocx.TypeDefault,
                            new XAttribute(XNamesDocx.TypeDefaultExtension, extension),
                            new XAttribute(XNamesDocx.TypeDefaultContentType, contentType)
                            );
            }
        }
        #endregion

        /// <summary>
        /// Сохранение документа в поток, актуально для всех изменяемых xml документов 
        /// </summary>
        /// <param name="documentXmlStream">Поток</param>
        /// <param name="document">Документ</param>
        private static async Task DocumentSave(Stream documentXmlStream, XDocument document, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();

            long oldStreamLength = documentXmlStream.Length;//Размера шаблона

            documentXmlStream.Seek(0, SeekOrigin.Begin);//Что бы документ заменялся, а не просто сохранился в конце имеющегося
            await document.SaveAsync(documentXmlStream, SaveOptions.None, token);

            //Если позиция в потоке после сохранения файла меньше изначальной длины потока - поток нужно урезать, что бы в итоговом файле не попали элементы шаблона
            if (oldStreamLength > documentXmlStream.Position)
            {
                documentXmlStream.SetLength(documentXmlStream.Position);
            }
        }
    }
}
