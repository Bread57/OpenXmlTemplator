using OpenXmlTemplator.Docx.Auxiliary;
using OpenXmlTemplator.Docx.Models.InnerModels;
using OpenXmlTemplator.Docx.Models.OuterModels;
using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Поиск и замена ключевых слов
    /// </summary>
    internal class SearchAndReplaceDocx
    {
        /// <summary>
        /// Рекурсивный перебор двера xml, для поиска и замены ключевых слов
        /// </summary>
        /// <param name="keyWordsHandler">Коллекции ключевых слов для замены</param>
        /// <param name="element">Элемент, дочерние узлы которого будет перебирать</param>
        /// <param name="search">Состояние поиска</param>
        /// <param name="toRemove">элементы, которые будут удаленны в конце</param>
        internal static void RecursiveSearch(XElement element, KeyWordsHandlerModelDocx keyWordsHandler, SearchingKeyWordModelDocx search, ICollection<XElement> toRemove)
        {
            foreach (XElement child in element.Elements())
            {
                if (child.Name == XNamesDocx.T)
                {
                    for (int i = 0; i < child.Value.Length; i++)
                    {
                        char symbol = child.Value[i];

                        //Если набрались стартовые ключи
                        if (search.HasAllStartingKeys)
                        {
                            //если символ входит в последовательность конечных ключей
                            if (search.IsPartOfEndingKeys(symbol: symbol))
                            {
                                //Проверяем, все ли конечные ключи найдены
                                if (search.HasAllEndingKeys)
                                {
                                    //Получаем ключевое слово и его параметры
                                    string[] keyWordParams = search.KeyWord.ToString().Split(search.AdditionalParameters.KeyWordSeparator);

                                    //Получаем отдельно ключевое слово
                                    string keyWord = keyWordParams[0];

                                    keyWordParams = keyWordParams.Skip(1).ToArray();

                                    if (keyWordsHandler.TableKeyWords.TryGetValue(keyWord, out IEnumerable<KeyWordsHandlerModelDocx>? rows))//Проверка на замену таблицы
                                    {
                                        try
                                        {
                                            int tempateRowCount = keyWordParams.Length != 0 ? Convert.ToInt32(keyWordParams[0]) : 1;//если не указано число строк-шаблонов, считаем что строка одна

                                            //Получаем строку-обозначение таблицы
                                            XElement tableSignTr = FindParentByXName(child: child, xName: XNamesDocx.TR) ?? throw new InvalidDataException($"Не найден родительский w:tr блок-обозначение. Ключевое слово - {keyWord ?? "is null"}");
                                            toRemove.Add(tableSignTr);//Добавляем строку-обозначение в список для итогового удаления

                                            //Получаем список строк-шаблонов, идущих после строки-обозначения таблицы
                                            //Обязательно Вызывакм To(Array/List и т.д.) для кэширования результата запроса, т.к. иначе изменения древа xml(вставки новых элементов) будут отражаться на этой коллекции, если оставим Ienumerable
                                            XElement[] trTemplates = tableSignTr.ElementsAfterSelf().Take(tempateRowCount).ToArray() ?? throw new NullReferenceException($"Не найдено шаблон-строка таблицы. Ключевое слово - {keyWord ?? "is null"}");

                                            //Берем первую строку шаблон для вставки новыз элементов перед ней
                                            XElement insertBefore = trTemplates.FirstOrDefault() ?? throw new InvalidDataException($"Не указаны шаблоны строк в таблице. Ключевое слово - {keyWord ?? " is null"}");

                                            //Добавляем шаблоны с писок на удаление
                                            foreach (XElement templateRow in trTemplates)
                                            {
                                                toRemove.Add(templateRow);
                                            }

                                            if (rows is not null)
                                            {
                                                foreach (KeyWordsHandlerModelDocx rowHandler in rows)
                                                {
                                                    foreach (XElement templateRow in trTemplates)
                                                    {
                                                        XElement row = new(templateRow);
                                                        insertBefore.AddBeforeSelf(row);

                                                        RecursiveSearch(keyWordsHandler: rowHandler, element: row, search: new SearchingKeyWordModelDocx(searchToCopy: search), toRemove: toRemove);//Вызываем рекурсивный поиск внутри строки таблицы, SearchKeyWord задаем новое
                                                    }
                                                }
                                            }
                                        }
                                        catch (FormatException ex)
                                        {
                                            throw new FormatException("Не удалось привести параметр 'число строк шаблонов' к int", ex);
                                        }
                                    }
                                    else if (keyWordsHandler.KeyWordsToInsert.TryGetValue(keyWord, out IEnumerable<string>? data))//Проверка на множественные замены
                                    {
                                        //Получем строку и ее стиль
                                        XElement run = FindParentByXName(child: child, xName: XNamesDocx.R) ?? throw new InvalidDataException($"Не найден родительский w:r блок. Ключевое слово - {keyWord ?? "is null"}");
                                        XElement? rStyle = run.Element(XNamesDocx.rPr);

                                        //Получаем параграф-обозначение и его стиль
                                        XElement paragraph = FindParentByXName(child: run, xName: XNamesDocx.P) ?? throw new InvalidDataException($"Не найден родительский w:p блок. Ключевое слово - {keyWord ?? "is null"}"); ;
                                        XElement? pStyle = paragraph.Element(XNamesDocx.pPr);

                                        foreach (string text in data)
                                        {
                                            XElement newParagraph = CreateParagraph(text: ApplyingParameters(search.AdditionalParameters, text, keyWordParams), pStyle: pStyle, rStyle: rStyle);//Создаем новый параграф с указаными стилями

                                            paragraph.AddBeforeSelf(newParagraph);//вставляем новый элемент
                                        }

                                        toRemove.Add(paragraph);//удаляем параграф-обозначнение
                                    }
                                    else
                                    {
                                        //Индекс ключевого слова
                                        int startIndex = search.StartIndex;

                                        //Удаляем ключевое слово из документа
                                        child.Value = child.Value.Remove(startIndex, i - startIndex + 1);

                                        if (keyWordsHandler.KeyWordsToReplace.TryGetValue(keyWord, out string? replaceValue))//Проверка на простую замену
                                        {
                                            if (keyWordParams.Length > 0)
                                            {
                                                replaceValue = ApplyingParameters(search.AdditionalParameters, replaceValue, keyWordParams);
                                            }
                                        }
                                        else//Если для ключевого слова нет обработчика - оставляем предупреждение на его месте
                                        {
                                            //Вставляем значение в индекс ключевого слова
                                            replaceValue = $"{keyWordsHandler.KeyWordHandlerNotFoundMessage}: {keyWord}";
                                        }

                                        //Вставляем значение в индекс ключевого слова
                                        child.Value = child.Value.Insert(startIndex, replaceValue);

                                        //Ставим переменной цикла индекс ключевого слова, что бы захватить все возможные ключевые слова
                                        i = startIndex;
                                    }

                                    //Сбрасываем параметры
                                    search.Reset();
                                }
                            }
                            else
                            {
                                //Добавляем символ в ключевое слово
                                search.KeyWord.Append(symbol);
                            }
                        }
                        else
                        {
                            //Смотри, является ли символ стартовым ключем
                            search.IsPartOfStartingKeys(symbol: symbol, index: i);
                        }
                    }

                    //Если переходим в другой блок/элемент и при этом ключевое слово начало собираться
                    if (search.HasAllStartingKeys)
                    {
                        //Удаляем все символы, начиная с ключевых и до конца блока текста
                        child.Value = child.Value.Remove(search.StartIndex);
                        search.StartIndex = 0;
                    }
                }

                if (child.HasElements)//если есть дочерние элементы
                {
                    //Продолжаем обход древа
                    RecursiveSearch(element: child, keyWordsHandler: keyWordsHandler, search: search, toRemove: toRemove);
                }
            }
        }

        /// <summary>
        /// Примененение параметров к замененному слову
        /// </summary>
        /// <param name="additionalParameters"></param>
        /// <param name="replaceValue"></param>
        /// <param name="keyWordParameters"></param>
        /// <returns></returns>
        private static string ApplyingParameters(AdditionalParametersDocx additionalParameters, string replaceValue, params string[] keyWordParameters)
        {
            foreach (string param in keyWordParameters)
            {
                string[] splitParam = param.Split(separator: additionalParameters.ParameterSeparator);

                if (additionalParameters.Handlers.TryGetValue(splitParam[0], out Func<string, string, string>? handler))
                {
                    replaceValue = handler!.Invoke(replaceValue, splitParam.Length > 1 ? splitParam[1] : string.Empty);
                }
            }

            return replaceValue;
        }

        /// <summary>
        /// Обход древа элемента, пока не найдем родительский класс с нужным именем
        /// </summary>
        /// <param name="child">Дочерний элемент, для которого ищем родителя</param>
        /// <param name="xName">Имя элемента</param>
        /// <returns></returns>
        private static XElement? FindParentByXName(XElement? child, XName xName)
        {
            if (child?.Name == xName)
            {
                return child;
            }

            return child?.Parent is not null ? FindParentByXName(child: child.Parent, xName: xName) : null;
        }

        /// <summary>
        /// Создание параграфа с текстом
        /// </summary>
        /// <param name="text">Текст</param>
        /// <param name="pStyle">Стиль параграфа</param>
        /// <param name="rStyle">Стиль строки</param>
        /// <returns></returns>
        private static XElement CreateParagraph(string text, XElement? pStyle, XElement? rStyle)
        {
            return new XElement(XNamesDocx.P,
                pStyle is not null && pStyle.Name == XNamesDocx.pPr ? new XElement(pStyle) : null,
                CreateRow(text: text, rStyle: rStyle));
        }

        /// <summary>
        /// Создание строки с текстом
        /// </summary>
        /// <param name="text">Текст</param>
        /// <param name="style">стиль</param>
        /// <returns></returns>
        private static XElement CreateRow(string text, XElement? rStyle)
        {
            return new XElement(XNamesDocx.R,
                 rStyle is not null && rStyle.Name == XNamesDocx.rPr ? new XElement(rStyle) : null, //если стиль не передан, или имя элемента не совпадает с принятой docx схемой - опускаем указание стиля
                 new XElement(XNamesDocx.T, text));
        }
    }
}
