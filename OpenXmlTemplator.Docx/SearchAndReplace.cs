using OpenXmlTemplator.Docx.Models;
using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Поиск и замена ключевых слов
    /// </summary>
    internal class SearchAndReplace
    {
        /// <summary>
        /// Рекурсивный перебор двера xml, для поиска и замены ключевых слов
        /// </summary>
        /// <param name="keyWords">Коллекции ключевых слов для замены</param>
        /// <param name="element">Элемент, дочерние узлы которого будет перебирать</param>
        /// <param name="search">Состояние поиска</param>
        /// <param name="toRemove">элементы, которые будут удаленны в конце</param>
        internal static void RecursiveSearch(XElement element, KeyWordsHandlerModel keyWordsHandler, SearchKeyWord search, ICollection<XElement> toRemove)
        {
            foreach (XElement child in element.Elements())
            {
                if (child.Name == DocxXNames.T)
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
                                    //Получаем ключевое слово
                                    string keyWord = search.KeyWord.ToString();

                                    if (keyWordsHandler.TableKeyWords.TryGetValue(keyWord, out IEnumerable<KeyWordsHandlerModel>? rows))//Проверка на замену таблицы
                                    {
                                        //Получаем строку-обозначение таблицы
                                        XElement tableSignTr = FindParentByXName(child: child, xName: DocxXNames.TR) ?? throw new InvalidDataException($"Не найден родительский w:tr блок-обозначение. Ключевое слово - {keyWord ?? "is null"}");

                                        //Получаем следующую строку, после строки-обозначения таблицы
                                        XElement trParent = tableSignTr.ElementsAfterSelf().FirstOrDefault() ?? throw new NullReferenceException($"Не найдено шаблон-строка таблицы. Ключевое слово - {keyWord ?? "is null"}");

                                        //Указатель на последний элемент, после которого и нужно добавлять новые, для сохранения порядка
                                        XElement insertAfter = trParent;

                                        //Ссылка на шаблон
                                        XElement? trTemplate = null;

                                        if (rows is not null)
                                        {
                                            foreach (KeyWordsHandlerModel rowHandler in rows)
                                            {
                                                if (trTemplate is not null)//Если мы установили шаблон, значит в самом документе строка-пример уже заменена
                                                {
                                                    trParent = new XElement(trTemplate);//создаем новый элемент на основе шаблона
                                                    insertAfter.AddAfterSelf(trParent);//встявляем новый элемент
                                                    insertAfter = trParent;//меняем ссылку на последний элемент
                                                }
                                                else
                                                {
                                                    trTemplate = new XElement(trParent);//Устанавливаем шаблон, для последующих копии
                                                }

                                                RecursiveSearch(keyWordsHandler: rowHandler, element: trParent, search: new SearchKeyWord(searchToCopy: search), toRemove: toRemove);//Вызываем рекурсивный поиск внутри строки таблицы, SearchKeyWord задаем новое
                                            }
                                        }

                                        toRemove.Add(tableSignTr);//Добавляем строку-обозначение в список для итогового удаления
                                    }
                                    else if (keyWordsHandler.KeyWordsToInsert.TryGetValue(keyWord, out IEnumerable<string>? data))//Проверка на множественные замены
                                    {
                                        //Получем строку и ее стиль
                                        XElement run = FindParentByXName(child: child, xName: DocxXNames.R) ?? throw new InvalidDataException($"Не найден родительский w:r блок. Ключевое слово - {keyWord ?? "is null"}");
                                        XElement? rStyle = run.Element(DocxXNames.rPr);

                                        //Получаем параграф-обозначение и его стиль
                                        XElement paragraph = FindParentByXName(child: run, xName: DocxXNames.P) ?? throw new InvalidDataException($"Не найден родительский w:p блок. Ключевое слово - {keyWord ?? "is null"}"); ;
                                        XElement? pStyle = paragraph.Element(DocxXNames.pPr);

                                        //Указатель на последний элемент, после которого и нужно добавлять новые, для сохранения порядка
                                        XElement insertAfter = paragraph;

                                        foreach (string text in data)
                                        {
                                            XElement newParagraph = CreateParagraph(text: text, pStyle: pStyle, rStyle: rStyle);//Создаем новый параграф с указаными стилями

                                            insertAfter.AddAfterSelf(newParagraph);//вставляем новый элемент
                                            insertAfter = newParagraph;//меняем ссылку на последний элемент
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
                                            //Вставляем значение в индекс ключевого слова
                                            child.Value = child.Value.Insert(startIndex, replaceValue);
                                        }
                                        else//Если для ключевого слова нет обработчика - выводим оставляем предупредение на документе
                                        {
                                            //Вставляем значение в индекс ключевого слова
                                            child.Value = child.Value.Insert(startIndex, $"{keyWordsHandler.KeyWordHandlerNotFoundMessage}: {keyWord}");
                                        }

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
            return new XElement(DocxXNames.P,
                pStyle is not null && pStyle.Name == DocxXNames.pPr ? new XElement(pStyle) : null,
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
            return new XElement(DocxXNames.R,
                 rStyle is not null && rStyle.Name == DocxXNames.rPr ? new XElement(rStyle) : null, //если стиль не передан, или имя элемента не совпадает с принятой docx схемой - опускаем указание стиля
                 new XElement(DocxXNames.T, text));
        }
    }
}
