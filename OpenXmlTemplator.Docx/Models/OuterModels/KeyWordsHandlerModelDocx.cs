namespace OpenXmlTemplator.Docx.Models.OuterModels
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Набор коллекции ключевых слов
    /// </summary>
    public class KeyWordsHandlerModelDocx
    {
        public KeyWordsHandlerModelDocx(string keyWordHandlerNotFoundMessage)
        {
            KeyWordHandlerNotFoundMessage = keyWordHandlerNotFoundMessage;
        }

        /// <summary>
        /// Сообщение, вставляем перед ключевым словом, когда для него нет обработчика(нет записи в словарях)
        /// </summary>
        public string KeyWordHandlerNotFoundMessage { get; set; }

        /// <summary>
        /// Ключевые слова со значениями для замен
        /// </summary>
        public IDictionary<string, string> KeyWordsToReplace { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Ключевые слова со значениями для вставок
        /// </summary>
        public IDictionary<string, IEnumerable<string>> KeyWordsToInsert { get; set; } = new Dictionary<string, IEnumerable<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Наборы для таблиц, таблицы для нас, все равно что отдельный документ, поэтому для его обработки нужны свои коллекции ключевых слов
        /// </summary>
        public IDictionary<string, IEnumerable<KeyWordsHandlerModelDocx>> TableKeyWords { get; set; } = new Dictionary<string, IEnumerable<KeyWordsHandlerModelDocx>>(StringComparer.OrdinalIgnoreCase);
    }
}
