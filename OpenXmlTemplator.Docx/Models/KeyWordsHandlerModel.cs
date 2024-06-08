namespace OpenXmlTemplator.Docx.Models
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// Набор коллекции ключевых слов
    /// </summary>
    public class KeyWordsHandlerModel
    {
        public KeyWordsHandlerModel(string keyWordHandlerNotFoundMessage)
        {
            KeyWordHandlerNotFoundMessage = keyWordHandlerNotFoundMessage;
        }

        /// <summary>
        /// Сообщение, вставляем перед ключевым словом, когда для него нет обработчика(нет ззаписи в словарях)
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
        public IDictionary<string, IEnumerable<KeyWordsHandlerModel>> TableKeyWords { get; set; } = new Dictionary<string, IEnumerable<KeyWordsHandlerModel>>(StringComparer.OrdinalIgnoreCase);
    }
}
