namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Набор коллекции ключевых слов
    /// </summary>
    public sealed class KeyWordsHandlerModelDocx
    {
        public KeyWordsHandlerModelDocx(string keyWordHandlerNotFoundMessage)
        {
            KeyWordHandlerNotFoundMessage = keyWordHandlerNotFoundMessage;
        }

        /// <summary>
        /// Сообщение, вставляем перед ключевым словом, когда для него нет обработчика(нет записи в словарях)
        /// </summary>
        public string KeyWordHandlerNotFoundMessage { get; private set; }

        #region Работа с файлами
        /// <summary>
        /// Ключевые слова с файлами под замену
        /// </summary>
        public IDictionary<string, File> KeyWordsToFileReplace { get; init; } = new Dictionary<string, File>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Ключевые слова, связанные с Relation.Id из файла document.xml.rels, для подстановки нужно id файла из папки media
        /// </summary>
        internal IDictionary<string, string> KeyWordsToFileWordId { get; init; } = new Dictionary<string, string>(capacity: 4, StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Id файлов, которые были ПОЛНОСТЬЮ заменены файлы из KeyWordsToFileReplace. Есть файл заменился один или несколько раз, но при этом он все равно остается где-то в документе - то мы его оставляем
        /// </summary>
        internal HashSet<string> FileReplaceIds { get; init; } = new HashSet<string>(capacity: 4, StringComparer.OrdinalIgnoreCase);
        #endregion

        #region  Работа с текстом
        /// <summary>
        /// Ключевые слова со значениями для замен
        /// </summary>
        public IDictionary<string, string> KeyWordsToReplace { get; init; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Ключевые слова со значениями для вставок
        /// </summary>
        public IDictionary<string, IEnumerable<string>> KeyWordsToInsert { get; init; } = new Dictionary<string, IEnumerable<string>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Наборы для таблиц, таблицы для нас, все равно что отдельный документ, поэтому для его обработки нужны свои коллекции ключевых слов
        /// </summary>
        public IDictionary<string, IEnumerable<KeyWordsHandlerModelDocx>> TableKeyWords { get; init; } = new Dictionary<string, IEnumerable<KeyWordsHandlerModelDocx>>(StringComparer.OrdinalIgnoreCase);
        #endregion
    }
}
