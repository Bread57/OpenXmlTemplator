namespace OpenXmlTemplator.Docx.Models.OuterModels
{
    using System.IO;

    /// <summary>
    /// Модель для формирования docx документа по шаблону
    /// </summary>
    public class DocumentModelDocx
    {
        /// <summary>
        /// Шаблон
        /// </summary>
        public Stream InStream { get; } = null!;

        /// <summary>
        /// Спиоск параметров для каждого документа
        /// </summary>
        public IEnumerable<(string documentName, KeyWordsHandlerModelDocx keyWords)> Documents { get; } = [];

        /// <summary>
        /// Модель, содержащая информацию для поиска ключевых слов
        /// </summary>
        public SearchModelDocx SearchModel { get; }

        /// <summary>
        /// Настройки встроенных обработчиков ключевых слов, могу совпадать с пользовательскими, но имеют приоритет ниже
        /// </summary>
        public BuiltInKeyWordsHandlersDocx BuiltInKeyWordsHandlers { get; init; } = new BuiltInKeyWordsHandlersDocx();

        private DocumentModelDocx(SearchModelDocx searchModel)
        {
            SearchModel = searchModel;
        }

        #region Stream constructor
        private DocumentModelDocx(Stream inStream, SearchModelDocx searchModel) : this(searchModel)
        {
            InStream = inStream;
        }

        /// <summary>
        /// Принимает один набор ключевых слов, без имени документа
        /// </summary>
        /// <param name="inStream">Шаблон</param>
        /// <param name="keyWords">Ключевые слова</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(Stream inStream, KeyWordsHandlerModelDocx keyWords, SearchModelDocx searchModel) : this(inStream, searchModel)
        {
            Documents = [(string.Empty, keyWords)];
        }

        /// <summary>
        /// Принимает коллекию наборов ключевых слов, вместо имен идут порядковые номера
        /// </summary>
        /// <param name="inStream">Шаблон</param>
        /// <param name="keyWordsCollection">коллекия ключевых слов</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(Stream inStream, IEnumerable<KeyWordsHandlerModelDocx> keyWordsCollection, SearchModelDocx searchModel) : this(inStream, searchModel)
        {
            int documentCount = 1;

            ICollection<(string, KeyWordsHandlerModelDocx)> documents = [];

            foreach (KeyWordsHandlerModelDocx keyWords in keyWordsCollection)
            {
                documents.Add(($"{documentCount++}", keyWords));
            }

            Documents = documents;
        }

        /// <summary>
        /// Принимает один набор ключевых слов с именем документа
        /// </summary>
        /// <param name="inStream">Шаблон</param>
        /// <param name="documentName">Имя документа</param>
        /// <param name="keyWords">Ключевые слова</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(Stream inStream, string documentName, KeyWordsHandlerModelDocx keyWords, SearchModelDocx searchModel) : this(inStream, searchModel)
        {
            Documents = [(documentName, keyWords)];
        }

        /// <summary>
        /// Принимает коллекцию кортежей (имя документа, набор ключевых слов)
        /// </summary>
        /// <param name="inStream">м</param>
        /// <param name="documents">Коллекция кортежей (имя документа, набор ключевых слов)</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(Stream inStream, IEnumerable<(string, KeyWordsHandlerModelDocx)> documents, SearchModelDocx searchModel) : this(inStream, searchModel)
        {
            Documents = documents;
        }
        #endregion

        #region Byte[] constructor
        private DocumentModelDocx(byte[] data, SearchModelDocx searchModel) : this(searchModel)
        {
            InStream = new MemoryStream(data);
        }

        /// <summary>
        /// Принимает один набор ключевых слов, без имени документа
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="keyWords">Ключевые слова</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(byte[] data, KeyWordsHandlerModelDocx keyWords, SearchModelDocx searchModel) : this(data, searchModel)
        {
            Documents = [(string.Empty, keyWords)];
        }

        /// <summary>
        /// Принимает коллекию наборов ключевых слов, вместо имен идут порядковые номера
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="keyWordsCollection">коллекия ключевых слов</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(byte[] data, IEnumerable<KeyWordsHandlerModelDocx> keyWordsCollection, SearchModelDocx searchModel) : this(data, searchModel)
        {
            int documentCount = 1;

            ICollection<(string, KeyWordsHandlerModelDocx)> documents = [];

            foreach (KeyWordsHandlerModelDocx keyWords in keyWordsCollection)
            {
                documents.Add(($"{documentCount++}", keyWords));
            }

            Documents = documents;
        }

        /// <summary>
        /// Принимает один набор ключевых слов с именем документа
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="documentName">Имя документа</param>
        /// <param name="keyWords">Ключевые слова</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(byte[] data, string documentName, KeyWordsHandlerModelDocx keyWords, SearchModelDocx searchModel) : this(data, searchModel)
        {
            Documents = [(documentName, keyWords)];
        }

        /// <summary>
        /// Принимает коллекцию кортежей (имя документа, набор ключевых слов)
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="documents">Коллекция кортежей (имя документа, набор ключевых слов)</param>
        /// <param name="searchModel">Модель для поиска</param>
        public DocumentModelDocx(byte[] data, IEnumerable<(string, KeyWordsHandlerModelDocx)> documents, SearchModelDocx searchModel) : this(data, searchModel)
        {
            Documents = documents;
        }
        #endregion
    }
}
