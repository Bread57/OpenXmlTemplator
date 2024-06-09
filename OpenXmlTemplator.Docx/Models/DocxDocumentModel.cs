namespace OpenXmlTemplator.Docx.Models
{
    using System.IO;

    /// <summary>
    /// Модель для формирования docx документа по шаблону
    /// </summary>
    public class DocxDocumentModel
    {
        /// <summary>
        /// Шаблон
        /// </summary>
        public Stream InStream { get; } = null!;

        /// <summary>
        /// Спиоск параметров для каждого документа
        /// </summary>
        public IEnumerable<(string documentName, KeyWordsHandlerModel keyWords)> Documents { get; } = [];

        /// <summary>
        /// Разделитель для параметров ключевого слова
        /// </summary>
        internal string KeyWordParamsSeparator { get; }

        /// <summary>
        /// Набор ключей для обозначения начала ключевого слова
        /// </summary>
        public char[] StartingKeys { get; }

        /// <summary>
        /// Набор ключей для обозначения окончания ключевого слова
        /// </summary>
        public char[] EndingKeys { get; }

        private DocxDocumentModel(char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator)
        {
            StartingKeys = startingKeys;
            EndingKeys = endingKeys;
            KeyWordParamsSeparator = keyWordParamsSeparator;
        }

        #region Stream constructor
        private DocxDocumentModel(Stream inStream, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(startingKeys, endingKeys, keyWordParamsSeparator)
        {
            InStream = inStream;
        }

        /// <summary>
        /// Принимает один набор ключевых слов, без имени документа
        /// </summary>
        /// <param name="inStream">Шаблон</param>
        /// <param name="keyWords">Ключевые слова</param>
        /// <param name="startingKeys">Обозначение начала ключевого слова</param>
        /// <param name="endingKeys">Обозначение окончания ключевого слова</param>
        public DocxDocumentModel(Stream inStream, KeyWordsHandlerModel keyWords, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(inStream, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = [(string.Empty, keyWords)];
        }

        /// <summary>
        /// Принимает коллекию наборов ключевых слов, вместо имен идут порядковые номера
        /// </summary>
        /// <param name="inStream">Шаблон</param>
        /// <param name="keyWordsCollection">коллекия ключевых слов</param>
        /// <param name="startingKeys">Обозначение начала ключевого слова</param>
        /// <param name="endingKeys">Обозначение окончания ключевого слова</param>
        public DocxDocumentModel(Stream inStream, IEnumerable<KeyWordsHandlerModel> keyWordsCollection, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(inStream, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            int documentCount = 1;

            ICollection<(string, KeyWordsHandlerModel)> documents = [];

            foreach (KeyWordsHandlerModel keyWords in keyWordsCollection)
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
        /// <param name="startingKeys">Обозначение начала ключевого слова</param>
        /// <param name="endingKeys">Обозначение окончания ключевого слова</param>
        public DocxDocumentModel(Stream inStream, string documentName, KeyWordsHandlerModel keyWords, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(inStream, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = [(documentName, keyWords)];
        }

        /// <summary>
        /// Принимает коллекцию кортежей (имя документа, набор ключевых слов)
        /// </summary>
        /// <param name="inStream">м</param>
        /// <param name="documents">Коллекция кортежей (имя документа, набор ключевых слов)</param>
        /// <param name="startingKeys">Обозначение начала ключевого слова</param>
        /// <param name="endingKeys">Обозначение окончания ключевого слова</param>
        public DocxDocumentModel(Stream inStream, IEnumerable<(string, KeyWordsHandlerModel)> documents, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(inStream, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = documents;
        }
        #endregion

        #region Byte[] constructor
        private DocxDocumentModel(byte[] data, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(startingKeys, endingKeys, keyWordParamsSeparator)
        {
            InStream = new MemoryStream(data);
        }

        /// <summary>
        /// Принимает один набор ключевых слов, без имени документа
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="keyWords">Ключевые слова</param>
        public DocxDocumentModel(byte[] data, KeyWordsHandlerModel keyWords, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(data, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = [(string.Empty, keyWords)];
        }

        /// <summary>
        /// Принимает коллекию наборов ключевых слов, вместо имен идут порядковые номера
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="keyWordsCollection">коллекия ключевых слов</param>
        public DocxDocumentModel(byte[] data, IEnumerable<KeyWordsHandlerModel> keyWordsCollection, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(data, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            int documentCount = 1;

            ICollection<(string, KeyWordsHandlerModel)> documents = [];

            foreach (KeyWordsHandlerModel keyWords in keyWordsCollection)
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
        public DocxDocumentModel(byte[] data, string documentName, KeyWordsHandlerModel keyWords, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(data, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = [(documentName, keyWords)];
        }

        /// <summary>
        /// Принимает коллекцию кортежей (имя документа, набор ключевых слов)
        /// </summary>
        /// <param name="data">Шаблон</param>
        /// <param name="documents">Коллекция кортежей (имя документа, набор ключевых слов)</param>
        public DocxDocumentModel(byte[] data, IEnumerable<(string, KeyWordsHandlerModel)> documents, char[] startingKeys, char[] endingKeys, string keyWordParamsSeparator) : this(data, startingKeys, endingKeys, keyWordParamsSeparator)
        {
            Documents = documents;
        }
        #endregion
    }
}
