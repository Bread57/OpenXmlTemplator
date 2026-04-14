namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Модель, содержащая информацию для поиска ключевых слов
    /// </summary>
    public sealed class SearchModelDocx
    {
        /// <summary>
        /// Набор дополнительных параметров для обработки ключевых слов
        /// </summary>
        public AdditionalParametersDocx AdditionalParameters { get; } = new AdditionalParametersDocx(keyWordSeparator: "&", parameterSeparator: ":", []);

        /// <summary>
        /// Набор ключей для обозначения начала ключевого слова, например ['[','#']
        /// </summary>
        public char[] StartingKeys { get; }

        /// <summary>
        /// Набор ключей для обозначения окончания ключевого слова, например ['#',']']
        /// </summary>
        public char[] EndingKeys { get; }

        public SearchModelDocx(char[] startingKeys, char[] endingKeys)
        {
            StartingKeys = startingKeys;
            EndingKeys = endingKeys;
        }

        public SearchModelDocx(char[] startingKeys, char[] endingKeys, AdditionalParametersDocx additionalParameters) : this(startingKeys: startingKeys, endingKeys: endingKeys)
        {
            AdditionalParameters = additionalParameters;
        }
    }
}
