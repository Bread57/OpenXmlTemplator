namespace OpenXmlTemplator.Docx.Models.OuterModels
{
    /// <summary>
    /// Модель, содержащая информацию для поиска ключевых слов
    /// </summary>
    public class SearchModelDocx
    {
        /// <summary>
        /// Набор дополнительных параметров для обработки ключевых слов
        /// </summary>
        public AdditionalParametersDocx AdditionalParameters { get; }

        /// <summary>
        /// Набор ключей для обозначения начала ключевого слова
        /// </summary>
        public char[] StartingKeys { get; }

        /// <summary>
        /// Набор ключей для обозначения окончания ключевого слова
        /// </summary>
        public char[] EndingKeys { get; }

        public SearchModelDocx(char[] startingKeys, char[] endingKeys)
        {
            StartingKeys = startingKeys;
            EndingKeys = endingKeys;
            AdditionalParameters = new AdditionalParametersDocx(keyWordSeparator: "&", parameterSeparator: ":", new Dictionary<string, Func<string, string, string>>(0));
        }

        public SearchModelDocx(char[] startingKeys, char[] endingKeys, AdditionalParametersDocx additionalParameters) : this(startingKeys: startingKeys, endingKeys: endingKeys)
        {
            AdditionalParameters = additionalParameters;
        }
    }
}
