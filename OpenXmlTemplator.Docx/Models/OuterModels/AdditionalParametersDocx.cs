namespace OpenXmlTemplator.Docx.Models.OuterModels
{
    /// <summary>
    /// Модель для передачи параметров с обработчиками, которые позволяют дополнительно поменять итоговое значение
    /// </summary>
    public class AdditionalParametersDocx
    {
        /// <summary>
        /// Отделяет ключевое слово от его параметров 
        /// </summary>
        public string KeyWordSeparator { get; }

        /// <summary>
        /// Разделяет имя параметра и его доп. данные
        /// </summary>
        public string ParameterSeparator { get; }

        /// <summary>
        /// (Key: Параметр,Value: обработчик), каждый делегат принимает на вход замененное ключевое слово и набор параметров для обработчки(в виде строки), возвращает обработанное слово
        /// </summary>
        public Dictionary<string, Func<string, string, string>> Handlers { get; } = new Dictionary<string, Func<string, string, string>>(StringComparer.OrdinalIgnoreCase);

        public AdditionalParametersDocx(string keyWordSeparator, string parameterSeparator, Dictionary<string, Func<string, string, string>> handlers)
        {
            KeyWordSeparator = keyWordSeparator;
            ParameterSeparator = parameterSeparator;
            Handlers = handlers;
        }
    }
}
