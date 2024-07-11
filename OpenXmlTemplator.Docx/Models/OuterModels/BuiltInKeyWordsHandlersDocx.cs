using System.Globalization;

namespace OpenXmlTemplator.Docx.Models.OuterModels
{
    /// <summary>
    /// Настройки встроенных обработчиков ключевых слов, могу совпадать с пользовательскими, но имеют приоритет ниже
    /// </summary>
    public class BuiltInKeyWordsHandlersDocx
    {
        #region Счетчик строк в таблицах
        /// <summary>
        /// Обозначение счетчика строк в таблице
        /// </summary>
        public string TableRowCounter_Sign { get; init; } = "№";

        /// <summary>
        /// Стартовое значение счетчика строк в таблице
        /// </summary>
        public int TableRowCounter_StartValue { get; init; } = 1;

        /// <summary>
        /// Текущее значение счетчика строк
        /// </summary>
        internal int TableRowCounter_CurrentValue { get; set; }

        /// <summary>
        /// Нужно ли сбрасывать счетчик строк при завершении заполнения шаблонов-строк для каждого элемента таблицы(если true - 1,2,3,1,2,3 и т.д.)
        /// </summary>
        public bool TableRowCounter_ResetByTemplateList { get; set; } = false;

        /// <summary>
        /// Использовать слова вместо цифо
        /// </summary>
        public bool TableRowCounter_UseWords { get; set; }

        /// <summary>
        /// Культура для перевода цифр в строку
        /// </summary>
        public CultureInfo TableRowCounter_WordsCulture { get; set; } = CultureInfo.CurrentCulture;
        #endregion
    }
}
