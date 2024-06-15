namespace OpenXmlTemplator.Docx.Auxiliary
{
    using System.Xml.Linq;

    internal static class XNamesDocx
    {
        #region Имена элементов docx
        internal static readonly XNamespace Scheme = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// Тело документа
        /// </summary>
        internal static readonly XName Body = Scheme + "body";

        /// <summary>
        /// Параграф
        /// </summary>
        internal static readonly XName P = Scheme + "p";

        /// <summary>
        /// Строка
        /// </summary>
        internal static readonly XName R = Scheme + "r";

        /// <summary>
        /// Стиль параграфа
        /// </summary>
        internal static readonly XName pPr = Scheme + "pPr";

        /// <summary>
        /// Стиль строки
        /// </summary>
        internal static readonly XName rPr = Scheme + "rPr";

        /// <summary>
        /// Блок с текстом
        /// </summary>
        internal static readonly XName T = Scheme + "t";

        /// <summary>
        /// Перенос/разрыв
        /// </summary>
        internal static readonly XName BR = Scheme + "br";

        /// <summary>
        /// Аттрибут-тип
        /// </summary>
        internal static readonly XName Type = Scheme + "type";

        /// <summary>
        /// Аттрибут-строка таблицы
        /// </summary>
        internal static readonly XName TR = Scheme + "tr";

        /// <summary>
        /// Аттрибут-ячейка в строке таблицы
        /// </summary>
        internal static readonly XName TC = Scheme + "tc";
        #endregion
    }
}
