using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
    internal static class XNamesDocx
    {
        #region XNamespace
        /// <summary>
        ///Основная схема
        /// </summary>
        private static readonly XNamespace WScheme = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        /// <summary>
        /// Схема для графики в целом
        /// </summary>
        private static readonly XNamespace AScheme = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>
        /// Схема для изображений
        /// </summary>
        private static readonly XNamespace WPScheme = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

        /// <summary>
        /// Схема для Relationship, внутри файла document.xml.rels, ссылочки короче
        /// </summary>
        private static readonly XNamespace RScheme = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>
        /// Схема для ссылок на файл document.xml.rels из document.xml, не путать с RScheme!!!
        /// </summary>
        private static readonly XNamespace RLinkScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Схема для Types объектов файла ContentTypes.xml, расширения файлов
        /// </summary>
        private static readonly XNamespace TScheme = "http://schemas.openxmlformats.org/package/2006/content-types";

        /// <summary>
        /// Схема для Relationship.Type, изображения
        /// </summary>
        internal static readonly XNamespace ImageScheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        #endregion

        #region XElement
        /// <summary>
        /// Тело документа
        /// </summary>
        internal static readonly XName Body = WScheme + "body";

        /// <summary>
        /// Параграф
        /// </summary>
        internal static readonly XName P = WScheme + "p";

        /// <summary>
        /// Строка
        /// </summary>
        internal static readonly XName R = WScheme + "r";

        /// <summary>
        /// Стиль параграфа
        /// </summary>
        internal static readonly XName pPr = WScheme + "pPr";

        /// <summary>
        /// Стиль строки
        /// </summary>
        internal static readonly XName rPr = WScheme + "rPr";

        /// <summary>
        /// Блок с текстом
        /// </summary>
        internal static readonly XName T = WScheme + "t";

        /// <summary>
        /// Перенос/разрыв
        /// </summary>
        internal static readonly XName BR = WScheme + "br";

        /// <summary>
        /// Тип
        /// </summary>
        internal static readonly XName Type = WScheme + "type";

        /// <summary>
        /// Строка таблицы
        /// </summary>
        internal static readonly XName TR = WScheme + "tr";

        /// <summary>
        /// Ячейка в строке таблицы
        /// </summary>
        internal static readonly XName TC = WScheme + "tc";

        /// <summary>
        /// Аттрибут-Alt text, доп сведения о картинке, тут аттрибут находится "descr"
        /// </summary>
        internal static readonly XName docPr = WPScheme + "docPr";

        /// <summary>
        /// Заполнение элемента-изображения,тут аттрибут находится "embed"
        /// </summary>
        internal static readonly XName Blip = AScheme + "blip";

        /// <summary>
        /// Элемент отношений в файле document.xml.rels, 
        /// </summary>
        internal static readonly XName Relationship = RScheme + "Relationship";

        /// <summary>
        /// Обработчик типа по умолчанию файлов в ContentTypes.xml
        /// </summary>
        internal static readonly XName TypeDefault = TScheme + "Default";
        #endregion

        #region XAttribute
        /// <summary>
        /// Аттрибут для ссылок на файл document.xml.rels, Relationship.Id аттрибут
        /// </summary>
        internal static readonly XName BlipEmbed = RLinkScheme + "embed";

        /// <summary>
        /// Аттрибут описания ALt Text для картинок
        /// </summary>
        internal static readonly XName docPrDescr = "descr";

        /// <summary>
        /// Аттрибут файла document.xml.rels, устанавливает Id для связи с элементами в document.xml и папке media
        /// </summary>
        internal static readonly XName RelationshipId = "Id";

        /// <summary>
        /// Аттрибут файла document.xml.rels, устанавливает Type схему для обозначения с каким типом файла связывается запись
        /// </summary>
        internal static readonly XName RelationshipType = "Type";

        /// <summary>
        /// Аттрибут файла document.xml.rels, устанавливает Target, путь до файла внутри папки word(например указывается media/image1.jpeg)
        /// </summary>
        internal static readonly XName RelationshipTarget = "Target";

        /// <summary>
        /// Аттрибут файла ContentTypes.xml, расширение файла, краткое
        /// </summary>
        internal static readonly XName TypeDefaultExtension = "Extension";

        /// <summary>
        /// Аттрибут файла ContentTypes.xml, contentType, более полное обозначение расширения файла
        /// </summary>
        internal static readonly XName TypeDefaultContentType = "ContentType";
        #endregion
    }
}
