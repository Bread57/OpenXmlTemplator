using System.Xml.Linq;

namespace OpenXmlTemplator.Docx
{
    /// <summary>
    /// Стандартные элементы docx документа
    /// </summary>
    internal static class CommonElementsDocx
    {
        /// <summary>
        /// Разрыв страницы
        /// </summary>
        internal static XElement PageBreak
        {
            get
            {
                return new(XNamesDocx.P,
                        new XElement(XNamesDocx.R,
                            new XElement(XNamesDocx.BR, new XAttribute(XNamesDocx.Type, "page"))
                    )
                );
            }
        }
    }
}
