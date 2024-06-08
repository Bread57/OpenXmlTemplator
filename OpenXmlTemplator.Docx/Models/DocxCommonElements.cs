namespace OpenXmlTemplator.Docx.Models
{
    using System.Xml.Linq;

    /// <summary>
    /// Стандартные элементы docx документа
    /// </summary>
    internal static class DocxCommonElements
    {
        /// <summary>
        /// Разрыв страницы
        /// </summary>
        internal static XElement PageBreak
        {
            get
            {
                return new(DocxXNames.P,
                        new XElement(DocxXNames.R,
                            new XElement(DocxXNames.BR,new XAttribute(DocxXNames.Type, "page"))
                    )
                );
            }
        } 
    }
}
