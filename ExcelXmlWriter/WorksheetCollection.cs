using System.Collections.Generic;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetCollection : List<Worksheet>, IXmlWriter
    {
        #region Fields

        private readonly Workbook m_Workbook;

        #endregion

        #region Initialization

        internal WorksheetCollection(Workbook workbook)
        {
            m_Workbook = workbook;
        }

        #endregion

        #region Public Methods

        public new void Add(Worksheet worksheet)
        {
            base.Add(worksheet);

            if (Count == 1)
                m_Workbook.ActiveWorksheet = worksheet;
        }

        public void ToXml(XmlWriter writer)
        {
            foreach (Worksheet item in this)
                item.ToXml(writer);
        }

        #endregion
    }
}
