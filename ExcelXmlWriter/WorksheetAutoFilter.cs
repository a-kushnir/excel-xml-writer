using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetAutoFilter: IXmlWriter
    {
        #region Fields

        public string m_Range;

        #endregion

        #region Initialization
        
        internal WorksheetAutoFilter()
        {
        }

        #endregion

        #region Public Properties

        public string Range
        {
            get { return m_Range; }
            set { m_Range = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
