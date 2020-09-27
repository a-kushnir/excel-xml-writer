using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class SpreadSheetToolbar: IXmlWriter
    {
        #region Fields

        public Boolean m_Hidden;

        #endregion

        #region Initialization

        internal SpreadSheetToolbar()
        {
        }

        #endregion

        #region Public Methods

        public Boolean Hidden
        {
            get { return m_Hidden; }
            set { m_Hidden = value; }
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
