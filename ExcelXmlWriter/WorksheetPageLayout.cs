using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetPageLayout : IXmlWriter
    {
        #region Fields

        private Boolean m_CenterHorizontal;
        private Boolean m_CenterVertical;
        private Orientation m_Orientation;
        private Int32 m_StartPageNumber;

        #endregion

        #region Public Properties

        public Boolean CenterHorizontal
        {
            get { return m_CenterHorizontal; }
            set { m_CenterHorizontal = value; }
        }

        public Boolean CenterVertical
        {
            get { return m_CenterVertical; }
            set { m_CenterVertical = value; }
        }

        public Orientation Orientation
        {
            get { return m_Orientation; }
            set { m_Orientation = value; }
        }

        public Int32 StartPageNumber
        {
            get { return m_StartPageNumber; }
            set { m_StartPageNumber = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Layout");

			if (m_Orientation != Orientation.NotSet)
				writer.WriteAttributeString("x", "Orientation", null, m_Orientation.ToString());
			if (m_CenterHorizontal)
				writer.WriteAttributeString("x", "CenterHorizontal", null, Helper.XmlConvert(m_CenterHorizontal));
			if (m_CenterVertical)
				writer.WriteAttributeString("x", "CenterVertical", null, Helper.XmlConvert(m_CenterVertical));
			if (m_StartPageNumber != 0)
				writer.WriteAttributeString("x", "StartPageNumber", null, XmlConvert.ToString(m_StartPageNumber));

        	writer.WriteEndElement();
        }

        #endregion
    }
}
