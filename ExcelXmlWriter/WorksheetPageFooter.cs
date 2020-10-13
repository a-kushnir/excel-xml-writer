using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetPageFooter: IXmlWriter
    {
        #region Field

        private String m_Data;
        private Double m_Margin;

        #endregion

        #region Public Properties

        public String Data
        {
            get { return m_Data; }
            set { m_Data = value; }
        }

        public Double Margin
        {
            get { return m_Margin; }
            set { m_Margin = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Footer");

			if (m_Margin != 0)
				writer.WriteAttributeString("x", "Margin", null, XmlConvert.ToString(m_Margin));
			if (!String.IsNullOrEmpty(Data))
				writer.WriteAttributeString("x", "Data", null, Helper.XmlConvert(m_Data));

        	writer.WriteEndElement();
        }

        #endregion
    }
}
