using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetPageMargins : IXmlWriter
    {
        #region Fields

        private Double m_Bottom;
        private Double m_Left;
        private Double m_Right;
        private Double m_Top;

        #endregion

        #region Public Properties

        public Double Bottom
        {
            get { return m_Bottom; }
            set { m_Bottom = value; }
        }

        public Double Left
        {
            get { return m_Left; }
            set { m_Left = value; }
        }

        public Double Right
        {
            get { return m_Right; }
            set { m_Right = value; }
        }

        public Double Top
        {
            get { return m_Top; }
            set { m_Top = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("PageMargins");

			if (m_Bottom != 0)
				writer.WriteAttributeString("x", "Bottom", null, XmlConvert.ToString(m_Bottom));
			if (m_Left != 0)
				writer.WriteAttributeString("x", "Left", null, XmlConvert.ToString(m_Left));
			if (m_Right != 0)
				writer.WriteAttributeString("x", "Right", null, XmlConvert.ToString(m_Right));
			if (m_Top != 0)
				writer.WriteAttributeString("x", "Top", null, XmlConvert.ToString(m_Top));

            writer.WriteEndElement();
        }

        #endregion
    }
}
