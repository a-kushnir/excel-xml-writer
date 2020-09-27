using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetStyleAlignment : IXmlWriter, ICloneable
    {
        #region Fields

        private StyleHorizontalAlignment m_Horizontal;
        private Int32 m_Indent;
        private StyleReadingOrder m_ReadingOrder;
        private Int32 m_Rotate;
        private Boolean m_ShrinkToFit;
        private StyleVerticalAlignment m_Vertical;
        private Boolean m_VerticalText;
        private Boolean m_WrapText;

        #endregion

        #region Initialization

        public WorksheetStyleAlignment()
        {
            m_Horizontal = StyleHorizontalAlignment.Automatic;
            m_Vertical = StyleVerticalAlignment.Automatic;
            m_ReadingOrder = StyleReadingOrder.NotSet;
        }

        public WorksheetStyleAlignment(StyleHorizontalAlignment horizontal, StyleVerticalAlignment vertical)
            : this()
        {
            m_Horizontal = horizontal;
            m_Vertical = vertical;
        }

        #endregion

        #region Public Properties

        public StyleHorizontalAlignment Horizontal
        {
            get { return m_Horizontal; }
            set { m_Horizontal = value; }
        }

        public Int32 Indent
        {
            get { return m_Indent; }
            set { m_Indent = value; }
        }

        public StyleReadingOrder ReadingOrder
        {
            get { return m_ReadingOrder; }
            set { m_ReadingOrder = value; }
        }

        public Int32 Rotate
        {
            get { return m_Rotate; }
            set { m_Rotate = value; }
        }

        public Boolean ShrinkToFit
        {
            get { return m_ShrinkToFit; }
            set { m_ShrinkToFit = value; }
        }

        public StyleVerticalAlignment Vertical
        {
            get { return m_Vertical; }
            set { m_Vertical = value; }
        }

        public Boolean VerticalText
        {
            get { return m_VerticalText; }
            set { m_VerticalText = value; }
        }

        public Boolean WrapText
        {
            get { return m_WrapText; }
            set { m_WrapText = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Alignment");

            if (m_Horizontal != StyleHorizontalAlignment.Automatic)
                writer.WriteAttributeString("ss", "Horizontal", null, m_Horizontal.ToString());
            if (m_Vertical != StyleVerticalAlignment.Automatic)
                writer.WriteAttributeString("ss", "Vertical", null, m_Vertical.ToString());
            if (m_ReadingOrder != StyleReadingOrder.NotSet)
                writer.WriteAttributeString("ss", "ReadingOrder", null, m_ReadingOrder.ToString());
            if (m_Indent > 0)
                writer.WriteAttributeString("ss", "Indent", null, XmlConvert.ToString(m_Indent));
            if (m_Rotate > 0)
                writer.WriteAttributeString("ss", "Rotate", null, XmlConvert.ToString(m_Rotate));
            if (m_ShrinkToFit)
                writer.WriteAttributeString("ss", "ShrinkToFit", null, Helper.XmlConvert(m_ShrinkToFit));
            if (m_VerticalText)
                writer.WriteAttributeString("ss", "VerticalText", null, Helper.XmlConvert(m_VerticalText));
            if (m_WrapText)
                writer.WriteAttributeString("ss", "WrapText", null, Helper.XmlConvert(m_WrapText));

            writer.WriteEndElement();
        }

        public object Clone()
        {
            WorksheetStyleAlignment clone = new WorksheetStyleAlignment();

            clone.Horizontal = Horizontal;
            clone.Indent = Indent;
            clone.ReadingOrder = ReadingOrder;
            clone.Rotate = Rotate;
            clone.ShrinkToFit = ShrinkToFit;
            clone.Vertical = Vertical;
            clone.VerticalText = VerticalText;
            clone.WrapText = WrapText;

            return clone;
        }

        #endregion
    }
}
