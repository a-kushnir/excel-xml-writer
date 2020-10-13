using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetColumn: IXmlWriter
    {
        #region Fields

        private Boolean m_AutoFitWidth;
        private Boolean m_Hidden;
        private Int32 m_Span;
        private Double m_Width;
        private WorksheetStyle m_Style;

        private Int32 m_Index;
        private WorksheetTable m_Parent;
        private Boolean m_IsEmpty = true;

        #endregion

        #region Initialization

        internal WorksheetColumn(WorksheetTable table, Double width)
        {
            m_Parent = table;
            m_Index = -1;
            m_Width = width;

            foreach (WorksheetRow row in table.Rows)
                row.NewCell(this);
        }

        #endregion

        #region Public Properties

        public Boolean AutoFitWidth
        {
            get { return m_AutoFitWidth; }
            set { m_AutoFitWidth = value; }
        }

        public Boolean Hidden
        {
            get { return m_Hidden; }
            set { m_Hidden = value; }
        }

        public Int32 Span
        {
            get { return m_Span; }
            set { m_Span = value; }
        }

        public Double Width
        {
            get { return m_Width; }
            set { m_Width = value; }
        }

        public WorksheetStyle Style
        {
            get { return m_Style; }
            set { m_Style = value; }
        }

        public Int32 Index
        {
            get { return m_Index; }
            internal set { m_Index = value; }
        }

        public WorksheetTable Table
        {
            get { return m_Parent; }
            internal set { m_Parent = value; }
        }

        internal Boolean IsEmpty
        {
            get { return m_IsEmpty; }
            set { m_IsEmpty = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Column");

            writer.WriteAttributeString("ss", "AutoFitWidth", null, Helper.XmlConvert(m_AutoFitWidth));
            if (m_Width > 0)
                writer.WriteAttributeString("ss", "Width", null, XmlConvert.ToString(m_Width));
            if (m_Style != null)
                writer.WriteAttributeString("ss", "StyleID", null, m_Style.StyleID);
            if (m_Hidden)
                writer.WriteAttributeString("ss", "Hidden", null, Helper.XmlConvert(m_Hidden));
            if (m_Span > 1)
                writer.WriteAttributeString("ss", "Span", null, XmlConvert.ToString(m_Span));

            writer.WriteEndElement();
        }

        #endregion
    }
}
