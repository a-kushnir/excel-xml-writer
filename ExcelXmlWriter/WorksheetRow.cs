using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetRow: IXmlWriter
    {
        #region Fields

        private readonly WorksheetCellCollection m_Cells;
        private Boolean m_AutoFitHeight;
        private Double m_Height;

        private Int32 m_Index;
        private WorksheetTable m_Parent;
        private Boolean m_IsEmpty = true;

        #endregion

        #region Initialization

        internal WorksheetRow(WorksheetTable table)
        {
            m_Parent = table;
            m_Cells = new WorksheetCellCollection(this);
            m_Index = -1;

            foreach (WorksheetColumn column in table.Columns)
                NewCell(column);
        }

        #endregion

        #region Public Properties

        public WorksheetCellCollection Cells
        {
            get { return m_Cells; }
        }

        public Boolean AutoFitHeight
        {
            get { return m_AutoFitHeight; }
            set { m_AutoFitHeight = value; }
        }

        public Double Height
        {
            get { return m_Height; }
            set { m_Height = value; }
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

        internal WorksheetCell NewCell(WorksheetColumn column)
        {
            return NewCell(column, null);
        }

        internal WorksheetCell NewCell(WorksheetColumn column, WorksheetCell referer)
        {
            WorksheetCell cell = new WorksheetCell(m_Parent, column, this, referer);
            m_Cells.Add(cell);
            return cell;
        }

        public WorksheetCell NewCell(String text)
        {
            return NewCell(text, DataType.String, null);
        }

        public WorksheetCell NewCell(String text, DataType type)
        {
            return NewCell(text, type, null);
        }

        public WorksheetCell NewCell(String text, WorksheetStyle style)
        {
            return NewCell(text, DataType.String, style);
        }

        public WorksheetCell NewCell(String text, DataType type, WorksheetStyle style)
        {
            WorksheetCell cell = new WorksheetCell(m_Parent, null, this, text, type, style);
            m_Cells.Add(cell);

            cell.Column.IsEmpty = false;
            cell.Row.IsEmpty = false;

            return cell;
        }

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Row");
            writer.WriteAttributeString("ss", "AutoFitHeight", null, Helper.XmlConvert(m_AutoFitHeight));
            if (m_Height > 0)
                writer.WriteAttributeString("ss", "Height", null, XmlConvert.ToString(m_Height));
            m_Cells.ToXml(writer);
            writer.WriteEndElement();
        }

        #endregion
    }
}
