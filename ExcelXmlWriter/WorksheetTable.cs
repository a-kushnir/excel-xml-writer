using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetTable: IXmlWriter
    {
        #region Fields

        private readonly Worksheet m_Parent;
        private readonly WorksheetColumnCollection m_Columns;
        private readonly WorksheetRowCollection m_Rows;
        private Double m_DefaultRowHeight;
        private Double m_DefaultColumnWidth;
        private WorksheetStyle m_Style;

        private readonly WorksheetTableCellCollection m_Cells;

        #endregion

        #region Initialization

        internal WorksheetTable(Worksheet parent)
        {
            m_Parent = parent;
            m_DefaultRowHeight = 15;
            m_Columns = new WorksheetColumnCollection();
            m_Rows = new WorksheetRowCollection();
            m_Cells = new WorksheetTableCellCollection(m_Rows);
        }

        #endregion

        #region Public Properties

        public Worksheet Worksheet
        {
            get { return m_Parent; }
        }

        public WorksheetColumnCollection Columns
        {
            get { return m_Columns; }
        }

        public WorksheetRowCollection Rows
        {
            get { return m_Rows; }
        }

        public Double DefaultRowHeight
        {
            get { return m_DefaultRowHeight; }
            set { m_DefaultRowHeight = value; }
        }

        public Double DefaultColumnWidth
        {
            get { return m_DefaultColumnWidth; }
            set { m_DefaultColumnWidth = value; }
        }

        public WorksheetStyle Style
        {
            get { return m_Style; }
            set { m_Style = value; }
        }

        public WorksheetTableCellCollection Cells
        {
            get { return m_Cells; }
        }

        #endregion

        #region Public Methods

        public WorksheetRow NewRow()
        {
            return NewRow(false);
        }

        internal WorksheetRow NewRow(Boolean force)
        {
            WorksheetRow result = null;
			if (!force)
				for (Int32 i = Rows.Count - 1; i >= 0; i--)
				{
					WorksheetRow row = Rows[i];
					if (row.IsEmpty)
						result = row;
					else
						break;
				}

        	if (result == null)
            {
                result = new WorksheetRow(this);
                m_Rows.Add(result);
            }

            return result;
        }

        public WorksheetColumn NewColumn()
        {
            return NewColumn(0);
        }

        public WorksheetColumn NewColumn(Double width)
        {
            return NewColumn(width, false);
        }

        internal WorksheetColumn NewColumn(Boolean force)
        {
            return NewColumn(0, false);
        }

        internal WorksheetColumn NewColumn(Double width, Boolean force)
        {
            WorksheetColumn result = null;
            if (force)
                foreach (WorksheetColumn column in Columns)
                    if (column.IsEmpty)
                    {
                        result = column;
                        break;
                    }

            if (result == null)
            {
                result = new WorksheetColumn(this, width);
                m_Columns.Add(result);
            }

            return result;
        }

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Table");

            int columns = m_Columns.Count;
            int rows = m_Rows.Count;

            writer.WriteAttributeString("ss", "ExpandedColumnCount", null, XmlConvert.ToString(columns));
            writer.WriteAttributeString("ss", "ExpandedRowCount", null, XmlConvert.ToString(rows));
            writer.WriteAttributeString("ss", "FullColumns", null, XmlConvert.ToString(1));
            writer.WriteAttributeString("ss", "FullRows", null, XmlConvert.ToString(1));

            if (m_DefaultRowHeight > 0)
                writer.WriteAttributeString("ss", "DefaultRowHeight", null, XmlConvert.ToString(m_DefaultRowHeight));
            if (m_DefaultColumnWidth > 0)
                writer.WriteAttributeString("ss", "DefaultColumnWidth", null, XmlConvert.ToString(m_DefaultColumnWidth));

            m_Columns.ToXml(writer);
            m_Rows.ToXml(writer);

            writer.WriteEndElement();
        }

        #endregion
    }
}
