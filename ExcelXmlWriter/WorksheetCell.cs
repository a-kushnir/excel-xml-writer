using System;
using System.Drawing;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetCell : IXmlWriter
    {
        #region Fields

        private WorksheetComment m_Comment;
        private WorksheetCellData m_Data;
        private String m_Formula;
        private WorksheetHyperlink m_Hyperlink;
        private Int32 m_MergeAcross;
        private Int32 m_MergeDown;
        private WorksheetStyle m_Style;

        private WorksheetTable m_Table;
        private WorksheetColumn m_Column;
        private WorksheetRow m_Row;

        private readonly Boolean m_Initialized;
        private WorksheetCell m_Referer;

        #endregion

        #region Initialization

        private WorksheetCell(WorksheetTable table, WorksheetColumn column, WorksheetRow row, Boolean initialized)
        {
            if (table == null)
                throw new ArgumentNullException("table");
            if (row == null)
                throw new ArgumentNullException("row");

            m_Table = table;
            m_Column = column;
            m_Row = row;

            m_Hyperlink = new WorksheetHyperlink();
            m_Comment = new WorksheetComment();

            m_Initialized = initialized;
        }

        internal WorksheetCell(WorksheetTable table, WorksheetColumn column, WorksheetRow row, WorksheetCell referer)
            : this(table, column, row, false)
        {
            m_Referer = referer;
            m_Data = new WorksheetCellData(null, DataType.String);
        }

        internal WorksheetCell(WorksheetTable table, WorksheetColumn column, WorksheetRow row, String text, DataType type, WorksheetStyle style)
            : this(table, column, row, true)
        {
            m_Style = style;
            m_Data = new WorksheetCellData(text, type);
        }

        #endregion

        #region Public Properties

        public WorksheetComment Comment
        {
            get { return m_Comment; }
            set { m_Comment = value; }
        }

        public WorksheetCellData Data
        {
            get { return m_Data; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                m_Data = value;
            }
        }

        public String Formula
        {
            get { return m_Formula; }
            set { m_Formula = value; }
        }

        public WorksheetHyperlink Hyperlink
        {
            get { return m_Hyperlink; }
            set { m_Hyperlink = value; }
        }

        public Int32 MergeAcross
        {
            get { return m_MergeAcross; }
            set { ResizeMergeArea(m_MergeDown, value); }
        }

        public Int32 MergeDown
        {
            get { return m_MergeDown; }
            set { ResizeMergeArea(value, m_MergeAcross); }
        }

        public WorksheetStyle Style
        {
            get { return m_Style; }
            set { m_Style = value; }
        }

        public WorksheetTable Table
        {
            get { return m_Table; }
            internal set { m_Table = value; }
        }

        public WorksheetColumn Column
        {
            get { return m_Column; }
            internal set { m_Column = value; }
        }

        public WorksheetRow Row
        {
            get { return m_Row; }
            internal set { m_Row = value; }
        }

        internal Boolean Initialized
        {
            get { return m_Initialized; }
        }

        internal WorksheetCell Referer 
        {
            get { return m_Referer; }
            private set
            {
                if (m_Referer != null && m_Referer != value)
                    throw new ExcelXmlWriterException(String.Format(
                        "Cell intersection is forbidden. Cell[{0},{1}] already occupied by Cell[{2},{3}] (MergeAcross={4}, MergeDown={5})",
                        m_Row.Index, m_Column.Index, m_Referer.Row.Index, m_Referer.Column.Index, m_Referer.MergeAcross, m_Referer.MergeDown));

                m_Referer = value;
            }
        }

        #endregion

        #region Public Methods

		/// <summary>
		/// Sets the style of specified number of characters beginning at specified position to bold.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
    	public void SetBold(Int32 position, Int32 length)
    	{
    		FormattedTextItem item = new FormattedTextItem();
    		item.Bold = true;
    		ChangeFormatting(position, length, item);
    	}

		/// <summary>
		/// Sets the style of specified number of characters beginning at specified position to italic.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		public void SetItalic(Int32 position, Int32 length)
		{
			FormattedTextItem item = new FormattedTextItem();
			item.Italic = true;
			ChangeFormatting(position, length, item);
		}

		/// <summary>
		/// Sets the style of specified number of characters beginning at specified position to color.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		/// <param name="color">The font color to apply.</param>
		public void SetColor(Int32 position, Int32 length, Color color)
		{
			FormattedTextItem item = new FormattedTextItem();
			item.Color = color;
			ChangeFormatting(position, length, item);
		}

		/// <summary>
		/// Sets the font style of specified number of characters beginning at specified position.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		/// <param name="fontName">The font name to apply.</param>
		/// <param name="size">The font size to apply.</param>
		public void SetFont(Int32 position, Int32 length, String fontName, Double? size)
		{
			SetFont(position, length, fontName, null, size, null);
		}

		/// <summary>
		/// Sets the font style of specified number of characters beginning at specified position.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		/// <param name="fontName">The font name to apply.</param>
		/// <param name="fontFamily">The font family to apply.</param>
		/// <param name="size">The font size to apply.</param>
		/// <param name="charSet">The font char set to apply.</param>
    	public void SetFont(Int32 position, Int32 length, String fontName, String fontFamily, Double? size, Int32? charSet)
		{
			FormattedTextItem item = new FormattedTextItem();
    		item.FontName = fontName;
			item.FontFamily = fontFamily;
			item.Size = size;
			item.CharSet = charSet;
			ChangeFormatting(position, length, item);
		}

		/// <summary>
		/// Sets the font style of specified number of characters beginning at specified position.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		/// <param name="font">The font to apply.</param>
		public void SetFont(Int32 position, Int32 length, WorksheetStyleFont font)
		{
			FormattedTextItem item = new FormattedTextItem();
			item.Bold = font.Bold;
			item.Italic = font.Italic;
			item.Color = font.Color;
			item.FontName = font.FontName;
			item.Size = font.Size;
			ChangeFormatting(position, length, item);
		}

		/// <summary>
		/// Sets the style of specified number of characters beginning at specified position.
		/// </summary>
		/// <param name="position">The starting position to apply the style.</param>
		/// <param name="length">The number of characters to apply the style.</param>
		/// <param name="style">The style to apply.</param>
		public void SetStyle(Int32 position, Int32 length, WorksheetStyle style)
		{
			SetFont(position, length, style.Font);
		}

        public void ToXml(XmlWriter writer)
        {
            if (m_Referer != null)
                return;

            writer.WriteStartElement("Cell");

            if (HasMissingCellsBefore())
                writer.WriteAttributeString("ss", "Index", null, XmlConvert.ToString(m_Column.Index + 1));

            if (!String.IsNullOrEmpty(m_Formula))
                writer.WriteAttributeString("ss", "Formula", null, m_Formula);

            if (m_MergeAcross > 0)
                writer.WriteAttributeString("ss", "MergeAcross", null, XmlConvert.ToString(m_MergeAcross));
            if (m_MergeDown > 0)
                writer.WriteAttributeString("ss", "MergeDown", null, XmlConvert.ToString(m_MergeDown));
            if (m_Style != null)
                writer.WriteAttributeString("ss", "StyleID", null, m_Style.StyleID);

            m_Comment.ToXml(writer);
            m_Hyperlink.ToXml(writer);
            m_Data.ToXml(writer);

            writer.WriteEndElement();
        }

        #endregion

        #region Internal Methods

        internal String GetExcelCellAddress()
        {
            /*
            String result = String.Empty;

            const Int32 EnglishLettersCount = 26;
             
            Int32 first = m_Column.Index / EnglishLettersCount;
            Int32 last = m_Column.Index - first * EnglishLettersCount;

            if (first > 0)
                result += (Char)('A' + first);
            result += (Char)('A' + last);
            result += m_Row.Index + 1;
            
            return result;
             */

            return String.Format("R{0}C{1}", m_Row.Index + 1, m_Column.Index + 1);
        }

        #endregion

        #region Private Methods

        private Boolean HasMissingCellsBefore()
        {
            Boolean result = false;

            foreach(WorksheetCell cell in m_Row.Cells)
            {
                if (cell == this)
                    break;

                result = cell.Referer != null;
                if (result)
                    break;
            }

            return result;
        }

		private void ChangeFormatting(Int32 position, Int32 length, FormattedTextItem item)
		{
			try
			{
				m_Data.FormattedText.Update(position, length, item);
			}
			catch (Exception ex)
			{
				throw new ExcelXmlWriterException("The specified text format cannot be applied.", ex);
			}
		}

        private void ResizeMergeArea(Int32 height, Int32 width)
        {
            Int32 maxHeight = Max(height, m_MergeDown);
            Int32 maxWidth = Max(width, m_MergeAcross);

            try
            {
                for (Int32 row = m_Row.Index; row <= m_Row.Index + maxHeight; row++)
                    for (Int32 column = m_Column.Index; column <= m_Column.Index + maxWidth; column++)
                    {
                        // Excluding itself
                        if (row == m_Row.Index && column == m_Column.Index)
                            continue;

                        // Expand table if needed
                        if (row == m_Table.Rows.Count)
                            m_Table.NewRow(true);
                        if (column == m_Table.Columns.Count)
                            m_Table.NewColumn(true);

                        // Set cell referer
                        if (row > m_Row.Index + height || column > m_Column.Index + width)
                            m_Table.Cells.GetShadowyCell(row, column).Referer = null;
                        else
                            m_Table.Cells.GetShadowyCell(row, column).Referer = this;
                    }

                m_MergeDown = height;
                m_MergeAcross = width;
            }
            catch (Exception ex)
            {
                RollbackMergeArea(height, width);
                throw new ExcelXmlWriterException("The cell cannot be merged.", ex);
            }
        }

        private void RollbackMergeArea(Int32 height, Int32 width)
        {
            Int32 maxHeight = Max(height, m_MergeDown);
            Int32 maxWidth = Max(width, m_MergeAcross);

            for (Int32 row = m_Row.Index; row <= m_Row.Index + maxHeight; row++)
                for (Int32 column = m_Column.Index; column <= m_Column.Index + maxWidth; column++)
                {
                    // Excluding of itself
                    if (row == m_Row.Index && column == m_Column.Index)
                        continue;

                    // Set cell referer
                    if (row < m_Table.Rows.Count && column < m_Table.Columns.Count)
                        if (row > m_Row.Index + m_MergeDown || column > m_Column.Index + m_MergeAcross)
                        {
                            WorksheetCell cell = m_Table.Cells.GetShadowyCell(row, column);
                            if (cell == this) 
                                cell.Referer = null;
                        }
                        else
                            m_Table.Cells.GetShadowyCell(row, column).Referer = this;
                }
        }

        private static Int32 Max(Int32 value1, Int32 value2)
        {
            return value1 > value2 ? value1 : value2;
        }

        #endregion
    }
}
