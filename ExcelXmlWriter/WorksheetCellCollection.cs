using System;
using System.Collections.Generic;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetCellCollection: List<WorksheetCell>, IXmlWriter
    {
        #region Fields

        private readonly WorksheetRow m_Row;

        #endregion

        #region Initialization

        internal WorksheetCellCollection(WorksheetRow row)
        {
            m_Row = row;
        }

        #endregion

        #region Public Methods

        public new void Add(WorksheetCell cell)
        {
            if (cell.Column != null)
            {
                base.Add(cell);
            }
            else
            {
                Int32 columnIndex = GetColumnIndex();
                cell.Column = m_Row.Table.Columns[columnIndex];
                if (columnIndex >= 0 && columnIndex < Count)
                    this[GetColumnIndex()] = cell;
                else
                    base.Add(cell);
            }
        }

        public void ToXml(XmlWriter writer)
        {
            foreach (WorksheetCell item in this)
                item.ToXml(writer);
        }

        /// <summary>
        /// Gets or sets the cell with specified column.
        /// </summary>
        /// <param name="column">The zero-based index of column.</param>
        /// <returns>The cell with specified column.</returns>
        public new WorksheetCell this[Int32 column]
        {
            get
            {
                if (column < 0 || column >= Count)
                    throw new ArgumentOutOfRangeException("column");

                WorksheetCell result = base[column];
                return result.Referer ?? result;
            }
            set { base[column] = value; }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Returns the cell with specified column ignoring Referer property.
        /// Note: For internal use only.
        /// </summary>
        /// <param name="column">The zero-based index of column.</param>
        /// <returns>The cell with specified column.</returns>
        internal WorksheetCell GetShadowyCell(Int32 column)
        {
            return base[column];
        }

        #endregion

        #region Private Methods

        private Int32 GetColumnIndex()
        {
            Int32 result = -1;

            foreach(WorksheetCell cell in this)
                if (cell.Referer == null && !cell.Initialized)
                {
                    result = cell.Column.Index;
                    break;
                }

            if (result < 0)
                result = m_Row.Table.NewColumn(true).Index;

            return result;
        }

        #endregion
    }
    
}
