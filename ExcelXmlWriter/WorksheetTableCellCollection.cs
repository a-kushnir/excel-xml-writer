using System;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetTableCellCollection
    {
        #region Fields

        private readonly WorksheetRowCollection m_Rows;

        #endregion

        #region Initialization

        internal WorksheetTableCellCollection(WorksheetRowCollection rows)
        {
            m_Rows = rows;
        }

        #endregion

        #region Public Methods

        /// <summary>
        /// Returns the cell collection of specified row.
        /// </summary>
        /// <param name="row">The zero-based index of row.</param>
        /// <returns>The cell collection of specified row.</returns>
        public WorksheetCellCollection this[Int32 row]
        {
            get
            {
                if (row < 0 || row >= m_Rows.Count)
                    throw new ArgumentOutOfRangeException("row");

                return m_Rows[row].Cells;
            }
        }

        #endregion

        #region Internal Methods

        /// <summary>
        /// Returns the cell with specified row and column ignoring Referer property.
        /// Note: For internal use only.
        /// </summary>
        /// <param name="row">The zero-based index of row.</param>
        /// <param name="column">The zero-based index of column.</param>
        /// <returns>The cell with specified row and column.</returns>
        internal WorksheetCell GetShadowyCell(Int32 row, Int32 column)
        {
            return m_Rows[row].Cells.GetShadowyCell(column);
        }

        #endregion
    }
}
