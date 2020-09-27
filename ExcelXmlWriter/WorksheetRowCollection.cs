using System;
using System.Collections.Generic;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetRowCollection : List<WorksheetRow>, IXmlWriter
    {
        #region Initialization

        internal WorksheetRowCollection()
        {
        }

        #endregion

        #region Public Methods

        public new void Add(WorksheetRow row)
        {
			if (Count > UInt16.MaxValue)
				throw new OverflowException(String.Format("The row count cannot be more than {0}", UInt16.MaxValue));

            row.Index = Count;
            base.Add(row);
        }

        public void ToXml(XmlWriter writer)
        {
            foreach (WorksheetRow item in this)
                item.ToXml(writer);
        }

        #endregion
    }
}
