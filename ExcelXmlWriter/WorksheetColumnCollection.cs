using System;
using System.Collections.Generic;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetColumnCollection : List<WorksheetColumn>, IXmlWriter
    {
        #region Initialization

        internal WorksheetColumnCollection()
        {
        }

        #endregion

        #region Public Methods

        public new void Add(WorksheetColumn column)
        {
			if (Count > Byte.MaxValue)
				throw new OverflowException(String.Format("The column count cannot be more than {0}", Byte.MaxValue));

            column.Index = Count;
            base.Add(column);
        }

        public void ToXml(XmlWriter writer)
        {
            foreach (WorksheetColumn item in this)
                item.ToXml(writer);
        }

        #endregion
    }
}
