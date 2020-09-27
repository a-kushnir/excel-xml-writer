using System;
using System.Collections.Generic;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetStyleCollection : Dictionary<String, WorksheetStyle>, IXmlWriter
    {
        #region Fields

        private readonly Workbook m_Parent;

        #endregion

        #region Initialization

        internal WorksheetStyleCollection(Workbook parent)
        {
            m_Parent = parent;
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Styles");

            foreach (WorksheetStyle item in Values)
                item.ToXml(writer);

            writer.WriteEndElement();
        }

        internal void Add(WorksheetStyle style)
        {
            if (String.IsNullOrEmpty(style.StyleID))
                style.StyleID = GenerateStyleID();

            Add(style.Name, style);
        }

        #endregion

        #region Private Methods

        private String GenerateStyleID()
        {
            bool exists = false;
            int index = 0;
            string result;
            do
            {
                result = string.Format("s{0}", index);
                foreach (String name in Keys)
                {
                    exists = name == result;
                    if (exists)
                        break;
                }
                index++;
            } while (exists);

            return result;
        }

        #endregion
    }
}
