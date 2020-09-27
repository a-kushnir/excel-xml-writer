using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class Worksheet: IXmlWriter
    {
        #region Fields

        private String m_Name;
        private readonly WorksheetTable m_Table;
        private readonly WorksheetOptions m_Options;
        private readonly Workbook m_Parent;
        private readonly WorksheetAutoFilter m_AutoFilter;
        private Boolean m_Protected;

        /*
        public .ExcelXmlWriter.WorksheetNamedRangeCollection Names { get; }
        public .ExcelXmlWriter.PivotTable PivotTable { get; }
        public System.Collections.Specialized.StringCollection Sorting { get; }
        */

        #endregion

        #region Initialization

        internal Worksheet(Workbook workbook, String name)
        {
            m_Parent = workbook;
            m_Name = name;
            m_Table = new WorksheetTable(this);
            m_Options = new WorksheetOptions(this);
            m_AutoFilter = new WorksheetAutoFilter();
        }

        #endregion

        #region Public Properties

        public WorksheetTable Table
        {
            get { return m_Table; }
        }

        public String Name
        {
            get { return m_Name; }
            set { m_Name = value; }
        }

        public Workbook Workbook
        {
            get { return m_Parent; }
        }

        public WorksheetOptions Options
        {
            get { return m_Options; }
        }

        public WorksheetAutoFilter AutoFilter
        {
            get { return m_AutoFilter; }
        }

        public Boolean Protected // ?
        {
            get { return m_Protected; }
            set { m_Protected = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Worksheet");
			writer.WriteAttributeString("ss", "Name", null, GetWorksheetName());

            m_Table.ToXml(writer);
            m_Options.ToXml(writer);

            writer.WriteEndElement();
        }

        #endregion

		#region Private Methods

	    private String GetWorksheetName()
	    {
		    String result = m_Name;

		    if (String.IsNullOrEmpty(result))
		    {
			    Int32 index = 1;
			    Boolean alreadyExist = false;
			    do
			    {
				    result = String.Format("Workbook{0}", index);
				    foreach (Worksheet worksheet in m_Parent.Worksheets)
				    {
					    if (worksheet.Name == result)
					    {
						    alreadyExist = true;
						    index++;
						    break;
					    }
				    }
    				
			    } while (alreadyExist);
		    }

		    return result;
	    }

	    #endregion
	}
}
