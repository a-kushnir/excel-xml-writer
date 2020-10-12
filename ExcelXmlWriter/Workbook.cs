using System;
using System.IO;
using System.Text;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class Workbook: IXmlWriter
    {
        #region Fields
        
        private Boolean m_GenerateExcelProcessingInstruction;
        private readonly SpreadSheetComponentOptions m_SpreadSheetComponentOptions;
        private readonly ExcelWorkbook m_ExcelWorkbook;
        private readonly DocumentProperties m_Properties;
        private readonly WorksheetStyleCollection m_Styles;
        private readonly WorksheetCollection m_Worksheets;
        private Worksheet m_ActiveWorksheet;

        #endregion

        #region Initialization

        public Workbook()
        {
            m_SpreadSheetComponentOptions = new SpreadSheetComponentOptions();
            m_ExcelWorkbook = new ExcelWorkbook(this);
            m_Properties = new DocumentProperties();
            m_Styles = new WorksheetStyleCollection(this);
            m_Worksheets = new WorksheetCollection(this);
        }

        #endregion

        #region Public Properties

        public Boolean GenerateExcelProcessingInstruction
        {
            get { return m_GenerateExcelProcessingInstruction; }
            set { m_GenerateExcelProcessingInstruction = value; }
        }

        public SpreadSheetComponentOptions SpreadSheetComponentOptions
        {
            get { return m_SpreadSheetComponentOptions; }
        }

        public ExcelWorkbook ExcelWorkbook
        {
            get { return m_ExcelWorkbook; }
        }

        public WorksheetCollection Worksheets
        {
            get { return m_Worksheets; }
        }

        public DocumentProperties Properties
        {
            get { return m_Properties; }
        }

        public WorksheetStyleCollection Styles
        {
            get { return m_Styles; }
        }

        public Worksheet ActiveWorksheet
        {
            get { return m_ActiveWorksheet; }
            set
            {
                if (value == null)
                    throw new ArgumentNullException("value");

                m_ActiveWorksheet = value;
                m_ActiveWorksheet.Options.Selected = true;
            }
        }

        #endregion

        #region Public Methods

		public Worksheet NewWorksheet()
		{
			return NewWorksheet(String.Empty);
		}

        public Worksheet NewWorksheet(String name)
        {
            if (!String.IsNullOrEmpty(name))
                foreach (Worksheet ws in m_Worksheets)
                    if (name == ws.Name)
                        throw new ArgumentException("An worksheet with the same name has already been added.");

			Worksheet worksheet = new Worksheet(this, name);
            m_Worksheets.Add(worksheet);
            return worksheet;
        }

        public Worksheet NewWorksheet(Int32 index, String name)
        {
            if (!String.IsNullOrEmpty(name))
                foreach (Worksheet ws in m_Worksheets)
                    if (name == ws.Name)
                        throw new ArgumentException("An worksheet with the same name has already been added.");

            Worksheet worksheet = new Worksheet(this, name);
            m_Worksheets.Insert(index, worksheet);
            return worksheet;
        }

        public WorksheetStyle NewStyle(String styleID)
        {
            return NewStyle(styleID, new WorksheetStyleFont(), null);
        }

        public WorksheetStyle NewStyle(String styleID, WorksheetStyle parent)
        {
            return NewStyle(styleID, new WorksheetStyleFont(), parent);
        }

        public WorksheetStyle NewStyle(String styleID, WorksheetStyleFont font)
        {
            return NewStyle(styleID, font, null);
        }

        public WorksheetStyle NewStyle(String styleID, WorksheetStyleFont font, WorksheetStyle parent)
        {
            WorksheetStyle style = new WorksheetStyle(this, styleID, font, parent);
            m_Styles.Add(style);
            return style;
        }

        public void Save(String fileName)
        {
        	String directory = Path.GetDirectoryName(fileName);
			if (!String.IsNullOrEmpty(directory) && !Directory.Exists(directory))
				Directory.CreateDirectory(directory);

            using (FileStream stream = new FileStream(fileName, FileMode.Create))
                WriteToStream(stream);
        }

        public void Save(Stream stream)
        {
            WriteToStream(stream);
        }

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartDocument();
            writer.WriteProcessingInstruction("mso-application", "progid=\"Excel.Sheet\"");
            writer.WriteStartElement("Workbook", "urn:schemas-microsoft-com:office:spreadsheet");
            writer.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
            writer.WriteAttributeString("xmlns", "x", null, "urn:schemas-microsoft-com:office:excel");
            writer.WriteAttributeString("xmlns", "ss", null, "urn:schemas-microsoft-com:office:spreadsheet");
            writer.WriteAttributeString("xmlns", "html", null, "http://www.w3.org/TR/REC-html40");

            m_Properties.ToXml(writer);
            m_ExcelWorkbook.ToXml(writer);

            m_Styles.ToXml(writer);
            m_Worksheets.ToXml(writer);

            m_SpreadSheetComponentOptions.ToXml(writer);

            writer.WriteEndElement();
            writer.Flush();
        }

        #endregion

        #region Private Methods

        private void WriteToStream(Stream stream)
        {
            XmlWriterSettings xmlWriterSettings = new XmlWriterSettings();
            xmlWriterSettings.Encoding = Encoding.UTF8;
            xmlWriterSettings.Indent = true;
            xmlWriterSettings.IndentChars = String.Empty;
            xmlWriterSettings.CloseOutput = false;

            using (XmlWriter writer = XmlTextWriter.Create(stream, xmlWriterSettings))
                ToXml(writer);
        }

        #endregion
    }
}
