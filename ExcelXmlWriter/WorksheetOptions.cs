using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetOptions : IXmlWriter
    {
        #region Fields

        private Worksheet m_Parent;

        private Int32 m_ActivePane;
        private Boolean m_FitToPage;
        private Boolean m_FreezePanes;
        private String m_GridLineColor;
        private Int32 m_LeftColumnRightPane;
        private WorksheetPageSetup m_PageSetup;
        private WorksheetPrintOptions m_Print;
        private Boolean m_ProtectObjects;
        private Boolean m_ProtectScenarios;
        private Boolean m_Selected;
        private Int32 m_SplitHorizontal;
        private Int32 m_SplitVertical;
        private Int32 m_TopRowBottomPane;
        private Int32 m_TopRowVisible;
        private String m_ViewableRange;
        private Boolean m_Visible;

        #endregion

		#region Initialization

		internal WorksheetOptions(Worksheet worksheet)
		{
            m_Parent = worksheet;

			m_PageSetup = new WorksheetPageSetup();
			m_Print = new WorksheetPrintOptions();
		    m_Visible = true;
		}

    	#endregion

		#region Public Properties

		public Int32 ActivePane
        {
            get { return m_ActivePane; }
            set { m_ActivePane = value; }
        }

        public Boolean FitToPage
        {
            get { return m_FitToPage; }
            set { m_FitToPage = value; }
        }

        public Boolean FreezePanes
        {
            get { return m_FreezePanes; }
            set { m_FreezePanes = value; }
        }

        public String GridLineColor
        {
            get { return m_GridLineColor; }
            set { m_GridLineColor = value; }
        }

        public Int32 LeftColumnRightPane
        {
            get { return m_LeftColumnRightPane; }
            set { m_LeftColumnRightPane = value; }
        }

        public WorksheetPageSetup PageSetup
        {
            get { return m_PageSetup; }
            set { m_PageSetup = value; }
        }

        public WorksheetPrintOptions Print
        {
            get { return m_Print; }
            set { m_Print = value; }
        }

        public Boolean ProtectObjects
        {
            get { return m_ProtectObjects; }
            set { m_ProtectObjects = value; }
        }

        public Boolean ProtectScenarios
        {
            get { return m_ProtectScenarios; }
            set { m_ProtectScenarios = value; }
        }

        public Boolean Selected
        {
            get { return m_Selected; }
            set
            {
                if (m_Selected != value)
                {
                    m_Selected = value;

                    Workbook workbook = m_Parent.Workbook;
                    if (value)
                        workbook.ActiveWorksheet = m_Parent;
                    else
                        workbook.ActiveWorksheet = workbook.Worksheets[0];
                }
            }
        }

        public Int32 SplitHorizontal
        {
            get { return m_SplitHorizontal; }
            set { m_SplitHorizontal = value; }
        }

        public Int32 SplitVertical
        {
            get { return m_SplitVertical; }
            set { m_SplitVertical = value; }
        }

        public Int32 TopRowBottomPane
        {
            get { return m_TopRowBottomPane; }
            set { m_TopRowBottomPane = value; }
        }

        public Int32 TopRowVisible
        {
            get { return m_TopRowVisible; }
            set { m_TopRowVisible = value; }
        }

        public String ViewableRange
        {
            get { return m_ViewableRange; }
            set { m_ViewableRange = value; }
        }

        public Boolean Visible
        {
            get { return m_Visible; }
            set { m_Visible = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
        	writer.WriteRaw("<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
			//writer.WriteStartElement("WorksheetOptions", String.Empty);
			//writer.WriteAttributeString("xmlns", "urn:schemas-microsoft-com:office:excel");

            m_PageSetup.ToXml(writer);
            m_Print.ToXml(writer);

            if (!m_Selected)
                writer.WriteElementString("Selected", null);
            if (!m_Visible)
                writer.WriteElementString("Visible", "SheetHidden");

			writer.WriteElementString("ProtectObjects", m_ProtectObjects.ToString());
			writer.WriteElementString("ProtectScenarios", m_ProtectScenarios.ToString());

        	//writer.WriteEndElement();
			writer.WriteRaw("</WorksheetOptions>\n");
        }

        #endregion
    }
}
