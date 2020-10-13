using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetPageSetup : IXmlWriter
    {
        #region Fields

        private readonly WorksheetPageHeader m_Header;
		private readonly WorksheetPageFooter m_Footer;
		private readonly WorksheetPageLayout m_Layout;
		private readonly WorksheetPageMargins m_PageMargins;

        #endregion

		#region Initialization

		internal WorksheetPageSetup()
		{
			m_Header = new WorksheetPageHeader();
			m_Footer = new WorksheetPageFooter();
			m_Layout = new WorksheetPageLayout();
			m_PageMargins = new WorksheetPageMargins();
		}

    	#endregion

		#region Public Properties

		public WorksheetPageHeader Header
        {
            get { return m_Header; }
        }

        public WorksheetPageFooter Footer
        {
            get { return m_Footer; }
        }

        public WorksheetPageLayout Layout
        {
            get { return m_Layout; }
        }

        public WorksheetPageMargins PageMargins
        {
            get { return m_PageMargins; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
			writer.WriteStartElement("PageSetup");

            m_Layout.ToXml(writer);
            m_Header.ToXml(writer);
            m_Footer.ToXml(writer);
            m_PageMargins.ToXml(writer);

			writer.WriteEndElement();
        }

        #endregion
    }
}
