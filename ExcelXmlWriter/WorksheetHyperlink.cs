using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetHyperlink : IXmlWriter 
    {
        #region Fields

        private WorksheetHyperlinkType m_Type;
        private String m_Destination;
        private String m_ScreenTip;

        #endregion

        #region Initialization

        internal WorksheetHyperlink()
        {
            m_Type = WorksheetHyperlinkType.NotSet;
        }

        #endregion

        #region Public Properties

        public WorksheetHyperlinkType Type
        {
            get { return m_Type; }
            set { m_Type = value; }
        }

        public String Destination
        {
            get { return m_Destination; }
            set { m_Destination = value; }
        }

        public String ScreenTip
        {
            get { return m_ScreenTip; }
            set { m_ScreenTip = value; }
        }

        #endregion

        #region Public Methods

        public void SetFile(String fileName)
        {
            SetFile(fileName, String.Empty);
        }

        public void SetFile(String fileName, String screenTip)
        {
            m_Type = WorksheetHyperlinkType.File;
            m_Destination = fileName;
        }

        public void SetUrl(Uri url)
        {
            SetUrl(url, String.Empty);
        }

        public void SetUrl(Uri url, String screenTip)
        {
            m_Type = WorksheetHyperlinkType.Url;
            m_Destination = url.AbsolutePath;
        }

        public void SetEmail(String email)
        {
            SetEmail(email, String.Empty, String.Empty);
        }

        public void SetEmail(String email, String subject)
        {
            SetEmail(email, subject, String.Empty);
        }

        public void SetEmail(String email, String subject, String screenTip)
        {
            m_Type = WorksheetHyperlinkType.Email;
            m_Destination = String.Format("mailto:{0}", email);
            if (!String.IsNullOrEmpty(subject))
                m_Destination += String.Format("?subject={0}", subject);
        }

        public void SetPlace(WorksheetCell cell)
        {
            SetPlace(cell, String.Empty);
        }

        public void SetPlace(WorksheetCell cell, String screenTip)
        {
            if (cell == null)
                throw new ArgumentNullException("cell");

            m_Type = WorksheetHyperlinkType.Place;
            m_Destination = String.Format("#'{0}'!{1}", cell.Table.Worksheet.Name, cell.GetExcelCellAddress());
        }

        public void ToXml(XmlWriter writer)
        {
            if (!String.IsNullOrEmpty(m_Destination))
                writer.WriteAttributeString("ss", "HRef", null, m_Destination);
            if (!String.IsNullOrEmpty(m_ScreenTip))
                writer.WriteAttributeString("ss", "HRefScreenTip", null, m_ScreenTip);
        }

        #endregion
    }
}
