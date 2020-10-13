using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class DocumentProperties : IXmlWriter
    {
        #region Fields

        private String m_Title;
        private String m_Subject;
        private String m_Author;
        private String m_Keywords; // note! make it as collection string better
        private String m_Description;
        private String m_LastAuthor;
        private DateTime m_Created;
        private DateTime m_LastSaved;
        private String m_Category;
        private String m_Manager;
        private String m_Company;
        private String m_HyperlinkBase;
        private String m_Version;

        #endregion

        #region Initialization

        internal DocumentProperties()
        {
            m_Author = m_LastAuthor = m_Version = String.Empty;
            m_Created = m_LastSaved = new DateTime();
        }

        #endregion

        #region Public Properties

        public String Title
        {
            get { return m_Title; }
            set { m_Title = value; }
        }

        public String Subject
        {
            get { return m_Subject; }
            set { m_Subject = value; }
        }

        public String Author
        {
            get { return m_Author; }
            set { m_Author = value; }
        }

        public String Keywords
        {
            get { return m_Keywords; }
            set { m_Keywords = value; }
        }

        public String Description
        {
            get { return m_Description; }
            set { m_Description = value; }
        }

        public String LastAuthor
        {
            get { return m_LastAuthor; }
            set { m_LastAuthor = value; }
        }

        public DateTime Created
        {
            get { return m_Created; }
            set { m_Created = value; }
        }

        public DateTime LastSaved
        {
            get { return m_LastSaved; }
            set { m_LastSaved = value; }
        }

        public String Category
        {
            get { return m_Category; }
            set { m_Category = value; }
        }

        public String Manager
        {
            get { return m_Manager; }
            set { m_Manager = value; }
        }

        public String Company
        {
            get { return m_Company; }
            set { m_Company = value; }
        }

        public String HyperlinkBase
        {
            get { return m_HyperlinkBase; }
            set { m_HyperlinkBase = value; }
        }

        public String Version
        {
            get { return m_Version; }
            set { m_Version = value; }            
        }

        #endregion

        #region Public Methods

        public void Create(String author, String version)
        {
            m_Author = author;
            m_LastAuthor = author;
            m_Created = DateTime.Now;
            m_LastSaved = DateTime.Now;
            m_Version = version;
        }

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("DocumentProperties");
            writer.WriteAttributeString("xmlns", "urn", null, "schemas-microsoft-com:office:office");

            writer.WriteElementString("Title", m_Title);
            writer.WriteElementString("Subject", m_Subject);
            writer.WriteElementString("Author", m_Author);
            writer.WriteElementString("Keywords", m_Keywords);
            writer.WriteElementString("Description", m_Description);
            writer.WriteElementString("LastAuthor", m_LastAuthor);
            writer.WriteElementString("Created", XmlConvert.ToString(m_Created, XmlDateTimeSerializationMode.Local));
            writer.WriteElementString("LastSaved", XmlConvert.ToString(m_LastSaved, XmlDateTimeSerializationMode.Local));
            writer.WriteElementString("Category", m_Category);
            writer.WriteElementString("Manager", m_Manager);
            writer.WriteElementString("Company", m_Company);
            writer.WriteElementString("HyperlinkBase", m_HyperlinkBase);
            writer.WriteElementString("Version", m_Version);

            writer.WriteEndElement();
        }

        #endregion
    }
}
