using System;
using System.Drawing;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetComment : IXmlWriter
    {
        #region Fields

        private string m_Author;
        private readonly WorksheetCellData m_Data;
        private bool m_ShowAlways;

        #endregion

        #region Initialization

        internal WorksheetComment()
        {
            m_Data = new WorksheetCellData(String.Empty, DataType.NotSet);
        }

        #endregion

        #region Public Methods

        public string Author
        {
            get { return m_Author; }
            set { m_Author = value; }
        }

        public WorksheetCellData Data
        {
            get { return m_Data; }
        }

        public bool ShowAlways
        {
            get { return m_ShowAlways; }
            set { m_ShowAlways = value; }
        }

        #endregion

        #region Public Methods

        public void Create(String author, String comment)
        {
            Create(author, comment, false);
        }

        public void Create(String author, String comment, Boolean showAlways)
        {
            if (author == null)
                throw new ArgumentNullException("author");
            if (comment == null)
                throw new ArgumentNullException("comment");

            m_Author = author;
            m_Data.Text = author + Environment.NewLine + comment;
            m_ShowAlways =  showAlways;

            FormattedTextItem format = new FormattedTextItem();
            format.FontName = "Tahoma";
            format.Size = 8;
            format.Color = Color.Black;
            m_Data.FormattedText.Update(0, m_Data.Text.Length, format);

            format.Bold = true;
            m_Data.FormattedText.Update(0, author.Length, format);
        }

        public void ToXml(XmlWriter writer)
        {
            if (!String.IsNullOrEmpty(m_Author) ||
                !String.IsNullOrEmpty(m_Data.Text))
            {
                writer.WriteStartElement("Comment");

                if (!String.IsNullOrEmpty(m_Author))
                    writer.WriteAttributeString("ss", "Author", null, m_Author);
                if (m_ShowAlways)
                    writer.WriteAttributeString("ss", "ShowAlways", null, Helper.XmlConvert(m_ShowAlways));

                m_Data.ToXml(writer);
                writer.WriteEndElement();
            }
        }

        #endregion
    }
}
