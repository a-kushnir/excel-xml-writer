using System;
using System.Drawing;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetStyleFont: IXmlWriter, ICloneable
    {
        #region Fields

        private Boolean m_Bold;
        private Color m_Color;
        private String m_FontName;
        private Boolean m_Italic;
        private Double m_Size;
        private Boolean m_Strikethrough;
        private Boolean m_Underline;
        private Int32 m_CharSet;

        #endregion

        #region Initialization

        public WorksheetStyleFont()
        {
        }

        public WorksheetStyleFont(string fontName)
            : this(fontName, 0, false)
        {
        }

        public WorksheetStyleFont(string fontName, int size)
            : this(fontName, size, false)
        {
        }

        public WorksheetStyleFont(string fontName, int size, bool bold)
        {
            m_FontName = fontName;
            m_Size = size;
            m_Bold = bold;
        }

        public WorksheetStyleFont(string fontName, int size, bool bold, bool italic)
        {
            m_FontName = fontName;
            m_Size = size;
            m_Bold = bold;
            m_Italic = italic;
        }

        #endregion

        #region Public Properties

        public Boolean Bold
        {
            get { return m_Bold; }
            set { m_Bold = value; }
        }

        public Color Color
        {
            get { return m_Color; }
            set { m_Color = value; }
        }

        public String FontName
        {
            get { return m_FontName; }
            set { m_FontName = value; }
        }

        public Boolean Italic
        {
            get { return m_Italic; }
            set { m_Italic = value; }
        }

        public Double Size
        {
            get { return m_Size; }
            set { m_Size = value; }
        }

        public Boolean Strikethrough
        {
            get { return m_Strikethrough; }
            set { m_Strikethrough = value; }
        }

        public Boolean Underline
        {
            get { return m_Underline; }
            set { m_Underline = value; }
        }

        public Int32 CharSet
        {
            get { return m_CharSet; }
            set { m_CharSet = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Font");

            if (!string.IsNullOrEmpty(m_FontName))
                writer.WriteAttributeString("ss", "FontName", null, m_FontName);
            if (m_Color.ToArgb() != 0)
                writer.WriteAttributeString("ss", "Color", null, Helper.XmlConvert(m_Color));
            if (m_CharSet > 0)
                writer.WriteAttributeString("x", "CharSet", null, XmlConvert.ToString(m_CharSet));
            if (m_Bold)
                writer.WriteAttributeString("ss", "Bold", null, Helper.XmlConvert(m_Bold));
            if (m_Italic)
                writer.WriteAttributeString("ss", "Italic", null, Helper.XmlConvert(m_Italic));
            if (m_Underline)
                writer.WriteAttributeString("ss", "Underline", null, Helper.XmlConvert(m_Underline));
            if (m_Size > 0)
                writer.WriteAttributeString("ss", "Size", null, XmlConvert.ToString(m_Size));

            writer.WriteEndElement();
        }

        public object Clone()
        {
            WorksheetStyleFont clone = new WorksheetStyleFont();

            clone.Bold = Bold;
            clone.CharSet = CharSet;
            clone.Color = Color;
            clone.FontName = FontName;
            clone.Italic = Italic;
            clone.Size = Size;
            clone.Strikethrough = Strikethrough;
            clone.Underline = Underline;

            return clone;
        }

        #endregion
    }
}
