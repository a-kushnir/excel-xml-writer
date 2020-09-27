using System;
using System.Drawing;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    internal class FormattedTextItem : IXmlWriter, ICloneable
    {
        #region Fields

        private String m_Text;
        private Boolean? m_Bold;
        private Boolean? m_Italic;
        private Color? m_Color;
        private String m_FontName;
        private String m_FontFamily;
        private Int32? m_CharSet;
        private Double? m_Size;

        #endregion

        #region Public Properties

        public String Text
        {
            get { return m_Text; }
            set { m_Text = value; }
        }

        public Boolean? Bold
        {
            get { return m_Bold; }
            set { m_Bold = value; }
        }

        public Boolean? Italic
        {
            get { return m_Italic; }
            set { m_Italic = value; }
        }

        public Color? Color
        {
            get { return m_Color; }
            set { m_Color = value; }
        }

        public String FontName
        {
            get { return m_FontName; }
            set { m_FontName = value; }
        }

        public String FontFamily
        {
            get { return m_FontFamily; }
            set { m_FontFamily = value; }
        }

        public Int32? CharSet
        {
            get { return m_CharSet; }
            set { m_CharSet = value; }
        }

        public Double? Size
        {
            get { return m_Size; }
            set { m_Size = value; }
        }

        public Boolean IsDefault
        {
            get
            {
                return String.IsNullOrEmpty(Text) || 
                    (Bold == null &&
                    Italic == null &&
                    Color == null && 
                    FontName == null &&
                    FontFamily == null &&
                    CharSet == null &&
                    Size == null);
            }
        }

        #endregion

        #region Initialization

		public FormattedTextItem()
		{
		}

        public FormattedTextItem(String value)
        {
            m_Text = value;
        }

        #endregion

        #region Public Methods

        public FormattedTextItem Combine(FormattedTextItem item)
        {
            FormattedTextItem result = new FormattedTextItem(null);

            result.Bold = Combine(Bold, item.Bold);
            result.Italic = Combine(Italic, item.Italic);
            result.Color = Combine(Color, item.Color);
            result.FontName = Combine(FontName, item.FontName);
            result.FontFamily = Combine(FontFamily, item.FontFamily);
            result.CharSet = Combine(CharSet, item.CharSet);
            result.Size = Combine(Size, item.Size);

            return result;
        }

        public void ToXml(XmlWriter writer)
        {
            if (m_Bold.GetValueOrDefault(false))
                writer.WriteStartElement("B");

            if (m_Italic.GetValueOrDefault(false))
                writer.WriteStartElement("I");

            if (IsDefault || m_Color != null || m_FontName != null || m_FontFamily != null || m_CharSet != null || m_Size != null)
            {
                writer.WriteStartElement("Font");
                if (m_Color != null)
                    writer.WriteAttributeString("html", "Color", null, Helper.XmlConvert(m_Color.Value));
                if (m_FontName != null)
                    writer.WriteAttributeString("html", "Face", null, m_FontName);
                if (m_FontFamily != null)
                    writer.WriteAttributeString("html", "Family", null, m_FontFamily);
                if (m_CharSet != null)
                    writer.WriteAttributeString("html", "CharSet", null, XmlConvert.ToString(m_CharSet.Value));
                if (m_Size != null)
                    writer.WriteAttributeString("html", "Size", null, XmlConvert.ToString(m_Size.Value));
            }

            writer.WriteRaw(Helper.XmlConvert(m_Text));

            if (IsDefault || m_Color != null || m_FontName != null || m_FontFamily != null || m_CharSet != null || m_Size != null)
                writer.WriteEndElement();

            if (m_Italic.GetValueOrDefault(false))
                writer.WriteEndElement();

            if (m_Bold.GetValueOrDefault(false))
                writer.WriteEndElement();
        }

        public object Clone()
        {
            FormattedTextItem clone = new FormattedTextItem(Text);

            clone.Bold = Bold;
            clone.Italic = Italic;
            clone.Color = Color;
            clone.FontName = FontName;
            clone.FontFamily = FontFamily;
            clone.CharSet = CharSet;
            clone.Size = Size;

            return clone;
        }

        #endregion

        #region Private Methods

        public T Combine<T>(T source, T update) where T : class
        {
            return update ?? source;
        }

        public Nullable<T> Combine<T>(Nullable<T> source, Nullable<T> update) where T : struct
        {
            return update ?? source;
        }

        #endregion
    }
}
