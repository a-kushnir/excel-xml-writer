using System;
using System.Drawing;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetStyleInterior: IXmlWriter, ICloneable
    {
        #region Fields

        private Color m_Color;
        private StyleInteriorPattern m_Pattern;

        #endregion

        #region Initialization

        public WorksheetStyleInterior()
        {
        }

        public WorksheetStyleInterior(Color ñolor)
            : this(ñolor, StyleInteriorPattern.Solid)
        {
        }

        public WorksheetStyleInterior(StyleInteriorPattern pattern)
            : this(new Color(), pattern)
        {
        }

        public WorksheetStyleInterior(Color ñolor, StyleInteriorPattern pattern)
        {
            m_Color = ñolor;
            m_Pattern = pattern;
        }

        #endregion

        #region Public Properties

        public Color Color
        {
            get { return m_Color; }
            set
            {
                m_Color = value;
                if (m_Pattern == StyleInteriorPattern.NotSet)
                    m_Pattern = StyleInteriorPattern.Solid;
            }
        }

        public StyleInteriorPattern Pattern
        {
            get { return m_Pattern; }
            set { m_Pattern = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Interior");
            if (m_Color.ToArgb() != 0)
                writer.WriteAttributeString("ss", "Color", null, Helper.XmlConvert(m_Color));
            if (m_Pattern != StyleInteriorPattern.NotSet)
                writer.WriteAttributeString("ss", "Pattern", null, m_Pattern.ToString());
            writer.WriteEndElement();
        }

        public object Clone()
        {
            WorksheetStyleInterior clone = new WorksheetStyleInterior();

            clone.Color = Color;
            clone.Pattern = Pattern;

            return clone;
        }
        #endregion
    }
}
