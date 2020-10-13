using System;
using System.Drawing;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetStyleBorder: IXmlWriter, ICloneable
    {
        #region Fields

        private Color m_Color;
        private LineStyleOption m_LineStyle;
        private StylePosition m_Position;
        private Int32 m_Weight;

        #endregion

        #region Initialization

        internal WorksheetStyleBorder()
        {
            m_LineStyle = LineStyleOption.NotSet;
            m_Position = StylePosition.NotSet;
        }

        internal WorksheetStyleBorder(StylePosition position, LineStyleOption lineStyle, Int32 weight, Color color)
        {
            m_LineStyle = lineStyle;
            m_Position = position;
            m_Weight = weight;
            m_Color = color;
        }

        #endregion

        #region Public Properties

        public Color Color
        {
            get { return m_Color; }
            set { m_Color = value; }
        }

        public LineStyleOption LineStyle
        {
            get { return m_LineStyle; }
            set { m_LineStyle = value; }
        }

        public StylePosition Position
        {
            get { return m_Position; }
            set { m_Position = value; }
        }

        public Int32 Weight
        {
            get { return m_Weight; }
            set { m_Weight = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Border");

            writer.WriteAttributeString("ss", "Color", null, Helper.XmlConvert(m_Color));
            if (m_LineStyle != LineStyleOption.NotSet)
                writer.WriteAttributeString("ss", "LineStyle", null, m_LineStyle.ToString());
            if (m_Position != StylePosition.NotSet)
                writer.WriteAttributeString("ss", "Position", null, m_Position.ToString());
            writer.WriteAttributeString("ss", "Weight", null, XmlConvert.ToString(m_Weight));

            writer.WriteEndElement();
        }

        public object Clone()
        {
            WorksheetStyleBorder clone = new WorksheetStyleBorder();

            clone.Color = Color;
            clone.LineStyle = LineStyle;
            clone.Position = Position;
            clone.Weight = Weight;

            return clone;
        }

        #endregion
    }
}
