using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetCellData: IXmlWriter
    {
        #region Fields

        private String m_Text;
        private DataType m_Type;
        private FormattedText m_FormattedText;

        #endregion

        #region Initialization

        internal WorksheetCellData(String text, DataType type)
        {
            Text = text;
            Type = type;
        }

        #endregion

        #region Public Properties

        public String Text
        {
            get { return m_Text; }
            set
            {
                if (m_Text != value)
                {
                    m_Text = value;
                    m_FormattedText = new FormattedText(value);
                }
            }
        }

        public DataType Type
        {
            get { return m_Type; }
            set { m_Type = value; }
        }

        internal FormattedText FormattedText
        {
            get { return m_FormattedText; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            if (m_FormattedText == null || m_FormattedText.IsDefault)
                writer.WriteStartElement("Data");
            else
                writer.WriteStartElement("ss", "Data", null);

            if (m_Type != DataType.NotSet)
                writer.WriteAttributeString("ss", "Type", null, m_Type.ToString());

            if (m_FormattedText == null || m_FormattedText.IsDefault)
            {
                writer.WriteRaw(Helper.XmlConvert(m_Text));
            }
            else
            {
                writer.WriteAttributeString("xmlns", "http://www.w3.org/TR/REC-html40");
                m_FormattedText.ToXml(writer);
            }

            writer.WriteEndElement();
        }

        public static implicit operator WorksheetCellData(Int16 value)
        {
            return new WorksheetCellData(XmlConvert.ToString(value), DataType.Integer);
        }

        public static implicit operator WorksheetCellData(Int32 value)
        {
            return new WorksheetCellData(XmlConvert.ToString(value), DataType.Integer);
        }

        public static implicit operator WorksheetCellData(Single value)
        {
            return new WorksheetCellData(XmlConvert.ToString(value), DataType.Number);
        }

        public static implicit operator WorksheetCellData(Double value)
        {
            return new WorksheetCellData(XmlConvert.ToString(value), DataType.Number);
        }

        public static implicit operator WorksheetCellData(DateTime value)
        {
            return new WorksheetCellData(XmlConvert.ToString(value, XmlDateTimeSerializationMode.Utc), DataType.DateTime);
        }

        public static implicit operator WorksheetCellData(String value)
        {
            return new WorksheetCellData(value, DataType.String);
        }

        #endregion
    }
}
