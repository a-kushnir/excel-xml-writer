using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class ExcelWorkbook
    {
        #region Fields

        private int m_WindowHeight;
        private int m_WindowWidth;
        private int m_WindowTopX;
        private int m_WindowTopY;
        private bool m_ProtectStructure;
        private bool m_ProtectWindows;

        #endregion

        #region Public Properties

        public int WindowHeight
        {
            get { return m_WindowHeight; }
            set { m_WindowHeight = value; }
        }

        public int WindowWidth
        {
            get { return m_WindowWidth; }
            set { m_WindowWidth = value; }
        }

        public int WindowTopX
        {
            get { return m_WindowTopX; }
            set { m_WindowTopX = value; }
        }

        public int WindowTopY
        {
            get { return m_WindowTopY; }
            set { m_WindowTopY = value; }
        }

        public bool ProtectStructure
        {
            get { return m_ProtectStructure; }
            set { m_ProtectStructure = value; }
        }

        public bool ProtectWindows
        {
            get { return m_ProtectWindows; }
            set { m_ProtectWindows = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("ExcelWorkbook");

            if (m_WindowHeight != 0)
                writer.WriteElementString("WindowHeight", XmlConvert.ToString(m_WindowHeight));
            if (m_WindowWidth != 0)
                writer.WriteElementString("WindowWidth", XmlConvert.ToString(m_WindowWidth));
            if (m_WindowTopX != 0)
                writer.WriteElementString("WindowTopX", XmlConvert.ToString(m_WindowTopX));
            if (m_WindowTopY != 0)
                writer.WriteElementString("WindowTopY", XmlConvert.ToString(m_WindowTopY));

            writer.WriteElementString("ProtectStructure", m_ProtectStructure ? "True" : "False");
            writer.WriteElementString("ProtectWindows", m_ProtectWindows ? "True" : "False");

            writer.WriteEndElement();
        }

        #endregion
    }
}
