using System;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetPrintOptions : IXmlWriter
    {
        #region Fields

        private Boolean m_BlackAndWhite;
        private PrintCommentsLayout m_CommentsLayout;
        private Boolean m_DraftQuality;
        private Int32 m_FitHeight;
        private Int32 m_FitWidth;
        private Boolean m_GridLines;
        private Int32 m_HorizontalResolution;
        private Boolean m_LeftToRight;
        private Int32 m_PaperSizeIndex;
        private PrintErrorsOption m_PrintErrors;
        private Boolean m_RowColHeadings;
        private Int32 m_Scale;
        private Boolean m_ValidPrinterInfo;
        private Int32 m_VerticalResolution;

        #endregion

        #region Public Properties

        public  Boolean BlackAndWhite
        {
            get { return m_BlackAndWhite; }
            set { m_BlackAndWhite = value; }
        }

        public  PrintCommentsLayout CommentsLayout
        {
            get { return m_CommentsLayout; }
            set { m_CommentsLayout = value; }
        }

        public  Boolean DraftQuality
        {
            get { return m_DraftQuality; }
            set { m_DraftQuality = value; }
        }

        public  Int32 FitHeight
        {
            get { return m_FitHeight; }
            set { m_FitHeight = value; }
        }

        public  Int32 FitWidth
        {
            get { return m_FitWidth; }
            set { m_FitWidth = value; }
        }

        public  Boolean GridLines
        {
            get { return m_GridLines; }
            set { m_GridLines = value; }
        }

        public  Int32 HorizontalResolution
        {
            get { return m_HorizontalResolution; }
            set { m_HorizontalResolution = value; }
        }

        public  Boolean LeftToRight
        {
            get { return m_LeftToRight; }
            set { m_LeftToRight = value; }
        }

        public  Int32 PaperSizeIndex
        {
            get { return m_PaperSizeIndex; }
            set { m_PaperSizeIndex = value; }
        }

        public  PrintErrorsOption PrintErrors
        {
            get { return m_PrintErrors; }
            set { m_PrintErrors = value; }
        }

        public  Boolean RowColHeadings
        {
            get { return m_RowColHeadings; }
            set { m_RowColHeadings = value; }
        }

        public  Int32 Scale
        {
            get { return m_Scale; }
            set { m_Scale = value; }
        }

        public  Boolean ValidPrinterInfo
        {
            get { return m_ValidPrinterInfo; }
            set { m_ValidPrinterInfo = value; }
        }

        public  Int32 VerticalResolution
        {
            get { return m_VerticalResolution; }
            set { m_VerticalResolution = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Print");

            if (m_BlackAndWhite)
                writer.WriteElementString("BlackAndWhite", null);
            if (m_DraftQuality)
                writer.WriteElementString("DraftQuality", null);
            if (m_LeftToRight)
                writer.WriteElementString("LeftToRight", null);
            if (m_CommentsLayout != PrintCommentsLayout.None)
                writer.WriteElementString("CommentsLayout", m_CommentsLayout.ToString());
            if (m_ValidPrinterInfo)
                writer.WriteElementString("ValidPrinterInfo", null);
            if (m_Scale != 0)
                writer.WriteElementString("Scale", XmlConvert.ToString(m_Scale));
            if (m_HorizontalResolution != 0)
                writer.WriteElementString("HorizontalResolution", XmlConvert.ToString(m_HorizontalResolution));
            if (m_VerticalResolution != 0)
                writer.WriteElementString("VerticalResolution", XmlConvert.ToString(m_VerticalResolution));
            if (m_GridLines)
                writer.WriteElementString("Gridlines", null);
            if (m_RowColHeadings)
                writer.WriteElementString("RowColHeadings", null);

            writer.WriteEndElement();
        }
        #endregion
    }
}
