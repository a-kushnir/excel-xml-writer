using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class SpreadSheetComponentOptions : IXmlWriter
    {
        #region Fields

        private Boolean m_DoNotEnableResize;
        private Double m_MaxHeight;
        private Double m_MaxWidth;
        private Int32 m_NextSheetNumber;
        private Boolean m_PreventPropBrowser;
        private Boolean m_SpreadsheetAutoFit;
        private readonly SpreadSheetToolbar m_Toolbar;

        #endregion

        #region Initialization

        internal SpreadSheetComponentOptions()
        {
            m_Toolbar = new SpreadSheetToolbar();
        }

        #endregion

        #region Public Properties

        public Boolean DoNotEnableResize
        {
            get { return m_DoNotEnableResize; }
            set { m_DoNotEnableResize = value; }
        }

        /// <summary>
        /// Gets or sets maximal SpreadSheetComponent height.
        /// Note! if value greater than 1.0 then value is described width in pixels, otherwise in percent.
        /// </summary>
        public Double MaxHeight
        {
            get { return m_MaxHeight; }
            set { m_MaxHeight = value; }
        }

        /// <summary>
        /// Gets or sets maximal SpreadSheetComponent width.
        /// Note! if value greater than 1.0 then value is described width in pixels, otherwise in percent.
        /// </summary>
        public Double MaxWidth
        {
            get { return m_MaxWidth; }
            set { m_MaxWidth = value; }
        }

        public Int32 NextSheetNumber
        {
            get { return m_NextSheetNumber; }
            set { m_NextSheetNumber = value; }
        }

        public Boolean PreventPropBrowser
        {
            get { return m_PreventPropBrowser; }
            set { m_PreventPropBrowser = value; }
        }

        public Boolean SpreadsheetAutoFit
        {
            get { return m_SpreadsheetAutoFit; }
            set { m_SpreadsheetAutoFit = value; }
        }

        public SpreadSheetToolbar Toolbar
        {
            get { return m_Toolbar; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            if (m_DoNotEnableResize)
                writer.WriteElementString("DoNotEnableResize", "x", null);

            if (m_MaxHeight > 1.0)
                writer.WriteElementString("MaxHeight", "x", XmlConvert.ToString((Int32)m_MaxHeight));
            else
                writer.WriteElementString("MaxHeight", "x", String.Format("{0}%", XmlConvert.ToString((Int32)(m_MaxHeight * 100))));

            if (m_MaxWidth > 1.0)
                writer.WriteElementString("MaxWidth", "x", XmlConvert.ToString((Int32)m_MaxWidth));
            else
                writer.WriteElementString("MaxWidth", "x", String.Format("{0}%", XmlConvert.ToString((Int32)(m_MaxWidth * 100))));

            if (m_PreventPropBrowser)
                writer.WriteElementString("PreventPropBrowser", "x", null);
            if (m_SpreadsheetAutoFit)
                writer.WriteElementString("SpreadsheetAutoFit", "x", null);

            if (m_NextSheetNumber > 0)
                writer.WriteElementString("NextSheetNumber", "x", XmlConvert.ToString(m_NextSheetNumber));
        }

        #endregion
    }
}
