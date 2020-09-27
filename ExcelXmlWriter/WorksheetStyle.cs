using System;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    public class WorksheetStyle : IXmlWriter
    {
        #region Fields

        private String m_StyleID;
        private WorksheetStyleFont m_Font;
        private String m_Name;
        private WorksheetStyle m_ParentStyle;
        private Workbook m_Parent;
        private WorksheetStyleInterior m_Interior;
        private WorksheetStyleAlignment m_Aligment;
        private WorksheetStyleBorderCollection m_Borders;

        #endregion

        #region Initialization

        internal WorksheetStyle()
        {
        }

        internal WorksheetStyle(Workbook parent, String styleID, WorksheetStyleFont font, WorksheetStyle parentStyle)
        {
            if (parentStyle != null)
            {
                m_Font = font ?? (WorksheetStyleFont) parentStyle.Font.Clone();
                m_Interior = (WorksheetStyleInterior)parentStyle.Interior.Clone();
                m_Aligment = (WorksheetStyleAlignment)parentStyle.Aligment.Clone();
                m_Borders = (WorksheetStyleBorderCollection) parentStyle.Borders.Clone();
            }
            else
            {
                m_Font = new WorksheetStyleFont();
                m_Interior = new WorksheetStyleInterior();
                m_Aligment = new WorksheetStyleAlignment();
                m_Borders = new WorksheetStyleBorderCollection();
            }
            m_ParentStyle = parentStyle;
            m_Parent = parent;
            m_Name = styleID;
            m_StyleID = styleID;
            m_Font = font;
        }

        #endregion

        #region Public Properties

        public String Name
        {
            get { return m_Name; }
            set { m_Name = value; }
        }

        public WorksheetStyleFont Font
        {
            get { return m_Font; }
            set { m_Font = value; }
        }

        public String StyleID
        {
            get { return m_StyleID; }
            set { m_StyleID = value; }
        }

        public WorksheetStyle Parent
        {
            get { return m_ParentStyle; }
            internal set { m_ParentStyle = value; }
        }

        public Workbook Workbook
        {
            get { return m_Parent; }
            internal set { m_Parent = value; }
        }

        public WorksheetStyleInterior Interior
        {
            get { return m_Interior; }
            set { m_Interior = value; }
        }

        public WorksheetStyleAlignment Aligment
        {
            get { return m_Aligment; }
            set { m_Aligment = value; }            
        }

        public WorksheetStyleBorderCollection Borders
        {
            get { return m_Borders; }
            internal set { m_Borders = value; }
        }

        #endregion

        #region Public Methods

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Style");

            writer.WriteAttributeString("ss", "ID", null, m_StyleID);
            if (!String.IsNullOrEmpty(m_Name))
                writer.WriteAttributeString("ss", "Name", null, m_Name);
            if (m_ParentStyle != null && !String.IsNullOrEmpty(m_ParentStyle.StyleID))
                writer.WriteAttributeString("ss", "Parent", null, m_ParentStyle.StyleID);

            if (m_Font != null)
                m_Font.ToXml(writer);
            if (m_Interior != null)
                m_Interior.ToXml(writer);
            if (m_Aligment != null)
                m_Aligment.ToXml(writer);
            if (m_Borders != null)
                m_Borders.ToXml(writer);

            writer.WriteEndElement();
        }

        /*
        public object Clone()
        {
            WorksheetStyle clone = new WorksheetStyle();

            clone.Aligment = (WorksheetStyleAlignment) Aligment.Clone();
            clone.Borders = (WorksheetStyleBorderCollection) Borders.Clone();
            clone.Font = (WorksheetStyleFont) Font.Clone();
            clone.Interior = (WorksheetStyleInterior) Interior.Clone();
            clone.Name = Name;
            clone.Parent = Parent;
            clone.StyleID = StyleID;
            clone.Workbook = Workbook;

            return clone;
        }*/

        #endregion
    }
}
