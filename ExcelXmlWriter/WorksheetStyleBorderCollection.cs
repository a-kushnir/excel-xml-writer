using System;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;

namespace ExcelXmlWriter
{
    public class WorksheetStyleBorderCollection: Dictionary<StylePosition, WorksheetStyleBorder>, IXmlWriter, ICloneable
    {
        #region Public Methods

        public void Add(StylePosition position, Int32 weight)
        {
            Add(position, weight, LineStyleOption.Continuous);
        }

        public void Add(StylePosition position, Int32 weight, LineStyleOption lineStyle)
        {
            Add(position, weight, lineStyle, Color.Black);
        }

        public void Add(StylePosition position, Int32 weight, LineStyleOption lineStyle, Color color)
        {
            Add(new WorksheetStyleBorder(position, lineStyle, weight, color));
        }

        public void Create(Int32 weight)
        {
            Create(weight, weight, weight, weight);
        }

        public void Create(Int32 weight, LineStyleOption lineStyle)
        {
            Create(weight, weight, weight, weight, lineStyle);
        }

        public void Create(Int32 weight, LineStyleOption lineStyle, Color color)
        {
            Create(weight, weight, weight, weight, lineStyle, color);
        }

        public void Create(Int32 top, Int32 right, Int32 bottom, Int32 left)
        {
            Create(top, right, bottom, left, LineStyleOption.Continuous, Color.Black);
        }

        public void Create(Int32 top, Int32 right, Int32 bottom, Int32 left, LineStyleOption lineStyle)
        {
            Create(top, right, bottom, left, lineStyle, Color.Black);
        }

        public void Create(Int32 top, Int32 right, Int32 bottom, Int32 left, LineStyleOption lineStyle, Color color)
        {
            Clear();
            if (top > 0)
                Add(new WorksheetStyleBorder(StylePosition.Top, lineStyle, top, color));
            if (right > 0)
                Add(new WorksheetStyleBorder(StylePosition.Right, lineStyle, right, color));
            if (bottom > 0)
                Add(new WorksheetStyleBorder(StylePosition.Bottom, lineStyle, bottom, color));
            if (left > 0)
                Add(new WorksheetStyleBorder(StylePosition.Left, lineStyle, left, color));
        }

        public void ToXml(XmlWriter writer)
        {
            writer.WriteStartElement("Borders");

            foreach (WorksheetStyleBorder item in Values)
                item.ToXml(writer);

            writer.WriteEndElement();
        }

        public object Clone()
        {
            WorksheetStyleBorderCollection clone = new WorksheetStyleBorderCollection();

            foreach(KeyValuePair<StylePosition, WorksheetStyleBorder> kpv in this)
                clone.Add(kpv.Key, (WorksheetStyleBorder)kpv.Value.Clone());

            return clone;
        }

        #endregion

        #region Private Methods

        private void Add(WorksheetStyleBorder border)
        {
            if (border == null)
                throw new ArgumentNullException("border");
            if (border.Position == StylePosition.NotSet)
                throw new ArgumentException("border");

            Add(border.Position, border);
        }

        #endregion
    }
}
