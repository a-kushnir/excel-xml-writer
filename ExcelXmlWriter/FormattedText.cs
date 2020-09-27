using System;
using System.Collections.Generic;
using System.Xml;

namespace Hess.ExcelXmlWriter
{
    internal class FormattedText : IXmlWriter
    {
        #region Fields

        private List<FormattedTextItem> m_Items;
        private readonly String m_Text;

        #endregion

        #region Public Properties

        public Boolean IsDefault
        {
            get
            {
                return m_Items.Count == 1 && m_Items[0].IsDefault;
            }
        }

        #endregion

        #region Initialization

        public FormattedText(String text)
        {
            m_Items = new List<FormattedTextItem>();
            m_Items.Add(new FormattedTextItem(text));
            m_Text = text;
        }

        #endregion

        #region Public Methods

        public void Update(Int32 position, Int32 length, FormattedTextItem item)
        {
            if (position < 0 || position > (m_Text ?? String.Empty).Length)
                throw new ArgumentOutOfRangeException("position");
            if (length < 0 || position + length > (m_Text ?? String.Empty).Length)
                throw new ArgumentOutOfRangeException("length");
            if (item == null)
                throw new ArgumentNullException("item");
            if (m_Text == null)
                throw new NullReferenceException("Text");

            Int32 positionBefore = 0;
            Int32 positionAfter = 0;
            List<FormattedTextItem> newItems = new List<FormattedTextItem>();
            foreach(FormattedTextItem textItem in m_Items)
            {
                positionAfter += (textItem.Text ?? String.Empty).Length;
                if ((positionBefore <= position && position <= positionAfter) ||
                    (positionBefore <= position + length && position + length <= positionAfter) ||
                    (positionBefore > position && position + length > positionAfter))
                {
                    Int32 beginPosition = position;
                    if (beginPosition < positionBefore)
                        beginPosition = positionBefore;

                    Int32 endPosition = position + length;
                    if (endPosition > positionAfter)
                        endPosition = positionAfter;

                    FormattedTextItem startItem = (FormattedTextItem)textItem.Clone();
                    startItem.Text = m_Text.Substring(positionBefore, beginPosition - positionBefore);
                    FormattedTextItem middleItem = textItem.Combine(item);
                    middleItem.Text = m_Text.Substring(beginPosition, endPosition - beginPosition);
                    FormattedTextItem finishItem = (FormattedTextItem)textItem.Clone();
                    finishItem.Text = m_Text.Substring(endPosition, positionAfter - endPosition);

                    if (!String.IsNullOrEmpty(startItem.Text))
                        newItems.Add(startItem);

                    newItems.Add(middleItem);

                    if (!String.IsNullOrEmpty(finishItem.Text))
                        newItems.Add(finishItem);
                }
                else
                    newItems.Add(textItem);

                positionBefore = positionAfter;
            }

            m_Items = newItems;
        }

        public void ToXml(XmlWriter writer)
        {
            foreach (FormattedTextItem item in m_Items)
                item.ToXml(writer);
        }

        #endregion
    }
}
