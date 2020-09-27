using System;
using System.Drawing;
using System.Web;

namespace Hess.ExcelXmlWriter
{
    internal class Helper
    {
        #region Public Methods

        public static String XmlConvert(Color value)
        {
            return String.Format("#{0:x2}{1:x2}{2:x2}", value.R, value.G, value.B);
        }

        public static String XmlConvert(Boolean value)
        {
            return (value ? 1 : 0).ToString();
        }

        public static String XmlConvert(String value)
        {
            if (!String.IsNullOrEmpty(value))
                value = HttpUtility.HtmlEncode(value)
                    .Replace(Environment.NewLine, "&#10;")
                    .Replace("\n", "&#10;")
                    .Replace("\r", "&#10;");

            return value;
        }

        #endregion
    }
}
