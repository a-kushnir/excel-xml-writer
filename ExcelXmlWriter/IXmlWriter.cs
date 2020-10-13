using System.Xml;

namespace ExcelXmlWriter
{
    public interface IXmlWriter
    {
		/// <summary>
		/// Serialize object as xml data to specified xml writer.
		/// </summary>
		/// <param name="writer">Xml writer for serialization.</param>
        void ToXml(XmlWriter writer);
    }
}
