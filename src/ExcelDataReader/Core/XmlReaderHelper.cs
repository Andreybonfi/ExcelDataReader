using System.Xml;

namespace ExcelDataReader.Core
{
    internal static class XmlReaderHelper
    {
        public static bool ReadFirstContent(XmlReader xmlReader)
        {
            if (xmlReader.IsEmptyElement)
            {
                xmlReader.Read();
                return false;
            }

            xmlReader.MoveToContent();
            xmlReader.Read();
            return true;
        }

        public static bool SkipContent(XmlReader xmlReader)
        {
            if (xmlReader.NodeType == XmlNodeType.EndElement)
            {
                xmlReader.Read();
                return false;
            }

            xmlReader.Skip();
            return true;
        }

        public static bool IsEndElement(XmlReader xmlReader, string localname, string ns)
        {
            if (xmlReader.MoveToContent() == XmlNodeType.EndElement)
            {
                if (xmlReader.LocalName == localname)
                {
                    return xmlReader.NamespaceURI == ns;
                }

                return false;
            }

            return false;
        }
    }
}
