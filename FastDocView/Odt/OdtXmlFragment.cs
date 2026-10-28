using System.Text;
using System.Xml.Linq;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>Parses one block cut out of content.xml / a .fodt: its bytes wrapped in the root open tag (namespaces) and a closing tag.</summary>
internal static class OdtXmlFragment
{
    public static XElement Parse(byte[] rootOpenTagBytes, string rootElementName, byte[] bytes, int start, int length)
    {
        byte[] closeTag = Encoding.UTF8.GetBytes("</" + rootElementName + ">");
        using var wrapped = new MemoryStream(rootOpenTagBytes.Length + length + closeTag.Length);
        wrapped.Write(rootOpenTagBytes, 0, rootOpenTagBytes.Length);
        wrapped.Write(bytes, start, length);
        wrapped.Write(closeTag, 0, closeTag.Length);
        wrapped.Position = 0;
        return WordDocumentXml.ReadFirstChildElement(wrapped, LoadOptions.PreserveWhitespace, tolerateUndeclaredPrefixes: true, "OpenDocument body", keepContentBeforeFirstError: true);
    }

    public static string RootElementName(string rootOpenTag) => DocxXmlFragment.RootElementName(rootOpenTag);

    public static byte[] ReadRange(Stream view, long start, long end) => DocxXmlFragment.ReadRange(view, start, end);
}
