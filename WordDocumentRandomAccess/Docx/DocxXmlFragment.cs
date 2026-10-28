using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Parses one element cut out of a part (a body block, a note): wraps its bytes in the part's root
/// opening tag (for the namespace declarations) and a closing tag, strips drawings first
/// (<see cref="DocxDrawingStripper"/>), and returns the element. A malformed element keeps what
/// precedes its first error (<see cref="WordDocumentXml.LoadUntilFirstError"/>).
/// </summary>
internal static class DocxXmlFragment
{
    public static XElement Parse(byte[] rootOpenTagBytes, string rootElementName, byte[] bytes, int start, int length, string partName)
    {
        ArraySegment<byte> element = DocxDrawingStripper.Strip(bytes, start, length);
        byte[] closeTag = Encoding.UTF8.GetBytes("</" + rootElementName + ">");
        using var wrapped = new MemoryStream(rootOpenTagBytes.Length + element.Count + closeTag.Length);
        wrapped.Write(rootOpenTagBytes, 0, rootOpenTagBytes.Length);
        wrapped.Write(element.Array!, element.Offset, element.Count);
        wrapped.Write(closeTag, 0, closeTag.Length);
        wrapped.Position = 0;
        return WordDocumentXml.ReadFirstChildElement(wrapped, LoadOptions.None, tolerateUndeclaredPrefixes: false, partName, keepContentBeforeFirstError: true);
    }

    /// <summary>"w:document" from "&lt;w:document xmlns:w=…&gt;".</summary>
    public static string RootElementName(string rootOpenTag)
    {
        int start = rootOpenTag.IndexOf('<') + 1;
        int end = start;
        while (end < rootOpenTag.Length && rootOpenTag[end] is not (' ' or '>' or '/' or '\t' or '\r' or '\n')) end++;
        return rootOpenTag.Substring(start, end - start);
    }

    public static byte[] ReadRange(Stream view, long start, long end)
    {
        var bytes = new byte[end - start];
        view.Position = start;
        int filled = 0;
        while (filled < bytes.Length)
        {
            int read = view.Read(bytes, filled, bytes.Length - filled);
            if (read <= 0) throw new EndOfStreamException("Part ended inside an element.");
            filled += read;
        }
        return bytes;
    }
}
