using System.Xml;
using System.Xml.Linq;

namespace WordDocumentRandomAccess;

/// <summary>
/// XML reading shared by the XML formats: a malformed part becomes an
/// <see cref="InvalidDataException"/> naming the part (damaged input, not a crash), and OpenDocument
/// reading tolerates undeclared namespace prefixes the way LibreOffice does — xmloff's
/// SvXMLImport::initialize (xmloff/source/core/xmlimp.cxx) always initialises the fast parser with
/// "IgnoreMissingNSDecl", and sax/source/fastparser/fastparser.cxx then gives an element or
/// attribute with an undeclared prefix an empty namespace (an unknown token) instead of failing.
/// Here such a prefix resolves to a namespace no reader looks for, so the element is likewise
/// unknown and skipped, while its text stays in the document.
/// </summary>
public static class WordDocumentXml
{
    /// <summary>The namespace an undeclared prefix resolves to under <see cref="CreateReader"/>'s tolerant mode.</summary>
    public const string UndeclaredPrefixNamespacePrefix = "urn:wordDocumentRandomAccess:undeclared-prefix:";

    public static XmlReader CreateReader(Stream input, bool tolerateUndeclaredPrefixes, bool ignoreWhitespace = false)
        => tolerateUndeclaredPrefixes ? XmlReader.Create(input, Settings(ignoreWhitespace), TolerantContext()) : XmlReader.Create(input, Settings(ignoreWhitespace));

    public static XmlReader CreateReader(TextReader input, bool tolerateUndeclaredPrefixes, bool ignoreWhitespace = false)
        => tolerateUndeclaredPrefixes ? XmlReader.Create(input, Settings(ignoreWhitespace), TolerantContext()) : XmlReader.Create(input, Settings(ignoreWhitespace));

    private static XmlReaderSettings Settings(bool ignoreWhitespace)
        => new() { DtdProcessing = DtdProcessing.Ignore, CloseInput = false, IgnoreWhitespace = ignoreWhitespace };

    private static XmlParserContext TolerantContext()
    {
        var names = new NameTable();
        return new XmlParserContext(names, new UndeclaredPrefixTolerantNamespaceManager(names), null, XmlSpace.None);
    }

    /// <summary>
    /// The first child element of a block or section wrapped in its part's root tag (for the
    /// namespace declarations). Only that element is read: bytes after it up to the next block
    /// start (such as a stray &lt;/w:body&gt;&lt;w:body&gt; between two bodies) are never parsed.
    /// </summary>
    public static XElement ReadFirstChildElement(Stream wrapped, LoadOptions options, bool tolerateUndeclaredPrefixes, string partName)
    {
        try
        {
            // XElement.Load(Stream, options) drops whitespace-only text unless PreserveWhitespace is asked for; the same here.
            using XmlReader reader = CreateReader(wrapped, tolerateUndeclaredPrefixes, ignoreWhitespace: (options & LoadOptions.PreserveWhitespace) == 0);
            reader.MoveToContent();
            if (reader.IsEmptyElement) return new XElement("empty");
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element) { using XmlReader element = reader.ReadSubtree(); return XElement.Load(element, options); }
                if (reader.NodeType == XmlNodeType.EndElement) break;
                reader.Read();
            }
            return new XElement("empty");
        }
        catch (XmlException exception) { throw Malformed(partName, exception); }
    }

    public static InvalidDataException Malformed(string partName, XmlException exception)
        => new($"{partName} is not well-formed XML: {exception.Message}", exception);

    private sealed class UndeclaredPrefixTolerantNamespaceManager(XmlNameTable names) : XmlNamespaceManager(names)
    {
        public override string? LookupNamespace(string prefix)
            => base.LookupNamespace(prefix) ?? (prefix.Length > 0 ? UndeclaredPrefixNamespacePrefix + prefix : null);
    }
}
