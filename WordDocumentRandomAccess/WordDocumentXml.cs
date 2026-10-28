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
    public static XElement ReadFirstChildElement(Stream wrapped, LoadOptions options, bool tolerateUndeclaredPrefixes, string partName, bool keepContentBeforeFirstError = false)
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
        catch (XmlException exception)
        {
            if (!keepContentBeforeFirstError || !wrapped.CanSeek) throw Malformed(partName, exception);
            wrapped.Position = 0;
            using XmlReader reader = CreateReader(wrapped, tolerateUndeclaredPrefixes, ignoreWhitespace: (options & LoadOptions.PreserveWhitespace) == 0);
            return LoadUntilFirstError(reader)?.Elements().FirstOrDefault() ?? new XElement("empty");
        }
    }

    /// <summary>
    /// A whole part, the way LibreOffice's repair mode keeps it: strict <see cref="XDocument.Load(Stream)"/>
    /// first (the normal path, untouched); only when that throws is the part read again by
    /// <see cref="LoadUntilFirstError"/>. Null when nothing before the error was an element.
    /// </summary>
    public static XDocument? LoadPartKeepingContentBeforeFirstError(Func<Stream> open)
    {
        try
        {
            using Stream stream = open();
            return XDocument.Load(stream);
        }
        catch (XmlException)
        {
            using Stream stream = open();
            using XmlReader reader = CreateReader(stream, tolerateUndeclaredPrefixes: false, ignoreWhitespace: true);
            XElement? root = LoadUntilFirstError(reader);
            return root is null ? null : new XDocument(root);
        }
    }

    /// <summary>
    /// Builds the tree from every node the reader delivers before its first well-formedness error and
    /// stops there; elements still open at the error are closed. This is what LibreOffice keeps from a
    /// malformed .docx part when the user accepts the repair prompt: sax/source/fastparser/fastparser.cxx
    /// runs libxml2's push parser without XML_PARSE_RECOVER (options are only XML_PARSE_NOENT |
    /// XML_PARSE_HUGE), so the first fatal error sets disableSAX (libxml2 parserInternals.c
    /// xmlCtxtVErr: "level == XML_ERR_FATAL &amp;&amp; ctxt->recovery == 0") and no event follows; the
    /// filter's SAXException is swallowed only when SfxObjectShell::IsContinueImportOnFilterExceptions
    /// (sfx2 objmisc.cxx) is true (the repair prompt was approved), and the import keeps what it got.
    /// The element whose start tag holds the error (a duplicate attribute) is not delivered: libxml2
    /// reports it in xmlParseStartTag2 before startElementNs, and XmlReader throws on that Read.
    /// Never resolves anything external: same DtdProcessing.Ignore reader as the strict path.
    /// </summary>
    public static XElement? LoadUntilFirstError(XmlReader reader)
    {
        XElement? root = null;
        var open = new Stack<XElement>();
        try
        {
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                    {
                        var element = new XElement(XName.Get(reader.LocalName, reader.NamespaceURI));
                        bool isEmpty = reader.IsEmptyElement;
                        if (reader.MoveToFirstAttribute())
                        {
                            do
                            {
                                // the default declaration is "xmlns" with no namespace in LINQ to XML
                                XName name = reader.NamespaceURI.Length == 0 || (reader.Prefix.Length == 0 && reader.LocalName == "xmlns")
                                    ? XName.Get(reader.LocalName) : XName.Get(reader.LocalName, reader.NamespaceURI);
                                element.Add(new XAttribute(name, reader.Value));
                            }
                            while (reader.MoveToNextAttribute());
                            reader.MoveToElement();
                        }
                        if (open.Count == 0) { if (root is not null) return root; root = element; }
                        else open.Peek().Add(element);
                        if (!isEmpty) open.Push(element);
                        break;
                    }
                    case XmlNodeType.EndElement:
                        if (open.Count > 0) open.Pop();
                        break;
                    case XmlNodeType.Text:
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        if (open.Count > 0) open.Peek().Add(new XText(reader.Value));
                        break;
                    case XmlNodeType.CDATA:
                        if (open.Count > 0) open.Peek().Add(new XCData(reader.Value));
                        break;
                    case XmlNodeType.Comment:
                        if (open.Count > 0) open.Peek().Add(new XComment(reader.Value));
                        break;
                    case XmlNodeType.ProcessingInstruction:
                        if (open.Count > 0) open.Peek().Add(new XProcessingInstruction(reader.Name, reader.Value));
                        break;
                }
            }
        }
        catch (XmlException) { }
        return root;
    }

    public static InvalidDataException Malformed(string partName, XmlException exception)
        => new($"{partName} is not well-formed XML: {exception.Message}", exception);

    private sealed class UndeclaredPrefixTolerantNamespaceManager(XmlNameTable names) : XmlNamespaceManager(names)
    {
        public override string? LookupNamespace(string prefix)
            => base.LookupNamespace(prefix) ?? (prefix.Length > 0 ? UndeclaredPrefixNamespacePrefix + prefix : null);
    }
}
