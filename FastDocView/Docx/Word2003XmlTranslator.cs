using System.Text;
using System.Xml;
using System.Xml.Linq;
using FastDocView.DocBinary;
using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// The Word 2003 XML vocabulary ("Word 2003 XML Document", wdFormatXML: one XML file rooted in the 2003 w:wordDocument)
/// mapped onto the .docx vocabulary, element by element, so the .docx reader renders it: <see cref="Word2003PackageStream"/>
/// presents the file as a package whose header parts (styles, numbering, settings, relationships) are translated at open
/// and whose body and notes parts are translated block by block as the reader reaches them (windowed, never whole).
/// The vocabulary is WordprocessingML's predecessor, so most elements map one to one; the differences handled here follow
/// LibreOffice's own import of the format (the XSLT filter filter/source/xslt/import/wordml/): wordml2ooo.xsl:244-250
/// (wx:sect / wx:sub-section only wrap the body), wordml2ooo_list.xsl:20-25 and :339-347 (w:listPr w:ilfo -&gt;
/// w:lists/w:list[@w:ilfo]/w:ilst -&gt; w:listDef[@w:listDefId]; w:nfc is the MS-DOC number format code, :489-547),
/// wordml2ooo_text.xsl:937-958 (aml:annotation bookmarks), :959-986 (w:hlink w:dest / w:bookmark / w:target), :1044-1065
/// (w:footnote / w:endnote inline in the run, the note's body in place). Further renames (hyphenated names to camel case:
/// w:b-cs -&gt; w:bCs, w:first-line -&gt; w:firstLine; w:fareast -&gt; w:eastAsia; w:vmerge -&gt; w:vMerge; pictures as
/// w:binData base64 named "wordml://...", referenced by v:imagedata src) are Word's own 2003 -&gt; 2007 mapping, checked
/// against Word's save of the same document as .docx (test-data/word-xml). Tracked insertions are kept and deletions
/// dropped (the printed result); comments, which do not print, are dropped.
/// </summary>
internal static class Word2003XmlTranslator
{
    public const string WordNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
    private const string AuxiliaryNamespace = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private const string AnnotationNamespace = "http://schemas.microsoft.com/aml/2001/core";
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    internal const string DocxMainNamespace = MainNamespace;
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    private const string OfficeNamespace = "urn:schemas-microsoft-com:office:office";
    private const string WordVmlNamespace = "urn:schemas-microsoft-com:office:word";
    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";
    private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";
    private const string ImageNamePrefix = "wordml://";

    /// <summary>
    /// The package parts made at open: content types, relationships and the translated header parts (each left out when
    /// the file has none). The relationships list every hyperlink target and picture of the whole file, which the open
    /// pass collected into <paramref name="state"/>, so a block translated later finds its ids there.
    /// </summary>
    internal static List<(string Name, byte[] Bytes)> SynthesizedParts(Header header, TranslationState state, bool hasFootnotes, bool hasEndnotes)
    {
        var parts = new List<(string Name, byte[] Bytes)>();
        if (header.Styles is null && header.DefaultFonts is not null) header.Styles = new XElement(XName.Get("styles", WordNamespace));
        if (header.Styles is not null)
        {
            // the document's default fonts lead the styles part as docDefaults (ECMA-376 17.7.5.1), as Word saves them
            if (header.DefaultFonts is { } fonts)
                header.Styles.AddFirst(new XElement(XName.Get("docDefaults", WordNamespace), new XElement(XName.Get("rPrDefault", WordNamespace),
                    new XElement(XName.Get("rPr", WordNamespace), new XElement(XName.Get("rFonts", WordNamespace), fonts.Attributes())))));
            XElement styles = header.Styles;
            parts.Add(("word/styles.xml", XmlPart(writer => TranslateTree(styles, writer, state, "styles"))));
        }
        if (header.Lists is { } lists) parts.Add(("word/numbering.xml", XmlPart(writer => TranslateTree(lists, writer, state, "numbering"))));
        if (header.DocumentProperties is { } settings) parts.Add(("word/settings.xml", XmlPart(writer => TranslateTree(settings, writer, state, "settings"))));

        parts.Add(("[Content_Types].xml", Utf8(ContentTypes(state, header, hasFootnotes, hasEndnotes))));
        parts.Add(("_rels/.rels", Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">"
            + "<Relationship Id=\"rId1\" Type=\"" + RelationshipNamespace + "/officeDocument\" Target=\"word/document.xml\"/></Relationships>")));
        string imageRelationships = ImageRelationships(state);
        var document = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">");
        if (header.Styles is not null) document.Append(Relationship("rIdStyles", "styles", "styles.xml"));
        if (header.Lists is not null) document.Append(Relationship("rIdNumbering", "numbering", "numbering.xml"));
        if (header.DocumentProperties is not null) document.Append(Relationship("rIdSettings", "settings", "settings.xml"));
        if (hasFootnotes) document.Append(Relationship("rIdFootnotes", "footnotes", "footnotes.xml"));
        if (hasEndnotes) document.Append(Relationship("rIdEndnotes", "endnotes", "endnotes.xml"));
        document.Append(imageRelationships).Append("</Relationships>");
        parts.Add(("word/_rels/document.xml.rels", Utf8(document.ToString())));
        if (header.Lists is not null)      // picture bullets (w:listPicBullet) name their pictures through numbering.xml's own relationships
            parts.Add(("word/_rels/numbering.xml.rels", Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">"
                                                         + imageRelationships + "</Relationships>")));
        foreach (string? notes in new[] { hasFootnotes ? "footnotes" : null, hasEndnotes ? "endnotes" : null })
            if (notes is not null)
                parts.Add(("word/_rels/" + notes + ".xml.rels", Utf8("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">"
                                                                   + imageRelationships + "</Relationships>")));
        return parts;
    }

    /// <summary>The package name of a picture (a w:binData name).</summary>
    internal static string MediaPartPath(string imageName) => "word/media/" + MediaPartName(imageName);

    private const string HyperlinkIdPrefix = "rIdL", ImageIdPrefix = "rIdI";

    /// <summary>A hyperlink target's relationship id: the target itself, base64url-encoded (letters, digits, '-', '_': a valid id).</summary>
    internal static string HyperlinkIdOf(string target) => HyperlinkIdPrefix + Base64Url(target);

    /// <summary>A picture's relationship id: its w:binData name, base64url-encoded.</summary>
    internal static string ImageIdOf(string imageName) => ImageIdPrefix + Base64Url(imageName);

    /// <summary>The target a <see cref="HyperlinkIdOf"/> id stands for; null for any other id.</summary>
    internal static string? HyperlinkTargetOf(string relationshipId)
        => relationshipId.StartsWith(HyperlinkIdPrefix, StringComparison.Ordinal) ? FromBase64Url(relationshipId.Substring(HyperlinkIdPrefix.Length)) : null;

    /// <summary>The picture name a <see cref="ImageIdOf"/> id stands for; null for any other id.</summary>
    internal static string? ImageNameOf(string relationshipId)
        => relationshipId.StartsWith(ImageIdPrefix, StringComparison.Ordinal) ? FromBase64Url(relationshipId.Substring(ImageIdPrefix.Length)) : null;

    /// <summary>The picture name of a <see cref="MediaPartPath"/> part name; null for any other part.</summary>
    internal static string? ImageNameOfMediaPart(string partName)
    {
        const string Media = "word/media/";
        if (!partName.StartsWith(Media + ImageIdPrefix, StringComparison.Ordinal)) return null;
        int dash = partName.IndexOf('-', Media.Length + ImageIdPrefix.Length);
        // base64url text can hold '-': the id is the part up to the dash that the name's own encoding ends at
        for (; dash > 0; dash = partName.IndexOf('-', dash + 1))
            if (ImageNameOf(partName.Substring(Media.Length, dash - Media.Length)) is { } name && MediaPartPath(name) == partName) return name;
        return null;
    }

    private static string Base64Url(string text)
        => Convert.ToBase64String(new UTF8Encoding(false).GetBytes(text)).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    private static string? FromBase64Url(string text)
    {
        string padded = text.Replace('-', '+').Replace('_', '/');
        padded += new string('=', (4 - padded.Length % 4) % 4);
        try { return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(Convert.FromBase64String(padded)); }
        catch (FormatException) { return null; }
        catch (ArgumentException) { return null; }
    }

    private static byte[] Utf8(string text) => new UTF8Encoding(false).GetBytes(text);

    private static string Relationship(string id, string type, string target)
        => "<Relationship Id=\"" + id + "\" Type=\"" + RelationshipNamespace + "/" + type + "\" Target=\"" + target + "\"/>";

    private static string ImageRelationships(TranslationState state)
    {
        var text = new StringBuilder();
        foreach (KeyValuePair<string, string> image in state.ImageRelationships)
            text.Append(Relationship(image.Value, "image", "media/" + MediaPartName(image.Key)));      // only safe characters
        return text.ToString();
    }

    private static string Escape(string text) => System.Security.SecurityElement.Escape(text) ?? "";

    /// <summary>"wordml://03000001.png" -> "03000001.png"; any character that is not safe in a part name becomes '_'.</summary>
    private static string MediaFileName(string imageName)
    {
        string name = imageName.StartsWith(ImageNamePrefix, StringComparison.OrdinalIgnoreCase) ? imageName.Substring(ImageNamePrefix.Length) : imageName;
        var safe = new StringBuilder(name.Length);
        foreach (char c in name) safe.Append(char.IsLetterOrDigit(c) && c < 128 || c is '.' or '-' or '_' ? c : '_');
        return safe.Length == 0 ? "image" : safe.ToString();
    }

    /// <summary>The picture's part name: its relationship id first, so two names that clean to the same text stay two parts.</summary>
    private static string MediaPartName(string imageName) => ImageIdOf(imageName) + "-" + MediaFileName(imageName);

    private static string ContentTypes(TranslationState state, Header header, bool hasFootnotes, bool hasEndnotes)
    {
        const string Wordprocessing = "application/vnd.openxmlformats-officedocument.wordprocessingml.";
        var text = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>");
        var extensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string image in state.ImageRelationships.Keys)
        {
            string extension = Path.GetExtension(MediaFileName(image)).TrimStart('.');
            if (extension.Length > 0 && extensions.Add(extension))
                text.Append("<Default Extension=\"").Append(Escape(extension)).Append("\" ContentType=\"").Append(ImageContentType(extension)).Append("\"/>");
        }
        text.Append("<Override PartName=\"/word/document.xml\" ContentType=\"").Append(Wordprocessing).Append("document.main+xml\"/>");
        if (header.Styles is not null) text.Append("<Override PartName=\"/word/styles.xml\" ContentType=\"").Append(Wordprocessing).Append("styles+xml\"/>");
        if (header.Lists is not null) text.Append("<Override PartName=\"/word/numbering.xml\" ContentType=\"").Append(Wordprocessing).Append("numbering+xml\"/>");
        if (header.DocumentProperties is not null) text.Append("<Override PartName=\"/word/settings.xml\" ContentType=\"").Append(Wordprocessing).Append("settings+xml\"/>");
        if (hasFootnotes) text.Append("<Override PartName=\"/word/footnotes.xml\" ContentType=\"").Append(Wordprocessing).Append("footnotes+xml\"/>");
        if (hasEndnotes) text.Append("<Override PartName=\"/word/endnotes.xml\" ContentType=\"").Append(Wordprocessing).Append("endnotes+xml\"/>");
        return text.Append("</Types>").ToString();
    }

    private static string ImageContentType(string extension) => extension.ToLowerInvariant() switch
    {
        "png" => "image/png", "jpg" or "jpeg" => "image/jpeg", "gif" => "image/gif", "bmp" => "image/bmp",
        "emf" => "image/x-emf", "wmf" => "image/x-wmf", "tif" or "tiff" => "image/tiff", _ => "application/octet-stream",
    };

    private static byte[] XmlPart(Action<XmlWriter> write)
    {
        using var stream = new MemoryStream();
        using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false }))
        {
            writer.WriteStartDocument(standalone: true);
            write(writer);
            writer.WriteEndDocument();
        }
        return stream.ToArray();
    }

    internal static XmlReader CreateReader(Stream source)
        => XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false, IgnoreComments = true, IgnoreProcessingInstructions = true });

    // ── header: w:styles, w:lists, w:docPr (small; read as trees) ─────────────────────────────

    internal sealed class Header
    {
        public XElement? Styles, Lists, DocumentProperties;
        public XElement? DefaultFonts;           // w:fonts/w:defaultFonts: Word 2007's docDefaults run fonts
    }

    /// <summary>The header parts before w:body, read from the file's start (the stream is not disposed).</summary>
    internal static Header ReadHeader(Stream source)
    {
        var header = new Header();
        source.Position = 0;
        using XmlReader reader = CreateReader(source);
        try
        {
            if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "wordDocument" || reader.NamespaceURI != WordNamespace)
                throw new InvalidDataException("Word 2003 XML: the root is not w:wordDocument.");
            if (reader.IsEmptyElement) return header;
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element) { if (reader.NodeType == XmlNodeType.EndElement) break; reader.Read(); continue; }
                if (reader.NamespaceURI == WordNamespace && reader.LocalName == "body") break;
                if (reader.NamespaceURI == WordNamespace && reader.LocalName is "styles" or "lists" or "docPr" or "fonts")
                {
                    string name = reader.LocalName;
                    var element = (XElement)XNode.ReadFrom(reader);
                    if (name == "styles") header.Styles = element;
                    else if (name == "lists") header.Lists = element;
                    else if (name == "fonts") header.DefaultFonts = element.Element(XName.Get("defaultFonts", WordNamespace));
                    else header.DocumentProperties = element;
                }
                else reader.Skip();
            }
        }
        catch (XmlException)
        {
            // malformed before the body: keep the header parts read so far (the body pass reports what it keeps)
        }
        return header;
    }

    // ── the streaming translation ─────────────────────────────────────────────────────────────

    internal enum NotePass { None, Footnotes, Endnotes }

    /// <summary>
    /// The ids a translation gives. A hyperlink's or picture's relationship id is made from its target / picture name
    /// itself (<see cref="HyperlinkIdOf"/>, <see cref="ImageIdOf"/>), so no pass over the file has to collect them before a
    /// block is translated: a block translated at any time, in any order, gets the same ids, and the package resolves an id
    /// when it is looked up (<see cref="IDocxOnDemandRelationships"/>). Note ids are made from the byte offset of the block
    /// that holds the note, so a block's ids never depend on the blocks before it.
    /// </summary>
    internal sealed class TranslationState
    {
        /// <summary>The pictures the header parts use (list picture bullets): listed in the relationship parts made at open.</summary>
        public readonly Dictionary<string, string> ImageRelationships = new(StringComparer.Ordinal);   // "wordml://..." -> relationship id
        private bool _isFrozen;
        private int _noteIdBase, _notesInBlock;

        /// <summary>Ends the header translation: pictures named after it are no longer listed (they resolve on demand).</summary>
        public void Freeze() => _isFrozen = true;

        public string HyperlinkId(string target) => HyperlinkIdOf(target);

        public string ImageId(string name)
        {
            string id = ImageIdOf(name);
            if (!_isFrozen && !ImageRelationships.ContainsKey(name)) ImageRelationships[name] = id;
            return id;
        }

        /// <summary>
        /// Starts a body block at <paramref name="blockOffset"/> in the file: its notes get the ids offset + 1, offset + 2, ...
        /// (unique: a note takes more than one byte of its block, so they stay below the next block's offset).
        /// </summary>
        public void BeginBlock(long blockOffset)
        {
            if (blockOffset >= int.MaxValue / 2) throw new NotSupportedException("Word 2003 XML files over 1 GB are not supported.");
            _noteIdBase = (int)blockOffset;
            _notesInBlock = 0;
        }

        public int NextNoteId() => _noteIdBase + ++_notesInBlock;
    }

    internal static void TranslateTree(XElement element, XmlWriter writer, TranslationState state, string rootName)
    {
        using XmlReader reader = element.CreateReader();
        reader.MoveToContent();
        new Translation(writer, state, NotePass.None).Root(reader, rootName);
    }

    /// <summary>The element-by-element translation (an explicit stack: no recursion, any nesting depth).</summary>
    internal sealed class Translation(XmlWriter writer, TranslationState state, NotePass pass)
    {
        private enum Frame { Written, Picture, Unwrapped, Note }

        private readonly Stack<Frame> _frames = new();
        private int _noteDepth;                // > 0 inside the note a notes pass writes
        private int _pictureDepth;             // > 0 inside w:pict / w:object: VML is copied as it is
        private long _nodes;
        public CancellationToken Cancellation { get; init; }

        private bool IsWriting => pass == NotePass.None || _noteDepth > 0;

        public void StartRoot(string rootName)
        {
            writer.WriteStartElement("w", rootName, MainNamespace);
            writer.WriteAttributeString("xmlns", "r", XmlnsNamespace, RelationshipNamespace);
            writer.WriteAttributeString("xmlns", "v", XmlnsNamespace, VmlNamespace);
            writer.WriteAttributeString("xmlns", "o", XmlnsNamespace, OfficeNamespace);
            writer.WriteAttributeString("xmlns", "w10", XmlnsNamespace, WordVmlNamespace);
        }

        /// <summary>A header tree (w:styles / w:lists / w:docPr) as a whole part under <paramref name="rootName"/>.</summary>
        public void Root(XmlReader reader, string rootName)
        {
            StartRoot(rootName);
            if (!reader.IsEmptyElement)
            {
                reader.Read();
                while (!reader.EOF && reader.NodeType != XmlNodeType.EndElement)
                {
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        // docPr's note separators are 2003 note bodies; Word's settings keep only their ids
                        if (rootName == "settings" && reader.NamespaceURI == WordNamespace && reader.LocalName is "footnotePr" or "endnotePr") { reader.Skip(); continue; }
                        Element(reader);
                    }
                    else reader.Read();
                }
            }
            writer.WriteEndElement();
        }

        /// <summary>Translates the element the reader is on and everything inside it; the reader ends after it.</summary>
        public void Element(XmlReader reader)
        {
            int depth = _frames.Count;
            Start(reader);
            while (_frames.Count > depth && !reader.EOF)
            {
                if ((++_nodes & 0xFFF) == 0) Cancellation.ThrowIfCancellationRequested();
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element: Start(reader); break;
                    case XmlNodeType.EndElement: End(); reader.Read(); break;
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                    case XmlNodeType.Whitespace:
                        if (IsWriting && _frames.Count > 0 && _frames.Peek() != Frame.Unwrapped) writer.WriteString(reader.Value);
                        reader.Read();
                        break;
                    default: reader.Read(); break;
                }
            }
        }

        /// <summary>Closes the elements still open (a malformed file's end).</summary>
        public void CloseAll()
        {
            while (_frames.Count > 0) End();
        }

        private void End()
        {
            Frame frame = _frames.Pop();
            if (frame == Frame.Note) { _noteDepth--; writer.WriteEndElement(); return; }
            if (frame == Frame.Unwrapped) return;
            writer.WriteEndElement();
            if (frame == Frame.Picture) _pictureDepth--;
        }

        /// <summary>Opens the reader's element (or skips it whole) and moves the reader past its start tag.</summary>
        private void Start(XmlReader reader)
        {
            string ns = reader.NamespaceURI, local = reader.LocalName;
            bool isEmpty = reader.IsEmptyElement;
            if (ns == WordNamespace && local is "footnote" or "endnote")
            {
                bool isFootnote = local == "footnote";
                int id = state.NextNoteId();
                if (pass == NotePass.None)
                {
                    writer.WriteStartElement("w", isFootnote ? "footnoteReference" : "endnoteReference", MainNamespace);
                    writer.WriteAttributeString("w", "id", MainNamespace, id.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    reader.Skip();
                    return;
                }
                if ((pass == NotePass.Footnotes) != isFootnote) { reader.Skip(); return; }
                writer.WriteStartElement("w", local, MainNamespace);
                writer.WriteAttributeString("w", "id", MainNamespace, id.ToString(System.Globalization.CultureInfo.InvariantCulture));
                _noteDepth++;
                if (isEmpty) { _noteDepth--; writer.WriteEndElement(); reader.Read(); return; }
                _frames.Push(Frame.Note);
                reader.Read();
                return;
            }

            if (_pictureDepth > 0 && ns != WordNamespace)
            {
                WriteVerbatim(reader, isEmpty);
                return;
            }

            if (ns == WordNamespace) { WordElement(reader, local, isEmpty); return; }
            if (ns == AuxiliaryNamespace)
            {
                if (local is "sect" or "sub-section" or "pBdrGroup") Unwrap(reader, isEmpty); else reader.Skip();
                return;
            }
            if (ns == AnnotationNamespace) { Annotation(reader, local, isEmpty); return; }
            if (ns is VmlNamespace or OfficeNamespace or WordVmlNamespace) { WriteVerbatim(reader, isEmpty); return; }
            Unwrap(reader, isEmpty);           // smart tags, custom XML, schema-library elements: their content stays
        }

        private void Unwrap(XmlReader reader, bool isEmpty)
        {
            if (isEmpty) { reader.Read(); return; }
            _frames.Push(Frame.Unwrapped);
            reader.Read();
        }

        /// <summary>A VML (or other foreign) element inside a picture, copied with all its attributes; a 2003 picture reference becomes a relationship.</summary>
        private void WriteVerbatim(XmlReader reader, bool isEmpty)
        {
            if (!IsWriting) { Unwrap(reader, isEmpty); return; }
            writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
            bool isImageData = reader.LocalName == "imagedata" && reader.NamespaceURI == VmlNamespace;
            if (reader.MoveToFirstAttribute())
            {
                do
                {
                    if (reader.NamespaceURI == XmlnsNamespace) continue;
                    if (isImageData && reader.LocalName == "src" && reader.NamespaceURI.Length == 0 && reader.Value.StartsWith(ImageNamePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        writer.WriteAttributeString("r", "id", RelationshipNamespace, state.ImageId(reader.Value));
                        continue;
                    }
                    string attributeNamespace = reader.NamespaceURI == WordNamespace ? MainNamespace : reader.NamespaceURI;
                    if (attributeNamespace is AuxiliaryNamespace or AnnotationNamespace) continue;
                    writer.WriteAttributeString(attributeNamespace == MainNamespace ? "w" : reader.Prefix, CamelCase(reader.LocalName), attributeNamespace, reader.Value);
                }
                while (reader.MoveToNextAttribute());
                reader.MoveToElement();
            }
            Opened(reader, isEmpty, picture: true);
        }

        private void Opened(XmlReader reader, bool isEmpty, bool picture = false)
        {
            if (isEmpty) { writer.WriteEndElement(); reader.Read(); return; }
            if (picture) _pictureDepth++;
            _frames.Push(picture ? Frame.Picture : Frame.Written);
            reader.Read();
        }

        private void Annotation(XmlReader reader, string local, bool isEmpty)
        {
            if (local == "content") { Unwrap(reader, isEmpty); return; }
            if (local != "annotation") { reader.Skip(); return; }
            string? type = reader.GetAttribute("type", WordNamespace);
            switch (type)
            {
                case "Word.Insertion": Unwrap(reader, isEmpty); return;      // the printed text keeps an insertion
                case "Word.Bookmark.Start":
                case "Word.Bookmark.End":
                    if (IsWriting)
                    {
                        bool isStart = type == "Word.Bookmark.Start";
                        writer.WriteStartElement("w", isStart ? "bookmarkStart" : "bookmarkEnd", MainNamespace);
                        writer.WriteAttributeString("w", "id", MainNamespace, reader.GetAttribute("id", AnnotationNamespace) ?? "0");
                        if (isStart) writer.WriteAttributeString("w", "name", MainNamespace, reader.GetAttribute("name", WordNamespace) ?? "");
                        writer.WriteEndElement();
                    }
                    reader.Skip();
                    return;
                default: reader.Skip(); return;                                // deletions, formatting changes, comments: not printed
            }
        }

        private void WordElement(XmlReader reader, string local, bool isEmpty)
        {
            switch (local)
            {
                case "binData": reader.Skip(); return;                         // written into word/media by its own pass
                case "listPicture": reader.Skip(); return;                     // not a .docx element (picture bullets come from w:listPicBullet)
                case "fonts": reader.Skip(); return;
                case "pict":
                case "object":
                    if (!IsWriting) { Unwrap(reader, isEmpty); return; }
                    writer.WriteStartElement("w", local, MainNamespace);
                    WordAttributes(reader, local);
                    Opened(reader, isEmpty, picture: true);
                    return;
                case "hlink":
                    if (!IsWriting) { Unwrap(reader, isEmpty); return; }
                    writer.WriteStartElement("w", "hyperlink", MainNamespace);
                    string? destination = reader.GetAttribute("dest", WordNamespace), bookmark = reader.GetAttribute("bookmark", WordNamespace);
                    if (!string.IsNullOrEmpty(destination)) writer.WriteAttributeString("r", "id", RelationshipNamespace, state.HyperlinkId(destination!));
                    if (!string.IsNullOrEmpty(bookmark)) writer.WriteAttributeString("w", "anchor", MainNamespace, bookmark);
                    if (reader.GetAttribute("screenTip", WordNamespace) is { } tip) writer.WriteAttributeString("w", "tooltip", MainNamespace, tip);
                    if (reader.GetAttribute("target", WordNamespace) is { } target) writer.WriteAttributeString("w", "tgtFrame", MainNamespace, target);
                    Opened(reader, isEmpty);
                    return;
            }
            if (!IsWriting) { Unwrap(reader, isEmpty); return; }
            string name = ElementName(local);
            writer.WriteStartElement("w", name, MainNamespace);
            if (local == "nfc") writer.WriteAttributeString("w", "val", MainNamespace, NumberFormat(reader.GetAttribute("val", WordNamespace)));
            else if (local == "plt") writer.WriteAttributeString("w", "val", MainNamespace, LowerFirst(reader.GetAttribute("val", WordNamespace) ?? "multilevel"));
            else WordAttributes(reader, local);
            if (name is "t" or "instrText" or "delText" && reader.GetAttribute("space", XmlNamespace) is null)
                writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve");     // the 2003 root's xml:space="preserve" is not inherited by a part
            Opened(reader, isEmpty);
        }

        private void WordAttributes(XmlReader reader, string element)
        {
            if (!reader.MoveToFirstAttribute()) return;
            do
            {
                string ns = reader.NamespaceURI;
                if (ns == XmlnsNamespace || ns is AuxiliaryNamespace or AnnotationNamespace) continue;
                if (ns == XmlNamespace) { writer.WriteAttributeString("xml", reader.LocalName, XmlNamespace, reader.Value); continue; }
                if (ns != WordNamespace) continue;                             // wsp:rsid* and other editing data
                string local = reader.LocalName, value = reader.Value;
                string name = (element, local) switch
                {
                    ("listDef", "listDefId") => "abstractNumId",
                    ("list", "ilfo") => "numId",
                    ("listPicBullet", "listPicBulletId") => "numPicBulletId",
                    (_, "fareast") => "eastAsia",
                    _ => CamelCase(local),
                };
                if (element == "tblLayout" && local == "type") value = LowerFirst(value);
                else if (value == "on") value = "1";
                else if (value == "off") value = "0";
                writer.WriteAttributeString("w", name, MainNamespace, value);
            }
            while (reader.MoveToNextAttribute());
            reader.MoveToElement();
        }
    }

    private static string ElementName(string local) => local switch
    {
        "listPr" => "numPr",
        "ilfo" => "numId",
        "lists" => "numbering",
        "listDef" => "abstractNum",
        "list" => "num",
        "ilst" => "abstractNumId",
        "lsid" => "nsid",
        "plt" => "multiLevelType",
        "nfc" => "numFmt",
        "listStyleLink" => "numStyleLink",
        "listPicBullet" => "numPicBullet",
        "vmerge" => "vMerge",
        "hmerge" => "hMerge",
        "docPr" => "settings",
        _ => CamelCase(local),
    };

    /// <summary>The 2003 hyphenated names in Word 2007's camel case: "b-cs" -> "bCs", "first-line" -> "firstLine", "panose-1" -> "panose1".</summary>
    private static string CamelCase(string name)
    {
        if (name.IndexOf('-') < 0) return name;
        var text = new StringBuilder(name.Length);
        bool upper = false;
        foreach (char c in name)
        {
            if (c == '-') { upper = true; continue; }
            text.Append(upper ? char.ToUpperInvariant(c) : c);
            upper = false;
        }
        return text.ToString();
    }

    private static string LowerFirst(string value) => value.Length == 0 ? value : char.ToLowerInvariant(value[0]) + value.Substring(1);

    /// <summary>w:nfc is the MS-DOC number format code (LibreOffice wordml2ooo_list.xsl:489-547 maps the same codes).</summary>
    private static string NumberFormat(string? code)
        => byte.TryParse(code, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out byte value)
            ? DocBinaryListDefinitions.NumberFormatName(value) : "decimal";
}
