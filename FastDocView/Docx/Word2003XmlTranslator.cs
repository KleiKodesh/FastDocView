using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using FastDocView.DocBinary;
using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// Word 2003 XML ("Word 2003 XML Document", wdFormatXML: one XML file rooted in the 2003 w:wordDocument) read by
/// translating it into a .docx package once, beside the document's sidecar index, then opening that package with the
/// .docx reader — no second renderer. The translation streams (XmlReader to XmlWriter, a few passes over the file,
/// bounded memory: only the small header parts w:styles / w:lists / w:docPr are held as trees), and its output is
/// deterministic, so an index built over one translation fits the next. The vocabulary is WordprocessingML's
/// predecessor, so most elements map one to one; the differences handled here follow LibreOffice's own import of the
/// format (the XSLT filter filter/source/xslt/import/wordml/): wordml2ooo.xsl:244-250 (wx:sect / wx:sub-section only
/// wrap the body), wordml2ooo_list.xsl:20-25 and :339-347 (w:listPr w:ilfo -> w:lists/w:list[@w:ilfo]/w:ilst ->
/// w:listDef[@w:listDefId]; w:nfc is the MS-DOC number format code, :489-547), wordml2ooo_text.xsl:937-958
/// (aml:annotation bookmarks), :959-986 (w:hlink w:dest / w:bookmark / w:target), :1044-1065 (w:footnote / w:endnote
/// inline in the run, the note's body in place). Further renames (hyphenated names to camel case: w:b-cs -> w:bCs,
/// w:first-line -> w:firstLine; w:fareast -> w:eastAsia; w:vmerge -> w:vMerge; pictures as w:binData base64 named
/// "wordml://...", referenced by v:imagedata src) are Word's own 2003 -> 2007 mapping, checked against Word's save of
/// the same document as .docx (test-data/word-xml). Tracked insertions are kept and deletions dropped (the printed
/// result); comments, which do not print, are dropped.
/// </summary>
internal static class Word2003XmlTranslator
{
    public const string WordNamespace = "http://schemas.microsoft.com/office/word/2003/wordml";
    private const string AuxiliaryNamespace = "http://schemas.microsoft.com/office/word/2003/auxHint";
    private const string AnnotationNamespace = "http://schemas.microsoft.com/aml/2001/core";
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    private const string OfficeNamespace = "urn:schemas-microsoft-com:office:office";
    private const string WordVmlNamespace = "urn:schemas-microsoft-com:office:word";
    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";
    private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";
    private const string ImageNamePrefix = "wordml://";
    private const string SourceStampEntry = "fastdocview/source";
    private static readonly DateTimeOffset FixedEntryTime = new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);   // deterministic output

    /// <summary>Opens a Word 2003 XML file: translates it when its package is missing or older than the file, then opens the package.</summary>
    public static DocxWindowedDocument Open(string filePath, WordIndexPolicy? policy, IProgress<WordDocumentIndexProgress>? progress,
                                            CancellationToken cancellation)
    {
        policy ??= WordIndexPolicy.Default;
        string packagePath = WordDocumentIndexCache.TranslatedPackagePathFor(policy.IndexPathFor(filePath));
        if (!IsCurrent(packagePath, filePath))
        {
            // a sidecar built over an older package must not read the new one at its offsets
            string indexPath = policy.IndexPathFor(filePath);
            try { if (File.Exists(indexPath)) File.Delete(indexPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            Translate(filePath, packagePath, cancellation);
        }
        return DocxWindowedDocument.Open(filePath, DocxWindowedDocument.DocxPackageForm.Zip, packagePath, policy, progress, cancellation);
    }

    /// <summary>The policy's answer without translating: the file's size stands for its text (it is all markup and text).</summary>
    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy)
    {
        bool hasCachedIndex = DocxRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, new FileInfo(filePath).Length, hasCachedIndex);
    }

    private static string SourceStamp(string filePath)
    {
        var info = new FileInfo(filePath);
        return info.Length + " " + info.LastWriteTimeUtc.Ticks;
    }

    /// <summary>True when the translated package exists and was made from this version of the file.</summary>
    private static bool IsCurrent(string packagePath, string filePath)
    {
        if (!File.Exists(packagePath)) return false;
        try
        {
            using FileStream package = WordDocumentFileStream.OpenShared(packagePath);
            if (ZipEntryLocator.Locate(package, SourceStampEntry) is not { IsStored: true } entry) return false;
            var bytes = new byte[entry.CompressedSize];
            package.Position = entry.DataOffset;
            int filled = 0;
            while (filled < bytes.Length) { int read = package.Read(bytes, filled, bytes.Length - filled); if (read <= 0) return false; filled += read; }
            return Encoding.UTF8.GetString(bytes) == SourceStamp(filePath);
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or UnauthorizedAccessException) { return false; }
    }

    /// <summary>Writes the package next to its final path and moves it into place.</summary>
    internal static void Translate(string filePath, string packagePath, CancellationToken cancellation)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(packagePath)!);
        string temporaryPath = packagePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            string stamp = SourceStamp(filePath);
            using (var output = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                Translate(() => WordDocumentFileStream.OpenShared(filePath), output, stamp, cancellation);
            try
            {
                if (File.Exists(packagePath)) File.Delete(packagePath);
                File.Move(temporaryPath, packagePath);
            }
            catch (IOException) when (IsCurrent(packagePath, filePath)) { }   // another open made the same package first
        }
        finally
        {
            try { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
    }

    /// <summary>The translation itself: <paramref name="openSource"/> is called once per pass.</summary>
    internal static void Translate(Func<Stream> openSource, Stream output, string sourceStamp, CancellationToken cancellation)
    {
        Header header = ReadHeader(openSource);
        var state = new TranslationState();
        using var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true);

        WriteEntry(zip, "word/document.xml", writer => TranslateBody(openSource, writer, state, NotePass.None, cancellation));
        if (header.Styles is null && header.DefaultFonts is not null) header.Styles = new XElement(XName.Get("styles", WordNamespace));
        if (header.Styles is not null)
        {
            // the document's default fonts lead the styles part as docDefaults (ECMA-376 17.7.5.1), as Word saves them
            if (header.DefaultFonts is { } fonts)
                header.Styles.AddFirst(new XElement(XName.Get("docDefaults", WordNamespace), new XElement(XName.Get("rPrDefault", WordNamespace),
                    new XElement(XName.Get("rPr", WordNamespace), new XElement(XName.Get("rFonts", WordNamespace), fonts.Attributes())))));
            WriteEntry(zip, "word/styles.xml", writer => TranslateTree(header.Styles, writer, state, "styles"));
        }
        if (header.Lists is not null) WriteEntry(zip, "word/numbering.xml", writer => TranslateTree(header.Lists, writer, state, "numbering"));
        if (header.DocumentProperties is not null) WriteEntry(zip, "word/settings.xml", writer => TranslateTree(header.DocumentProperties, writer, state, "settings"));
        if (state.FootnoteCount > 0) WriteEntry(zip, "word/footnotes.xml", writer => TranslateBody(openSource, writer, new TranslationState(state), NotePass.Footnotes, cancellation));
        if (state.EndnoteCount > 0) WriteEntry(zip, "word/endnotes.xml", writer => TranslateBody(openSource, writer, new TranslationState(state), NotePass.Endnotes, cancellation));
        if (state.ImageRelationships.Count > 0) WriteImages(openSource, zip, state, cancellation);

        WriteText(zip, "[Content_Types].xml", ContentTypes(state, header));
        WriteText(zip, "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">"
            + "<Relationship Id=\"rId1\" Type=\"" + RelationshipNamespace + "/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        string imageRelationships = ImageRelationships(state);
        var document = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">");
        if (header.Styles is not null) document.Append(Relationship("rIdStyles", "styles", "styles.xml"));
        if (header.Lists is not null) document.Append(Relationship("rIdNumbering", "numbering", "numbering.xml"));
        if (header.DocumentProperties is not null) document.Append(Relationship("rIdSettings", "settings", "settings.xml"));
        if (state.FootnoteCount > 0) document.Append(Relationship("rIdFootnotes", "footnotes", "footnotes.xml"));
        if (state.EndnoteCount > 0) document.Append(Relationship("rIdEndnotes", "endnotes", "endnotes.xml"));
        foreach ((string id, string target) in state.Hyperlinks)
            document.Append("<Relationship Id=\"").Append(Escape(id)).Append("\" Type=\"").Append(RelationshipNamespace).Append("/hyperlink\" Target=\"")
                    .Append(Escape(target)).Append("\" TargetMode=\"External\"/>");
        document.Append(imageRelationships).Append("</Relationships>");
        WriteText(zip, "word/_rels/document.xml.rels", document.ToString());
        foreach (string notes in new[] { state.FootnoteCount > 0 ? "footnotes" : null, state.EndnoteCount > 0 ? "endnotes" : null })
            if (notes is not null)
                WriteText(zip, "word/_rels/" + notes + ".xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"" + PackageRelationshipNamespace + "\">"
                                                                   + HyperlinkRelationships(state) + imageRelationships + "</Relationships>");
        ZipArchiveEntry stampEntry = zip.CreateEntry(SourceStampEntry, CompressionLevel.NoCompression);
        stampEntry.LastWriteTime = FixedEntryTime;
        using (Stream stamp = stampEntry.Open()) { byte[] bytes = Encoding.UTF8.GetBytes(sourceStamp); stamp.Write(bytes, 0, bytes.Length); }
    }

    private static string Relationship(string id, string type, string target)
        => "<Relationship Id=\"" + id + "\" Type=\"" + RelationshipNamespace + "/" + type + "\" Target=\"" + target + "\"/>";

    private static string HyperlinkRelationships(TranslationState state)
    {
        var text = new StringBuilder();
        foreach ((string id, string target) in state.Hyperlinks)
            text.Append("<Relationship Id=\"").Append(Escape(id)).Append("\" Type=\"").Append(RelationshipNamespace).Append("/hyperlink\" Target=\"")
                .Append(Escape(target)).Append("\" TargetMode=\"External\"/>");
        return text.ToString();
    }

    private static string ImageRelationships(TranslationState state)
    {
        var text = new StringBuilder();
        foreach (KeyValuePair<string, string> image in state.ImageRelationships)
            text.Append(Relationship(image.Value, "image", "media/" + MediaPartName(state, image.Key)));      // only safe characters
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
    private static string MediaPartName(TranslationState state, string imageName) => state.ImageRelationships[imageName] + "-" + MediaFileName(imageName);

    private static string ContentTypes(TranslationState state, Header header)
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
        if (state.FootnoteCount > 0) text.Append("<Override PartName=\"/word/footnotes.xml\" ContentType=\"").Append(Wordprocessing).Append("footnotes+xml\"/>");
        if (state.EndnoteCount > 0) text.Append("<Override PartName=\"/word/endnotes.xml\" ContentType=\"").Append(Wordprocessing).Append("endnotes+xml\"/>");
        return text.Append("</Types>").ToString();
    }

    private static string ImageContentType(string extension) => extension.ToLowerInvariant() switch
    {
        "png" => "image/png", "jpg" or "jpeg" => "image/jpeg", "gif" => "image/gif", "bmp" => "image/bmp",
        "emf" => "image/x-emf", "wmf" => "image/x-wmf", "tif" or "tiff" => "image/tiff", _ => "application/octet-stream",
    };

    private static void WriteText(ZipArchive zip, string name, string text)
    {
        ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.NoCompression);
        entry.LastWriteTime = FixedEntryTime;
        using Stream stream = entry.Open();
        byte[] bytes = new UTF8Encoding(false).GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteEntry(ZipArchive zip, string name, Action<XmlWriter> write)
    {
        ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.NoCompression);
        entry.LastWriteTime = FixedEntryTime;
        using Stream stream = entry.Open();
        using XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false });
        writer.WriteStartDocument(standalone: true);
        write(writer);
        writer.WriteEndDocument();
    }

    private static XmlReader CreateReader(Stream source)
        => XmlReader.Create(source, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false, IgnoreComments = true, IgnoreProcessingInstructions = true });

    // ── header: w:styles, w:lists, w:docPr (small; read as trees) ─────────────────────────────

    private sealed class Header
    {
        public XElement? Styles, Lists, DocumentProperties;
        public XElement? DefaultFonts;           // w:fonts/w:defaultFonts: Word 2007's docDefaults run fonts
    }

    private static Header ReadHeader(Func<Stream> openSource)
    {
        var header = new Header();
        using Stream source = openSource();
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

    private enum NotePass { None, Footnotes, Endnotes }

    private sealed class TranslationState
    {
        public int FootnoteCount, EndnoteCount;                   // notes met so far (the ids: 1, 2, ... in document order)
        public readonly Dictionary<string, string> ImageRelationships;   // "wordml://..." -> relationship id
        public readonly List<(string id, string target)> Hyperlinks;
        private readonly Dictionary<string, string> _hyperlinkIds;
        public TranslationState() { ImageRelationships = new(StringComparer.Ordinal); Hyperlinks = new(); _hyperlinkIds = new(StringComparer.Ordinal); }
        /// <summary>A notes pass: the body pass's relationships (one id per link target, per picture), note counts from 0.</summary>
        public TranslationState(TranslationState body) { ImageRelationships = body.ImageRelationships; Hyperlinks = body.Hyperlinks; _hyperlinkIds = body._hyperlinkIds; }

        public string HyperlinkId(string target)
        {
            if (_hyperlinkIds.TryGetValue(target, out string? id)) return id;
            _hyperlinkIds[target] = id = "rIdLink" + (Hyperlinks.Count + 1);
            Hyperlinks.Add((id, target));
            return id;
        }

        public string ImageId(string name)
        {
            if (!ImageRelationships.TryGetValue(name, out string? id)) ImageRelationships[name] = id = "rIdImage" + (ImageRelationships.Count + 1);
            return id;
        }
    }

    private static void TranslateTree(XElement element, XmlWriter writer, TranslationState state, string rootName)
    {
        using XmlReader reader = element.CreateReader();
        reader.MoveToContent();
        new Translation(writer, state, NotePass.None).Root(reader, rootName);
    }

    /// <summary>One pass over w:body: the document part, or (a notes pass) only the notes' bodies.</summary>
    private static void TranslateBody(Func<Stream> openSource, XmlWriter writer, TranslationState state, NotePass pass, CancellationToken cancellation)
    {
        using Stream source = openSource();
        using XmlReader reader = CreateReader(source);
        var translation = new Translation(writer, state, pass) { Cancellation = cancellation };
        string rootName = pass switch { NotePass.Footnotes => "footnotes", NotePass.Endnotes => "endnotes", _ => "document" };
        translation.StartRoot(rootName);
        try
        {
            reader.MoveToContent();
            if (!reader.IsEmptyElement)
            {
                reader.Read();
                while (!reader.EOF && reader.NodeType != XmlNodeType.EndElement)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == WordNamespace && reader.LocalName == "body")
                    {
                        translation.Element(reader);      // w:body (translated as itself in the document pass)
                        break;
                    }
                    if (reader.NodeType == XmlNodeType.Element) reader.Skip(); else reader.Read();
                }
            }
        }
        catch (XmlException)
        {
            // Malformed: keep what precedes the first error, as for a .docx part (WordDocumentXml.LoadUntilFirstError):
            // the elements still open are closed below, the rest of the file is not read.
        }
        translation.CloseAll();
        writer.WriteEndElement();         // the root
    }

    /// <summary>Every w:binData a v:imagedata names, decoded into word/media.</summary>
    private static void WriteImages(Func<Stream> openSource, ZipArchive zip, TranslationState state, CancellationToken cancellation)
    {
        var written = new HashSet<string>(StringComparer.Ordinal);
        using Stream source = openSource();
        using XmlReader reader = CreateReader(source);
        var buffer = new byte[1 << 16];
        try
        {
            while (!reader.EOF)
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "binData" || reader.NamespaceURI != WordNamespace) { reader.Read(); continue; }
                string? name = reader.GetAttribute("name", WordNamespace);
                if (name is null || !state.ImageRelationships.ContainsKey(name) || !written.Add(name) || reader.IsEmptyElement) { reader.Skip(); continue; }
                cancellation.ThrowIfCancellationRequested();
                ZipArchiveEntry entry = zip.CreateEntry("word/media/" + MediaPartName(state, name), CompressionLevel.NoCompression);
                entry.LastWriteTime = FixedEntryTime;
                using Stream media = entry.Open();
                int read;
                while ((read = reader.ReadElementContentAsBase64(buffer, 0, buffer.Length)) > 0) media.Write(buffer, 0, read);
            }
        }
        catch (XmlException) { }          // malformed: the pictures before the error are kept
        catch (FormatException) { }       // bad base64: that picture is cut where it breaks
    }

    /// <summary>The element-by-element translation (an explicit stack: no recursion, any nesting depth).</summary>
    private sealed class Translation(XmlWriter writer, TranslationState state, NotePass pass)
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
                int id = isFootnote ? ++state.FootnoteCount : ++state.EndnoteCount;
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
                case "listPicture": reader.Skip(); return;                     // picture bullets' pictures (not translated)
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
