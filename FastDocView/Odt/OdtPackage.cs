using System.Text;
using System.Xml;
using System.Xml.Linq;
using SeekableDeflate;

namespace FastDocView.Odt;

/// <summary>
/// What kind of OpenDocument text file a stream holds and where its parts are.
///   • .odt / .ott — a zip package whose `mimetype` entry (or, without one, the manifest's root
///     entry) says application/vnd.oasis.opendocument.text (or -template). content.xml holds the body, styles.xml the named styles.
///   • .fodt — one plain XML file whose root is office:document; body offsets are file offsets.
/// A package whose manifest lists encryption-data for content.xml is password-protected and refused.
/// </summary>
public static class OdtPackage
{
    public const string ContentEntryName = "content.xml";
    public const string StylesEntryName = "styles.xml";
    public const string ManifestEntryName = "META-INF/manifest.xml";
    public const string TextMimeType = "application/vnd.oasis.opendocument.text";
    public const string TextTemplateMimeType = "application/vnd.oasis.opendocument.text-template";
    private const int MaxManifestBytes = 4 << 20;
    private const string OfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private const int MimeTypeReadBytes = 1024;
    private const string ManifestNamespace = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";

    /// <summary>True when the zip carries an OpenDocument text (or text template) mimetype entry.</summary>
    public static bool IsOpenDocumentTextPackage(Stream zip)
    {
        string? mimeType = ReadMimeType(zip);
        return mimeType is TextMimeType or TextTemplateMimeType;
    }

    /// <summary>
    /// The package's media type, or null when the zip is not an ODF package. Read the way
    /// LibreOffice's ZipPackage::parseManifest does (package/source/zippackage/ZipPackage.cxx): the
    /// `mimetype` entry is looked up by name — no check of its position or compression (ODF 1.3
    /// part 2 §3.3 asks writers to store it first and uncompressed, but real files deflate it or put
    /// it later) — and its first 1024 bytes are the type (LibreOffice's readBytes( aData, 1024 ));
    /// without a `mimetype` entry the manifest's root file-entry (manifest:full-path="/",
    /// ODF 1.3 part 2, manifest) names it.
    /// </summary>
    public static string? ReadMimeType(Stream zip)
    {
        ZipEntryLocation? entry;
        try { entry = ZipEntryLocator.Locate(zip, "mimetype"); }
        catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException) { return null; }
        if (entry is null) return ReadManifestRootMediaType(zip);
        if (!entry.IsStored && !entry.IsDeflate) return null;
        try
        {
            using Stream view = OpenEntry(zip, entry);
            var bytes = new byte[MimeTypeReadBytes];
            int filled = 0;
            while (filled < bytes.Length)
            {
                int read = view.Read(bytes, filled, bytes.Length - filled);
                if (read <= 0) break;
                filled += read;
            }
            string mimeType = Encoding.ASCII.GetString(bytes, 0, filled).Trim();
            return mimeType.Length > 0 ? mimeType : ReadManifestRootMediaType(zip);   // an empty mimetype says nothing (LibreOffice: the manifest's type wins anyway)
        }
        catch (Exception exception) when (exception is InvalidDataException or IOException) { return null; }
    }

    /// <summary>manifest:media-type of the manifest's root file-entry (full-path "/"), or null.</summary>
    private static string? ReadManifestRootMediaType(Stream zip)
    {
        try
        {
            ZipEntryLocation? manifest = ZipEntryLocator.Locate(zip, ManifestEntryName);
            if (manifest is null || (!manifest.IsStored && !manifest.IsDeflate)) return null;
            using Stream view = OpenEntry(zip, manifest);
            using XmlReader reader = WordDocumentXml.CreateReader(view, tolerateUndeclaredPrefixes: true);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "file-entry") continue;
                if (reader.GetAttribute("full-path", ManifestNamespace) == "/") return reader.GetAttribute("media-type", ManifestNamespace)?.Trim();
            }
            return null;
        }
        catch (Exception exception) when (exception is InvalidDataException or IOException or XmlException) { return null; }
    }

    /// <summary>
    /// True when the file is a flat OpenDocument text: XML whose root element is office:document
    /// (in the office namespace, whatever its prefix, and however its attributes are laid out over
    /// lines) with an office:mimetype naming a text or text template. Only the root start tag is read.
    /// </summary>
    public static bool IsFlatOpenDocumentText(Stream file)
    {
        file.Position = 0;
        try
        {
            using XmlReader reader = WordDocumentXml.CreateReader(file, tolerateUndeclaredPrefixes: true);
            if (reader.MoveToContent() != XmlNodeType.Element || reader.LocalName != "document" || reader.NamespaceURI != OfficeNamespace) return false;
            string? mimeType = reader.GetAttribute("mimetype", OfficeNamespace)?.Trim();
            return mimeType is TextMimeType or TextTemplateMimeType;
        }
        catch (XmlException) { return false; }           // not XML (or broken before its root tag): not a flat OpenDocument
    }

    /// <summary>content.xml of a package; throws when the package is password-protected or has no content.xml.</summary>
    public static ZipEntryLocation LocateContentEntry(Stream zip)
    {
        ThrowWhenEncrypted(zip);
        ZipEntryLocation entry = ZipEntryLocator.Locate(zip, ContentEntryName)
                                 ?? throw new InvalidDataException("Not an OpenDocument text (no content.xml).");
        if (!entry.IsDeflate && !entry.IsStored) throw new InvalidDataException($"Unsupported zip compression method {entry.CompressionMethod} for content.xml.");
        return entry;
    }

    /// <summary>
    /// ODF 1.3 part 2 §4: an encrypted file's manifest:file-entry carries a manifest:encryption-data
    /// child. The manifest is a small listing of the package's files (bounded read).
    /// </summary>
    public static void ThrowWhenEncrypted(Stream zip)
    {
        ZipEntryLocation? manifest = ZipEntryLocator.Locate(zip, ManifestEntryName);
        if (manifest is null || manifest.UncompressedSize > MaxManifestBytes) return;
        using Stream view = OpenEntry(zip, manifest);
        using XmlReader reader = WordDocumentXml.CreateReader(view, tolerateUndeclaredPrefixes: true);
        string? currentPath = null;
        try
        {
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.LocalName == "file-entry") currentPath = reader.GetAttribute("full-path", ManifestNamespace);
                else if (reader.LocalName == "encryption-data" && currentPath == ContentEntryName)
                    throw new NotSupportedException("password-protected OpenDocument is not supported");
            }
        }
        catch (XmlException exception) { throw WordDocumentXml.Malformed(ManifestEntryName, exception); }
    }

    /// <summary>A forward-only view of a whole entry (small parts only: manifest, styles.xml) — .NET's inflater, no checkpoints needed.</summary>
    public static Stream OpenEntry(Stream zip, ZipEntryLocation entry)
        => entry.IsDeflate
            ? new System.IO.Compression.DeflateStream(new ZipStoredEntryStream(zip, entry.DataOffset, entry.CompressedSize), System.IO.Compression.CompressionMode.Decompress)
            : new ZipStoredEntryStream(zip, entry.DataOffset, entry.CompressedSize);

    /// <summary>Page layouts and master pages: styles.xml of a package, or the head of a flat .fodt (read up to office:body).</summary>
    public static OdtPageSetup ReadPageSetup(Stream file, bool isFlat)
    {
        if (isFlat)
        {
            long position = file.Position;
            file.Position = 0;
            try { return OdtPageSetup.Read(new NonClosingStream(file)); }
            finally { file.Position = position; }
        }
        ZipEntryLocation? entry = ZipEntryLocator.Locate(file, StylesEntryName);
        if (entry is null) return OdtPageSetup.Empty;
        using Stream view = OpenEntry(file, entry);
        return OdtPageSetup.Read(view);
    }

    /// <summary>Keeps the caller's stream open when a reader disposes its input.</summary>
    private sealed class NonClosingStream(Stream inner) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    /// <summary>
    /// One streaming pass over styles.xml (XmlReader, no DOM of the whole part) that keeps only what indexing and rendering
    /// read: office:font-face-decls and office:styles become sections of <paramref name="styles"/> (ODF 1.3 part 3 3.14,
    /// 3.15.2) - except their text:list-style elements (16.32), which are cut from the bytes and parsed into
    /// <see cref="OdtListStyle"/> without XmlReader (a LibreOffice conversion of a Word file carries thousands: 15 MB,
    /// 3,058 for the sample; as XML trees they cost most of a second) - and the page layouts / master pages (office:automatic-styles, office:master-styles,
    /// 3.15.3 / 3.15.4) give the page setup. Header / footer content is stepped over, never built.
    /// </summary>
    public static void ReadStyles(Stream zip, OdtStyleCascade styles)
    {
        ZipEntryLocation? entry = ZipEntryLocator.Locate(zip, StylesEntryName);
        if (entry is null) { styles.PageSetup = OdtPageSetup.Empty; return; }
        var pageSetup = new OdtPageSetup.Collector();
        // list styles are cut out of the bytes before the XmlReader sees them (OdtStylesListStyleFilterStream)
        using (Stream view = new OdtStylesListStyleFilterStream(OpenEntry(zip, entry), styles.AddListStyle))
        using (XmlReader reader = WordDocumentXml.CreateReader(view, tolerateUndeclaredPrefixes: true))
        {
            try
            {
                reader.MoveToContent();
                if (reader.Read())
                    while (!reader.EOF)
                    {
                        if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
                        if (reader.Depth == 1 && reader.NamespaceURI == OfficeNamespace && reader.LocalName == "font-face-decls")
                        {
                            styles.AddSection((XElement)XNode.ReadFrom(reader));
                            continue;
                        }
                        if (reader.Depth == 1 && reader.NamespaceURI == OfficeNamespace && reader.LocalName == "styles")
                        {
                            ReadNamedStyles(reader, styles);
                            continue;
                        }
                        pageSetup.Visit(reader);
                        reader.Read();
                    }
            }
            // Malformed styles.xml: keep what was read before the first error, as for .docx parts.
            catch (XmlException) { }
        }
        styles.PageSetup = pageSetup.Finish();
    }

    /// <summary>office:styles: its list styles without a DOM, every other child as a small tree in one section.</summary>
    private static void ReadNamedStyles(XmlReader reader, OdtStyleCascade styles)
    {
        var section = new XElement(XName.Get(reader.LocalName, reader.NamespaceURI));
        try
        {
            if (reader.IsEmptyElement) { reader.Read(); return; }
            int depth = reader.Depth;
            reader.Read();
            while (!reader.EOF && reader.Depth > depth)
            {
                if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
                if (reader.LocalName == "list-style" && reader.NamespaceURI == OdtStyleCascade.Text.NamespaceName)
                {
                    string? name = reader.GetAttribute("name", OdtStyleCascade.Style.NamespaceName);
                    OdtListStyle listStyle = OdtListStyle.Read(reader);
                    if (name is not null) styles.AddListStyle(name, listStyle);
                    continue;
                }
                section.Add(XNode.ReadFrom(reader));
            }
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) reader.Read();
        }
        finally { styles.AddSection(section); }      // what was read, also when the part breaks off
    }
}
