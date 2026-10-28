using System.Text;
using System.Xml;
using System.Xml.Linq;
using SeekableDeflate;

namespace WordDocumentRandomAccess.Odt;

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
    private const int FlatSniffBytes = 8192;
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
            return Encoding.ASCII.GetString(bytes, 0, filled).Trim();
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

    /// <summary>True when the file is a flat OpenDocument text: XML whose root is office:document with a text mimetype.</summary>
    public static bool IsFlatOpenDocumentText(Stream file)
    {
        var head = new byte[FlatSniffBytes];
        file.Position = 0;
        int count = file.Read(head, 0, head.Length);
        string text = Encoding.UTF8.GetString(head, 0, count);
        int root = text.IndexOf("<office:document ", StringComparison.Ordinal);
        if (root < 0) root = text.IndexOf("<office:document>", StringComparison.Ordinal);
        return root >= 0 && (text.IndexOf(TextMimeType + "\"", root, StringComparison.Ordinal) >= 0
                          || text.IndexOf(TextTemplateMimeType + "\"", root, StringComparison.Ordinal) >= 0);
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

    /// <summary>
    /// styles.xml's office:font-face-decls and office:styles (named styles, default styles, list
    /// styles, notes configuration), streamed with XmlReader — page layouts, master pages and
    /// header/footer content are skipped without being built.
    /// </summary>
    public static List<XElement> ReadStylesSections(Stream zip)
    {
        var sections = new List<XElement>();
        ZipEntryLocation? entry = ZipEntryLocator.Locate(zip, StylesEntryName);
        if (entry is null) return sections;
        using Stream view = OpenEntry(zip, entry);
        using XmlReader reader = WordDocumentXml.CreateReader(view, tolerateUndeclaredPrefixes: true);
        try
        {
            reader.MoveToContent();
            if (!reader.Read()) return sections;
            while (!reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Depth == 1)
                {
                    if (reader.LocalName is "font-face-decls" or "styles") { sections.Add((XElement)XNode.ReadFrom(reader)); continue; }
                    reader.Skip();
                    continue;
                }
                reader.Read();
            }
        }
        catch (XmlException exception) { throw WordDocumentXml.Malformed(StylesEntryName, exception); }
        return sections;
    }
}
