using System.Text;
using System.Xml;
using System.Xml.Linq;
using SeekableDeflate;

namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// What kind of OpenDocument text file a stream holds and where its parts are.
///   • .odt / .ott — a zip package; the stored `mimetype` entry says
///     application/vnd.oasis.opendocument.text (or -template) (ODF 1.3 part 2 §3.3: the first
///     file of the zip, stored). content.xml holds the body, styles.xml the named styles.
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

    /// <summary>True when the zip carries an OpenDocument text (or text template) mimetype entry.</summary>
    public static bool IsOpenDocumentTextPackage(Stream zip)
    {
        string? mimeType = ReadMimeType(zip);
        return mimeType is TextMimeType or TextTemplateMimeType;
    }

    /// <summary>The `mimetype` entry's content, or null (not an ODF package, or not stored as the spec requires).</summary>
    public static string? ReadMimeType(Stream zip)
    {
        ZipEntryLocation? entry;
        try { entry = ZipEntryLocator.Locate(zip, "mimetype"); }
        catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException) { return null; }
        if (entry is null || !entry.IsStored || entry.CompressedSize > 256) return null;
        var bytes = new byte[entry.CompressedSize];
        zip.Position = entry.DataOffset;
        int filled = 0;
        while (filled < bytes.Length)
        {
            int read = zip.Read(bytes, filled, bytes.Length - filled);
            if (read <= 0) return null;
            filled += read;
        }
        return Encoding.ASCII.GetString(bytes).Trim();
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
        using var reader = XmlReader.Create(view, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
        string? currentPath = null;
        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.LocalName == "file-entry") currentPath = reader.GetAttribute("full-path", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0");
            else if (reader.LocalName == "encryption-data" && currentPath == ContentEntryName)
                throw new NotSupportedException("password-protected OpenDocument is not supported");
        }
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
        using var reader = XmlReader.Create(view, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore, CloseInput = false });
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
        return sections;
    }
}
