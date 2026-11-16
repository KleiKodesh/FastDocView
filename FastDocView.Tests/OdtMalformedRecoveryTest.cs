using System.IO.Compression;
using System.Text;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// A malformed content.xml / styles.xml (.odt and .fodt) keeps the blocks before the first XML error, as a malformed .docx
/// part does (WordDocumentXml.LoadUntilFirstError; LibreOffice: libxml2 / the fast parser stop at the first fatal error and
/// keep every element delivered before it - sax/source/fastparser/fastparser.cxx parseStream rethrows the SAXParseException
/// after the elements already handed to xmloff), and the rest of the body becomes the "could not be read" placeholder
/// (<see cref="WordDocumentBlockRenderGuard.PlaceholderHtml"/>) instead of failing the whole file. Each case is opened in
/// memory, as a sidecar and on demand.
/// </summary>
public static class OdtMalformedRecoveryTest
{
    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
        + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\"";

    private const string GoodStyles = "<office:styles><style:style style:name=\"Bold\" style:family=\"paragraph\"><style:text-properties fo:font-weight=\"bold\"/></style:style></office:styles>";

    private static string Paragraphs(int count) => string.Concat(Enumerable.Range(1, count).Select(i => $"<text:p text:style-name=\"Bold\">Paragraph {i}</text:p>"));

    // (case, body, styles, paragraphs shown, blocks, last block the placeholder, a .fodt fails instead)
    private static readonly (string name, string body, string styles, int keptParagraphs, int blocks, bool placeholder, bool flatFails)[] Cases =
    [
        // the first fatal error ends the body: the blocks before it, then the placeholder
        ("mismatched end tag in paragraph 3", Paragraphs(2) + "<text:p>Paragraph 3</text:span></text:p>" + Paragraphs(2), GoodStyles, 2, 3, true, false),
        ("body cut off inside paragraph 4", Paragraphs(3) + "<text:p>Paragraph 4 is cut", GoodStyles, 3, 4, true, false),
        ("body cut off inside paragraph 1", "<text:p>Paragraph 1 is cut", GoodStyles, 0, 1, true, false),
        // errors the block scan does not see are handled per block, as a .docx block is: the block keeps what precedes the error
        ("illegal character between paragraphs", Paragraphs(2) + "" + Paragraphs(2), GoodStyles, 4, 4, false, false),
        ("duplicate attribute in paragraph 2", Paragraphs(1) + "<text:p text:style-name=\"Bold\" text:style-name=\"Bold\">x</text:p>" + Paragraphs(2), GoodStyles, 3, 4, false, false),
        // styles.xml keeps the styles before its error; in a .fodt the same error comes before the body, so nothing is left to show
        ("malformed styles", Paragraphs(4), GoodStyles.Replace("</office:styles>", "<style:style></office:styles>"), 4, 4, false, true),
    ];

    public static void Run()
    {
        TestReport.Step("ODT malformed XML: blocks before the first error kept, the rest a placeholder (.odt and .fodt)");
        string folder = Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestOutput", "odt-malformed")).FullName;
        foreach ((string name, string body, string styles, int kept, int expectedBlocks, bool placeholder, bool flatFails) in Cases)
        foreach (bool isFlat in new[] { false, true })
        {
            string path = Path.Combine(folder, Slug(name) + (isFlat ? ".fodt" : ".odt"));
            if (isFlat) WriteFlat(path, body, styles); else WritePackage(path, body, styles);
            string cache = Path.Combine(folder, "cache-" + Slug(name) + (isFlat ? "-flat" : ""));
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
            foreach ((string mode, WordIndexPolicy policy) in new[]
            {
                ("in memory", new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = cache }),
                ("sidecar", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache }),
                ("on demand", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache + "-demand", IndexesOnDemand = true }),
            })
            {
                string label = $"{name} ({(isFlat ? ".fodt" : ".odt")}, {mode})";
                try
                {
                    using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                    if (!document.IsIndexComplete) document.ScanAhead(1000);
                    var blocks = document.RenderBlocks(0, document.BlockCount);
                    int paragraphs = blocks.Count(block => block.Html.Contains("Paragraph "));
                    bool hasPlaceholder = blocks.Count > 0 && blocks[blocks.Count - 1].Html == WordDocumentBlockRenderGuard.PlaceholderHtml;
                    bool bold = blocks.Where(block => block.Html.Contains("Paragraph 1")).All(block => block.Html.Contains("bold"));
                    TestReport.Check(!(isFlat && flatFails) && paragraphs == kept && hasPlaceholder == placeholder && blocks.Count == expectedBlocks
                                     && document.BlockCount == expectedBlocks && (bold || styles != GoodStyles),
                                     $"{label}: {paragraphs} paragraphs (expected {kept}), placeholder {hasPlaceholder} (expected {placeholder}), {blocks.Count} blocks (expected {expectedBlocks})");
                }
                catch (Exception exception) when (isFlat && flatFails)
                {
                    TestReport.Check(exception is InvalidDataException && exception.Message.Contains("not well-formed"), $"{label}: fails with a clear error - {exception.Message}");
                }
                catch (Exception exception)
                {
                    TestReport.Check(false, $"{label}: opens - {exception.GetType().Name}: {exception.Message}");
                }
            }
        }
    }

    private static string Slug(string name) => new(name.Select(c => char.IsLetterOrDigit(c) ? c : '-').ToArray());

    private static void WritePackage(string path, string body, string styles)
    {
        if (File.Exists(path)) File.Delete(path);
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        Write(zip, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
        Write(zip, "META-INF/manifest.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">"
            + "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>"
            + "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>");
        Write(zip, "styles.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles {Namespaces}>{styles}</office:document-styles>");
        Write(zip, "content.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {Namespaces}><office:body><office:text>{body}</office:text></office:body></office:document-content>");
    }

    private static void WriteFlat(string path, string body, string styles)
        => File.WriteAllText(path, $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document {Namespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                                   + $"{styles}<office:body><office:text>{body}</office:text></office:body></office:document>", new UTF8Encoding(false));

    private static void Write(ZipArchive zip, string name, string text, CompressionLevel level = CompressionLevel.Optimal)
    {
        using Stream stream = zip.CreateEntry(name, level).Open();
        byte[] bytes = new UTF8Encoding(false).GetBytes(text);
        stream.Write(bytes, 0, bytes.Length);
    }
}
