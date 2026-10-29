using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// The first open of an .odt whose styles.xml is huge (a LibreOffice conversion of a Word file carries thousands of list
/// styles: 15 MB, 3,058 text:list-style for the sample): time to the first window on a cold, on-demand open (no sidecar).
/// Synthetic file in --fidelity (same shape: 3,000 list styles of 10 levels); real files with --odt-first-open path...
/// </summary>
public static class OdtFirstOpenTiming
{
    private const int ListStyles = 3000;

    public static void RunSynthetic()
    {
        TestReport.Step($"OpenDocument first open with a huge styles.xml ({ListStyles} list styles)");
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-huge-styles.odt");
        if (!File.Exists(path)) Generate(path);
        (double milliseconds, string html) = Measure(path);
        TestReport.Check(html.Contains("1.") && html.Contains("item 0"), "huge styles: the list label is resolved from styles.xml (" + html.Substring(0, Math.Min(200, html.Length)) + ")");
        TestReport.Check(milliseconds < 300, $"huge styles: first window in {milliseconds:0} ms (< 300 ms)");

        // the byte-level list-style reader with other prefixes, an entity in a value, a bullet level and a self-closing style
        string prefixed = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-styles-prefixes.odt");
        if (File.Exists(prefixed)) File.Delete(prefixed);
        const string other = "xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:s=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
                           + "xmlns:t=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:f=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"";
        using (var zip = ZipFile.Open(prefixed, ZipArchiveMode.Create))
        {
            void Add(string name, string text)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("mimetype", "application/vnd.oasis.opendocument.text");
            Add("styles.xml", "<?xml version=\"1.0\"?><o:document-styles " + other + "><o:styles><t:list-style s:name=\"Empty\"/><t:list-style-other/>"
                + "<t:list-style s:name=\"L&amp;1\"><t:list-level-style-number t:level=\"1\" s:num-suffix=\"&amp;)\" s:num-format=\"a\"/>"
                + "<t:list-level-style-bullet t:level=\"2\" t:bullet-char=\"*\"/></t:list-style></o:styles></o:document-styles>");
            Add("content.xml", "<?xml version=\"1.0\"?><o:document-content " + other + "><o:body><o:text><t:list t:style-name=\"L&amp;1\"><t:list-item><t:p>one</t:p>"
                + "<t:list><t:list-item><t:p>two</t:p></t:list-item></t:list></t:list-item></t:list></o:text></o:body></o:document-content>");
        }
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(prefixed))
        {
            while (!document.IsIndexComplete) document.ScanAhead(1000);
            string first = document.RenderBlocks(0, 1)[0].Html, second = document.RenderBlocks(1, 1)[0].Html;
            TestReport.Check(first.Contains(">a&amp;)<") && second.Contains(">*<"), "list styles read from bytes: other prefixes, entities, bullet level (" + first + " / " + second + ")");
        }
    }

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"OpenDocument first open: {Path.GetFileName(path)}");
            Measure(path);
        }
    }

    /// <summary>Best of three cold on-demand opens + the first 40 blocks; prints every run.</summary>
    private static (double milliseconds, string firstHtml) Measure(string path)
    {
        double best = double.MaxValue;
        string html = "";
        for (int run = 1; run <= 3; run++)
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-odt-first-open"),
                                               ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
            string sidecar = policy.IndexPathFor(path);
            if (File.Exists(sidecar)) File.Delete(sidecar);
            var watch = Stopwatch.StartNew();
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            double opened = watch.Elapsed.TotalMilliseconds;
            html = document.RenderBlocks(0, Math.Min(40, document.BlockCount))[0].Html;
            double shown = watch.Elapsed.TotalMilliseconds;
            Console.WriteLine($"   run {run}: open {opened:0.#} ms | first 40 blocks shown at {shown:0.#} ms");
            best = Math.Min(best, shown);
        }
        return (best, html);
    }

    private static void Generate(string path)
    {
        const string namespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
                                + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\"";
        var styles = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles " + namespaces + "><office:styles>");
        styles.Append("<style:default-style style:family=\"paragraph\"><style:text-properties fo:font-size=\"12pt\"/></style:default-style>");
        styles.Append("<style:style style:name=\"Standard\" style:family=\"paragraph\"/>");
        for (int list = 0; list < ListStyles; list++)
        {
            styles.Append("<text:list-style style:name=\"WWNum").Append(list).Append("\">");
            for (int level = 1; level <= 10; level++)
                styles.Append("<text:list-level-style-number text:level=\"").Append(level).Append("\" style:num-suffix=\".\" style:num-format=\"1\">")
                      .Append("<style:list-level-properties text:list-level-position-and-space-mode=\"label-alignment\"><style:list-level-label-alignment ")
                      .Append("text:label-followed-by=\"listtab\" text:list-tab-stop-position=\"").Append(level * 0.5).Append("in\" fo:text-indent=\"-0.25in\" fo:margin-left=\"")
                      .Append(level * 0.5).Append("in\"/></style:list-level-properties><style:text-properties style:font-name=\"Arial\" fo:font-size=\"11pt\"/></text:list-level-style-number>");
            styles.Append("</text:list-style>");
        }
        styles.Append("</office:styles><office:master-styles/></office:document-styles>");
        var content = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content " + namespaces + "><office:body><office:text>");
        content.Append("<text:list text:style-name=\"WWNum2999\">");
        for (int i = 0; i < 200; i++) content.Append("<text:list-item><text:p text:style-name=\"Standard\">item ").Append(i).Append("</text:p></text:list-item>");
        content.Append("</text:list></office:text></office:body></office:document-content>");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text, CompressionLevel level)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name, level).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
        Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
        Add("styles.xml", styles.ToString(), CompressionLevel.Optimal);
        Add("content.xml", content.ToString(), CompressionLevel.Optimal);
    }
}
