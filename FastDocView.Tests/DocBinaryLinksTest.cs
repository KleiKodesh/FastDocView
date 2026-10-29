using System.Diagnostics;
using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Bookmarks and in-document links in .doc: the file is written by LibreOffice (headless, below-normal priority) from a
/// small OpenDocument file, so Word's own binary structures are produced by a mature writer: bookmarks
/// (SttbfBkmk / PlcfBkf / PlcfBkl), HYPERLINK fields (\l for in-document targets) and REF ... \h cross-references.
/// Skipped when LibreOffice is not installed. Checks: lookup (complete index and after a reopen from the sidecar), the
/// internal link, the cross-reference, an app link kept for the host, javascript: never a link.
/// </summary>
public static class DocBinaryLinksTest
{
    private const string SofficePath = @"C:\Program Files\LibreOffice\program\soffice.exe";
    private const int FillerParagraphs = 300;

    public static void Run()
    {
        TestReport.Step(".doc links: bookmarks, HYPERLINK \\l, REF \\h (file written by LibreOffice)");
        if (!File.Exists(SofficePath)) { TestReport.Step("LibreOffice not installed - .doc links not checked"); return; }
        string directory = Path.Combine(TestReport.OutputDirectory, "doc-links-v1");
        Directory.CreateDirectory(directory);
        string odt = Path.Combine(directory, "doc-links-v1.odt"), doc = Path.Combine(directory, "doc-links-v1.doc");
        if (!File.Exists(doc))
        {
            WriteSource(odt);
            var start = new ProcessStartInfo(SofficePath)
            {
                UseShellExecute = false, CreateNoWindow = true,
                Arguments = $"-env:UserInstallation=file:///{directory.Replace(Path.DirectorySeparatorChar, '/')}/lo-profile --headless --norestore --convert-to doc --outdir \"{directory}\" \"{odt}\"",
            };
            using Process process = Process.Start(start)!;
            try { process.PriorityClass = ProcessPriorityClass.BelowNormal; } catch (InvalidOperationException) { }
            if (!process.WaitForExit(180_000)) { process.Kill(); TestReport.Check(false, "LibreOffice wrote the .doc within 3 minutes"); return; }
        }
        TestReport.Check(File.Exists(doc), "LibreOffice wrote the .doc");
        if (!File.Exists(doc)) return;

        const int lateBlock = 3 + FillerParagraphs;
        foreach (bool reopen in new[] { false, true })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-doc-links"), ForcedMode = WordIndexMode.Sidecar };
            if (!reopen && File.Exists(policy.IndexPathFor(doc))) File.Delete(policy.IndexPathFor(doc));
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(doc, policy);
            string kind = reopen ? "reopened from the saved index" : "complete index";
            TestReport.Check(document.FindBookmark("Early") == 0, $"doc: early bookmark ({kind}): {document.FindBookmark("Early")}");
            TestReport.Check(document.FindBookmark("Late") == lateBlock, $"doc: late bookmark ({kind}): {document.FindBookmark("Late")} vs {lateBlock}");
            TestReport.Check(document.FindBookmark("missing") is null, $"doc: unknown bookmark is null ({kind})");
            if (reopen) continue;

            string links = document.RenderBlocks(1, 1)[0].Html, refs = document.RenderBlocks(2, 1)[0].Html;
            TestReport.Check(links.Contains("href=\"#bookmark-Late\"") && (links.Contains("to late</a>") || links.Contains("to late</span></a>")), "doc: HYPERLINK \\l is an in-document link (" + links + ")");
            TestReport.Check(links.Contains("class=\"link app-link\"") && links.Contains("href=\"otzaria://open/book/12\""), "doc: an otzaria:// link is kept as an app link");
            TestReport.Check(!links.Contains("javascript") && links.Contains("bad"), "doc: javascript: is not a link, its text shows");
            TestReport.Check(links.Contains("href=\"https://example.org/x\""), "doc: an https link");
            TestReport.Check(refs.Contains("href=\"#bookmark-Late\""), "doc: the cross-reference (REF \\h) links to its bookmark (" + refs + ")");
            TestReport.Check(document.RenderBlocks(0, 1)[0].Html.Contains("id=\"bookmark-Early\"") && document.RenderBlocks(lateBlock, 1)[0].Html.Contains("id=\"bookmark-Late\""),
                             "doc: bookmarks render as their element ids");
        }
    }

    private static void WriteSource(string path)
    {
        var body = new StringBuilder();
        body.Append("<text:p><text:bookmark text:name=\"Early\"/>early</text:p>");
        body.Append("<text:p><text:a xlink:type=\"simple\" xlink:href=\"#Late\">to late</text:a> <text:a xlink:type=\"simple\" xlink:href=\"otzaria://open/book/12\">in otzaria</text:a> ");
        body.Append("<text:a xlink:type=\"simple\" xlink:href=\"javascript:alert(1)\">bad</text:a> <text:a xlink:type=\"simple\" xlink:href=\"https://example.org/x\">web</text:a></text:p>");
        body.Append("<text:p>see <text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"Late\">late</text:bookmark-ref></text:p>");
        for (int i = 0; i < FillerParagraphs; i++) body.Append("<text:p>filler paragraph number ").Append(i).Append("</text:p>");
        body.Append("<text:p><text:bookmark-start text:name=\"Late\"/>late target<text:bookmark-end text:name=\"Late\"/></text:p>");
        string content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
                       + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
                       + "xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\"><office:body><office:text>" + body + "</office:text></office:body></office:document-content>";
        if (File.Exists(path)) File.Delete(path);
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string text, CompressionLevel level)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name, level).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
        Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
        Add("content.xml", content, CompressionLevel.Optimal);
        Add("META-INF/manifest.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">"
                                     + "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>"
                                     + "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>", CompressionLevel.Optimal);
    }
}
