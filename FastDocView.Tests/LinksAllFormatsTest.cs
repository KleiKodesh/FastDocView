using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Bookmarks and in-document links in OpenDocument and RTF (the .docx checks are in DocxFidelityTest.Links, .doc in
/// DocBinaryLinksTest): bookmark lookup on a complete and on an on-demand index (and after a reopen from the
/// saved index), an internal link, a cross-reference, an app link kept for the host, javascript: never a link.
/// ODF: ODF 1.3 part 3 6.1.8 (text:a), 6.2.1 / 6.2.2 (bookmarks, reference marks), 7.7.5 / 7.7.6 (reference-ref, bookmark-ref).
/// RTF: RTF 1.5 spec "Bookmarks" and "Fields"; the HYPERLINK / REF \h instruction syntax is Word's field syntax.
/// </summary>
public static class LinksAllFormatsTest
{
    private const int FillerParagraphs = 30000;      // puts the late bookmark well past the on-demand open's first chunk

    public static void Run()
    {
        RunOdt();
        RunRtf();
    }

    private static void RunOdt()
    {
        TestReport.Step("Links: OpenDocument bookmarks, text:a, bookmark-ref / reference-ref");
        var body = new StringBuilder();
        body.Append("<text:p><text:bookmark text:name=\"Early\"/>early</text:p>");
        body.Append("<text:p><text:a xlink:href=\"#Late\">to late</text:a> <text:a xlink:href=\"otzaria://open/book/12\">in otzaria</text:a> ");
        body.Append("<text:a xlink:href=\"javascript:alert(1)\">bad</text:a> <text:a xlink:href=\"https://example.org/x\">web</text:a></text:p>");
        body.Append("<text:p>see <text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"Late\">late</text:bookmark-ref>");
        body.Append(" and <text:reference-ref text:reference-format=\"text\" text:ref-name=\"Mark\">mark</text:reference-ref></text:p>");
        for (int i = 0; i < FillerParagraphs; i++) body.Append("<text:p>filler paragraph number ").Append(i).Append("</text:p>");
        body.Append("<text:p><text:bookmark-start text:name=\"Late\"/>late target<text:bookmark-end text:name=\"Late\"/></text:p>");
        body.Append("<text:p><text:reference-mark-start text:name=\"Mark\"/>mark target<text:reference-mark-end text:name=\"Mark\"/></text:p>");
        string content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
                       + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
                       + "xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\"><office:body><office:text>" + body + "</office:text></office:body></office:document-content>";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-links.odt");
        if (File.Exists(path)) File.Delete(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            void Add(string name, string text, CompressionLevel level)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name, level).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
            Add("content.xml", content, CompressionLevel.Optimal);
        }
        const int lateBlock = 3 + FillerParagraphs, markBlock = lateBlock + 1;
        CheckLookups("odt", path, lateBlock, markBlock, expectsOnDemandScan: true);

        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string links = document.RenderBlocks(1, 1)[0].Html, refs = document.RenderBlocks(2, 1)[0].Html;
        TestReport.Check(links.Contains("<a class=\"link\" href=\"#bookmark-Late\">to late</a>"), "odt: text:a #name is an in-document link (" + links + ")");
        TestReport.Check(links.Contains("class=\"link app-link\"") && links.Contains("href=\"otzaria://open/book/12\""), "odt: an otzaria:// link is kept as an app link");
        TestReport.Check(!links.Contains("javascript") && links.Contains("bad"), "odt: javascript: is not a link, its text shows");
        TestReport.Check(links.Contains("href=\"https://example.org/x\"") && links.Contains("target=\"_blank\""), "odt: an https link opens outside");
        TestReport.Check(refs.Contains("href=\"#bookmark-Late\">late</a>") && refs.Contains("href=\"#bookmark-Mark\">mark</a>"), "odt: bookmark-ref and reference-ref link to their targets (" + refs + ")");
        TestReport.Check(document.RenderBlocks(0, 1)[0].Html.Contains("id=\"bookmark-Early\"") && document.RenderBlocks(lateBlock, 1)[0].Html.Contains("id=\"bookmark-Late\""),
                         "odt: bookmarks render as their element ids");
    }

    private static void RunRtf()
    {
        TestReport.Step("Links: RTF \\bkmkstart, HYPERLINK / REF \\h fields");
        var rtf = new StringBuilder(@"{\rtf1\ansi{\fonttbl{\f0 Arial;}}" + "\n");
        rtf.Append(@"\pard {\*\bkmkstart Early}early{\*\bkmkend Early}\par" + "\n");
        rtf.Append(@"\pard {\field{\*\fldinst HYPERLINK \\l ""Late"" }{\fldrslt to late}} {\field{\*\fldinst HYPERLINK ""otzaria://open/book/12"" }{\fldrslt in otzaria}} ");
        rtf.Append(@"{\field{\*\fldinst HYPERLINK ""javascript:alert(1)"" }{\fldrslt bad}} {\field{\*\fldinst {HYPERLINK ""https://example.org/x"" }}{\fldrslt {web}}}\par" + "\n");
        rtf.Append(@"\pard see {\field{\*\fldinst  REF Late \\h }{\fldrslt late}} and {\field{\*\fldinst  PAGEREF Late }{\fldrslt 7}}\par" + "\n");
        for (int i = 0; i < 2000; i++) rtf.Append(@"\pard filler ").Append(i).Append(@"\par").Append('\n');
        rtf.Append(@"\pard {\*\bkmkstart Late}late target{\*\bkmkend Late}\par" + "\n}");
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-links.rtf");
        File.WriteAllText(path, rtf.ToString());
        const int lateBlock = 3 + 2000;
        CheckLookups("rtf", path, lateBlock, null, expectsOnDemandScan: false);

        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string links = document.RenderBlocks(1, 1)[0].Html, refs = document.RenderBlocks(2, 1)[0].Html;
        TestReport.Check(links.Contains("<a class=\"link\" href=\"#bookmark-Late\">to late</a>"), "rtf: HYPERLINK \\l is an in-document link (" + links + ")");
        TestReport.Check(links.Contains("class=\"link app-link\"") && links.Contains("href=\"otzaria://open/book/12\""), "rtf: an otzaria:// link is kept as an app link");
        TestReport.Check(!links.Contains("javascript") && links.Contains("bad"), "rtf: javascript: is not a link, its text shows");
        TestReport.Check(links.Contains("href=\"https://example.org/x\"") && links.Contains(">web</a>"), "rtf: an https link (instruction in a nested group)");
        TestReport.Check(refs.Contains("href=\"#bookmark-Late\">late</a>") && !refs.Contains(">7</a>"), "rtf: REF \\h links, PAGEREF without \\h does not (" + refs + ")");
        TestReport.Check(document.RenderBlocks(0, 1)[0].Html.Contains("id=\"bookmark-Early\"") && document.RenderBlocks(lateBlock, 1)[0].Html.Contains("id=\"bookmark-Late\""),
                         "rtf: bookmarks render as their element ids");
    }

    /// <summary>FindBookmark on a complete index, on an on-demand one (scanning on only until found), and after a reopen.</summary>
    internal static void CheckLookups(string format, string path, int lateBlock, int? markBlock, bool expectsOnDemandScan)
    {
        foreach (bool onDemand in new[] { false, true })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-links-" + format + "-" + onDemand),
                                               ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = onDemand };
            string sidecar = policy.IndexPathFor(path);
            if (File.Exists(sidecar)) File.Delete(sidecar);
            string kind = onDemand ? "on-demand" : "complete";
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
            {
                if (onDemand && expectsOnDemandScan) TestReport.Check(!document.IsIndexComplete, $"{format}: the on-demand open has not indexed the whole body");
                TestReport.Check(document.FindBookmark("Early") == 0, $"{format}: early bookmark lookup ({kind} index)");
                if (onDemand && expectsOnDemandScan) TestReport.Check(!document.IsIndexComplete, $"{format}: an early bookmark does not scan the rest");
                TestReport.Check(document.FindBookmark("Late") == lateBlock, $"{format}: late bookmark lookup ({kind} index): {document.FindBookmark("Late")} vs {lateBlock}");
                if (markBlock is { } mark) TestReport.Check(document.FindBookmark("Mark") == mark, $"{format}: reference mark lookup ({kind} index)");
                TestReport.Check(document.FindBookmark("missing") is null, $"{format}: unknown bookmark is null ({kind} index)");
            }
            using (IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(path, policy))
                TestReport.Check(reopened.FindBookmark("Late") == lateBlock, $"{format}: bookmark lookup after a reopen from the saved index ({kind})");
        }
    }
}
