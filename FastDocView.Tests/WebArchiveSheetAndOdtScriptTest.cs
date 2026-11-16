using System.Text;
using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// (1) A web archive's head sheet resolved per element into checked inline styles: type / class / id / descendant / child
/// selectors, :link and :first-child, specificity and order, the element's own style last, and nothing unsafe passing
/// (url(), behavior, expression, position, @import, :hover / :visited). (2) ODT script split over the whole paragraph text
/// (LibreOffice SwScriptInfo::InitScriptInfo, porlay.cxx:1368, 1467): digits after an English word in a right-to-left paragraph
/// stay in the Latin set even when a span starts between them.
/// </summary>
public static class WebArchiveSheetAndOdtScriptTest
{
    public static void Run()
    {
        TestReport.Step("Web archive class styles resolved inline; ODT script split per paragraph");
        WebArchiveSheet();
        WordSavedClassStyles();
        OdtParagraphScript();
    }

    private static void WebArchiveSheet()
    {
        const string boundary = "----=_NextPart_01DA0000.12345678";
        const string css = "<!--\r\n@import url(evil.css);\r\n@font-face {font-family:David; src:url(x.ttf);}\r\n"
            + "body {font-family:\"David\"; margin:5cm;}\r\n"
            + "p.MsoNormal, li.MsoNormal {margin:0cm; font-size:12.0pt; color:red;}\r\n"
            + "p.Big {font-size:20.0pt;}\r\n#special {color:green;}\r\n"
            + "div.Box > p {font-style:italic;}\r\ndiv.Box p.Deep {text-decoration:underline;}\r\n"
            + "a:link, span.MsoHyperlink {color:blue;}\r\na:visited {color:purple;}\r\np:first-child {letter-spacing:2px;}\r\n"
            + ".Huge {margin-left:-5in; padding:1cm 99in; text-indent:-9in; width:9000px; font-size:900pt; letter-spacing:-3cm; margin-top:calc(1px); border-top:1e9px solid;}"+"\r\n"
            + "p:hover {color:orange;}\r\np[title] {color:orange;}\r\np::first-letter {color:orange;}\r\n"
            + ".Evil {background:url(http://example.invalid/y.png); color:navy; position:absolute; behavior:url(x.htc); width:expression(alert(1)); font-family:\"A\\\"B\";}\r\n"
            + "@page WordSection1 {size:595.3pt 841.9pt;}\r\ndiv.WordSection1 {page:WordSection1;}\r\n-->";
        string body = "<html><head><style>" + css + "</style></head><body lang=EN-US>"
            + "<p class=MsoNormal>Plain</p>\r\n"
            + "<p class=\"MsoNormal Big\" style='color:black;position:fixed;top:0'>Both</p>\r\n"
            + "<p class=MsoNormal id=special>Green</p>\r\n"
            + "<div class=Box><p class=MsoNormal>Child</p><p class=Deep>Deep</p></div>\r\n"
            + "<p><a href=\"#x\">Link</a></p>\r\n"
            + "<p class=Huge style='margin-right:-2cm;line-height:400'>Huge</p>\r\n"
            + "<p style=\"color:teal\"class=Big>Tight</p>\r\n"
            + "<p class=Evil style=\"background-image:url(javascript:alert(1))\">Evil</p>\r\n"
            + "</body></html>\r\n";
        string archive = "MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=\"" + boundary + "\"\r\n\r\n"
            + "--" + boundary + "\r\nContent-Location: file:///C:/x/doc.htm\r\nContent-Transfer-Encoding: 8bit\r\nContent-Type: text/html; charset=\"us-ascii\"\r\n\r\n"
            + body + "\r\n--" + boundary + "--\r\n";
        string path = Path.Combine(TestReport.OutputDirectory, "mhtml-class-styles-v1.mht");
        File.WriteAllText(path, archive, Encoding.ASCII);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-mhtml-class-styles") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string all = string.Join("\n", html);
        string Tag(string text)
        {
            int at = all.IndexOf(">" + text + "<", StringComparison.Ordinal);
            if (at < 0) return "";
            int start = all.LastIndexOf('<', at);
            return all.Substring(start, at + 1 - start);
        }
        void Expect(bool condition, string what, string text) => TestReport.Check(condition, $"mhtml sheet: {what} [{Tag(text)}]");
        Expect(Tag("Plain").Contains("style=\"margin:0cm;font-size:12.0pt;color:red;\""), "class rule inline", "Plain");
        Expect(Tag("Both").Contains("font-size:12.0pt;color:red;font-size:20.0pt;color:black;") && !Tag("Both").Contains("position") && !Tag("Both").Contains("top:"),
               "two classes in order, own style last, position dropped", "Both");
        Expect(Tag("Green").EndsWith("color:red;color:green;\">", StringComparison.Ordinal), "id beats class", "Green");
        Expect(Tag("Child").Contains("font-style:italic;") && Tag("Child").Contains("letter-spacing:2px;"), "child combinator and :first-child", "Child");
        Expect(Tag("Deep").Contains("font-style:italic;text-decoration:underline;") && !Tag("Deep").Contains("letter-spacing"), "descendant, not first child", "Deep");
        Expect(Tag("Link").Contains("color:blue;") && !Tag("Link").Contains("purple"), ":link prints, :visited does not", "Link");
        Expect(Tag("Tight").StartsWith("<p ", StringComparison.Ordinal) && Tag("Tight").Contains("class=\"Big\" style=\"font-size:20.0pt;color:teal;\""), "style removed without joining the next attribute", "Tight");
        Expect(Tag("Huge").Contains("style=\"margin-left:-1in;padding:1cm 10in;text-indent:-1.5in;width:1920px;font-size:400pt;letter-spacing:-0.706cm;margin-right:-2cm;line-height:10;\""),
               "lengths from the sheet and the element's own style kept in bounds, calc() dropped", "Huge");
        Expect(Tag("Evil").Contains("style=\"color:navy;\"") && !Tag("Evil").Contains("url") && !Tag("Evil").Contains("javascript"), "only safe declarations", "Evil");
        TestReport.Check(!all.Contains("orange") && !all.Contains("@") && !all.Contains("behavior") && !all.Contains("expression"), "mhtml sheet: no :hover / attribute / pseudo-element rule, nothing unsafe");
        TestReport.Check(document.Stylesheet.Contains(".word-document { font-family:\"David\"; }") && !document.Stylesheet.Contains("MsoNormal")
                         && !document.Stylesheet.Contains("5cm") && !document.Stylesheet.Contains("url("),
                         "mhtml sheet: the page stylesheet carries only the body's inherited declarations, never the sheet");
    }

    /// <summary>
    /// test-data/word-mhtml/class-styles-v1.mht and .pdf, both saved by Word 16 from one document (make-class-styles-fixture.ps1):
    /// the PDF prints "Fancy styled paragraph" in Courier New bold 20 pt red centred, "Heading one text" in Calibri Light 20 pt
    /// #2F5496 and "green" italic 16 pt #008000 (read with PyMuPDF); the styles come only from the archive's class rules.
    /// </summary>
    private static void WordSavedClassStyles()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-mhtml", "class-styles-v1.mht");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-mhtml/class-styles-v1.mht present"); return; }
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        string all = string.Join("\n", document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        string Opening(string text)
        {
            int at = all.IndexOf(text, StringComparison.Ordinal);
            int start = at < 0 ? -1 : Math.Max(all.LastIndexOf("<p ", at, StringComparison.Ordinal), all.LastIndexOf("<h1 ", at, StringComparison.Ordinal));
            return start < 0 ? "" : all.Substring(start, all.IndexOf('>', start) + 1 - start);
        }
        string fancy = Opening("Fancy styled paragraph"), heading = Opening("Heading one text");
        int green = all.IndexOf(">green<", StringComparison.Ordinal);
        string greenRun = green < 0 ? "" : all.Substring(Math.Max(0, green - 200), 201);
        TestReport.Check(fancy.Contains("font-family:&quot;Courier New&quot;") && fancy.Contains("font-size:20.0pt") && fancy.Contains("color:red") && fancy.Contains("font-weight:bold") && fancy.Contains("text-align:center"),
                         $"Word .mht: paragraph style as the PDF prints it (Courier New bold 20 pt red centred) [{fancy}]");
        TestReport.Check(heading.Contains("font-family:&quot;Calibri Light&quot;") && heading.Contains("font-size:20.0pt") && heading.Contains("color:#2F5496"),
                         $"Word .mht: Heading 1 as the PDF prints it (Calibri Light 20 pt #2F5496) [{heading}]");
        TestReport.Check(greenRun.Contains("color:green;font-style:italic;") && greenRun.Contains("font-size:16.0pt"), "Word .mht: character style as the PDF prints it (green italic 16 pt)");
        TestReport.Check(!document.Stylesheet.Contains("MsoNormal") && !document.Stylesheet.Contains("FancyPara"), "Word .mht: the archive sheet is not in the page stylesheet");

        // table-margin-v1.mht / .pdf (make-table-margin-fixture.ps1): Word's PDF draws the table border 5.4 pt left of the body
        // text (84.7 pt vs 90.05 pt), so the cell text lines up with it - the table's margin-left:-5.4pt must survive the bounds
        string tablePath = Path.Combine(Path.GetDirectoryName(path)!, "table-margin-v1.mht");
        using IWordWindowedDocument table = WordWindowedDocumentOpener.Open(tablePath, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        string tableHtml = string.Join("\n", table.RenderBlocks(0, table.BlockCount).Select(block => block.Html));
        int tableAt = tableHtml.IndexOf("<table", StringComparison.Ordinal);
        string tableTag = tableAt < 0 ? "" : tableHtml.Substring(tableAt, tableHtml.IndexOf('>', tableAt) + 1 - tableAt);
        TestReport.Check(tableTag.Contains("margin-left:-5.4pt;"), $"Word .mht: a table's negative margin is kept, as Word prints it [{tableTag}]");
    }

    private static void OdtParagraphScript()
    {
        const string Namespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
                                + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"";
        string hebrew = "" + (char)0x05E9 + (char)0x05DC + (char)0x05D5 + (char)0x05DD;
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles><style:style style:name=\"P1\" style:family=\"paragraph\"><style:paragraph-properties style:writing-mode=\"rl-tb\"/>"
            + "<style:text-properties fo:font-size=\"11pt\" style:font-size-complex=\"16pt\"/></style:style>"
            + "<style:style style:name=\"T1\" style:family=\"text\"><style:text-properties fo:font-style=\"italic\" style:font-style-complex=\"italic\"/></style:style></office:automatic-styles>"
            + "<office:body><office:text><text:p text:style-name=\"P1\">" + hebrew + " English<text:span text:style-name=\"T1\"> 2024</text:span> " + hebrew + "</text:p>"
            + "<text:p text:style-name=\"P1\">" + hebrew + " Eng<text:span text:style-name=\"T1\">lish</text:span>, " + hebrew + "</text:p>"
            + "</office:text></office:body></office:document>";
        string path = Path.Combine(TestReport.OutputDirectory, "odt-paragraph-script-v1.fodt");
        File.WriteAllText(path, document, new UTF8Encoding(false));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-odt-paragraph-script") };
        using IWordWindowedDocument opened = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = opened.RenderBlocks(0, opened.BlockCount).Select(block => block.Html).ToList();
        TestReport.Check(html.Count == 2 && Regex.IsMatch(html[0], "font-size:11pt;?[^\"]*\"> 2024</span>"),
                         $"odt: digits after an English word stay Latin across a span [{Ascii(html.Count > 0 ? html[0] : "")}]");
        TestReport.Check(html.Count == 2 && Regex.IsMatch(html[1], "font-size:11pt;?[^\"]*\">Eng</span>") && Regex.IsMatch(html[1], "font-size:11pt;?[^\"]*\">lish</span>")
                         && !Regex.IsMatch(html[1], "font-size:11pt;?[^\"]*\">lish, "),
                         $"odt: a word split by a span is one Latin run, the comma after it takes the paragraph side [{Ascii(html.Count > 1 ? html[1] : "")}]");
    }

    private static string Ascii(string text)
    {
        var result = new StringBuilder(text.Length);
        foreach (char c in text) result.Append(c < 128 ? c.ToString() : $"U+{(int)c:X4}");
        return result.ToString();
    }
}
