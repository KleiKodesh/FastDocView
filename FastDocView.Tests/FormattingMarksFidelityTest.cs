using System.Text.RegularExpressions;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Emphasis marks, underline colour, character and paragraph borders, patterned shading in .doc / .rtf / .odt:
/// test-data/word-formatting/marks-borders-v1.doc / .rtf saved by Word 16 (make-marks-borders-fixture.ps1), plus hand-written
/// RTF (what Word writes in other ways: \box with modifiers, \shading + \cfpat, resets, borders that must not leak) and a
/// hand-written .fodt for style:text-emphasize.
/// </summary>
public static class FormattingMarksFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Emphasis marks, underline colour, borders, shading (.doc / .rtf / .odt)");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-formatting"));
        WordSavedEmphasis(Path.Combine(folder, "marks-borders-v1.doc"), ".doc");
        WordSavedEmphasis(Path.Combine(folder, "marks-borders-v1.rtf"), ".rtf");
        WordSavedRtfBordersAndShading(Path.Combine(folder, "marks-borders-v1.rtf"));
        HandWrittenRtf();
        OdtEmphasis();
    }

    private static List<string> Render(string path, string cacheName)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-" + cacheName) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    private static string Block(List<string> html, string text) => html.FirstOrDefault(block => WordFieldsFidelityTest.Strip(block).Contains(text)) ?? "";

    private static void Expect(List<string> html, string text, string css, string what)
    {
        string block = Block(html, text);
        TestReport.Check(block.Contains(css), $"{what}: [{css}] in [{(block.Length > 400 ? block.Substring(0, 400) : block)}]");
    }

    /// <summary>Word 16: Font.EmphasisMark 1..4 = over solid circle, over comma, over white circle, under solid circle (sprmCKcd / \acc*).</summary>
    private static void WordSavedEmphasis(string path, string format)
    {
        if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-formatting/{Path.GetFileName(path)} present"); return; }
        List<string> html = Render(path, "marks-" + format.TrimStart('.'));
        Expect(html, "dot marks", "text-emphasis:filled dot;", format + " dot above");
        Expect(html, "comma marks", "text-emphasis:filled sesame;", format + " comma");
        Expect(html, "circle marks", "text-emphasis:open circle;", format + " circle");
        Expect(html, "under dot marks", "text-emphasis:filled dot;text-emphasis-position:under right;", format + " dot below");
        TestReport.Check(!Block(html, "bottom rule").Contains("text-emphasis"), format + " no mark where none is set");
    }

    /// <summary>Word 16's RTF: \ul\ulc6 (red), \chbrdr\brdrdb\brdrw30\brdrcf2, a \brdrt/l/b/r box \brdrw20 \brdrcf11 with \cbpat7,
    /// \brdrb\brdrdashsm\brdrw10, \chshdng2500\chcfpat1\chcbpat8 (25 % black over white = BFBFBF).</summary>
    private static void WordSavedRtfBordersAndShading(string path)
    {
        if (!File.Exists(path)) return;
        List<string> html = Render(path, "marks-rtf-2");
        Expect(html, "red underline", "text-decoration-color:light-dark(#FF0000", ".rtf \\ulc");
        Expect(html, "bordered word", "border:1.5pt double #0000FF;", ".rtf \\chbrdr double 1.5 pt blue");
        string boxed = Block(html, "boxed paragraph");
        TestReport.Check(boxed.Contains("border-top:1pt solid #008000;") && boxed.Contains("border-bottom:1pt solid #008000;")
                         && boxed.Contains("border-inline-start:1pt solid #008000;") && boxed.Contains("border-inline-end:1pt solid #008000;"),
                         ".rtf paragraph box: four 1 pt green sides [" + boxed + "]");
        TestReport.Check(boxed.Contains("background-color:light-dark(#FFFF00"), ".rtf \\cbpat: the paragraph's yellow fill");
        string rule = Block(html, "bottom rule");
        TestReport.Check(rule.Contains("border-bottom:0.5pt dashed #000;") && !rule.Contains("border-top"), ".rtf bottom border only, dashed [" + rule + "]");
        string shade = Block(html, "patterned shade");
        TestReport.Check(shade.Contains("background-color:#BFBFBF") || shade.Contains("background:#BFBFBF") || shade.Contains("#BFBFBF"),
                         ".rtf \\chshdng2500 black over white = BFBFBF, as shading [" + shade + "]");
    }

    private static void HandWrittenRtf()
    {
        string rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}{\\colortbl;\\red255\\green0\\blue0;\\red0\\green0\\blue255;\\red255\\green255\\blue255;}\n"
                   + "\\pard\\box\\brdrdb\\brdrw40\\brdrcf2 boxbox\\par\n"
                   + "\\pard plainpara\\par\n"
                   + "\\pard\\rtlpar\\brdrl\\brdrs\\brdrw20\\brdrcf1 rtlleft\\par\n"
                   + "\\pard\\shading5000\\cfpat2\\cbpat3 halfshade\\par\n"
                   + "\\pard\\trowd\\clbrdrt\\brdrs\\brdrw60\\cellx2000\\intbl cellpara\\cell\\row\n"
                   + "\\pard{\\chbrdr\\brdrs\\brdrcf1 inborder}{\\plain afterplain}{\\accdot\\accdot0 dotoff}{\\chshdng0\\chcbpat1 flatshade}{\\chbgfdiag\\chcfpat2 diag}\\par\n"
                   + "\\pard\\brdrt\\brdrnone nonepara\\par}";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-marks-borders.rtf");
        File.WriteAllText(path, rtf, Encoding.ASCII);
        List<string> html = Render(path, "marks-rtf-hand");
        string box = Block(html, "boxbox");
        TestReport.Check(new[] { "border-top", "border-bottom", "border-inline-start", "border-inline-end" }.All(side => box.Contains(side + ":2pt double #0000FF;")),
                         "\\box: the modifiers after it reach all four sides [" + box + "]");
        TestReport.Check(!Block(html, "plainpara").Contains("border"), "\\pard: the borders end");
        string rtl = Block(html, "rtlleft");
        TestReport.Check(rtl.Contains("border-inline-start:1pt solid #FF0000;") && !rtl.Contains("border-inline-end"), "\\brdrl in a right-to-left paragraph: the start side, as .docx w:left and .doc brcLeft [" + rtl + "]");
        TestReport.Check(Block(html, "halfshade").Contains("background-color:light-dark(#7F7FFF"), "\\shading5000 \\cfpat blue over \\cbpat white: 7F7FFF");
        TestReport.Check(Regex.Match(Block(html, "cellpara"), "<p[^>]*>").Value is { Length: > 0 } cellParagraph && !cellParagraph.Contains("border"), "\\clbrdrt modifiers do not reach the paragraph");
        string runs = Block(html, "inborder");
        TestReport.Check(runs.Contains("border:0.5pt solid #FF0000;"), "\\chbrdr without \\brdrw: 0.5 pt, red [" + runs + "]");
        TestReport.Check(runs.Contains("inborder</span>afterplain"), "\\plain ends the character border");
        TestReport.Check(!runs.Contains("text-emphasis"), "\\accdot0: no mark");
        TestReport.Check(runs.Contains("<span style=\"background-color:light-dark(#FF0000,") && runs.Contains(");\">flatshade"), "\\chcbpat alone: the fill, drawn as shading");
        TestReport.Check(runs.Contains("#AAAAFF"), "\\chbg* pattern: a third of the pattern colour (blue) over white = AAAAFF [" + runs + "]");
        TestReport.Check(!Block(html, "nonepara").Contains("border-top:"), "\\brdrnone: no line");
    }

    /// <summary>style:text-emphasize (ODF 1.3 part 3 20.381), read as LibreOffice's txtprhdl.cxx importXML.</summary>
    private static void OdtEmphasis()
    {
        const string Namespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
                                + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"";
        string Style(string name, string value) => $"<style:style style:name=\"{name}\" style:family=\"text\"><style:text-properties style:text-emphasize=\"{value}\"/></style:style>";
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document " + Namespaces + " office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
                        + "<office:automatic-styles>" + Style("E1", "dot above") + Style("E2", "accent above") + Style("E3", "circle below") + Style("E4", "disc above") + Style("E5", "none")
                        + "</office:automatic-styles><office:body><office:text>"
                        + "<text:p><text:span text:style-name=\"E1\">odtdot</text:span></text:p><text:p><text:span text:style-name=\"E2\">odtaccent</text:span></text:p>"
                        + "<text:p><text:span text:style-name=\"E3\">odtcircle</text:span></text:p><text:p><text:span text:style-name=\"E4\">odtdisc</text:span></text:p>"
                        + "<text:p><text:span text:style-name=\"E5\">odtnone</text:span></text:p></office:text></office:body></office:document>";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-emphasis.fodt");
        File.WriteAllText(path, document, new UTF8Encoding(false));
        List<string> html = Render(path, "marks-odt");
        Expect(html, "odtdot", "text-emphasis:filled dot;", ".odt dot above");
        Expect(html, "odtaccent", "text-emphasis:filled sesame;", ".odt accent");
        Expect(html, "odtcircle", "text-emphasis:open circle;text-emphasis-position:under right;", ".odt circle below");
        Expect(html, "odtdisc", "text-emphasis:filled dot;", ".odt disc");
        TestReport.Check(!Block(html, "odtnone").Contains("text-emphasis"), ".odt none: no mark");
    }
}
