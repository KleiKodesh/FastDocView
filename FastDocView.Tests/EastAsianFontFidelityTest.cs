using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// The East Asian font slot of body text: test-data/word-east-asian/east-asian-font-v1.* (make-east-asian-font-docx.py, then Word 16
/// saves .docx / .doc / .rtf / .odt and the reference PDF: make-east-asian-font-fixture.ps1; pdf-fonts.py lists each character's
/// face). Word draws, per paragraph (P1-P8), these characters in the East Asian face and every other one in the Latin or
/// complex-script face; the check compares the characters inside East Asian font spans with Word's.
/// </summary>
public static class EastAsianFontFidelityTest
{
    private const string Mincho = "MS Mincho";

    // Word 16's PDF (pdf-fonts.py): the characters each paragraph draws in the East Asian face (Hangul U+AC00 included: Word asks
    // MS Mincho, which has no Hangul, and falls back to Malgun Gothic, as the browser does)
    private static readonly (string Tag, string Face, string Expected)[] WordFaces =
    [
        ("P1", Mincho, "日本あ、가Ａ"),                                               // no hint: punctuation stays Arial
        ("P2", Mincho, "“—…·日本あ、가Ａ"),                         // hint: shared punctuation too, not e-acute / U+00AB
        ("P3", "SimSun", "日本あ"),                                                                // eastAsiaTheme minorEastAsia, themeFontLang zh-CN: Hans
        ("P4", Mincho, "日本あ"),                                                                  // right-to-left run: CJK still East Asian
        ("P5", Mincho, "日本あ"),                                                                  // w:cs run
        ("P6", Mincho, "日本あ"),                                                                  // hint in a right-to-left run: quotes stay
        ("P7", Mincho, "¡¤§¨ª­¯°±´¶·º¼¿×÷ˇΑαАё‐‘†‰※€℃№Ⅰ←∀⌒①─■★♀　〈！￥"),
        ("P8", Mincho, "　〈！￥"),                                                            // the same without the hint
    ];

    public static void Run()
    {
        TestReport.Step("Body text: East Asian font slot per character (.docx / .doc / .rtf / .odt, Word 16 PDF; Word's .rtf splits the runs by script itself)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-east-asian");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "east-asian-font-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-east-asian/east-asian-font-v1{extension} present"); continue; }
            Check(path, extension);
        }
        CheckOdtVariants(Path.GetFullPath(Path.Combine(folder, "..", "odt-east-asian", "odt-east-asian-variants-v1.odt")));
        CheckRtfRuns(Path.GetFullPath(Path.Combine(folder, "..", "rtf-east-asian", "rtf-east-asian-runs-v1.rtf")));
        foreach (string extension in new[] { ".docx", ".doc" })
            CheckEastAsianLayout(Path.GetFullPath(Path.Combine(folder, "..", "word-east-asian-layout", "east-asian-layout-v1" + extension)));
    }

    /// <summary>
    /// w:eastAsianLayout (two lines in one with each bracket, horizontal in vertical, vertCompress) and w:fitText, .docx and .doc
    /// (sprmCFELayout): Word 16 with no East Asian editing language prints each run as plain text in one line at its own size,
    /// no brackets, no rotation, no spacing (test-data/word-east-asian-layout/east-asian-layout-v1.pdf: the glyph origins are
    /// those of unformatted text), so the HTML adds nothing to these runs and keeps their text.
    /// </summary>
    private static void CheckEastAsianLayout(string path)
    {
        string format = Path.GetExtension(path);
        if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-east-asian-layout/east-asian-layout-v1{format} present"); return; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-east-asian-layout-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        const string Han = "\u65e5\u672c\u8a9e\u6587\u5b57\u5217";
        (string Tag, string Text)[] runs =
        [
            ("L1", "ABCDEF"), ("L2", "ABCDEF"), ("L3", "ABCDE"), ("L4", Han), ("L5", Han),
            ("L6", "12"), ("L7", "ABCD"), ("L8", "ABC"), ("L9", "ABCDEFGHIJ"), ("L10", Han.Substring(0, 3)),
        ];
        string text = WordDocumentText.GetText(document);
        foreach ((string tag, string runText) in runs)
        {
            string html = blocks.FirstOrDefault(block => Regex.IsMatch(block, ">" + tag + " x")) ?? "";
            string shown = WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", ""));
            bool isPlain = !Regex.IsMatch(html, "letter-spacing|transform|writing-mode|text-combine|scale");
            TestReport.Check(shown.Contains(tag + " x" + runText + "y") && isPlain, $"{format} {tag}: drawn as plain text in one line, as Word 16 prints it");
            TestReport.Check(text.Contains(tag + " x" + runText + "y"), $"{format} {tag}: GetText keeps the text");
        }
    }

    /// <summary>
    /// .rtf: East Asian characters in right-to-left / complex-script runs take the run's \dbch\af face (else the document's
    /// \stshfdbch) at the Latin slot's size and weight (Word 16's PDF of the fixture, see make-rtf-east-asian-runs.py).
    /// </summary>
    private static void CheckRtfRuns(string path)
    {
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/rtf-east-asian/rtf-east-asian-runs-v1.rtf present"); return; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-east-asian-rtf-runs") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        const string Han = "\u65e5\u672c";
        (string Tag, string Css)[] expected =
        [
            ("A1", "font-family:'SimSun'"),                                            // no \dbch\af: \stshfdbch
            ("A3", "font-family:'SimSun';font-size:1.2em"),                            // 12pt beside the 10pt complex-script text
            ("A5", "font-family:'MS Mincho';font-size:0.5em"),                         // \dbch\af20, 10pt beside 20pt
            ("A6", "font-family:'Microsoft YaHei';font-size:0.5em;font-weight:normal"), // the Latin \b0, not the complex \ab
            ("A8", "font-family:'MS Mincho'"),                                         // \f after \dbch\af keeps it
            ("A10", "font-family:'Microsoft YaHei'"),                                  // \dbch\f while none is set
        ];
        foreach ((string tag, string css) in expected)
        {
            string html = blocks.FirstOrDefault(block => block.Contains(">" + tag + " ")) ?? "";
            TestReport.Check(html.Contains("<span style=\"" + css + "\">" + Han + "</span>"), $".rtf {tag}: the Han characters draw as Word draws them ({css})");
        }
        // A9: \f after \dbch in a left-to-right run names both faces, so the run's own face draws them (no switch)
        string a9 = blocks.FirstOrDefault(block => block.Contains(">A9 ")) ?? "";
        TestReport.Check(a9.Contains("font-family:'Microsoft YaHei'") && !a9.Contains("\">" + Han + "</span>"), ".rtf A9: the Han characters draw in the run's Microsoft YaHei");
        string text = WordDocumentText.GetText(document);
        TestReport.Check(text.Contains("A5 \u05e9\u05dc\u05d5\u05dd " + Han + " Ab"), ".rtf runs: GetText keeps the text as written");
    }

    /// <summary>
    /// .odt style:font-size-asian / -weight-asian / -style-asian: the Han characters take the Asian set, the rest the Western /
    /// complex one (LibreOffice 25's PDF of the fixture, see make-odt-east-asian-variants.py).
    /// </summary>
    private static void CheckOdtVariants(string path)
    {
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/odt-east-asian/odt-east-asian-variants-v1.odt present"); return; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-east-asian-odt-variants") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        (string Tag, string Css)[] expected =
        [
            ("P1", "font-family:'MS Mincho';font-size:1.667em;font-weight:bold;font-style:italic"),   // 20pt bold italic over 12pt
            ("P2", "font-family:'MS Mincho';font-size:1.5em;font-weight:normal"),                     // 18pt, not bold, in a bold run
            ("P3", "font-family:'MS Mincho';font-size:0.75em"),                                       // fo:font-size 16pt leaves it 12pt
            ("P4", "font-family:'MS Mincho';font-size:1.667em;font-weight:bold;font-style:italic"),   // right-to-left paragraph
        ];
        foreach ((string tag, string css) in expected)
        {
            string html = blocks.FirstOrDefault(block => block.Contains(">" + tag + " ")) ?? "";
            bool found = html.Contains("<span style=\"" + css + "\">漢字</span>");
            TestReport.Check(found, $".odt {tag}: the Han characters take the Asian size / bold / italic ({css})");
            TestReport.Check(!Regex.IsMatch(html, ">[^<]*Ab[^<]*漢"), $".odt {tag}: the Latin text stays outside the Asian span");
        }
        string text = WordDocumentText.GetText(document);
        TestReport.Check(text.Contains("P1 Ab 漢字 Z"), ".odt variants: GetText keeps the text as written");
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-east-asian-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        foreach ((string tag, string face, string wordExpected) in WordFaces)
        {
            string expected = wordExpected;
            // .odt: Word's ODT export names the themed run's resolved face and drops w:hint (no ODF attribute carries it), so the
            // hinted paragraphs draw as the unhinted ones (LibreOffice: weak characters follow the script before them)
            string html = blocks.FirstOrDefault(block => Regex.IsMatch(block, ">" + tag + "(</span>)? ")) ?? "";
            if (format == ".odt") expected = tag switch { "P2" => WordFaces[0].Expected, "P7" => WordFaces[7].Expected, _ => expected };
            // .rtf: Word's RTF writes no East Asian font in a right-to-left / complex-script run's group (only \dbch\afN in
            // Latin runs), so Word reopening it draws those characters in the document's \stshfdbch face, SimSun
            // (test-data/rtf-east-asian/east-asian-font-v1-rtf.pdf, Word 16's PDF of the .rtf)
            string drawnFace = format == ".rtf" && tag is "P4" or "P5" or "P6" ? "SimSun" : face;
            string got = FaceText(html, drawnFace);
            TestReport.Check(got == expected, $"{format} {tag}: the characters in {drawnFace} are Word's (got [{Ascii(got)}], want [{Ascii(expected)}])");
        }
        // the text is unchanged: GetText / search see the characters, not the spans
        string text = WordDocumentText.GetText(document);
        TestReport.Check(text.Contains("P2 Ab1 é«“—…· 日本あ、가Ａ Z"), $"{format}: GetText keeps the text as written");
    }

    /// <summary>The characters inside spans that switch to <paramref name="face"/>, in order.</summary>
    private static string FaceText(string html, string face)
    {
        var text = new StringBuilder();
        foreach (Match match in Regex.Matches(html, "<span style=\"[^\"]*font-family:'" + Regex.Escape(face) + "';?\">([^<]*)</span>"))
            text.Append(WebUtility.HtmlDecode(match.Groups[1].Value));
        return text.ToString();
    }

    private static string Ascii(string text) => BodyScriptFidelityTest.Ascii(text);
}
