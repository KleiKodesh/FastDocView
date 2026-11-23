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
            // Latin runs, which the RTF reader already follows), so those runs are not checked (docs/UNSUPPORTED-FEATURES.md)
            if (format == ".rtf" && tag is "P4" or "P5" or "P6") continue;
            string got = FaceText(html, face);
            TestReport.Check(got == expected, $"{format} {tag}: the characters in {face} are Word's (got [{Ascii(got)}], want [{Ascii(expected)}])");
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
