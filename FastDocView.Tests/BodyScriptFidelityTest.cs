using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Latin text in a right-to-left paragraph: test-data/word-script/body-script-v1.* (Word 16, make-body-script-fixture.ps1; .odt is
/// LibreOffice's conversion). Word 16's PDF export (body-script-v1.pdf) draws the Hebrew words and the digits in the complex-script set
/// (Arial 16) and "English" in the Latin set (Courier New 11): the set follows each run, not the paragraph's direction.
/// </summary>
public static class BodyScriptFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Body text: Latin or complex-script set per run in right-to-left paragraphs (.docx / .doc / .rtf / .odt)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-script");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "body-script-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-script/body-script-v1{extension} present"); continue; }
            Check(path, extension);
        }
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-body-script-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string mixed = blocks.FirstOrDefault(html => html.Contains("English")) ?? "";
        TestReport.Check(Regex.IsMatch(mixed, "font-size:11pt;font-family:'Courier New';\">English</span>"),
                         $"{format}: English in a right-to-left paragraph draws in the Latin set, Courier New 11 pt (got [{Ascii(mixed)}])");
        int english = mixed.IndexOf("English</span>", StringComparison.Ordinal);
        int digits = mixed.IndexOf("123", StringComparison.Ordinal);
        TestReport.Check(english > 0 && digits > english && !mixed.Substring(english, digits - english).Contains("<span style=\"font-size:11pt")
                         && Regex.IsMatch(mixed, "<span style=\"font-size:16pt;[^\"]*\">" + (char)0x05E9),
                         $"{format}: the Hebrew words and the digits draw in the complex-script set, 16 pt (got [{Ascii(mixed)}])");
        // a symbol-code bullet whose Wingdings font is set only for complex script: the label draws in the Latin set (its code is
        // strong left-to-right), so Word looks the code up in the Latin font (Courier New) - not Wingdings (Word 16 PDF: no Wingdings glyph)
        string bullet = blocks.FirstOrDefault(html => html.Contains("wingdings")) ?? "";
        int marker = bullet.IndexOf("<span class=\"list-marker\"", StringComparison.Ordinal);
        string label = marker < 0 ? "" : bullet.Substring(marker, Math.Max(0, bullet.IndexOf("</span></span>", marker, StringComparison.Ordinal) - marker));
        TestReport.Check(label.Contains((char)0x2022) && !label.Contains("Wingdings") && !label.Contains((char)0xF0A7),
                         $"{format}: a bullet with a complex-script-only Wingdings font is not drawn in Wingdings (got [{Ascii(label)}])");
    }

    /// <summary>Dev aid (--body-script-sample FILE FROM COUNT): the rendered blocks, ASCII only.</summary>
    public static void DumpSample(string path, int from, int count)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-body-script-sample") });
        foreach (var block in document.RenderBlocks(from, Math.Min(count, document.BlockCount - from))) Console.WriteLine(Ascii(block.Html));
    }

    internal static string Ascii(string text)
    {
        var s = new StringBuilder(text.Length);
        foreach (char c in text) s.Append(c < 128 ? c.ToString() : "U+" + ((int)c).ToString("X4"));
        return s.ToString();
    }
}
