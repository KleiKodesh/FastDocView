using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// Line spacing rules (LibreOffice PDF comparison, NOTES.md): auto spacing counts 240ths of the font's single line height
/// (1.15 em for Times New Roman / Arial), exact is the stated height, at least never drops below a single line.
/// </summary>
public static class LineSpacingTest
{
    public static void Run()
    {
        TestReport.Step("Line spacing: auto multiple of the single line, exact, at least");
        static string Css(int value, bool auto, bool exact) => new WordParagraphFormatting { LineSpacingTwipsOrAuto = value, IsLineSpacingAuto = auto, IsLineSpacingExact = exact }.ToCss();
        TestReport.Check(Css(240, true, false).Contains("line-height:1.15;"), "auto 240: one single line (1.15 em)");
        TestReport.Check(Css(480, true, false).Contains("line-height:2.3;"), "auto 480: two single lines");
        TestReport.Check(Css(360, false, true).Contains("line-height:18pt;"), "exact: the stated height");
        TestReport.Check(Css(200, false, false).Contains("line-height:max(1.15em,10pt);"), "at least: never below a single line");
        var formatting = new WordParagraphFormatting();
        byte[] grpprl = [0x12, 0x64, 0x98, 0xFE, 0, 0];      // sprmPDyaLine: dyaLine -360, fMultLinespace 0
        FastDocView.DocBinary.DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, grpprl, 0, grpprl.Length);
        TestReport.Check(formatting.IsLineSpacingExact && formatting.LineSpacingTwipsOrAuto == 360, ".doc negative dyaLine is exact");
    }
}
