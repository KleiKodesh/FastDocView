using FastDocView.DocBinary;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// .doc right-to-left paragraph properties, checked against [MS-DOC] 2.6.2 (Paragraph Properties) and LibreOffice
/// ww8par6.cxx (Read_Justify / Read_RTLJustify, Read_LR), and against LibreOffice's rendering of real files (the Word
/// comparison, NOTES.md). Grpprls are built by hand: a 2-byte opcode, then the operand.
/// </summary>
public static class DocBinaryBidiTest
{
    public static void Run()
    {
        TestReport.Step("Doc bidi: physical vs logical justification and indents");
        // sprmPJc80 (0x2403) is physical: 2 = physically right = the start side of a bidi paragraph.
        string css = Css(Sprm(0x2441, 1), Sprm(0x2403, 2));
        TestReport.Check(css.Contains("text-align:right;"), $"bidi + sprmPJc80 right renders right ({css})");
        css = Css(Sprm(0x2441, 1), Sprm(0x2403, 0));
        TestReport.Check(css.Contains("text-align:left;"), $"bidi + sprmPJc80 left renders left ({css})");
        // sprmPJc (0x2461) is logical: 2 = logical right = end; in a bidi paragraph that is physically left.
        css = Css(Sprm(0x2441, 1), Sprm(0x2461, 2));
        TestReport.Check(css.Contains("text-align:left;"), $"bidi + sprmPJc end renders left ({css})");
        // both written (Word 2000+): the later one wins, and both say the same
        css = Css(Sprm(0x2441, 1), Sprm(0x2403, 0), Sprm(0x2461, 2));
        TestReport.Check(css.Contains("text-align:left;"), $"bidi + sprmPJc80 then sprmPJc renders left ({css})");
        css = Css(Sprm(0x2403, 2));
        TestReport.Check(css.Contains("text-align:right;"), $"ltr + sprmPJc80 right renders right ({css})");
        // sprmPDxaLeft80 (0x840F) is the physical left indent: the end side of a bidi paragraph.
        css = Css(Sprm(0x2441, 1), Sprm16(0x840F, 720), Sprm16(0x840E, 360));
        TestReport.Check(css.Contains("margin-inline-end:36pt;") && css.Contains("margin-inline-start:18pt;"), $"bidi + physical indents swap to end / start ({css})");
        css = Css(Sprm(0x2441, 1), Sprm16(0x845E, 720));
        TestReport.Check(css.Contains("margin-inline-start:36pt;"), $"bidi + logical left indent stays start ({css})");

    }

    private static string Css(params byte[][] sprms)
    {
        byte[] grpprl = sprms.SelectMany(sprm => sprm).ToArray();
        var formatting = new WordParagraphFormatting();
        DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, grpprl, 0, grpprl.Length);
        return formatting.ToCss();
    }

    private static byte[] Sprm(ushort opcode, byte operand) => [(byte)opcode, (byte)(opcode >> 8), operand];
    private static byte[] Sprm16(ushort opcode, short operand) => [(byte)opcode, (byte)(opcode >> 8), (byte)operand, (byte)(operand >> 8)];
}
