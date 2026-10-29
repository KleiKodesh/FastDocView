using System.Text;
using FastDocView.DocBinary;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// List label layout (found by the LibreOffice PDF comparison, NOTES.md): the label starts at start indent + first-line
/// indent, a tab suffix reaches the next stop (custom stop, the hanging indent, else a default stop); .doc list paragraphs
/// take their LVL's grpprlPapx indents (LibreOffice ww8par3.cxx), its ixchFollow and jc.
/// </summary>
public static class DocBinaryListLayoutTest
{
    public static void Run()
    {
        TestReport.Step("List label layout: tab gap, alignment, .doc LVL indents");
        static WordParagraphFormatting Paragraph(int left, int firstLine, params int[] stops)
            => new() { LeftIndentTwips = left, FirstLineIndentTwips = firstLine, TabStops = stops.Select(s => new WordTabStop(s, "left", "none")).ToArray() };
        TestReport.Check(WordListLabelHtml.TabGapTwips(Paragraph(720, -360), 720) == 360, "hanging indent: the tab goes to the start indent");
        TestReport.Check(WordListLabelHtml.TabGapTwips(Paragraph(0, 0), 720) == 720, "no indent: the next default stop");
        TestReport.Check(WordListLabelHtml.TabGapTwips(Paragraph(0, 200), 567) == 367, "first-line indent: the next default stop after the label");
        TestReport.Check(WordListLabelHtml.TabGapTwips(Paragraph(720, -360, 500), 720) == 140, "a custom stop before the indent wins");
        TestReport.Check(WordListLabelHtml.TabGapTwips(Paragraph(720, -360, 1000), 720) == 360, "a custom stop after the indent loses to the indent");
        TestReport.Check(WordListLabelHtml.TabGapTwips(new WordParagraphFormatting { IsRightToLeft = true, IsLeftIndentPhysical = true, IsRightIndentPhysical = true,
                                                                                    LeftIndentTwips = 100, RightIndentTwips = 720, FirstLineIndentTwips = -360 }, 720) == 360,
                         "bidi + physical indents: the right indent is the start side");

        var html = new StringBuilder();
        WordListLabelHtml.Append(html, "1.", null, WordListLabelSuffix.Tab, Paragraph(720, -360), 720, WordListLabelAlignment.End);
        TestReport.Check(html.ToString() == "<span class=\"list-marker\" style=\"min-width:18pt;\"><span style=\"display:inline-flex;width:0;white-space:nowrap;justify-content:flex-end\">1.</span></span>",
                         "end-aligned label: zero-width flex box reaching back, tab gap as min-width (" + html + ")");
        html.Clear();
        WordListLabelHtml.Append(html, "a)", null, WordListLabelSuffix.Space, Paragraph(720, -360), 720);
        TestReport.Check(html.ToString() == "<span class=\"list-marker\">a)</span> ", "space suffix: label then a space (" + html + ")");

        // sprmPChgTabsPapx: delete none, add one stop at 1000 twips, centre aligned
        var formatting = new WordParagraphFormatting();
        byte[] grpprl = [0x0D, 0xC6, 5, 0, 1, 0xE8, 0x03, 0x01];
        DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, grpprl, 0, grpprl.Length);
        TestReport.Check(formatting.TabStops is [{ PositionTwips: 1000, Alignment: "center" }], "sprmPChgTabsPapx fills the tab stops");
        // sprmPChgTabs: delete 1005 within close tolerance 10 (rgdxaClose follows rgdxaDel), then add 500
        formatting.TabStops = [new WordTabStop(1005, "left", "none")];
        byte[] withClose = [0x15, 0xC6, 9, 1, 0xEC, 0x03, 0x0A, 0x00, 1, 0xF4, 0x01, 0x00];
        DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, withClose, 0, withClose.Length);
        TestReport.Check(formatting.TabStops is [{ PositionTwips: 500 }], "sprmPChgTabs skips rgdxaClose and deletes within the tolerance");

        // Apache POI Lists.doc: level 0 LVL grpprlPapx = left 720, first line -360, ixchFollow tab, no direct indents
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "apache-poi", "Lists.doc");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/apache-poi/Lists.doc present"); return; }
        using var document = DocBinaryWindowedDocument.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        int block = Enumerable.Range(0, document.BlockIndex.BlockCount).First(b => document.BlockIndex.BlockListLabelIds[b] >= 0);
        string rendered = document.RenderBlocks(block, 1)[0].Html;
        TestReport.Check(rendered.Contains("margin-inline-start:36pt;") && rendered.Contains("text-indent:-18pt;") && rendered.Contains("min-width:18pt;"),
                         "Lists.doc: the LVL indents reach the paragraph, the label reaches the hanging indent");
    }
}
