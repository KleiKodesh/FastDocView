using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// RTF drawing objects ({\shp} text boxes, shapes, {\shpgrp} groups): a hand-written .rtf with each object type, and the .odt
/// shapes fixture (OdtShapesFidelityTest) saved as .rtf by LibreOffice (headless, into TestOutput, never committed) - placement,
/// look, group offsets, text-box paragraphs that never split the block, and text-box text found by the search exactly once.
/// </summary>
public static class RtfShapesFidelityTest
{
    private static string Shape(string words, string properties, string text = "")
        => @"{\shp{\*\shpinst" + words + properties + (text.Length > 0 ? @"{\shptxt \pard\plain " + text + "}" : "") + @"}{\shprslt {\pard RESULTCOPY\par}}}";

    private static string Property(string name, string value) => @"{\sp{\sn " + name + @"}{\sv " + value + "}}";

    /// <summary>Blocks in order; the test checks each by its index.</summary>
    private static readonly string[] Body =
    [
        // 0 text box: 2 x 1 in, square wrap, text right of it, insets 0, green line 2 pt, yellow fill, two paragraphs
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright2880\shpbottom1440\shpwr2\shpwrk2", Property("shapeType", "202") + Property("dxTextLeft", "0") + Property("dyTextTop", "0")
            + Property("dxTextRight", "0") + Property("dyTextBottom", "0") + Property("lineColor", "32768") + Property("lineWidth", "25400") + Property("fillColor", "65535"),
            @"boxtext one\par second line\par") + @"outside text\par",
        // 1 ellipse, red, no line
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "3") + Property("fillColor", "255") + Property("fLine", "0")) + @"\par",
        // 2 arrow, blue line, not filled, with text
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright2880\shpbottom720\shpwr1", Property("shapeType", "13") + Property("fFilled", "0") + Property("lineColor", "16711680"), "arrowtext") + @"\par",
        // 3 group: two text boxes side by side in a 0..200 x 0..100 coordinate space over a 4 x 1 in box
        @"\pard {\shpgrp{\*\shpinst\shpleft0\shptop0\shpright5760\shpbottom1440\shpwr1" + Property("groupLeft", "0") + Property("groupTop", "0") + Property("groupRight", "200") + Property("groupBottom", "100")
            + @"{\shp{\*\shpinst" + Property("relLeft", "0") + Property("relTop", "0") + Property("relRight", "100") + Property("relBottom", "100") + Property("shapeType", "202") + @"{\shptxt groupone}}}"
            + @"{\shp{\*\shpinst" + Property("relLeft", "100") + Property("relTop", "50") + Property("relRight", "200") + Property("relBottom", "100") + Property("shapeType", "1") + Property("rotation", "5898240") + @"{\shptxt grouptwo}}}"
            + @"}{\shprslt {\pard GROUPCOPY\par}}}\par",
        // 4 line with an end arrow
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright2880\shpbottom720\shpwr3", Property("shapeType", "20") + Property("lineEndArrowhead", "1")) + @"\par",
        // 5 isosceles triangle, round rectangle, diamond, unknown type (a box)
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440", Property("shapeType", "5")) + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720", Property("shapeType", "2"))
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440", Property("shapeType", "4")) + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440", Property("shapeType", "96")) + @"last\par",
        // 6 a group turned 90 degrees: its left-half child turns about the group's centre
        @"\pard {\shpgrp{\*\shpinst\shpleft0\shptop0\shpright2880\shpbottom1440" + Property("rotation", "5898240") + Property("groupLeft", "0") + Property("groupTop", "0") + Property("groupRight", "200") + Property("groupBottom", "100")
            + @"{\shp{\*\shpinst" + Property("relLeft", "0") + Property("relTop", "0") + Property("relRight", "100") + Property("relBottom", "100") + Property("shapeType", "1") + @"}}}}\par",
        // 7 system colours (MSO_CLR 0x10..): button face darkened by 0x80 as the fill; a red fill and "use the fill colour" (0xF0) as the line
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright720\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillColor", "276824320")) + Shape(@"\shpleft0\shptop0\shpright720\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillColor", "255") + Property("lineColor", "268435696")) + @"\par",
        // 8 fills: linear shade (fillType 7, angle 0, focus 0), axial (focus 50), shade shape (6), shade centre (5, focus 0.5 / 0.5),
        // an 8 x 8 1-bit pattern (1) and a picture fill (3)
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillType", "7") + Property("fillColor", "255") + Property("fillBackColor", "16711680"))
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillType", "7") + Property("fillFocus", "50") + Property("fillColor", "255") + Property("fillBackColor", "16711680"))
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillType", "6") + Property("fillColor", "255") + Property("fillBackColor", "16711680"))
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillType", "5") + Property("fillToRight", "32768") + Property("fillToBottom", "32768") + Property("fillColor", "255") + Property("fillBackColor", "16711680"))
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "1") + Property("fillType", "1") + Property("fillColor", "255") + Property("fillBackColor", "16711680") + @"{\sp{\sn fillBlip}{\sv {\pict\dibitmap0 2800000008000000080000000100010000000000200000000000000000000000020000000000000000000000ffffff00aa00000055000000aa00000055000000aa00000055000000aa00000055000000}}}")
            + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom720\shpwr1", Property("shapeType", "3") + Property("fillType", "3") + @"{\sp{\sn fillBlip}{\sv {\pict\pngblip 89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d49444154789c636060f80f0000010100005c1b0d8b0000000049454e44ae426082}}}") + @"\par",
        // 9 the shape's own geometry (pVerticies / pSegmentInfo: move, 2 lines, close, end) and WordArt (gtextUNICODE in Arial 36 pt)
        @"\pard " + Shape(@"\shpleft0\shptop0\shpright1440\shpbottom1440\shpwr1", Property("shapeType", "0") + Property("pVerticies", "8;3;(0,0);(21600,0);(10800,21600)") + Property("pSegmentInfo", "2;4;16384;2;24577;32768"))
            + Shape(@"\shpleft0\shptop0\shpright2880\shpbottom720\shpwr1", Property("shapeType", "136") + Property("fGtext", "1") + Property("fillColor", "255") + Property("gtextUNICODE", "WordArtText") + Property("gtextFont", "Arial") + Property("gtextSize", "2359296")) + @"\par",
        // 10 a text box holding a two-cell table, inside a one-cell outer table row (the box's \trowd must not replace the outer row's)
        @"\trowd\cellx3000\pard\intbl " + Shape(@"\shpleft0\shptop0\shpright2880\shpbottom1440\shpwr1", Property("shapeType", "202"),
            @"\trowd\cellx1000\cellx2000\pard\intbl cellA\cell cellB\cell\row\pard after table") + @" outer\cell\row",
    ];

    public static void Run()
    {
        TestReport.Step("RTF shapes: hand-written .rtf");
        string rtf = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-drawn-shapes.rtf");
        string document = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}" + string.Concat(Body) + "}";
        if (!File.Exists(rtf) || File.ReadAllText(rtf) != document) File.WriteAllText(rtf, document, new UTF8Encoding(false));
        CheckHandWritten(rtf);

        TestReport.Step(@"RTF: a \par inside a header table (a Shift-JIS trail byte '{' in a style name, as LibreOffice fdo79384.rtf)");
        string trail = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-header-par.rtf");
        File.WriteAllBytes(trail, [.. Encoding.ASCII.GetBytes(@"{\rtf1{\fonttbl{\f5\fcharset128 X;}}{\stylesheet{\*\cs35 \f5 A"), 0x84, (byte)'{', .. Encoding.ASCII.GetBytes(";}\r\n}\r\n\\par text}")]);
        try
        {
            var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
            using IWordWindowedDocument opened = WordWindowedDocumentOpener.Open(trail, policy);
            int rendered = opened.RenderBlocks(0, opened.BlockCount).Count;
            TestReport.Check(rendered == opened.BlockCount, $"the index and a window segment alike ({opened.BlockCount} blocks)");
        }
        catch (Exception exception) { TestReport.Check(false, "the index and a window segment alike: " + exception.Message); }

        TestReport.Step("RTF shapes: the .odt shapes fixture saved as .rtf by LibreOffice");
        string fodt = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-shapes.fodt");
        if (!File.Exists(fodt)) OdtShapesFidelityTest.Run();
        if (File.Exists(fodt) && LibreOfficeConversion.Convert(fodt, ".rtf", "rtf") is { } saved) CheckLibreOfficeSaved(saved);
    }

    private static List<string> Render(string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    private static int SearchCount(string path, string query)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        return document.Search(query, new WordDocumentSearchOptions(), 0, document.BlockCount).HitCounts.Sum();
    }

    private static int Occurrences(string text, string part)
    {
        int count = 0;
        for (int at = text.IndexOf(part, StringComparison.Ordinal); at >= 0; at = text.IndexOf(part, at + part.Length, StringComparison.Ordinal)) count++;
        return count;
    }

    private static void CheckHandWritten(string path)
    {
        List<string> html = Render(path);
        TestReport.Check(html.Count == Body.Length, $"text-box paragraphs do not split blocks: {html.Count} blocks, expected {Body.Length}");
        if (html.Count < Body.Length) return;
        string all = string.Concat(html);
        TestReport.Check(!all.Contains("RESULTCOPY") && !all.Contains("GROUPCOPY"), "\\shprslt is skipped when \\shpinst is read");
        TestReport.Check(html[0].Contains("class=\"text-box\"") && html[0].Contains("width:144pt") && html[0].Contains("height:72pt") && html[0].Contains("float:left")
                         && html[0].Contains("border:2pt solid #008000") && html[0].Contains("background:#ffff00") && html[0].Contains("padding:0pt 0pt 0pt 0pt"),
                         "text box: size, square wrap, line, fill, insets");
        TestReport.Check(Occurrences(html[0], "text-box-paragraph") == 2 && html[0].Contains("boxtext one") && html[0].Contains("outside text"), "text box: two paragraphs, outside text kept");
        TestReport.Check(html[1].Contains("<path d=\"M72 18L71.83 19.76") && html[1].Contains("fill=\"#ff0000\"") && !html[1].Contains("stroke=\"#"), "ellipse (LibreOffice's msoEllipse: a full angle ellipse): red fill, no line (" + html[1] + ")");
        TestReport.Check(html[2].Contains("<path d=\"M0 9") && html[2].Contains("fill=\"none\"") && html[2].Contains("stroke=\"#0000ff\"") && html[2].Contains("arrowtext"),
                         "arrow: MSO outline, no fill, blue line, its text");
        TestReport.Check(html[3].Contains("left:0pt;top:0pt;width:144pt;height:72pt") && html[3].Contains("left:144pt;top:36pt;width:144pt;height:36pt")
                         && html[3].Contains("rotate(90)") && html[3].Contains("groupone") && html[3].Contains("grouptwo"), "group: children placed by relLeft.. in groupLeft.., rotation");
        TestReport.Check(html[4].Contains("<path d=\"M0 0L138.57 34.64\" fill=\"none\"") && Occurrences(html[4], "<path") == 2 && html[4].Contains("<path d=\"M144 36L"), "line with an end arrow (no fill: a line is not filled by default; the 2 pt minimum medium arrow, 6 pt long) (" + html[4] + ")");
        TestReport.Check(html[5].Contains("M36 0L72 72L0 72Z") && html[5].Contains("M12 0L12 0L10.82 0.03") && html[5].Contains("M36 0L72 36L36 72L0 36L36 0Z") && html[5].Contains("M16.23 48.37C28.93 61.73 43.07 61.73 55.77 48.37"),
                         "triangle, round rectangle (corner 3600 / 21600), diamond, smiley face (type 96: eyes and the mouth curve) (" + html[5] + ")");
        TestReport.Check(html[7].Contains("fill=\"#606060\"") && html[7].Contains("stroke=\"#ff0000\""), "system colours: button face C0C0C0 darkened by 128/256; a line that takes the fill colour (" + html[7] + ")");
        TestReport.Check(html[8].Contains("x1=\"36\" y1=\"0\" x2=\"36\" y2=\"36\"><stop offset=\"0\" stop-color=\"#0000ff\"/><stop offset=\"1\" stop-color=\"#ff0000\"/>"),
                         "linear shade: angle 0 runs top to bottom, the colours swapped by the angle and the zero focus (ImportGradientColor) (" + html[8] + ")");
        TestReport.Check(html[8].Contains("<stop offset=\"0\" stop-color=\"#ff0000\"/><stop offset=\"0.5\" stop-color=\"#0000ff\"/><stop offset=\"1\" stop-color=\"#ff0000\"/>"),
                         "axial shade (focus 50): the start colour on the axis");
        TestReport.Check(Occurrences(html[8], "word-officeart-shade-") >= 4 && html[8].Contains("word-officeart-pattern-") && html[8].Contains("word-shape-fill-picture")
                         && html[8].Contains("data-word-picture=\"8:0\""), "rectangular shades, the 8 x 8 pattern, the picture fill clipped to the ellipse");
        TestReport.Check(html[9].Contains("<path d=\"M0 0L72 0L36 72Z\"") && html[9].Contains("class=\"word-art\" style=\"display:block;text-align:center;font-family:'Arial';color:#ff0000;font-size:36pt;\"") && Occurrences(html[9], "<path") == 2
                         && SearchCount(path, "WordArtText") == 1, "custom geometry from pVerticies / pSegmentInfo; WordArt text in its font and size, found once (" + html[9] + ")");
        TestReport.Check(html[10].Contains("<td style=\"width:150pt;\">") && html[10].Contains("<td style=\"width:50pt;\"><span class=\"text-box-paragraph\"")
                         && html[10].Contains("cellA") && html[10].Contains("cellB") && html[10].Contains("after table") && SearchCount(path, "cellA") == 1 && SearchCount(path, "cellB") == 1,
                         "a table inside a text box is a table; the outer row keeps its own width (" + html[10] + ")");
        TestReport.Check(html[6].Contains("translate(72 0) rotate(90)"), "a rotated group turns its child about the group's centre");
        foreach ((string query, int expected) in new[] { ("boxtext", 1), ("second line", 1), ("outside", 1), ("arrowtext", 1), ("groupone", 1), ("grouptwo", 1), ("RESULTCOPY", 0), ("GROUPCOPY", 0) })
        {
            int count = SearchCount(path, query);
            TestReport.Check(count == expected, $"search '{query}': {count}, expected {expected}");
        }
        TestReport.Check(SearchCount(path, "one second") == 0, "a text-box paragraph end is a separator");
    }

    private static void CheckLibreOfficeSaved(string path)
    {
        List<string> html = Render(path);
        string all = string.Concat(html);
        foreach (string word in new[] { "boxtext one", "inlinebox", "centrebox", "rtlbox", "shapetext", "gradtext" })   // LibreOffice's RTF export drops the draw:g group (no "groupone" in the file)
        {
            TestReport.Check(all.Contains(word), $"LibreOffice .rtf: '{word}' rendered");
            int count = SearchCount(path, word);
            TestReport.Check(count == 1, $"LibreOffice .rtf: search '{word}' {count}, expected 1");
        }
        TestReport.Check(all.Contains("<svg") && all.Contains("text-box"), "LibreOffice .rtf: shapes drawn, text boxes placed");
        TestReport.Check(all.Contains("word-officeart-gradient-"), "LibreOffice .rtf: the gradient rectangle is a gradient");
    }
}
