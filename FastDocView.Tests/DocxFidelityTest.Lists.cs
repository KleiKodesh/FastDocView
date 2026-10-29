namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    private static string ListParagraph(int numId, int level, string text)
        => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"{numId}\"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

    private static string Marker(string html)
    {
        int start = html.IndexOf("list-marker", StringComparison.Ordinal);
        if (start < 0) return "";
        start = html.IndexOf('>', start) + 1;
        return html.Substring(start, html.IndexOf("</span>", start, StringComparison.Ordinal) - start);
    }

    /// <summary>§17.9: lvlOverride/lvl replaces the abstract level, numStyleLink follows the numbering style, lvlRestart 0 never restarts, isLgl shows decimals.</summary>
    private static void ListNumbering()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="0">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperLetter"/><w:lvlText w:val="%1"/></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2"/></w:lvl>
              <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:isLgl/><w:lvlText w:val="%1.%2.%3"/></w:lvl>
            </w:abstractNum>
            <w:abstractNum w:abstractNumId="1">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:lvlRestart w:val="0"/><w:numFmt w:val="decimal"/><w:lvlText w:val="(%2)"/></w:lvl>
            </w:abstractNum>
            <w:abstractNum w:abstractNumId="2"><w:numStyleLink w:val="ListStyle"/></w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
            <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
            <w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
            <w:num w:numId="4"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/></w:lvl></w:lvlOverride></w:num>
            """;
        const string styles = """<w:style w:type="numbering" w:styleId="ListStyle"><w:name w:val="ListStyle"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style>""";
        string body = ListParagraph(1, 0, "a") + ListParagraph(1, 1, "b") + ListParagraph(1, 1, "c") + ListParagraph(1, 2, "d")   // A, A.a, A.b, 1.2.1
                    + ListParagraph(2, 1, "e") + ListParagraph(2, 0, "f") + ListParagraph(2, 1, "g")                               // (1), 1., (2) — no restart
                    + ListParagraph(3, 0, "h")                                                                                     // numStyleLink → abstract 1, continues: 2.
                    + ListParagraph(4, 0, "i");                                                                                    // lvlOverride lvl: upperRoman → I)
        List<string> html = RenderAll("lists", body, styles, numbering);
        List<string> markers = html.Select(Marker).ToList();
        TestReport.Check(markers.Take(4).SequenceEqual(["A", "A.a", "A.b", "1.2.1"]), "isLgl: every level of that label in decimal (" + string.Join(" | ", markers.Take(4)) + ")");
        TestReport.Check(markers[4] == "(1)" && markers[5] == "1." && markers[6] == "(2)", "lvlRestart 0: the level never restarts (" + string.Join(" | ", markers.Skip(4).Take(3)) + ")");
        TestReport.Check(markers[7] == "2.", "numStyleLink: the list continues the linked style's abstractNum (" + markers[7] + ")");
        TestReport.Check(markers[8] == "I)", "lvlOverride/lvl replaces the abstract level (" + markers[8] + ")");
    }

    /// <summary>
    /// List label drawing: suffix (§17.9.28: tab by default, space, nothing), the level's rPr (§17.9.6) and the paragraph mark's
    /// rPr — which formats the label and an empty paragraph's height but not the paragraph's runs (§17.3.1.29).
    /// </summary>
    private static void ListLabelsAndParagraphMark()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="0">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="space"/><w:lvlText w:val="%2)"/></w:lvl>
              <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:suff w:val="nothing"/><w:lvlText w:val="%3]"/></w:lvl>
              <w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl>
            </w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
            """;
        string body = ListParagraph(1, 0, "a") + ListParagraph(1, 1, "b") + ListParagraph(1, 2, "c") + ListParagraph(1, 3, "d")
                    + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr><w:rPr><w:i/></w:rPr></w:pPr><w:r><w:t>e</w:t></w:r></w:p>"
                    + "<w:p><w:pPr><w:rPr><w:sz w:val=\"48\"/></w:rPr></w:pPr></w:p>";
        List<string> html = RenderAll("list-labels", body, "", numbering);
        TestReport.Check(html[0].Contains("<span class=\"list-marker\" style=\"font-weight:bold;color:#FF0000;min-width:36pt;\">1.</span>"), "level rPr formats the label; default suffix is a tab");
        TestReport.Check(html[1].Contains(">1)</span> a") == false && html[1].Contains(">1)</span> b"), "suff space");
        TestReport.Check(html[2].Contains(">1]</span>c"), "suff nothing");
        TestReport.Check(html[3].Contains(">•</span>") && !html[3].Contains("Symbol"), "a Symbol-font bullet is mapped to a bullet and drops the symbol font");
        TestReport.Check(html[4].Contains("style=\"font-style:italic;\">2)</span>") && !html[4].Contains("<p style=\"font-style:italic"), "paragraph-mark rPr formats the label, not the text");
        TestReport.Check(html[5].Contains("font-size:24pt;"), "an empty paragraph takes its height from the mark's size");
    }

    /// <summary>
    /// The numbering level's pPr (17.9.23) indents a list paragraph: direct numbering overrides the paragraph style's
    /// indent, numbering from the style leaves the style's indent in charge (LibreOffice DomainMapper_Impl), direct w:ind
    /// wins over both; the tab after the label reaches the hanging indent; lvlJc right ends the label at its position.
    /// </summary>
    private static void ListLevelIndents()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="0">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
            </w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
            """;
        const string styles = """
            <w:style w:type="paragraph" w:styleId="Indented"><w:name w:val="Indented"/><w:pPr><w:ind w:left="2000"/></w:pPr></w:style>
            <w:style w:type="paragraph" w:styleId="NumberedIndented"><w:name w:val="NumberedIndented"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:ind w:left="3000"/></w:pPr></w:style>
            """;
        string body = ListParagraph(1, 0, "a")
                    + "<w:p><w:pPr><w:pStyle w:val='Indented'/><w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr></w:pPr><w:r><w:t>b</w:t></w:r></w:p>"
                    + "<w:p><w:pPr><w:pStyle w:val='NumberedIndented'/></w:pPr><w:r><w:t>c</w:t></w:r></w:p>"
                    + "<w:p><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='1'/></w:numPr><w:ind w:left='1000' w:hanging='500'/></w:pPr><w:r><w:t>d</w:t></w:r></w:p>"
                    + ListParagraph(1, 1, "e");
        List<string> html = RenderAll("list-indents", body, styles, numbering);
        TestReport.Check(html[0].Contains("margin-inline-start:36pt;") && html[0].Contains("text-indent:-18pt;") && html[0].Contains("min-width:18pt;"), "level pPr indents the paragraph; the tab reaches the hanging indent");
        TestReport.Check(html[1].Contains("margin-inline-start:36pt;"), "direct numbering overrides the paragraph style's indent");
        TestReport.Check(html[2].Contains("margin-inline-start:150pt;") && html[2].Contains("text-indent:-18pt;"), "numbering from the style keeps the style's indent, takes the level's hanging");
        TestReport.Check(html[3].Contains("margin-inline-start:50pt;") && html[3].Contains("text-indent:-25pt;") && html[3].Contains("min-width:25pt;"), "direct w:ind wins over the level");
        TestReport.Check(html[4].Contains("justify-content:flex-end"), "lvlJc right: the label ends at its position");
    }
}
