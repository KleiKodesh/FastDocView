namespace WordDocumentRandomAccess.Tests;

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
}
