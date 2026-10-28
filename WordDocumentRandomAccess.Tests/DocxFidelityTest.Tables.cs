namespace WordDocumentRandomAccess.Tests;

public static partial class DocxFidelityTest
{
    private static string Cell(string text) => $"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
    private static string Row(params string[] texts) => "<w:tr>" + string.Concat(texts.Select(Cell)) + "</w:tr>";

    private const string BandedStyle = """
        <w:style w:type="table" w:styleId="Grid"><w:name w:val="Grid"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="18"/></w:rPr>
          <w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="4472C4"/></w:tcPr></w:tblStylePr>
          <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr></w:tblStylePr>
          <w:tblStylePr w:type="lastCol"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:tblStylePr>
          <w:tblStylePr w:type="neCell"><w:rPr><w:i/></w:rPr></w:tblStylePr>
        </w:style>
        """;

    /// <summary>
    /// Table style conditional parts (tblStylePr) chosen and layered as LibreOffice's DomainMapperTableHandler does, and its
    /// rule for the default paragraph style's 11 / 12 pt size and left alignment (overrideTableStyleFontSizeAndJustification).
    /// </summary>
    private static void TableConditionalFormatting()
    {
        string table = "<w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/><w:tblLook w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"1\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr>"
                     + "<w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>" + Row("h1", "h2") + Row("r1", "r1e") + Row("r2", "r2e") + Row("r3", "r3e") + "</w:tbl>";
        string html = RenderAll("table-conditions", table, BandedStyle)[0];
        string CellOf(string text) { int at = html.IndexOf(">" + text + "<", StringComparison.Ordinal); return html.Substring(html.LastIndexOf("<td", at, StringComparison.Ordinal), at - html.LastIndexOf("<td", at, StringComparison.Ordinal)); }
        TestReport.Check(CellOf("h1").Contains("background:#4472C4") && CellOf("h1").Contains("font-weight:bold"), "firstRow: shading and bold");
        TestReport.Check(CellOf("h2").Contains("font-style:italic") && CellOf("h2").Contains("color:#FF0000"), "neCell + lastCol on the top-right cell");
        TestReport.Check(CellOf("r1").Contains("background:#D9E2F3") && !CellOf("r2").Contains("background:#D9E2F3") && CellOf("r3").Contains("background:#D9E2F3"), "band1Horz on odd body rows only");
        TestReport.Check(CellOf("r1e").Contains("color:#FF0000") && !CellOf("r1").Contains("color:#FF0000"), "lastCol only on the last cell");

        // default paragraph style 12 pt + left: the table style's 9 pt and centering win, unless overrideTableStyleFontSizeAndJustification
        const string normal12 = """<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>""";
        string styles = normal12 + BandedStyle.Replace("<w:style w:type=\"table\"", "<w:style w:type=\"paragraph\" w:styleId=\"Body\"><w:name w:val=\"Body\"/><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:rPr><w:sz w:val=\"24\"/></w:rPr></w:style><w:style w:type=\"table\"");
        string plainTable = "<w:tbl><w:tblPr><w:tblStyle w:val=\"Grid\"/><w:tblLook w:val=\"0000\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"1000\"/></w:tblGrid>"
                          + "<w:tr><w:tc><w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        string defaultIsBody = styles.Replace("w:styleId=\"Body\"", "w:default=\"1\" w:styleId=\"Body\"");
        string overridden = RenderAll("table-compat-default", plainTable, defaultIsBody)[0];
        string kept = RenderAll("table-compat-override", plainTable, defaultIsBody, settings: "<w:compat><w:compatSetting w:name=\"overrideTableStyleFontSizeAndJustification\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"1\"/></w:compat>")[0];
        TestReport.Check(overridden.Contains("font-size:9pt") && overridden.Contains("text-align:center"), "default paragraph style 12 pt / left: the table style's size and alignment win");
        TestReport.Check(kept.Contains("font-size:12pt") && !kept.Contains("text-align:center"), "overrideTableStyleFontSizeAndJustification: the paragraph style keeps them");
    }
}
