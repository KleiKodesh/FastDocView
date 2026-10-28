namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>w:shd patterns mixed as LibreOffice CellColorHandler does: color share (thousandths) over fill; highlight wins over run shading.</summary>
    private static void ShadingPatterns()
    {
        string body = Paragraph(null, "", "pct25", "<w:shd w:val=\"pct25\" w:color=\"000000\" w:fill=\"FFFFFF\"/>")
                    + Paragraph(null, "", "solid", "<w:shd w:val=\"solid\" w:color=\"FF0000\" w:fill=\"00FF00\"/>")
                    + Paragraph(null, "<w:shd w:val=\"clear\" w:fill=\"00FF00\"/><w:highlight w:val=\"yellow\"/>", "run")
                    + Paragraph(null, "", "nil", "<w:shd w:val=\"nil\" w:fill=\"FF0000\"/>")
                    + "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:val=\"pct50\" w:color=\"0000FF\" w:fill=\"auto\"/></w:tcPr>"
                    + Paragraph(null, "", "cell") + "</w:tc></w:tr></w:tbl>";
        List<string> html = RenderAll("shading", body);
        TestReport.Check(html[0].Contains("background-color:#BFBFBF;"), "pct25 black over white → #BFBFBF");
        TestReport.Check(html[1].Contains("background-color:#FF0000;"), "solid → the pattern color");
        TestReport.Check(html[2].Contains("background-color:yellow;") && !html[2].Contains("#00FF00"), "a highlight is drawn instead of run shading");
        TestReport.Check(!html[3].Contains("background"), "shd nil: no shading");
        TestReport.Check(html[4].Contains("background:#7F7FFF;"), "cell pct50 blue over auto (white) fill → #7F7FFF");
    }
}
