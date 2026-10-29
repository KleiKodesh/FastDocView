namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>Tab stops with leaders (ST_TabJc / ST_TabTlc): a last tab to an end-aligned stop becomes a flex filler with the leader line.</summary>
    private static void TabLeaders()
    {
        const string styles = """<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr></w:style>""";
        static string Line(string style, string paragraphProperties, string tabRun)
            => $"<w:p><w:pPr>{(style.Length > 0 ? $"<w:pStyle w:val=\"{style}\"/>" : "")}{paragraphProperties}</w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Chapter</w:t></w:r>{tabRun}<w:r><w:t>12</w:t></w:r></w:p>";
        List<string> html = RenderAll("tabs",
            Line("TOC1", "", "<w:r><w:tab/></w:r>")                                                                     // 0 style right/dot
          + Line("TOC1", "<w:tabs><w:tab w:val=\"clear\" w:pos=\"9350\"/><w:tab w:val=\"left\" w:pos=\"720\"/></w:tabs>", "<w:r><w:tab/></w:r>")   // 1 cleared
          + Line("", "", "<w:r><w:ptab w:relativeTo=\"margin\" w:alignment=\"right\" w:leader=\"hyphen\"/></w:r>"), styles);   // 2 ptab
        TestReport.Check(html[0].Contains("display:flex") && html[0].Contains("class=\"tab-leader\" style=\"flex:1 1 0;min-width:0.5em;margin:0 0.2em;border-bottom:0.15em dotted currentColor;\"")
                         && html[0].Contains("Chapter</span></span><span class=\"tab-leader\"") && html[0].Contains("<span class=\"tab-after\" style=\"flex:none\">"),
                         "right tab with a dot leader: text, dotted filler, the rest at the end (the run span is closed and reopened)");
        TestReport.Check(!html[1].Contains("display:flex") && html[1].Contains("\t"), "a cleared end stop leaves a plain tab");
        TestReport.Check(html[2].Contains("dashed"), "ptab alignment right with a hyphen leader");
    }
}
