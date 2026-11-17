namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>Tab stops with leaders (ST_TabJc / ST_TabTlc): the stops and the absolute position tab become data for the page's tab pass.</summary>
    private static void TabLeaders()
    {
        const string styles = """<w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/></w:tabs></w:pPr></w:style>""";
        static string Line(string style, string paragraphProperties, string tabRun)
            => $"<w:p><w:pPr>{(style.Length > 0 ? $"<w:pStyle w:val=\"{style}\"/>" : "")}{paragraphProperties}</w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Chapter</w:t></w:r>{tabRun}<w:r><w:t>12</w:t></w:r></w:p>";
        List<string> html = RenderAll("tabs",
            Line("TOC1", "", "<w:r><w:tab/></w:r>")                                                                     // 0 style right/dot
          + Line("TOC1", "<w:tabs><w:tab w:val=\"clear\" w:pos=\"9350\"/><w:tab w:val=\"left\" w:pos=\"720\"/></w:tabs>", "<w:r><w:tab/></w:r>")   // 1 cleared
          + Line("", "", "<w:r><w:ptab w:relativeTo=\"margin\" w:alignment=\"right\" w:leader=\"hyphen\"/></w:r>"), styles);   // 2 ptab
        TestReport.Check(html[0].Contains("data-tabs=\"36 0 0 0;467.5 r d\"") && html[0].Contains("	"),
                         "right tab with a dot leader: the style's stop goes to the page's tab pass as data, the tab stays a tab");
        TestReport.Check(html[1].Contains("data-tabs=\"36 0 0 0;36 l \"") && html[1].Contains("	"), "a cleared end stop leaves only the paragraph's own left stop");
        TestReport.Check(html[2].Contains("data-ptab=\"r h m\""), "ptab alignment right with a hyphen leader, relative to the margin");
    }
}
