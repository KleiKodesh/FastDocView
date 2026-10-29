namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    private const string Mode15 = "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>";
    private const string Mode14 = "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"14\"/></w:compat>";

    /// <summary>Compatibility-dependent layout read from LibreOffice dmapper (see DocxCompatibilitySettings and the renderer comments).</summary>
    private static void CompatibilityLayout()
    {
        // 2. a justified line ending in a manual break is stretched, unless doNotExpandShiftReturn
        string justified = "<w:p><w:pPr><w:jc w:val=\"both\"/><w:ind w:firstLine=\"360\"/></w:pPr><w:r><w:t>one</w:t><w:br/><w:t>two</w:t></w:r></w:p>";
        string stretched = RenderAll("compat-justify", justified)[0];
        string plain = RenderAll("compat-justify-off", justified, settings: "<w:compat><w:doNotExpandShiftReturn/></w:compat>")[0];
        TestReport.Check(stretched.Contains("<span class=\"justified-lines\" style=\"display:block;text-align-last:justify\">one</span><span class=\"after-justified-lines\" style=\"display:block;text-indent:0\">two</span>"),
                         "justified paragraph: the line before a manual break is stretched, the last line is not");
        TestReport.Check(!plain.Contains("justified-lines") && plain.Contains("one<br>two"), "doNotExpandShiftReturn: a plain break");

        // 3. autospacing: 280 twips (14pt), 100 (5pt) with doNotUseHTMLParagraphAutoSpacing; none before the document's first paragraph
        string auto = "<w:p><w:pPr><w:spacing w:before=\"0\" w:beforeAutospacing=\"1\" w:after=\"0\" w:afterAutospacing=\"1\"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>";
        List<string> automatic = RenderAll("compat-autospacing", auto + auto);
        List<string> small = RenderAll("compat-autospacing-html", auto + auto, settings: "<w:compat><w:doNotUseHTMLParagraphAutoSpacing/></w:compat>");
        TestReport.Check(!automatic[0].Contains("margin-top") && automatic[0].Contains("margin-bottom:14pt") && automatic[1].Contains("margin-top:14pt"), "autospacing: 14pt, none before the first paragraph");
        TestReport.Check(small[0].Contains("margin-top:5pt") && small[1].Contains("margin-bottom:5pt"), "doNotUseHTMLParagraphAutoSpacing: 5pt and no adjustments");

        // 4. mode 15: a paragraph that starts with a page break has no space before
        string broken = "<w:p><w:pPr><w:spacing w:before=\"240\"/></w:pPr><w:r><w:br w:type=\"page\"/><w:t>after</w:t></w:r></w:p>";
        TestReport.Check(!RenderAll("compat-break-15", broken, settings: Mode15)[0].Contains("margin-top") && RenderAll("compat-break-14", broken, settings: Mode14)[0].Contains("margin-top:12pt"),
                         "mode 15: no space before a paragraph that starts with a page break (mode 14 keeps it)");

        // 5. mode 11-14 (stated): trailing spaces of a table-cell paragraph are removed
        string cell = "<w:tbl><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc><w:p><w:pPr><w:jc w:val=\"right\"/></w:pPr><w:r><w:t xml:space=\"preserve\">right   </w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        TestReport.Check(RenderAll("compat-cell-14", cell, settings: Mode14)[0].Contains(">right</p>") && RenderAll("compat-cell-15", cell, settings: Mode15)[0].Contains(">right   </p>"),
                         "mode 14: trailing spaces in a table cell are removed (mode 15 keeps them)");
    }
}
