namespace WordDocumentRandomAccess.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>w:sym keeps its font (LibreOffice: run font = w:font, symbol charset); codes without the F0 byte move to U+F0xx.</summary>
    private static void Symbols()
    {
        List<string> html = RenderAll("symbols", "<w:p><w:r><w:sym w:font=\"Wingdings\" w:char=\"F04A\"/><w:sym w:font=\"Symbol\" w:char=\"0061\"/><w:sym w:char=\"F041\"/></w:r></w:p>");
        TestReport.Check(html[0].Contains("<span style=\"font-family:'Wingdings'\"></span>"), "w:sym keeps w:font and the private-use code");
        TestReport.Check(html[0].Contains("<span style=\"font-family:'Symbol'\"></span>"), "w:sym 0061 in a symbol font → U+F061");
        TestReport.Check(html[0].Contains("</span>A"), "w:sym without a font: the plain character");
    }
}
