using System.IO.Compression;
using System.Text;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// Placement of list label gaps, tables and pictures in every format that has them (--fidelity, seconds):
/// .docx (w:jc / w:tblInd, wp:inline / wp:anchor), .odt (table:align, draw:frame + style:wrap), .rtf (\trleft / \trqc, \pict),
/// and the list label's least gap after a tab suffix.
/// </summary>
public static class PagePlacementFidelityTest
{
    /// <summary>A 2 x 1 red PNG.</summary>
    private static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAJ/gH/0wkYjQAAAABJRU5ErkJggg==");

    public static void Run()
    {
        ListLabelGap();
        DocxTables();
        DocxPictures();
        OdtTablesAndFrames();
        RtfRowsAndPictures();
    }

    /// <summary>A tab suffix reaches its stop (min-width) and keeps a quarter em after a label that fills the gap; space and nothing do not.</summary>
    private static void ListLabelGap()
    {
        TestReport.Step("Placement: list label gaps");
        var paragraph = new WordParagraphFormatting { LeftIndentTwips = 720, FirstLineIndentTwips = -360 };
        var tab = new StringBuilder();
        WordListLabelHtml.Append(tab, "1.", null, WordListLabelSuffix.Tab, paragraph, 720);
        TestReport.Check(tab.ToString() == "<span class=\"list-marker\" style=\"min-width:18pt;\"><span style=\"padding-inline-end:0.25em\">1.</span></span>",
                         "tab: the label box reaches the hanging indent and keeps a least gap (" + tab + ")");
        var space = new StringBuilder();
        WordListLabelHtml.Append(space, "1.", null, WordListLabelSuffix.Space, paragraph, 720);
        TestReport.Check(space.ToString() == "<span class=\"list-marker\">1.</span> ", "space: a space after the label (" + space + ")");
        var nothing = new StringBuilder();
        WordListLabelHtml.Append(nothing, "1.", null, WordListLabelSuffix.Nothing, paragraph, 720);
        TestReport.Check(nothing.ToString() == "<span class=\"list-marker\">1.</span>", "nothing: the text follows the label (" + nothing + ")");
        string stylesheet = WordDocumentHtmlStylesheet.Build(new WordRunFormatting(), isRightToLeft: true);
        TestReport.Check(stylesheet.Contains(".word-document .list-marker { display: inline-block;"), "the label is an inline-block (its min-width applies)");
    }

    private static string Table(string properties)
        => "<w:tbl><w:tblPr><w:tblW w:w=\"2880\" w:type=\"dxa\"/>" + properties + "</w:tblPr><w:tblGrid><w:gridCol w:w=\"2880\"/></w:tblGrid>"
         + "<w:tr><w:tc><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    /// <summary>w:jc (LibreOffice convertTableJustification), w:tblInd less the cell margin before Word 2013, from the table's own start edge.</summary>
    private static void DocxTables()
    {
        TestReport.Step("Placement: .docx tables");
        const string margins = "<w:tblCellMar><w:left w:w=\"108\" w:type=\"dxa\"/></w:tblCellMar>";
        List<string> html = DocxFidelityTest.RenderAll("placement-tables",
            Table("<w:jc w:val=\"center\"/>") + Table("<w:jc w:val=\"end\"/>") + Table("<w:tblInd w:w=\"720\" w:type=\"dxa\"/>" + margins)
          + Table("<w:bidiVisual/><w:tblInd w:w=\"720\" w:type=\"dxa\"/>" + margins) + Table(""));
        TestReport.Check(html[0].Contains("margin-inline:auto;"), "jc center: centred (" + html[0] + ")");
        TestReport.Check(html[1].Contains("margin-inline-start:auto;"), "jc end: at the end side");
        TestReport.Check(html[2].Contains("margin-inline-start:30.6pt;margin-inline-end:auto;"), "tblInd 720 - cell margin 108 (compatibility mode 12) from the start (" + html[2] + ")");
        TestReport.Check(html[3].Contains("dir=\"rtl\"") && html[3].Contains("margin-inline-start:30.6pt;margin-inline-end:auto;"), "bidiVisual: the indent from the right (the table's own start)");
        TestReport.Check(html[4].Contains("margin-inline-end:auto;") && !html[4].Contains("margin-inline-start"), "no indent: at the table's start edge");
        List<string> modern = DocxFidelityTest.RenderAll("placement-tables-15", Table("<w:tblInd w:w=\"720\" w:type=\"dxa\"/>" + margins),
                                                         settings: "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/></w:compat>");
        TestReport.Check(modern[0].Contains("margin-inline-start:36pt;"), "compatibility mode 15: the indent is the table edge (" + modern[0] + ")");
    }

    private const string DrawingNamespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"";

    private static string Graphic => "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic><pic:blipFill><a:blip r:embed=\"rIdP\"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>";

    private static string Anchored(string wrap, string position, string behind = "0")
        => $"<w:p><w:r><w:drawing><wp:anchor {DrawingNamespaces} distT=\"127000\" distB=\"127000\" distL=\"254000\" distR=\"254000\" behindDoc=\"{behind}\" simplePos=\"0\" relativeHeight=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
         + $"<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\">{position}</wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>"
         + $"<wp:extent cx=\"914400\" cy=\"457200\"/>{wrap}<wp:docPr id=\"1\" name=\"p\"/>{Graphic}</wp:anchor></w:drawing></w:r><w:r><w:t>text</w:t></w:r></w:p>";

    /// <summary>wp:inline in the line, wp:anchor square / topAndBottom / none placed as WordPictureHtml documents, the picture from its relationship.</summary>
    private static void DocxPictures()
    {
        TestReport.Step("Placement: .docx pictures");
        string body = $"<w:p><w:r><w:drawing><wp:inline {DrawingNamespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"p\"/>{Graphic}</wp:inline></w:drawing></w:r></w:p>"
                    + Anchored("<wp:wrapSquare wrapText=\"bothSides\"/>", "<wp:align>right</wp:align>")
                    + Anchored("<wp:wrapTight wrapText=\"bothSides\"><wp:wrapPolygon edited=\"0\"><wp:start x=\"0\" y=\"0\"/></wp:wrapPolygon></wp:wrapTight>", "<wp:align>left</wp:align>")
                    + Anchored("<wp:wrapTopAndBottom/>", "<wp:align>center</wp:align>")
                    + Anchored("<wp:wrapNone/>", "<wp:posOffset>127000</wp:posOffset>", behind: "1")
                    + Anchored("<wp:wrapSquare wrapText=\"bothSides\"/>", "<wp:align>center</wp:align>");
        List<string> html = DocxFidelityTest.RenderAll("placement-pictures", body,
            extraDocumentRelationships: "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = Png });
        TestReport.Check(html[0].Contains("<img class=\"word-picture\" alt=\"\" src=\"data:image/png;base64," + Convert.ToBase64String(Png) + "\" style=\"width:72pt;aspect-ratio:72/36;\">"),
                         "inline: the picture from its relationship at its extent, in the line (" + html[0] + ")");
        TestReport.Check(html[1].Contains("float:right;margin:10pt 0pt 10pt 20pt;"), "wrapSquare right: floats right, the gap only on the text side (" + html[1] + ")");
        TestReport.Check(html[2].Contains("float:left;margin:10pt 20pt 10pt 0pt;"), "wrapTight left: floats left");
        TestReport.Check(html[3].Contains("display:block;margin:10pt auto 10pt auto;"), "wrapTopAndBottom center: a centred block with the top / bottom gap");
        TestReport.Check(html[4].Contains("display:block;margin:0pt auto 0pt 10pt;"), "wrapNone: a block without gaps, at its offset from the left (" + html[4] + ")");
        TestReport.Check(html[5].Contains("display:block;margin:10pt auto 10pt auto;"), "wrapSquare centred: a centred block (CSS cannot float centre)");
        TestReport.Check(html.All(block => !block.Contains("[image]")), "every picture is drawn");
    }

    /// <summary>table:align left / right / center / margins (ODF 1.3 20.414) and draw:frame pictures by anchor type and style:wrap (19.759, 20.400, 20.298); bullets without prefix / suffix.</summary>
    private static void OdtTablesAndFrames()
    {
        TestReport.Step("Placement: .odt tables, frames and bullets");
        const string styles = """
            <office:automatic-styles>
              <style:style style:name="TRight" style:family="table"><style:table-properties style:width="2in" fo:margin-right="0.5in" table:align="right"/></style:style>
              <style:style style:name="TLeft" style:family="table"><style:table-properties style:width="2in" fo:margin-left="1in" table:align="left"/></style:style>
              <style:style style:name="TCenter" style:family="table"><style:table-properties style:width="2in" table:align="center"/></style:style>
              <style:style style:name="Square" style:family="graphic"><style:graphic-properties style:wrap="parallel" style:horizontal-pos="right" fo:margin-left="0.25in" fo:margin-right="0.25in" fo:margin-top="0.1in" fo:margin-bottom="0.1in"/></style:style>
              <style:style style:name="Block" style:family="graphic"><style:graphic-properties style:wrap="none" style:horizontal-pos="center"/></style:style>
              <text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" style:num-suffix="&#x2022;" text:bullet-char="&#x2022;"/></text:list-style>
            </office:automatic-styles>
            """;
        string picture = "<draw:image xlink:href=\"Pictures/p.png\"/>";
        string body = "<table:table table:style-name=\"TRight\"><table:table-column/><table:table-row><table:table-cell><text:p>r</text:p></table:table-cell></table:table-row></table:table>"
                    + "<table:table table:style-name=\"TLeft\"><table:table-column/><table:table-row><table:table-cell><text:p>l</text:p></table:table-cell></table:table-row></table:table>"
                    + "<table:table table:style-name=\"TCenter\"><table:table-column/><table:table-row><table:table-cell><text:p>c</text:p></table:table-cell></table:table-row></table:table>"
                    + "<text:p><draw:frame draw:style-name=\"Square\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.5in\">" + picture + "</draw:frame>text</text:p>"
                    + "<text:p><draw:frame draw:style-name=\"Block\" text:anchor-type=\"paragraph\" svg:width=\"1in\" svg:height=\"0.5in\">" + picture + "</draw:frame>text</text:p>"
                    + "<text:p>in <draw:frame draw:style-name=\"Square\" text:anchor-type=\"as-char\" svg:width=\"1in\" svg:height=\"0.5in\">" + picture + "</draw:frame> line</text:p>"
                    + "<text:list text:style-name=\"L1\"><text:list-item><text:p>bullet</text:p></text:list-item></text:list>";
        string content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
                       + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
                       + "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" "
                       + "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" "
                       + "xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\">" + styles + "<office:body><office:text>" + body + "</office:text></office:body></office:document-content>";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-placement.odt");
        if (File.Exists(path)) File.Delete(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            void Add(string name, byte[] bytes, CompressionLevel level)
            {
                using Stream stream = zip.CreateEntry(name, level).Open();
                stream.Write(bytes, 0, bytes.Length);
            }
            Add("mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text"), CompressionLevel.NoCompression);
            Add("content.xml", new UTF8Encoding(false).GetBytes(content), CompressionLevel.Optimal);
            Add("Pictures/p.png", Png, CompressionLevel.NoCompression);
        }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        TestReport.Check(html[0].Contains("margin-left:auto;margin-right:36pt;"), "table:align right: at the right margin, fo:margin-right kept (" + html[0] + ")");
        TestReport.Check(html[1].Contains("margin-left:72pt;margin-right:auto;"), "table:align left: fo:margin-left from the left");
        TestReport.Check(html[2].Contains("margin-left:auto;margin-right:auto;"), "table:align center");
        TestReport.Check(html[3].Contains("src=\"data:image/png;base64,") && html[3].Contains("float:right;margin:7.2pt 0pt 7.2pt 18pt;"),
                         "frame, wrap parallel, right: floats right with its fo:margin gap, the picture from the package (" + html[3] + ")");
        TestReport.Check(html[4].Contains("display:block;margin:0pt auto 0pt auto;"), "frame, wrap none, center: a centred block");
        TestReport.Check(html[5].Contains("style=\"width:72pt;aspect-ratio:72/36;margin:7.2pt 18pt 7.2pt 18pt;\"") && !html[5].Contains("float"), "anchor-type as-char: in the line, with its fo:margin spacing (" + html[5] + ")");
        TestReport.Check(html[6].Contains(">•<") && !html[6].Contains("••"), "a bullet is drawn alone (no num-suffix) (" + html[6] + ")");
    }

    /// <summary>\trleft / \trqc (LibreOffice: tblInd / jc), the first cell from \trleft; a {\pict \pngblip} at its goal size.</summary>
    private static void RtfRowsAndPictures()
    {
        TestReport.Step("Placement: .rtf rows and pictures");
        string hex = Convert.ToHexString(Png);
        string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Arial;}}" + "\n"
                   + @"\trowd\trleft720\cellx2160\cellx3600\pard\intbl a\cell b\cell\row" + "\n"
                   + @"\trowd\trqc\cellx1440\pard\intbl c\cell\row" + "\n"
                   + @"\pard x {\*\shppict{\pict\pngblip\picw2\pich1\picwgoal1440\pichgoal720\picscalex50\picscaley50 " + hex + @"}}{\nonshppict{\pict\wmetafile8 00}} y\par" + "\n"
                   + @"\pard {\pict\emfblip\picwgoal720\pichgoal720 0102}\par" + "\n}";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-placement.rtf");
        File.WriteAllText(path, rtf);
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        TestReport.Check(html[0].Contains("width:144pt;margin-inline-start:36pt;margin-inline-end:auto;") && html[0].Contains("<td style=\"width:72pt;"),
                         "\\trleft: the row's indent, the first cell from it (" + html[0] + ")");
        TestReport.Check(html[1].Contains("margin-left:auto;margin-right:auto;"), "\\trqc: a centred row (" + html[1] + ")");
        TestReport.Check(html[2].Contains("src=\"data:image/png;base64," + Convert.ToBase64String(Png) + "\" style=\"width:36pt;aspect-ratio:36/18;\"")
                         && html[2].Split("word-picture").Length == 2, "\\pict \\pngblip in \\shppict: drawn once, at goal x scale (" + html[2] + ")");
        TestReport.Check(html[3].Contains("word-picture-missing") && html[3].Contains("width:36pt;"), "an EMF keeps its box (" + html[3] + ")");
    }
}
