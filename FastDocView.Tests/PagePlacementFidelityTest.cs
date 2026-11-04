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
        DocxVmlPictures();
        RtfFloatingShapes();
        DocBinaryPicturesAndRows();
        DocxTextBoxes();
        DocxShapesFidelityTest.Run();
        DocxDrawingFidelityTest.Run();
        DocxDrawingLayoutFidelityTest.Run();
    }

    /// <summary>The picture ids of a block's &lt;img&gt; elements (data-word-picture), in order.</summary>
    internal static List<string> PictureIds(string html)
        => System.Text.RegularExpressions.Regex.Matches(html, WordPictureHtml.PictureIdAttribute + "=\"([^\"]*)\"")
               .Select(match => System.Net.WebUtility.HtmlDecode(match.Groups[1].Value)).ToList();

    /// <summary>No picture bytes in the block HTML: no data: URI, an id, lazy loading and its size.</summary>
    private static bool IsLazyPicture(string html)
        => !html.Contains("data:image") && html.Contains(WordPictureHtml.PictureIdAttribute + "=") && html.Contains("loading=\"lazy\" decoding=\"async\"");

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

    private static string TextBoxAnchor(string wrap, string position, string spPr, string bodyPr)
        => $"<w:p><w:r><w:drawing><wp:anchor {DrawingNamespaces} distT=\"0\" distB=\"0\" distL=\"127000\" distR=\"127000\" behindDoc=\"0\" simplePos=\"0\" relativeHeight=\"1\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
         + $"<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\">{position}</wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>"
         + $"<wp:extent cx=\"1270000\" cy=\"635000\"/>{wrap}<wp:docPr id=\"2\" name=\"Text Box 2\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
         + $"<wps:wsp xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:cNvSpPr txBox=\"1\"/><wps:spPr>{spPr}</wps:spPr>"
         + "<wps:txbx><w:txbxContent><w:p><w:r><w:t>inbox</w:t></w:r></w:p></w:txbxContent></wps:txbx>" + bodyPr + "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t>after</w:t></w:r></w:p>";

    /// <summary>A wps text box at its anchor: size (fixed / auto-grow), placement shared with pictures, insets, fill, line, rotation, vertical text.</summary>
    private static void DocxTextBoxes()
    {
        TestReport.Step("Placement: .docx text boxes");
        string body = TextBoxAnchor("<wp:wrapSquare wrapText=\"bothSides\"/>", "<wp:align>right</wp:align>",
                          "<a:xfrm rot=\"5400000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1270000\" cy=\"635000\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"FFFF00\"/></a:solidFill><a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>",
                          "<wps:bodyPr lIns=\"0\" tIns=\"12700\" rIns=\"25400\" bIns=\"38100\" vert=\"vert\"><a:spAutoFit/></wps:bodyPr>")
                    + TextBoxAnchor("<wp:wrapTopAndBottom/>", "<wp:align>left</wp:align>", "<a:noFill/><a:ln><a:noFill/></a:ln>", "<wps:bodyPr/>");
        List<string> html = DocxFidelityTest.RenderAll("placement-textboxes", body);
        TestReport.Check(html[0].Contains("<span class=\"text-box\" style=\"box-sizing:border-box;overflow-wrap:break-word;width:100pt;min-height:50pt;float:right;margin:0pt 0pt 0pt 10pt;padding:1pt 2pt 3pt 0pt;border:2pt dashed #FF0000;background:#FFFF00;transform:rotate(90deg);writing-mode:vertical-rl;\">"),
                         "square right: floats at its extent, grows with its text, insets / dashed red line / yellow fill / 90 deg / vertical (" + html[0] + ")");
        TestReport.Check(html[1].Contains("width:100pt;height:50pt;overflow:hidden;display:block;margin:0pt auto 0pt 0pt;padding:3.6pt 7.2pt 3.6pt 7.2pt;border:0;\">") && !html[1].Contains("background:"),
                         "top and bottom left: fixed height, default insets 0.1in / 0.05in, no line, no fill (" + html[1] + ")");
        TestReport.Check(html.All(h => System.Text.RegularExpressions.Regex.Matches(h, "inbox").Count == 1 && h.Contains("after")), "the box text once, the paragraph text after it");
    }

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
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = Png },
            inspect: document =>
            {
                WordDocumentPicture? read = document.ReadPicture("m:rIdP");
                TestReport.Check(read is { ContentType: "image/png" } && read.Bytes.SequenceEqual(Png), "ReadPicture(m:rIdP): the part's bytes, image/png");
                TestReport.Check(document.ReadPicture("m:rIdS") is null && document.ReadPicture("x:rIdP") is null && document.ReadPicture("m:nothing") is null,
                                 "ReadPicture: only picture relationships resolve (not the styles part, not a bad prefix)");
            });
        TestReport.Check(html[0].Contains("<img class=\"word-picture\" alt=\"\" data-word-picture=\"m:rIdP\" width=\"96\" height=\"48\" loading=\"lazy\" decoding=\"async\" style=\"width:72pt;aspect-ratio:72/36;\">"),
                         "inline: the picture by id (no bytes) at its extent, in the line, px size for layout (" + html[0] + ")");
        TestReport.Check(html.Take(6).All(IsLazyPicture), "every picture is lazy, by id, without a data: URI");
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
        TestReport.Check(PictureIds(html[3]).SequenceEqual(["p:Pictures/p.png"]) && IsLazyPicture(html[3]), "frame: the picture by its package entry id (" + html[3] + ")");
        TestReport.Check(document.ReadPicture("p:Pictures/p.png") is { ContentType: "image/png" } read && read.Bytes.SequenceEqual(Png), "ReadPicture(p:...): the package entry");
        TestReport.Check(document.ReadPicture("p:content.xml") is null && document.ReadPicture("p:Pictures/none.png") is null, "ReadPicture: no non-picture entry, nothing missing");
        TestReport.Check(html[0].Contains("margin-left:auto;margin-right:36pt;"), "table:align right: at the right margin, fo:margin-right kept (" + html[0] + ")");
        TestReport.Check(html[1].Contains("margin-left:72pt;margin-right:auto;"), "table:align left: fo:margin-left from the left");
        TestReport.Check(html[2].Contains("margin-left:auto;margin-right:auto;"), "table:align center");
        TestReport.Check(html[3].Contains("float:right;margin:7.2pt 0pt 7.2pt 18pt;"),
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
        TestReport.Check(html[0].Contains("width:144pt;margin-left:36pt;margin-right:auto;") && html[0].Contains("<td style=\"width:72pt;"),
                         "\\trleft: the row's indent from the physical left (RTF spec), the first cell from it (" + html[0] + ")");
        TestReport.Check(html[1].Contains("margin-left:auto;margin-right:auto;"), "\\trqc: a centred row (" + html[1] + ")");
        TestReport.Check(html[2].Contains("data-word-picture=\"2:0\" width=\"48\" height=\"24\" loading=\"lazy\" decoding=\"async\" style=\"width:36pt;aspect-ratio:36/18;\"")
                         && html[2].Split("class=\"word-picture").Length == 2 && IsLazyPicture(html[2]), "\\pict \\pngblip in \\shppict: drawn once, at goal x scale, by id (" + html[2] + ")");
        TestReport.Check(document.ReadPicture("2:0") is { ContentType: "image/png" } read && read.Bytes.SequenceEqual(Png), "ReadPicture(2:0): the block parsed again, the \\pict's bytes");
        TestReport.Check(document.ReadPicture("2:1") is null && document.ReadPicture("99:0") is null && document.ReadPicture("3:0") is null,
                         "ReadPicture: no second picture, no block 99, an EMF is not served");
        TestReport.Check(html[3].Contains("word-picture-missing") && html[3].Contains("width:36pt;"), "an EMF keeps its box (" + html[3] + ")");
    }

    private const string VmlNamespaces = "xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\"";

    private static string VmlPicture(string style, string wrap = "")
        => $"<w:p><w:r><w:pict {VmlNamespaces}><v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\"/>"
         + $"<v:shape id=\"s1\" type=\"#_x0000_t75\" style=\"{style}\"><v:imagedata r:id=\"rIdP\" o:title=\"\"/>{wrap}</v:shape></w:pict></w:r><w:r><w:t>text</w:t></w:r></w:p>";

    /// <summary>Old-style VML pictures (w:pict / v:shape / v:imagedata): size from the style, inline when static, wrap and side as LibreOffice's oox vml import.</summary>
    private static void DocxVmlPictures()
    {
        TestReport.Step("Placement: .docx VML pictures");
        string body = VmlPicture("width:72pt;height:36pt")
                    + VmlPicture("position:absolute;margin-left:0;margin-top:0;width:1in;height:0.5in;mso-position-horizontal:right;mso-wrap-distance-left:9pt", "<w10:wrap type=\"square\"/>")
                    + VmlPicture("position:absolute;width:96px;height:48px;mso-position-horizontal:center", "<w10:wrap type=\"topAndBottom\"/>")
                    + VmlPicture("position:absolute;margin-left:18pt;width:72pt;height:36pt")
                    + "<w:p><w:r><w:pict " + VmlNamespaces + "><v:rect style=\"width:72pt;height:36pt\"/></w:pict></w:r></w:p>";
        List<string> html = DocxFidelityTest.RenderAll("placement-vml", body,
            extraDocumentRelationships: "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = Png },
            inspect: document => TestReport.Check(document.ReadPicture("m:rIdP") is { ContentType: "image/png" }, "VML: ReadPicture by the imagedata's r:id"));
        TestReport.Check(PictureIds(html[0]).SequenceEqual(["m:rIdP"]) && html[0].Contains("style=\"width:72pt;aspect-ratio:72/36;\"") && !html[0].Contains("float"),
                         "static (no position): in the line at the style's size (" + html[0] + ")");
        TestReport.Check(html[1].Contains("width:72pt;aspect-ratio:72/36;float:right;margin:0pt 0pt 0pt 9pt;"), "absolute, wrap square, right: floats right, mso-wrap-distance-left kept (" + html[1] + ")");
        TestReport.Check(html[2].Contains("display:block;margin:0pt auto 0pt auto;") && html[2].Contains("width:72pt;"), "wrap topAndBottom, centre: a centred block; 96px = 72pt (" + html[2] + ")");
        TestReport.Check(html[3].Contains("display:block;margin:0pt auto 0pt 18pt;"), "absolute, no w10:wrap: through the text (a block without gaps) at margin-left (" + html[3] + ")");
        TestReport.Check(!html[4].Contains("word-picture") && html[4].Contains("<svg class=\"word-shape-geometry\"") && html[4].Contains("<path d=\"M0 0L72 0L72 36L0 36Z\" fill-rule=\"evenodd\" fill=\"#FFFFFF\"") && html[4].Contains("fill=\"none\" stroke=\"#000000\""),
                         "a VML shape without an image: drawn at its size, VML default white fill and black line (LibreOffice vmlformatting.cxx:706-770; the outline is the DrawingML rect preset, filled and stroked as LibreOffice draws a custom shape) (" + html[4] + ")");
    }

    /// <summary>{\shp} with a pib picture: floating at \shpleft..\shpbottom, \shpwr / \shpwrk / posh as LibreOffice reads them; the \shprslt copy not drawn.</summary>
    private static void RtfFloatingShapes()
    {
        TestReport.Step("Placement: .rtf floating shapes");
        string hex = Convert.ToHexString(Png);
        string Shape(string words, string properties)
            => @"{\shp{\*\shpinst\shpleft720\shptop0\shpright2160\shpbottom720" + words + properties
             + @"{\sp{\sn pib}{\sv {\pict\pngblip\picw2\pich1\picwgoal1440\pichgoal720 " + hex + @"}}}}{\shprslt{\pict\pngblip\picwgoal100\pichgoal100 " + hex + "}}}";
        string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Arial;}}" + "\n"
                   + @"\pard " + Shape(@"\shpwr2\shpwrk0", @"{\sp{\sn posh}{\sv 3}}{\sp{\sn dxWrapDistLeft}{\sv 114300}}") + @" beside\par" + "\n"
                   + @"\pard " + Shape(@"\shpwr1", @"{\sp{\sn posh}{\sv 2}}") + @" above\par" + "\n"
                   + @"\pard " + Shape(@"\shpwr3", "") + @" through\par" + "\n"
                   + @"\pard {\shp{\*\shpinst\shpleft0\shpright720{\sp{\sn shapeType}{\sv 1}}}} no picture\par" + "\n}";
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-shapes.rtf");
        File.WriteAllText(path, rtf);
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        TestReport.Check(html[0].Split("class=\"word-picture").Length == 2 && PictureIds(html[0]).SequenceEqual(["0:0"])
                         && html[0].Contains("width:72pt;aspect-ratio:72/36;float:right;margin:0pt 0pt 0pt 9pt;") && html[0].Contains("beside"),
                         "\\shpwr2 + posh 3: floats right at the shape's size, dxWrapDistLeft as the gap, \\shprslt not drawn (" + html[0] + ")");
        TestReport.Check(html[1].Contains("display:block;margin:0pt auto 0pt auto;"), "\\shpwr1 + posh 2: a centred block above and below (" + html[1] + ")");
        TestReport.Check(html[2].Contains("display:block;margin:0pt auto 0pt 36pt;"), "\\shpwr3, no posh: through the text, at \\shpleft (" + html[2] + ")");
        TestReport.Check(html[3].Contains("<path d=\"M0 0L36 0L36 0L0 0L0 0Z\"") && !html[3].Contains("[image]") && html[3].Contains("no picture"), "a shape without a picture is drawn as its shape type (RtfShapesFidelityTest) (" + html[3] + ")");
        TestReport.Check(document.ReadPicture("0:0") is { ContentType: "image/png" } read && read.Bytes.SequenceEqual(Png), "ReadPicture of a shape's pib picture");
        TestReport.Check(document.ReadPicture("0:1") is null, "the \\shprslt copy is not a second picture");
    }

    /// <summary>
    /// .doc: inline PICF / OfficeArt pictures (sprmCPicLocation) and row placement (sprmTJc, sprmTDxaLeft) - the .doc is the
    /// placement .docx converted by a local LibreOffice (skipped without it), plus Apache POI's PngPicture.doc / testPictures.doc.
    /// </summary>
    private static void DocBinaryPicturesAndRows()
    {
        TestReport.Step("Placement: .doc pictures and rows");
        // a quarter inch: LibreOffice writes mx = shown / natural size x 1000 into 16 bits (wrtww8gr.cxx WritePICFHeader), which
        // overflows for this 2 x 1 pixel picture shown an inch wide
        string body = $"<w:p><w:r><w:t>pic </w:t></w:r><w:r><w:drawing><wp:inline {DrawingNamespaces}><wp:extent cx=\"228600\" cy=\"114300\"/><wp:docPr id=\"1\" name=\"p\"/>{Graphic}</wp:inline></w:drawing></w:r></w:p>"
                    + Table("<w:jc w:val=\"center\"/>") + "<w:p/>" + Table("<w:tblInd w:w=\"1440\" w:type=\"dxa\"/>") + "<w:p/>" + Table("<w:jc w:val=\"right\"/>");
        DocxFidelityTest.RenderAll("placement-for-doc", body,
            extraDocumentRelationships: "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = Png });
        string? doc = LibreOfficeConversion.Convert(Path.Combine(TestReport.OutputDirectory, "fidelity-v1-placement-for-doc.docx"), ".doc", "doc:MS Word 97");
        if (doc is null) TestReport.Step("LibreOffice not installed - .doc conversion skipped");
        else
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(doc, policy);
            List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
            string picture = html.FirstOrDefault(block => block.Contains("pic")) ?? html[0];
            List<string> ids = PictureIds(picture);
            TestReport.Check(ids.Count == 1 && IsLazyPicture(picture) && picture.Contains("width:18pt;aspect-ratio:18/9;"), "inline PICF picture: by id at dxaGoal x mx (" + picture + ")");
            TestReport.Check(ids.Count == 1 && document.ReadPicture(ids[0]) is { ContentType: "image/png" } read && read.Bytes.AsSpan(0, 4).SequenceEqual(Png.AsSpan(0, 4)), "ReadPicture: the OfficeArt PNG blip's bytes (LibreOffice re-encodes them)");
            TestReport.Check(document.ReadPicture("12345678") is null && document.ReadPicture("x") is null, "ReadPicture: a wrong offset or id gives nothing");
            List<string> rows = html.Where(block => block.Contains("table-row")).ToList();
            TestReport.Check(rows.Count == 3, "three one-row tables (" + rows.Count + ")");
            if (rows.Count == 3)
            {
                TestReport.Check(rows[0].Contains("width:144pt;margin-left:auto;margin-right:auto;"), "sprmTJc 1: a centred row at its width (" + rows[0] + ")");
                TestReport.Check(System.Text.RegularExpressions.Regex.IsMatch(rows[1], "margin-left:[1-9][0-9.]*pt;margin-right:auto;"), "the indented row: from the left (" + rows[1] + ")");
                TestReport.Check(rows[2].Contains("margin-left:auto;margin-right:0;"), "sprmTJc 2: at the right (" + rows[2] + ")");
            }
        }
        // (Apache POI's PngPicture.doc holds a floating picture - a 0x08 character, FSPA + the drawing group - not read yet)
        foreach (string name in new[] { "testPictures.doc", "two_images.doc", "testCroppedPictures.doc" })
        {
            string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "apache-poi", name);
            if (!File.Exists(path)) { TestReport.Step(name + " not found - skipped"); continue; }
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(Path.GetFullPath(path));
            List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
            List<string> ids = blocks.SelectMany(PictureIds).ToList();
            int shown = ids.Count(id => document.ReadPicture(id) is { Bytes.Length: > 0 });
            TestReport.Check(ids.Count > 0 && shown == ids.Count, $"{name}: {ids.Count} picture(s) by id, {shown} read back");
        }
    }
}
