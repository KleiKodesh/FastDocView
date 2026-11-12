using System.Globalization;
using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Picture cropping in every format, against Word 16's PDF of test-data/word-picture-crop/picture-crop-v1 (make-crop-fixture.ps1:
/// quadrants.png, 200 x 100 px = 150 x 75 pt, inline cropped left 15 / top 7.5 / right 30 / bottom 15 pt, inline with a negative
/// crop left -15 / top -7.5 pt, floating cropped left 30 / bottom 15 pt; .docx a:srcRect, .doc OfficeArt cropFrom*, .rtf
/// \piccrop*) and LibreOffice's .odt of it (make-crop-odt.ps1: fo:clip). Word's PDF (picture-crop-v1.pdf content stream):
/// the clip boxes 104.93 x 52.45, 164.9 x 82.4 and 119.95 x 59.95 pt; the negative crop draws the whole picture 149.91 pt wide at
/// +14.99 / +7.49 pt inside its box; the cut ones are baked into 140 x 70 and 160 x 80 px images (the kept part: left 10 % / top
/// 10 % and left 20 %), i.e. the whole picture 149.9 pt wide at -14.99 / -7.49 and -29.99 / 0 pt. Measured here: the box width,
/// the whole picture's width (box x the inner width %) and its offset (margins in % of the box width).
/// </summary>
public static class PictureCropFidelityTest
{
    private static readonly (double Box, double Whole, double Left, double Top)[] Expected =
    [
        (104.93, 149.9, -14.99, -7.49),
        (164.9, 149.91, 14.99, 7.49),
        (119.95, 149.94, -29.99, 0),
    ];

    public static void Run()
    {
        TestReport.Step("Picture crop: .docx / .doc / .rtf / .odt against Word's PDF");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-picture-crop");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf", ".odt" })
        {
            string path = Path.Combine(folder, "picture-crop-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-picture-crop/picture-crop-v1{extension} present"); continue; }
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
            var boxes = Crops(html);
            TestReport.Check(boxes.Count == 3, $"{extension}: three cropped pictures ({boxes.Count})");
            for (int i = 0; i < Math.Min(3, boxes.Count); i++)
            {
                var (box, whole, left, top) = boxes[i];
                var expected = Expected[i];
                // Word's RTF keeps \picscalex / \picscaley as whole percents (the padded picture: 99 for 99.94), so its sizes may be 1 % off
                double tolerance = extension == ".rtf" ? 0.011 * expected.Box + 0.3 : 0.3;
                TestReport.Check(Math.Abs(box - expected.Box) < tolerance && Math.Abs(whole - expected.Whole) < tolerance && Math.Abs(left - expected.Left) < tolerance && Math.Abs(top - expected.Top) < tolerance,
                                 $"{extension} picture {i + 1}: box {box:0.##} pt, whole picture {whole:0.##} pt at {left:0.##} / {top:0.##} pt (Word: {expected.Box} / {expected.Whole} at {expected.Left} / {expected.Top})");
            }
            if (extension != ".odt")
                TestReport.Check(document.ReadPicture(Regex.Match(html, "data-word-picture=\"([^\"]*)\"").Groups[1].Value) is { } picture
                                 && FastDocView.Html.WordPictureSize.Read(picture.Bytes) is { } size && Math.Abs(size.WidthPt - 150) < 0.5,
                                 $"{extension}: the served picture is the whole uncropped one (the box crops it)");
        }
        VmlCrop();
    }

    /// <summary>VML v:imagedata croptop / cropleft ("Nf" = N / 65536, or a plain fraction; vmlshape.cxx:118-128) on a 72 x 36 pt shape.</summary>
    private static void VmlCrop()
    {
        string body = "<w:p><w:r><w:pict xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\"><v:shape id=\"s1\" type=\"#_x0000_t75\" style=\"width:72pt;height:36pt\">"
                    + "<v:imagedata r:id=\"rIdP\" croptop=\"13107f\" cropleft=\"0.25\" o:title=\"\"/></v:shape></w:pict></w:r></w:p>";
        List<string> html = DocxFidelityTest.RenderAll("picture-crop-vml", body,
            extraDocumentRelationships: "<Relationship Id=\"rIdP\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/p.png\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/media/p.png"] = PagePlacementFidelityTest.Png });
        var boxes = Crops(html[0]);
        // whole = 72 / 0.75 = 96 pt wide, 36 / 0.8 = 45 pt high: left -24 pt, top -9 pt
        TestReport.Check(boxes.Count == 1 && Math.Abs(boxes[0].Whole - 96) < 0.1 && Math.Abs(boxes[0].Left + 24) < 0.1 && Math.Abs(boxes[0].Top + 9) < 0.1,
                         ".docx VML: croptop 13107f (0.2) and cropleft 0.25 crop the picture (" + html[0] + ")");
    }

    private static List<(double Box, double Whole, double Left, double Top)> Crops(string html)
    {
        var found = new List<(double, double, double, double)>();
        foreach (Match match in Regex.Matches(html, "class=\"word-picture word-picture-crop\" style=\"[^\"]*?width:([0-9.]+)pt;[^\"]*\"><img [^>]*style=\"[^\"]*;width:([-0-9.]+)%;[^\"]*margin:([-0-9.]+)% 0 0 ([-0-9.]+)%\""))
        {
            double Number(int group) => double.Parse(match.Groups[group].Value, CultureInfo.InvariantCulture);
            double box = Number(1);
            found.Add((box, box * Number(2) / 100, box * Number(4) / 100, box * Number(3) / 100));
        }
        return found;
    }
}
