using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .docx drawing details (--fidelity, seconds), each expected value worked out by hand from the LibreOffice source it cites:
/// dash patterns (preset, round cap, custom), every line-end type with its sizes, alpha on fills / lines / gradient stops,
/// arcTo on an ellipse (angles on the ellipse, a negative sweep), a path without w / h, every guide operator, presets with
/// darken / lighten paths, theme fill styles (fillStyleLst, bgFillStyleLst, past the end), radial gradients, a:grpFill, and VML
/// dash styles and arrows. Hand-written XML in the shape Word writes.
/// </summary>
public static class DocxDrawingFidelityTest
{
    private const string Namespaces = "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        + "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" "
        + "xmlns:v=\"urn:schemas-microsoft-com:vml\"";

    private const string Theme = """
        <?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements>
        <a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
        <a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
        <a:fmtScheme name="T">
        <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:alpha val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>
        <a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:fillStyleLst>
        <a:lnStyleLst><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></a:lnStyleLst>
        <a:bgFillStyleLst><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></a:bgFillStyleLst>
        </a:fmtScheme></a:themeElements></a:theme>
        """;

    public static void Run()
    {
        TestReport.Step("Drawing details: dashes, line ends, alpha, arcs, guides, presets, theme fills, VML strokes");
        var body = new List<string>
        {
            /* 0 */ Line("<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:prstDash val=\"sysDashDot\"/></a:ln>"),
            /* 1 */ Line("<a:ln w=\"12700\" cap=\"rnd\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>"),
            /* 2 */ Line("<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:custDash><a:ds d=\"300000\" sp=\"100000\"/><a:ds d=\"100%\" sp=\"100%\"/></a:custDash></a:ln>"),
            /* 3 */ Line("<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:headEnd type=\"oval\" w=\"sm\" len=\"sm\"/><a:tailEnd type=\"triangle\"/></a:ln>"),
            /* 4 */ Line("<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:tailEnd type=\"stealth\" w=\"lg\" len=\"lg\"/></a:ln>"),
            /* 5 */ Line("<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill><a:headEnd type=\"diamond\"/><a:tailEnd type=\"arrow\"/></a:ln>"),
            /* 6 */ Shape("rect", "<a:solidFill><a:srgbClr val=\"FF0000\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill>"
                              + "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"0000FF\"><a:alpha val=\"80000\"/><a:alphaMod val=\"50000\"/></a:srgbClr></a:solidFill></a:ln>"),
            /* 7 */ Shape("rect", "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FF0000\"><a:alpha val=\"30000\"/><a:alphaOff val=\"20000\"/></a:srgbClr></a:gs>"
                              + "<a:gs pos=\"100000\"><a:srgbClr val=\"00FF00\"/></a:gs></a:gsLst><a:lin ang=\"0\"/></a:gradFill>"),
            /* 8 */ Custom("<a:path w=\"200\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"50\"/></a:moveTo><a:arcTo wR=\"100\" hR=\"50\" stAng=\"10800000\" swAng=\"10800000\"/></a:path>"
                          + "<a:path w=\"200\" h=\"100\"><a:moveTo><a:pt x=\"200\" y=\"50\"/></a:moveTo><a:arcTo wR=\"100\" hR=\"50\" stAng=\"0\" swAng=\"2700000\"/></a:path>"
                          + "<a:path w=\"200\" h=\"100\"><a:moveTo><a:pt x=\"200\" y=\"50\"/></a:moveTo><a:arcTo wR=\"100\" hR=\"50\" stAng=\"0\" swAng=\"-5400000\"/></a:path>"
                          + "<a:path><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"914400\" y=\"457200\"/></a:lnTo></a:path>"),
            /* 9 */ Custom("<a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"cb\" y=\"c\"/></a:lnTo><a:lnTo><a:pt x=\"f\" y=\"e\"/></a:lnTo>"
                          + "<a:lnTo><a:pt x=\"g\" y=\"hh\"/></a:lnTo><a:lnTo><a:pt x=\"t1\" y=\"t2\"/></a:lnTo><a:lnTo><a:pt x=\"md\" y=\"pd\"/></a:lnTo><a:lnTo><a:pt x=\"v\" y=\"v\"/></a:lnTo></a:path>",
                          "<a:gd name=\"a\" fmla=\"at2 1 1\"/><a:gd name=\"cb\" fmla=\"cos 100 a\"/><a:gd name=\"c\" fmla=\"mod 30 40 0\"/><a:gd name=\"d\" fmla=\"pin 0 150 100\"/>"
                          + "<a:gd name=\"e\" fmla=\"?: -1 0 d\"/><a:gd name=\"f\" fmla=\"+/ cb 0 2\"/><a:gd name=\"g\" fmla=\"cat2 100 1 1\"/><a:gd name=\"hh\" fmla=\"sat2 100 1 1\"/>"
                          + "<a:gd name=\"tv\" fmla=\"tan 10 a\"/><a:gd name=\"sq\" fmla=\"sqrt 2500\"/><a:gd name=\"av\" fmla=\"abs -40\"/><a:gd name=\"mx\" fmla=\"max 3 7\"/><a:gd name=\"mn\" fmla=\"min 3 7\"/>"
                          + "<a:gd name=\"t1\" fmla=\"+- tv sq 0\"/><a:gd name=\"t2\" fmla=\"+- av mx mn\"/><a:gd name=\"md\" fmla=\"*/ 10 30 3\"/><a:gd name=\"pd\" fmla=\"+- 10 30 3\"/><a:gd name=\"v\" fmla=\"val 12\"/>"),
            /* 10 */ Shape("cube", "<a:solidFill><a:srgbClr val=\"808080\"/></a:solidFill>"),
            /* 11 */ Shape("rect", "", style: "<wps:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent2\"/></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"/></wps:style>"),
            /* 12 */ Shape("rect", "", style: "<wps:style><a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"2\"><a:schemeClr val=\"accent6\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"/></wps:style>"),
            /* 13 */ Shape("rect", "", style: "<wps:style><a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"5\"><a:schemeClr val=\"accent4\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"/></wps:style>"),
            /* 14 */ Shape("rect", "", style: "<wps:style><a:lnRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"1001\"><a:schemeClr val=\"accent4\"/></a:fillRef><a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"/></wps:style>"),
            /* 15 */ GroupFill(),
            /* 16 */ "<w:p><w:r><w:pict " + Namespaces + "><v:line from=\"0,0\" to=\"72pt,0\" strokeweight=\"1pt\" strokecolor=\"#00FF00\"><v:stroke dashstyle=\"longdashdot\" endarrow=\"block\" endarrowwidth=\"wide\" endarrowlength=\"short\"/></v:line></w:pict></w:r></w:p>",
            /* 17 */ "<w:p><w:r><w:pict " + Namespaces + "><v:rect style=\"width:72pt;height:36pt\" strokeweight=\"2pt\"><v:stroke dashstyle=\"3 1 1 1\" endcap=\"round\"/></v:rect></w:pict></w:r></w:p>",
            /* 18 */ Shape("rect", "<a:solidFill><a:prstClr val=\"dkGray\"/></a:solidFill>", rotation: " rot=\"5400000\"", extraFill: ""),
            /* 19 */ Shape("rect", "<a:gradFill rotWithShape=\"0\"><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"000000\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"FFFFFF\"/></a:gs></a:gsLst><a:lin ang=\"0\"/></a:gradFill>", rotation: " rot=\"5400000\""),
        };
        List<string> html = DocxFidelityTest.RenderAll("drawing-details", string.Concat(body),
            extraDocumentRelationships: "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>",
            binaryParts: new Dictionary<string, byte[]> { ["word/theme/theme1.xml"] = new UTF8Encoding(false).GetBytes(Theme) });

        // dashes: lclConvertPresetDash lineproperties.cxx:60-84 (dots / dashes followed by Distance, in line widths); round cap: 0.99 off the dash, onto the gap (:451-462)
        TestReport.Check(html[0].Contains("stroke-dasharray=\"3 1 1 1\""), "sysDashDot at 1pt: a 3-width dash, gap 1, a 1-width dot, gap 1 (" + html[0] + ")");
        TestReport.Check(html[1].Contains("stroke-linecap=\"round\"") && html[1].Contains("stroke-dasharray=\"3.01 3.99\""), "dash with a round cap: 4 - 0.99 and 3 + 0.99 widths, the round cap drawn (" + html[1] + ")");
        TestReport.Check(html[2].Contains("stroke-dasharray=\"6 2 2 2\""), "custDash at 2pt: ds 300000 / 100000 (1000ths of a percent) and 100% / 100% (" + html[2] + ")");

        // line ends: lclPushMarkerProperties lineproperties.cxx:223-367; placement b2dlinegeometry.cxx:38-178; shaft polygonprimitive2d.cxx:704-770.
        // a 2pt line, oval sm: 2 x 2pt wide and long, centred on the start (consumed 2, overlap 4 / 15); triangle med: 6 wide and long, tip on the end
        TestReport.Check(html[3].Contains("<path d=\"M1.73 0L66.4 0\"") && html[3].Contains("<path d=\"M72 0L66 3L66 -3L72 0Z\"") && html[3].Contains("<path d=\"M-2 0L"),
                         "oval sm head centred on the start, triangle tail with its tip on the end, the shaft cut back (" + html[3] + ")");
        // stealth lg at 1pt: 5 x 1.98pt (the 70 / 100 mm minimum) wide and long; its notch at 60 % of the length: 72 - 0.6 x 9.92 = 66.05
        TestReport.Check(html[4].Contains("L66.05 0L"), "stealth lg tail: the notch 60 % back from the tip (" + html[4] + ")");
        // diamond med at 2pt: 6 x 6 centred on the start; arrow (open) med: 3.5 x 2pt = 7 wide, a stem of the line's width
        TestReport.Check(html[5].Contains("<path d=\"M-3 0L0 -3L3 0L0 3L-3 0Z\"") && Regex.IsMatch(html[5], "<path d=\"M72 0(L[0-9.-]+ [0-9.-]+){9}Z\""),
                         "diamond head straddles the start; open arrow tail has its 10-point outline (" + html[5] + ")");

        // alpha (color.cxx:467-477, lclSetValue / lclModValue / lclOffValue :193-210)
        TestReport.Check(html[6].Contains("fill=\"#FF0000\" fill-opacity=\"0.5\"") && html[6].Contains("stroke-opacity=\"0.4\""), "alpha 50 % fill; alpha 80 % x alphaMod 50 % line (" + html[6] + ")");
        TestReport.Check(html[7].Contains("stop-color=\"#FF0000\" stop-opacity=\"0.5\"") && html[7].Contains("<stop offset=\"1\" stop-color=\"#00FF00\"/>"), "gradient stop alpha 30 % + alphaOff 20 % (" + html[7] + ")");

        // arcTo (EnhancedCustomShape2d.cxx:2414-2496): w 200 h 100 scaled to 72 x 36; angles on the ellipse; a negative sweep runs backwards
        TestReport.Check(Regex.IsMatch(html[8], "<path d=\"M0 18(C[^\"]*) 72 18\" fill=\"none\""), "arcTo half ellipse from 180 deg sweeping 180: ends at the right (" + html[8] + ")");
        TestReport.Check(html[8].Contains(" 52.1 34.1\""), "arcTo 45 deg on a 2:1 ellipse: the point on the 45-degree ray, (100 + 44.72, 50 + 44.72) x 0.36 (" + html[8] + ")");
        TestReport.Check(Regex.IsMatch(html[8], "<path d=\"M72 18(C[^\"]*) 36 0\" fill=\"none\""), "arcTo sweeping -90 deg: up to the top (" + html[8] + ")");
        TestReport.Check(html[8].Contains("<path d=\"M0 0L72 36\" fill=\"none\""), "a path without w / h: coordinates in the shape's EMU (customshapeproperties.cxx:236-244) (" + html[8] + ")");

        // every guide operator (customshapegeometry.cxx:486-606), path w = h = 100 on 72 x 36; guide names avoid the built-ins (l t r b w h ..., looked up first :180-372)
        TestReport.Check(html[9].Contains("L50.91 18L25.46 36L50.91 25.46L43.2 15.84L72 13.32L8.64 4.32"),
                         "guides: cos / at2 / mod, +/ / ?: / pin, cat2 / sat2, tan / sqrt / abs / max / min / +-, */ / +-, val (" + html[9] + ")");

        // cube: darkenLess / lightenLess paths (EnhancedCustomShape2d.cxx:2122-2133, GetColorData :1095-1118): 0x80 x 0.8 = 0x66, 0x80 x 0.8 + 51 = 0x99
        TestReport.Check(html[10].Contains("fill=\"#666666\"") && html[10].Contains("fill=\"#999999\"") && html[10].Contains("fill=\"#808080\""), "cube: its darker and lighter faces (" + html[10] + ")");

        // theme styles (theme.cxx:41-55): fillRef 1 = solid phClr, 2 = the gradient with phClr stops, past the end = the last, 1001 = bgFillStyleLst 1
        TestReport.Check(html[11].Contains("fill=\"#4472C4\" stroke=\"none\"") && html[11].Contains("stroke=\"#ED7D31\" stroke-width=\"1\" stroke-dasharray=\"1 1\""),
                         "fillRef 1: the solid style in accent1; lnRef 1: the theme a:ln (1pt sysDot) in accent2 (" + html[11] + ")");
        TestReport.Check(html[12].Contains("stop-color=\"#70AD47\"/>") && html[12].Contains("stop-color=\"#70AD47\" stop-opacity=\"0.5\""), "fillRef 2: the gradient style, phClr = accent6 (" + html[12] + ")");
        // a:path circle, fillToRect at the centre: a circle of the diagonal (hypot(72, 36) / 2 = 40.25) about (36, 18) (gradienttools.cxx:127-190)
        TestReport.Check(html[13].Contains("<radialGradient id=") && html[13].Contains("gradientUnits=\"userSpaceOnUse\" cx=\"36\" cy=\"18\" r=\"40.25\">") && html[13].Contains("<stop offset=\"1\" stop-color=\"#FFC000\"/>"),
                         "fillRef 5 (past the end): the last style, a radial gradient, phClr = accent4 (" + html[13] + ")");
        TestReport.Check(html[14].Contains("fill=\"#ABCDEF\""), "fillRef 1001: bgFillStyleLst 1 (" + html[14] + ")");
        TestReport.Check(html[15].Contains("fill=\"#112233\""), "a:grpFill: the group's fill (shape.cxx:3143-3146) (" + html[15] + ")");

        // VML strokes (vmlformatting.cxx:565-664): longdashdot = lgDashDot; block = triangle, wide = lg width, short = sm length
        TestReport.Check(html[16].Contains("stroke=\"#00FF00\" stroke-width=\"1\" stroke-dasharray=\"8 3 1 3\"") && html[16].Contains("<path d=\"M72 0L68.03 4.96L68.03 -4.96L72 0Z\""),
                         "VML line: longdashdot, a wide short block arrow: 5 x 1.98pt wide, 2 x 1.98pt long (" + html[16] + ")");
        TestReport.Check(html[17].Contains("stroke-linecap=\"round\"") && html[17].Contains("stroke-dasharray=\"4.02 3.98 0.02 3.98\""),
                         "VML custom dash 3 1 1 1 in line widths with round caps (the cap inside the dash) (" + html[17] + ")");
        TestReport.Check(html[18].Contains("fill=\"#A9A9A9\""), "prstClr dkGray: LibreOffice's A9A9A9 (color.cxx constDmlColors), not CSS darkgray (" + html[18] + ")");
        TestReport.Check(html[19].Contains("x1=\"36\" y1=\"36\" x2=\"36\" y2=\"0\"") && html[19].Contains("rotate(90)"),
                         "rotWithShape=0 on a shape turned 90 deg: the gradient stays left-to-right on the page (fillproperties.cxx:521-525) (" + html[19] + ")");
    }

    private static string Line(string line)
        => Drawing("<a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>" + line, 914400, 0);

    private static string Shape(string preset, string fill, string style = "", string rotation = "", string extraFill = "")
        => Drawing($"<a:prstGeom prst=\"{preset}\"><a:avLst/></a:prstGeom>{fill}{extraFill}", 914400, 457200, style, rotation);

    private static string Custom(string paths, string guides = "")
        => Drawing($"<a:custGeom><a:avLst/><a:gdLst>{guides}</a:gdLst><a:pathLst>{paths}</a:pathLst></a:custGeom><a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>", 914400, 457200);

    private static string Drawing(string geometryAndFill, long cx, long cy, string style = "", string rotation = "")
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"{cx}\" cy=\"{cy}\"/><wp:docPr id=\"1\" name=\"s\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">"
         + $"<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm{rotation}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>{geometryAndFill}</wps:spPr>{style}<wps:bodyPr/>"
         + "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static string GroupFill()
        => $"<w:p><w:r><w:drawing><wp:inline {Namespaces}><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"2\" name=\"g\"/><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\">"
         + "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"914400\" cy=\"457200\"/></a:xfrm>"
         + "<a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></wpg:grpSpPr>"
         + "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:grpFill/></wps:spPr><wps:bodyPr/></wps:wsp>"
         + "</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
}
