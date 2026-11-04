using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// draw:measure and draw:caption geometry against numbers worked by hand from LibreOffice's code
/// (sdrmeasureprimitive2d.cxx, svdomeas.cxx TakeRepresentation, svdmodel.cxx GetMetricString, svdocapt.cxx ImpCalcTail*).
/// </summary>
public static class OdtMeasureCaptionTest
{
    private const string Ns = "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" "
                            + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"";
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0", Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
                                       Fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
    private const double Mm = 72 / 25.4;

    public static void Run()
    {
        TestReport.Step("ODT draw:measure / draw:caption geometry");
        Measures();
        Labels();
        Captions();
    }

    private static Func<XName, string?> Props(params (XName Name, string Value)[] values)
    {
        var map = values.ToDictionary(v => v.Name, v => v.Value);
        return name => map.TryGetValue(name, out string? v) ? v : null;
    }

    private static XElement Measure(string x2, string y2, string body = "") =>
        XElement.Parse($"<draw:measure {Ns} svg:x1=\"0pt\" svg:y1=\"0pt\" svg:x2=\"{x2}\" svg:y2=\"{y2}\">{body}</draw:measure>", LoadOptions.PreserveWhitespace);

    private static bool Near(double a, double b) => Math.Abs(a - b) < 0.001;
    private static bool Line(WordShapeSubpath s, double x1, double y1, double x2, double y2) =>
        s.Segments.Count == 1 && Near(s.StartX, x1) && Near(s.StartY, y1) && Near(s.Segments[0].X, x2) && Near(s.Segments[0].Y, y2);
    private static string Show(WordShapeSubpath s) =>
        string.Join(" ", new[] { $"{s.StartX:0.###},{s.StartY:0.###}" }.Concat(s.Segments.Select(g => $"{g.X:0.###},{g.Y:0.###}")));

    private static void Measures()
    {
        // 1. Defaults, 2in horizontal, 6x6 markers, text 30x10. Pool defaults: line distance 8 mm (22.677pt), overhang 2 mm, guide distance 1 mm.
        //    Arrows fit (6+6+3 < 144; small need 12+1.5+30 < 144) -> one main line with both markers at y = -8 mm.
        //    Help lines x=0 / 144 from leftDelta-lower = -1 mm to -upper-distance = -10 mm. Text above: Y = -8mm - (10 + 12*0.1) -> centre y -8mm-6.2.
        var m = OdtMeasureShape.Layout(Measure("2in", "0in"), Props(), 30, 10, new OdtMeasureArrow(6, 6), new OdtMeasureArrow(6, 6));
        TestReport.Check(m.MainLines.Count == 1 && Line(m.MainLines[0].Line, 0, -8 * Mm, 144, -8 * Mm) && m.MainLines[0].StartMarker && m.MainLines[0].EndMarker,
            "measure defaults: one dimension line 8 mm above, markers at both ends (" + string.Join(" | ", m.MainLines.Select(p => Show(p.Line))) + ")");
        TestReport.Check(m.HelpLines.Count == 2 && Line(m.HelpLines[0], 0, -1 * Mm, 0, -10 * Mm) && Line(m.HelpLines[1], 144, -1 * Mm, 144, -10 * Mm),
            "measure defaults: extension lines from 1 mm to 10 mm above (" + string.Join(" | ", m.HelpLines.Select(Show)) + ")");
        TestReport.Check(Near(m.TextCenterX, 72) && Near(m.TextCenterY, -8 * Mm - 6.2) && Near(m.TextRotationDegrees, 0),
            $"measure defaults: text centred above the line ({m.TextCenterX:0.###},{m.TextCenterY:0.###} rot {m.TextRotationDegrees})");

        // 2. line-distance 1cm, guide-overhang 0.5cm, guide-distance 0.2cm, start-guide 0.1cm, placing below, 2pt stroke, 10x10 markers, text 40x12.
        //    Below: line at +1cm; help left from 0.2-0.1 = 0.1cm to 1.5cm, right from 0.2cm; text below the line (above swapped):
        //    Y = 1cm + (40/2)*0.1 + 1 -> centre 1cm + 3 + 6.
        var p2 = Props((Draw + "line-distance", "1cm"), (Draw + "guide-overhang", "0.5cm"), (Draw + "guide-distance", "0.2cm"), (Draw + "start-guide", "0.1cm"),
                       (Draw + "placing", "below"), (Svg + "stroke-width", "2pt"));
        m = OdtMeasureShape.Layout(Measure("144pt", "0pt"), p2, 40, 12, new OdtMeasureArrow(10, 10), new OdtMeasureArrow(10, 10));
        TestReport.Check(m.MainLines.Count == 1 && Line(m.MainLines[0].Line, 0, 10 * Mm, 144, 10 * Mm), "measure below: dimension line 1 cm below (" + Show(m.MainLines[0].Line) + ")");
        TestReport.Check(Line(m.HelpLines[0], 0, 1 * Mm, 0, 15 * Mm) && Line(m.HelpLines[1], 144, 2 * Mm, 144, 15 * Mm),
            "measure below: start guide shortens the gap, overhang 0.5 cm past the line (" + string.Join(" | ", m.HelpLines.Select(Show)) + ")");
        TestReport.Check(Near(m.TextCenterX, 72) && Near(m.TextCenterY, 10 * Mm + 9), $"measure below: text under the line ({m.TextCenterX:0.###},{m.TextCenterY:0.###})");

        // 3. Vertical (0,0)->(0,100): angle 90° > 45° -> auto upside down, text side swapped; frame (x,y) -> anchor (-y, x).
        //    Line at frame y=-8mm -> anchor x = 8mm; text centre frame (50, -8mm+5) -> anchor (8mm-5, 50); rotation 90+180.
        m = OdtMeasureShape.Layout(Measure("0pt", "100pt"), Props(), 20, 10, default, default);
        TestReport.Check(Line(m.MainLines[0].Line, 8 * Mm, 0, 8 * Mm, 100) && !m.MainLines[0].StartMarker, "measure vertical: line 8 mm to the right (" + Show(m.MainLines[0].Line) + ")");
        TestReport.Check(Near(m.TextCenterX, 8 * Mm - 5) && Near(m.TextCenterY, 50) && Near(m.TextRotationDegrees, 270),
            $"measure vertical: text turned to read upright ({m.TextCenterX:0.###},{m.TextCenterY:0.###} rot {m.TextRotationDegrees})");

        // 4. measure-align left-outside: arrows outside, three pieces; left piece = startH + text width.
        m = OdtMeasureShape.Layout(Measure("144pt", "0pt"), Props((Draw + "measure-align", "left-outside")), 30, 10, new OdtMeasureArrow(6, 6), new OdtMeasureArrow(6, 6));
        TestReport.Check(m.MainLines.Count == 3 && Line(m.MainLines[0].Line, -36, -8 * Mm, 0, -8 * Mm) && m.MainLines[0].EndMarker && !m.MainLines[0].StartMarker
                         && Line(m.MainLines[1].Line, 144, -8 * Mm, 156, -8 * Mm) && m.MainLines[1].StartMarker && !m.MainLines[2].StartMarker && !m.MainLines[2].EndMarker
                         && Near(m.TextCenterX, -(6 + 30 + 12 * 0.18) + 15),
            "measure left-outside: arrows outside, text left of the start (" + string.Join(" | ", m.MainLines.Select(p => Show(p.Line))) + $" text {m.TextCenterX:0.###})");
    }

    private static void Labels()
    {
        const string template = "<text:p><text:measure text:kind=\"gap\"/><text:measure text:kind=\"value\">x</text:measure> <text:measure text:kind=\"unit\"/><text:measure text:kind=\"gap\"/></text:p>";
        // 2in = 5080/100 mm = 2880 twips; 2880 * 127/72 = 5080 -> "50.80" -> auto decimals trim -> "50.8"; show-unit off: unit field empty.
        Label(Measure("2in", "0in", template), Props(), "50.8 ", "default: millimetres, trailing zeros trimmed, no unit");
        Label(Measure("2in", "0in", template), Props((Draw + "show-unit", "true")), "50.8 mm", "show-unit: Writer's default unit mm");
        // 10cm = 10000/100 mm -> round(10000*1440/2540) = 5669 twips -> 5669*127/72 = 9999.4 -> "99.99" (Writer's twip model).
        Label(Measure("10cm", "0cm", template), Props(), "99.99 ", "10 cm through Writer's twips");
        // cm, 1 place: 2880*0.1 -> 288 (truncated); 288*127/72 = 508 (mark 2) /10 -> 50.8 -> 51 -> "5.1".
        Label(Measure("2in", "0in", template), Props((Draw + "unit", "cm"), (Draw + "decimal-places", "1"), (Draw + "show-unit", "true")), "5.1 cm", "cm with one decimal place");
        // inch, 3 places: 2880*5/127 = 113.39 -> 113; 113*127/72 = 199.32 *10 -> 1993 -> "1.993".
        Label(Measure("2in", "0in", template), Props((Draw + "unit", "inch"), (Draw + "decimal-places", "3"), (Draw + "show-unit", "true")), "1.993 \"", "inches truncate in the twip model");
        Label(Measure("2in", "0in", template), Props((Draw + "parallel", "true")), " 50.8  ", "draw:parallel=true (TextRotate90): gap fields are spaces");
        Label(Measure("2in", "0in"), Props(), "", "no text:p: no text (LibreOffice drops its pre-made fields)");
        // 20m = 2,000,000/100 mm -> 1133858 twips -> 1133858*127/72 = 2000000.0 -> "20,000" mm.
        Label(Measure("2000cm", "0cm", template), Props(), "20,000 ", "thousands separator");
    }

    private static void Label(XElement measure, Func<XName, string?> get, string expected, string what)
    {
        string actual = OdtMeasureShape.Label(measure, get);
        TestReport.Check(actual == expected, $"measure text, {what}: \"{actual}\"");
    }

    private static XElement Caption(string x, string pointX, string pointY, string extra = "") =>
        XElement.Parse($"<draw:caption {Ns} svg:x=\"{x}\" svg:y=\"0pt\" svg:width=\"100pt\" svg:height=\"50pt\" draw:caption-point-x=\"{pointX}\" draw:caption-point-y=\"{pointY}\" {extra}/>");

    private static void Tail(OdtCaptionGeometry g, string what, params double[] xy)
    {
        var points = new List<double> { g.Tail.StartX, g.Tail.StartY };
        foreach (var s in g.Tail.Segments) { points.Add(s.X); points.Add(s.Y); }
        bool ok = points.Count == xy.Length && points.Zip(xy, Near).All(b => b);
        TestReport.Check(ok, $"caption {what}: tail {string.Join(" ", points.Select(v => v.ToString("0.###", CultureInfo.InvariantCulture)))}");
    }

    private static void Captions()
    {
        // Defaults: angled-connector-line, horizontal, escape 50%, fit line length. Rect x 10..110, y 0..50; point (150,120) rel -> (160,120).
        // Left (10,25) vs right (110,25): 160-10 < 110-160 false -> right; middle x = (160+110)/2.
        var g = OdtCaptionShape.Layout(Caption("10pt", "150pt", "120pt"), Props());
        Tail(g, "angled connector (default)", 160, 120, 135, 25, 110, 25);
        TestReport.Check(Near(g.Left, 10) && Near(g.Width, 100) && g.CornerRadiusX == 0, "caption rectangle");

        // angled-line, vertical, gap 5pt: top (50,-5) vs bottom (50,55): 10+5 < 55-10 -> top.
        g = OdtCaptionShape.Layout(Caption("0pt", "-40pt", "10pt"), Props((Draw + "caption-type", "angled-line"), (Draw + "caption-escape-direction", "vertical"), (Draw + "caption-gap", "5pt")));
        Tail(g, "angled line", -40, 10, 50, -5);

        // straight-line, horizontal (tries vertical for this type), escape 20pt absolute: top (20,0) / bottom (20,50): 90 < -40 false -> bottom;
        // the tail point slides to y = 50.
        g = OdtCaptionShape.Layout(Caption("0pt", "30pt", "90pt"), Props((Draw + "caption-type", "straight-line"), (Draw + "caption-escape", "20pt")));
        Tail(g, "straight line", 30, 50, 20, 50);

        // angled-connector, vertical, fixed line length 10pt: top (50,0); middle (50,-10).
        g = OdtCaptionShape.Layout(Caption("0pt", "60pt", "-30pt"), Props((Draw + "caption-escape-direction", "vertical"), (Draw + "caption-fit-line-length", "false"), (Draw + "caption-line-length", "10pt")));
        Tail(g, "connector with line length", 60, -30, 50, -10, 50, 0);

        // auto: horizontal pick right (100,25), vertical pick top (50,0); 60²+40² = 5200 < 10²+65² = 4325 false -> keep right.
        g = OdtCaptionShape.Layout(Caption("0pt", "110pt", "-40pt", "draw:corner-radius=\"40pt\""), Props((Draw + "caption-escape-direction", "auto")));
        Tail(g, "best fit", 110, -40, 105, 25, 100, 25);
        TestReport.Check(Near(g.CornerRadiusX, 40) && Near(g.CornerRadiusY, 25), $"caption corner radius clamped to half the height ({g.CornerRadiusX}, {g.CornerRadiusY})");
    }
}
