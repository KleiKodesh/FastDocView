using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// ODF fills that are more than one colour, as SVG paints in the shape's own points: gradients of every draw:style, hatches, and
/// transparency gradients (draw:opacity), the way LibreOffice draws them (master read 2026-11-04):
/// - colour stops: xmloff GradientStyle.cxx:60-145 (draw:start-color / end-color, start- / end-intensity blended toward black
///   per stop offset - basegfx bgradient.cxx:249-272 blendToIntensity; LibreOffice's loext:gradient-stop children :147-215
///   replace the two colours), angle :112-131 (sax Converter::convert10thDegAngle converter.cxx:878-912: "deg" / "grad" /
///   "rad" units; a plain number is tenths of a degree before ODF 1.2 and in ODF 1.2 from OpenOffice.org / AOO or LibreOffice
///   before 7, degrees otherwise - decided per document from meta:generator, <see cref="OdtGeneratorVersion.GradientAngleIsTenths"/>); draw:cx / draw:cy default 0 (Gradient2 XOffset / YOffset = 0, :70-71); draw:border;
/// - geometry: basegfx/source/tools/gradienttools.cxx init1DGradientInfo :52-121 (linear, axial: the box grown to hold the
///   rotated gradient, rotated about its centre by -angle, the border at the start), initEllipticalGradientInfo :127-190
///   (radial: a circle of the box's diagonal; ellipsoid: the box times sqrt 2, rotated), initRectGradientInfo :196-270
///   (square: the larger side first; rectangular), offsets after rotation; the colour at each point:
///   getLinearGradientAlpha :589-608, getAxialGradientAlpha :611-627 with the reversed stops of GeoTexSvxGradientAxial
///   (drawinglayer texture.cxx:249-383: the start colour in the middle), getRadialGradientAlpha :629-639,
///   getEllipticalGradientAlpha :641-667 (the aspect correction), getSquareGradientAlpha / getRectangularGradientAlpha :669-757;
///   outside the unit area the start colour (alpha 0);
/// - steps: BColorStops::getInterpolatedBColor (bgradient.cxx:85-190) with calculateNumberOfSteps (gradienttools.cxx:479-497):
///   draw:gradient-step-count bands of interpolate(start, end, k / (n - 1)) per stop range, one band of the start colour when
///   n = 1; 0 (automatic) = as many as the colour distance allows (distance * 127.5 rounded, at least 1) - banded as the
///   primitive decomposition draws every style (drawinglayer texture.cxx:169-222 GeoTexSvxGradientLinear and the same loop
///   for axial :316-357, radial :425-461, elliptical :532-571, square :638-674, rect :745-784), except a plain two-colour
///   linear / axial / radial fill without transparency, which Windows LibreOffice draws smooth through Skia
///   (<see cref="IsDrawnSmoothBySkia"/>) - transparency gradients alike unless they share the fill's definition;
/// - hatches: drawinglayer fillhatchprimitive2d.cxx:36-130 and texture.cxx:810-930 GeoTexSvxHatch (lines along the box grown
///   for the rotation, every draw:distance from its top, rotated by -draw:rotation about the centre; double adds the set at
///   rotation - 90 degrees, triple also rotation - 45; hairlines; draw:fill-hatch-solid fills the background first);
/// - transparency gradients: xmloff TransGradientStyle.cxx:60-165 (draw:start / draw:end opacity percent, the same geometry),
///   drawn as an SVG luminance mask (opacity = luminance).
/// Everything is one SVG pattern tile covering the box and far beyond it, so pads continue where an outline leaves its box.
/// </summary>
internal static class OdtFillPaint
{
    /// <summary>A colour stop: offset 0..1 and RGB 0..1.</summary>
    internal readonly record struct Stop(double Offset, double R, double G, double B);

    internal sealed class GradientModel
    {
        public string Style = "linear";
        public List<Stop> Stops = new();
        public double AngleRadians, Border, CenterX, CenterY;
        public int Steps;
    }

    /// <summary>A draw:gradient (or a draw:opacity as grey levels of opacity) as its model; null when the element is missing.</summary>
    public static GradientModel? Gradient(XElement? element, bool isOpacity, bool angleIsTenths, int steps)
    {
        if (element is null) return null;
        XNamespace draw = OdtStyleCascade.Draw;
        var model = new GradientModel { Style = (string?)element.Attribute(draw + "style") ?? "linear", Steps = Math.Max(0, steps) };
        model.CenterX = Percent((string?)element.Attribute(draw + "cx")) ?? 0;
        model.CenterY = Percent((string?)element.Attribute(draw + "cy")) ?? 0;
        model.Border = Math.Max(0, Math.Min(1, Percent((string?)element.Attribute(draw + "border")) ?? 0));
        model.AngleRadians = TenthDegrees((string?)element.Attribute(draw + "angle"), angleIsTenths) / 10.0 * Math.PI / 180;
        if (isOpacity)
        {
            // grey = transparency (100 - opacity), stored as a byte (TransGradientStyle.cxx:108-127); the mask wants opacity
            double Level(string name) { double opacity = Percent((string?)element.Attribute(draw + name)) ?? 0; int n = (int)((100 - opacity * 100) * 255 / 100); return 1 - n / 255.0; }
            double start = Level("start"), end = Level("end");
            model.Stops.Add(new Stop(0, start, start, start));
            model.Stops.Add(new Stop(1, end, end, end));
            return model;
        }
        (double r, double g, double b) Rgb(string? css) => OdtShapeStyle.ParseColor(css) is { } c ? c : (0, 0, 0);
        var loStops = element.Elements().Where(e => e.Name.LocalName == "gradient-stop").ToList();
        if (loStops.Count > 0)
        {
            foreach (XElement stop in loStops)
            {
                if ((string?)stop.Attributes().FirstOrDefault(a => a.Name.LocalName == "color-type")?.Value != "rgb") continue;
                string? offsetText = (string?)stop.Attributes().FirstOrDefault(a => a.Name.LocalName == "offset")?.Value;
                if (!double.TryParse(offsetText, NumberStyles.Float, CultureInfo.InvariantCulture, out double offset)) continue;
                (double r, double g, double b) = Rgb((string?)stop.Attributes().FirstOrDefault(a => a.Name.LocalName == "color-value")?.Value);
                model.Stops.Add(new Stop(Math.Max(0, Math.Min(1, offset)), r, g, b));
            }
            model.Stops.Sort((x, y) => x.Offset.CompareTo(y.Offset));
        }
        if (model.Stops.Count == 0)
        {
            (double r, double g, double b) start = Rgb((string?)element.Attribute(draw + "start-color")), end = Rgb((string?)element.Attribute(draw + "end-color"));
            model.Stops.Add(new Stop(0, start.r, start.g, start.b));
            model.Stops.Add(new Stop(1, end.r, end.g, end.b));
        }
        double startIntensity = Percent((string?)element.Attribute(draw + "start-intensity")) ?? 1, endIntensity = Percent((string?)element.Attribute(draw + "end-intensity")) ?? 1;
        startIntensity = Math.Max(0, Math.Min(1, startIntensity)); endIntensity = Math.Max(0, Math.Min(1, endIntensity));
        if (startIntensity != 1 || endIntensity != 1)
            model.Stops = model.Stops.Select(s => { double k = startIntensity * (1 - s.Offset) + endIntensity * s.Offset; return new Stop(s.Offset, s.R * k, s.G * k, s.B * k); }).ToList();
        return model;
    }

    /// <summary>sax Converter::convert10thDegAngle (converter.cxx:878-912); <paramref name="unitlessIsTenths"/> is its isWrongOOo10thDegAngle.</summary>
    internal static double TenthDegrees(string? value, bool unitlessIsTenths)
    {
        if (string.IsNullOrWhiteSpace(value)) return 0;
        string text = value!.Trim();
        int end = 0;
        while (end < text.Length && (char.IsDigit(text[end]) || text[end] is '.' or '-' or '+' or 'e' or 'E') && !(text[end] is 'e' or 'E' && end + 1 < text.Length && !char.IsDigit(text[end + 1]) && text[end + 1] is not ('-' or '+'))) end++;
        if (!double.TryParse(text.Substring(0, end), NumberStyles.Float, CultureInfo.InvariantCulture, out double angle)) return 0;
        string unit = text.Substring(end);
        angle = unit switch
        {
            "deg" => angle * 10,
            "grad" => angle * 9,
            "rad" => angle * 1800 / Math.PI,
            "" => unitlessIsTenths ? angle : angle * 10,
            _ => 0,
        };
        angle = Math.Round(angle) % 3600;
        return angle < 0 ? angle + 3600 : angle;
    }

    internal static double? Percent(string? value)
    {
        if (value is null) return null;
        string v = value.Trim();
        if (v.EndsWith("%", StringComparison.Ordinal)) v = v.Substring(0, v.Length - 1);
        return double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d / 100 : null;
    }

    /// <summary>The (offset, colour) breakpoints of the stops over 0..1, banded as LibreOffice bands them; a band is two stops at its ends.</summary>
    private static List<(double T, double R, double G, double B)> Breakpoints(GradientModel model, Func<(double, double, double), (double, double, double)> adapt, bool smooth)
    {
        var result = new List<(double T, double R, double G, double B)>();
        List<Stop> stops = model.Stops;
        void Add(double t, (double r, double g, double b) c)
        {
            var a = adapt(c);
            if (result.Count > 0 && result[result.Count - 1] == (t, a.Item1, a.Item2, a.Item3)) return;     // the same breakpoint twice draws nothing more
            result.Add((t, a.Item1, a.Item2, a.Item3));
        }
        if (stops.Count == 1) { Add(0, (stops[0].R, stops[0].G, stops[0].B)); Add(1, (stops[0].R, stops[0].G, stops[0].B)); return result; }
        Add(0, (stops[0].R, stops[0].G, stops[0].B));
        for (int i = 0; i + 1 < stops.Count; i++)
        {
            Stop a = stops[i], b = stops[i + 1];
            if (b.Offset <= a.Offset) continue;
            if (smooth) { Add(a.Offset, (a.R, a.G, a.B)); Add(b.Offset, (b.R, b.G, b.B)); continue; }
            int n = StepCount(model.Steps, a, b);
            for (int k = 0; k < n; k++)
            {
                double f = n == 1 ? 0 : k / (double)(n - 1);
                var c = (a.R + (b.R - a.R) * f, a.G + (b.G - a.G) * f, a.B + (b.B - a.B) * f);
                Add(a.Offset + (b.Offset - a.Offset) * k / n, c);
                Add(a.Offset + (b.Offset - a.Offset) * (k + 1) / n, c);
            }
        }
        Stop last = stops[stops.Count - 1];
        Add(1, (last.R, last.G, last.B));
        return result;
    }

    /// <summary>
    /// True when LibreOffice on Windows draws this fill gradient smooth: Skia is the only renderer there (vcl/skia/SkiaHelper.cxx:446-447
    /// "macOS/win can __only__ render via skia"); VclPixelProcessor2D hands a FillGradientPrimitive2D without alpha to
    /// OutputDevice::DrawGradient (vclpixelprocessor2d.cxx:911-1051) unless FillGradientAttribute::cannotBeHandledByVCL - more or
    /// other than two stops at 0 and 1, or a style other than linear / axial / radial (fillgradientattribute.cxx:144-176); the
    /// definition range is the outline's own range (svx sdrdecompositiontools.cxx createPolyPolygonFillPrimitive without a range),
    /// so it always holds the output range. DrawGradient first asks the graphics to draw it natively (vcl/source/outdev/gradient.cxx:108-113)
    /// and SkiaSalGraphicsImpl::drawGradient (vcl/skia/gdiimpl.cxx:1942-2023) does, as one Skia linear / radial shader, whenever
    /// the step count is 0 - so the banded fallback (GetGradientSteps, gradient.cxx:569-591: height / 4 bands) never runs on
    /// Windows. A uniform fill transparency, or a transparency gradient of the same definition (sameDefinitionThanAlpha,
    /// fillgradientattribute.cxx:226-257), makes it a gradient with alpha (sdrdecompositiontools.cxx createPolyPolygonFillPrimitive),
    /// which the processor decomposes into bands (vclpixelprocessor2d.cxx:914-919). The Cairo renderer (Linux) also draws
    /// automatic steps smooth for these and more styles (cairopixelprocessor2d.cxx:3492-3504).
    /// </summary>
    private static bool IsDrawnSmoothBySkia(GradientModel model)
        => model.Steps == 0 && model.Style is "linear" or "axial" or "radial"
           && model.Stops.Count == 2 && model.Stops[0].Offset == 0 && model.Stops[1].Offset == 1;

    /// <summary>FillGradientAttribute::sameDefinitionThanAlpha (fillgradientattribute.cxx:226-257).</summary>
    private static bool HasSameDefinition(GradientModel fill, GradientModel alpha)
    {
        if (fill.Style != alpha.Style || fill.Border != alpha.Border || fill.Steps != alpha.Steps) return false;
        if (fill.Style is not ("linear" or "axial") && (fill.CenterX != alpha.CenterX || fill.CenterY != alpha.CenterY)) return false;
        if (fill.Style != "radial" && fill.AngleRadians != alpha.AngleRadians) return false;
        return fill.Stops.Count == alpha.Stops.Count && fill.Stops.Select(stop => stop.Offset).SequenceEqual(alpha.Stops.Select(stop => stop.Offset));
    }

    /// <summary>calculateNumberOfSteps (gradienttools.cxx:479-497): at most the colour distance times 127.5 (rounded), at least 1.</summary>
    private static int StepCount(int requested, Stop a, Stop b)
    {
        double distance = Math.Max(Math.Abs(a.R - b.R), Math.Max(Math.Abs(a.G - b.G), Math.Abs(a.B - b.B)));
        int max = (int)(distance * 127.5 + 0.5);
        if (requested == 0 || requested > max) requested = max;
        return Math.Max(1, requested);
    }

    private static string Css(double r, double g, double b)
        => "#" + ((Byte(r) << 16) | (Byte(g) << 8) | Byte(b)).ToString("x6", CultureInfo.InvariantCulture);

    private static int Byte(double c) => (int)Math.Round(Math.Max(0, Math.Min(1, c)) * 255);

    private static string N(double v) => WordShapeHtml.Number6(v);

    /// <summary>
    /// The paint of a gradient over the definition rectangle (x, y, w, h) in the shape's points, with an optional transparency
    /// gradient, uniform opacity handled by the caller. <paramref name="adapt"/> changes every colour (a path's brightness).
    /// </summary>
    public static (string Definitions, string Paint) GradientPaint(GradientModel gradient, GradientModel? transparency, double x, double y, double w, double h,
                                                                 Func<(double, double, double), (double, double, double)> adapt, bool hasUniformTransparency = false,
                                                                 bool isSolidFill = false)
    {
        var defs = new StringBuilder();
        // isSolidFill: one colour under a transparency gradient - no fill gradient, so never "the same definition" (sdrdecompositiontools.cxx:428-449)
        bool sharesDefinition = !isSolidFill && transparency is not null && HasSameDefinition(gradient, transparency);
        bool smooth = IsDrawnSmoothBySkia(gradient) && !hasUniformTransparency && !sharesDefinition;
        string content = GradientContent(gradient, x, y, w, h, adapt, defs, smooth);
        // a transparency gradient of another definition is its own FillGradientPrimitive2D (sdrdecompositiontools.cxx:532), drawn
        // by the same processor into the alpha buffer (vclprocessor2d.cxx:1045-1065 RenderTransparencePrimitive2D) - smooth too
        bool isMaskSmooth = transparency is not null && IsDrawnSmoothBySkia(transparency) && !sharesDefinition;
        return Pattern(defs, content, transparency, x, y, w, h, isMaskSmooth);
    }

    /// <summary>A hatch (lines over an optional background colour) as a paint, with an optional transparency gradient.</summary>
    public static (string Definitions, string Paint) HatchPaint(string style, (double r, double g, double b) color, double distancePt, double angleRadians, (double r, double g, double b)? background,
                                                              GradientModel? transparency, double x, double y, double w, double h,
                                                              Func<(double, double, double), (double, double, double)> adapt)
    {
        var defs = new StringBuilder();
        var content = new StringBuilder();
        if (background is { } back)
        {
            var b = adapt(back);
            content.Append("<rect x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" width=\"").Append(N(w)).Append("\" height=\"").Append(N(h)).Append("\" fill=\"").Append(Css(b.Item1, b.Item2, b.Item3)).Append("\"/>");
        }
        var line = adapt(color);
        string stroke = Css(line.Item1, line.Item2, line.Item3);
        void Set(double angle)
        {
            // GeoTexSvxHatch: the box grown for the rotation; lines at a * distance (a = 1 .. steps - 1) across its width
            double a = -angle, cos = Math.Abs(Math.Cos(a)), sin = Math.Abs(Math.Sin(a));
            double sizeX = a != 0 ? w * cos + h * sin : w, sizeY = a != 0 ? h * cos + w * sin : h;
            // lines closer than 1/2000 of the area are drawn that far apart (closer would be below a device pixel)
            if (!(distancePt > 0) || !(!double.IsNaN(distancePt) && !double.IsInfinity(distancePt))) distancePt = 0;
            else distancePt = Math.Max(distancePt, sizeY / 2000);
            double steps = distancePt != 0 ? sizeY / distancePt : 10;
            int count = (int)Math.Round(steps + 0.5, MidpointRounding.AwayFromZero);
            double spacing = sizeY / steps;
            content.Append("<g transform=\"translate(").Append(N(x + w / 2)).Append(' ').Append(N(y + h / 2)).Append(") rotate(").Append(N(a * 180 / Math.PI))
                   .Append(") translate(").Append(N(-sizeX / 2)).Append(' ').Append(N(-sizeY / 2)).Append(")\" stroke=\"").Append(stroke).Append("\" stroke-width=\"0.75\">");
            var d = new StringBuilder();
            for (int k = 1; k < count; k++) d.Append('M').Append(N(0)).Append(' ').Append(N(spacing * k)).Append('H').Append(N(sizeX));
            content.Append("<path d=\"").Append(d).Append("\" fill=\"none\"/></g>");
        }
        // fillhatchprimitive2d.cxx:60-100: triple = + the set at angle - 45 degrees, double = + angle - 90
        if (style == "triple") Set(angleRadians - Math.PI / 4);
        if (style is "triple" or "double") Set(angleRadians - Math.PI / 2);
        Set(angleRadians);
        return Pattern(defs, content.ToString(), transparency, x, y, w, h, transparency is not null && IsDrawnSmoothBySkia(transparency));
    }

    /// <summary>A transparency gradient alone as an SVG luminance mask in the shape's points (for a picture fill): its definitions and id.</summary>
    public static (string Definitions, string Id) Mask(GradientModel transparency, double x, double y, double w, double h)
    {
        var defs = new StringBuilder();
        string content = GradientContent(transparency, x, y, w, h, c => c, defs, IsDrawnSmoothBySkia(transparency));
        double reach = Reach(w, h);
        string mask = "<mask maskUnits=\"userSpaceOnUse\" x=\"" + N(x - reach) + "\" y=\"" + N(y - reach) + "\" width=\"" + N(w + 2 * reach) + "\" height=\"" + N(h + 2 * reach) + "\">" + content + "</mask>";
        string id = "odt-fill-mask-" + WordShapeHtml.StableHashOf(defs + mask);
        defs.Append(mask.Insert("<mask".Length, " id=\"" + id + "\""));
        return (defs.ToString(), id);
    }

    private static (string Definitions, string Paint) Pattern(StringBuilder defs, string content, GradientModel? transparency, double x, double y, double w, double h,
                                                              bool isMaskSmooth)
    {
        string masked = content;
        if (transparency is not null)
        {
            var maskDefs = new StringBuilder();
            string maskContent = GradientContent(transparency, x, y, w, h, c => c, maskDefs, isMaskSmooth);
            double reach = Reach(w, h);
            string mask = "<mask maskUnits=\"userSpaceOnUse\" x=\"" + N(x - reach) + "\" y=\"" + N(y - reach) + "\" width=\"" + N(w + 2 * reach) + "\" height=\"" + N(h + 2 * reach) + "\">" + maskContent + "</mask>";
            string maskId = "odt-fill-mask-" + WordShapeHtml.StableHashOf(maskDefs + mask);
            defs.Append(maskDefs).Append(mask.Insert("<mask".Length, " id=\"" + maskId + "\""));
            masked = "<g mask=\"url(#" + maskId + ")\">" + content + "</g>";
        }
        double r2 = Reach(w, h);
        // the tile's content is drawn from the tile's corner: shift it back so the content stays in the shape's coordinates
        string tile = "<pattern patternUnits=\"userSpaceOnUse\" x=\"" + N(x - r2) + "\" y=\"" + N(y - r2) + "\" width=\"" + N(w + 2 * r2) + "\" height=\"" + N(h + 2 * r2) + "\">"
                    + "<g transform=\"translate(" + N(r2 - x) + " " + N(r2 - y) + ")\">" + masked + "</g></pattern>";
        string id = "odt-fill-" + WordShapeHtml.StableHashOf(defs + tile);
        defs.Append(tile.Insert("<pattern".Length, " id=\"" + id + "\""));
        return (defs.ToString(), "url(#" + id + ")");
    }

    /// <summary>How far beyond the definition box a fill must reach (outlines may leave their box; pads continue there).</summary>
    private static double Reach(double w, double h) => 4 * Math.Max(Math.Max(w, h), 1) + 100;

    /// <summary>The gradient's SVG content over a huge area around the box: the start colour as a base, the gradient over it.</summary>
    private static string GradientContent(GradientModel model, double x, double y, double w, double h, Func<(double, double, double), (double, double, double)> adapt, StringBuilder defs,
                                          bool smooth)
    {
        List<(double T, double R, double G, double B)> points = Breakpoints(model, adapt, smooth);
        string ColorAt(int i) => Css(points[i].R, points[i].G, points[i].B);
        double reach = Reach(w, h);
        string Huge(string fill) => "<rect x=\"" + N(x - reach) + "\" y=\"" + N(y - reach) + "\" width=\"" + N(w + 2 * reach) + "\" height=\"" + N(h + 2 * reach) + "\" fill=\"" + fill + "\"/>";
        string start = ColorAt(0), endColor = ColorAt(points.Count - 1);
        string Define(string element)
        {
            string id = "odt-gradient-" + WordShapeHtml.StableHashOf(element);
            int space = element.IndexOf(' ');
            defs.Append(element.Insert(space, " id=\"" + id + "\""));
            return id;
        }
        string StopsHtml(IEnumerable<(double offset, string color)> stops)
        {
            var html = new StringBuilder();
            double last = 0;
            string? previous = null;
            foreach ((double offset, string color) in stops)
            {
                last = Math.Max(last, Math.Max(0, Math.Min(1, offset)));
                string key = N(last) + color;
                if (key == previous) continue;           // the mirrored halves of an axial gradient meet in one stop
                previous = key;
                html.Append("<stop offset=\"").Append(N(last)).Append("\" stop-color=\"").Append(color).Append("\"/>");
            }
            return html.ToString();
        }
        double border = model.Border, angle = -model.AngleRadians;
        var content = new StringBuilder();
        switch (model.Style)
        {
            case "axial":
            case "linear":
            default:
            {
                bool axial = model.Style == "axial";
                double cos = Math.Abs(Math.Cos(angle)), sin = Math.Abs(Math.Sin(angle));
                double sizeX = angle != 0 ? w * cos + h * sin : w, sizeY = angle != 0 ? h * cos + w * sin : h;
                IEnumerable<(double, string)> stops;
                if (!axial)
                    // unit y = (s - border) / (1 - border): the colour at s
                    stops = points.Select((p, i) => (border + (1 - border) * p.T, ColorAt(i)));
                else
                {
                    // unit |y| = |s - 0.5| / ((1 - border) / 2), the colour at |y| (the start colour in the middle)
                    var half = points.Select((p, i) => (0.5 + (1 - border) / 2 * p.T, ColorAt(i))).ToList();
                    stops = half.AsEnumerable().Reverse().Select(p => (1 - p.Item1, p.Item2)).Concat(half);
                }
                string gradient = Define("<linearGradient gradientUnits=\"userSpaceOnUse\" x1=\"0\" y1=\"0\" x2=\"0\" y2=\"" + N(sizeY) + "\">" + StopsHtml(stops) + "</linearGradient>");
                content.Append(Huge(start));
                content.Append("<g transform=\"translate(").Append(N(x + w / 2)).Append(' ').Append(N(y + h / 2)).Append(") rotate(").Append(N(angle * 180 / Math.PI))
                       .Append(") translate(").Append(N(-sizeX / 2)).Append(' ').Append(N(-sizeY / 2)).Append(")\">")
                       .Append("<rect x=\"").Append(N(-reach - sizeX)).Append("\" y=\"").Append(N(-reach - sizeY)).Append("\" width=\"").Append(N(2 * reach + 3 * sizeX))
                       .Append("\" height=\"").Append(N(2 * reach + 3 * sizeY)).Append("\" fill=\"url(#").Append(gradient).Append(")\"/></g>");
                break;
            }
            case "radial":
            case "ellipsoid":
            {
                bool circular = model.Style == "radial";
                double sizeX, sizeY;
                if (circular) { double diagonal = Math.Sqrt(w * w + h * h); sizeX = sizeY = diagonal; }
                else { sizeX = Math.Sqrt(2) * w; sizeY = Math.Sqrt(2) * h; }
                // Skia's native radial (gdiimpl.cxx:2006-2014): radius half the diagonal, the end colour solid out to the border,
                // then smooth to the start colour - not the decomposition's (1 - border) unit radius
                bool skiaRadial = smooth && circular;
                double unitX = (skiaRadial ? 1 : 1 - border) / 2 * sizeX, unitY = (skiaRadial ? 1 : 1 - border) / 2 * sizeY;
                double cx = x + w / 2 + (model.CenterX - 0.5) * w, cy = y + h / 2 + (model.CenterY - 0.5) * h;
                double rotation = circular ? 0 : angle;
                // alpha = 1 - radius: offset q = 1 - t
                var stops = skiaRadial ? new[] { (0.0, ColorAt(points.Count - 1)), (border, ColorAt(points.Count - 1)), (1.0, ColorAt(0)) }
                                       : points.Select((p, i) => (1 - p.T, ColorAt(i))).Reverse();
                double aspect = sizeY == 0 ? 1 : sizeX / sizeY;
                string scale = circular ? "" : aspect > 1 ? " gradientTransform=\"scale(" + N(aspect) + " 1)\"" : aspect > 0 ? " gradientTransform=\"scale(1 " + N(1 / aspect) + ")\"" : "";
                string gradient = Define("<radialGradient gradientUnits=\"userSpaceOnUse\" cx=\"0\" cy=\"0\" r=\"1\"" + scale + ">" + StopsHtml(stops) + "</radialGradient>");
                content.Append(Huge(start));
                content.Append("<g transform=\"translate(").Append(N(cx)).Append(' ').Append(N(cy)).Append(") rotate(").Append(N(rotation * 180 / Math.PI))
                       .Append(") scale(").Append(N(Math.Max(unitX, 1e-9))).Append(' ').Append(N(Math.Max(unitY, 1e-9))).Append(")\">")
                       .Append("<rect x=\"-1\" y=\"-1\" width=\"2\" height=\"2\" fill=\"url(#").Append(gradient).Append(")\"/></g>");
                break;
            }
            case "square":
            case "rectangular":
            {
                double sizeX = w, sizeY = h;
                if (model.Style == "square") sizeX = sizeY = Math.Max(w, h);
                double cos = Math.Abs(Math.Cos(angle)), sin = Math.Abs(Math.Sin(angle));
                if (angle != 0) { double nx = sizeX * cos + sizeY * sin, ny = sizeY * cos + sizeX * sin; sizeX = nx; sizeY = ny; }
                double ux = (1 - border) / 2 * sizeX, uy = (1 - border) / 2 * sizeY, m = Math.Min(ux, uy);
                double cx = x + w / 2 + (model.CenterX - 0.5) * w, cy = y + h / 2 + (model.CenterY - 0.5) * h;
                content.Append(Huge(start));
                if (m <= 0) break;
                // alpha = distance to the nearest edge / the shorter half side: four regions, each a gradient across it
                var stops = points.Select((p, i) => (p.T, ColorAt(i))).ToList();
                content.Append("<g transform=\"translate(").Append(N(cx)).Append(' ').Append(N(cy)).Append(") rotate(").Append(N(angle * 180 / Math.PI)).Append(")\">");
                string P(double px, double py) => N(px) + " " + N(py);
                (double x1, double y1, double x2, double y2, string polygon)[] regions =
                [
                    (0, -uy, 0, -uy + m, "M" + P(-ux, -uy) + "L" + P(ux, -uy) + "L" + P(ux - m, -uy + m) + "L" + P(-ux + m, -uy + m) + "Z"),
                    (0, uy, 0, uy - m, "M" + P(-ux, uy) + "L" + P(-ux + m, uy - m) + "L" + P(ux - m, uy - m) + "L" + P(ux, uy) + "Z"),
                    (-ux, 0, -ux + m, 0, "M" + P(-ux, -uy) + "L" + P(-ux + m, -uy + m) + "L" + P(-ux + m, uy - m) + "L" + P(-ux, uy) + "Z"),
                    (ux, 0, ux - m, 0, "M" + P(ux, -uy) + "L" + P(ux, uy) + "L" + P(ux - m, uy - m) + "L" + P(ux - m, -uy + m) + "Z"),
                ];
                foreach ((double x1, double y1, double x2, double y2, string polygon) in regions)
                {
                    string gradient = Define("<linearGradient gradientUnits=\"userSpaceOnUse\" x1=\"" + N(x1) + "\" y1=\"" + N(y1) + "\" x2=\"" + N(x2) + "\" y2=\"" + N(y2) + "\">" + StopsHtml(stops) + "</linearGradient>");
                    content.Append("<path d=\"").Append(polygon).Append("\" fill=\"url(#").Append(gradient).Append(")\"/>");
                }
                content.Append("</g>");
                break;
            }
        }
        _ = endColor;
        return content.ToString();
    }
}
