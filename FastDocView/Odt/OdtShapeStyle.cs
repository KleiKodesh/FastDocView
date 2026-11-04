using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// A drawing shape's graphic style (ODF 1.3 part 3 20.x graphic properties, read through the style's parent chain and the
/// document's default graphic style), applied to a <see cref="WordShape"/> as LibreOffice applies it (xmloff
/// sdpropls.cxx:98-140 maps every attribute read here; master read 2026-11-04).
/// Defaults when neither the style nor the default style sets a property are LibreOffice's item-pool defaults
/// (svx svdattr.cxx:175-206): draw:fill solid (include/svx/xfillit0.hxx:37) in #729fcf, draw:stroke solid
/// (include/svx/xlineit0.hxx:34) in #3465a4 (COL_DEFAULT_SHAPE_FILLING / _STROKE, include/svx/xdef.hxx:83-84), stroke width 0
/// = a hairline (include/svx/xlnwtit.hxx:33; drawn one device pixel, here 0.75 pt), line join round (include/svx/xlinjoit.hxx:35),
/// cap butt (xlncapit.hxx:35), marker width 2 mm (xlnstwit.hxx:26 LINE_START_WIDTH_DEFAULT 200), not centred (xlnstcit.hxx:31).
/// </summary>
internal sealed class OdtShapeStyle
{
    public const string DefaultFillColor = "#729fcf", DefaultStrokeColor = "#3465a4";
    private static readonly XNamespace Draw = OdtStyleCascade.Draw, Svg = OdtStyleCascade.Svg, Fo = OdtStyleCascade.Fo, Style = OdtStyleCascade.Style;

    private readonly OdtStyleCascade _styles;
    private readonly string? _styleName;
    private readonly double _officeVersion;

    public OdtShapeStyle(OdtStyleCascade styles, string? styleName, double officeVersion)
    {
        _styles = styles;
        _styleName = styleName;
        _officeVersion = officeVersion;
    }

    public string? Get(XName attribute) => _styles.GraphicProperty(_styleName, attribute);

    public string FillKind => Get(Draw + "fill") ?? "solid";
    public string StrokeKind => Get(Draw + "stroke") ?? "solid";
    public bool IsFilled => FillKind != "none";
    public bool IsStroked => StrokeKind != "none";

    /// <summary>A length attribute in points (cm / mm / in / pt / pc / px); null when unset or not a length.</summary>
    public static double? Length(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        string v = value!.Trim();
        int unitStart = v.Length;
        while (unitStart > 0 && char.IsLetter(v[unitStart - 1])) unitStart--;
        if (!double.TryParse(v.Substring(0, unitStart), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return null;
        return v.Substring(unitStart) switch
        {
            "pt" => number, "in" => number * 72, "cm" => number * 72 / 2.54, "mm" => number * 72 / 25.4, "pc" => number * 12, "px" => number * 0.75,
            "" => number * 72 / 2540,         // a unitless measure is 1/100 mm (the core unit; LibreOffice's converter falls back to it)
            _ => null,
        };
    }

    /// <summary>"#rrggbb" as RGB 0..1; null otherwise.</summary>
    public static (double R, double G, double B)? ParseColor(string? value)
    {
        if (value is not { Length: 7 } || value[0] != '#' || !int.TryParse(value.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb)) return null;
        return (((rgb >> 16) & 0xFF) / 255.0, ((rgb >> 8) & 0xFF) / 255.0, (rgb & 0xFF) / 255.0);
    }

    private static string Hex((double R, double G, double B) c)
        => "#" + (((int)Math.Round(c.R * 255) << 16) | ((int)Math.Round(c.G * 255) << 8) | (int)Math.Round(c.B * 255)).ToString("x6", CultureInfo.InvariantCulture);

    /// <summary>A path brightness / colour-data change on an RGB colour (the same rules as <see cref="WordShapeColor.Brightened"/>).</summary>
    private static Func<(double, double, double), (double, double, double)> Adapter(double brightness)
        => c =>
        {
            if (brightness == 0) return c;
            string adapted = WordShapeColor.Brightened(Hex(c).ToUpperInvariant(), brightness);
            return ParseColor(adapted.ToLowerInvariant()) is { } rgb ? (rgb.R, rgb.G, rgb.B) : c;
        };

    /// <summary>
    /// The fill (draw:fill 20.118: none / solid / gradient / hatch / bitmap) over the definition rectangle (the shape's box,
    /// LibreOffice's fill range), with draw:opacity (20.124, uniform: FillTransparence = 100 - opacity) or draw:opacity-name
    /// (20.125, a transparency gradient). <paramref name="pictureOf"/> gives a fill image's id and its natural size.
    /// </summary>
    public void ApplyFill(WordShape shape, (double X, double Y, double W, double H) range, Func<XElement, (string Id, (double W, double H)? Size)?>? pictureOf)
    {
        string kind = FillKind;
        shape.FillColor = null; shape.GradientStops = null; shape.FillPainter = null;
        if (kind == "none") return;
        if (Get(Draw + "opacity") is { } opacityText && OdtFillPaint.Percent(opacityText) is { } opacity) shape.FillOpacity = Math.Max(0, Math.Min(1, opacity));
        OdtFillPaint.GradientModel? transparency = Get(Draw + "opacity-name") is { } opacityName
            ? OdtFillPaint.Gradient(_styles.DrawingStyle("opacity", opacityName), isOpacity: true, _officeVersion, 0) : null;
        if (transparency is not null) shape.FillOpacity = 1;            // a transparency gradient replaces the uniform transparency
        string color = ParseColor(Get(Draw + "fill-color")) is { } fillColor ? Hex(fillColor) : DefaultFillColor;
        switch (kind)
        {
            case "gradient":
            {
                int steps = int.TryParse(Get(Draw + "gradient-step-count"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int s) ? s : 0;
                if (OdtFillPaint.Gradient(_styles.Gradient(Get(Draw + "fill-gradient-name")), isOpacity: false, _officeVersion, steps) is not { } gradient) { shape.FillColor = color; break; }
                shape.FillPainter = brightness => OdtFillPaint.GradientPaint(gradient, transparency, range.X, range.Y, range.W, range.H, Adapter(brightness));
                return;
            }
            case "hatch":
            {
                XElement? hatch = _styles.DrawingStyle("hatch", Get(Draw + "fill-hatch-name"));
                if (hatch is null) { shape.FillColor = color; break; }
                string style = (string?)hatch.Attribute(Draw + "style") ?? "single";
                (double, double, double) lineColor = ParseColor((string?)hatch.Attribute(Draw + "color")) is { } hc ? hc : (0, 0, 0);
                double distance = Length((string?)hatch.Attribute(Draw + "distance")) ?? 0;
                // draw:rotation: unitless values stay tenths of a degree (HatchStyle.cxx:96-112, tdf#161327)
                double angle = OdtFillPaint.TenthDegrees((string?)hatch.Attribute(Draw + "rotation"), _officeVersion, alwaysTenths: true) / 10.0 * Math.PI / 180;
                (double, double, double)? background = Get(Draw + "fill-hatch-solid") == "true" ? ParseColor(color) is { } bc ? (bc.R, bc.G, bc.B) : null : null;
                shape.FillPainter = brightness => OdtFillPaint.HatchPaint(style, lineColor, distance, angle, background, transparency, range.X, range.Y, range.W, range.H, Adapter(brightness));
                return;
            }
            case "bitmap":
                if (pictureOf is not null && _styles.DrawingStyle("fill-image", Get(Draw + "fill-image-name")) is { } image && pictureOf(image) is { } picture)
                {
                    ApplyBitmap(shape, range, picture.Id, picture.Size);
                    shape.FillPicture = shape.FillPicture with { Opacity = shape.FillOpacity };
                    if (transparency is not null) { shape.FillPictureMask = OdtFillPaint.Mask(transparency, range.X, range.Y, range.W, range.H); shape.FillPicture = shape.FillPicture with { Opacity = 1 }; }
                    return;
                }
                return;
            default:
                shape.FillColor = color;
                break;
        }
        if (transparency is not null && shape.FillColor is { } solid)
        {
            // a solid colour under a transparency gradient: the gradient machinery with one colour
            var single = new OdtFillPaint.GradientModel();
            (double r, double g, double b) = ParseColor(solid) ?? (0, 0, 0);
            single.Stops.Add(new OdtFillPaint.Stop(0, r, g, b)); single.Stops.Add(new OdtFillPaint.Stop(1, r, g, b));
            shape.FillColor = null;
            shape.FillPainter = brightness => OdtFillPaint.GradientPaint(single, transparency, range.X, range.Y, range.W, range.H, Adapter(brightness));
        }
    }

    /// <summary>
    /// A bitmap fill (draw:fill-image-name) placed as LibreOffice places it (drawinglayer sdrfillgraphicattribute.cxx:200-300
    /// createFillGraphicAttribute): style:repeat stretch (the default in an item set without XATTR_FILLBMP_STRETCH set is
    /// stretch, svx xflbstit / sdpropls XML_SD_TYPE_BITMAP_MODE) = the box; repeat = tiles of draw:fill-image-width / -height
    /// (a percentage = of the box, absent = the picture's own size), aligned at draw:fill-image-ref-point (centre / sides, -1 0 1),
    /// shifted by draw:fill-image-ref-point-x / -y percent of a tile, every other row or column moved by draw:tile-repeat-offset;
    /// no-repeat = one picture of that size at the reference point.
    /// </summary>
    private void ApplyBitmap(WordShape shape, (double X, double Y, double W, double H) range, string pictureId, (double W, double H)? natural)
    {
        string repeat = Get(Style + "repeat") ?? "repeat";
        shape.FillPictureId = pictureId;
        if (repeat == "stretch")
        {
            shape.FillPicture = new WordShapePictureFill(0, 0, 0, 0, false, 0, 0, 0, 0, null, 1);
            return;
        }
        double w = range.W > 0 ? range.W : 1, h = range.H > 0 ? range.H : 1;
        double SizeOf(string attribute, double box, double own)
        {
            string? value = Get(Draw + attribute);
            if (value is not null && value.EndsWith("%", StringComparison.Ordinal) && OdtFillPaint.Percent(value) is { } p && p != 0) return box * p;
            if (Length(value) is { } absolute && absolute != 0) return absolute;
            return own;
        }
        double tileW = SizeOf("fill-image-width", w, natural?.W ?? w), tileH = SizeOf("fill-image-height", h, natural?.H ?? h);
        string refPoint = Get(Draw + "fill-image-ref-point") ?? "center";
        double rx = refPoint.Contains("left") ? -1 : refPoint.Contains("right") ? 1 : 0, ry = refPoint.StartsWith("top", StringComparison.Ordinal) ? -1 : refPoint.StartsWith("bottom", StringComparison.Ordinal) ? 1 : 0;
        double left = rx == 0 ? (w - tileW) / 2 : rx == 1 ? w - tileW : 0;
        double top = ry == 0 ? (h - tileH) / 2 : ry == 1 ? h - tileH : 0;
        bool tiled = repeat == "repeat";
        if (tiled)
        {
            left += tileW * (OdtFillPaint.Percent(Get(Draw + "fill-image-ref-point-x")) ?? 0);
            top += tileH * (OdtFillPaint.Percent(Get(Draw + "fill-image-ref-point-y")) ?? 0);
            if (Get(Draw + "tile-repeat-offset") is { } offset)
            {
                string[] parts = offset.Split(' ');
                double amount = OdtFillPaint.Percent(parts[0]) ?? 0;
                if (parts.Length > 1 && parts[1] == "vertical") shape.TileColumnOffset = amount; else shape.TileRowOffset = amount;
            }
            shape.FillPicture = new WordShapePictureFill(0, 0, 0, 0, true, left, top, tileW, tileH, null, 1);
        }
        else
            shape.FillPicture = new WordShapePictureFill(left / w, top / h, (w - left - tileW) / w, (h - top - tileH) / h, false, 0, 0, 0, 0, null, 1);
    }

    /// <summary>
    /// The line: draw:stroke (20.167 none / solid / dash), svg:stroke-color, svg:stroke-width (0 = hairline), svg:stroke-opacity
    /// (LineTransparence), draw:stroke-linejoin (miter / middle = miter, round, bevel, none = no join, sdrattributecreator.cxx:123-145),
    /// svg:stroke-linecap; a dash from its draw:stroke-dash style (XDash::CreateDotDashArray, svx xattr.cxx:506-645: dots then
    /// dashes, each followed by the distance; lengths in percent are relative to the line width, a zero length is a dot of the
    /// line width; absolute lengths at least SMALLEST_DASH_WIDTH 26.95 hundredths of a mm, :504); the dash's draw:style rect / round
    /// does not change the drawing (createNewSdrLineAttribute, sdrattributecreator.cxx:288-332, uses only the dash array and
    /// XATTR_LINECAP); markers (draw:marker-start / -end with -width / -center).
    /// </summary>
    public void ApplyStroke(WordShape shape)
    {
        string kind = StrokeKind;
        if (kind == "none") { shape.LineColor = null; return; }
        shape.LineColor = ParseColor(Get(Svg + "stroke-color")) is { } c ? Hex(c) : DefaultStrokeColor;
        double width = Length(Get(Svg + "stroke-width")) ?? 0;
        shape.LineWidthPt = width > 0 ? width : 0.75;                 // a zero-width line is a one-pixel hairline
        if (Get(Svg + "stroke-opacity") is { } opacity && OdtFillPaint.Percent(opacity) is { } o) shape.LineOpacity = Math.Max(0, Math.Min(1, o));
        shape.LineJoin = (Get(Draw + "stroke-linejoin") ?? "round") switch { "bevel" => WordLineJoin.Bevel, "miter" or "middle" => WordLineJoin.Miter, "none" => WordLineJoin.None, _ => WordLineJoin.Round };
        shape.LineCap = Get(Svg + "stroke-linecap") switch { "round" => WordLineCap.Round, "square" => WordLineCap.Square, _ => WordLineCap.Flat };
        if (kind == "dash")
        {
            // no (known) draw:stroke-dash: the item pool's dash, XDash() = 1 dot and 1 dash of 20 and distance 20 hundredths of a mm,
            // rectangular (svdattr.cxx:176, include/svx/xdash.hxx:41-43)
            XElement dash = _styles.DrawingStyle("stroke-dash", Get(Draw + "stroke-dash"))
                         ?? new XElement(Draw + "stroke-dash", new XAttribute(Draw + "dots1", "1"), new XAttribute(Draw + "dots1-length", "0.2mm"),
                                         new XAttribute(Draw + "dots2", "1"), new XAttribute(Draw + "dots2-length", "0.2mm"), new XAttribute(Draw + "distance", "0.2mm"));
            if (DashArray(dash, width) is { Length: > 0 } array) shape.Dash = array.Select(d => d / shape.LineWidthPt).ToArray();
        }
        shape.HeadMarker = Marker("start", width);
        shape.TailMarker = Marker("end", width);
    }

    /// <summary>The dash array in points (XDash::CreateDotDashArray).</summary>
    internal static double[] DashArray(XElement dash, double lineWidthPt)
    {
        const double Smallest = 26.95 * 72 / 2540;
        // counts are 16-bit in LibreOffice (XDash sal_uInt16); a pattern of more than 100 repeats one period anyway
        int dots = Math.Max(0, Math.Min(100, int.TryParse((string?)dash.Attribute(Draw + "dots1"), out int d1) ? d1 : 0));
        int dashes = Math.Max(0, Math.Min(100, int.TryParse((string?)dash.Attribute(Draw + "dots2"), out int d2) ? d2 : 0));
        if (dots + dashes == 0) return [];
        string? dotText = (string?)dash.Attribute(Draw + "dots1-length"), dashText = (string?)dash.Attribute(Draw + "dots2-length"), distanceText = (string?)dash.Attribute(Draw + "distance");
        bool relative = (dotText?.Contains('%') ?? false) || (dashText?.Contains('%') ?? false) || (distanceText?.Contains('%') ?? false);
        double Value(string? text, double fallbackPercentOrLength) => text is null ? fallbackPercentOrLength : text.Contains('%') ? (OdtFillPaint.Percent(text) ?? 0) * 100 : Length(text) ?? 0;
        double dotLength = Value(dotText, 0), dashLength = Value(dashText, 0);
        double distance = distanceText is null ? 20 * 72 / 2540.0 : Value(distanceText, 0);   // DashStyle.cxx:68 Distance = 20 (1/100 mm)
        double lineWidth = lineWidthPt == 0 ? Smallest : lineWidthPt;
        if (relative)
        {
            // percentages: a percent value is a fraction of the line width; a length that is not a percentage keeps its value
            double factor = lineWidth / 100.0;
            double Rel(string? text, double raw) => text is not null && text.Contains('%') ? raw * factor : raw;
            if (dashes > 0) dashLength = dashLength != 0 ? Rel(dashText, dashLength) : lineWidth;
            if (dots > 0) dotLength = dotLength != 0 ? Rel(dotText, dotLength) : lineWidth;
            distance = distance != 0 ? Rel(distanceText, distance) : lineWidth;
        }
        else
        {
            if (dashes > 0) dashLength = dashLength != 0 ? Math.Max(dashLength, Smallest) : Math.Max(dashLength, lineWidth);
            if (dots > 0) dotLength = dotLength != 0 ? Math.Max(dotLength, Smallest) : Math.Max(dotLength, lineWidth);
            distance = distance != 0 ? Math.Max(distance, Smallest) : Math.Max(distance, lineWidth);
        }
        var array = new List<double>();
        for (int i = 0; i < dots; i++) { array.Add(dotLength); array.Add(distance); }
        for (int i = 0; i < dashes; i++) { array.Add(dashLength); array.Add(distance); }
        return array.ToArray();
    }

    /// <summary>
    /// A line-end marker (draw:marker, 16.42.8: svg:viewBox + svg:d; MarkerStyle.cxx:60-115 maps the view box to (0, 0)..(w, h))
    /// placed as basegfx createAreaGeometryForLineStartEnd places it (b2dlinegeometry.cxx:38-178): centred in x on the line,
    /// its top at the end point, scaled to draw:marker-*-width (absent: 2 mm; negative = a percentage of the line width,
    /// sdrattributecreator.cxx:338-395), docked at its middle when draw:marker-*-center is true; the line is shortened by the
    /// marker's length (less for a marker with a notched back, :96-140).
    /// </summary>
    private WordLineEndMarker? Marker(string end, double lineWidthPt)
    {
        if (_styles.DrawingStyle("marker", Get(Draw + ("marker-" + end))) is not { } marker) return null;
        double width = Length(Get(Draw + ("marker-" + end + "-width"))) ?? 200 * 72 / 2540.0;
        if (width == 0) return null;
        width = Math.Abs(width);
        bool centred = Get(Draw + ("marker-" + end + "-center")) == "true";
        (double X, double Y, double Width, double Height)? viewBox = OdtSvgPath.ViewBox((string?)marker.Attribute(Svg + "viewBox"));
        List<WordShapeSubpath>? paths = OdtSvgPath.Parse((string?)marker.Attribute(Svg + "d"));
        if (viewBox is null || paths is not { Count: > 0 }) return null;
        var polygons = paths.Select(p => Flatten(p).Select(q => (q.X - viewBox.Value.X, q.Y - viewBox.Value.Y)).ToList()).Where(p => p.Count > 0).ToList();
        if (polygons.Count == 0) return null;
        var all = polygons.SelectMany(p => p).ToList();
        double minX = all.Min(p => p.Item1), maxX = all.Max(p => p.Item1), minY = all.Min(p => p.Item2), maxY = all.Max(p => p.Item2);
        if (maxX - minX <= 0) return null;
        double centreX = (minX + maxX) / 2, scale = width / (maxX - minX), length = (maxY - minY) * scale;
        double docking = centred ? 0.5 : 0;
        List<(double X, double Y)> Place(List<(double, double)> polygon) => polygon.Select(p => ((p.Item1 - centreX) * scale, (p.Item2 - minY) * scale - length * docking)).ToList();
        var placed = polygons.Select(Place).ToList();
        double consumed = length * (1 - docking);
        // a notched back: the centre line leaves the marker before its full length (b2dlinegeometry.cxx:96-140)
        var crossings = new List<double>();
        foreach (var polygon in placed)
            for (int i = 0; i < polygon.Count; i++)
            {
                var p1 = polygon[i]; var p2 = polygon[(i + 1) % polygon.Count];
                if ((p1.X < 0 && p2.X > 0) || (p1.X > 0 && p2.X < 0)) crossings.Add(p1.Y + (-p1.X / (p2.X - p1.X)) * (p2.Y - p1.Y));
                else if (p1.X == 0 && p2.X != 0)
                {
                    var prior = polygon[(i + polygon.Count - 1) % polygon.Count];
                    if ((prior.X < 0 && p2.X > 0) || (prior.X > 0 && p2.X < 0)) crossings.Add(p1.Y);
                }
            }
        crossings.Sort();
        if (crossings.Count(y => y < 1e-9) % 2 != 0)
            foreach (double y in crossings)
                if (y > 1e-9) { if (y < consumed) consumed = y; break; }
        return new WordLineEndMarker(placed[0], width, consumed) { MorePolygons = placed.Skip(1).ToList() };
    }

    internal static List<(double X, double Y)> Flatten(WordShapeSubpath subpath)
    {
        var points = new List<(double X, double Y)> { (subpath.StartX, subpath.StartY) };
        foreach (WordShapeSegment segment in subpath.Segments)
        {
            if (!segment.IsCurve) { points.Add((segment.X, segment.Y)); continue; }
            (double x0, double y0) = points[points.Count - 1];
            for (int step = 1; step <= 16; step++)
            {
                double t = step / 16.0, u = 1 - t;
                points.Add((u * u * u * x0 + 3 * u * u * t * segment.Control1X + 3 * u * t * t * segment.Control2X + t * t * t * segment.X,
                            u * u * u * y0 + 3 * u * u * t * segment.Control1Y + 3 * u * t * t * segment.Control2Y + t * t * t * segment.Y));
            }
        }
        return points;
    }
}
