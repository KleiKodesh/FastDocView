using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>The outline a drawn shape takes (format-neutral). <see cref="Box"/> = a preset this library does not draw: a plain box of the right size.</summary>
public enum WordShapeGeometry : byte { Rectangle, RoundRectangle, Ellipse, Line, Triangle, RightArrow, LeftArrow, UpArrow, DownArrow, Path, Box, Picture }

/// <summary>One segment of a <see cref="WordShapeSubpath"/>: a line to (X, Y), or a cubic Bezier through the two control points.</summary>
public readonly record struct WordShapeSegment(bool IsCurve, double Control1X, double Control1Y, double Control2X, double Control2Y, double X, double Y)
{
    public static WordShapeSegment Line(double x, double y) => new(false, 0, 0, 0, 0, x, y);
}

/// <summary>A run of segments from a start point; closed = its end joins its start.</summary>
public sealed class WordShapeSubpath
{
    public double StartX, StartY;
    public readonly List<WordShapeSegment> Segments = new();
    public bool Closed;
}

/// <summary>How a path is filled (ECMA-376 ST_PathFillMode): the shape's fill, none, or the fill darkened / lightened.</summary>
public enum WordShapePathFill : byte { Normal, None, Darken, DarkenLess, Lighten, LightenLess }

/// <summary>One drawn path of a shape's outline (DrawingML a:path): its subpaths, how it is filled, whether it is stroked.</summary>
public sealed class WordShapePath
{
    public readonly List<WordShapeSubpath> Subpaths = new();
    public WordShapePathFill Fill;
    public bool Stroke = true;
}

/// <summary>A line end (DrawingML a:headEnd / a:tailEnd, ST_LineEndType) with its width and length (0 small, 1 medium, 2 large).</summary>
public enum WordLineEndType : byte { None, Triangle, Stealth, Diamond, Oval, Arrow }

public readonly record struct WordLineEnd(WordLineEndType Type, int Width, int Length);

/// <summary>Line cap (DrawingML a:ln cap: flat, rnd, sq).</summary>
public enum WordLineCap : byte { Flat, Round, Square }

/// <summary>
/// One drawn shape inside a drawing's frame: its box in points from the frame's top left, rotation (degrees clockwise about
/// its centre) and flips, outline, fill (a CSS colour, or gradient stops), line, arrow ends, and which text box (in the
/// drawing's order) it holds (-1: none) or which picture (by its <see cref="WordPictureHtml.PictureIdAttribute"/> id).
/// </summary>
public sealed class WordShape
{
    public WordShapeGeometry Geometry;
    public double X, Y, Width, Height;
    public double RotationDegrees;
    public bool FlipHorizontal, FlipVertical;
    /// <summary>Preset adjust values (ECMA-376 a:avLst gd "val n", 100000ths); empty = the preset's defaults.</summary>
    public double[] Adjust = [];
    /// <summary>A <see cref="WordShapeGeometry.Path"/> outline as SVG path data in the shape's own points (0..Width, 0..Height).</summary>
    public string? PathData;
    public string? FillColor;
    public List<(double Offset, string Color)>? GradientStops;
    /// <summary>Linear gradient direction in degrees clockwise from left-to-right (DrawingML a:lin ang).</summary>
    public double GradientAngleDegrees;
    public string? LineColor;
    /// <summary>Fill / line opacity 0..1 (1 = opaque; OfficeArt fillOpacity / lineOpacity 16.16).</summary>
    public double FillOpacity = 1, LineOpacity = 1;
    public double LineWidthPt;
    /// <summary>SVG stroke-dasharray in multiples of the line width (null = solid).</summary>
    public double[]? Dash;
    public bool HeadArrow, TailArrow;
    /// <summary>The evaluated outline (DrawingML preset or custom geometry) in the shape's points; when set it is drawn instead of <see cref="Geometry"/>.</summary>
    public List<WordShapePath>? Paths;
    /// <summary>Line ends at the start (head) and end (tail) of every open subpath of <see cref="Paths"/>.</summary>
    public WordLineEnd HeadEnd, TailEnd;
    /// <summary>A line end drawn by its own format's geometry (OfficeArt binary arrows) instead of <see cref="WordLineEndGeometry"/>.</summary>
    internal WordLineEndMarker? HeadMarker, TailMarker;
    public WordLineCap LineCap;
    /// <summary>Per stop of <see cref="GradientStops"/> (null = all opaque).</summary>
    public List<double>? GradientStopOpacities;
    /// <summary>A radial gradient (DrawingML a:path circle) centred at (<see cref="GradientCenterX"/>, <see cref="GradientCenterY"/>) as fractions of the box.</summary>
    public bool GradientRadial;
    public double GradientCenterX = 0.5, GradientCenterY = 0.5;
    public int TextIndex = -1;
    public string? PictureId;
    public WordTextBoxStyle TextStyle;
}

/// <summary>
/// Draws <see cref="WordShape"/>s: a frame placed like a picture (<see cref="WordPictureHtml.AppendPlacement"/>), one inline
/// &lt;svg&gt; filling it with every outline in frame points, and each shape's text box / picture as an absolutely placed box over its
/// outline. A frame of one shape whose text grows (auto-fit) keeps its text in the flow instead, so the frame grows and the
/// outline (drawn with non-scaling strokes) stretches with it.
/// </summary>
public static class WordShapeHtml
{
    public static void OpenFrame(StringBuilder html, WordPictureLayout layout, bool growsWithText)
    {
        var css = new StringBuilder("position:relative;box-sizing:border-box;");
        css.Append(layout.Wrap == WordPictureWrap.Inline ? "display:inline-block;vertical-align:bottom;" : "");
        css.Append("width:").Append(Points(layout.WidthPt)).Append("pt;");
        css.Append(growsWithText ? "min-height:" : "height:").Append(Points(layout.HeightPt)).Append("pt;");
        WordPictureHtml.AppendPlacement(css, layout);
        html.Append("<span class=\"word-shape\" style=\"").Append(css).Append("\">");
    }

    public static void CloseFrame(StringBuilder html) => html.Append("</span>");

    /// <summary>The outlines of every shape (pictures are drawn as boxes by <see cref="AppendPicture"/>).</summary>
    public static void AppendGeometry(StringBuilder html, IReadOnlyList<WordShape> shapes, double width, double height, bool stretches)
    {
        width = Math.Max(width, 0.01); height = Math.Max(height, 0.01);
        html.Append("<svg class=\"word-shape-geometry\" aria-hidden=\"true\" viewBox=\"0 0 ").Append(Points(width)).Append(' ').Append(Points(height))
            .Append("\" width=\"").Append(Points(width)).Append("pt\" height=\"").Append(Points(height)).Append("pt\"")
            .Append(stretches ? " preserveAspectRatio=\"none\"" : "")
            .Append(" style=\"position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none\">");
        foreach (WordShape shape in shapes)
            if (shape.Geometry != WordShapeGeometry.Picture) AppendShape(html, shape, stretches);
        html.Append("</svg>");
    }

    private static void AppendShape(StringBuilder html, WordShape shape, bool stretches)
    {
        double w = shape.Width, h = shape.Height;
        html.Append("<g transform=\"translate(").Append(Number(shape.X + w / 2)).Append(' ').Append(Number(shape.Y + h / 2)).Append(')');
        if (shape.RotationDegrees != 0) html.Append(" rotate(").Append(Number(shape.RotationDegrees)).Append(')');
        if (shape.FlipHorizontal || shape.FlipVertical) html.Append(" scale(").Append(shape.FlipHorizontal ? "-1" : "1").Append(' ').Append(shape.FlipVertical ? "-1" : "1").Append(')');
        html.Append(" translate(").Append(Number(-w / 2)).Append(' ').Append(Number(-h / 2)).Append(")\">");
        if (shape.Paths is not null)
        {
            AppendPaths(html, shape, stretches);
            html.Append("</g>");
            return;
        }
        string? gradientId = AppendGradient(html, shape, 0);
        bool isLine = shape.Geometry == WordShapeGeometry.Line;
        string fill = isLine ? "none" : gradientId is not null ? "url(#" + gradientId + ")" : shape.FillColor ?? "none";
        var paint = new StringBuilder(" fill=\"").Append(fill).Append('"');
        if (shape.FillOpacity < 1 && !isLine) paint.Append(" fill-opacity=\"").Append(Number(Math.Max(0, shape.FillOpacity))).Append('"');
        if (shape.LineOpacity < 1 && shape.LineColor is not null) paint.Append(" stroke-opacity=\"").Append(Number(Math.Max(0, shape.LineOpacity))).Append('"');
        double lineWidth = Math.Max(0.25, shape.LineWidthPt);
        if (shape.LineColor is { } line)
        {
            paint.Append(" stroke=\"").Append(line).Append("\" stroke-width=\"").Append(Number(lineWidth)).Append('"');
            if (shape.Dash is { Length: > 0 } dash) paint.Append(" stroke-dasharray=\"").Append(string.Join(" ", dash.Select(d => Number(d * lineWidth)))).Append('"');
        }
        else paint.Append(" stroke=\"none\"");
        if (stretches) paint.Append(" vector-effect=\"non-scaling-stroke\"");
        switch (shape.Geometry)
        {
            case WordShapeGeometry.Rectangle or WordShapeGeometry.Box:
                html.Append("<rect width=\"").Append(Number(w)).Append("\" height=\"").Append(Number(h)).Append('"').Append(paint).Append("/>");
                break;
            case WordShapeGeometry.RoundRectangle:
            {
                // presetShapeDefinitions.xml roundRect: a = pin 0 adj 50000 (default 16667), corner radius ss * a / 100000
                double radius = Math.Min(w, h) * Clamp(Adjust(shape, 0, 16667), 0, 50000) / 100000;
                html.Append("<rect width=\"").Append(Number(w)).Append("\" height=\"").Append(Number(h)).Append("\" rx=\"").Append(Number(radius)).Append("\" ry=\"").Append(Number(radius)).Append('"').Append(paint).Append("/>");
                break;
            }
            case WordShapeGeometry.Ellipse:
                html.Append("<ellipse cx=\"").Append(Number(w / 2)).Append("\" cy=\"").Append(Number(h / 2)).Append("\" rx=\"").Append(Number(w / 2)).Append("\" ry=\"").Append(Number(h / 2)).Append('"').Append(paint).Append("/>");
                break;
            case WordShapeGeometry.Line:
                html.Append("<path d=\"M0 0L").Append(Number(w)).Append(' ').Append(Number(h)).Append('"').Append(paint).Append("/>");
                if (shape.LineColor is { } color)
                {
                    if (shape.HeadArrow) AppendArrowHead(html, w, h, 0, 0, lineWidth, color);
                    if (shape.TailArrow) AppendArrowHead(html, 0, 0, w, h, lineWidth, color);
                }
                break;
            default:
                html.Append("<path d=\"").Append(PresetPath(shape)).Append('"').Append(paint).Append("/>");
                break;
        }
        html.Append("</g>");
    }

    /// <summary>A filled triangle at (<paramref name="toX"/>, <paramref name="toY"/>) pointing away from (fromX, fromY), three line widths long.</summary>
    private static void AppendArrowHead(StringBuilder html, double fromX, double fromY, double toX, double toY, double lineWidth, string color)
    {
        double dx = toX - fromX, dy = toY - fromY, length = Math.Sqrt(dx * dx + dy * dy);
        if (length == 0) return;
        double ux = dx / length, uy = dy / length, size = Math.Max(3 * lineWidth, 3);
        double baseX = toX - ux * size, baseY = toY - uy * size;
        html.Append("<path d=\"M").Append(Number(toX)).Append(' ').Append(Number(toY))
            .Append('L').Append(Number(baseX - uy * size / 2)).Append(' ').Append(Number(baseY + ux * size / 2))
            .Append('L').Append(Number(baseX + uy * size / 2)).Append(' ').Append(Number(baseY - ux * size / 2))
            .Append("Z\" fill=\"").Append(color).Append("\" stroke=\"none\"/>");
    }

    /// <summary>
    /// The shape's gradient as an SVG definition in its own points; null when it has none. <paramref name="brightness"/> is a
    /// path's darken / lighten (applied to every stop as LibreOffice AdaptObjColor does, svx EnhancedCustomShape2d.cxx:2829-2849).
    /// Linear: LibreOffice init1DGradientInfo (basegfx/source/tools/gradienttools.cxx:52-121) spans the box grown to hold the
    /// rotated gradient, so the first and last stop touch the box's first and last corner along the direction: length
    /// w |cos| + h |sin| through the centre. Radial (a:path circle): initEllipticalGradientInfo (:127-190) with bCircular, a
    /// circle of the box's diagonal centred at the fillToRect centre; pos 0 is the centre (the stops are reversed with the
    /// radial alpha 1 - distance, oox fillproperties.cxx:528-551).
    /// </summary>
    private static string? AppendGradient(StringBuilder html, WordShape shape, double brightness)
    {
        if (shape.GradientStops is not { Count: > 0 } stops) return null;
        double w = shape.Width, h = shape.Height;
        var stopsHtml = new StringBuilder();
        for (int i = 0; i < stops.Count; i++)
        {
            double opacity = shape.GradientStopOpacities is { } opacities && i < opacities.Count ? opacities[i] : 1;
            stopsHtml.Append("<stop offset=\"").Append(Number(stops[i].Offset)).Append("\" stop-color=\"").Append(WordShapeColor.Brightened(stops[i].Color, brightness)).Append('"');
            if (opacity < 1) stopsHtml.Append(" stop-opacity=\"").Append(Opacity(opacity)).Append('"');
            stopsHtml.Append("/>");
        }
        var definition = new StringBuilder();
        if (shape.GradientRadial)
        {
            double radius = Math.Sqrt(w * w + h * h) / 2;
            definition.Append("<radialGradient gradientUnits=\"userSpaceOnUse\" cx=\"").Append(Number(shape.GradientCenterX * w)).Append("\" cy=\"").Append(Number(shape.GradientCenterY * h))
                      .Append("\" r=\"").Append(Number(radius)).Append("\">").Append(stopsHtml).Append("</radialGradient>");
        }
        else
        {
            double radians = shape.GradientAngleDegrees * Math.PI / 180, cos = Math.Cos(radians), sin = Math.Sin(radians);
            double half = (w * Math.Abs(cos) + h * Math.Abs(sin)) / 2;
            definition.Append("<linearGradient gradientUnits=\"userSpaceOnUse\" x1=\"").Append(Number(w / 2 - cos * half)).Append("\" y1=\"").Append(Number(h / 2 - sin * half))
                      .Append("\" x2=\"").Append(Number(w / 2 + cos * half)).Append("\" y2=\"").Append(Number(h / 2 + sin * half)).Append("\">").Append(stopsHtml).Append("</linearGradient>");
        }
        // the id is the gradient's own content: two blocks that both define it define the same thing
        string id = "word-shape-gradient-" + StableHash(definition.ToString());
        int open = definition.ToString().IndexOf(' ');
        definition.Insert(open, " id=\"" + id + "\"");
        html.Append("<defs>").Append(definition).Append("</defs>");
        return id;
    }

    private static string Opacity(double value) => Math.Round(Math.Max(0, Math.Min(1, value)), 3).ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>
    /// Draws an evaluated outline the way LibreOffice draws a custom shape (svx EnhancedCustomShape2d.cxx CreateSubPath
    /// :2078-2698): each a:path is one object; its fill (unless fill="none") closes every subpath and fills them even-odd
    /// (vcl/skia/gdiimpl.cxx:658 kEvenOdd), with the darken / lighten brightness of its fill mode (:2116-2133, GetColorData
    /// :1095-1118); its stroke (unless stroke="0") draws the subpaths as they are; line ends go on every open subpath
    /// (each is its own polyline object carrying the shape's line start / end).
    /// </summary>
    private static void AppendPaths(StringBuilder html, WordShape shape, bool stretches)
    {
        double lineWidth = Math.Max(0.25, shape.LineWidthPt);
        foreach (WordShapePath path in shape.Paths!)
        {
            if (path.Subpaths.Count == 0) continue;
            double brightness = path.Fill switch
            {
                WordShapePathFill.Darken => -0.4, WordShapePathFill.DarkenLess => -0.2,
                WordShapePathFill.Lighten => 0.4, WordShapePathFill.LightenLess => 0.2, _ => 0,
            };
            if (path.Fill != WordShapePathFill.None && (shape.FillColor is not null || shape.GradientStops is { Count: > 0 }))
            {
                string? gradientId = AppendGradient(html, shape, brightness);
                html.Append("<path d=\"").Append(PathData(path.Subpaths, closeAll: true)).Append("\" fill-rule=\"evenodd\" fill=\"")
                    .Append(gradientId is not null ? "url(#" + gradientId + ")" : WordShapeColor.Brightened(shape.FillColor!, brightness)).Append('"');
                if (gradientId is null && shape.FillOpacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(shape.FillOpacity)).Append('"');
                html.Append(" stroke=\"none\"/>");
            }
            if (!path.Stroke || shape.LineColor is not { } line) continue;
            var stroke = new StringBuilder(" fill=\"none\" stroke=\"").Append(line).Append("\" stroke-width=\"").Append(Number(lineWidth)).Append('"');
            if (shape.LineOpacity < 1) stroke.Append(" stroke-opacity=\"").Append(Opacity(shape.LineOpacity)).Append('"');
            if (shape.LineCap != WordLineCap.Flat) stroke.Append(" stroke-linecap=\"").Append(shape.LineCap == WordLineCap.Round ? "round" : "square").Append('"');
            if (shape.Dash is { Length: > 0 } dash) stroke.Append(" stroke-dasharray=\"").Append(string.Join(" ", dash.Select(d => Number(Math.Max(0, d) * lineWidth)))).Append('"');
            if (stretches) stroke.Append(" vector-effect=\"non-scaling-stroke\"");
            bool hasEnds = shape.HeadEnd.Type != WordLineEndType.None || shape.TailEnd.Type != WordLineEndType.None;
            var plain = new List<WordShapeSubpath>();
            foreach (WordShapeSubpath subpath in path.Subpaths)
            {
                if (!hasEnds || subpath.Closed || subpath.Segments.Count == 0) { plain.Add(subpath); continue; }
                AppendWithLineEnds(html, subpath, shape, stroke.ToString(), line);
            }
            if (plain.Count > 0) html.Append("<path d=\"").Append(PathData(plain, closeAll: false)).Append('"').Append(stroke).Append("/>");
        }
    }

    private static string PathData(List<WordShapeSubpath> subpaths, bool closeAll)
    {
        var d = new StringBuilder();
        foreach (WordShapeSubpath subpath in subpaths)
        {
            d.Append('M').Append(Number(subpath.StartX)).Append(' ').Append(Number(subpath.StartY));
            foreach (WordShapeSegment segment in subpath.Segments)
            {
                if (segment.IsCurve)
                    d.Append('C').Append(Number(segment.Control1X)).Append(' ').Append(Number(segment.Control1Y)).Append(' ').Append(Number(segment.Control2X)).Append(' ')
                     .Append(Number(segment.Control2Y)).Append(' ').Append(Number(segment.X)).Append(' ').Append(Number(segment.Y));
                else d.Append('L').Append(Number(segment.X)).Append(' ').Append(Number(segment.Y));
            }
            if (closeAll || subpath.Closed) d.Append('Z');
        }
        return d.ToString();
    }

    /// <summary>
    /// An open subpath with its line ends, as LibreOffice draws them (drawinglayer polygonprimitive2d.cxx:704-770
    /// PolygonStrokeArrowPrimitive2D, basegfx b2dlinegeometry.cxx:38-178 createAreaGeometryForLineStartEnd): each end's
    /// marker is scaled to its width, its tip on the end point (centred markers, diamond and oval, straddle it), turned along
    /// the line toward the point the marker's length back; the shaft is cut back by that length less width / 15 of overlap.
    /// Curves are flattened for the cut (LibreOffice measures the same polygon).
    /// </summary>
    private static void AppendWithLineEnds(StringBuilder html, WordShapeSubpath subpath, WordShape shape, string stroke, string color)
    {
        List<(double X, double Y)> points = Flatten(subpath);
        var cumulative = new double[points.Count];
        for (int i = 1; i < points.Count; i++) cumulative[i] = cumulative[i - 1] + Distance(points[i - 1], points[i]);
        double length = cumulative[points.Count - 1];
        if (length <= 0) return;
        double start = 0, end = 0, startOverlap = 0, endOverlap = 0;
        var markers = new StringBuilder();
        if ((shape.HeadMarker ?? WordLineEndGeometry.Marker(shape.HeadEnd, shape.LineWidthPt)) is { } head)
        {
            start = head.Consumed;
            startOverlap = head.Width / 15;
            AppendMarker(markers, head, points[0], PointAt(points, cumulative, Math.Min(start, length)), color, shape.LineOpacity);
        }
        if ((shape.TailMarker ?? WordLineEndGeometry.Marker(shape.TailEnd, shape.LineWidthPt)) is { } tail)
        {
            end = tail.Consumed;
            endOverlap = tail.Width / 15;
            AppendMarker(markers, tail, points[points.Count - 1], PointAt(points, cumulative, Math.Max(0, length - end)), color, shape.LineOpacity);
        }
        double from = Math.Max(0, start - startOverlap), to = Math.Min(length, length - end + endOverlap);
        if (to > from)
        {
            var d = new StringBuilder();
            (double x, double y) first = PointAt(points, cumulative, from);
            d.Append('M').Append(Number(first.x)).Append(' ').Append(Number(first.y));
            for (int i = 1; i < points.Count; i++)
                if (cumulative[i] > from && cumulative[i] < to) d.Append('L').Append(Number(points[i].X)).Append(' ').Append(Number(points[i].Y));
            (double x, double y) last = PointAt(points, cumulative, to);
            d.Append('L').Append(Number(last.x)).Append(' ').Append(Number(last.y));
            html.Append("<path d=\"").Append(d).Append('"').Append(stroke).Append("/>");
        }
        html.Append(markers);
    }

    private static void AppendMarker(StringBuilder html, WordLineEndMarker marker, (double X, double Y) head, (double X, double Y) toward, string color, double opacity)
    {
        double rotation = Math.Atan2(head.Y - toward.Y, head.X - toward.X) + Math.PI / 2, cos = Math.Cos(rotation), sin = Math.Sin(rotation);
        var d = new StringBuilder();
        for (int i = 0; i < marker.Points.Count; i++)
        {
            (double x, double y) = marker.Points[i];
            d.Append(i == 0 ? 'M' : 'L').Append(Number(head.X + x * cos - y * sin)).Append(' ').Append(Number(head.Y + x * sin + y * cos));
        }
        html.Append("<path d=\"").Append(d).Append("Z\" fill=\"").Append(color).Append('"');
        if (opacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(opacity)).Append('"');
        html.Append(" stroke=\"none\"/>");
    }

    private static double Distance((double X, double Y) a, (double X, double Y) b) => Math.Sqrt((b.X - a.X) * (b.X - a.X) + (b.Y - a.Y) * (b.Y - a.Y));

    private static (double X, double Y) PointAt(List<(double X, double Y)> points, double[] cumulative, double distance)
    {
        for (int i = 1; i < points.Count; i++)
        {
            if (cumulative[i] < distance && i < points.Count - 1) continue;
            double span = cumulative[i] - cumulative[i - 1], t = span > 0 ? Math.Max(0, Math.Min(1, (distance - cumulative[i - 1]) / span)) : 0;
            return (points[i - 1].X + (points[i].X - points[i - 1].X) * t, points[i - 1].Y + (points[i].Y - points[i - 1].Y) * t);
        }
        return points[0];
    }

    /// <summary>The subpath as a polyline, each cubic split into 32 chords, repeated points dropped (removeDoublePoints, polygonprimitive2d.cxx:709).</summary>
    private static List<(double X, double Y)> Flatten(WordShapeSubpath subpath)
    {
        var points = new List<(double X, double Y)> { (subpath.StartX, subpath.StartY) };
        void Add(double x, double y) { if (points[points.Count - 1] != (x, y)) points.Add((x, y)); }
        foreach (WordShapeSegment segment in subpath.Segments)
        {
            if (!segment.IsCurve) { Add(segment.X, segment.Y); continue; }
            (double x0, double y0) = points[points.Count - 1];
            for (int step = 1; step <= 32; step++)
            {
                double t = step / 32.0, u = 1 - t;
                Add(u * u * u * x0 + 3 * u * u * t * segment.Control1X + 3 * u * t * t * segment.Control2X + t * t * t * segment.X,
                    u * u * u * y0 + 3 * u * u * t * segment.Control1Y + 3 * u * t * t * segment.Control2Y + t * t * t * segment.Y);
            }
        }
        return points;
    }

    /// <summary>
    /// Outlines of the presets drawn, transcribed from LibreOffice oox/source/drawingml/customshapes/presetShapeDefinitions.xml
    /// (the ECMA-376 presetShapeDefinitions.xml): triangle (line 18658), rightArrow (15317), leftArrow (10866), upArrow
    /// (18914), downArrow (6567), with their default adjust values and pins.
    /// </summary>
    private static string PresetPath(WordShape shape)
    {
        double w = shape.Width, h = shape.Height, ss = Math.Min(w, h), hc = w / 2, vc = h / 2;
        if (ss <= 0) ss = Math.Max(w, h);
        var d = new StringBuilder();
        void Move(double x, double y) => d.Append('M').Append(Number(x)).Append(' ').Append(Number(y));
        void Line(double x, double y) => d.Append('L').Append(Number(x)).Append(' ').Append(Number(y));
        switch (shape.Geometry)
        {
            case WordShapeGeometry.Path:
                return shape.PathData ?? "";
            case WordShapeGeometry.Triangle:
            {
                double x2 = w * Clamp(Adjust(shape, 0, 50000), 0, 100000) / 100000;
                Move(0, h); Line(x2, 0); Line(w, h); d.Append('Z');
                break;
            }
            case WordShapeGeometry.RightArrow or WordShapeGeometry.LeftArrow:
            {
                double a1 = Clamp(Adjust(shape, 0, 50000), 0, 100000), a2 = Clamp(Adjust(shape, 1, 50000), 0, 100000 * w / ss);
                double head = ss * a2 / 100000, dy1 = h * a1 / 200000, y1 = vc - dy1, y2 = vc + dy1;
                if (shape.Geometry == WordShapeGeometry.RightArrow) { double x1 = w - head; Move(0, y1); Line(x1, y1); Line(x1, 0); Line(w, vc); Line(x1, h); Line(x1, y2); Line(0, y2); }
                else { double x2 = head; Move(0, vc); Line(x2, 0); Line(x2, y1); Line(w, y1); Line(w, y2); Line(x2, y2); Line(x2, h); }
                d.Append('Z');
                break;
            }
            case WordShapeGeometry.UpArrow or WordShapeGeometry.DownArrow:
            {
                double a1 = Clamp(Adjust(shape, 0, 50000), 0, 100000), a2 = Clamp(Adjust(shape, 1, 50000), 0, 100000 * h / ss);
                double head = ss * a2 / 100000, dx1 = w * a1 / 200000, x1 = hc - dx1, x2 = hc + dx1;
                if (shape.Geometry == WordShapeGeometry.UpArrow) { double y2 = head; Move(0, y2); Line(hc, 0); Line(w, y2); Line(x2, y2); Line(x2, h); Line(x1, h); Line(x1, y2); }
                else { double y1 = h - head; Move(0, y1); Line(x1, y1); Line(x1, 0); Line(x2, 0); Line(x2, y1); Line(w, y1); Line(hc, h); }
                d.Append('Z');
                break;
            }
        }
        return d.ToString();
    }

    /// <summary>Opens the box a shape's text sits in: over the shape's outline, or (a growing single shape) in the frame's flow.</summary>
    public static void OpenShapeText(StringBuilder html, WordShape shape, bool inFlow)
    {
        WordTextBoxStyle style = shape.TextStyle;
        var css = new StringBuilder("box-sizing:border-box;overflow-wrap:break-word;display:block;margin:0;border:0;");
        if (inFlow) css.Append("position:relative;width:100%;");
        else
        {
            css.Append("position:absolute;left:").Append(Points(shape.X)).Append("pt;top:").Append(Points(shape.Y)).Append("pt;width:").Append(Points(shape.Width))
               .Append("pt;height:").Append(Points(shape.Height)).Append("pt;overflow:hidden;");
            if (shape.RotationDegrees != 0) css.Append("transform:rotate(").Append(Number(shape.RotationDegrees)).Append("deg);");
        }
        css.Append("padding:").Append(Points(style.TopInsetPt)).Append("pt ").Append(Points(style.RightInsetPt)).Append("pt ")
           .Append(Points(style.BottomInsetPt)).Append("pt ").Append(Points(style.LeftInsetPt)).Append("pt;");
        if (style.WritingMode is { } writingMode) css.Append(writingMode);
        html.Append("<span class=\"text-box word-shape-text\" style=\"").Append(css).Append("\">");
    }

    public static void CloseShapeText(StringBuilder html) => html.Append("</span>");

    /// <summary>A picture inside a group, at its box over the frame.</summary>
    public static void AppendPicture(StringBuilder html, WordShape shape)
    {
        html.Append("<span class=\"word-shape-part\" style=\"position:absolute;display:block;left:").Append(Points(shape.X)).Append("pt;top:").Append(Points(shape.Y)).Append("pt;");
        if (shape.RotationDegrees != 0) html.Append("transform:rotate(").Append(Number(shape.RotationDegrees)).Append("deg);");
        html.Append("\">");
        WordPictureHtml.Append(html, new WordPictureLayout(shape.Width, shape.Height, WordPictureWrap.Inline, WordPictureSide.Left, 0, 0, 0, 0), shape.PictureId);
        html.Append("</span>");
    }

    private static double Adjust(WordShape shape, int index, double fallback) => index < shape.Adjust.Length ? shape.Adjust[index] : fallback;

    private static double Clamp(double value, double min, double max) => value < min ? min : value > max ? Math.Max(min, max) : value;

    private static string StableHash(string text)
    {
        uint hash = 2166136261;
        foreach (char c in text) { hash ^= c; hash *= 16777619; }
        return hash.ToString("x8", CultureInfo.InvariantCulture);
    }

    internal static string Number(double value) { double rounded = Math.Round(value, 2); return (rounded == 0 ? 0 : rounded).ToString("0.##", CultureInfo.InvariantCulture); }

    private static string Points(double value) => Math.Max(0, value).ToString("0.##", CultureInfo.InvariantCulture);
}

/// <summary>A line-end marker ready to place: its outline with the tip (or, centred, its middle) at (0, 0) and its body along +y, in points.</summary>
internal sealed record WordLineEndMarker(List<(double X, double Y)> Points, double Width, double Consumed);

/// <summary>
/// Line-end markers as LibreOffice builds them for DrawingML (oox/source/drawingml/lineproperties.cxx:223-367,
/// lclPushMarkerProperties): sm / med / lg give width and length factors 2 / 3 / 5 line widths (2.5 / 3.5 / 5.5 for the open
/// "arrow"), the marker width is the width factor times the line width but at least 70 hundredths of a millimetre
/// (:302-303); the outline is given in percent of (width, length) per type: triangle, open arrow (a stem of the line's
/// width, :328-338), stealth (notch at 60 %), diamond, oval (12 points); diamond and oval are centred on the end
/// (bMarkerCenter). Placement: basegfx b2dlinegeometry.cxx:74-94 (centre in x, scale to the width, docking 0.5 for centred).
/// </summary>
internal static class WordLineEndGeometry
{
    private const double MinimumWidthPt = 0.7 / 25.4 * 72;   // 70 hundredths of a millimetre

    public static WordLineEndMarker? Marker(WordLineEnd end, double lineWidthPt)
    {
        if (end.Type == WordLineEndType.None) return null;
        bool isArrow = end.Type == WordLineEndType.Arrow;
        double Factor(int size) => size switch { 0 => isArrow ? 2.5 : 2.0, 2 => isArrow ? 5.5 : 5.0, _ => isArrow ? 3.5 : 3.0 };
        double widthFactor = Factor(end.Width), lengthFactor = Factor(end.Length);
        double lineWidth = Math.Max(0, lineWidthPt);
        double markerWidth = widthFactor * Math.Max(lineWidth, MinimumWidthPt);
        double half = Math.Max(100.0 * 0.5 * lineWidth / markerWidth, 1);
        (double, double)[] percent = end.Type switch
        {
            WordLineEndType.Triangle => [(50, 0), (100, 100), (0, 100), (50, 0)],
            WordLineEndType.Arrow => [(50, 0), (100, 100 - half * 1.5), (100 - half * 1.5, 100), (50.0 + half, 5.5 * half), (50.0 + half, 100), (50.0 - half, 100),
                                      (50.0 - half, 5.5 * half), (half * 1.5, 100), (0, 100 - half * 1.5), (50, 0)],
            WordLineEndType.Stealth => [(50, 0), (100, 100), (50, 60), (0, 100), (50, 0)],
            WordLineEndType.Diamond => [(50, 0), (100, 50), (50, 100), (0, 50), (50, 0)],
            _ => [(50, 0), (75, 7), (93, 25), (100, 50), (93, 75), (75, 93), (50, 100), (25, 93), (7, 75), (0, 50), (7, 25), (25, 7), (50, 0)],
        };
        // OOX_ARROW_POINT truncates to integers in its own units (fArrowWidth * x, fArrowLength * y)
        var raw = percent.Select(p => ((double)(int)(widthFactor * p.Item1), (double)(int)(lengthFactor * p.Item2))).ToList();
        double minX = raw.Min(p => p.Item1), maxX = raw.Max(p => p.Item1), minY = raw.Min(p => p.Item2), maxY = raw.Max(p => p.Item2);
        if (maxX - minX <= 0) return null;
        double scale = markerWidth / (maxX - minX), centreX = (minX + maxX) / 2, length = (maxY - minY) * scale;
        bool centred = end.Type is WordLineEndType.Diamond or WordLineEndType.Oval;
        double docking = centred ? 0.5 : 0;
        var points = raw.Select(p => ((p.Item1 - centreX) * scale, (p.Item2 - minY) * scale - length * docking)).ToList();
        return new WordLineEndMarker(points, markerWidth, length * (1 - docking));
    }
}

/// <summary>Colour arithmetic the drawing needs on "#RRGGBB" values.</summary>
public static class WordShapeColor
{
    /// <summary>LibreOffice EnhancedCustomShape2d::GetColorData (svx EnhancedCustomShape2d.cxx:1095-1118): blend toward white by
    /// a positive brightness, toward black by a negative one (darken -0.4, darkenLess -0.2, lighten 0.4, lightenLess 0.2, :2122-2133).</summary>
    public static string Brightened(string css, double brightness)
    {
        if (brightness == 0 || css.Length != 7 || css[0] != '#'
            || !int.TryParse(css.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb)) return css;
        int Channel(int c) => (int)Math.Max(0, Math.Min(255, brightness >= 0 ? c * (1.0 - brightness) + brightness * 255.0 : c * (1.0 + brightness)));
        return "#" + ((Channel((rgb >> 16) & 0xFF) << 16) | (Channel((rgb >> 8) & 0xFF) << 8) | Channel(rgb & 0xFF)).ToString("X6", CultureInfo.InvariantCulture);
    }
}
