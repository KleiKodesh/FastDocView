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
    /// <summary>LibreOffice's path brightness (ODF H / I / J / K, OOXML darken / lighten): -0.4 .. 0.4 (GetColorData :1095-1118); 0 = none.</summary>
    public double Brightness;
    /// <summary>The OpenOffice colour data's luminance change in percent for this object (GetLuminanceChange :1082-1093, applied in HSV :1120-1149); 0 = none.</summary>
    public int LuminancePercent;
    /// <summary>The fill half of a sub-path group LibreOffice draws as a fill and a stroke object; a text effect lays its text on the other half only.</summary>
    internal bool IsFillCopy;
}

/// <summary>A line end (DrawingML a:headEnd / a:tailEnd, ST_LineEndType) with its width and length (0 small, 1 medium, 2 large).</summary>
public enum WordLineEndType : byte { None, Triangle, Stealth, Diamond, Oval, Arrow }

public readonly record struct WordLineEnd(WordLineEndType Type, int Width, int Length);

/// <summary>A preset pattern: foreground colour and opacity, the 8 rows of its 8 x 8 cell (bit 0x80 = the left pixel).</summary>
public sealed record WordShapePattern(string Color, double Opacity, byte[] Rows);

/// <summary>
/// How a picture fills a shape: stretched into the box inset by (Left, Top, Right, Bottom) fractions (negative = beyond), or
/// tiled from an offset (points) at a tile size (points), mirrored per <see cref="TileFlip"/> ("x", "y", "xy").
/// </summary>
public readonly record struct WordShapePictureFill(double Left, double Top, double Right, double Bottom, bool Tiled, double TileOffsetX, double TileOffsetY,
                                                   double TileWidth, double TileHeight, string? TileFlip, double Opacity);

/// <summary>Line cap (DrawingML a:ln cap: flat, rnd, sq).</summary>
public enum WordLineCap : byte { Flat, Round, Square }

/// <summary>Line join (SVG stroke-linejoin; ODF draw:stroke-linejoin miter / round / bevel, middle and none drawn as LibreOffice draws them).</summary>
public enum WordLineJoin : byte { Miter, Round, Bevel, None }

/// <summary>A 2-D affine transform (x' = A x + C y + E, y' = B x + D y + F), SVG matrix(a b c d e f) order.</summary>
public readonly record struct WordAffine(double A, double B, double C, double D, double E, double F)
{
    public static readonly WordAffine Identity = new(1, 0, 0, 1, 0, 0);
    public static WordAffine Translate(double x, double y) => new(1, 0, 0, 1, x, y);
    public static WordAffine Scale(double x, double y) => new(x, 0, 0, y, 0, 0);
    /// <summary>A rotation by <paramref name="radians"/> in y-down coordinates (positive = clockwise on screen, basegfx rotate).</summary>
    public static WordAffine Rotate(double radians) { double c = Math.Cos(radians), s = Math.Sin(radians); return new(c, s, -s, c, 0, 0); }
    /// <summary><paramref name="first"/> applied, then <paramref name="then"/> (basegfx: then * first).</summary>
    public static WordAffine Then(WordAffine first, WordAffine then)
        => new(then.A * first.A + then.C * first.B, then.B * first.A + then.D * first.B,
               then.A * first.C + then.C * first.D, then.B * first.C + then.D * first.D,
               then.A * first.E + then.C * first.F + then.E, then.B * first.E + then.D * first.F + then.F);
    public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);
    public bool IsIdentity => A == 1 && B == 0 && C == 0 && D == 1 && E == 0 && F == 0;
    public string Svg => "matrix(" + WordShapeHtml.Number6(A) + " " + WordShapeHtml.Number6(B) + " " + WordShapeHtml.Number6(C) + " " + WordShapeHtml.Number6(D) + " "
                       + WordShapeHtml.Number6(E) + " " + WordShapeHtml.Number6(F) + ")";
    /// <summary>The same transform for CSS (translation in px: 1 pt = 4/3 px).</summary>
    public string Css => "matrix(" + WordShapeHtml.Number6(A) + "," + WordShapeHtml.Number6(B) + "," + WordShapeHtml.Number6(C) + "," + WordShapeHtml.Number6(D) + ","
                       + WordShapeHtml.Number6(E * 4 / 3) + "," + WordShapeHtml.Number6(F * 4 / 3) + ")";
}

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
    /// <summary>A fill drawn by its own format's SVG paint (OfficeArt rectangular shades, patterns, picture fills): the &lt;defs&gt;
    /// content (written once before the shape's fill) and the paint (e.g. "url(#id)") used instead of <see cref="FillColor"/>.</summary>
    internal string? FillDefinitions, FillPaint;
    /// <summary>Text-effect (WordArt / Fontwork) lines set on paths (<see cref="WordOfficeArtFontwork"/>), drawn in their font and colour.</summary>
    internal List<WordShapeTextOnPath>? TextOnPaths;
    internal string? TextOnPathFont, TextOnPathColor;
    /// <summary>The text effect's glyph outline (ODF Fontwork glyphs carry the shape's line); null = none.</summary>
    internal string? TextOnPathStroke;
    internal double TextOnPathStrokeWidth;
    /// <summary>gtextAlign (MSO_GEOTEXTALIGN, msdffdef.hxx:594-601: 0 stretch, 1 centre - the default, msdffimp.cxx:4466 -, 2 left, 3 right, 4 letter / 5 word justify).</summary>
    internal int TextOnPathAlign = 1;
    public WordLineCap LineCap;
    /// <summary>Per stop of <see cref="GradientStops"/> (null = all opaque).</summary>
    public List<double>? GradientStopOpacities;
    /// <summary>A radial gradient (DrawingML a:path circle) centred at (<see cref="GradientCenterX"/>, <see cref="GradientCenterY"/>) as fractions of the box.</summary>
    public bool GradientRadial;
    public double GradientCenterX = 0.5, GradientCenterY = 0.5;
    /// <summary>A rectangular gradient (DrawingML a:path rect / shape) about the same centre.</summary>
    public bool GradientRectangular;
    /// <summary>The geometry's text rectangle in the shape's own points (unflipped); null = the whole box.</summary>
    public (double X, double Y, double Width, double Height)? TextRect;
    /// <summary>A preset pattern (pattFill) drawn over <see cref="FillColor"/> (its background, or none).</summary>
    public WordShapePattern? Pattern;
    /// <summary>A picture fill (blipFill) by picture id, stretched to the box less its fill rectangle, or tiled.</summary>
    public string? FillPictureId;
    public WordShapePictureFill FillPicture;
    /// <summary>DrawingML text of the shape itself (SmartArt, charts) as HTML, drawn in its text box; <see cref="DrawingTextAnchor"/> t / ctr / b.</summary>
    public string? DrawingTextHtml;
    public string DrawingTextAnchor = "t";
    public int TextIndex = -1;
    /// <summary>A DrawingML WordArt (a:prstTxWarp): its text-effect shape type and adjust values; the renderer lays its text box text on the outlines.</summary>
    public int WordArtShapeType;
    public List<int?>? WordArtAdjust;
    public string? PictureId;
    public WordTextBoxStyle TextStyle;
    /// <summary>Maps the shape's own box (0..Width, 0..Height) to the frame; when set it replaces X / Y / rotation / flips (ODF draw:transform with shear).</summary>
    public WordAffine? Transform;
    public WordLineJoin LineJoin;
    /// <summary>A fill drawn by its own format's paint for a path's brightness (ODF gradients, hatches, transparency gradients): the &lt;defs&gt; content and the paint.</summary>
    internal Func<double, (string Definitions, string Paint)>? FillPainter;
    /// <summary>A tiled picture fill's alternate-row / alternate-column shift in tile fractions (ODF draw:tile-repeat-offset).</summary>
    public double TileRowOffset, TileColumnOffset;
    /// <summary>A picture fill's transparency gradient: an SVG luminance mask in the shape's own points (definitions, id).</summary>
    internal (string Definitions, string Id)? FillPictureMask;
    /// <summary>stroke-miterlimit for miter joins (0 = the SVG default 4); LibreOffice bevels below 15 degrees (b2dlinegeometry.hxx:142): 1 / sin(7.5 deg).</summary>
    public double MiterLimit;
}

/// <summary>
/// Draws <see cref="WordShape"/>s: a frame placed like a picture (<see cref="WordPictureHtml.AppendPlacement"/>), one inline
/// &lt;svg&gt; filling it with every outline in frame points, and each shape's text box / picture as an absolutely placed box over its
/// outline. A frame of one shape whose text grows (auto-fit) keeps its text in the flow instead, so the frame grows and the
/// outline (drawn with non-scaling strokes) stretches with it.
/// </summary>
public static class WordShapeHtml
{
    public static void OpenFrame(StringBuilder html, WordPictureLayout layout, bool growsWithText, string? extraCss = null)
    {
        var css = new StringBuilder("position:relative;box-sizing:border-box;");
        css.Append(layout.Wrap == WordPictureWrap.Inline ? "display:inline-block;vertical-align:bottom;" : "");
        css.Append("width:").Append(Points(layout.WidthPt)).Append("pt;");
        css.Append(growsWithText ? "min-height:" : "height:").Append(Points(layout.HeightPt)).Append("pt;");
        WordPictureHtml.AppendPlacement(css, layout);
        css.Append(extraCss);
        html.Append("<span class=\"word-shape\" style=\"").Append(css).Append("\">");
    }

    public static void CloseFrame(StringBuilder html) => html.Append("</span>");

    /// <summary>The outlines of every shape (pictures are drawn as boxes by <see cref="AppendPicture"/>); a picture-filled shape splits the SVG so its picture layer keeps its place in the drawing order.</summary>
    public static void AppendGeometry(StringBuilder html, IReadOnlyList<WordShape> shapes, double width, double height, bool stretches)
    {
        width = Math.Max(width, 0.01); height = Math.Max(height, 0.01);
        void Open() => html.Append("<svg class=\"word-shape-geometry\" aria-hidden=\"true\" viewBox=\"0 0 ").Append(Points(width)).Append(' ').Append(Points(height))
            .Append("\" width=\"").Append(Points(width)).Append("pt\" height=\"").Append(Points(height)).Append("pt\"")
            .Append(stretches ? " preserveAspectRatio=\"none\"" : "")
            .Append(" style=\"position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none\">");
        Open();
        foreach (WordShape shape in shapes)
        {
            if (shape.Geometry == WordShapeGeometry.Picture) continue;
            if (shape.FillPictureId is not null && shape.Paths is not null)
            {
                html.Append("</svg>");
                AppendPictureFill(html, shape, width, height);
                Open();
            }
            AppendShape(html, shape, stretches);
        }
        html.Append("</svg>");
    }

    private static void AppendShape(StringBuilder html, WordShape shape, bool stretches)
    {
        double w = shape.Width, h = shape.Height;
        if (shape.Transform is { } transform) html.Append("<g transform=\"").Append(transform.Svg).Append("\">");
        else
        {
            html.Append("<g transform=\"translate(").Append(Number(shape.X + w / 2)).Append(' ').Append(Number(shape.Y + h / 2)).Append(')');
            if (shape.RotationDegrees != 0) html.Append(" rotate(").Append(Number(shape.RotationDegrees)).Append(')');
            if (shape.FlipHorizontal || shape.FlipVertical) html.Append(" scale(").Append(shape.FlipHorizontal ? "-1" : "1").Append(' ').Append(shape.FlipVertical ? "-1" : "1").Append(')');
            html.Append(" translate(").Append(Number(-w / 2)).Append(' ').Append(Number(-h / 2)).Append(")\">");
        }
        if (shape.Paths is not null)
        {
            AppendPaths(html, shape, stretches);
            if (shape.TextOnPaths is { Count: > 0 }) AppendTextOnPaths(html, shape);
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

    /// <summary>A clip path of <paramref name="pathData"/> (user space), defined once by its content; returns its id.</summary>
    private static string AppendClip(StringBuilder html, string pathData)
    {
        string id = "word-shape-clip-" + StableHash(pathData);
        html.Append("<defs><clipPath id=\"").Append(id).Append("\"><path d=\"").Append(pathData).Append("\" clip-rule=\"evenodd\"/></clipPath></defs>");
        return id;
    }

    private static string StopsHtml(WordShape shape, double brightness)
    {
        var stops = new StringBuilder();
        var list = shape.GradientStops!;
        for (int i = 0; i < list.Count; i++)
        {
            double opacity = shape.GradientStopOpacities is { } opacities && i < opacities.Count ? opacities[i] : 1;
            stops.Append("<stop offset=\"").Append(Number(list[i].Offset)).Append("\" stop-color=\"").Append(WordShapeColor.Brightened(list[i].Color, brightness)).Append('"');
            if (opacity < 1) stops.Append(" stop-opacity=\"").Append(Opacity(opacity)).Append('"');
            stops.Append("/>");
        }
        return stops.ToString();
    }

    /// <summary>
    /// LibreOffice's rectangular gradient (a:path rect / shape): basegfx gradienttools.cxx initRectGradientInfo :196-270 places
    /// the box at the fillToRect centre, getRectangularGradientAlpha :694-755 gives 1 - max(|x|, |y|) with the aspect correction,
    /// that is the distance to the box's nearest edge over half its shorter side, and the stops run reversed from the centre
    /// (fillproperties.cxx:549-551): position = 1 - distance / (shorter side / 2). Drawn as the four edge regions (each region
    /// is nearest one edge, split on the 45-degree lines from the corners), each with a linear gradient across it, clipped
    /// to the path; outside the box (a moved centre) the last stop (alpha 0).
    /// </summary>
    private static void AppendRectangularGradient(StringBuilder html, WordShape shape, string fillData, double brightness)
    {
        double w = shape.Width, h = shape.Height, bx = (shape.GradientCenterX - 0.5) * w, by = (shape.GradientCenterY - 0.5) * h, m = Math.Min(w, h) / 2;
        if (m <= 0) return;
        string clip = AppendClip(html, fillData);
        var list = shape.GradientStops!;
        double lastOpacity = shape.GradientStopOpacities is { } o && o.Count == list.Count ? o[list.Count - 1] : 1;
        html.Append("<g clip-path=\"url(#").Append(clip).Append(")\"><path d=\"").Append(fillData).Append("\" fill-rule=\"evenodd\" fill=\"")
            .Append(WordShapeColor.Brightened(list[list.Count - 1].Color, brightness)).Append('"');
        if (lastOpacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(lastOpacity)).Append('"');
        html.Append(" stroke=\"none\"/>");
        string stops = StopsHtml(shape, brightness);
        string P(double x, double y) => Number(x) + " " + Number(y);
        (double x1, double y1, double x2, double y2, string polygon)[] regions =
        [
            (0, by + m, 0, by, "M" + P(bx, by) + "L" + P(bx + w, by) + "L" + P(bx + w - m, by + m) + "L" + P(bx + m, by + m) + "Z"),
            (0, by + h - m, 0, by + h, "M" + P(bx, by + h) + "L" + P(bx + m, by + h - m) + "L" + P(bx + w - m, by + h - m) + "L" + P(bx + w, by + h) + "Z"),
            (bx + m, 0, bx, 0, "M" + P(bx, by) + "L" + P(bx + m, by + m) + "L" + P(bx + m, by + h - m) + "L" + P(bx, by + h) + "Z"),
            (bx + w - m, 0, bx + w, 0, "M" + P(bx + w, by) + "L" + P(bx + w, by + h) + "L" + P(bx + w - m, by + h - m) + "L" + P(bx + w - m, by + m) + "Z"),
        ];
        foreach ((double x1, double y1, double x2, double y2, string polygon) in regions)
        {
            string definition = "<linearGradient gradientUnits=\"userSpaceOnUse\" x1=\"" + Number(x1) + "\" y1=\"" + Number(y1) + "\" x2=\"" + Number(x2) + "\" y2=\"" + Number(y2) + "\">" + stops + "</linearGradient>";
            string id = "word-shape-gradient-" + StableHash(definition);
            html.Append("<defs>").Append(definition.Insert("<linearGradient".Length, " id=\"" + id + "\"")).Append("</defs>")
                .Append("<path d=\"").Append(polygon).Append("\" fill=\"url(#").Append(id).Append(")\" stroke=\"none\"/>");
        }
        html.Append("</g>");
    }

    /// <summary>
    /// A preset pattern over the fill (a:pattFill): its 8 x 8 bit cell (DocxPresetPattern) as an SVG pattern tile, one cell
    /// pixel = 0.75 pt (one CSS pixel at 96 per inch, the device pixel the GDI+ hatch brush draws), set bits in the foreground
    /// colour, the background under it already drawn as the shape's fill; clipped to the path. The colour takes the path's
    /// darken / lighten too (AdaptObjColor :2851-2865).
    /// </summary>
    private static void AppendPattern(StringBuilder html, WordShape shape, string fillData, WordShapePattern pattern, double brightness)
    {
        const double Cell = 0.75;
        var cells = new StringBuilder();
        for (int y = 0; y < 8; y++)
            for (int x = 0; x < 8; x++)
                if ((pattern.Rows[y] & (0x80 >> x)) != 0) cells.Append('M').Append(Number(x * Cell)).Append(' ').Append(Number(y * Cell)).Append("h0.75v0.75h-0.75Z");
        string color = WordShapeColor.Brightened(pattern.Color, brightness);
        string definition = "<pattern patternUnits=\"userSpaceOnUse\" width=\"6\" height=\"6\"><path d=\"" + cells + "\" fill=\"" + color + "\""
                          + (pattern.Opacity < 1 ? " fill-opacity=\"" + Opacity(pattern.Opacity) + "\"" : "") + "/></pattern>";
        string id = "word-shape-pattern-" + StableHash(definition);
        html.Append("<defs>").Append(definition.Insert("<pattern".Length, " id=\"" + id + "\"")).Append("</defs>")
            .Append("<path d=\"").Append(fillData).Append("\" fill-rule=\"evenodd\" fill=\"url(#").Append(id).Append(")\" stroke=\"none\"/>");
    }

    /// <summary>
    /// A picture fill (a:blipFill) as an HTML layer between two SVG layers, so the shapes after it still draw over it: the
    /// picture element the host gives its source to, placed and turned like the shape, clipped to the outline by an SVG clip
    /// path in the box's own units. Stretched: into the box inset by the fill rectangle; tiled: copies from the tile origin.
    /// </summary>
    private static void AppendPictureFill(StringBuilder html, WordShape shape, double frameWidth, double frameHeight)
    {
        StringBuilder outer = html;
        bool masked = shape.FillPictureMask is not null && shape.Transform is not null;
        if (masked) html = new StringBuilder();
        double w = Math.Max(shape.Width, 0.01), h = Math.Max(shape.Height, 0.01);
        var unit = new StringBuilder();
        foreach (WordShapePath path in shape.Paths!)
            if (path.Fill != WordShapePathFill.None)
                foreach (WordShapeSubpath subpath in path.Subpaths)
                {
                    unit.Append('M').Append(Unit(subpath.StartX / w)).Append(' ').Append(Unit(subpath.StartY / h));
                    foreach (WordShapeSegment s in subpath.Segments)
                    {
                        if (s.IsCurve)
                            unit.Append('C').Append(Unit(s.Control1X / w)).Append(' ').Append(Unit(s.Control1Y / h)).Append(' ').Append(Unit(s.Control2X / w)).Append(' ')
                                .Append(Unit(s.Control2Y / h)).Append(' ').Append(Unit(s.X / w)).Append(' ').Append(Unit(s.Y / h));
                        else unit.Append('L').Append(Unit(s.X / w)).Append(' ').Append(Unit(s.Y / h));
                    }
                    unit.Append('Z');
                }
        string clip = "word-shape-clip-" + StableHash("bbox" + unit);
        html.Append("<svg width=\"0\" height=\"0\" style=\"position:absolute\" aria-hidden=\"true\"><defs><clipPath id=\"").Append(clip)
            .Append("\" clipPathUnits=\"objectBoundingBox\"><path d=\"").Append(unit).Append("\" clip-rule=\"evenodd\"/></clipPath></defs></svg>");
        var css = new StringBuilder("position:absolute;display:block;overflow:hidden;left:").Append(Points(shape.X)).Append("pt;top:").Append(Points(shape.Y))
            .Append("pt;width:").Append(Points(w)).Append("pt;height:").Append(Points(h)).Append("pt;clip-path:url(#").Append(clip).Append(");");
        if (masked)
            css.Clear().Append("position:absolute;display:block;overflow:hidden;left:0;top:0;width:").Append(Points(w)).Append("pt;height:").Append(Points(h))
               .Append("pt;clip-path:url(#").Append(clip).Append(");");
        else if (shape.Transform is { } transform)
        {
            css.Clear().Append("position:absolute;display:block;overflow:hidden;left:0;top:0;width:").Append(Points(w)).Append("pt;height:").Append(Points(h))
               .Append("pt;clip-path:url(#").Append(clip).Append(");transform-origin:0 0;transform:").Append(transform.Css).Append(';');
        }
        else if (shape.RotationDegrees != 0 || shape.FlipHorizontal || shape.FlipVertical)
            css.Append("transform:rotate(").Append(Number(shape.RotationDegrees)).Append("deg) scale(").Append(shape.FlipHorizontal ? "-1" : "1").Append(',').Append(shape.FlipVertical ? "-1" : "1").Append(");");
        WordShapePictureFill fill = shape.FillPicture;
        if (fill.Opacity < 1) css.Append("opacity:").Append(Opacity(fill.Opacity)).Append(';');
        html.Append("<span class=\"word-shape-fill-picture\" style=\"").Append(css).Append("\">");
        void Picture(double left, double top, double width, double height)
            => html.Append("<img class=\"word-picture\" alt=\"\" ").Append(WordPictureHtml.PictureIdAttribute).Append("=\"").Append(WordHtmlText.Encode(shape.FillPictureId!))
                   .Append("\" loading=\"lazy\" decoding=\"async\" style=\"position:absolute;max-width:none;left:").Append(Number(left)).Append("pt;top:").Append(Number(top))
                   .Append("pt;width:").Append(Points(width)).Append("pt;height:").Append(Points(height)).Append("pt;\">");
        if (fill.Tiled && fill.TileWidth > 0 && fill.TileHeight > 0)
        {
            double startX = fill.TileOffsetX - Math.Ceiling(fill.TileOffsetX / fill.TileWidth) * fill.TileWidth, startY = fill.TileOffsetY - Math.Ceiling(fill.TileOffsetY / fill.TileHeight) * fill.TileHeight;
            // ODF draw:tile-repeat-offset: every other row (column) shifted by a fraction of the tile
            // a tile smaller than 1/256 of the box repeats below a pixel on screen: drawn at that size (bounds the picture elements)
            if (!(fill.TileWidth >= w / 256) || !(fill.TileHeight >= h / 256) || !(!double.IsNaN(startX) && !double.IsInfinity(startX)) || !(!double.IsNaN(startY) && !double.IsInfinity(startY)))
                fill = fill with { TileWidth = Math.Max(fill.TileWidth, w / 256), TileHeight = Math.Max(fill.TileHeight, h / 256) };
            if (!(!double.IsNaN(startX) && !double.IsInfinity(startX))) startX = 0;
            if (!(!double.IsNaN(startY) && !double.IsInfinity(startY))) startY = 0;
            double rowShift = shape.TileRowOffset * fill.TileWidth, columnShift = shape.TileColumnOffset * fill.TileHeight;
            if (rowShift > 0) startX -= fill.TileWidth;
            if (columnShift > 0) startY -= fill.TileHeight;
            int row = 0;
            for (double y = startY; y < h + columnShift; y += fill.TileHeight, row++)
            {
                int column = 0;
                for (double x = startX; x < w + rowShift; x += fill.TileWidth, column++)
                    Picture(x + ((row & 1) != 0 ? rowShift : 0), y + ((column & 1) != 0 ? columnShift : 0), fill.TileWidth, fill.TileHeight);
            }
        }
        else Picture(fill.Left * w, fill.Top * h, w * (1 - fill.Left - fill.Right), h * (1 - fill.Top - fill.Bottom));
        html.Append("</span>");
        if (!masked) return;
        // a transparency gradient: the picture layer inside the SVG (foreignObject), masked in the shape's own points
        (string definitions, string maskId) = shape.FillPictureMask!.Value;
        outer.Append("<svg class=\"word-shape-geometry\" aria-hidden=\"true\" viewBox=\"0 0 ").Append(Points(frameWidth)).Append(' ').Append(Points(frameHeight))
             .Append("\" style=\"position:absolute;left:0;top:0;width:100%;height:100%;overflow:visible;pointer-events:none\"><g transform=\"")
             .Append(shape.Transform!.Value.Svg).Append("\"><defs>").Append(definitions).Append("</defs><g mask=\"url(#").Append(maskId).Append(")\"><foreignObject x=\"0\" y=\"0\" width=\"")
             .Append(Number(w)).Append("\" height=\"").Append(Number(h)).Append("\">").Append(html).Append("</foreignObject></g></g></svg>");
    }

    private static string Unit(double value) => Math.Round(value, 5).ToString("0.#####", CultureInfo.InvariantCulture);

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
            double brightness = path.Brightness != 0 ? path.Brightness : path.Fill switch
            {
                WordShapePathFill.Darken => -0.4, WordShapePathFill.DarkenLess => -0.2,
                WordShapePathFill.Lighten => 0.4, WordShapePathFill.LightenLess => 0.2, _ => 0,
            };
            if (path.LuminancePercent != 0) brightness = WordShapeColor.LuminanceMarker(path.LuminancePercent);
            if (path.Fill != WordShapePathFill.None && shape.FillPainter is { } painter)
            {
                (string definitions, string painted) = painter(brightness);
                html.Append("<defs>").Append(definitions).Append("</defs>");
                html.Append("<path d=\"").Append(PathData(path.Subpaths, closeAll: true)).Append("\" fill-rule=\"evenodd\" fill=\"").Append(painted).Append('"');
                if (shape.FillOpacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(shape.FillOpacity)).Append('"');
                html.Append(" stroke=\"none\"/>");
            }
            else if (path.Fill != WordShapePathFill.None && shape.FillPaint is { } paint)
            {
                if (shape.FillDefinitions is { } definitions) { html.Append("<defs>").Append(definitions).Append("</defs>"); shape.FillDefinitions = null; }
                html.Append("<path d=\"").Append(PathData(path.Subpaths, closeAll: true)).Append("\" fill-rule=\"evenodd\" fill=\"").Append(paint).Append('"');
                if (shape.FillOpacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(shape.FillOpacity)).Append('"');
                html.Append(" stroke=\"none\"/>");
            }
            else if (path.Fill != WordShapePathFill.None)
            {
                string fillData = PathData(path.Subpaths, closeAll: true);
                if (shape.GradientRectangular && shape.GradientStops is { Count: > 0 }) AppendRectangularGradient(html, shape, fillData, brightness);
                else if (shape.FillColor is not null || shape.GradientStops is { Count: > 0 })
                {
                    string? gradientId = AppendGradient(html, shape, brightness);
                    html.Append("<path d=\"").Append(fillData).Append("\" fill-rule=\"evenodd\" fill=\"")
                        .Append(gradientId is not null ? "url(#" + gradientId + ")" : WordShapeColor.Brightened(shape.FillColor!, brightness)).Append('"');
                    if (gradientId is null && shape.FillOpacity < 1) html.Append(" fill-opacity=\"").Append(Opacity(shape.FillOpacity)).Append('"');
                    html.Append(" stroke=\"none\"/>");
                }
                if (shape.Pattern is { } pattern) AppendPattern(html, shape, fillData, pattern, brightness);
            }
            if (!path.Stroke || shape.LineColor is not { } line) continue;
            var stroke = new StringBuilder(" fill=\"none\" stroke=\"").Append(line).Append("\" stroke-width=\"").Append(Number(lineWidth)).Append('"');
            if (shape.LineOpacity < 1) stroke.Append(" stroke-opacity=\"").Append(Opacity(shape.LineOpacity)).Append('"');
            if (shape.LineCap != WordLineCap.Flat) stroke.Append(" stroke-linecap=\"").Append(shape.LineCap == WordLineCap.Round ? "round" : "square").Append('"');
            if (shape.LineJoin is WordLineJoin.Round or WordLineJoin.Bevel) stroke.Append(" stroke-linejoin=\"").Append(shape.LineJoin == WordLineJoin.Round ? "round" : "bevel").Append('"');
            else if (shape.MiterLimit > 0) stroke.Append(" stroke-miterlimit=\"").Append(Number(shape.MiterLimit)).Append('"');
            if (shape.Dash is { Length: > 0 } dash) stroke.Append(" stroke-dasharray=\"").Append(string.Join(" ", dash.Select(d => Number(Math.Max(0, d) * lineWidth)))).Append('"');
            if (stretches) stroke.Append(" vector-effect=\"non-scaling-stroke\"");
            bool hasEnds = shape.HeadEnd.Type != WordLineEndType.None || shape.TailEnd.Type != WordLineEndType.None || shape.HeadMarker is not null || shape.TailMarker is not null;
            var plain = new List<WordShapeSubpath>();
            foreach (WordShapeSubpath subpath in path.Subpaths)
            {
                if (!hasEnds || subpath.Closed || subpath.Segments.Count == 0) { plain.Add(subpath); continue; }
                AppendWithLineEnds(html, subpath, shape, stroke.ToString(), line);
            }
            if (plain.Count > 0 && shape.LineJoin == WordLineJoin.None)
            {
                // B2DLineJoin::NONE: every segment stroked on its own, nothing fills the corners (b2dlinegeometry.cxx:886)
                foreach (WordShapeSubpath subpath in plain) foreach (WordShapeSubpath segment in Segments(subpath))
                    html.Append("<path d=\"").Append(PathData([segment], closeAll: false)).Append('"').Append(stroke).Append("/>");
            }
            else if (plain.Count > 0) html.Append("<path d=\"").Append(PathData(plain, closeAll: false)).Append('"').Append(stroke).Append("/>");
        }
    }

    /// <summary>Each text-effect line along its path: SVG textPath places and turns every glyph, its centre on the path, the line stretched to the path's length.</summary>
    /// <remarks>
    /// The page finishes what needs the browser's font metrics (KitveiHakodesh vue-frontend fastDocViewFontwork.ts), from these
    /// attributes on each &lt;text&gt;: data-fontwork="single" with data-fontwork-height (the line height before the fit factor) and
    /// data-fontwork-length (its outline's length) - the shape's lines share one fit factor, so they sit in one data-fontwork-shape
    /// group; data-fontwork="pair" with data-fontwork-upper / -lower (the two outlines) and data-fontwork-line / -lines (its band).
    /// Without the page's script the text stays on the path as written here.
    /// </remarks>
    private static void AppendTextOnPaths(StringBuilder html, WordShape shape)
    {
        if (shape.FillDefinitions is { } definitions) { html.Append("<defs>").Append(definitions).Append("</defs>"); shape.FillDefinitions = null; }
        html.Append("<g data-fontwork-shape=\"\">");
        foreach (WordShapeTextOnPath line in shape.TextOnPaths!)
        {
            if (line.Text.Length == 0 || line.PathLength <= 0) continue;
            string id = "word-shape-textpath-" + StableHash(line.PathData + line.Text);
            html.Append("<defs><path id=\"").Append(id).Append("\" d=\"").Append(line.PathData).Append("\"/></defs>");
            html.Append("<text font-size=\"").Append(Number(line.FontSizePt)).Append("\" dominant-baseline=\"central\" fill=\"").Append(shape.TextOnPathColor ?? "#000000").Append('"');
            if (shape.TextOnPathStroke is { } textStroke)
                html.Append(" stroke=\"").Append(textStroke).Append("\" stroke-width=\"").Append(Number(Math.Max(0.25, shape.TextOnPathStrokeWidth))).Append('"');
            if (shape.TextOnPathFont is { Length: > 0 } font) html.Append(" font-family=\"").Append(WordHtmlText.Encode(font)).Append('"');
            if (line.UpperData is { } upper && line.LowerData is { } lower)
                html.Append(" data-fontwork=\"pair\" data-fontwork-upper=\"").Append(upper).Append("\" data-fontwork-lower=\"").Append(lower)
                    .Append("\" data-fontwork-line=\"").Append(line.LineIndex).Append("\" data-fontwork-lines=\"").Append(line.LineCount).Append('"');
            else
                html.Append(" data-fontwork=\"single\" data-fontwork-height=\"").Append(Number(line.LineHeightPt)).Append("\" data-fontwork-length=\"").Append(Number(line.PathLength)).Append('"');
            html.Append(" data-fontwork-align=\"").Append(shape.TextOnPathAlign).Append('"');
            html.Append("><textPath href=\"#").Append(id).Append("\" textLength=\"").Append(Number(line.PathLength)).Append("\" lengthAdjust=\"spacingAndGlyphs\">")
                .Append(WordHtmlText.Encode(line.Text)).Append("</textPath></text>");
        }
        html.Append("</g>");
    }

    /// <summary>Each segment of a sub-path (its closing edge too) as a sub-path of its own.</summary>
    private static IEnumerable<WordShapeSubpath> Segments(WordShapeSubpath subpath)
    {
        double x = subpath.StartX, y = subpath.StartY;
        foreach (WordShapeSegment segment in subpath.Segments)
        {
            var one = new WordShapeSubpath { StartX = x, StartY = y };
            one.Segments.Add(segment);
            yield return one;
            x = segment.X; y = segment.Y;
        }
        if (subpath.Closed && (x != subpath.StartX || y != subpath.StartY))
        {
            var closing = new WordShapeSubpath { StartX = x, StartY = y };
            closing.Segments.Add(WordShapeSegment.Line(subpath.StartX, subpath.StartY));
            yield return closing;
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
        foreach (List<(double X, double Y)> polygon in new[] { marker.Points }.Concat(marker.MorePolygons ?? []))
        {
            for (int i = 0; i < polygon.Count; i++)
            {
                (double x, double y) = polygon[i];
                d.Append(i == 0 ? 'M' : 'L').Append(Number(head.X + x * cos - y * sin)).Append(' ').Append(Number(head.Y + x * sin + y * cos));
            }
            d.Append('Z');
        }
        html.Append("<path d=\"").Append(d).Append(marker.MorePolygons is { Count: > 0 } ? "\" fill-rule=\"evenodd\" fill=\"" : "\" fill=\"").Append(color).Append('"');
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

    /// <summary>
    /// Opens the box a shape's text sits in: over the shape's outline, or (a growing single shape) in the frame's flow. The box is
    /// the geometry's text rectangle when it has one, mirrored with the shape's flips (svx EnhancedCustomShape2d::GetTextRect
    /// :1152-1175), turned with the shape about the shape's centre; a vertical flip turns the text 180 degrees more and a
    /// horizontal one leaves it (LibreOffice inverts the stored rotation of a once-flipped shape, oox shape.cxx:1501-1515, and
    /// SdrObjCustomShape takes 180 - angle for a mirrored Y, svdoashp.cxx:1551-1566).
    /// </summary>
    public static void OpenShapeText(StringBuilder html, WordShape shape, bool inFlow)
    {
        WordTextBoxStyle style = shape.TextStyle;
        var css = new StringBuilder("box-sizing:border-box;overflow-wrap:break-word;display:block;margin:0;border:0;");
        if (inFlow) css.Append("position:relative;width:100%;");
        else
        {
            (double x, double y, double w, double h) = shape.TextRect ?? (0, 0, shape.Width, shape.Height);
            if (shape.TextRect is not null && shape.FlipHorizontal) x = shape.Width - x - w;
            if (shape.TextRect is not null && shape.FlipVertical) y = shape.Height - y - h;
            css.Append("position:absolute;left:").Append(Points(shape.X + x)).Append("pt;top:").Append(Points(shape.Y + y)).Append("pt;width:").Append(Points(w))
               .Append("pt;height:").Append(Points(h)).Append("pt;overflow:hidden;");
            double rotation = shape.RotationDegrees + (shape.FlipVertical ? 180 : 0);
            rotation -= 360 * Math.Floor(rotation / 360);
            if (rotation != 0)
            {
                if (x != 0 || y != 0 || w != shape.Width || h != shape.Height)
                    css.Append("transform-origin:").Append(Number(shape.Width / 2 - x)).Append("pt ").Append(Number(shape.Height / 2 - y)).Append("pt;");
                css.Append("transform:rotate(").Append(Number(rotation)).Append("deg);");
            }
        }
        css.Append("padding:").Append(Points(style.TopInsetPt)).Append("pt ").Append(Points(style.RightInsetPt)).Append("pt ")
           .Append(Points(style.BottomInsetPt)).Append("pt ").Append(Points(style.LeftInsetPt)).Append("pt;");
        if (style.WritingMode is { } writingMode) css.Append(writingMode);
        html.Append("<span class=\"text-box word-shape-text\" style=\"").Append(css).Append("\">");
    }

    public static void CloseShapeText(StringBuilder html) => html.Append("</span>");

    /// <summary>A shape's own DrawingML text (<see cref="WordShape.DrawingTextHtml"/>) in its text box, anchored top / centre / bottom (bodyPr anchor).</summary>
    public static void AppendDrawingText(StringBuilder html, WordShape shape, string drawingText)
    {
        OpenShapeText(html, shape, inFlow: false);
        string justify = shape.DrawingTextAnchor switch { "ctr" => "center", "b" => "flex-end", _ => "flex-start" };
        html.Append("<span style=\"display:flex;flex-direction:column;justify-content:").Append(justify).Append(";height:100%;\">").Append(drawingText).Append("</span>");
        CloseShapeText(html);
    }

    /// <summary>A picture inside a group, at its box over the frame, turned and mirrored like the shapes (flips at its centre).</summary>
    public static void AppendPicture(StringBuilder html, WordShape shape)
    {
        html.Append("<span class=\"word-shape-part\" style=\"position:absolute;display:block;left:").Append(Points(shape.X)).Append("pt;top:").Append(Points(shape.Y)).Append("pt;");
        if (shape.RotationDegrees != 0 || shape.FlipHorizontal || shape.FlipVertical)
        {
            html.Append("transform:");
            if (shape.RotationDegrees != 0) html.Append("rotate(").Append(Number(shape.RotationDegrees)).Append("deg)");
            if (shape.FlipHorizontal || shape.FlipVertical) html.Append(" scale(").Append(shape.FlipHorizontal ? "-1" : "1").Append(',').Append(shape.FlipVertical ? "-1" : "1").Append(')');
            html.Append(';');
        }
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

    internal static string Number6(double value) { double rounded = Math.Round(value, 6); return (rounded == 0 ? 0 : rounded).ToString("0.######", CultureInfo.InvariantCulture); }

    internal static string StableHashOf(string text) => StableHash(text);

    private static string Points(double value) => Math.Max(0, value).ToString("0.##", CultureInfo.InvariantCulture);
}

/// <summary>A line-end marker ready to place: its outline with the tip (or, centred, its middle) at (0, 0) and its body along +y, in points.</summary>
internal sealed record WordLineEndMarker(List<(double X, double Y)> Points, double Width, double Consumed)
{
    /// <summary>A marker's further polygons (an ODF draw:marker path may have several), placed like <see cref="Points"/>.</summary>
    public List<List<(double X, double Y)>>? MorePolygons { get; init; }
}

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
        if (brightness > 1 || brightness < -1) return Luminance(rgb, (int)Math.Round(brightness > 0 ? brightness - 1000 : brightness + 1000));
        int Channel(int c) => (int)Math.Max(0, Math.Min(255, brightness >= 0 ? c * (1.0 - brightness) + brightness * 255.0 : c * (1.0 + brightness)));
        return "#" + ((Channel((rgb >> 16) & 0xFF) << 16) | (Channel((rgb >> 8) & 0xFF) << 8) | Channel(rgb & 0xFF)).ToString("X6", CultureInfo.InvariantCulture);
    }

    /// <summary>A luminance change carried through the brightness parameter: percent + 1000 (positive) / - 1000 (negative), out of the -1..1 brightness range.</summary>
    internal static double LuminanceMarker(int percent) => percent > 0 ? 1000 + percent : -1000 + percent;

    /// <summary>
    /// The OpenOffice colour-data way (EnhancedCustomShape2d::GetColorData :1120-1149): in HSV, a positive change scales the
    /// saturation by (1 - n) and moves the value toward 1 by n; a negative one scales the value by (1 + n); rounded to bytes.
    /// </summary>
    private static string Luminance(int rgb, int percent)
    {
        double r = ((rgb >> 16) & 0xFF) / 255.0, g = ((rgb >> 8) & 0xFF) / 255.0, b = (rgb & 0xFF) / 255.0;
        double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b)), delta = max - min;
        double hue = 0, saturation = max == 0 ? 0 : delta / max, value = max;
        if (delta != 0)
        {
            if (r == max) hue = (g - b) / delta; else if (g == max) hue = 2 + (b - r) / delta; else hue = 4 + (r - g) / delta;
            hue *= 60; if (hue < 0) hue += 360;
        }
        double n = percent / 100.0;
        if (n > 0) { saturation *= 1 - n; value = n + (1 - n) * value; }
        else value *= 1 + n;
        // basegfx hsv2rgb
        double red, green, blue;
        if (saturation == 0) red = green = blue = value;
        else
        {
            double h = hue == 360 ? 0 : hue / 60;
            int sector = (int)Math.Floor(h);
            double f = h - sector, p = value * (1 - saturation), q = value * (1 - saturation * f), t = value * (1 - saturation * (1 - f));
            (red, green, blue) = sector switch { 0 => (value, t, p), 1 => (q, value, p), 2 => (p, value, t), 3 => (p, q, value), 4 => (t, p, value), _ => (value, p, q) };
        }
        int Byte(double c) => (int)(Math.Max(0, Math.Min(1, c)) * 255.0 + 0.5);
        return "#" + ((Byte(red) << 16) | (Byte(green) << 8) | Byte(blue)).ToString("X6", CultureInfo.InvariantCulture);
    }
}
