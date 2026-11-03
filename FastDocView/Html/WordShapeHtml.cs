using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>The outline a drawn shape takes (format-neutral). <see cref="Box"/> = a preset this library does not draw: a plain box of the right size.</summary>
public enum WordShapeGeometry : byte { Rectangle, RoundRectangle, Ellipse, Line, Triangle, RightArrow, LeftArrow, UpArrow, DownArrow, Path, Box, Picture }

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
        string? gradientId = null;
        if (shape.GradientStops is { Count: > 0 } stops)
        {
            // the id is the gradient's own content: two blocks that both define it define the same thing
            var key = new StringBuilder();
            foreach ((double offset, string color) in stops) key.Append(Number(offset)).Append(color);
            key.Append(Number(shape.GradientAngleDegrees));
            gradientId = "word-shape-gradient-" + StableHash(key.ToString());
            double radians = shape.GradientAngleDegrees * Math.PI / 180;
            double dx = Math.Cos(radians) / 2, dy = Math.Sin(radians) / 2;
            html.Append("<defs><linearGradient id=\"").Append(gradientId).Append("\" x1=\"").Append(Number(0.5 - dx)).Append("\" y1=\"").Append(Number(0.5 - dy))
                .Append("\" x2=\"").Append(Number(0.5 + dx)).Append("\" y2=\"").Append(Number(0.5 + dy)).Append("\">");
            foreach ((double offset, string color) in stops)
                html.Append("<stop offset=\"").Append(Number(offset)).Append("\" stop-color=\"").Append(color).Append("\"/>");
            html.Append("</linearGradient></defs>");
        }
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

    internal static string Number(double value) => Math.Round(value, 2).ToString("0.##", CultureInfo.InvariantCulture);

    private static string Points(double value) => Math.Max(0, value).ToString("0.##", CultureInfo.InvariantCulture);
}
