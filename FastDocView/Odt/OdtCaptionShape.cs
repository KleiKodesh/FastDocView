using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>A laid-out draw:caption: its rectangle (with corner radii) and the tail polyline to the caption point, all in points in the anchor's coordinates.</summary>
internal sealed class OdtCaptionGeometry
{
    public double Left, Top, Width, Height;
    /// <summary>Corner radii, each clamped to half the side (calculateRelativeCornerRadius, svx/source/sdr/primitive2d/sdrattributecreator.cxx:1131-).</summary>
    public double CornerRadiusX, CornerRadiusY;
    /// <summary>The tail, from the caption point to the rectangle; it carries the shape's markers (sdrcaptionprimitive2d.cxx: line start / end on the tail only).</summary>
    public WordShapeSubpath Tail = new();
}

/// <summary>
/// draw:caption (ODF 1.3 part 3 §10.3.11) as LibreOffice builds it: SdrCaptionObj::ImpCalcTail1..3 and ImpCaptParams::CalcEscPos
/// (svx/source/svdraw/svdocapt.cxx:85-161, 421-505), drawn by SdrCaptionPrimitive2D (rectangle outline unrotated by its own transform; the
/// caller applies draw:transform as for any shape).
/// </summary>
internal static class OdtCaptionShape
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw, Svg = OdtStyleCascade.Svg;

    private enum Escape { Left, Right, Top, Bottom }   // LKS, RTS, OBN, UNT (svdocapt.cxx:57)
    private enum CaptionType { StraightLine, AngledLine, AngledConnectorLine }
    private enum EscapeDirection { Horizontal, Vertical, BestFit }

    public static OdtCaptionGeometry Layout(XElement caption, Func<XName, string?> get)
    {
        var g = new OdtCaptionGeometry
        {
            Left = Length((string?)caption.Attribute(Svg + "x")) ?? 0,
            Top = Length((string?)caption.Attribute(Svg + "y")) ?? 0,
            Width = Length((string?)caption.Attribute(Svg + "width")) ?? 0,
            Height = Length((string?)caption.Attribute(Svg + "height")) ?? 0,
        };

        // draw:corner-radius is read for captions (ximpshap.cxx:2420-2423, set as CornerRadius :2387-2400), clamped per axis.
        double radius = Length((string?)caption.Attribute(Draw + "corner-radius")) ?? 0;
        g.CornerRadiusX = Math.Max(0, Math.Min(radius, g.Width * 0.5));
        g.CornerRadiusY = Math.Max(0, Math.Min(radius, g.Height * 0.5));

        // The caption point is relative to the shape's top left (svx/source/unodraw/unoshape.cxx:2012-2034).
        double tailX = g.Left + (Length((string?)caption.Attribute(Draw + "caption-point-x")) ?? 0);
        double tailY = g.Top + (Length((string?)caption.Attribute(Draw + "caption-point-y")) ?? 0);

        // Parameters (sdpropls.cxx:326-336, enums :785-799; defaults include/svx/sxctitm.hxx:33, sxcecitm.hxx:36/53/73/93, sxcgitm.hxx:33, sxcllitm.hxx:34/51).
        CaptionType type = (get(Draw + "caption-type")?.Trim()) switch
        {
            "straight-line" => CaptionType.StraightLine, "angled-line" => CaptionType.AngledLine, _ => CaptionType.AngledConnectorLine,
        };
        EscapeDirection direction = (get(Draw + "caption-escape-direction")?.Trim()) switch
        {
            "vertical" => EscapeDirection.Vertical, "auto" => EscapeDirection.BestFit, _ => EscapeDirection.Horizontal,
        };
        double gap = Length(get(Draw + "caption-gap")) ?? 0;
        double lineLength = Length(get(Draw + "caption-line-length")) ?? 0;
        bool fitLineLength = get(Draw + "caption-fit-line-length")?.Trim() != "false";
        bool escapeRelative = true;
        double escapeFraction = 0.5, escapeAbsolute = 0;
        if (get(Draw + "caption-escape")?.Trim() is { Length: > 0 } escape)
        {
            if (escape.EndsWith("%", StringComparison.Ordinal))                   // XMLIsPercentagePropertyHandler / XMLCaptionEscapeRelative (sdpropls.cxx:818-828)
            {
                if (double.TryParse(escape.Substring(0, escape.Length - 1), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double percent))
                    escapeFraction = percent / 100;
            }
            else if (Length(escape) is { } absolute) { escapeRelative = false; escapeAbsolute = absolute; }
        }

        // CalcEscPos (svdocapt.cxx:85-161).
        double right = g.Left + g.Width, bottom = g.Top + g.Height;
        double nx = (escapeRelative ? g.Width * escapeFraction : escapeAbsolute) + g.Left;
        double ny = (escapeRelative ? g.Height * escapeFraction : escapeAbsolute) + g.Top;
        bool tryHorizontal = direction == EscapeDirection.BestFit
            || (type != CaptionType.StraightLine ? direction == EscapeDirection.Horizontal : direction == EscapeDirection.Vertical);
        bool tryVertical = direction == EscapeDirection.BestFit
            || (type != CaptionType.StraightLine ? direction == EscapeDirection.Vertical : direction == EscapeDirection.Horizontal);
        (double X, double Y) best = default;
        Escape bestDirection = Escape.Left;
        if (tryHorizontal)
        {
            (double X, double Y) leftPoint = (g.Left - gap, ny), rightPoint = (right + gap, ny);
            if (tailX - leftPoint.X < rightPoint.X - tailX) { bestDirection = Escape.Left; best = leftPoint; }
            else { bestDirection = Escape.Right; best = rightPoint; }
        }
        if (tryVertical)
        {
            (double X, double Y) topPoint = (nx, g.Top - gap), bottomPoint = (nx, bottom + gap);
            bool top = tailY - topPoint.Y < bottomPoint.Y - tailY;
            var second = top ? topPoint : bottomPoint;
            Escape secondDirection = top ? Escape.Top : Escape.Bottom;
            bool take = direction != EscapeDirection.BestFit;
            if (!take)
            {
                double horizontalSquared = Square(best.X - tailX) + Square(best.Y - tailY);
                double verticalSquared = Square(second.X - tailX) + Square(second.Y - tailY);
                take = type != CaptionType.StraightLine ? verticalSquared < horizontalSquared : verticalSquared >= horizontalSquared;
            }
            if (take) { best = second; bestDirection = secondDirection; }
        }
        bool sideways = bestDirection == Escape.Left || bestDirection == Escape.Right;

        var points = new List<(double X, double Y)>();
        switch (type)
        {
            case CaptionType.StraightLine:                                       // ImpCalcTail1 (:421-445): the tail point slides onto the escape's line
                points.Add(sideways ? (best.X, tailY) : (tailX, best.Y));
                points.Add(best);
                break;
            case CaptionType.AngledLine:                                         // ImpCalcTail2 (:447-462)
                points.Add((tailX, tailY));
                points.Add(best);
                break;
            default:                                                             // ImpCalcTail3 (:464-495)
            {
                var middle = best;
                if (sideways)
                {
                    if (fitLineLength) middle.X = (tailX + best.X) / 2;
                    else middle.X += bestDirection == Escape.Left ? -lineLength : lineLength;
                }
                else
                {
                    if (fitLineLength) middle.Y = (tailY + best.Y) / 2;
                    else middle.Y += bestDirection == Escape.Top ? -lineLength : lineLength;
                }
                points.Add((tailX, tailY));
                points.Add(middle);
                points.Add(best);
                break;
            }
        }
        g.Tail.StartX = points[0].X;
        g.Tail.StartY = points[0].Y;
        for (int i = 1; i < points.Count; i++) g.Tail.Segments.Add(WordShapeSegment.Line(points[i].X, points[i].Y));
        return g;
    }

    private static double Square(double v) => v * v;
    private static double? Length(string? value) => OdtShapeStyle.Length(value);
}
