using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>One line of text-effect (WordArt / Fontwork) text set on a path: SVG path data in the shape's points, the font size, and the text.</summary>
public sealed record WordShapeTextOnPath(string PathData, double PathLength, double FontSizePt, string Text);

/// <summary>
/// Text effects (WordArt / Fontwork) laid out as LibreOffice lays them out (svx/source/customshapes/EnhancedCustomShapeFontWork.cxx,
/// master read 2026-11-04), with SVG &lt;textPath&gt; doing the per-glyph placement:
/// - the outlines are every sub-path of the shape's geometry, curves subdivided (GetOutlinesFromShape2d :630-650);
/// - an odd number of outlines = single-line mode, one text area per outline; an even number = pairs of outlines (upper, lower)
///   per text area (InitializeFontWorkData :93-181, nOutlinesCount2d &amp; 1);
/// - the lines (paragraphs, split at line breaks) are shared out over the text areas in order, ((left - 1) / areas) + 1 per area
///   (:155-176);
/// - a line's height is the shape's height / the most lines in an area (GetFontWorkOutline :520-525, ScaleX off - the MSO import
///   default), and its width is scaled to the outline's length (CalculateHorizontalScalingFactor :197-290: fWidth / fTextWidth;
///   for a pair the mean of both lengths) - SVG textLength with lengthAdjust="spacingAndGlyphs";
/// - single-line mode: each glyph turned with the outline at its place, its centre on the outline (FitTextOutlinesToShapeOutlines
///   :756-1015: rotate about the glyph centre, move the paragraph's centre onto the curve) - textPath with a central baseline;
///   several lines in one area go on outlines moved along the normals by the line's distance from the area's centre (:806-900);
/// - pairs: every glyph point is placed between the two outlines at the same fraction of their lengths (:1020-1080). SVG cannot
///   bend a glyph's own shape without its outline, which an HTML writer does not have: the text runs along the curve halfway
///   between the pair (the same fraction-by-length rule), its size the mean gap between them - glyphs follow the path but keep
///   their shapes.
/// </summary>
public static class WordOfficeArtFontwork
{
    public static List<WordShapeTextOnPath> Layout(List<WordShapePath> geometry, string text, double height)
    {
        var outlines = new List<List<(double X, double Y)>>();
        foreach (WordShapePath path in geometry)
            foreach (WordShapeSubpath subpath in path.Subpaths)
                if (Flatten(subpath) is { Count: > 1 } points) outlines.Add(points);
        var result = new List<WordShapeTextOnPath>();
        if (outlines.Count == 0 || text.Length == 0) return result;
        bool singleLine = (outlines.Count & 1) == 1;
        int areas = singleLine ? outlines.Count : outlines.Count / 2;
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n').Split('\n');
        int maxPerArea = (lines.Length - 1) / areas + 1;
        double lineHeight = height / maxPerArea;
        int next = 0, left = lines.Length, area = 0, areasLeft = areas;
        while (left > 0 && areasLeft > 0)
        {
            int count = (left - 1) / areasLeft + 1;
            List<(double X, double Y)> curve;
            double fontSize;
            if (singleLine) { curve = outlines[area]; fontSize = lineHeight; }
            else
            {
                List<(double X, double Y)> upper = outlines[area * 2], lower = outlines[area * 2 + 1];
                curve = Between(upper, lower, 0.5);
                fontSize = MeanGap(upper, lower) / count;
            }
            for (int i = 0; i < count; i++)
            {
                // the line's distance from the area's centre, along the outline's normals (centred lines: 0 for one line)
                double offset = (i - (count - 1) / 2.0) * fontSize;
                List<(double X, double Y)> line = offset == 0 ? curve : Offset(curve, offset);
                result.Add(new WordShapeTextOnPath(PathData(line), Length(line), fontSize, lines[next + i]));
            }
            next += count; left -= count; area++; areasLeft--;
        }
        return result;
    }

    private static List<(double X, double Y)> Flatten(WordShapeSubpath subpath)
    {
        var points = new List<(double X, double Y)> { (subpath.StartX, subpath.StartY) };
        foreach (WordShapeSegment segment in subpath.Segments)
        {
            (double x0, double y0) = points[points.Count - 1];
            if (!segment.IsCurve) { points.Add((segment.X, segment.Y)); continue; }
            for (int k = 1; k <= 16; k++)
            {
                double t = k / 16.0, u = 1 - t;
                points.Add((u * u * u * x0 + 3 * u * u * t * segment.Control1X + 3 * u * t * t * segment.Control2X + t * t * t * segment.X,
                            u * u * u * y0 + 3 * u * u * t * segment.Control1Y + 3 * u * t * t * segment.Control2Y + t * t * t * segment.Y));
            }
        }
        if (subpath.Closed) points.Add(points[0]);
        return points;
    }

    private static double Length(List<(double X, double Y)> points)
    {
        double length = 0;
        for (int i = 1; i < points.Count; i++) length += Math.Sqrt(Square(points[i].X - points[i - 1].X) + Square(points[i].Y - points[i - 1].Y));
        return length;
    }

    private static double Square(double v) => v * v;

    /// <summary>The point at a fraction of the outline's length (GetPoint :728-754).</summary>
    private static (double X, double Y) At(List<(double X, double Y)> points, double fraction)
    {
        double total = Length(points), target = Math.Max(0, Math.Min(1, fraction)) * total, walked = 0;
        for (int i = 1; i < points.Count; i++)
        {
            double step = Math.Sqrt(Square(points[i].X - points[i - 1].X) + Square(points[i].Y - points[i - 1].Y));
            if (walked + step >= target && step > 0)
            {
                double t = (target - walked) / step;
                return (points[i - 1].X + (points[i].X - points[i - 1].X) * t, points[i - 1].Y + (points[i].Y - points[i - 1].Y) * t);
            }
            walked += step;
        }
        return points[points.Count - 1];
    }

    private static List<(double X, double Y)> Between(List<(double X, double Y)> upper, List<(double X, double Y)> lower, double weight)
    {
        var points = new List<(double X, double Y)>();
        for (int k = 0; k <= 64; k++)
        {
            double f = k / 64.0;
            (double ux, double uy) = At(upper, f);
            (double lx, double ly) = At(lower, f);
            points.Add((ux + (lx - ux) * weight, uy + (ly - uy) * weight));
        }
        return points;
    }

    private static double MeanGap(List<(double X, double Y)> upper, List<(double X, double Y)> lower)
    {
        double sum = 0;
        for (int k = 0; k <= 16; k++)
        {
            (double ux, double uy) = At(upper, k / 16.0);
            (double lx, double ly) = At(lower, k / 16.0);
            sum += Math.Sqrt(Square(lx - ux) + Square(ly - uy));
        }
        return sum / 17;
    }

    /// <summary>The outline moved along its normals (the neighbouring points' direction turned by 90 degrees, :826-847).</summary>
    private static List<(double X, double Y)> Offset(List<(double X, double Y)> points, double distance)
    {
        var moved = new List<(double X, double Y)>(points.Count);
        for (int i = 0; i < points.Count; i++)
        {
            (double ax, double ay) = points[i == 0 ? i : i - 1];
            (double bx, double by) = points[i == points.Count - 1 ? i : i + 1];
            double dx = bx - ax, dy = by - ay, length = Math.Sqrt(dx * dx + dy * dy);
            (double nx, double ny) = length > 0 ? (-dy / length, dx / length) : (0, 0);
            moved.Add((points[i].X + nx * distance, points[i].Y + ny * distance));
        }
        return moved;
    }

    private static string PathData(List<(double X, double Y)> points)
    {
        var d = new StringBuilder();
        for (int i = 0; i < points.Count; i++)
            d.Append(i == 0 ? 'M' : 'L').Append(WordShapeHtml.Number(points[i].X)).Append(' ').Append(WordShapeHtml.Number(points[i].Y));
        return d.ToString();
    }

    internal static string N(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);

    /// <summary>Turns <paramref name="shape"/> (its geometry evaluated) into its text effect: the text on the outlines, filled with the
    /// shape's fill (a gradient / pattern paint too), the outline itself not drawn (EnhancedCustomShapeEngine.cxx:236-262).</summary>
    public static void Apply(WordShape shape, string text, string? font)
    {
        shape.TextOnPaths = Layout(shape.Paths ?? [], text, shape.Height);
        shape.TextOnPathFont = font;
        shape.TextOnPathColor = shape.FillPaint ?? shape.FillColor ?? "#000000";
        shape.Geometry = WordShapeGeometry.Path;
        shape.Paths = [];
    }
}
