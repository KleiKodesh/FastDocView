using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>
/// Contour wrapping for every format: the text runs along an object's outline instead of its box (.docx wp:wrapTight /
/// wp:wrapThrough with wp:wrapPolygon, VML wrapcoords, .doc / RTF pWrapPolygonVertices, ODF style:wrap-contour). Drawn as CSS
/// shape-outside on the float (CSS Shapes 1): the outline's extreme on the text side at every point of height (1 pt rows, at most
/// 2000), so the text meets it line by line, with the gap on the text side as shape-margin. A float keeps its text on one side
/// (CSS Shapes 1, section 2), so "through" (text also into the outline's indentations) is drawn as tight.
/// </summary>
public static class WordPictureContour
{
    /// <summary>The size of the coordinate space of Word's wrap polygons: (0,0) is the object's top-left, (21600,21600) its bottom-right.</summary>
    /// <remarks>
    /// [MS-OI29500] (msdn ee342530) as quoted by LibreOffice sw/source/writerfilter/dmapper/WrapPolygonHandler.cxx:120-129
    /// ("Image wrapping polygons in Microsoft Word use a fixed coordinate space that is 21600 units x 21600 units"); the same
    /// space for .doc / RTF pWrapPolygonVertices (sw/source/filter/ww8/ww8graf.cxx:2117-2128 scales by ww::nWrap100Percent).
    /// LibreOffice's further 15-twip shift (WrapPolygonHandler.cxx:99-114, ww8graf.cxx:2075-2125 "I have no idea") is its own
    /// guess for its contour engine and is not applied here.
    /// </remarks>
    public const double WordUnits = 21600;

    /// <summary>A Word wrap polygon (21600 units) in the object's box points; null when it has fewer than 3 points.</summary>
    public static IReadOnlyList<IReadOnlyList<(double X, double Y)>>? FromWordUnits(IReadOnlyList<(double X, double Y)>? points, double widthPt, double heightPt)
    {
        if (points is not { Count: > 2 } || !(widthPt > 0) || !(heightPt > 0)) return null;
        var outline = new List<(double X, double Y)>(points.Count);
        foreach ((double x, double y) in points) outline.Add((x / WordUnits * widthPt, y / WordUnits * heightPt));
        return [outline];
    }

    /// <summary>Points "x,y x,y ..." / "x y x y" / "x,y,x,y" (any mix of commas and spaces) as numbers in pairs; null when fewer than 3 pairs.</summary>
    public static List<(double X, double Y)>? ParsePairs(string? value)
    {
        if (value is null) return null;
        var numbers = new List<double>();
        foreach (string part in value.Split(new[] { ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            if (double.TryParse(part, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && !double.IsNaN(number) && !double.IsInfinity(number))
                numbers.Add(number);
            else return null;
        if (numbers.Count < 6 || numbers.Count % 2 != 0) return null;
        var points = new List<(double X, double Y)>(numbers.Count / 2);
        for (int i = 0; i < numbers.Count; i += 2) points.Add((numbers[i], numbers[i + 1]));
        return points;
    }

    /// <summary>
    /// The shape-outside / shape-margin CSS of a float of <paramref name="width"/> x <paramref name="height"/> points whose text is on
    /// its right (<paramref name="textOnRight"/>) or left, from its outlines in its own box points; null when nothing usable.
    /// </summary>
    public static string? ShapeOutsideCss(IReadOnlyList<IReadOnlyList<(double X, double Y)>> outlines, double width, double height, bool textOnRight, double gap)
    {
        if (!(width > 0) || !(height > 0) || double.IsInfinity(width) || double.IsInfinity(height) || outlines.Count == 0) return null;
        double step = Math.Max(1, height / 2000);
        var polygon = new List<(double X, double Y)>();
        for (double y = 0; ; y = Math.Min(height, y + step))
        {
            double? extreme = null;
            foreach (IReadOnlyList<(double X, double Y)> outline in outlines)
                for (int i = 0; i < outline.Count; i++)
                {
                    (double x1, double y1) = outline[i]; (double x2, double y2) = outline[(i + 1) % outline.Count];
                    if (y < Math.Min(y1, y2) || y > Math.Max(y1, y2)) continue;
                    double x = y1 == y2 ? (textOnRight ? Math.Max(x1, x2) : Math.Min(x1, x2)) : x1 + (y - y1) / (y2 - y1) * (x2 - x1);
                    extreme = extreme is null ? x : textOnRight ? Math.Max(extreme.Value, x) : Math.Min(extreme.Value, x);
                }
            if (extreme is { } e) polygon.Add((Math.Max(0, Math.Min(width, e)), y));
            if (y >= height) break;
        }
        if (polygon.Count < 2) return null;
        double edge = textOnRight ? 0 : width;
        var css = new StringBuilder("shape-outside:polygon(");
        css.Append(Css(edge)).Append("pt ").Append(Css(polygon[0].Y)).Append("pt");
        foreach ((double x, double y) in polygon) css.Append(',').Append(Css(x)).Append("pt ").Append(Css(y)).Append("pt");
        css.Append(',').Append(Css(edge)).Append("pt ").Append(Css(polygon[polygon.Count - 1].Y)).Append("pt) border-box;");
        if (gap > 0) css.Append("shape-margin:").Append(Css(gap)).Append("pt;");
        return css.ToString();
    }

    private static string Css(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);
}
