using System.Text;

namespace FastDocView.Html;

/// <summary>
/// OfficeArt shape types (MSOSPT: RTF's shapeType property and the .doc FSP instance share the numbers - "Word 97 RTF for Drawing
/// Objects (Shapes)" ShapeType table): every type LibreOffice has a table for is drawn by <see cref="WordOfficeArtGeometry"/> from
/// <see cref="WordOfficeArtShapeDefinitions"/> (generated from svx EnhancedCustomShapeGeometry.cxx); a type without one is a box
/// of its size. Also the line dashes, line ends and default fill / stroke rules shared by RTF and .doc.
/// </summary>
public static class WordOfficeArtShapeTypes
{
    public const int Rectangle = 1, RoundRectangle = 2, Ellipse = 3, Diamond = 4, IsoscelesTriangle = 5, RightTriangle = 6, Arrow = 13,
                     Line = 20, LeftArrow = 66, DownArrow = 67, UpArrow = 68, LeftRightArrow = 69, PictureFrame = 75, TextBox = 202;

    /// <summary>
    /// The SVG dash pattern (multiples of the line width; null = solid) of an OfficeArt lineDashing value (MSO_LineDashing,
    /// LibreOffice include/svx/msdffdef.hxx:696-708), as LibreOffice draws it (filter/source/msfilter/msdffimp.cxx:941-1036:
    /// one "dot" of dotLen, "dashes" of dashLen, each followed by the distance, in percent of the line width; defaults dotLen
    /// 300, distance 300).
    /// </summary>
    public static double[]? Dash(int lineDashing)
    {
        // (dotLen, dashes, dashLen, distance) per msdffimp.cxx:962-1033
        (double dot, int dashes, double dash, double distance) = lineDashing switch
        {
            0 => (0, 0, 0, 0),
            1 => (3, 0, 0, 1),        // mso_lineDashSys
            5 => (1, 0, 0, 3),        // mso_lineDotGEL
            6 => (4, 0, 0, 3),        // mso_lineDashGEL
            7 => (8, 0, 0, 3),        // mso_lineLongDashGEL
            8 => (4, 1, 1, 3),        // mso_lineDashDotGEL
            9 => (8, 1, 1, 3),        // mso_lineLongDashDotGEL
            10 => (8, 2, 1, 3),       // mso_lineLongDashDotDotGEL
            3 => (3, 1, 1, 1),        // mso_lineDashDotSys
            4 => (3, 2, 1, 1),        // mso_lineDashDotDotSys
            _ => (1, 0, 0, 1),        // mso_lineDotSys and the switch's default
        };
        if (dot == 0) return null;
        var pattern = new List<double> { dot, distance };
        for (int i = 0; i < dashes; i++) { pattern.Add(dash); pattern.Add(distance); }
        return pattern.ToArray();
    }

    /// <summary>OfficeArt fills a shape type by default unless LibreOffice svx EnhancedCustomShapeGeometry.cxx:6155-6167 (mso_DefaultFillingTable) says not.</summary>
    public static bool IsFilledByDefault(int type)
    {
        ushort[] notFilled = [0x0000, 0x0018, 0x01ff, 0x0000, 0x0c00, 0x01e0, 0x0000, 0x0000, 0x0000, 0x0000, 0x0000, 0x0600, 0x0000, 0x0000, 0x0000, 0x0000];
        return type is < 0 or >= 0x100 || (notFilled[type >> 4] & (1 << (type & 0xF))) == 0;
    }

    /// <summary>OfficeArt strokes a shape type by default unless EnhancedCustomShapeGeometry.cxx:6197-6213 (mso_DefaultStrokingTable: only 75, the picture frame) says not.</summary>
    public static bool IsStrokedByDefault(int type) => type != PictureFrame;

    /// <summary>Sets the outline of <paramref name="shape"/> (Width / Height set) for shape type <paramref name="type"/> with adjust values (null = the type's default).</summary>
    public static void Apply(WordShape shape, int type, params int?[] adjust)
    {
        if (WordOfficeArtGeometry.Preset(type, shape.Width, shape.Height, adjust) is { } paths)
        {
            shape.Geometry = WordShapeGeometry.Path;
            shape.Paths = paths;
        }
        else
        {
            // a type without a table: a plain box of its size, as paths so fills and line ends draw as for every other shape
            shape.Geometry = WordShapeGeometry.Path;
            var box = new WordShapeSubpath { Closed = true };
            box.Segments.Add(WordShapeSegment.Line(shape.Width, 0));
            box.Segments.Add(WordShapeSegment.Line(shape.Width, shape.Height));
            box.Segments.Add(WordShapeSegment.Line(0, shape.Height));
            var path = new WordShapePath();
            path.Subpaths.Add(box);
            shape.Paths = [path];
        }
    }

    /// <summary>
    /// A line end (MSOLINEEND, [MS-ODRAW] 2.4.16: 1 arrow, 2 stealth, 3 diamond, 4 oval, 5 open arrow; 6 / 7 "MUST be ignored";
    /// the RTF spec's own list numbers the open arrow 6, against [MS-ODRAW] and LibreOffice msdffdef.hxx:711-718 - not followed) with
    /// its width (MSOLINEENDWIDTH 2.4.17: narrow, medium, wide) and length (2.4.18: short, medium, long), drawn as LibreOffice
    /// draws a binary arrow (msdffimp.cxx:786-900 GetLineArrow: a line under 2 pt has the 2 pt arrow; width / length 2, 3, 5 times
    /// it, the open arrow 3.5 / 4.5 / 6; diamond and oval centred on the line's end).
    /// </summary>
    internal static (WordLineEnd end, WordLineEndMarker? marker) LineEnd(int type, int width, int length, double lineWidthPt)
    {
        if (type is < 1 or > 5) return (default, null);
        double line = Math.Max(lineWidthPt, 2);
        double lengthFactor = length switch { 0 => 2, 2 => 5, _ => 3 }, widthFactor = width switch { 0 => 2, 2 => 5, _ => 3 };
        if (type == 5) { lengthFactor = length switch { 0 => 3.5, 2 => 6, _ => 4.5 }; widthFactor = width switch { 0 => 3.5, 2 => 6, _ => 4.5 }; }
        double w = widthFactor * line, l = lengthFactor * line;
        List<(double X, double Y)> points = type switch
        {
            1 => [(w * 0.5, 0), (w, l), (0, l)],
            2 => [(w * 0.5, 0), (w, l), (w * 0.5, l * 0.6), (0, l)],
            3 => [(w * 0.5, 0), (w, l * 0.5), (w * 0.5, l), (0, l * 0.5)],
            5 => [(w * 0.5, 0), (w, l * 0.91), (w * 0.85, l), (w * 0.5, l * 0.36), (w * 0.15, l), (0, l * 0.91)],
            _ => Enumerable.Range(0, 33).Select(i => (w * 0.5 + w * 0.5 * Math.Cos(i * Math.PI / 16), l * 0.5 + l * 0.5 * Math.Sin(i * Math.PI / 16))).ToList(),
        };
        bool centred = type is 3 or 4;
        double docking = centred ? 0.5 : 0;
        var marker = new WordLineEndMarker(points.Select(p => (p.X - w / 2, p.Y - l * docking)).ToList(), w, l * (1 - docking));
        WordLineEndType kind = type switch { 1 => WordLineEndType.Triangle, 2 => WordLineEndType.Stealth, 3 => WordLineEndType.Diamond, 4 => WordLineEndType.Oval, _ => WordLineEndType.Arrow };
        return (new WordLineEnd(kind, width, length), marker);
    }

    /// <summary>Sets both line ends of <paramref name="shape"/> (its LineWidthPt set) from lineStart / lineEnd Arrowhead, ArrowWidth, ArrowLength (defaults: none, medium, medium - msdffimp.cxx:1071-1073).</summary>
    internal static void ApplyLineEnds(WordShape shape, int? startType, int? startWidth, int? startLength, int? endType, int? endWidth, int? endLength)
    {
        (shape.HeadEnd, shape.HeadMarker) = LineEnd(startType ?? 0, startWidth ?? 1, startLength ?? 1, shape.LineWidthPt);
        (shape.TailEnd, shape.TailMarker) = LineEnd(endType ?? 0, endWidth ?? 1, endLength ?? 1, shape.LineWidthPt);
        shape.HeadArrow = shape.TailArrow = false;
    }
}
