using System.Text;

namespace FastDocView.Html;

/// <summary>
/// The outline of an OfficeArt shape type (MSOSPT: RTF's shapeType property and the .doc FSP instance share the numbers -
/// "Word 97 RTF for Drawing Objects (Shapes)" ShapeType table, techshelps.github.io/MSDN/SPECS/newrtf/
/// word97rtffordrawingobjectsshapes.htm) on a <see cref="WordShape"/> of a given box. The outlines are transcribed from LibreOffice
/// svx/source/customshapes/EnhancedCustomShapeGeometry.cxx (master, read 2026-11-03), whose shapes live in a 21600 x 21600
/// space scaled to the box, with their default adjust values (adjustValue / adjust2Value in that space):
/// round rectangle (lines 173-215: corner = adjustValue on each axis, default 3600), right triangle (219-222), diamond
/// (324-327), isosceles triangle (480-518: apex at adjustValue, default 10800), arrow / left / down / up (684-819: head at
/// adjustValue, shaft at adjust2Value; defaults 16200 / 5400, left and up 5400 / 5400) and left-right arrow (827-868: defaults
/// 4300 / 5400). Rectangle 1, text box 202 and picture frame 75 are boxes, ellipse 3 an ellipse, line 20 a line. Any other
/// type is drawn as a plain box of its size (<see cref="WordShapeGeometry.Box"/>).
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

    /// <summary>Sets the geometry of <paramref name="shape"/> (its Width / Height already set) for shape type <paramref name="type"/>.</summary>
    public static void Apply(WordShape shape, int type, int? adjust1, int? adjust2)
    {
        double w = shape.Width, h = shape.Height;
        double X(double x) => x * w / 21600;
        double Y(double y) => y * h / 21600;
        string Polygon(params (double x, double y)[] points)
        {
            var d = new StringBuilder();
            for (int i = 0; i < points.Length; i++)
                d.Append(i == 0 ? 'M' : 'L').Append(WordShapeHtml.Number(X(points[i].x))).Append(' ').Append(WordShapeHtml.Number(Y(points[i].y)));
            return d.Append('Z').ToString();
        }
        void Path(string data) { shape.Geometry = WordShapeGeometry.Path; shape.PathData = data; }
        switch (type)
        {
            case Rectangle: case TextBox: case PictureFrame: shape.Geometry = WordShapeGeometry.Rectangle; break;
            case Ellipse: shape.Geometry = WordShapeGeometry.Ellipse; break;
            case Line: shape.Geometry = WordShapeGeometry.Line; break;
            case RoundRectangle:
            {
                double a = Math.Max(0, Math.Min(10800, adjust1 ?? 3600)), rx = X(a), ry = Y(a);
                string N(double v) => WordShapeHtml.Number(v);
                string arc = "A" + N(rx) + " " + N(ry) + " 0 0 1 ";
                Path("M" + N(rx) + " 0L" + N(w - rx) + " 0" + arc + N(w) + " " + N(ry) + "L" + N(w) + " " + N(h - ry) + arc + N(w - rx) + " " + N(h)
                     + "L" + N(rx) + " " + N(h) + arc + "0 " + N(h - ry) + "L0 " + N(ry) + arc + N(rx) + " 0Z");
                break;
            }
            case RightTriangle: Path(Polygon((0, 0), (21600, 21600), (0, 21600))); break;
            case Diamond: Path(Polygon((10800, 0), (21600, 10800), (10800, 21600), (0, 10800))); break;
            case IsoscelesTriangle: Path(Polygon((adjust1 ?? 10800, 0), (21600, 21600), (0, 21600))); break;
            case Arrow:
            {
                double a1 = adjust1 ?? 16200, a2 = adjust2 ?? 5400;
                Path(Polygon((0, a2), (a1, a2), (a1, 0), (21600, 10800), (a1, 21600), (a1, 21600 - a2), (0, 21600 - a2)));
                break;
            }
            case LeftArrow:
            {
                double a1 = adjust1 ?? 5400, a2 = adjust2 ?? 5400;
                Path(Polygon((21600, a2), (a1, a2), (a1, 0), (0, 10800), (a1, 21600), (a1, 21600 - a2), (21600, 21600 - a2)));
                break;
            }
            case DownArrow:
            {
                double a1 = adjust1 ?? 16200, a2 = adjust2 ?? 5400;
                Path(Polygon((a2, 0), (a2, a1), (0, a1), (10800, 21600), (21600, a1), (21600 - a2, a1), (21600 - a2, 0)));
                break;
            }
            case UpArrow:
            {
                double a1 = adjust1 ?? 5400, a2 = adjust2 ?? 5400;
                Path(Polygon((a2, 21600), (a2, a1), (0, a1), (10800, 0), (21600, a1), (21600 - a2, a1), (21600 - a2, 21600)));
                break;
            }
            case LeftRightArrow:
            {
                double a1 = adjust1 ?? 4300, a2 = adjust2 ?? 5400;
                Path(Polygon((0, 10800), (a1, 0), (a1, a2), (21600 - a1, a2), (21600 - a1, 0), (21600, 10800), (21600 - a1, 21600),
                             (21600 - a1, 21600 - a2), (a1, 21600 - a2), (a1, 21600)));
                break;
            }
            default: shape.Geometry = WordShapeGeometry.Box; break;
        }
    }
}
