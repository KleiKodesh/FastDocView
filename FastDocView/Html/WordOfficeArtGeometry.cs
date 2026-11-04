namespace FastDocView.Html;

/// <summary>
/// Evaluates an OfficeArt outline - a preset from <see cref="WordOfficeArtShapeDefinitions"/> or a shape's own geometry
/// (pVertices / pSegmentInfo / pGuides) - into <see cref="WordShapePath"/>s in the shape's points, the way LibreOffice draws it
/// (svx, master read 2026-11-03):
/// - Equations: EnhancedCustomShape2d.cxx:88-330 GetEquation (operations 0..16, 0x80..0x82; a parameter is special when flag
///   0x2000 / 0x4000 / 0x8000 is set: 0x400 | n = equation n, adjustValue..adjust10Value = adjustment, geoLeft..geoBottom =
///   the coordinate space) and :332-400 AppendEnhancedCustomShapeEquationParameter.
/// - Coordinates: a vertex whose value has the MSO_I mark (EnhancedCustomShapeGeometry.hxx:88; :72-86 SetEnhancedCustomShapeParameter)
///   is equation n's result; the coordinate space (0, 0)..(coordinate width, height) is scaled to the box.
/// - Segments: svdoashp.cxx:96-165 lcl_ShapeSegmentFromBinary (0x00 line, 0x20 curve, 0x40 move, 0x60 close, 0x80 end,
///   0xa1 / 0xa2 angle ellipse (to), 0xa3..0xa6 arc (to) / clockwise, 0xa7 / 0xa8 elliptical quadrant x / y, 0xaa no fill,
///   0xab no stroke); drawn as EnhancedCustomShape2d.cxx:2078-2690 CreateSubPath (no segments = one closed polygon; ARC from
///   tools/source/generic/poly.cxx:246-330 with ImplGetParameter :60-67; ANGLEELLIPSE ODF rules and the preset-with-U full
///   ellipses :2189-2340; quadrants :2490-2590), one path per sub-path object with its own fill / stroke.
/// Arcs are flattened to short lines (LibreOffice draws polygons too).
/// </summary>
public static class WordOfficeArtGeometry
{
    private const int GeoLeft = 320, GeoTop = 321, GeoRight = 322, GeoBottom = 323, AdjustFirst = 327, AdjustLast = 336;
    private const int EquationMark = 0x40000000;

    /// <summary>Shape types whose angle ellipse LibreOffice always draws whole (EnhancedCustomShape2d.cxx:2203-2205 with
    /// EnhancedCustomShapeTypeNames.cxx:38, 58, 92, 131, 141, 155, 158, 159, 218).</summary>
    private static readonly HashSet<int> FullEllipseTypes = [3, 23, 57, 96, 106, 120, 123, 124, 183];

    /// <summary>A preset's paths for a box of <paramref name="width"/> x <paramref name="height"/> points; null when the type has no table.</summary>
    public static List<WordShapePath>? Preset(int shapeType, double width, double height, IReadOnlyList<int?> adjust)
    {
        if (WordOfficeArtShapeDefinitions.Of(shapeType) is not { } definition) return null;
        var values = new double[10];
        for (int i = 0; i < 10; i++) values[i] = i < adjust.Count && adjust[i] is { } given ? given : i < definition.Defaults.Length ? definition.Defaults[i] : 0;
        var vertices = new (int, bool)[definition.Vertices.Length];
        for (int i = 0; i < vertices.Length; i++)
        {
            int v = definition.Vertices[i];
            vertices[i] = (v & EquationMark) != 0 && v > 0 ? (v & 0xFFFF, true) : (v, false);
        }
        var segments = new List<(int command, int count)>(definition.Segments.Length);
        foreach (ushort data in definition.Segments) segments.Add(TableSegment(data));
        var evaluator = new Equations(definition.Calculations, values, 0, 0, definition.CoordinateWidth, definition.CoordinateHeight);
        return Build(vertices, segments, evaluator, 0, 0, definition.CoordinateWidth, definition.CoordinateHeight, width, height,
                     angleMode: FullEllipseTypes.Contains(shapeType) ? AngleMode.Full : AngleMode.Odf);
    }

    /// <summary>
    /// A shape's own geometry (FOPT pVertices 325, pSegmentInfo 326, pGuides 342; geoLeft..geoBottom 320-323): vertices as
    /// (value, isEquation) pairs - an equation reference is a vertex value with 0x80000000 set (the same mark as the presets),
    /// segments, guide records (flags, p1, p2, p3); angle ellipses use the binary-import rule (degrees in 16.16, a swing angle,
    /// EnhancedCustomShape2d.cxx:2246-2297).
    /// </summary>
    public static List<WordShapePath> Custom(IReadOnlyList<int> vertexValues, IReadOnlyList<ushort> segments, int[] guides,
                                             IReadOnlyList<int?> adjust, int left, int top, int right, int bottom, double width, double height)
    {
        var values = new double[10];
        for (int i = 0; i < 10; i++) values[i] = i < adjust.Count && adjust[i] is { } given ? given : 0;
        // a value whose high 16 bits are 0x8000 is equation n (EnhancedCustomShape2d.cxx:72-86 SetEnhancedCustomShapeParameter)
        var vertices = vertexValues.Select(v => ((uint)v >> 16) == 0x8000 ? (v & 0xFFFF, true) : (v, false)).ToArray();
        var evaluator = new Equations(guides, values, left, top, right, bottom);
        return Build(vertices, segments.Select(BinarySegment).ToList(), evaluator, left, top, right, bottom, width, height, AngleMode.Binary);
    }

    /// <summary>A preset table's segment (svdoashp.cxx:96-165 lcl_ShapeSegmentFromBinary) as (command byte, element count).</summary>
    private static (int command, int count) TableSegment(ushort data)
    {
        int command = data >> 8, low = data & 0xFF;
        return command switch
        {
            0x00 or 0x20 or 0x40 => (command, Math.Max(1, low)),
            0xA1 or 0xA2 => (command, low / 3),
            0xA3 or 0xA4 or 0xA5 or 0xA6 => (command, low >> 2),
            _ => (command, low),
        };
    }

    /// <summary>
    /// A shape's own pSegmentInfo element (msdffimp.cxx:2267-2390: the top 3 bits the command - line, curve, move, close, end -
    /// and 13 bits the count; 5 / 6 = an escape whose next 5 bits pick extension (a line), angle ellipse to / angle ellipse, arc to,
    /// arc, clockwise arc to / arc, quadrant x / y, no fill, no stroke) as (command byte, element count). RTF writes the same values
    /// (rtfsdrimport.cxx:531-590 reads 0x0001 / 0x4000 / 0x2000+n / 0x6001 / 0x8000 / 0xaa00 / 0xab00 alike).
    /// </summary>
    private static (int command, int count) BinarySegment(ushort data)
    {
        int count = data & 0x1FFF;
        switch (data >> 13)
        {
            case 0: return (0x00, Math.Max(1, count));
            case 1: return (0x20, Math.Max(1, count));
            case 2: return (0x40, Math.Max(1, count));
            case 3: return (0x60, 0);
            case 4: return (0x80, 0);
            case 5: case 6:
                return ((data >> 8) & 0x1F) switch
                {
                    0x0 => (0x00, Math.Max(1, count)),
                    0x1 => (0xA1, (data & 0xFF) / 3),
                    0x2 => (0xA2, (data & 0xFF) / 3),
                    0x3 => (0xA3, (data & 0xFF) >> 2),
                    0x4 => (0xA4, (data & 0xFF) >> 2),
                    0x5 => (0xA5, (data & 0xFF) >> 2),
                    0x6 => (0xA6, (data & 0xFF) >> 2),
                    0x7 => (0xA7, data & 0xFF),
                    0x8 => (0xA8, data & 0xFF),
                    0xA => (0xAA, 0),
                    0xB => (0xAB, 0),
                    _ => (-1, 0),
                };
            default: return (-1, 0);
        }
    }

    private enum AngleMode { Odf, Full, Binary }

    private sealed class Equations(int[] records, double[] adjust, double left, double top, double right, double bottom)
    {
        private readonly double?[] _results = new double?[records.Length / 4];
        private readonly bool[] _active = new bool[records.Length / 4];

        public double Result(int index)
        {
            if (index < 0 || index >= _results.Length) return 0;
            if (_results[index] is { } done) return done;
            if (_active[index]) return 0;                       // a cycle: LibreOffice's evaluation fails it to 0
            _active[index] = true;
            double value = Evaluate(index);
            _active[index] = false;
            if (double.IsNaN(value) || double.IsInfinity(value)) value = 0;
            _results[index] = value;
            return value;
        }

        private double Parameter(int value, bool isSpecial)
        {
            if (!isSpecial) return value;
            if ((value & 0x400) != 0) return Result(value & 0xFF);
            if (value is >= AdjustFirst and <= AdjustLast) return adjust[value - AdjustFirst];
            return value switch { GeoLeft => left, GeoTop => top, GeoRight => right, GeoBottom => bottom, _ => 0 };
        }

        private double Evaluate(int index)
        {
            int flags = records[index * 4];
            bool s1 = (flags & 0x2000) != 0, s2 = (flags & 0x4000) != 0, s3 = (flags & 0x8000) != 0;
            int r1 = records[index * 4 + 1], r2 = records[index * 4 + 2], r3 = records[index * 4 + 3];
            double p1 = Parameter(r1, s1), p2 = Parameter(r2, s2), p3 = Parameter(r3, s3);
            const double Degree = Math.PI / 180;
            switch (flags & 0xFF)
            {
                case 0: case 14: return p1 + p2 - p3;
                case 1:
                {
                    double result = p1;
                    if (s2 || r2 != 1) result *= p2;
                    if (s3 || (r3 != 1 && r3 != 0)) result /= p3;
                    return result;
                }
                case 2: return (p1 + p2) / 2;
                case 3: return Math.Abs(p1);
                case 4: return Math.Min(p1, p2);
                case 5: return Math.Max(p1, p2);
                case 6: return p1 > 0 ? p2 : p3;                   // ODF if(): the condition is positive
                case 7: return Math.Sqrt(p1 * p1 + p2 * p2 + p3 * p3);
                case 8: return Math.Atan2(p2, p1) / Degree;
                case 9: return p1 * Math.Sin(p2 * Degree);
                case 10: return p1 * Math.Cos(p2 * Degree);
                case 11: return p1 * Math.Cos(Math.Atan2(p3, p2));
                case 12: return p1 * Math.Sin(Math.Atan2(p3, p2));
                case 13: return Math.Sqrt(p1);
                case 15: return p3 * Math.Sqrt(1 - p1 / p2 * (p1 / p2));
                case 16: return p1 * Math.Tan(p2);
                case 0x80: return Math.Sqrt(p3 * p3 - p1 * p1);
                case 0x81: return Math.Cos(p3 * Degree) * (p1 - 10800) + Math.Sin(p3 * Degree) * (p2 - 10800) + 10800;
                case 0x82: return -(Math.Sin(p3 * Degree) * (p1 - 10800) - Math.Cos(p3 * Degree) * (p2 - 10800)) + 10800;
                default: return 0;
            }
        }
    }

    private static List<WordShapePath> Build((int value, bool isEquation)[] vertices, List<(int command, int count)> segments, Equations equations,
                                             double left, double top, double right, double bottom, double width, double height, AngleMode angleMode)
    {
        double scaleX = right != left ? width / (right - left) : 0, scaleY = bottom != top ? height / (bottom - top) : 0;
        double Raw(int index, bool isX)
        {
            (int value, bool isEquation) = vertices[index];
            return isEquation ? equations.Result(value) : value;
        }
        (double X, double Y) Point(int pair) => ((Raw(pair * 2, true) - left) * scaleX, (Raw(pair * 2 + 1, false) - top) * scaleY);
        int pointCount = vertices.Length / 2;
        var paths = new List<WordShapePath>();
        if (segments.Count == 0)
        {
            var path = new WordShapePath();
            if (pointCount > 0)
            {
                var only = new WordShapeSubpath { Closed = true };
                (only.StartX, only.StartY) = Point(0);
                for (int i = 1; i < pointCount; i++) { var p = Point(i); only.Segments.Add(WordShapeSegment.Line(p.X, p.Y)); }
                path.Subpaths.Add(only);
            }
            paths.Add(path);
            return paths;
        }
        int source = 0, segment = 0;
        while (segment < segments.Count)
        {
            var path = new WordShapePath();
            WordShapeSubpath? current = null;
            void Finish() { if (current is { Segments.Count: > 0 }) path.Subpaths.Add(current); current = null; }
            void LineTo(double x, double y)
            {
                if (current is null) { current = new WordShapeSubpath { StartX = x, StartY = y }; return; }
                current.Segments.Add(WordShapeSegment.Line(x, y));
            }
            (double X, double Y) Last() => current is null ? (0, 0) : current.Segments.Count > 0 ? (current.Segments[current.Segments.Count - 1].X, current.Segments[current.Segments.Count - 1].Y) : (current.StartX, current.StartY);
            bool ended = false;
            while (segment < segments.Count && !ended)
            {
                (int command, int count) = segments[segment++];
                switch (command)
                {
                    case 0x00:
                        for (int i = 0; i < count && source < pointCount; i++) { var p = Point(source++); LineTo(p.X, p.Y); }
                        break;
                    case 0x20:
                        for (int i = 0; i < count && source + 2 < pointCount; i++)
                        {
                            var a = Point(source++); var b = Point(source++); var e = Point(source++);
                            current ??= new WordShapeSubpath { StartX = a.X, StartY = a.Y };
                            current.Segments.Add(new WordShapeSegment(true, a.X, a.Y, b.X, b.Y, e.X, e.Y));
                        }
                        break;
                    case 0x40:
                        Finish();
                        if (source < pointCount) { var p = Point(source++); current = new WordShapeSubpath { StartX = p.X, StartY = p.Y }; }
                        break;
                    case 0x60:
                        if (current is not null) { current.Closed = true; Finish(); }
                        break;
                    case 0x80: ended = true; break;
                    case 0xA1: case 0xA2:
                        for (int i = 0; i < count && source + 2 < pointCount; i++)
                        {
                            if (command == 0xA2) Finish();
                            var centre = Point(source);
                            double rx = Raw((source + 1) * 2, true) * scaleX, ry = Raw((source + 1) * 2 + 1, false) * scaleY;
                            double a1 = Raw((source + 2) * 2, true), a2 = Raw((source + 2) * 2 + 1, false);
                            source += 3;
                            foreach (var p in AngleEllipse(centre, rx, ry, a1, a2, angleMode, rx / Math.Max(scaleX, 1e-9), ry / Math.Max(scaleY, 1e-9))) LineTo(p.X, p.Y);
                        }
                        break;
                    case 0xA3: case 0xA4: case 0xA5: case 0xA6:
                    {
                        bool clockwise = command is 0xA5 or 0xA6, moves = command is 0xA4 or 0xA6;
                        for (int i = 0; i < count && source + 3 < pointCount; i++)
                        {
                            if (moves) Finish();
                            var c1 = Point(source); var c2 = Point(source + 1);
                            var start = Point(source + (clockwise ? 3 : 2)); var end = Point(source + (clockwise ? 2 : 3));
                            source += 4;
                            List<(double X, double Y)> arc = Arc(c1, c2, start, end);
                            if (clockwise) arc.Reverse();
                            foreach (var p in arc) LineTo(p.X, p.Y);
                        }
                        break;
                    }
                    case 0xA7: case 0xA8:
                    {
                        bool isX = command == 0xA7;
                        for (int i = 0; i < count && source < pointCount; i++)
                        {
                            var from = Last();
                            if (current is null) { var p0 = Point(source++); LineTo(p0.X, p0.Y); continue; }
                            var to = Point(source++);
                            foreach (var p in Quadrant(from, to, isX)) LineTo(p.X, p.Y);
                            isX = !isX;
                        }
                        break;
                    }
                    case 0xAA: path.Fill = WordShapePathFill.None; break;
                    case 0xAB: path.Stroke = false; break;
                }
            }
            Finish();
            if (path.Subpaths.Count > 0) paths.Add(path);
        }
        return paths;
    }

    /// <summary>Points of an ellipse arc from <paramref name="start"/> to <paramref name="end"/> radians, increasing (basegfx createPolygonFromEllipseSegment: y down).</summary>
    private static IEnumerable<(double X, double Y)> EllipseSegment((double X, double Y) centre, double rx, double ry, double start, double end)
    {
        if (end <= start) end += 2 * Math.PI;
        int steps = Math.Max(4, (int)Math.Ceiling((end - start) / (Math.PI / 32)));
        for (int i = 0; i <= steps; i++)
        {
            double t = start + (end - start) * i / steps;
            yield return (centre.X + rx * Math.Cos(t), centre.Y + ry * Math.Sin(t));
        }
    }

    private static double Normalized(double degrees)
    {
        double d = degrees % 360;
        if (d < 0) d += 360;
        return d * Math.PI / 180;
    }

    /// <summary>lcl_getNormalizedCircleAngleRad (EnhancedCustomShape2d.cxx:2051-2068).</summary>
    private static double CircleAngle(double wr, double hr, double degrees)
    {
        double d = degrees % 360;
        if (d < 0) d += 360;
        if (d is 0 or 90 or 180 or 270) return d * Math.PI / 180;
        double x = hr * Math.Cos(d * Math.PI / 180), y = wr * Math.Sin(d * Math.PI / 180);
        double result = Math.Atan2(y, x);
        return result < 0 ? result + 2 * Math.PI : result;
    }

    private static List<(double X, double Y)> AngleEllipse((double X, double Y) centre, double rx, double ry, double first, double second, AngleMode mode, double wr, double hr)
    {
        var points = new List<(double X, double Y)>();
        if (rx == 0 && ry == 0) { points.Add(centre); return points; }
        switch (mode)
        {
            case AngleMode.Full:
                points.AddRange(EllipseSegment(centre, rx, ry, 0, 2 * Math.PI));
                break;
            case AngleMode.Binary:
            {
                double startAngle = -first / 65536.0, swing = -second / 65536.0, endAngle = startAngle + swing;
                if (swing < 0) (startAngle, endAngle) = (endAngle, startAngle);
                double from = startAngle, to = from + 180;
                while (to < endAngle) { points.AddRange(EllipseSegment(centre, rx, ry, Normalized(from), Normalized(to))); from = to; to += 180; }
                points.AddRange(EllipseSegment(centre, rx, ry, Normalized(from), Normalized(endAngle)));
                if (swing < 0) points.Reverse();
                break;
            }
            default:
                if (Math.Abs(Math.Abs(second - first) - 360) < 1e-15)
                {
                    double s = CircleAngle(wr, hr, first), h = CircleAngle(wr, hr, first + 180), e = CircleAngle(wr, hr, second);
                    points.AddRange(EllipseSegment(centre, rx, ry, s, h));
                    points.AddRange(EllipseSegment(centre, rx, ry, h, e));
                }
                else points.AddRange(EllipseSegment(centre, rx, ry, CircleAngle(wr, hr, first), CircleAngle(wr, hr, second)));
                break;
        }
        return points;
    }

    /// <summary>tools::Polygon(rect, start, end, PolyStyle::Arc) counter-clockwise on screen (poly.cxx:246-330).</summary>
    private static List<(double X, double Y)> Arc((double X, double Y) a, (double X, double Y) b, (double X, double Y) start, (double X, double Y) end)
    {
        var points = new List<(double X, double Y)>();
        double left = Math.Min(a.X, b.X), right = Math.Max(a.X, b.X), top = Math.Min(a.Y, b.Y), bottom = Math.Max(a.Y, b.Y);
        if (right - left <= 0 || bottom - top <= 0) return points;
        // the caller normalises the rectangle first (EnhancedCustomShape2d.cxx:2402 Rectangle::Normalize), so CreateArc's swap never applies
        double cx = (left + right) / 2, cy = (top + bottom) / 2, rx = cx - left, ry = cy - top;
        double Parameter((double X, double Y) p) { double angle = Math.Atan2(cy - p.Y, p.X - cx); return Math.Atan2(rx * Math.Sin(angle), ry * Math.Cos(angle)); }
        double from = Parameter(start), diff = Parameter(end) - from;
        if (diff <= 0) diff += 2 * Math.PI;
        int steps = Math.Max(16, (int)Math.Ceiling(diff / (Math.PI / 32)));
        for (int i = 0; i <= steps; i++)
        {
            double t = from + diff * i / steps;
            points.Add((cx + rx * Math.Cos(t), cy - ry * Math.Sin(t)));
        }
        return points;
    }

    /// <summary>ELLIPTICALQUADRANTX / Y (EnhancedCustomShape2d.cxx:2490-2590): a quarter ellipse from the current point, leaving along x (or y) first.</summary>
    private static List<(double X, double Y)> Quadrant((double X, double Y) from, (double X, double Y) to, bool isX)
    {
        double rx = Math.Abs(to.X - from.X), ry = Math.Abs(to.Y - from.Y);
        (double X, double Y) centre = isX ? (from.X, to.Y) : (to.X, from.Y);
        double half = Math.PI / 2;
        double start, end;
        bool flip;
        if (isX)
        {
            if (to.X < from.X) { if (to.Y < from.Y) { start = half; end = Math.PI; flip = false; } else { start = Math.PI; end = 1.5 * Math.PI; flip = true; } }
            else { if (to.Y < from.Y) { start = 0; end = half; flip = true; } else { start = 1.5 * Math.PI; end = 2 * Math.PI; flip = false; } }
        }
        else
        {
            if (to.X < from.X) { if (to.Y < from.Y) { start = 1.5 * Math.PI; end = 2 * Math.PI; flip = true; } else { start = 0; end = half; flip = false; } }
            else { if (to.Y < from.Y) { start = Math.PI; end = 1.5 * Math.PI; flip = false; } else { start = half; end = Math.PI; flip = true; } }
        }
        var points = EllipseSegment(centre, rx, ry, start, end).ToList();
        if (flip) points.Reverse();
        return points;
    }
}
