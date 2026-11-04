using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// DrawingML geometry (ECMA-376 Part 1 20.1.9: a:prstGeom / a:custGeom) evaluated into <see cref="WordShapePath"/>s in the
/// shape's points, as LibreOffice does it (all checked online 2026-11-03, LibreOffice core master):
/// - presets: every ST_ShapeType from ECMA-376 Part 1 annex D presetShapeDefinitions.xml, the copy LibreOffice ships
///   (oox/source/drawingml/customshapes/presetShapeDefinitions.xml), embedded and read once; a shape's a:avLst replaces the
///   preset's adjust values by name.
/// - guides (a:avLst then a:gdLst, in order): the formula operators of oox/source/drawingml/customshapegeometry.cxx:71-90
///   with LibreOffice's meaning (:486-606): "*/" x*y/z, "+-" x+y-z, "+/" (x+y)/z, "?:" / ifelse x>0 ? y : z, abs, at2
///   atan2(y, x) in 60000ths of a degree, cat2 x*cos(atan2(z, y)), sat2 x*sin(atan2(z, y)), cos / sin / tan x*f(y) with y in
///   60000ths of a degree, max, min, mod sqrt(x^2+y^2+z^2), pin (if x>y then x, else if z>y then y else z), sqrt, val.
///   Arguments: the built-in names first (:180-372: l t r b w h hc vc ls ss, wd2..wd32, hd2..hd10, ssd2..ssd32, cd2 cd3 cd4
///   cd8 3cd4 3cd8 5cd8 7cd8), then a number, then an adjust value, then a guide. The shape's size is its own a:ext in EMU.
/// - paths: a:path w / h, when both are set, scale that path from its own coordinate space (EnhancedCustomShape2d.cxx
///   SetPathSize :650-700, SubViewSize customshapeproperties.cxx:298-308); otherwise the coordinates are the shape's own
///   (EMU of its a:ext: customshapeproperties.cxx:236-244, a ViewBox of the shape size for plain numbers, scale 1 for guides).
///   fill (norm / none / darken / darkenLess / lighten / lightenLess) and stroke per path (customshapegeometry.cxx:997-1038).
/// - commands as EnhancedCustomShape2d::CreateSubPath builds the polygon (:2134-2496): moveTo ends the subpath (one whose
///   last point equals its first becomes closed, checkClosed); a:close right after moveTo is ignored
///   (customshapegeometry.cxx:1045-1052), otherwise it closes the subpath and the next command starts a new one at its own
///   point; arcTo with no current point is skipped. arcTo: swAng clamped to +-360 degrees (:2429-2435); stAng / swAng are
///   angles on the ellipse (lcl_getNormalizedCircleAngleRad :2051-2068 turns them into the parameter); the arc is
///   LibreOffice's Bezier ellipse segment (basegfx b2dpolygontools.cxx createPolygonFromUnitEllipseSegment :1796-1890, 30
///   degree steps, kappa 4/3 tan(a/4) :1548-1557), moved so it starts at the current point; a negative sweep runs backwards.
///   quadBezTo becomes the exact cubic.
/// </summary>
internal static class DocxShapeGeometry
{
    private const double EmuPerPoint = 12700;
    private static Dictionary<string, XElement>? _presets;
    private static readonly object PresetsGate = new();

    /// <summary>The preset's definition (avLst, gdLst, pathLst...), namespace-free; null when the name is not a preset.</summary>
    public static XElement? Preset(string? name)
    {
        if (name is null) return null;
        Dictionary<string, XElement> presets = _presets ?? LoadPresets();
        return presets.TryGetValue(name, out XElement? definition) ? definition : null;
    }

    private static Dictionary<string, XElement> LoadPresets()
    {
        lock (PresetsGate)
        {
            if (_presets is not null) return _presets;
            var presets = new Dictionary<string, XElement>(StringComparer.Ordinal);
            using Stream stream = typeof(DocxShapeGeometry).Assembly.GetManifestResourceStream("FastDocView.presetShapeDefinitions.xml")
                                  ?? throw new InvalidOperationException("presetShapeDefinitions.xml resource is missing");
            XElement root = XElement.Load(stream, LoadOptions.None);
            foreach (XElement preset in root.Elements()) presets[preset.Name.LocalName] = WithoutNamespaces(preset);
            return _presets = presets;
        }
    }

    private static XElement WithoutNamespaces(XElement element)
        => new(element.Name.LocalName, element.Attributes().Where(a => !a.IsNamespaceDeclaration).Select(a => new XAttribute(a.Name.LocalName, a.Value)),
               element.Elements().Select(WithoutNamespaces));

    /// <summary>
    /// The outline of a prefix-free a:prstGeom / a:custGeom for a shape drawn <paramref name="widthPt"/> x <paramref name="heightPt"/>
    /// whose own a:ext is <paramref name="ownWidthEmu"/> x <paramref name="ownHeightEmu"/>; null when there is nothing to draw
    /// (a name that is no preset, or no path).
    /// </summary>
    public static List<WordShapePath>? Paths(XElement geometry, double widthPt, double heightPt, double ownWidthEmu, double ownHeightEmu)
        => Paths(geometry, widthPt, heightPt, ownWidthEmu, ownHeightEmu, out _);

    /// <summary>As above, with the text rectangle (a:rect l t r b, in points from the shape's top left; null = none) as LibreOffice reads it
    /// into TextFrames (customshapegeometry.cxx a:rect, EnhancedCustomShape2d::GetTextRect :1152-1175: in the shape's own coordinates; a
    /// rectangle of 1 unit or less = the whole shape).</summary>
    public static List<WordShapePath>? Paths(XElement geometry, double widthPt, double heightPt, double ownWidthEmu, double ownHeightEmu, out (double X, double Y, double Width, double Height)? textRect)
    {
        textRect = null;
        XElement? adjustList, guideList, pathList, rectangle;
        if (geometry.Name.LocalName == "prstGeom")
        {
            XElement? definition = Preset((string?)geometry.Attribute("prst"));
            if (definition is null) return null;
            adjustList = definition.Element("avLst"); guideList = definition.Element("gdLst"); pathList = definition.Element("pathLst"); rectangle = definition.Element("rect");
        }
        else { adjustList = geometry.Element("avLst"); guideList = geometry.Element("gdLst"); pathList = geometry.Element("pathLst"); rectangle = geometry.Element("rect"); }
        if (pathList is null) return null;
        if (ownWidthEmu <= 0) ownWidthEmu = widthPt * EmuPerPoint;
        if (ownHeightEmu <= 0) ownHeightEmu = heightPt * EmuPerPoint;

        var guides = new Guides(ownWidthEmu, ownHeightEmu);
        var overrides = new Dictionary<string, string>(StringComparer.Ordinal);
        if (geometry.Name.LocalName == "prstGeom")
            foreach (XElement guide in geometry.Element("avLst")?.Elements("gd") ?? [])
                if ((string?)guide.Attribute("name") is { } name && (string?)guide.Attribute("fmla") is { } formula) overrides[name] = formula;
        foreach (XElement guide in adjustList?.Elements("gd") ?? [])
        {
            string? name = (string?)guide.Attribute("name");
            if (name is null) continue;
            guides.Define(name, overrides.TryGetValue(name, out string? replaced) ? replaced : (string?)guide.Attribute("fmla") ?? "");
            overrides.Remove(name);
        }
        foreach (KeyValuePair<string, string> extra in overrides) guides.Define(extra.Key, extra.Value);
        foreach (XElement guide in guideList?.Elements("gd") ?? [])
            if ((string?)guide.Attribute("name") is { } name) guides.Define(name, (string?)guide.Attribute("fmla") ?? "");

        if (rectangle is not null && ownWidthEmu > 0 && ownHeightEmu > 0)
        {
            double left = guides.Value((string?)rectangle.Attribute("l")), top = guides.Value((string?)rectangle.Attribute("t"));
            double right = guides.Value((string?)rectangle.Attribute("r")), bottom = guides.Value((string?)rectangle.Attribute("b"));
            double sx = widthPt / ownWidthEmu, sy = heightPt / ownHeightEmu;
            if (right - left > 360 && bottom - top > 360)   // 1/100 mm = 360 EMU
                textRect = (left * sx, top * sy, (right - left) * sx, (bottom - top) * sy);
        }
        var paths = new List<WordShapePath>();
        foreach (XElement path in pathList.Elements("path"))
        {
            double pathWidth = Attribute(path, "w"), pathHeight = Attribute(path, "h");
            bool ownBox = pathWidth > 0 && pathHeight > 0;
            double scaleX = widthPt / (ownBox ? pathWidth : ownWidthEmu), scaleY = heightPt / (ownBox ? pathHeight : ownHeightEmu);
            if (double.IsNaN(scaleX) || double.IsInfinity(scaleX)) scaleX = 0;
            if (double.IsNaN(scaleY) || double.IsInfinity(scaleY)) scaleY = 0;
            var result = new WordShapePath
            {
                Fill = (string?)path.Attribute("fill") switch
                {
                    "none" => WordShapePathFill.None, "darken" => WordShapePathFill.Darken, "darkenLess" => WordShapePathFill.DarkenLess,
                    "lighten" => WordShapePathFill.Lighten, "lightenLess" => WordShapePathFill.LightenLess, _ => WordShapePathFill.Normal,
                },
                Stroke = (string?)path.Attribute("stroke") is not ("0" or "false"),
            };
            var builder = new PolygonBuilder(result.Subpaths);
            bool previousWasMove = false;
            foreach (XElement command in path.Elements())
            {
                (double x, double y) Point(XElement? point) => (guides.Value((string?)point?.Attribute("x")), guides.Value((string?)point?.Attribute("y")));
                switch (command.Name.LocalName)
                {
                    case "moveTo": builder.Move(Point(command.Element("pt"))); previousWasMove = true; continue;
                    case "lnTo": builder.Line(Point(command.Element("pt"))); break;
                    case "cubicBezTo":
                    {
                        List<XElement> points = command.Elements("pt").ToList();
                        if (points.Count == 3) builder.Cubic(Point(points[0]), Point(points[1]), Point(points[2]));
                        break;
                    }
                    case "quadBezTo":
                    {
                        List<XElement> points = command.Elements("pt").ToList();
                        if (points.Count == 2) builder.Quadratic(Point(points[0]), Point(points[1]));
                        break;
                    }
                    case "arcTo":
                        builder.Arc(guides.Value((string?)command.Attribute("wR")), guides.Value((string?)command.Attribute("hR")),
                                    guides.Value((string?)command.Attribute("stAng")) / 60000, guides.Value((string?)command.Attribute("swAng")) / 60000);
                        break;
                    case "close":
                        if (!previousWasMove) builder.Close();
                        break;
                    default: continue;
                }
                previousWasMove = false;
            }
            builder.Finish();
            foreach (WordShapeSubpath subpath in result.Subpaths) Scale(subpath, scaleX, scaleY);
            paths.Add(result);
        }
        return paths.Count == 0 ? null : paths;
    }

    private static void Scale(WordShapeSubpath subpath, double sx, double sy)
    {
        subpath.StartX *= sx; subpath.StartY *= sy;
        for (int i = 0; i < subpath.Segments.Count; i++)
        {
            WordShapeSegment s = subpath.Segments[i];
            subpath.Segments[i] = new WordShapeSegment(s.IsCurve, s.Control1X * sx, s.Control1Y * sy, s.Control2X * sx, s.Control2Y * sy, s.X * sx, s.Y * sy);
        }
    }

    private static double Attribute(XElement element, string name)
        => double.TryParse((string?)element.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : 0;

    /// <summary>The guide values of one shape (adjust values and guides share one name space; the built-ins are looked up first).</summary>
    private sealed class Guides(double width, double height)
    {
        private readonly Dictionary<string, double> _values = new(StringComparer.Ordinal);

        public void Define(string name, string formula) => _values[name] = Evaluate(formula);

        public double Value(string? token)
        {
            if (string.IsNullOrEmpty(token)) return 0;
            double ss = Math.Min(width, height);
            switch (token)
            {
                case "l" or "t": return 0;
                case "r" or "w": return width;
                case "b" or "h": return height;
                case "hc" or "wd2": return width / 2;
                case "vc" or "hd2": return height / 2;
                case "wd3": return width / 3; case "wd4": return width / 4; case "wd5": return width / 5; case "wd6": return width / 6;
                case "wd8": return width / 8; case "wd10": return width / 10; case "wd12": return width / 12; case "wd32": return width / 32;
                case "hd3": return height / 3; case "hd4": return height / 4; case "hd5": return height / 5; case "hd6": return height / 6;
                case "hd8": return height / 8; case "hd10": return height / 10;
                case "ls": return Math.Max(width, height);
                case "ss": return ss;
                case "ssd2": return ss / 2; case "ssd4": return ss / 4; case "ssd6": return ss / 6; case "ssd8": return ss / 8;
                case "ssd16": return ss / 16; case "ssd32": return ss / 32;
                case "cd2": return 180 * 60000; case "cd3": return 120 * 60000; case "cd4": return 90 * 60000; case "cd8": return 45 * 60000;
                case "3cd4": return 270 * 60000; case "3cd8": return 135 * 60000; case "5cd8": return 225 * 60000; case "7cd8": return 315 * 60000;
            }
            if (double.TryParse(token, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowExponent, CultureInfo.InvariantCulture, out double number)) return number;
            return _values.TryGetValue(token!, out double value) ? value : 0;
        }

        private double Evaluate(string formula)
        {
            string[] parts = formula.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return 0;
            double x = parts.Length > 1 ? Value(parts[1]) : 0, y = parts.Length > 2 ? Value(parts[2]) : 0, z = parts.Length > 3 ? Value(parts[3]) : 0;
            const double PerRadian = 10800000 / Math.PI;
            double result = parts[0] switch
            {
                "*/" => z == 0 ? 0 : x * y / z,
                "+-" => x + y - z,
                "+/" => z == 0 ? 0 : (x + y) / z,
                "?:" or "ifelse" => x > 0 ? y : z,
                "abs" => Math.Abs(x),
                "at2" => Math.Atan2(y, x) * PerRadian,
                "cat2" => x * Math.Cos(Math.Atan2(z, y)),
                "sat2" => x * Math.Sin(Math.Atan2(z, y)),
                "cos" => x * Math.Cos(y / PerRadian),
                "sin" => x * Math.Sin(y / PerRadian),
                "tan" => x * Math.Tan(y / PerRadian),
                "max" => Math.Max(x, y),
                "min" => Math.Min(x, y),
                "mod" => Math.Sqrt(x * x + y * y + z * z),
                "pin" => x > y ? x : z > y ? y : z,
                "sqrt" => x > 0 ? Math.Sqrt(x) : 0,
                "val" => x,
                _ => 0,
            };
            return double.IsNaN(result) || double.IsInfinity(result) ? 0 : result;
        }
    }

    /// <summary>Builds subpaths with LibreOffice's polygon rules (see the class summary), in path coordinates.</summary>
    private sealed class PolygonBuilder(List<WordShapeSubpath> output)
    {
        private WordShapeSubpath? _current;
        private double _x, _y;

        private int PointCount => _current is null ? 0 : 1 + _current.Segments.Count;

        public void Move((double x, double y) point)
        {
            Finish();
            Begin(point);
        }

        private void Begin((double x, double y) point)
        {
            _current = new WordShapeSubpath { StartX = point.x, StartY = point.y };
            (_x, _y) = point;
        }

        public void Line((double x, double y) point)
        {
            if (_current is null) { Begin(point); return; }
            _current.Segments.Add(WordShapeSegment.Line(point.x, point.y));
            (_x, _y) = point;
        }

        public void Cubic((double x, double y) c1, (double x, double y) c2, (double x, double y) end)
        {
            if (_current is null) { Begin(end); return; }
            _current.Segments.Add(new WordShapeSegment(true, c1.x, c1.y, c2.x, c2.y, end.x, end.y));
            (_x, _y) = end;
        }

        public void Quadratic((double x, double y) control, (double x, double y) end)
            => Cubic((_x + 2.0 / 3 * (control.x - _x), _y + 2.0 / 3 * (control.y - _y)),
                     (end.x + 2.0 / 3 * (control.x - end.x), end.y + 2.0 / 3 * (control.y - end.y)), end);

        public void Close()
        {
            if (_current is null) return;
            if (PointCount > 1) { _current.Closed = true; output.Add(_current); }
            _current = null;
        }

        public void Finish()
        {
            if (_current is not null && PointCount > 1)
            {
                if (Same(_current.StartX, _x) && Same(_current.StartY, _y)) _current.Closed = true;   // checkClosed
                output.Add(_current);
            }
            _current = null;
        }

        private static bool Same(double a, double b) => Math.Abs(a - b) <= 1e-9 * Math.Max(1, Math.Max(Math.Abs(a), Math.Abs(b)));

        /// <summary>EnhancedCustomShape2d.cxx ARCANGLETO (:2414-2496).</summary>
        public void Arc(double radiusX, double radiusY, double startDegrees, double sweepDegrees)
        {
            if (_current is null) return;
            sweepDegrees = Math.Max(-360, Math.Min(360, sweepDegrees));
            if (radiusX == 0 && radiusY == 0) return;   // a single point at the current point
            double endDegrees = startDegrees + sweepDegrees, from = startDegrees, to = endDegrees;
            if (sweepDegrees < 0) (from, to) = (to, from);
            var points = new List<WordShapeSegment>();
            double startX = 0, startY = 0;
            bool started = false;
            void Append(double fromRadians, double toRadians)
            {
                List<WordShapeSegment> piece = UnitEllipseSegment(fromRadians, toRadians, out double pieceX, out double pieceY);
                if (!started) { startX = radiusX + radiusX * pieceX; startY = radiusY + radiusY * pieceY; started = true; }
                foreach (WordShapeSegment s in piece)
                    points.Add(new WordShapeSegment(s.IsCurve, radiusX + radiusX * s.Control1X, radiusY + radiusY * s.Control1Y, radiusX + radiusX * s.Control2X,
                                                    radiusY + radiusY * s.Control2Y, radiusX + radiusX * s.X, radiusY + radiusY * s.Y));
            }
            if (sweepDegrees >= 360 || sweepDegrees <= -360)
            {
                double next = from + 180;
                while (next < to) { Append(CircleAngle(radiusX, radiusY, from), CircleAngle(radiusX, radiusY, next)); from = next; next += 180; }
            }
            Append(CircleAngle(radiusX, radiusY, from), CircleAngle(radiusX, radiusY, to));
            if (sweepDegrees < 0) Reverse(ref startX, ref startY, points);
            double dx = _x - startX, dy = _y - startY;
            foreach (WordShapeSegment s in points)
            {
                double ex = s.X + dx, ey = s.Y + dy;
                if (!s.IsCurve && Same(ex, _x) && Same(ey, _y)) continue;
                _current.Segments.Add(new WordShapeSegment(s.IsCurve, s.Control1X + dx, s.Control1Y + dy, s.Control2X + dx, s.Control2Y + dy, ex, ey));
                (_x, _y) = (ex, ey);
            }
        }

        private static void Reverse(ref double startX, ref double startY, List<WordShapeSegment> segments)
        {
            if (segments.Count == 0) return;
            var reversed = new List<WordShapeSegment>(segments.Count);
            double endX = segments[segments.Count - 1].X, endY = segments[segments.Count - 1].Y;
            for (int i = segments.Count - 1; i >= 0; i--)
            {
                double targetX = i == 0 ? startX : segments[i - 1].X, targetY = i == 0 ? startY : segments[i - 1].Y;
                WordShapeSegment s = segments[i];
                reversed.Add(new WordShapeSegment(s.IsCurve, s.Control2X, s.Control2Y, s.Control1X, s.Control1Y, targetX, targetY));
            }
            segments.Clear();
            segments.AddRange(reversed);
            (startX, startY) = (endX, endY);
        }

        /// <summary>lcl_getNormalizedCircleAngleRad: an angle on the ellipse (degrees) as the ellipse parameter in [0, 2pi).</summary>
        private static double CircleAngle(double radiusX, double radiusY, double degrees)
        {
            double angle = degrees % 360.0;
            if (angle < 0) angle += 360;
            if (angle is 0 or 90 or 180 or 270) return angle * Math.PI / 180;
            double x = radiusY * Math.Cos(angle * Math.PI / 180), y = radiusX * Math.Sin(angle * Math.PI / 180);
            if (x == 0 && y == 0) return 0;
            double result = Math.Atan2(y, x);
            return result < 0 ? result + 2 * Math.PI : result;
        }

        /// <summary>basegfx createPolygonFromUnitEllipseSegment: the unit-circle arc from <paramref name="start"/> to <paramref name="end"/> (radians, increasing, wrapping).</summary>
        private static List<WordShapeSegment> UnitEllipseSegment(double start, double end, out double startX, out double startY)
        {
            if (start < 0 || start >= 2 * Math.PI - 1e-12) start = 0;
            if (end < 0 || end >= 2 * Math.PI - 1e-12) end = 0;
            var segments = new List<WordShapeSegment>();
            startX = Math.Cos(start); startY = Math.Sin(start);
            if (Math.Abs(start - end) <= 1e-12) return segments;   // a single point
            const int Steps = 3 * 4;
            const double PerSegment = Math.PI / 2 / 3;
            int startSegment = (int)(start / PerSegment) % Steps, endSegment = (int)(end / PerSegment) % Steps;
            double Kappa(double angle) => angle is >= 0 and <= Math.PI / 2 ? 4.0 / 3.0 * Math.Tan(angle / 4.0) : 0;
            void Add(double a0, double a1, double factor)
            {
                if (Math.Abs(a1 - a0) <= 1e-12) return;   // removeDoublePoints drops an empty piece
                double x0 = Math.Cos(a0), y0 = Math.Sin(a0), x1 = Math.Cos(a1), y1 = Math.Sin(a1);
                segments.Add(new WordShapeSegment(true, x0 - y0 * factor, y0 + x0 * factor, x1 + y1 * factor, y1 - x1 * factor, x1, y1));
            }
            if (startSegment == endSegment && end > start) { Add(start, end, Kappa(end - start)); return segments; }
            double segmentEnd = (startSegment + 1) * PerSegment;
            Add(start, segmentEnd, Kappa(segmentEnd - start));
            int segment = (startSegment + 1) % Steps;
            while (segment != endSegment)
            {
                Add(segment * PerSegment, (segment + 1) * PerSegment, Kappa(PerSegment));
                segment = (segment + 1) % Steps;
            }
            double segmentStart = segment * PerSegment;
            Add(segmentStart, end, Kappa(end - segmentStart));
            return segments;
        }
    }
}
