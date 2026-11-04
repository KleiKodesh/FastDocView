using System.Globalization;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// ODF path data (svg:d, ODF 1.3 part 3 19.528, the SVG path syntax) and point lists (draw:points 19.208), read as
/// LibreOffice reads them (basegfx/source/polygon/b2dsvgpolypolygon.cxx importFromSvgD :76-660, master read 2026-11-04):
/// M / m L / l H / h V / v C / c S / s Q / q T / t A / a Z / z, implicit repetition, a sub-path started without a moveto
/// begins at the current point (:96-102), after Z the current point is the sub-path's start (SVG; :117-124 - the old
/// LibreOffice compatibility mode is for files of OpenOffice-era writers only and is not followed), arcs per SVG appendix
/// F.6 (start = end skipped, a zero radius is a line, radii scaled up when too small). Quadratic curves become cubics
/// (exact); arcs are flattened to short lines. Coordinates are returned as given (the caller maps the view box).
/// </summary>
internal static class OdtSvgPath
{
    public static List<WordShapeSubpath>? Parse(string? data)
    {
        if (string.IsNullOrWhiteSpace(data)) return null;
        string d = data!;
        var result = new List<WordShapeSubpath>();
        WordShapeSubpath? current = null;
        double lastX = 0, lastY = 0, controlX = 0, controlY = 0;
        char previous = ' ';
        int i = 0;
        void Start(double x, double y) { current = new WordShapeSubpath { StartX = x, StartY = y }; }
        void Line(double x, double y)
        {
            if (current is null) Start(lastX, lastY);
            current!.Segments.Add(WordShapeSegment.Line(x, y));
        }
        void Curve(double c1x, double c1y, double c2x, double c2y, double x, double y)
        {
            if (current is null) Start(lastX, lastY);
            current!.Segments.Add(new WordShapeSegment(true, c1x, c1y, c2x, c2y, x, y));
        }
        void Flush() { if (current is not null && (current.Segments.Count > 0 || current.Closed)) result.Add(current); else if (current is not null) result.Add(current); current = null; }
        SkipSpaces(d, ref i);
        while (i < d.Length)
        {
            char c = d[i];
            if (!char.IsLetter(c)) return result.Count > 0 || current is not null ? Finish() : null;   // a number without a command: stop (LibreOffice fails the import)
            i++;
            SkipSpaces(d, ref i);
            bool relative = char.IsLower(c);
            switch (char.ToUpperInvariant(c))
            {
                case 'Z':
                    if (current is not null)
                    {
                        current.Closed = true;
                        lastX = current.StartX; lastY = current.StartY;
                        result.Add(current);
                        current = null;
                    }
                    break;
                case 'M':
                {
                    bool first = true;
                    while (Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x += lastX; y += lastY; }
                        if (first) { if (current is not null) Flush(); Start(x, y); first = false; }
                        else Line(x, y);
                        lastX = x; lastY = y;
                    }
                    break;
                }
                case 'L':
                    while (Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x += lastX; y += lastY; }
                        Line(x, y); lastX = x; lastY = y;
                    }
                    break;
                case 'H':
                    while (Number(d, ref i, out double x)) { if (relative) x += lastX; Line(x, lastY); lastX = x; }
                    break;
                case 'V':
                    while (Number(d, ref i, out double y)) { if (relative) y += lastY; Line(lastX, y); lastY = y; }
                    break;
                case 'C':
                    while (Number(d, ref i, out double x1) && Number(d, ref i, out double y1) && Number(d, ref i, out double x2) && Number(d, ref i, out double y2)
                           && Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x1 += lastX; y1 += lastY; x2 += lastX; y2 += lastY; x += lastX; y += lastY; }
                        Curve(x1, y1, x2, y2, x, y);
                        controlX = x2; controlY = y2; lastX = x; lastY = y;
                    }
                    break;
                case 'S':
                    while (Number(d, ref i, out double x2) && Number(d, ref i, out double y2) && Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x2 += lastX; y2 += lastY; x += lastX; y += lastY; }
                        // the first control point reflects the previous curve's second one (SVG 1.1 8.3.6)
                        double x1 = lastX, y1 = lastY;
                        if (previous is 'C' or 'c' or 'S' or 's') { x1 = 2 * lastX - controlX; y1 = 2 * lastY - controlY; }
                        Curve(x1, y1, x2, y2, x, y);
                        controlX = x2; controlY = y2; lastX = x; lastY = y;
                        previous = 'S';
                    }
                    break;
                case 'Q':
                    while (Number(d, ref i, out double qx) && Number(d, ref i, out double qy) && Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { qx += lastX; qy += lastY; x += lastX; y += lastY; }
                        Curve(lastX + 2.0 / 3 * (qx - lastX), lastY + 2.0 / 3 * (qy - lastY), x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y);
                        controlX = qx; controlY = qy; lastX = x; lastY = y;
                    }
                    break;
                case 'T':
                    while (Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x += lastX; y += lastY; }
                        double qx = lastX, qy = lastY;
                        if (previous is 'Q' or 'q' or 'T' or 't') { qx = 2 * lastX - controlX; qy = 2 * lastY - controlY; }
                        Curve(lastX + 2.0 / 3 * (qx - lastX), lastY + 2.0 / 3 * (qy - lastY), x + 2.0 / 3 * (qx - x), y + 2.0 / 3 * (qy - y), x, y);
                        controlX = qx; controlY = qy; lastX = x; lastY = y;
                        previous = 'T';
                    }
                    break;
                case 'A':
                    while (Number(d, ref i, out double rx) && Number(d, ref i, out double ry) && Number(d, ref i, out double phi)
                           && Flag(d, ref i, out bool large) && Flag(d, ref i, out bool sweep) && Number(d, ref i, out double x) && Number(d, ref i, out double y))
                    {
                        if (relative) { x += lastX; y += lastY; }
                        if (Math.Abs(x - lastX) < 1e-12 && Math.Abs(y - lastY) < 1e-12) continue;
                        if (rx == 0 || ry == 0) Line(x, y);
                        else foreach ((double px, double py) in Arc(lastX, lastY, rx, ry, phi, large, sweep, x, y)) Line(px, py);
                        lastX = x; lastY = y;
                    }
                    break;
                default:
                    return Finish();
            }
            if (char.ToUpperInvariant(c) is not ('S' or 'T')) previous = c;
            SkipSpaces(d, ref i);
        }
        return Finish();

        List<WordShapeSubpath> Finish()
        {
            if (current is not null) result.Add(current);
            current = null;
            return result.Where(s => s.Segments.Count > 0).ToList();
        }
    }

    /// <summary>draw:points: "x,y x,y ..." (pairs separated by white space, x and y by a comma; basegfx importFromSvgPoints).</summary>
    public static List<(double X, double Y)>? Points(string? data)
    {
        if (string.IsNullOrWhiteSpace(data)) return null;
        var points = new List<(double X, double Y)>();
        int i = 0;
        SkipSpaces(data!, ref i);
        while (i < data!.Length && Number(data, ref i, out double x) && Number(data, ref i, out double y)) points.Add((x, y));
        return points.Count > 0 ? points : null;
    }

    /// <summary>svg:viewBox: x y width height.</summary>
    public static (double X, double Y, double Width, double Height)? ViewBox(string? data)
    {
        if (data is null) return null;
        int i = 0;
        SkipSpaces(data, ref i);
        return Number(data, ref i, out double x) && Number(data, ref i, out double y) && Number(data, ref i, out double w) && Number(data, ref i, out double h) ? (x, y, w, h) : null;
    }

    private static void SkipSpaces(string s, ref int i) { while (i < s.Length && (char.IsWhiteSpace(s[i]) || s[i] == ',')) i++; }

    private static bool Flag(string s, ref int i, out bool value)
    {
        SkipSpaces(s, ref i);
        value = false;
        if (i >= s.Length || s[i] is not ('0' or '1')) return false;
        value = s[i++] == '1';
        SkipSpaces(s, ref i);
        return true;
    }

    internal static bool Number(string s, ref int i, out double value)
    {
        SkipSpaces(s, ref i);
        value = 0;
        int start = i;
        if (i < s.Length && s[i] is '+' or '-') i++;
        bool digits = false;
        while (i < s.Length && char.IsDigit(s[i])) { i++; digits = true; }
        if (i < s.Length && s[i] == '.') { i++; while (i < s.Length && char.IsDigit(s[i])) { i++; digits = true; } }
        if (!digits) { i = start; return false; }
        if (i < s.Length && s[i] is 'e' or 'E')
        {
            int e = i++;
            if (i < s.Length && s[i] is '+' or '-') i++;
            if (i < s.Length && char.IsDigit(s[i])) while (i < s.Length && char.IsDigit(s[i])) i++;
            else i = e;
        }
        value = double.Parse(s.Substring(start, i - start), NumberStyles.Float, CultureInfo.InvariantCulture);
        SkipSpaces(s, ref i);
        return true;
    }

    /// <summary>An SVG elliptical arc (SVG 1.1 appendix F.6.5 / F.6.6) as points after the start, ending exactly at the end point.</summary>
    private static List<(double X, double Y)> Arc(double x1, double y1, double rx, double ry, double phiDegrees, bool large, bool sweep, double x2, double y2)
    {
        rx = Math.Abs(rx); ry = Math.Abs(ry);
        double phi = phiDegrees * Math.PI / 180, cos = Math.Cos(phi), sin = Math.Sin(phi);
        double dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
        double x1p = cos * dx + sin * dy, y1p = -sin * dx + cos * dy;
        double radicant = (rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p) / (rx * rx * y1p * y1p + ry * ry * x1p * x1p);
        double cxp = 0, cyp = 0;
        if (radicant < 0)
        {
            double ratio = rx / ry;
            ry = Math.Sqrt(y1p * y1p + x1p * x1p / (ratio * ratio));
            rx = ratio * ry;
        }
        else
        {
            double factor = (large == sweep ? -1 : 1) * Math.Sqrt(radicant);
            cxp = factor * rx * y1p / ry; cyp = -factor * ry * x1p / rx;
        }
        double cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
        double Angle(double ux, double uy, double vx, double vy) { double a = Math.Atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; }
        double theta1 = Angle(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
        double delta = Angle((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
        if (!sweep && delta > 0) delta -= 2 * Math.PI;
        else if (sweep && delta < 0) delta += 2 * Math.PI;
        int steps = Math.Max(4, (int)Math.Ceiling(Math.Abs(delta) / (Math.PI / 32)));
        var points = new List<(double X, double Y)>(steps);
        for (int k = 1; k <= steps; k++)
        {
            if (k == steps) { points.Add((x2, y2)); break; }
            double t = theta1 + delta * k / steps;
            double ex = rx * Math.Cos(t), ey = ry * Math.Sin(t);
            points.Add((cos * ex - sin * ey + cx, sin * ex + cos * ey + cy));
        }
        return points;
    }
}
