using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// 3-D bar / column and pie charts drawn as projected, flat-shaded SVG polygons, the way LibreOffice imports and draws them:
/// c:view3D (21.2.2.228) rotX / rotY / perspective / rAngAx as oox's View3DConverter maps them
/// (oox/source/drawingml/chart/plotareaconverter.cxx:262-323: bars rotY default 20, rotX default 15 limited to -90..90, perspective
/// halved to 0..100, parallel projection when right-angled or perspective 0; pies: rotX 0..90 becomes the elevation rotX - 90, rotY the
/// first slice angle; flat shading, ambient grey 20 % (#CCCCCC) and one light grey 60 % (#666666) from (0.2, 0.4, 1) for bars, ambient
/// #B3B3B3 and light #4C4C4C for pies), the camera distance from the perspective (chart2/source/tools/ThreeDHelper.cxx:605-645:
/// distance range 0.75..20 scene volumes, perspective = a / distance + b). The scene is turned by rotY about the vertical axis, then by
/// rotX about the horizontal one, and projected (parallel, or a pinhole at the camera distance); right-angled axes (Word's default for
/// columns) keep the front faces flat and draw the depth obliquely (the axes stay at right angles, ThreeDHelper.cxx:598-602 clips the
/// angles for it). Faces turned away are not drawn and each face takes one colour (ambient + light x max(0, n . l)), as the extruded
/// shapes in Odt/OdtExtrusion.cs do (drawinglayer defaultprocessor3d.cxx:348-431). The whole scene is fitted into the plot rectangle;
/// boxes are painted far to near. 3-D line and area charts keep the flat wall-and-floor drawing.
/// </summary>
internal sealed partial class DocxChart
{
    /// <summary>c:view3D values (with OOXML's defaults), and the scene lighting (ODF dr3d:ambient-color / light when given).</summary>
    private sealed class View3D
    {
        public double RotX = 15, RotY = 20, Perspective = 30, DepthPercent = 100, GapDepth = 150;
        public bool RightAngled;
        public double Ambient = 0.8, Light = 0.4;
        public (double X, double Y, double Z) LightDirection = (0.2, 0.4, 1.0);
    }

    private View3D? _view3D;

    private static View3D ReadView3D(XElement? view, bool pie)
    {
        var result = new View3D();
        if (pie) { result.RotX = 30; result.RotY = 0; result.Ambient = 0xB3 / 255.0; result.Light = 0x4C / 255.0; }
        if (view is null) return result;
        if (Number(view.Element("rotX")) is { } rotX) result.RotX = rotX;
        if (Number(view.Element("rotY")) is { } rotY) result.RotY = rotY;
        if (Number(view.Element("perspective")) is { } perspective) result.Perspective = perspective;
        if (Number(view.Element("depthPercent")) is { } depth) result.DepthPercent = Math.Max(20, Math.Min(2000, depth));
        result.RightAngled = Value(view.Element("rAngAx")) is "1" or "true";
        if (double.TryParse((string?)view.Attribute("ambient"), NumberStyles.Float, CultureInfo.InvariantCulture, out double ambient)) result.Ambient = ambient;
        if (double.TryParse((string?)view.Attribute("light"), NumberStyles.Float, CultureInfo.InvariantCulture, out double light)) result.Light = light;
        return result;
    }

    /// <summary>A projection of scene points (x right, y up, z away from the viewer) to the page, fitted into a rectangle.</summary>
    private sealed class Projection
    {
        private readonly View3D _view;
        private readonly double _cx, _cy, _cz, _distance, _sinX, _cosX, _sinY, _cosY;
        private double _scale = 1, _offsetX, _offsetY;

        public Projection(View3D view, double width, double height, double depth)
        {
            _view = view;
            (_cx, _cy, _cz) = (width / 2, height / 2, depth / 2);
            double x = view.RotX * Math.PI / 180, y = view.RotY * Math.PI / 180;
            (_sinX, _cosX, _sinY, _cosY) = (Math.Sin(x), Math.Cos(x), Math.Sin(y), Math.Cos(y));
            // ThreeDHelper.cxx:635-645 PerspectiveToCameraDistance, in scene volumes (the scene's largest extent)
            double perspective = Math.Max(0, Math.Min(100, view.Perspective / 2));
            const double minimum = 0.75, maximum = 20;
            double a = 100 * maximum * minimum / (maximum - minimum), b = -a / maximum;
            _distance = view.RightAngled || perspective <= 0 ? double.PositiveInfinity : a / (perspective - b) * Math.Max(width, Math.Max(height, depth));
            // never inside the scene (OOXML perspective goes to 240, past LibreOffice's 100): at least the scene's diagonal away
            _distance = Math.Max(_distance, 1.5 * Math.Sqrt(width * width + height * height + depth * depth) / 2);
        }

        /// <summary>A direction turned by the scene rotation (z toward the viewer): visibility and shading of a face.</summary>
        public (double X, double Y, double Z) Turn((double X, double Y, double Z) v)
        {
            // scene z runs away from the viewer; turned space has z toward the viewer
            double x = v.X, y = v.Y, z = -v.Z;
            double x1 = x * _cosY - z * _sinY, z1 = x * _sinY + z * _cosY;
            double y2 = y * _cosX - z1 * _sinX, z2 = y * _sinX + z1 * _cosX;
            return (x1, y2, z2);
        }

        /// <summary>The unfitted projection: oblique for right-angled axes, else turned and seen from the camera.</summary>
        private (double X, double Y, double Depth) Raw(double x, double y, double z)
        {
            if (_view.RightAngled)
            {
                // the front stays flat; depth goes up and to the side by the angles (Word's right-angled axes)
                return (x + z * _sinY * Math.Abs(_cosX), y + z * _sinX, z);
            }
            (double tx, double ty, double tz) = Turn((x - _cx, y - _cy, z - _cz));
            double factor = double.IsInfinity(_distance) ? 1 : _distance / Math.Max(1e-6, _distance - tz);
            return (tx * factor, ty * factor, -tz);
        }

        /// <summary>Fits the projected scene box into the rectangle (uniform scale, centred).</summary>
        public void Fit(double width, double height, double depth, double left, double top, double right, double bottom)
        {
            double minX = double.MaxValue, maxX = double.MinValue, minY = double.MaxValue, maxY = double.MinValue;
            foreach (double x in new[] { 0, width })
                foreach (double y in new[] { 0, height })
                    foreach (double z in new[] { 0, depth })
                    {
                        var p = Raw(x, y, z);
                        minX = Math.Min(minX, p.X); maxX = Math.Max(maxX, p.X); minY = Math.Min(minY, p.Y); maxY = Math.Max(maxY, p.Y);
                    }
            _scale = Math.Min((right - left) / Math.Max(1e-9, maxX - minX), (bottom - top) / Math.Max(1e-9, maxY - minY));
            _offsetX = left + ((right - left) - (maxX - minX) * _scale) / 2 - minX * _scale;
            _offsetY = bottom - ((bottom - top) - (maxY - minY) * _scale) / 2 + minY * _scale;
        }

        public (double X, double Y) Point(double x, double y, double z)
        {
            var p = Raw(x, y, z);
            return (_offsetX + p.X * _scale, _offsetY - p.Y * _scale);
        }

        public double DepthOf(double x, double y, double z) => Raw(x, y, z).Depth;
    }

    private static string Shade(string css, double factor)
    {
        if (css.Length != 7 || css[0] != '#') return css;
        int Channel(int at) => (int)Math.Round(Math.Max(0, Math.Min(255, int.Parse(css.Substring(at, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture) * factor)));
        return "#" + Channel(1).ToString("X2", CultureInfo.InvariantCulture) + Channel(3).ToString("X2", CultureInfo.InvariantCulture) + Channel(5).ToString("X2", CultureInfo.InvariantCulture);
    }

    /// <summary>Flat shading of a face: ambient + light x max(0, n . l) with the face's turned normal (OdtExtrusion's model, one light).</summary>
    private static double Lighting(View3D view, (double X, double Y, double Z) turnedNormal)
    {
        var l = view.LightDirection;
        double length = Math.Sqrt(l.X * l.X + l.Y * l.Y + l.Z * l.Z);
        double dot = (turnedNormal.X * l.X + turnedNormal.Y * l.Y + turnedNormal.Z * l.Z) / length;
        return view.Ambient + view.Light * Math.Max(0, dot);
    }

    private static void Polygon(StringBuilder svg, IEnumerable<(double X, double Y)> points, string fill, string stroke, string? className = null)
    {
        svg.Append("<path").Append(className is null ? "" : " class=\"" + className + "\"").Append(" d=\"");
        bool first = true;
        foreach (var p in points) { svg.Append(first ? 'M' : 'L').Append(N(p.X)).Append(' ').Append(N(p.Y)); first = false; }
        svg.Append("Z\" fill=\"").Append(fill).Append("\" stroke=\"").Append(stroke).Append("\" stroke-width=\"0.5\" stroke-linejoin=\"round\"/>");
    }

    /// <summary>3-D columns: walls, floor, gridlines, value and category labels and every bar as a shaded box, far to near.</summary>
    private void Draw3DColumns(StringBuilder svg, int categoryCount, double left, double top, double right, double bottom, Func<double, double> valueToAxis,
                               List<(double Value, string Text)> valueLabels, List<string> categories, bool showCategories, string textColor, string gridColor, string? plotFill)
    {
        View3D view = _view3D ?? new View3D();
        Group group = _groups[0];
        bool standard = group.Grouping == "standard";
        int rows = standard ? group.Series.Count : 1;
        double width = 1, height = Math.Max(0.2, (bottom - top) / Math.Max(1, right - left));
        double slot = width / categoryCount, rowDepth = slot * view.DepthPercent / 100, depth = rowDepth * rows;
        var projection = new Projection(view, width, height, depth);
        projection.Fit(width, height, depth, left, top, right, bottom);
        string wallFill = PropertiesFill(_backWallProperties, plotFill) ?? "none", floorFill = PropertiesFill(_floorProperties, plotFill) ?? "none";
        // walls (back, left) and floor, then gridlines on the walls
        Polygon(svg, new[] { projection.Point(0, 0, depth), projection.Point(width, 0, depth), projection.Point(width, height, depth), projection.Point(0, height, depth) }, wallFill, gridColor, "word-chart-wall");
        Polygon(svg, new[] { projection.Point(0, 0, 0), projection.Point(0, 0, depth), projection.Point(0, height, depth), projection.Point(0, height, 0) }, wallFill, gridColor, "word-chart-wall");
        Polygon(svg, new[] { projection.Point(0, 0, 0), projection.Point(width, 0, 0), projection.Point(width, 0, depth), projection.Point(0, 0, depth) }, floorFill, gridColor, "word-chart-floor");
        foreach ((double value, string text) in valueLabels)
        {
            double y = valueToAxis(value) * height;
            var a = projection.Point(0, y, 0); var b = projection.Point(0, y, depth); var c = projection.Point(width, y, depth);
            svg.Append("<path d=\"M").Append(N(a.X)).Append(' ').Append(N(a.Y)).Append('L').Append(N(b.X)).Append(' ').Append(N(b.Y)).Append('L').Append(N(c.X)).Append(' ').Append(N(c.Y))
               .Append("\" fill=\"none\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
            Text(svg, a.X - 4, a.Y + 3, text, 9, textColor, "end");
        }
        if (showCategories)
            for (int i = 0; i < categoryCount; i++)
            {
                var p = projection.Point((i + 0.5) * slot, 0, 0);
                Text(svg, p.X, p.Y + 11, i < categories.Count ? categories[i] : (i + 1).ToString(CultureInfo.InvariantCulture), 9, textColor, "middle");
            }
        // the boxes
        bool stacked = group.Grouping is "stacked" or "percentStacked", percent = group.Grouping == "percentStacked";
        int count = group.Series.Count, totalSeries = _groups.Sum(g => g.Series.Count);
        double gap = group.GapWidth / 100, overlap = group.Overlap / 100;
        double bars = stacked || standard ? 1 : count - (count - 1) * overlap;
        double barWidth = slot / (bars + gap), barDepth = rowDepth / (1 + view.GapDepth / 100);
        var boxes = new List<(double X0, double X1, double Y0, double Y1, double Z0, double Z1, string Color, string? Label)>();
        for (int i = 0; i < categoryCount; i++)
        {
            double positive = 0, negative = 0, total = group.Series.Sum(s => Math.Abs(i < s.Values.Count ? s.Values[i] ?? 0 : 0));
            for (int s = 0; s < count; s++)
            {
                Series series = group.Series[s];
                if (i >= series.Values.Count || series.Values[i] is not { } raw) continue;
                double value = percent ? (total == 0 ? 0 : raw / total) : raw;
                double from, to;
                if (stacked) { if (value >= 0) { from = positive; positive += value; to = positive; } else { from = negative; negative += value; to = negative; } }
                else { from = 0; to = value; }
                double x0 = i * slot + gap * barWidth / 2 + (stacked || standard ? 0 : s * barWidth * (1 - overlap));
                double z0 = (standard ? s : 0) * rowDepth + (rowDepth - barDepth) / 2;
                double y0 = Math.Max(0, Math.Min(1, valueToAxis(Math.Min(from, to)))) * height, y1 = Math.Max(0, Math.Min(1, valueToAxis(Math.Max(from, to)))) * height;
                DocxResolvedColor fill = FillOf(group, series, group.VaryColors && count == 1 ? i : -1, group.VaryColors && count == 1 ? categoryCount : totalSeries);
                boxes.Add((x0, x0 + barWidth, y0, y1, z0, z0 + barDepth, fill.Css, LabelText(group, series, i)));
            }
        }
        // far to near: by the projected depth of each box's centre
        // ties (right-angled axes: one row shares its depth) - lower boxes first when seen from above, and the boxes away from the
        // side the viewer looks from first, so a box's top and side faces never cover its neighbour's front
        var ordered = boxes.OrderByDescending(b => Math.Round(projection.DepthOf((b.X0 + b.X1) / 2, (b.Y0 + b.Y1) / 2, (b.Z0 + b.Z1) / 2), 9))
                           .ThenBy(b => view.RotX >= 0 ? b.Y0 : -b.Y1)
                           .ThenBy(b => view.RotY >= 0 ? b.X0 : -b.X1).ToList();
        foreach (var box in ordered)
        {
            var faces = new (double X, double Y, double Z, (double, double, double)[] Corners)[]
            {
                (0, 0, -1, new[] { (box.X0, box.Y0, box.Z0), (box.X1, box.Y0, box.Z0), (box.X1, box.Y1, box.Z0), (box.X0, box.Y1, box.Z0) }),   // front
                (0, 1, 0, new[] { (box.X0, box.Y1, box.Z0), (box.X1, box.Y1, box.Z0), (box.X1, box.Y1, box.Z1), (box.X0, box.Y1, box.Z1) }),    // top
                (1, 0, 0, new[] { (box.X1, box.Y0, box.Z0), (box.X1, box.Y0, box.Z1), (box.X1, box.Y1, box.Z1), (box.X1, box.Y1, box.Z0) }),    // right
                (-1, 0, 0, new[] { (box.X0, box.Y0, box.Z0), (box.X0, box.Y0, box.Z1), (box.X0, box.Y1, box.Z1), (box.X0, box.Y1, box.Z0) }),   // left
                (0, -1, 0, new[] { (box.X0, box.Y0, box.Z0), (box.X1, box.Y0, box.Z0), (box.X1, box.Y0, box.Z1), (box.X0, box.Y0, box.Z1) }),   // bottom
            };
            foreach (var face in faces)
            {
                var turned = projection.Turn((face.X, face.Y, face.Z));
                bool visible = view.RightAngled
                    ? face.Z < 0 || face.Y > 0 && view.RotX > 0 || face.Y < 0 && view.RotX < 0 || face.X > 0 && view.RotY > 0 || face.X < 0 && view.RotY < 0
                    : turned.Z > 1e-9;
                if (!visible) continue;
                string color = Shade(box.Color, Lighting(view, view.RightAngled ? projection.Turn((face.X, face.Y, face.Z)) : turned));
                Polygon(svg, face.Corners.Select(c => projection.Point(c.Item1, c.Item2, c.Item3)), color, color, "word-chart-box");
            }
            if (box.Label is { } label)
            {
                var p = projection.Point((box.X0 + box.X1) / 2, box.Y1, box.Z0);
                Text(svg, p.X, p.Y - 3, label, _dataLabelSize, textColor, "middle");
            }
        }
    }

    /// <summary>
    /// A 3-D pie: the slices as a tilted disc of some thickness - the elevation from rotX (plotareaconverter.cxx:275-279: rotX 0..90 to
    /// rotX - 90, so 90 looks straight down and 0 edge on), the first slice from rotY; the visible rim below each slice darker, the tops
    /// lit, drawn rim first (front half only) then the tops.
    /// </summary>
    private void Draw3DPie(StringBuilder svg, Group group, double left, double top, double right, double bottom, string textColor)
    {
        View3D view = _view3D ?? ReadView3D(null, true);
        Series series = group.Series[0];
        double total = series.Values.Sum(v => Math.Abs(v ?? 0));
        if (total <= 0) return;
        double tilt = Math.Max(0.15, Math.Sin(Math.Max(0, Math.Min(90, view.RotX)) * Math.PI / 180));
        double thicknessRatio = 0.12 * Math.Cos(Math.Max(0, Math.Min(90, view.RotX)) * Math.PI / 180);
        double rx = Math.Max(1, Math.Min((right - left) / 2 - 4, ((bottom - top) / 2 - 4) / (tilt + thicknessRatio)));
        double ry = rx * tilt, thickness = rx * thicknessRatio;
        double cx = (left + right) / 2, cy = (top + bottom) / 2 - thickness / 2;
        int count = series.Values.Count;
        var slices = new List<(double Start, double Sweep, string Color, int Index)>();
        double start = view.RotY != 0 ? view.RotY : group.FirstSliceAngle;
        double angle = (double.IsNaN(start) || double.IsInfinity(start) ? 0 : ((start % 360) + 360) % 360) - 90;
        for (int i = 0; i < count; i++)
        {
            double value = Math.Abs(series.Values[i] ?? 0);
            if (value <= 0) continue;
            double sweep = value / total * 360;
            slices.Add((angle, sweep, FillOf(group, series, group.VaryColors ? i : -1, count).Css, i));
            angle += sweep;
        }
        (double X, double Y) At(double degrees, double dy) { double r = degrees * Math.PI / 180; return (cx + rx * Math.Cos(r), cy + ry * Math.Sin(r) + dy); }
        double sideLight = view.Ambient + view.Light * 0.2, topLight = view.Ambient + view.Light;
        // rims: the part of each slice's arc on the front half (sin > 0), as a band down by the thickness
        foreach (var slice in slices)
        {
            double to = slice.Start + slice.Sweep;
            for (double a0 = slice.Start; a0 < to; )
            {
                double turn = Math.Floor(a0 / 360) * 360, frontFrom = Math.Max(a0, turn), frontTo = Math.Min(to, turn + 180);
                if (frontTo > frontFrom)
                {
                    var band = new List<(double X, double Y)>();
                    for (int k = 0; k <= 24; k++) band.Add(At(frontFrom + (frontTo - frontFrom) * k / 24, 0));
                    for (int k = 24; k >= 0; k--) band.Add(At(frontFrom + (frontTo - frontFrom) * k / 24, thickness));
                    string rim = Shade(slice.Color, sideLight);
                    Polygon(svg, band, rim, rim, "word-chart-pie-side");
                }
                a0 = turn + 360;
            }
        }
        foreach (var slice in slices)
        {
            var outline = new List<(double X, double Y)> { (cx, cy) };
            int steps = Math.Max(2, (int)Math.Ceiling(slice.Sweep / 5));
            for (int k = 0; k <= steps; k++) outline.Add(At(slice.Start + slice.Sweep * k / steps, 0));
            string lit = Shade(slice.Color, Math.Min(1, topLight));
            Polygon(svg, outline, lit, "#FFFFFF", "word-chart-pie-top");
            if (LabelText(group, series, slice.Index) is { } label)
            {
                var p = At(slice.Start + slice.Sweep / 2, 0);
                Text(svg, cx + (p.X - cx) * 0.65, cy + (p.Y - cy) * 0.65 + 3, label, _dataLabelSize, textColor, "middle");
            }
        }
    }
}
