using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>One drawn ODF shape: its outline and paint in its own box, the transform to the anchor's coordinates, and its text.</summary>
internal sealed class OdtDrawnShape
{
    public WordShape Shape = new();
    public XElement Element = null!;
    public OdtShapeStyle Style = null!;
    /// <summary>The text area in the shape's own points (a custom shape's text frame, else its box).</summary>
    public (double X, double Y, double W, double H) TextArea;
    /// <summary>The text area's transform to the anchor (the shape's rotation and position, no shear, no mirroring).</summary>
    public WordAffine TextTransform = WordAffine.Identity;
    public OdtEnhancedGeometry? Geometry;
    public bool HasText;
    /// <summary>A draw:measure's label: text, centre in the anchor's points, clockwise rotation.</summary>
    public (string Text, double X, double Y, double Rotation)? MeasureLabel;
}

/// <summary>
/// ODF drawing shapes (ODF 1.3 part 3 10.3) as <see cref="WordShape"/>s, positioned as LibreOffice positions them
/// (xmloff/source/draw/ximpshap.cxx, master read 2026-11-04):
/// - SetTransformation :539-616: the shape's box (svg:width / svg:height) is moved to svg:x / svg:y, then draw:transform
///   applies (xexptran.cxx:346-470 parse, GetFullTransform :480-540: rotate(a) turns by -a in y-down coordinates -
///   counter-clockwise on screen -, skewX(a) shears by -tan(a), skewY likewise, scale, translate with lengths, matrix(a..f));
///   each operation is applied after the ones before it (basegfx left multiplication, b2dhommatrix.cxx:113-215);
/// - draw:rect (:948-1009, draw:corner-radius only - svg:rx / svg:ry are not read), draw:line (:1011-1115, the points'
///   bounding box), draw:circle / draw:ellipse (:1117-1252, svg:cx / cy / r / rx / ry, draw:kind full / section / cut / arc
///   with draw:start-angle / draw:end-angle in degrees; svx svdocirc.cxx:242-300 ImpCalcXPolyCirc: a segment runs from
///   360 - end to 360 - start in y-down angles, a section adds the centre first, an arc stays open; a full circle starts at
///   the bottom), draw:polygon / draw:polyline (:1254-1355, draw:points in svg:viewBox mapped to the box from the view box's
///   own corner), draw:path (:1357-1506, svg:d likewise; closed only when every sub-path is closed - then filled - else
///   a poly-line that is not filled), draw:connector (:1804-2141, svg:d when its first and last point are the connector's
///   ends, else the straight line between them), draw:measure (:2143-2247, the measured line), draw:caption (a rectangle),
///   draw:custom-shape (<see cref="OdtEnhancedGeometry"/>; mirroring after the rotation about the snap rectangle's centre,
///   EnhancedCustomShapeEngine.cxx:228-310), draw:regular-polygon (not imported by LibreOffice - shapeimport.cxx:241-410 has
///   no case for it - drawn per ODF 1.3 part 3 19.123 / 19.128 / 19.211: corners on an ellipse filling the box, a star with
///   draw:concave, its inner points at (1 - sharpness) of the way to the centre, an odd count starting at the top of the
///   outer circle, an even one at the top of the inner; a convex polygon's first corner at the top - the spec gives no
///   start for it), draw:g (children in the same coordinates; a group's own draw:transform is not applied, as LibreOffice's
///   group context sets no transformation).
/// </summary>
internal sealed class OdtShapeBuilder
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw, Svg = OdtStyleCascade.Svg, Text = OdtStyleCascade.Text;
    private readonly OdtStyleCascade _styles;
    private readonly bool _gradientAngleIsTenths;
    private readonly Func<XElement, (string Id, (double W, double H)? Size)?>? _fillPicture;

    public OdtShapeBuilder(OdtStyleCascade styles, bool gradientAngleIsTenths, Func<XElement, (string Id, (double W, double H)? Size)?>? fillPicture)
    {
        _styles = styles;
        _gradientAngleIsTenths = gradientAngleIsTenths;
        _fillPicture = fillPicture;
    }

    public static bool IsShape(XElement element)
        => element.Name.Namespace == Draw && element.Name.LocalName is "custom-shape" or "rect" or "ellipse" or "circle" or "line" or "polygon" or "polyline"
                                                                     or "path" or "connector" or "measure" or "caption" or "regular-polygon" or "g";

    /// <summary>The element's shapes (a group's, flattened in drawing order).</summary>
    public List<OdtDrawnShape> Build(XElement element)
    {
        var result = new List<OdtDrawnShape>();
        Add(element, result, 0);
        return result;
    }

    private void Add(XElement element, List<OdtDrawnShape> result, int depth)
    {
        if (element.Name == Draw + "g")
        {
            if (depth > 256) return;
            foreach (XElement child in element.Elements())
                if (IsShape(child)) Add(child, result, depth + 1);
            return;
        }
        if (Shape(element) is { } shape) result.Add(shape);
    }

    private static double Length(XElement element, XName name) => OdtShapeStyle.Length((string?)element.Attribute(name)) ?? 0;

    /// <summary>draw:transform as a matrix (GetFullTransform, xexptran.cxx:480-540; lengths of translate in points).</summary>
    internal static WordAffine Transform(string? value)
    {
        WordAffine matrix = WordAffine.Identity;
        if (string.IsNullOrWhiteSpace(value)) return matrix;
        string s = value!;
        int i = 0;
        while (i < s.Length)
        {
            while (i < s.Length && (char.IsWhiteSpace(s[i]) || s[i] == ',')) i++;
            if (i >= s.Length) break;
            string? name = new[] { "rotate", "scale", "translate", "skewX", "skewY", "matrix" }.FirstOrDefault(n => string.CompareOrdinal(s, i, n, 0, n.Length) == 0);
            if (name is null) { i++; continue; }
            i += name.Length;
            while (i < s.Length && (char.IsWhiteSpace(s[i]) || s[i] == '(')) i++;
            int close = s.IndexOf(')', i);
            string arguments = close < 0 ? s.Substring(i) : s.Substring(i, close - i);
            i = close < 0 ? s.Length : close + 1;
            string[] parts = arguments.Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            double Number(int k, double fallback) => k < parts.Length && double.TryParse(parts[k], NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : fallback;
            double Measure(int k) => k < parts.Length ? OdtShapeStyle.Length(parts[k]) ?? 0 : 0;
            WordAffine step = name switch
            {
                "rotate" => WordAffine.Rotate(-Number(0, 0)),
                "scale" => WordAffine.Scale(Number(0, 1), Number(1, 1)),
                "translate" => WordAffine.Translate(Measure(0), Measure(1)),
                "skewX" => new WordAffine(1, 0, -Math.Tan(Number(0, 0)), 1, 0, 0),
                "skewY" => new WordAffine(1, -Math.Tan(Number(0, 0)), 0, 1, 0, 0),
                _ => new WordAffine(Number(0, 1), Number(1, 0), Number(2, 0), Number(3, 1), Measure(4), Measure(5)),
            };
            matrix = WordAffine.Then(matrix, step);
        }
        return matrix;
    }

    private OdtDrawnShape? Shape(XElement element)
    {
        var drawn = new OdtDrawnShape { Element = element, Style = new OdtShapeStyle(_styles, (string?)element.Attribute(Draw + "style-name"), _gradientAngleIsTenths) };
        WordShape shape = drawn.Shape;
        WordAffine transform = Transform((string?)element.Attribute(Draw + "transform") ?? (string?)element.Attribute(Svg + "transform"));
        double x = Length(element, Svg + "x"), y = Length(element, Svg + "y");
        double width = Length(element, Svg + "width"), height = Length(element, Svg + "height");
        string local = element.Name.LocalName;
        bool filled = drawn.Style.IsFilled, stroked = drawn.Style.IsStroked;
        WordAffine placement;
        List<WordShapePath>? paths = null;
        bool closedForFill = true;          // a poly-line / open path / arc / line is not filled
        switch (local)
        {
            case "rect":
                paths = [Path(RoundedRectangle(width, height, Length(element, Draw + "corner-radius")), fill: true)];
                placement = WordAffine.Translate(x, y);
                break;
            case "caption":
            {
                // the rectangle and its tail (svdocapt.cxx ImpCalcTail, OdtCaptionShape); the tail carries the markers
                OdtCaptionGeometry caption = OdtCaptionShape.Layout(element, drawn.Style.Get);
                var box = RoundedRectangle(caption.Width, caption.Height, Math.Min(caption.CornerRadiusX, caption.CornerRadiusY));
                var all = new List<WordShapeSubpath> { Shift(box, caption.Left, caption.Top, 1, 1, 0, 0), caption.Tail };
                (double l, double t, _, _) = Bounds(all);
                var fillPath = new WordShapePath { Fill = WordShapePathFill.Normal };
                fillPath.Subpaths.Add(Shift(all[0], -l, -t, 1, 1, 0, 0));
                var tailPath = new WordShapePath { Fill = WordShapePathFill.None };
                tailPath.Subpaths.Add(Shift(caption.Tail, -l, -t, 1, 1, 0, 0));
                paths = [fillPath, tailPath];
                (_, _, width, height) = Bounds(all);
                placement = WordAffine.Translate(l, t);
                drawn.TextArea = (caption.Left - l, caption.Top - t, caption.Width, caption.Height);
                break;
            }
            case "ellipse":
            case "circle":
            {
                if (element.Attribute(Svg + "cx") is not null || element.Attribute(Svg + "r") is not null || element.Attribute(Svg + "rx") is not null)
                {
                    double rx = element.Attribute(Svg + "r") is not null ? Length(element, Svg + "r") : Length(element, Svg + "rx");
                    double ry = element.Attribute(Svg + "r") is not null ? rx : Length(element, Svg + "ry");
                    if (Length(element, Svg + "cx") != 0 || Length(element, Svg + "cy") != 0 || rx != 0 || ry != 0)
                    { width = 2 * rx; height = 2 * ry; x = Length(element, Svg + "cx") - rx; y = Length(element, Svg + "cy") - ry; }
                }
                string kind = (string?)element.Attribute(Draw + "kind") ?? "full";
                double Angle(string name) => OdtFillPaint.TenthDegrees((string?)element.Attribute(Draw + name), unitlessIsTenths: false) / 10.0;     // convertAngle: degrees (converter.cxx:952-977)
                (WordShapeSubpath outline, bool isClosed) = Ellipse(width, height, kind, Angle("start-angle"), Angle("end-angle"));
                closedForFill = isClosed;
                paths = [Path(outline, fill: isClosed)];
                placement = WordAffine.Translate(x, y);
                break;
            }
            case "line":
            case "measure":
            {
                double x1 = Length(element, Svg + "x1"), y1 = Length(element, Svg + "y1"), x2 = Length(element, Svg + "x2"), y2 = Length(element, Svg + "y2");
                if (local == "line" && element.Attribute(Svg + "x2") is null && element.Attribute(Svg + "y2") is null) { x2 = 1 * 72 / 2540.0; y2 = 1 * 72 / 2540.0; }   // mnX2 / mnY2 default 1
                var line = new WordShapeSubpath { StartX = x1, StartY = y1 };
                line.Segments.Add(WordShapeSegment.Line(x2, y2));
                closedForFill = false;
                if (local == "measure")
                {
                    // dimension and help lines (sdrmeasureprimitive2d.cxx, OdtMeasureShape); the label is laid out by the renderer
                    var marker = OdtMeasureArrow.FromMarker(0, 1, 1, false);
                    string label = OdtMeasureShape.Label(element, drawn.Style.Get);
                    double fontSize = 12;
                    OdtMeasureGeometry measure = OdtMeasureShape.Layout(element, drawn.Style.Get, label.Length * fontSize * 0.5, fontSize * 1.2, marker, marker);
                    var lines = measure.MainLines.Select(m => m.Line).Concat(measure.HelpLines).ToList();
                    (paths, placement, width, height) = Absolute(lines, fill: false);
                    drawn.MeasureLabel = (label, measure.TextCenterX, measure.TextCenterY, measure.TextRotationDegrees);
                    break;
                }
                (paths, placement, width, height) = Absolute([line], fill: false);
                break;
            }
            case "polygon":
            case "polyline":
            {
                bool closed = local == "polygon";
                if (OdtSvgPath.Points((string?)element.Attribute(Draw + "points")) is not { Count: > 0 } points) return null;
                var outline = new WordShapeSubpath { StartX = points[0].X, StartY = points[0].Y, Closed = closed };
                foreach (var p in points.Skip(1)) outline.Segments.Add(WordShapeSegment.Line(p.X, p.Y));
                List<WordShapeSubpath> mapped = ViewBoxMap([outline], element, width, height);
                closedForFill = closed;
                (paths, placement, _, _) = Absolute(mapped, fill: closed);
                placement = WordAffine.Then(placement, WordAffine.Translate(x, y));
                (width, height) = (Bounds(mapped).W, Bounds(mapped).H);
                break;
            }
            case "path":
            {
                if (OdtSvgPath.Parse((string?)element.Attribute(Svg + "d")) is not { Count: > 0 } subpaths) return null;
                List<WordShapeSubpath> mapped = ViewBoxMap(subpaths, element, width, height);
                bool allClosed = mapped.All(s => s.Closed);
                closedForFill = allClosed;
                (paths, placement, _, _) = Absolute(mapped, fill: allClosed);
                placement = WordAffine.Then(placement, WordAffine.Translate(x, y));
                (width, height) = (Bounds(mapped).W, Bounds(mapped).H);
                break;
            }
            case "connector":
            {
                double x1 = Length(element, Svg + "x1"), y1 = Length(element, Svg + "y1"), x2 = Length(element, Svg + "x2"), y2 = Length(element, Svg + "y2");
                if ((string?)element.Attribute(Draw + "start-shape") is null && (string?)element.Attribute(Draw + "end-shape") is null && x1 == x2 && y1 == y2
                    && element.Attribute(Draw + "line-skew") is null) return null;    // empty connectors are not added (:1984-2002)
                (x1, y1) = transform.Apply(x1, y1); (x2, y2) = transform.Apply(x2, y2);     // #121965# the transform moves the ends (:2011-2026)
                transform = WordAffine.Identity;
                // svg:d of a connector has no view box: its numbers are the document's 1/100 mm (ximpshap.cxx:1957-1972)
                List<WordShapeSubpath>? track = OdtSvgPath.Parse((string?)element.Attribute(Svg + "d"))?.Select(t => Shift(t, 0, 0, 72 / 2540.0, 72 / 2540.0, 0, 0)).ToList();
                const double Tolerance = 0.5 * 72 / 2540;             // the ends are compared in whole hundredths of a millimetre
                WordShapeSegment? last = track is { Count: > 0 } && track[track.Count - 1].Segments.Count > 0 ? track[track.Count - 1].Segments[track[track.Count - 1].Segments.Count - 1] : null;
                if (track is not { Count: > 0 } || last is null || Math.Abs(track[0].StartX - x1) > Tolerance || Math.Abs(track[0].StartY - y1) > Tolerance
                    || Math.Abs(last.Value.X - x2) > Tolerance || Math.Abs(last.Value.Y - y2) > Tolerance)
                {
                    // no usable stored route: LibreOffice's own routing (svdoedge.cxx ImpCalcEdgeTrack, OdtConnectorRoute); a connected
                    // shape may stand in another block - read windowed, ends are taken where the file puts them
                    double[] skew = ((string?)element.Attribute(Draw + "line-skew") ?? "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
                        .Select(v => OdtShapeStyle.Length(v) ?? 0).ToArray();
                    track = OdtConnectorRoute.Route((string?)element.Attribute(Draw + "type") ?? "standard", (x1, y1), (x2, y2), skew, null, null);
                }
                closedForFill = false;
                (paths, placement, width, height) = Absolute(track, fill: false);
                break;
            }
            case "regular-polygon":
            {
                paths = [Path(RegularPolygon(element, width, height), fill: true)];
                placement = WordAffine.Translate(x, y);
                break;
            }
            case "custom-shape":
            {
                OdtEnhancedGeometry geometry = OdtEnhancedGeometry.Read(element.Element(Draw + "enhanced-geometry"), width, height, filled, stroked);
                drawn.Geometry = geometry;
                if (geometry.IsRectangle) paths = [Path(RoundedRectangle(width, height, 0), fill: true)];
                else
                {
                    paths = geometry.Engine?.Build() ?? [];
                    if (geometry.Engine?.TextRect() is { } textRect) drawn.TextArea = textRect;
                }
                placement = WordAffine.Translate(x, y);
                break;
            }
            default:
                return null;
        }
        shape.Geometry = WordShapeGeometry.Path;
        shape.Width = Math.Max(width, 0);
        shape.Height = Math.Max(height, 0);
        shape.Paths = paths;
        WordAffine full = WordAffine.Then(placement, transform);
        if (drawn.Geometry is { } g && (g.MirrorHorizontal || g.MirrorVertical))
        {
            // mirrored after the rotation, about the transformed box's centre (EnhancedCustomShapeEngine.cxx:286-298)
            (double cx, double cy) = full.Apply(shape.Width / 2, shape.Height / 2);
            WordAffine mirror = WordAffine.Then(WordAffine.Then(WordAffine.Translate(-cx, -cy), WordAffine.Scale(g.MirrorHorizontal ? -1 : 1, g.MirrorVertical ? -1 : 1)), WordAffine.Translate(cx, cy));
            shape.Transform = WordAffine.Then(full, mirror);
        }
        else shape.Transform = full;
        // the text turns with the shape (its rotation; shear and mirroring do not apply to text, svdoashp.cxx:2638-2658)
        (double ox, double oy) = full.Apply(0, 0);
        double angle = Math.Atan2(full.B, full.A);
        drawn.TextTransform = WordAffine.Then(WordAffine.Rotate(angle), WordAffine.Translate(ox, oy));
        if (drawn.TextArea.W <= 0 || drawn.TextArea.H <= 0) drawn.TextArea = (0, 0, shape.Width, shape.Height);
        drawn.HasText = element.Elements().Any(e => e.Name == Text + "p" || e.Name == Text + "h" || e.Name == Text + "list");

        // paint
        drawn.Style.ApplyFill(shape, (0, 0, shape.Width, shape.Height), _fillPicture);
        if (!closedForFill) { shape.FillColor = null; shape.FillPainter = null; shape.FillPictureId = null; }
        drawn.Style.ApplyStroke(shape);
        if (shape.LineJoin == WordLineJoin.Miter) shape.MiterLimit = 1 / Math.Sin(7.5 * Math.PI / 180);
        if (drawn.Geometry?.Extrusion is { } extrusion) Extrude(drawn, extrusion, full);
        return drawn;
    }

    /// <summary>
    /// An extruded custom shape: its outline objects become the projected faces (<see cref="OdtExtrusion"/>). The 3-D scene takes the
    /// shape's rotation and mirroring itself and the result is not rotated or mirrored again (EnhancedCustomShapeEngine.cxx:262-266);
    /// a shear of the box still applies after (:270-282). Faces have no outline unless the shade mode is draft (Create3DObject :379-388).
    /// </summary>
    private void Extrude(OdtDrawnShape drawn, OdtExtrusion extrusion, WordAffine full)
    {
        WordShape shape = drawn.Shape;
        OdtEnhancedGeometry geometry = drawn.Geometry!;
        var sources = new List<OdtExtrusionSource>();
        foreach (WordShapePath path in shape.Paths ?? [])
        {
            // a path with neither fill nor line is a placeholder that is removed (:415-416, :1083-1086); the shape's own colours are used,
            // not the path's darken / lighten (the 3-D object takes the shape's item set, :418); filled or not follows the shape's fill
            // style, not the path's (aLocalFillStyle, EnhancedCustomShape3d.cxx:378, :606-613)
            if (path.Fill == WordShapePathFill.None && !path.Stroke) continue;
            bool filled = drawn.Style.IsFilled;
            var source = new OdtExtrusionSource { Filled = filled };
            foreach (WordShapeSubpath subpath in path.Subpaths)
                source.Rings.Add((OdtShapeStyle.Flatten(subpath), filled || subpath.Closed));
            sources.Add(source);
        }
        var fill = OdtShapeStyle.ParseColor(drawn.Style.Get(Draw + "fill-color")) ?? OdtShapeStyle.ParseColor(OdtShapeStyle.DefaultFillColor)!.Value;
        var secondary = OdtShapeStyle.ParseColor(drawn.Style.Get(Draw + "secondary-fill-color")) ?? OdtShapeStyle.ParseColor(OdtShapeStyle.DefaultFillColor)!.Value;  // xpool.cxx:200
        var line = OdtShapeStyle.ParseColor(drawn.Style.Get(Svg + "stroke-color")) ?? OdtShapeStyle.ParseColor(OdtShapeStyle.DefaultStrokeColor)!.Value;
        bool textured = drawn.Style.FillKind is "gradient" or "bitmap";
        double rotationDegrees = -Math.Atan2(full.B, full.A) * 180 / Math.PI;      // draw:transform rotate(a) is Rotate(-a) here
        var faces = extrusion.Project(sources, shape.Width, shape.Height, rotationDegrees, geometry.MirrorHorizontal, geometry.MirrorVertical, fill, secondary, line, textured);
        var paths = new List<WordShapePath>();
        foreach (OdtExtrusionFace face in faces)
        {
            var path = new WordShapePath { Fill = face.OutlineOnly ? WordShapePathFill.None : WordShapePathFill.Normal, Stroke = face.OutlineOnly, SolidFill = face.OutlineOnly ? null : face.Color };
            foreach (var ring in face.Rings)
            {
                if (ring.Count == 0) continue;
                var subpath = new WordShapeSubpath { StartX = ring[0].X, StartY = ring[0].Y, Closed = true };
                for (int i = 1; i < ring.Count; i++) subpath.Segments.Add(WordShapeSegment.Line(ring[i].X, ring[i].Y));
                path.Subpaths.Add(subpath);
            }
            paths.Add(path);
        }
        if (paths.Count == 0) return;   // nothing to extrude: no scene, the flat 2-D shape stays (EnhancedCustomShape3d.cxx:620)
        shape.Paths = paths;
        if (extrusion.ShadeMode == "draft") { shape.LineColor = OdtShapeStyle.Hex(line); shape.Dash = null; }
        else shape.LineColor = null;
        shape.HeadEnd = shape.TailEnd = default;
        shape.HeadMarker = shape.TailMarker = null;
        // the box's centre where the transform puts it; the rotation taken out of the transform (shear and scale stay)
        (double cx, double cy) = full.Apply(shape.Width / 2, shape.Height / 2);
        WordAffine linear = new(full.A, full.B, full.C, full.D, 0, 0);
        WordAffine unrotated = WordAffine.Then(linear, WordAffine.Rotate(-Math.Atan2(full.B, full.A)));
        shape.Transform = WordAffine.Then(WordAffine.Then(WordAffine.Translate(-shape.Width / 2, -shape.Height / 2), unrotated), WordAffine.Translate(cx, cy));
    }

    private static WordShapePath Path(WordShapeSubpath subpath, bool fill)
    {
        var path = new WordShapePath { Fill = fill ? WordShapePathFill.Normal : WordShapePathFill.None };
        path.Subpaths.Add(subpath);
        return path;
    }

    private static (double X, double Y, double W, double H) Bounds(IEnumerable<WordShapeSubpath> subpaths)
    {
        double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
        foreach (WordShapeSubpath s in subpaths)
            foreach ((double px, double py) in OdtShapeStyle.Flatten(s))
            { left = Math.Min(left, px); top = Math.Min(top, py); right = Math.Max(right, px); bottom = Math.Max(bottom, py); }
        return left == double.MaxValue ? (0, 0, 0, 0) : (left, top, right - left, bottom - top);
    }

    /// <summary>Point coordinates as the shape's own box: shifted to their bounding box's corner, the shift as the placement.</summary>
    private static (List<WordShapePath> paths, WordAffine placement, double width, double height) Absolute(List<WordShapeSubpath> subpaths, bool fill)
    {
        (double left, double top, double w, double h) = Bounds(subpaths);
        var path = new WordShapePath { Fill = fill ? WordShapePathFill.Normal : WordShapePathFill.None };
        foreach (WordShapeSubpath s in subpaths) path.Subpaths.Add(Shift(s, -left, -top, 1, 1, 0, 0));
        return ([path], WordAffine.Translate(left, top), w, h);
    }

    private static WordShapeSubpath Shift(WordShapeSubpath s, double dx, double dy, double sx, double sy, double ox, double oy)
    {
        (double X, double Y) P(double px, double py) => (ox + (px - ox) * sx + dx, oy + (py - oy) * sy + dy);
        var start = P(s.StartX, s.StartY);
        var result = new WordShapeSubpath { StartX = start.X, StartY = start.Y, Closed = s.Closed };
        foreach (WordShapeSegment g in s.Segments)
        {
            var e = P(g.X, g.Y);
            if (g.IsCurve) { var c1 = P(g.Control1X, g.Control1Y); var c2 = P(g.Control2X, g.Control2Y); result.Segments.Add(new WordShapeSegment(true, c1.X, c1.Y, c2.X, c2.Y, e.X, e.Y)); }
            else result.Segments.Add(WordShapeSegment.Line(e.X, e.Y));
        }
        return result;
    }

    /// <summary>svg:viewBox mapped to the box (svg:width / svg:height when both are set, else the view box's size), from the view box's own corner (:1307-1330).</summary>
    private static List<WordShapeSubpath> ViewBoxMap(List<WordShapeSubpath> subpaths, XElement element, double width, double height)
    {
        if (OdtSvgPath.ViewBox((string?)element.Attribute(Svg + "viewBox")) is not { } box) return subpaths;
        // the view box is in 1/100 mm units of the drawing (LibreOffice's SdXMLImExViewBox reads plain numbers)
        double toPoints = 72 / 2540.0;
        double targetW = width != 0 && height != 0 ? width : box.Width * toPoints, targetH = width != 0 && height != 0 ? height : box.Height * toPoints;
        double sx = box.Width != 0 ? targetW / box.Width : 1, sy = box.Height != 0 ? targetH / box.Height : 1;
        // source (vx, vy, vw, vh) -> target (vx, vy, targetW, targetH): x' = vx + (x - vx) * sx, in the document unit
        return subpaths.Select(s => Shift(s, box.X * toPoints - box.X, box.Y * toPoints - box.Y, sx, sy, box.X, box.Y)).ToList();
    }

    /// <summary>A rectangle with corner radius r (SdrRectObj: radii limited to half the sides; quarter ellipses as Bezier curves, basegfx createPolygonFromRect).</summary>
    internal static WordShapeSubpath RoundedRectangle(double w, double h, double radius)
    {
        var s = new WordShapeSubpath { Closed = true };
        double rx = Math.Min(radius, w / 2), ry = Math.Min(radius, h / 2);
        if (rx <= 0 || ry <= 0)
        {
            s.StartX = 0; s.StartY = 0;
            s.Segments.Add(WordShapeSegment.Line(w, 0)); s.Segments.Add(WordShapeSegment.Line(w, h)); s.Segments.Add(WordShapeSegment.Line(0, h));
            return s;
        }
        const double Kappa = 0.5522847498307935;
        double kx = rx * Kappa, ky = ry * Kappa;
        s.StartX = rx; s.StartY = 0;
        s.Segments.Add(WordShapeSegment.Line(w - rx, 0));
        s.Segments.Add(new WordShapeSegment(true, w - rx + kx, 0, w, ry - ky, w, ry));
        s.Segments.Add(WordShapeSegment.Line(w, h - ry));
        s.Segments.Add(new WordShapeSegment(true, w, h - ry + ky, w - rx + kx, h, w - rx, h));
        s.Segments.Add(WordShapeSegment.Line(rx, h));
        s.Segments.Add(new WordShapeSegment(true, rx - kx, h, 0, h - ry + ky, 0, h - ry));
        s.Segments.Add(WordShapeSegment.Line(0, ry));
        s.Segments.Add(new WordShapeSegment(true, 0, ry - ky, rx - kx, 0, rx, 0));
        return s;
    }

    /// <summary>ImpCalcXPolyCirc (svdocirc.cxx:242-300).</summary>
    private static (WordShapeSubpath outline, bool closed) Ellipse(double w, double h, string kind, double startDegrees, double endDegrees)
    {
        double cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2;
        var s = new WordShapeSubpath();
        if (kind == "full")
        {
            // createPolygonFromUnitCircle(1): from the bottom, increasing angle (clockwise on screen)
            var points = WordCustomShapeEngine.EllipseSegment((cx, cy), rx, ry, Math.PI / 2, Math.PI / 2 + 2 * Math.PI);
            s.StartX = points[0].X; s.StartY = points[0].Y;
            foreach (var p in points.Skip(1).Take(points.Count - 2)) s.Segments.Add(WordShapeSegment.Line(p.X, p.Y));
            s.Closed = true;
            return (s, true);
        }
        double Normal(double degrees) { double d = (360 - degrees) % 360; return (d < 0 ? d + 360 : d) * Math.PI / 180; }
        double start = Normal(endDegrees), end = Normal(startDegrees);
        var arc = WordCustomShapeEngine.EllipseSegment((cx, cy), rx, ry, start, end);
        if (kind == "section") arc.Insert(0, (cx, cy));
        s.StartX = arc[0].X; s.StartY = arc[0].Y;
        foreach (var p in arc.Skip(1)) s.Segments.Add(WordShapeSegment.Line(p.X, p.Y));
        bool closed = kind != "arc";
        s.Closed = closed;
        return (s, closed);
    }

    /// <summary>draw:regular-polygon (ODF 1.3 part 3 19.123 draw:concave, 19.128 draw:corners, 19.211 draw:sharpness).</summary>
    private static WordShapeSubpath RegularPolygon(XElement element, double w, double h)
    {
        int corners = Math.Min(1000, Math.Max(3, int.TryParse((string?)element.Attribute(Draw + "corners"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int c) ? c : 3));      // over 1000 corners is a circle at any size on screen
        bool concave = (string?)element.Attribute(Draw + "concave") == "true";
        double sharpness = OdtFillPaint.Percent((string?)element.Attribute(Draw + "sharpness")) ?? 0;
        var points = new List<(double X, double Y)>();
        if (!concave)
            for (int i = 0; i < corners; i++) { double a = -Math.PI / 2 + 2 * Math.PI * i / corners; points.Add((Math.Cos(a), Math.Sin(a))); }
        else
        {
            // R is the middle of the shorter arc between two outer points: its radius 1; T at (1 - s) of MR
            double inner = 1 - sharpness;
            bool odd = (corners & 1) == 1;
            for (int i = 0; i < 2 * corners; i++)
            {
                double a = -Math.PI / 2 + Math.PI * i / corners;
                bool outer = odd ? (i & 1) == 0 : (i & 1) == 1;
                double r = outer ? 1 : inner;
                points.Add((r * Math.Cos(a), r * Math.Sin(a)));
            }
        }
        // the basic model scaled so that its bounding box is the box
        double minX = points.Min(p => p.X), maxX = points.Max(p => p.X), minY = points.Min(p => p.Y), maxY = points.Max(p => p.Y);
        double sx = maxX > minX ? w / (maxX - minX) : 0, sy = maxY > minY ? h / (maxY - minY) : 0;
        var s = new WordShapeSubpath { Closed = true, StartX = (points[0].X - minX) * sx, StartY = (points[0].Y - minY) * sy };
        foreach (var p in points.Skip(1)) s.Segments.Add(WordShapeSegment.Line((p.X - minX) * sx, (p.Y - minY) * sy));
        return s;
    }
}
