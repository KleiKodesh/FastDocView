using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// Turns the shape tree <see cref="Docx.DocxDrawingStripper"/> kept (prefix-free DrawingML) or a VML shape's layout element
/// into <see cref="WordShape"/>s in the drawing frame's points. Sources (checked online 2026-11-03, LibreOffice core master):
/// - groups: a child's a:off / a:ext are in its group's child space; it lands at (x - chOff) * ext / chExt from the group's
///   own offset, a missing or zero chExt meaning scale 1 (oox/source/drawingml/shape.cxx:1433-1456). The top group fills
///   the frame (wp:extent).
/// - a:xfrm rot in 60000ths of a degree clockwise, flipH / flipV (oox/source/drawingml/transform2dcontext.cxx:48-49).
/// - geometry: <see cref="DocxShapeGeometry"/> (every preset, guides, custom paths).
/// - fill (Shape::getActualFillProperties, shape.cxx:3122-3148): none, then the theme fill style a:fillRef idx names
///   (theme.cxx:49-54, its phClr = the fillRef colour), then the shape's own fill over it (FillProperties::assignUsed:
///   a gradient without its own stops keeps the style's), a:grpFill = the parent group's fill.
/// - line (getActualLineProperties, shape.cxx:3150-3170): none, then the theme a:ln a:lnRef idx names (phClr = the lnRef
///   colour), then the shape's own a:ln, each attribute / child used only when set (LineProperties::assignUsed,
///   lineproperties.cxx:407-430); its colour: a solid fill's, a gradient's first stop (its second when it has more than two,
///   FillProperties::getBestSolidColor fillproperties.cxx:401-424).
/// - dashes (lineproperties.cxx:60-84 lclConvertPresetDash): a preset is dots of DotLen and dashes of DashLen each followed
///   by Distance, in line widths; with a round or square cap (a preset) or a round cap (a custom dash) the cap is inside the
///   length, so LibreOffice takes 0.99 off each dash and adds it to the gap (:451-462), and the SVG cap adds it back. A
///   custom dash (a:custDash a:ds d / sp, 1000ths of a percent of the line width, linepropertiescontext.cxx:135-185) is
///   drawn stop by stop.
/// - gradients: gs pos in 100000ths (misccontexts.cxx:74-76), a:lin ang in 60000ths of a degree (:99), a:path circle =
///   radial about the a:fillToRect centre (fillproperties.cxx:528-551); rotWithShape="0" keeps the angle off the shape's
///   rotation (:521-525). a:path rect / shape and blip / pattern fills are not drawn (the fill below them stays).
/// - a:headEnd is the line's start, a:tailEnd its end (oox/source/drawingml/linepropertiescontext.cxx:216-221).
/// - alpha: every colour's opacity (DocxDrawingColor) goes to fill-opacity / stroke-opacity / stop-opacity.
/// </summary>
internal static class DocxShapeReader
{
    private const double EmuPerPoint = 12700;

    public delegate string? PictureIdResolver(string embed);

    public static List<WordShape> Read(XElement shapeTree, double frameWidth, double frameHeight, DocxTheme? theme, PictureIdResolver pictureId)
    {
        var shapes = new List<WordShape>();
        int textIndex = 0;
        foreach (XElement child in shapeTree.Elements())
            Walk(child, shapes, theme, pictureId, ref textIndex, x => x, y => y, 1, 1, frameWidth, frameHeight, isTop: true, groupFill: null);
        return shapes;
    }

    /// <summary>
    /// One element of the tree. <paramref name="mapX"/> / <paramref name="mapY"/> turn this level's EMU coordinates into frame
    /// points, <paramref name="scaleX"/> / <paramref name="scaleY"/> its EMU lengths; a top-level shape or group fills the frame.
    /// </summary>
    private static void Walk(XElement element, List<WordShape> shapes, DocxTheme? theme, PictureIdResolver pictureId, ref int textIndex,
                             Func<double, double> mapX, Func<double, double> mapY, double scaleX, double scaleY, double frameWidth, double frameHeight, bool isTop,
                             ShapeFill? groupFill)
    {
        switch (element.Name.LocalName)
        {
            case "wgp" or "grpSp":
            {
                XElement? properties = element.Element("grpSpPr");
                XElement? transform = properties?.Element("xfrm");
                (double x, double y, double width, double height) = Box(transform, mapX, mapY, scaleX, scaleY, frameWidth, frameHeight, isTop);
                double childOffsetX = Attribute(transform?.Element("chOff"), "x"), childOffsetY = Attribute(transform?.Element("chOff"), "y");
                double childWidth = Attribute(transform?.Element("chExt"), "cx"), childHeight = Attribute(transform?.Element("chExt"), "cy");
                double childScaleX = childWidth > 0 ? width / childWidth : scaleX, childScaleY = childHeight > 0 ? height / childHeight : scaleY;
                ShapeFill? fill = ShapeFill.Apply(ShapeFill.None, OwnFill(properties), theme, null, groupFill);
                foreach (XElement child in element.Elements())
                    Walk(child, shapes, theme, pictureId, ref textIndex, cx => x + (cx - childOffsetX) * childScaleX, cy => y + (cy - childOffsetY) * childScaleY,
                         childScaleX, childScaleY, frameWidth, frameHeight, isTop: false, fill);
                break;
            }
            case "wsp":
            {
                XElement? properties = element.Element("spPr");
                XElement? transform = properties?.Element("xfrm");
                (double x, double y, double width, double height) = Box(transform, mapX, mapY, scaleX, scaleY, frameWidth, frameHeight, isTop);
                WordShape shape = Shape(properties, element.Element("style"), theme, x, y, width, height, groupFill);
                if (element.Element("txbx") is not null) { shape.TextIndex = textIndex++; shape.TextStyle = TextStyle(element.Element("bodyPr")); }
                shapes.Add(shape);
                break;
            }
            case "pic":
            {
                XElement? transform = element.Element("spPr")?.Element("xfrm");
                (double x, double y, double width, double height) = Box(transform, mapX, mapY, scaleX, scaleY, frameWidth, frameHeight, isTop);
                string? embed = (string?)element.Element("blipFill")?.Element("blip")?.Attribute("embed");
                shapes.Add(new WordShape
                {
                    Geometry = WordShapeGeometry.Picture, X = x, Y = y, Width = width, Height = height,
                    RotationDegrees = Attribute(transform, "rot") / 60000, PictureId = embed is { Length: > 0 } ? pictureId(embed) : null,
                });
                break;
            }
        }
    }

    private static (double x, double y, double width, double height) Box(XElement? transform, Func<double, double> mapX, Func<double, double> mapY,
                                                                         double scaleX, double scaleY, double frameWidth, double frameHeight, bool isTop)
    {
        if (isTop || transform is null) return (0, 0, frameWidth, frameHeight);
        return (mapX(Attribute(transform.Element("off"), "x")), mapY(Attribute(transform.Element("off"), "y")),
                Attribute(transform.Element("ext"), "cx") * scaleX, Attribute(transform.Element("ext"), "cy") * scaleY);
    }

    private static XElement? OwnFill(XElement? properties)
        => properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "grpFill" or "blipFill" or "pattFill");

    private static WordShape Shape(XElement? properties, XElement? style, DocxTheme? theme, double x, double y, double width, double height, ShapeFill? groupFill)
    {
        XElement? transform = properties?.Element("xfrm");
        var shape = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = Attribute(transform, "rot") / 60000,
            FlipHorizontal = IsOn(transform, "flipH"), FlipVertical = IsOn(transform, "flipV"),
        };
        double ownWidth = Attribute(transform?.Element("ext"), "cx"), ownHeight = Attribute(transform?.Element("ext"), "cy");
        XElement? geometry = properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "prstGeom" or "custGeom");
        if (geometry is not null && DocxShapeGeometry.Paths(geometry, width, height, ownWidth, ownHeight) is { } paths)
        {
            shape.Geometry = WordShapeGeometry.Path;
            shape.Paths = paths;
        }
        else if (geometry is null)
        {
            // no geometry: a rectangle (the ECMA-376 default shape)
            shape.Geometry = WordShapeGeometry.Path;
            shape.Paths = DocxShapeGeometry.Paths(new XElement("prstGeom", new XAttribute("prst", "rect")), width, height, ownWidth, ownHeight);
        }
        else shape.Geometry = WordShapeGeometry.Box;   // a name that is no ECMA-376 preset, or a custom geometry without a path

        // fill: none, the theme style, the shape's own
        XElement? fillReference = style?.Element("fillRef");
        string? fillPlaceholder = DocxDrawingColor.Hex(fillReference, theme);
        ShapeFill fill = ShapeFill.None;
        if (fillReference is not null && theme?.FillStyle((int)Attribute(fillReference, "idx")) is { } themeFill)
            fill = ShapeFill.Apply(fill, themeFill, theme, fillPlaceholder, groupFill) ?? fill;
        fill = ShapeFill.Apply(fill, OwnFill(properties), theme, fillPlaceholder, groupFill) ?? fill;
        fill.ApplyTo(shape);

        // line: none, the theme a:ln, the shape's own
        XElement? lineReference = style?.Element("lnRef");
        string? linePlaceholder = DocxDrawingColor.Hex(lineReference, theme);
        var line = new ShapeLine();
        if (lineReference is not null && theme?.LineStyle((int)Attribute(lineReference, "idx")) is { } themeLine) line.Apply(themeLine, theme, linePlaceholder);
        if (properties?.Element("ln") is { } ownLine) line.Apply(ownLine, theme, linePlaceholder);
        line.ApplyTo(shape);
        return shape;
    }

    /// <summary>A resolved DrawingML fill (oox FillProperties as far as it is drawn here).</summary>
    private sealed class ShapeFill
    {
        public static readonly ShapeFill None = new() { Kind = FillKind.None };
        public enum FillKind { None, Solid, Gradient }
        public FillKind Kind;
        public DocxResolvedColor Color;
        public List<(double Position, DocxResolvedColor Color)>? Stops;
        public double? AngleDegrees;
        public bool? RotateWithShape;
        public string? PathKind;
        public double CenterX = 0.5, CenterY = 0.5;

        /// <summary>
        /// <paramref name="element"/> (a fill element) over <paramref name="under"/> as assignUsed does: noFill / solidFill replace
        /// it, a gradFill keeps the stops, angle and path it does not set itself, a:grpFill takes the group's; a blip or
        /// pattern fill (not drawn) or no element leaves <paramref name="under"/>. Null only when grpFill has no group.
        /// </summary>
        public static ShapeFill? Apply(ShapeFill under, XElement? element, DocxTheme? theme, string? placeholder, ShapeFill? groupFill)
        {
            switch (element?.Name.LocalName)
            {
                case "noFill": return None;
                case "solidFill":
                    return DocxDrawingColor.Resolve(element, theme, placeholder) is { } color ? new ShapeFill { Kind = FillKind.Solid, Color = color } : under;
                case "grpFill": return groupFill ?? under;
                case "gradFill":
                {
                    var fill = new ShapeFill { Kind = FillKind.Gradient };
                    ShapeFill? baseGradient = under.Kind == FillKind.Gradient ? under : null;
                    if (element.Element("gsLst") is { } list)
                    {
                        fill.Stops = new List<(double, DocxResolvedColor)>();
                        foreach (XElement stop in list.Elements("gs"))
                            if (DocxDrawingColor.Resolve(stop, theme, placeholder) is { } stopColor)
                                fill.Stops.Add((Math.Max(0, Math.Min(1, Attribute(stop, "pos") / 100000)), stopColor));
                        fill.Stops.Sort((a, b) => a.Position.CompareTo(b.Position));
                    }
                    else fill.Stops = baseGradient?.Stops;
                    if (element.Element("lin") is { } linear) { fill.AngleDegrees = Attribute(linear, "ang") / 60000; fill.PathKind = null; }
                    else if (element.Element("path") is { } path)
                    {
                        fill.PathKind = (string?)path.Attribute("path") ?? "circle";
                        XElement? rect = path.Element("fillToRect");
                        // fillproperties.cxx:530-536: centre = (100000 + l - r) / 2, (100000 + t - b) / 2, limited to 0..100 %
                        fill.CenterX = Math.Max(0, Math.Min(1, (100000 + Attribute(rect, "l") - Attribute(rect, "r")) / 2 / 100000));
                        fill.CenterY = Math.Max(0, Math.Min(1, (100000 + Attribute(rect, "t") - Attribute(rect, "b")) / 2 / 100000));
                    }
                    else if (baseGradient is not null)
                    {
                        fill.AngleDegrees = baseGradient.AngleDegrees; fill.PathKind = baseGradient.PathKind;
                        fill.CenterX = baseGradient.CenterX; fill.CenterY = baseGradient.CenterY;
                    }
                    fill.RotateWithShape = element.Attribute("rotWithShape") is not null ? IsOn(element, "rotWithShape") : baseGradient?.RotateWithShape;
                    return fill.Stops is { Count: > 0 } ? fill : under;
                }
                default: return under;
            }
        }

        public void ApplyTo(WordShape shape)
        {
            switch (Kind)
            {
                case FillKind.Solid:
                    shape.FillColor = Color.Css; shape.FillOpacity = Color.Opacity;
                    break;
                case FillKind.Gradient when PathKind is null or "circle":
                    shape.GradientStops = Stops!.Select(s => (s.Position, s.Color.Css)).ToList();
                    shape.GradientStopOpacities = Stops!.Select(s => s.Color.Opacity).ToList();
                    shape.FillColor = Stops![0].Color.Css;
                    shape.GradientRadial = PathKind == "circle";
                    shape.GradientCenterX = CenterX; shape.GradientCenterY = CenterY;
                    double angle = AngleDegrees ?? 0;
                    if (RotateWithShape == false) angle -= shape.FlipHorizontal ^ shape.FlipVertical ? -shape.RotationDegrees : shape.RotationDegrees;
                    shape.GradientAngleDegrees = angle;
                    break;
                case FillKind.Gradient:
                    // a:path rect / shape: LibreOffice's rectangular gradient is not drawn here; the first stop's colour fills the shape
                    shape.FillColor = Stops![0].Color.Css; shape.FillOpacity = Stops[0].Color.Opacity;
                    break;
            }
        }
    }

    /// <summary>A resolved DrawingML line (oox LineProperties): each setting used only when an a:ln sets it.</summary>
    private sealed class ShapeLine
    {
        private ShapeFill _fill = ShapeFill.None;
        private double? _widthEmu;
        private string? _cap, _presetDash;
        private List<(double Dash, double Space)>? _customDash;
        private string? _headType, _headWidth, _headLength, _tailType, _tailWidth, _tailLength;

        public void Apply(XElement line, DocxTheme? theme, string? placeholder)
        {
            if (line.Attribute("w") is not null) _widthEmu = Attribute(line, "w");
            if ((string?)line.Attribute("cap") is { } cap) _cap = cap;
            XElement? fill = line.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "pattFill");
            if (fill?.Name.LocalName == "pattFill")
            {
                // getBestSolidColor: the pattern's background colour, else its foreground (fillproperties.cxx:419-421)
                if ((DocxDrawingColor.Resolve(fill.Element("bgClr"), theme, placeholder) ?? DocxDrawingColor.Resolve(fill.Element("fgClr"), theme, placeholder)) is { } pattern)
                    _fill = new ShapeFill { Kind = ShapeFill.FillKind.Solid, Color = pattern };
            }
            else if (fill is not null) _fill = ShapeFill.Apply(_fill, fill, theme, placeholder, null) ?? _fill;
            if ((string?)line.Element("prstDash")?.Attribute("val") is { } preset) _presetDash = preset;
            if (line.Element("custDash") is { } custom)
            {
                var stops = new List<(double, double)>();
                foreach (XElement stop in custom.Elements("ds")) stops.Add((DashPercent((string?)stop.Attribute("d")), DashPercent((string?)stop.Attribute("sp"))));
                if (stops.Count > 0) _customDash = stops;
            }
            if (line.Element("headEnd") is { } head) { _headType = (string?)head.Attribute("type") ?? _headType; _headWidth = (string?)head.Attribute("w") ?? _headWidth; _headLength = (string?)head.Attribute("len") ?? _headLength; }
            if (line.Element("tailEnd") is { } tail) { _tailType = (string?)tail.Attribute("type") ?? _tailType; _tailWidth = (string?)tail.Attribute("w") ?? _tailWidth; _tailLength = (string?)tail.Attribute("len") ?? _tailLength; }
        }

        /// <summary>ST_PositivePercentage as linepropertiescontext.cxx:139-178 reads it: "n%" or 1000ths of a percent; in line widths.</summary>
        private static double DashPercent(string? value)
        {
            if (value is null) return 0;
            value = value.Trim();
            if (value.EndsWith("%", StringComparison.Ordinal))
                return int.TryParse(value.Substring(0, value.Length - 1).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int percent) ? percent / 100.0 : 0;
            return int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int thousandths) ? thousandths / 100000.0 : 0;
        }

        public void ApplyTo(WordShape shape)
        {
            DocxResolvedColor? color = _fill.Kind switch
            {
                ShapeFill.FillKind.Solid => _fill.Color,
                ShapeFill.FillKind.Gradient => _fill.Stops![_fill.Stops.Count > 2 ? 1 : 0].Color,
                _ => null,
            };
            shape.LineWidthPt = (_widthEmu ?? 0) / EmuPerPoint;
            if (color is null) { shape.LineColor = null; return; }
            shape.LineColor = color.Value.Css;
            shape.LineOpacity = color.Value.Opacity;
            shape.LineCap = _cap switch { "rnd" => WordLineCap.Round, "sq" => WordLineCap.Square, _ => WordLineCap.Flat };
            shape.Dash = Dash(shape.LineCap);
            shape.HeadEnd = End(_headType, _headWidth, _headLength);
            shape.TailEnd = End(_tailType, _tailWidth, _tailLength);
            shape.HeadArrow = shape.HeadEnd.Type != WordLineEndType.None;
            shape.TailArrow = shape.TailEnd.Type != WordLineEndType.None;
        }

        private double[]? Dash(WordLineCap cap)
        {
            if (_presetDash is not null && _presetDash != "solid")
            {
                // lclConvertPresetDash (lineproperties.cxx:60-84): dots, dot length, dashes, dash length, distance (line widths)
                (int dots, double dotLength, int dashes, double dashLength, double distance) = _presetDash switch
                {
                    "dot" => (1, 1, 0, 0, 3),
                    "dash" => (1, 4, 0, 0, 3),
                    "dashDot" => (1, 4, 1, 1, 3),
                    "lgDash" => (1, 8, 0, 0, 3),
                    "lgDashDot" => (1, 8, 1, 1, 3),
                    "lgDashDotDot" => (1, 8, 2, 1, 3),
                    "sysDot" => (1, 1, 0, 0, 1),
                    "sysDash" => (1, 3, 0, 0, 1),
                    "sysDashDot" => (1, 3, 1, 1, 1),
                    "sysDashDotDot" => (1, 3, 2, 1, 1),
                    _ => (1, 4, 0, 0, 3),   // an unknown preset: LibreOffice's fallback "dash" (:78-79)
                };
                if (cap != WordLineCap.Flat)
                {
                    if (dotLength >= 1 || dashLength >= 1) distance += 0.99;
                    if (dotLength >= 1) dotLength -= 0.99;
                    if (dashLength >= 1) dashLength -= 0.99;
                }
                var pattern = new List<double>();
                for (int i = 0; i < dots; i++) { pattern.Add(dotLength); pattern.Add(distance); }
                for (int i = 0; i < dashes; i++) { pattern.Add(dashLength); pattern.Add(distance); }
                return pattern.ToArray();
            }
            if (_customDash is { Count: > 0 } custom)
            {
                var pattern = new List<double>();
                foreach ((double dash, double space) in custom)
                {
                    // a custom dash includes a round cap (and not a square one): lineproperties.cxx:451-454
                    bool trim = cap == WordLineCap.Round && dash >= 1;
                    pattern.Add(trim ? dash - 0.99 : dash);
                    pattern.Add(trim ? space + 0.99 : space);
                }
                return pattern.ToArray();
            }
            return null;
        }

        private static WordLineEnd End(string? type, string? width, string? length)
        {
            WordLineEndType kind = type switch
            {
                "triangle" => WordLineEndType.Triangle, "stealth" => WordLineEndType.Stealth, "diamond" => WordLineEndType.Diamond,
                "oval" => WordLineEndType.Oval, "arrow" => WordLineEndType.Arrow, _ => WordLineEndType.None,
            };
            static int Size(string? value) => value switch { "sm" => 0, "lg" => 2, _ => 1 };   // lclGetArrowSize: default medium
            return new WordLineEnd(kind, Size(width), Size(length));
        }
    }

    /// <summary>A shape's text area from wps:bodyPr (same defaults and readings as the plain text box: DocxBlockHtmlRenderer.TextBoxStyle).</summary>
    private static WordTextBoxStyle TextStyle(XElement? body)
    {
        double Inset(string name, double fallback) => body?.Attribute(name) is not null ? Attribute(body, name) / EmuPerPoint : fallback / EmuPerPoint;
        return new WordTextBoxStyle(Inset("lIns", 91440), Inset("tIns", 45720), Inset("rIns", 91440), Inset("bIns", 45720), null, 0, false, null, 0,
                                    DocxTextBoxWritingMode((string?)body?.Attribute("vert")), body?.Element("spAutoFit") is not null);
    }

    /// <summary>bodyPr vert as CSS (LibreOffice textbodypropertiescontext.cxx:133-158; see DocxBlockHtmlRenderer.TextBoxStyle).</summary>
    internal static string? DocxTextBoxWritingMode(string? vert) => vert switch
    {
        "eaVert" or "vert" => "writing-mode:vertical-rl;",
        "vert270" => "writing-mode:sideways-lr;",
        "mongolianVert" => "writing-mode:vertical-lr;",
        "wordArtVert" => "writing-mode:vertical-rl;text-orientation:upright;",
        _ => null,
    };

    private static double Attribute(XElement? element, string name)
        => double.TryParse((string?)element?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : 0;

    private static bool IsOn(XElement? element, string name) => (string?)element?.Attribute(name) is "1" or "true";

    // ── VML ────────────────────────────────────────────────────────────────────────────

    /// <summary>
    /// A VML shape (ECMA-376 Part 4 14.1) from its layout element, in its own frame (style width / height): v:rect / v:shape
    /// = rectangle, v:roundrect (arcsize: the corner radius as a fraction of half the shorter side, default 0.2: W3C
    /// NOTE-VML 1998 roundrect; LibreOffice vmlshape.cxx:1066-1077 likewise takes min(w, h) / 2 times n "f" / 65536 or
    /// n "%" / 100), v:oval, v:line (from / to). Fill: filled unless filled="f", fillcolor default white; line: stroked unless
    /// stroked="f", strokecolor default black, strokeweight default 1 EMU = the thinnest line (LibreOffice
    /// oox/source/vml/vmlformatting.cxx:706-712 and :757-770). Colours #RRGGBB / #RGB, else a colour name (the part before a
    /// space; vmlformatting.cxx:240-279). Text insets: v:textbox inset, unset parts 0.1in / 0.05in / 0.1in / 0.05in
    /// (vmltextboxcontext.cxx:186-211; insetmode auto = the defaults). style rotation in degrees, flip x / y.
    /// </summary>
    public static WordShape ReadVml(XElement layout, Dictionary<string, string> style, double width, double height)
    {
        string kind = (string?)layout.Attribute("vmlShape") ?? "rect";
        var shape = new WordShape { Width = width, Height = height };
        double arcFraction = 0.2;
        if (kind == "roundrect")
        {
            string arc = ((string?)layout.Attribute("arcsize") ?? "0.2").Trim();
            arcFraction = arc.EndsWith("f", StringComparison.Ordinal) && double.TryParse(arc.TrimEnd('f'), NumberStyles.Float, CultureInfo.InvariantCulture, out double f) ? f / 65536
                        : double.TryParse(arc.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out double plain) ? (arc.EndsWith("%", StringComparison.Ordinal) ? plain / 100 : plain) : 0.2;
        }
        // the outline as the DrawingML preset it is (rect, roundRect with radius = ss * adj / 100000, so arcsize 1 = adj 50000, ellipse) or the line
        shape.Geometry = WordShapeGeometry.Path;
        if (kind == "line")
        {
            var line = new WordShapeSubpath { StartX = 0, StartY = 0 };
            line.Segments.Add(WordShapeSegment.Line(width, height));
            var path = new WordShapePath { Fill = WordShapePathFill.None };
            path.Subpaths.Add(line);
            shape.Paths = [path];
        }
        else
        {
            var preset = new XElement("prstGeom", new XAttribute("prst", kind switch { "roundrect" => "roundRect", "oval" => "ellipse", _ => "rect" }));
            if (kind == "roundrect")
                preset.Add(new XElement("avLst", new XElement("gd", new XAttribute("name", "adj"), new XAttribute("fmla", "val " + Math.Round(arcFraction * 50000).ToString(CultureInfo.InvariantCulture)))));
            shape.Paths = DocxShapeGeometry.Paths(preset, width, height, 0, 0);
        }
        bool filled = !IsOff((string?)layout.Attribute("filled")) && !IsOff((string?)layout.Attribute("fillOn"));
        if (filled && kind != "line") shape.FillColor = VmlColor((string?)layout.Attribute("fillcolor")) ?? "#FFFFFF";
        bool stroked = !IsOff((string?)layout.Attribute("stroked")) && !IsOff((string?)layout.Attribute("strokeOn"));
        if (stroked)
        {
            double widthPt = (string?)layout.Attribute("strokeweight") is { } weight ? DocxBlockHtmlRenderer.VmlLengthPoints(weight.Contains("pt") || weight.Any(char.IsLetter) ? weight : weight + "pt") : 1 / EmuPerPoint;
            // the v:stroke settings as the DrawingML line LibreOffice turns them into (oox/source/vml/vmlformatting.cxx:565-664, :700-715)
            var line = new XElement("ln", new XAttribute("w", Math.Round(widthPt * EmuPerPoint).ToString(CultureInfo.InvariantCulture)),
                                    new XAttribute("cap", (string?)layout.Attribute("endcap") switch { "square" => "sq", "round" => "rnd", _ => "flat" }),
                                    new XElement("solidFill", new XElement("srgbClr", new XAttribute("val", "000000"))));
            if (VmlDash((string?)layout.Attribute("dashstyle")) is { } dash) line.Add(dash);
            line.Add(VmlArrow("headEnd", layout, "start"), VmlArrow("tailEnd", layout, "end"));
            var resolved = new ShapeLine();
            resolved.Apply(line, null, null);
            resolved.ApplyTo(shape);
            shape.LineColor = VmlColor((string?)layout.Attribute("strokecolor")) ?? "#000000";
        }
        if (style.TryGetValue("rotation", out string? rotation) && double.TryParse(rotation.TrimEnd('f', 'd', 'e', 'g'), NumberStyles.Float, CultureInfo.InvariantCulture, out double degrees))
            shape.RotationDegrees = degrees;
        if (style.TryGetValue("flip", out string? flip)) { shape.FlipHorizontal = flip.Contains('x'); shape.FlipVertical = flip.Contains('y'); }
        if (layout.Attribute("textBox") is not null)
        {
            shape.TextIndex = 0;
            double[] insets = [7.2, 3.6, 7.2, 3.6];
            if ((string?)layout.Attribute("insetmode") != "auto" && (string?)layout.Attribute("inset") is { } inset)
            {
                string[] parts = inset.Split(',');
                for (int i = 0; i < 4 && i < parts.Length; i++)
                    if (parts[i].Trim().Length > 0) insets[i] = DocxBlockHtmlRenderer.VmlLengthPoints(parts[i]);
            }
            shape.TextStyle = new WordTextBoxStyle(insets[0], insets[1], insets[2], insets[3], null, 0, false, null, 0, null,
                                                   layout.Attribute("autoFit") is not null || (style.TryGetValue("mso-fit-shape-to-text", out string? fit) && fit == "t"));
        }
        return shape;
    }

    /// <summary>VML dashstyle as LibreOffice lclGetDmlLineDash (vmlformatting.cxx:565-597) maps it to a DrawingML preset; a list of numbers
    /// is dash / space pairs in line widths (W3C NOTE-VML 1998, stroke dashstyle: "a length of 1 is equal to the line width", an odd last number ignored).</summary>
    private static XElement? VmlDash(string? dashStyle)
    {
        if (dashStyle is null) return null;
        string? preset = dashStyle.Trim() switch
        {
            "solid" => "solid", "shortdot" => "sysDot", "shortdash" => "sysDash", "shortdashdot" => "sysDashDot", "shortdashdotdot" => "sysDashDotDot",
            "dot" => "dot", "dash" => "dash", "dashdot" => "dashDot", "longdash" => "lgDash", "longdashdot" => "lgDashDot", "longdashdotdot" => "lgDashDotDot",
            _ => null,
        };
        if (preset is not null) return new XElement("prstDash", new XAttribute("val", preset));
        string[] numbers = dashStyle.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var custom = new XElement("custDash");
        for (int i = 0; i + 1 < numbers.Length; i += 2)
            if (double.TryParse(numbers[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double d) && double.TryParse(numbers[i + 1], NumberStyles.Float, CultureInfo.InvariantCulture, out double sp))
                custom.Add(new XElement("ds", new XAttribute("d", Math.Round(d * 100000).ToString(CultureInfo.InvariantCulture)), new XAttribute("sp", Math.Round(sp * 100000).ToString(CultureInfo.InvariantCulture))));
        return custom.HasElements ? custom : null;
    }

    /// <summary>A VML arrow (startarrow / endarrow and their width / length) as the DrawingML end: lclGetDmlArrowType / Width / Length (vmlformatting.cxx:599-633).</summary>
    private static XElement VmlArrow(string name, XElement layout, string prefix)
        => new(name,
               new XAttribute("type", (string?)layout.Attribute(prefix + "arrow") switch { "block" => "triangle", "classic" => "stealth", "diamond" => "diamond", "oval" => "oval", "open" => "arrow", _ => "none" }),
               new XAttribute("w", (string?)layout.Attribute(prefix + "arrowwidth") switch { "narrow" => "sm", "wide" => "lg", _ => "med" }),
               new XAttribute("len", (string?)layout.Attribute(prefix + "arrowlength") switch { "short" => "sm", "long" => "lg", _ => "med" }));

    private static bool IsOff(string? value) => value is "f" or "false" or "0" or "off";

    private static string? VmlColor(string? value)
    {
        if (value is null) return null;
        string name = value.Trim().Split(' ')[0];
        if (name.Length == 7 && name[0] == '#' && name.Skip(1).All(Uri.IsHexDigit)) return name.ToUpperInvariant();
        if (name.Length == 4 && name[0] == '#' && name.Skip(1).All(Uri.IsHexDigit))
            return ("#" + name[1] + name[1] + name[2] + name[2] + name[3] + name[3]).ToUpperInvariant();
        return name.Length > 0 && name.All(char.IsLetter) ? name.ToLowerInvariant() : null;
    }
}
