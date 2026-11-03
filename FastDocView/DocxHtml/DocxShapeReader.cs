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
/// - fill: no fill unless the theme style (wps:style a:fillRef) or the shape's own fill says so, the shape's own winning
///   (shape.cxx:3122-3141); line: likewise a:lnRef then a:ln (shape.cxx:1777-1784). A style reference is drawn as a solid
///   colour of its own colour (the theme's fill style list is not read: approximation) and a line of its colour with the
///   theme line style's width (theme.cxx:41-58).
/// - a:gradFill: gs pos in 100000ths (oox/source/drawingml/misccontexts.cxx:74-76), a:lin ang in 60000ths of a degree (:99);
///   drawn as an SVG linear gradient (path / radial gradients approximated as linear).
/// - a:headEnd is the line's start, a:tailEnd its end (oox/source/drawingml/linepropertiescontext.cxx:216-221); every end
///   type other than none is drawn as a triangle (approximation).
/// - custGeom paths: a:path w / h are the path's own coordinate space (customshapegeometry.cxx:997-998), scaled to the shape;
///   moveTo / lnTo / cubicBezTo / quadBezTo / arcTo / close. Points must be numbers (guide names are not evaluated: such a
///   shape is drawn as a box). Unverified: an a:path without w / h is taken to use the shape's own size in EMU.
/// - presets drawn: rect, roundRect, ellipse, line, straightConnector1, triangle, right / left / up / down arrows; every other
///   preset is a box of the right size.
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
            Walk(child, shapes, theme, pictureId, ref textIndex, x => x, y => y, 1, 1, frameWidth, frameHeight, isTop: true);
        return shapes;
    }

    /// <summary>
    /// One element of the tree. <paramref name="mapX"/> / <paramref name="mapY"/> turn this level's EMU coordinates into frame
    /// points, <paramref name="scaleX"/> / <paramref name="scaleY"/> its EMU lengths; a top-level shape or group fills the frame.
    /// </summary>
    private static void Walk(XElement element, List<WordShape> shapes, DocxTheme? theme, PictureIdResolver pictureId, ref int textIndex,
                             Func<double, double> mapX, Func<double, double> mapY, double scaleX, double scaleY, double frameWidth, double frameHeight, bool isTop)
    {
        switch (element.Name.LocalName)
        {
            case "wgp" or "grpSp":
            {
                XElement? transform = element.Element("grpSpPr")?.Element("xfrm");
                (double x, double y, double width, double height) = Box(transform, mapX, mapY, scaleX, scaleY, frameWidth, frameHeight, isTop);
                double childOffsetX = Attribute(transform?.Element("chOff"), "x"), childOffsetY = Attribute(transform?.Element("chOff"), "y");
                double childWidth = Attribute(transform?.Element("chExt"), "cx"), childHeight = Attribute(transform?.Element("chExt"), "cy");
                double childScaleX = childWidth > 0 ? width / childWidth : scaleX, childScaleY = childHeight > 0 ? height / childHeight : scaleY;
                foreach (XElement child in element.Elements())
                    Walk(child, shapes, theme, pictureId, ref textIndex, cx => x + (cx - childOffsetX) * childScaleX, cy => y + (cy - childOffsetY) * childScaleY,
                         childScaleX, childScaleY, frameWidth, frameHeight, isTop: false);
                break;
            }
            case "wsp":
            {
                XElement? properties = element.Element("spPr");
                XElement? transform = properties?.Element("xfrm");
                (double x, double y, double width, double height) = Box(transform, mapX, mapY, scaleX, scaleY, frameWidth, frameHeight, isTop);
                WordShape shape = Shape(properties, element.Element("style"), theme, x, y, width, height);
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

    private static WordShape Shape(XElement? properties, XElement? style, DocxTheme? theme, double x, double y, double width, double height)
    {
        XElement? transform = properties?.Element("xfrm");
        var shape = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = Attribute(transform, "rot") / 60000,
            FlipHorizontal = IsOn(transform, "flipH"), FlipVertical = IsOn(transform, "flipV"),
        };
        if (properties?.Element("prstGeom") is { } preset)
        {
            shape.Geometry = (string?)preset.Attribute("prst") switch
            {
                "rect" => WordShapeGeometry.Rectangle,
                "roundRect" => WordShapeGeometry.RoundRectangle,
                "ellipse" => WordShapeGeometry.Ellipse,
                "line" or "straightConnector1" => WordShapeGeometry.Line,
                "triangle" => WordShapeGeometry.Triangle,
                "rightArrow" => WordShapeGeometry.RightArrow,
                "leftArrow" => WordShapeGeometry.LeftArrow,
                "upArrow" => WordShapeGeometry.UpArrow,
                "downArrow" => WordShapeGeometry.DownArrow,
                _ => WordShapeGeometry.Box,
            };
            shape.Adjust = AdjustValues(preset.Element("avLst"));
        }
        else if (properties?.Element("custGeom") is { } custom && CustomPath(custom, width, height) is { } path)
        {
            shape.Geometry = WordShapeGeometry.Path;
            shape.PathData = path;
        }
        else shape.Geometry = properties?.Element("custGeom") is not null ? WordShapeGeometry.Box : WordShapeGeometry.Rectangle;

        // fill: the style reference, then the shape's own
        XElement? fillReference = style?.Element("fillRef");
        string? placeholder = DocxDrawingColor.Hex(fillReference, theme);
        if (fillReference is not null && Attribute(fillReference, "idx") > 0) shape.FillColor = DocxDrawingColor.Css(fillReference, theme);
        XElement? ownFill = properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill");
        switch (ownFill?.Name.LocalName)
        {
            case "noFill": shape.FillColor = null; break;
            case "solidFill": shape.FillColor = DocxDrawingColor.Css(ownFill, theme, placeholder); break;
            case "gradFill":
                var stops = new List<(double, string)>();
                foreach (XElement stop in ownFill.Element("gsLst")?.Elements("gs") ?? [])
                    if (DocxDrawingColor.Css(stop, theme, placeholder) is { } color)
                        stops.Add((Math.Max(0, Math.Min(1, Attribute(stop, "pos") / 100000)), color));
                if (stops.Count > 0) { shape.GradientStops = stops; shape.FillColor = stops[0].Item2; }
                shape.GradientAngleDegrees = Attribute(ownFill.Element("lin"), "ang") / 60000;
                break;
        }

        // line: the style reference, then the shape's own a:ln
        XElement? lineReference = style?.Element("lnRef");
        int lineIndex = (int)Attribute(lineReference, "idx");
        string? linePlaceholder = DocxDrawingColor.Hex(lineReference, theme);
        if (lineReference is not null && lineIndex > 0)
        {
            shape.LineColor = DocxDrawingColor.Css(lineReference, theme);
            long themeWidth = theme?.LineStyleWidthEmu(lineIndex) ?? -1;
            shape.LineWidthPt = themeWidth > 0 ? themeWidth / EmuPerPoint : 0;
        }
        if (properties?.Element("ln") is { } line)
        {
            if (line.Attribute("w") is not null) shape.LineWidthPt = Attribute(line, "w") / EmuPerPoint;
            if (line.Element("noFill") is not null) shape.LineColor = null;
            else if (line.Element("solidFill") is { } lineFill) shape.LineColor = DocxDrawingColor.Css(lineFill, theme, linePlaceholder);
            else if (line.Element("gradFill")?.Element("gsLst")?.Element("gs") is { } firstStop) shape.LineColor = DocxDrawingColor.Css(firstStop, theme, linePlaceholder);
            shape.Dash = (string?)line.Element("prstDash")?.Attribute("val") switch
            {
                // ST_PresetLineDashVal names as dash / space lengths in line widths: an approximation, the exact patterns not checked against the spec or Word
                "dot" or "sysDot" => [1, 1],
                "dash" or "sysDash" => [3, 1],
                "lgDash" => [8, 3],
                "dashDot" or "sysDashDot" => [4, 3, 1, 3],
                "lgDashDot" => [8, 3, 1, 3],
                "lgDashDotDot" or "sysDashDotDot" => [8, 3, 1, 3, 1, 3],
                _ => null,
            };
            shape.HeadArrow = (string?)line.Element("headEnd")?.Attribute("type") is { } head && head != "none";
            shape.TailArrow = (string?)line.Element("tailEnd")?.Attribute("type") is { } tail && tail != "none";
        }
        return shape;
    }

    /// <summary>a:avLst gd fmla="val n" values in order (only constants: Word writes those).</summary>
    private static double[] AdjustValues(XElement? list)
    {
        if (list is null) return [];
        var values = new List<double>();
        foreach (XElement guide in list.Elements("gd"))
        {
            string formula = ((string?)guide.Attribute("fmla") ?? "").Trim();
            if (formula.StartsWith("val ", StringComparison.Ordinal) && double.TryParse(formula.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) values.Add(value);
            else return values.ToArray();
        }
        return values.ToArray();
    }

    /// <summary>The custGeom outline as SVG path data in the shape's points, or null when a point is not a number.</summary>
    private static string? CustomPath(XElement custom, double width, double height)
    {
        var d = new StringBuilder();
        foreach (XElement path in custom.Element("pathLst")?.Elements("path") ?? [])
        {
            double pathWidth = Attribute(path, "w"), pathHeight = Attribute(path, "h");
            double sx = width / (pathWidth > 0 ? pathWidth : width * EmuPerPoint), sy = height / (pathHeight > 0 ? pathHeight : height * EmuPerPoint);
            if (double.IsNaN(sx) || double.IsInfinity(sx)) sx = 0;
            if (double.IsNaN(sy) || double.IsInfinity(sy)) sy = 0;
            double currentX = 0, currentY = 0;
            foreach (XElement command in path.Elements())
            {
                var points = new List<(double x, double y)>();
                foreach (XElement point in command.Elements("pt"))
                {
                    if (!double.TryParse((string?)point.Attribute("x"), NumberStyles.Float, CultureInfo.InvariantCulture, out double px)
                        || !double.TryParse((string?)point.Attribute("y"), NumberStyles.Float, CultureInfo.InvariantCulture, out double py)) return null;
                    points.Add((px * sx, py * sy));
                }
                string letter;
                switch (command.Name.LocalName)
                {
                    case "moveTo" when points.Count == 1: letter = "M"; break;
                    case "lnTo" when points.Count == 1: letter = "L"; break;
                    case "cubicBezTo" when points.Count == 3: letter = "C"; break;
                    case "quadBezTo" when points.Count == 2: letter = "Q"; break;
                    case "close": d.Append('Z'); continue;
                    case "arcTo":
                    {
                        // an arc of the ellipse with radii wR / hR from angle stAng sweeping swAng (60000ths of a degree) starting at the current point;
                        // read as the attribute names say, the spec text not consulted (unverified)
                        if (!TryNumber(command, "wR", out double radiusX) || !TryNumber(command, "hR", out double radiusY)
                            || !TryNumber(command, "stAng", out double start) || !TryNumber(command, "swAng", out double sweep)) return null;
                        radiusX *= sx; radiusY *= sy;
                        double startRadians = start / 60000 * Math.PI / 180, endRadians = (start + sweep) / 60000 * Math.PI / 180;
                        double centerX = currentX - radiusX * Math.Cos(startRadians), centerY = currentY - radiusY * Math.Sin(startRadians);
                        currentX = centerX + radiusX * Math.Cos(endRadians); currentY = centerY + radiusY * Math.Sin(endRadians);
                        d.Append('A').Append(WordShapeHtml.Number(radiusX)).Append(' ').Append(WordShapeHtml.Number(radiusY)).Append(" 0 ")
                         .Append(Math.Abs(sweep) > 180 * 60000 ? '1' : '0').Append(' ').Append(sweep > 0 ? '1' : '0').Append(' ')
                         .Append(WordShapeHtml.Number(currentX)).Append(' ').Append(WordShapeHtml.Number(currentY));
                        continue;
                    }
                    default: return null;
                }
                d.Append(letter);
                for (int i = 0; i < points.Count; i++) d.Append(i == 0 ? "" : " ").Append(WordShapeHtml.Number(points[i].x)).Append(' ').Append(WordShapeHtml.Number(points[i].y));
                (currentX, currentY) = points[points.Count - 1];
            }
        }
        return d.Length == 0 ? null : d.ToString();
    }

    private static bool TryNumber(XElement element, string name, out double value)
        => double.TryParse((string?)element.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out value);

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
    /// = rectangle, v:roundrect (arcsize: a fraction, or n "f" in 65536ths; unset = the VML default, unverified here, taken as
    /// 0.2), v:oval, v:line (from / to). Fill: filled unless filled="f", fillcolor default white; line: stroked unless
    /// stroked="f", strokecolor default black, strokeweight default 1 EMU = the thinnest line (LibreOffice
    /// oox/source/vml/vmlformatting.cxx:706-712 and :757-770). Colours #RRGGBB / #RGB, else a colour name (the part before a
    /// space; vmlformatting.cxx:240-279). Text insets: v:textbox inset, unset parts 0.1in / 0.05in / 0.1in / 0.05in
    /// (vmltextboxcontext.cxx:186-211; insetmode auto = the defaults). style rotation in degrees, flip x / y.
    /// </summary>
    public static WordShape ReadVml(XElement layout, Dictionary<string, string> style, double width, double height)
    {
        string kind = (string?)layout.Attribute("vmlShape") ?? "rect";
        var shape = new WordShape { Width = width, Height = height };
        shape.Geometry = kind switch
        {
            "roundrect" => WordShapeGeometry.RoundRectangle,
            "oval" => WordShapeGeometry.Ellipse,
            "line" => WordShapeGeometry.Line,
            _ => WordShapeGeometry.Rectangle,
        };
        if (kind == "roundrect")
        {
            string arc = ((string?)layout.Attribute("arcsize") ?? "0.2").Trim();
            double fraction = arc.EndsWith("f", StringComparison.Ordinal) && double.TryParse(arc.TrimEnd('f'), NumberStyles.Float, CultureInfo.InvariantCulture, out double f) ? f / 65536
                            : double.TryParse(arc.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out double plain) ? (arc.EndsWith("%", StringComparison.Ordinal) ? plain / 100 : plain) : 0.2;
            shape.Adjust = [fraction * 100000];
        }
        bool filled = !IsOff((string?)layout.Attribute("filled")) && !IsOff((string?)layout.Attribute("fillOn"));
        if (filled && kind != "line") shape.FillColor = VmlColor((string?)layout.Attribute("fillcolor")) ?? "#FFFFFF";
        bool stroked = !IsOff((string?)layout.Attribute("stroked")) && !IsOff((string?)layout.Attribute("strokeOn"));
        if (stroked)
        {
            shape.LineColor = VmlColor((string?)layout.Attribute("strokecolor")) ?? "#000000";
            shape.LineWidthPt = (string?)layout.Attribute("strokeweight") is { } weight ? DocxBlockHtmlRenderer.VmlLengthPoints(weight.Contains("pt") || weight.Any(char.IsLetter) ? weight : weight + "pt") : 1 / EmuPerPoint;
            string? dash = (string?)layout.Attribute("dashstyle");
            if (dash is not null && dash != "solid") shape.Dash = dash.Contains("dot") && !dash.Contains("dash") ? [1, 1] : [3, 1];
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
