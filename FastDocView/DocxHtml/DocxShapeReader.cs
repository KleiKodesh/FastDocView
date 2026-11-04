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

    /// <summary>A picture's original size in points (its pixels at its resolution), for tiled picture fills; null when unknown.</summary>
    public delegate (double WidthPt, double HeightPt)? PictureSizeResolver(string embed);


    public static List<WordShape> Read(XElement shapeTree, double frameWidth, double frameHeight, DocxTheme? theme, PictureIdResolver pictureId, PictureSizeResolver? pictureSize = null)
    {
        var shapes = new List<WordShape>();
        int textIndex = 0;
        foreach (XElement child in shapeTree.Elements())
            Walk(child, shapes, theme, pictureId, pictureSize, ref textIndex, null, frameWidth, frameHeight, groupFill: null);
        return shapes;
    }

    /// <summary>
    /// A group as its children see it (LibreOffice Shape::createAndInsert, oox/source/drawingml/shape.cxx:1394-1493): its box
    /// in its parent's unrotated space, its child space (chOff, ext / chExt; a missing or zero chExt = scale 1), its own
    /// rotation and flips, and the flips and rotation it has on the page (a flip of an odd number of ancestors turns the
    /// rotation the other way: nOrientation :1481-1488).
    /// </summary>
    private sealed record GroupFrame(double X, double Y, double Width, double Height, double ChildOffsetX, double ChildOffsetY, double ChildScaleX, double ChildScaleY,
                                     double RotationDegrees, bool FlipHorizontal, bool FlipVertical, GroupFrame? Parent, bool PageFlipHorizontal, bool PageFlipVertical, double PageRotationDegrees)
    {
        /// <summary>A point of this group's space on the frame: mirrored at the group's centre, turned about it, then through the parents.</summary>
        public (double X, double Y) ToFrame(double x, double y)
        {
            double cx = X + Width / 2, cy = Y + Height / 2;
            if (FlipHorizontal) x = 2 * cx - x;
            if (FlipVertical) y = 2 * cy - y;
            if (RotationDegrees != 0)
            {
                double radians = RotationDegrees * Math.PI / 180, dx = x - cx, dy = y - cy;
                (x, y) = (cx + dx * Math.Cos(radians) - dy * Math.Sin(radians), cy + dx * Math.Sin(radians) + dy * Math.Cos(radians));
            }
            return Parent is null ? (x, y) : Parent.ToFrame(x, y);
        }

        public bool Mirrors => PageFlipHorizontal != PageFlipVertical;
    }

    /// <summary>
    /// A shape's box, flips and rotation on the frame: its a:xfrm in its group's child space scaled to the group (with MSO's
    /// width / height swap for a child turned 45-135 or 225-315 degrees, lcl_doSpecialMSOWidthHeightToggle shape.cxx:986-999,
    /// :1456-1473, and 180 degrees more under a mirroring parent), its centre carried through the groups, its flips toggled
    /// by theirs (:1406-1416), its rotation turned by their orientation and added to theirs (:1481-1491). A top-level element
    /// fills the frame.
    /// </summary>
    private static (double x, double y, double width, double height, double rotation, bool flipH, bool flipV) Place(XElement? transform, GroupFrame? group, double frameWidth, double frameHeight)
    {
        double rotation = Attribute(transform, "rot") / 60000;
        bool flipH = IsOn(transform, "flipH"), flipV = IsOn(transform, "flipV");
        if (group is null) return (0, 0, frameWidth, frameHeight, rotation, flipH, flipV);
        double offX = Attribute(transform?.Element("off"), "x"), offY = Attribute(transform?.Element("off"), "y");
        double cx = Attribute(transform?.Element("ext"), "cx"), cy = Attribute(transform?.Element("ext"), "cy");
        double centreX = group.X + (offX + cx / 2 - group.ChildOffsetX) * group.ChildScaleX, centreY = group.Y + (offY + cy / 2 - group.ChildOffsetY) * group.ChildScaleY;
        double width = cx * group.ChildScaleX, height = cy * group.ChildScaleY;
        int degrees = (int)(Attribute(transform, "rot") / 60000);
        bool toggle = (group.ChildScaleX != group.ChildScaleY) && ((degrees >= 45 && degrees < 135) || (degrees >= 225 && degrees < 315));
        double extra = 0;
        if (toggle)
        {
            (width, height) = (cx * group.ChildScaleY, cy * group.ChildScaleX);
            if (group.Mirrors) extra = 180;
        }
        (double x, double y) = group.ToFrame(centreX, centreY);
        return (x - width / 2, y - height / 2, width, height, (group.Mirrors ? -rotation : rotation) + group.PageRotationDegrees + extra,
                flipH ^ group.PageFlipHorizontal, flipV ^ group.PageFlipVertical);
    }

    /// <summary>One element of the tree inside <paramref name="group"/> (null: the drawing's top level, which fills the frame).</summary>
    private static void Walk(XElement element, List<WordShape> shapes, DocxTheme? theme, PictureIdResolver pictureId, PictureSizeResolver? pictureSize, ref int textIndex,
                             GroupFrame? group, double frameWidth, double frameHeight, ShapeFill? groupFill)
    {
        switch (element.Name.LocalName)
        {
            case "wgp" or "grpSp":
            {
                XElement? properties = element.Element("grpSpPr");
                XElement? transform = properties?.Element("xfrm");
                // a group keeps its own unrotated box in its parent's space; its turn and flips apply to its children
                double x, y, width, height;
                if (group is null) (x, y, width, height) = (0, 0, frameWidth, frameHeight);
                else
                {
                    double offX = Attribute(transform?.Element("off"), "x"), offY = Attribute(transform?.Element("off"), "y");
                    width = Attribute(transform?.Element("ext"), "cx") * group.ChildScaleX; height = Attribute(transform?.Element("ext"), "cy") * group.ChildScaleY;
                    x = group.X + (offX - group.ChildOffsetX) * group.ChildScaleX; y = group.Y + (offY - group.ChildOffsetY) * group.ChildScaleY;
                }
                double childWidth = Attribute(transform?.Element("chExt"), "cx"), childHeight = Attribute(transform?.Element("chExt"), "cy");
                double ownWidth = Attribute(transform?.Element("ext"), "cx"), ownHeight = Attribute(transform?.Element("ext"), "cy");
                double emuX = group?.ChildScaleX ?? (ownWidth > 0 ? width / ownWidth : 1 / EmuPerPoint), emuY = group?.ChildScaleY ?? (ownHeight > 0 ? height / ownHeight : 1 / EmuPerPoint);
                double rotation = Attribute(transform, "rot") / 60000;
                bool flipH = IsOn(transform, "flipH"), flipV = IsOn(transform, "flipV");
                bool parentMirrors = group?.Mirrors ?? false;
                var frame = new GroupFrame(x, y, width, height, Attribute(transform?.Element("chOff"), "x"), Attribute(transform?.Element("chOff"), "y"),
                                           childWidth > 0 ? width / childWidth : emuX, childHeight > 0 ? height / childHeight : emuY,
                                           rotation, flipH, flipV, group, flipH ^ (group?.PageFlipHorizontal ?? false), flipV ^ (group?.PageFlipVertical ?? false),
                                           (parentMirrors ? -rotation : rotation) + (group?.PageRotationDegrees ?? 0));
                ShapeFill? fill = ShapeFill.Apply(ShapeFill.None, OwnFill(properties), theme, null, groupFill);
                foreach (XElement child in element.Elements()) Walk(child, shapes, theme, pictureId, pictureSize, ref textIndex, frame, frameWidth, frameHeight, fill);
                break;
            }
            case "wpc":
            {
                // a drawing canvas (LibreOffice WordprocessingCanvasContext.cxx:34-52): a group whose child space is its own size, with a
                // background rectangle taking the wpc:bg fill and the wpc:whole line (:67-72)
                var frame = new GroupFrame(0, 0, frameWidth, frameHeight, 0, 0, 1 / EmuPerPoint, 1 / EmuPerPoint, 0, false, false, null, false, false, 0);
                var background = new XElement("spPr");
                foreach (XElement part in element.Elements().Where(e => e.Name.LocalName is "bg" or "whole")) background.Add(part.Elements());
                WordShape canvas = Shape(background, null, theme, 0, 0, frameWidth, frameHeight, null, pictureId, pictureSize, null);
                if (canvas.FillColor is not null || canvas.GradientStops is not null || canvas.LineColor is not null || canvas.FillPictureId is not null) shapes.Add(canvas);
                foreach (XElement child in element.Elements()) Walk(child, shapes, theme, pictureId, pictureSize, ref textIndex, frame, frameWidth, frameHeight, null);
                break;
            }
            case "spTree":
            {
                // a SmartArt drawing (dsp:drawing / dsp:spTree, diagramDrawing part): its shapes are placed in EMU from the graphic
                // frame's top left, as LibreOffice imports the drawing into the frame (oox diagram.cxx loadDiagram :615-631, the
                // ext drawing's shapes replace the layout); a group of the frame's own size
                var frame = new GroupFrame(0, 0, frameWidth, frameHeight, 0, 0, 1 / EmuPerPoint, 1 / EmuPerPoint, 0, false, false, null, false, false, 0);
                foreach (XElement child in element.Elements()) Walk(child, shapes, theme, pictureId, pictureSize, ref textIndex, frame, frameWidth, frameHeight, null);
                break;
            }
            case "sp":
            {
                // a dsp:sp: a shape with its own DrawingML text (dsp:txBody) in its text rectangle (dsp:txXfrm, in the drawing's space)
                XElement? properties = element.Element("spPr");
                XElement? transform = properties?.Element("xfrm");
                (double x, double y, double width, double height, double rotation, bool flipH, bool flipV) = Place(transform, group, frameWidth, frameHeight);
                WordShape shape = Shape(properties, element.Element("style"), theme, x, y, width, height, groupFill, pictureId, pictureSize, (rotation, flipH, flipV));
                if (element.Element("txBody") is { } body && DocxDrawingTextHtml.Text(body).Trim().Length > 0)
                {
                    shape.DrawingTextHtml = DocxDrawingTextHtml.Render(body, element.Element("style"), theme);
                    shape.DrawingTextAnchor = DocxDrawingTextHtml.Anchor(body);
                    (double left, double top, double right, double bottom) = DocxDrawingTextHtml.Insets(body);
                    shape.TextStyle = new WordTextBoxStyle(left, top, right, bottom, null, 0, false, null, 0,
                                                           DocxTextBoxWritingMode((string?)body.Element("bodyPr")?.Attribute("vert")), false);
                    if (element.Element("txXfrm") is { } text && group is not null)
                    {
                        double scaleX = group.ChildScaleX, scaleY = group.ChildScaleY;
                        shape.TextRect = ((Attribute(text.Element("off"), "x") - Attribute(transform?.Element("off"), "x")) * scaleX,
                                          (Attribute(text.Element("off"), "y") - Attribute(transform?.Element("off"), "y")) * scaleY,
                                          Attribute(text.Element("ext"), "cx") * scaleX, Attribute(text.Element("ext"), "cy") * scaleY);
                    }
                }
                shapes.Add(shape);
                break;
            }
            case "wsp":
            {
                XElement? properties = element.Element("spPr");
                (double x, double y, double width, double height, double rotation, bool flipH, bool flipV) = Place(properties?.Element("xfrm"), group, frameWidth, frameHeight);
                WordShape shape = Shape(properties, element.Element("style"), theme, x, y, width, height, groupFill, pictureId, pictureSize, (rotation, flipH, flipV));

                if (element.Element("txbx") is not null) { shape.TextIndex = textIndex++; shape.TextStyle = TextStyle(element.Element("bodyPr")); }
                shapes.Add(shape);
                break;
            }
            case "pic":
            {
                (double x, double y, double width, double height, double rotation, bool flipH, bool flipV) = Place(element.Element("spPr")?.Element("xfrm"), group, frameWidth, frameHeight);
                string? embed = (string?)element.Element("blipFill")?.Element("blip")?.Attribute("embed");
                shapes.Add(new WordShape
                {
                    Geometry = WordShapeGeometry.Picture, X = x, Y = y, Width = width, Height = height,
                    RotationDegrees = rotation, FlipHorizontal = flipH, FlipVertical = flipV, PictureId = embed is { Length: > 0 } ? pictureId(embed) : null,
                });
                break;
            }
        }
    }

    private static XElement? OwnFill(XElement? properties)
        => properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "noFill" or "solidFill" or "gradFill" or "grpFill" or "blipFill" or "pattFill");

    private static WordShape Shape(XElement? properties, XElement? style, DocxTheme? theme, double x, double y, double width, double height, ShapeFill? groupFill,
                                   PictureIdResolver? pictureId, PictureSizeResolver? pictureSize, (double Rotation, bool FlipH, bool FlipV)? placement)
    {
        XElement? transform = properties?.Element("xfrm");
        var shape = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = placement?.Rotation ?? Attribute(transform, "rot") / 60000,
            FlipHorizontal = placement?.FlipH ?? IsOn(transform, "flipH"), FlipVertical = placement?.FlipV ?? IsOn(transform, "flipV"),
        };
        double ownWidth = Attribute(transform?.Element("ext"), "cx"), ownHeight = Attribute(transform?.Element("ext"), "cy");
        XElement? geometry = properties?.Elements().FirstOrDefault(e => e.Name.LocalName is "prstGeom" or "custGeom");
        if (geometry is not null && DocxShapeGeometry.Paths(geometry, width, height, ownWidth, ownHeight, out var textRect) is { } paths)
        {
            shape.Geometry = WordShapeGeometry.Path;
            shape.Paths = paths;
            shape.TextRect = textRect;
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
        fill.ApplyTo(shape, pictureId, pictureSize);

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
        public enum FillKind { None, Solid, Gradient, Pattern, Picture }
        public FillKind Kind;
        public DocxResolvedColor Color;
        public List<(double Position, DocxResolvedColor Color)>? Stops;
        public double? AngleDegrees;
        public bool? RotateWithShape;
        public string? PathKind;
        public double CenterX = 0.5, CenterY = 0.5;
        // a:pattFill
        public string? PatternPreset;
        public DocxResolvedColor? Foreground, Background;
        // a:blipFill
        public string? Embed;
        public XElement? Blip;

        /// <summary>
        /// <paramref name="element"/> (a fill element) over <paramref name="under"/> as assignUsed does: noFill / solidFill /
        /// pattFill / blipFill replace it, a gradFill keeps the stops, angle and path it does not set itself, a:grpFill takes
        /// the group's; no element leaves <paramref name="under"/>. Null only when grpFill has no group.
        /// </summary>
        public static ShapeFill? Apply(ShapeFill under, XElement? element, DocxTheme? theme, string? placeholder, ShapeFill? groupFill)
        {
            switch (element?.Name.LocalName)
            {
                case "noFill": return None;
                case "solidFill":
                    return DocxDrawingColor.Resolve(element, theme, placeholder) is { } color ? new ShapeFill { Kind = FillKind.Solid, Color = color } : under;
                case "grpFill": return groupFill ?? under;
                case "pattFill":
                    return new ShapeFill
                    {
                        Kind = FillKind.Pattern, PatternPreset = (string?)element.Attribute("prst"),
                        Foreground = DocxDrawingColor.Resolve(element.Element("fgClr"), theme, placeholder), Background = DocxDrawingColor.Resolve(element.Element("bgClr"), theme, placeholder),
                    };
                case "blipFill":
                    return (string?)element.Element("blip")?.Attribute("embed") is { Length: > 0 } embed ? new ShapeFill { Kind = FillKind.Picture, Embed = embed, Blip = element } : under;
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

        public void ApplyTo(WordShape shape, PictureIdResolver? pictureId = null, PictureSizeResolver? pictureSize = null)
        {
            switch (Kind)
            {
                case FillKind.Solid:
                    shape.FillColor = Color.Css; shape.FillOpacity = Color.Opacity;
                    break;
                case FillKind.Gradient:
                    shape.GradientStops = Stops!.Select(s => (s.Position, s.Color.Css)).ToList();
                    shape.GradientStopOpacities = Stops!.Select(s => s.Color.Opacity).ToList();
                    shape.FillColor = Stops![0].Color.Css;
                    // fillproperties.cxx:539-547: circle = radial; rect and shape = LibreOffice's rectangular gradient ("shape" is not implemented there)
                    shape.GradientRadial = PathKind == "circle";
                    shape.GradientRectangular = PathKind is "rect" or "shape";
                    shape.GradientCenterX = CenterX; shape.GradientCenterY = CenterY;
                    double angle = AngleDegrees ?? 0;
                    if (RotateWithShape == false) angle -= shape.FlipHorizontal ^ shape.FlipVertical ? -shape.RotationDegrees : shape.RotationDegrees;
                    shape.GradientAngleDegrees = angle;
                    break;
                case FillKind.Pattern:
                    // the preset's cell in the foreground colour over the background colour (DocxPresetPattern); no foreground = the background solid
                    if (Foreground is { } foreground && PatternPreset is not null && DocxPresetPattern.Rows(PatternPreset) is { } rows)
                    {
                        shape.Pattern = new WordShapePattern(foreground.Css, foreground.Opacity, rows);
                        if (Background is { } back) { shape.FillColor = back.Css; shape.FillOpacity = back.Opacity; }
                    }
                    else if (Background is { } back) { shape.FillColor = back.Css; shape.FillOpacity = back.Opacity; }
                    break;
                case FillKind.Picture:
                    shape.FillPictureId = pictureId?.Invoke(Embed!);
                    shape.FillPicture = PictureFill(Blip!, shape, () => pictureSize?.Invoke(Embed!));
                    break;
            }
        }

        /// <summary>
        /// a:blipFill placement: a:stretch a:fillRect insets of the box (100000ths, negative = beyond; ECMA-376 20.1.8.30), or a:tile
        /// as LibreOffice reads it (fillproperties.cxx:624-649: tile size = the picture's original size x sx / sy, offset tx / ty in EMU,
        /// anchored at the algn point; flip is not used there either); alphaModFix = opacity (:749-750).
        /// </summary>
        private static WordShapePictureFill PictureFill(XElement blipFill, WordShape shape, Func<(double WidthPt, double HeightPt)?> originalSize)
        {
            double opacity = blipFill.Element("blip")?.Element("alphaModFix") is { } alpha ? Attribute(alpha, "amt") / 100000 : 1;
            if (blipFill.Element("tile") is { } tile && originalSize() is { } size)
            {
                double sx = tile.Attribute("sx") is not null ? Attribute(tile, "sx") / 100000 : 1, sy = tile.Attribute("sy") is not null ? Attribute(tile, "sy") / 100000 : 1;
                double tileWidth = Math.Max(0.01, size.WidthPt * sx), tileHeight = Math.Max(0.01, size.HeightPt * sy);
                string align = (string?)tile.Attribute("algn") ?? "tl";
                double originX = align switch { "t" or "ctr" or "b" => (shape.Width - tileWidth) / 2, "tr" or "r" or "br" => shape.Width - tileWidth, _ => 0 };
                double originY = align switch { "l" or "ctr" or "r" => (shape.Height - tileHeight) / 2, "bl" or "b" or "br" => shape.Height - tileHeight, _ => 0 };
                return new WordShapePictureFill(0, 0, 0, 0, true, originX + Attribute(tile, "tx") / EmuPerPoint, originY + Attribute(tile, "ty") / EmuPerPoint,
                                                tileWidth, tileHeight, null, opacity);
            }
            XElement? fillRect = blipFill.Element("stretch")?.Element("fillRect");
            return new WordShapePictureFill(Attribute(fillRect, "l") / 100000, Attribute(fillRect, "t") / 100000, Attribute(fillRect, "r") / 100000, Attribute(fillRect, "b") / 100000,
                                            false, 0, 0, 0, 0, null, opacity);
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
        if (style.TryGetValue("rotation", out string? rotation)) shape.RotationDegrees = VmlRotation(rotation);
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

    /// <summary>VML style rotation in degrees clockwise: a number, or n "fd" in 65536ths (LibreOffice ConversionHelper::decodeRotation, vmlformatting.cxx:130-152).</summary>
    internal static double VmlRotation(string value)
    {
        value = value.Trim();
        if (value.EndsWith("fd", StringComparison.Ordinal))
            return double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out double fixedPoint) ? fixedPoint / 65536 : 0;
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double degrees) ? degrees : 0;
    }

    private static Dictionary<string, string> VmlStyleOf(XElement element)
    {
        var style = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string declaration in ((string?)element.Attribute("style") ?? "").Split(';'))
        {
            int colon = declaration.IndexOf(':');
            if (colon > 0) style[declaration.Substring(0, colon).Trim()] = declaration.Substring(colon + 1).Trim();
        }
        return style;
    }

    private static double VmlNumber(Dictionary<string, string> style, string name)
        => style.TryGetValue(name, out string? value) && double.TryParse(new string(value.TakeWhile(c => char.IsDigit(c) || c is '-' or '+' or '.').ToArray()),
                                                                          NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : 0;

    private static (double X, double Y) VmlPair(string? value, (double, double) fallback)
    {
        if (value is null) return fallback;
        string[] parts = value.Split(',');
        double Part(int i, double otherwise) => i < parts.Length && double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : otherwise;
        return (Part(0, fallback.Item1), Part(1, fallback.Item2));
    }

    /// <summary>
    /// A VML group (the vgroup tree DocxDrawingStripper.WriteVmlGroupLayout keeps) as shapes in the frame's points: a child's
    /// style left / top / width / height are plain numbers in its group's coordinate system (coordorigin, default 0,0;
    /// coordsize, default 1000,1000, a zero part 1: LibreOffice ShapeType::getCoordSystem vmlshape.cxx:171-180,
    /// getRelRectangle :224-238 with a zero size 1), placed by lclGetAbsRect (:105-116) into the group's box; a v:line's from / to
    /// the same way (lclGetAbsPoint :95-103). Group rotation turns its children about its centre (:1621-1622, rotation as
    /// decodeRotation); flips (style flip x / y) mirror at the group's centre. The top group fills the frame. Pictures (an
    /// r:id on v:imagedata) are drawn as group pictures; the n-th child with a text box takes the n-th text box.
    /// </summary>
    public static List<WordShape> ReadVmlGroup(XElement layout, double frameWidth, double frameHeight, PictureIdResolver pictureId)
    {
        var shapes = new List<WordShape>();
        if (layout.Element("vgroup") is not { } root) return shapes;
        int textIndex = 0;
        WalkVml(root, null, 0, 0, frameWidth, frameHeight, shapes, pictureId, ref textIndex);
        return shapes;
    }

    private static void WalkVml(XElement group, GroupFrame? parent, double x, double y, double width, double height, List<WordShape> shapes, PictureIdResolver pictureId, ref int textIndex)
    {
        Dictionary<string, string> groupStyle = VmlStyleOf(group);
        (double originX, double originY) = VmlPair((string?)group.Attribute("coordorigin"), (0, 0));
        (double sizeX, double sizeY) = VmlPair((string?)group.Attribute("coordsize"), (1000, 1000));
        if (sizeX == 0) sizeX = 1;
        if (sizeY == 0) sizeY = 1;
        double rotation = groupStyle.TryGetValue("rotation", out string? r) ? VmlRotation(r) : 0;
        string flip = groupStyle.TryGetValue("flip", out string? f) ? f : "";
        bool flipH = flip.Contains('x'), flipV = flip.Contains('y'), parentMirrors = parent?.Mirrors ?? false;
        var frame = new GroupFrame(x, y, width, height, originX, originY, width / sizeX, height / sizeY, rotation, flipH, flipV, parent,
                                   flipH ^ (parent?.PageFlipHorizontal ?? false), flipV ^ (parent?.PageFlipVertical ?? false),
                                   (parentMirrors ? -rotation : rotation) + (parent?.PageRotationDegrees ?? 0));
        foreach (XElement child in group.Elements())
        {
            Dictionary<string, string> style = VmlStyleOf(child);
            double left = VmlNumber(style, "left"), top = VmlNumber(style, "top"), childWidth = VmlNumber(style, "width"), childHeight = VmlNumber(style, "height");
            string kind = (string?)child.Attribute("vmlShape") ?? "";
            if (kind == "line")
            {
                (double x1, double y1) = VmlPair((string?)child.Attribute("from"), (0, 0));
                (double x2, double y2) = VmlPair((string?)child.Attribute("to"), (0, 0));
                left = Math.Min(x1, x2); top = Math.Min(y1, y2); childWidth = Math.Abs(x2 - x1); childHeight = Math.Abs(y2 - y1);
                if (x2 < x1) style["flip"] = (style.TryGetValue("flip", out string? lf) ? lf : "") + "x";
                if (y2 < y1) style["flip"] = (style.TryGetValue("flip", out string? lf2) ? lf2 : "") + "y";
            }
            else
            {
                if (childWidth == 0) childWidth = 1;
                if (childHeight == 0) childHeight = 1;
            }
            double boxX = frame.X + (left - frame.ChildOffsetX) * frame.ChildScaleX, boxY = frame.Y + (top - frame.ChildOffsetY) * frame.ChildScaleY;
            double boxWidth = childWidth * frame.ChildScaleX, boxHeight = childHeight * frame.ChildScaleY;
            if (child.Name.LocalName == "vgroup")
            {
                // a nested group keeps its box in this group's space; its own turn and flips apply to its children
                WalkVml(child, frame, boxX, boxY, boxWidth, boxHeight, shapes, pictureId, ref textIndex);
                continue;
            }
            (double cx, double cy) = frame.ToFrame(boxX + boxWidth / 2, boxY + boxHeight / 2);
            double ownRotation = style.TryGetValue("rotation", out string? sr) ? VmlRotation(sr) : 0;
            string ownFlip = style.TryGetValue("flip", out string? of) ? of : "";
            bool shapeFlipH = ownFlip.Contains('x') ^ frame.PageFlipHorizontal, shapeFlipV = ownFlip.Contains('y') ^ frame.PageFlipVertical;
            double finalRotation = (frame.Mirrors ? -ownRotation : ownRotation) + frame.PageRotationDegrees;
            WordShape shape;
            if (child.Attribute("embed") is { } embed && kind is "image" or "shape" or "rect")
                shape = new WordShape { Geometry = WordShapeGeometry.Picture, PictureId = pictureId(embed.Value) };
            else
            {
                shape = ReadVml(child, style, boxWidth, boxHeight);
                if (shape.TextIndex >= 0) shape.TextIndex = textIndex++;
            }
            shape.X = cx - boxWidth / 2; shape.Y = cy - boxHeight / 2; shape.Width = boxWidth; shape.Height = boxHeight;
            shape.RotationDegrees = finalRotation; shape.FlipHorizontal = shapeFlipH; shape.FlipVertical = shapeFlipV;
            shapes.Add(shape);
        }
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
