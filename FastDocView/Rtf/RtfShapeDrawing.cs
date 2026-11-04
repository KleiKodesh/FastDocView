using FastDocView.Html;

namespace FastDocView.Rtf;

/// <summary>An RTF shape's look as the format-neutral <see cref="WordShape"/> model (property meanings: <see cref="RtfShapePlacement"/>).</summary>
internal static class RtfShapeDrawing
{
    private const double EmuPerPoint = 12700;

    /// <summary>A shape's text-box insets (dxTextLeft.. in EMU; defaults 91440 / 45720, LibreOffice rtfsdrimport.cxx:113-124).</summary>
    public static WordTextBoxStyle TextStyle(RtfShapePlacement shape)
        => new((shape.TextLeft ?? 91440) / EmuPerPoint, (shape.TextTop ?? 45720) / EmuPerPoint, (shape.TextRight ?? 91440) / EmuPerPoint,
               (shape.TextBottom ?? 45720) / EmuPerPoint, null, 0, false, null, 0, null, shape.FitsShapeToText);

    /// <summary>The fill, or null: fFilled defaults to on (rtfsdrimport.cxx:199) - except for the types OfficeArt does not fill by default (a line becomes
    /// a LineShape without fill, rtfsdrimport.cxx:329-330; the same table as .doc), fillColor to white ("White in Word", rtfsdrimport.cxx:360-363).</summary>
    public static string? FillColorCss(RtfShapePlacement shape)
        => (shape.Filled ?? WordOfficeArtShapeTypes.IsFilledByDefault(shape.ShapeType < 0 ? WordOfficeArtShapeTypes.Rectangle : shape.ShapeType)) ? Css(shape, shape.FillColor ?? 0xFFFFFF, WordOfficeArtColor.FillColor) : null;

    /// <summary>The line, or null: fLine 0 = no line (rtfsdrimport.cxx:183-190), lineColor defaults to black (rtfsdrimport.cxx:382).</summary>
    public static string? LineColorCss(RtfShapePlacement shape) => shape.HasLine == false ? null : Css(shape, shape.LineColor ?? 0, WordOfficeArtColor.LineColor);

    /// <summary>lineWidth in points; Word's default 0.75 pt (rtfsdrimport.cxx:384).</summary>
    public static double LineWidthPt(RtfShapePlacement shape) => shape.LineWidth is { } emu ? emu / EmuPerPoint : 0.75;

    /// <summary>A colour value: the same MSO_CLR numbers as .doc (24-bit, red in the low byte - rtfsdrimport.cxx:459 BGRToRGB - plus the
    /// scheme / system colour bits LibreOffice's RTF import drops; resolved by <see cref="WordOfficeArtColor"/> as for .doc).</summary>
    private static string Css(RtfShapePlacement shape, int value, int property)
        => WordOfficeArtColor.Css(unchecked((uint)value), property, other => other switch
        {
            WordOfficeArtColor.FillColor => shape.FillColor is { } fill ? unchecked((uint)fill) : null,
            WordOfficeArtColor.LineColor => shape.LineColor is { } line ? unchecked((uint)line) : null,
            WordOfficeArtColor.FillBackColor => shape.FillBackColor is { } back ? unchecked((uint)back) : null,
            _ => null,
        }, shape.HasLine != false);

    public delegate string? PictureIdOf(in RtfPicture picture);

    /// <summary>
    /// A text effect (WordArt) shape is drawn as its text only: LibreOffice renders a TextPath shape as Fontwork in place of its
    /// outline (svx EnhancedCustomShapeEngine.cxx:236-262); the text takes the shape's fill colour (Fontwork text is filled with it).
    /// </summary>
    private static void TextEffect(WordShape drawn, RtfShapePlacement shape)
    {
        if (shape.WordArtText is { } text) WordOfficeArtFontwork.Apply(drawn, text, shape.WordArtFont);
        else { drawn.Geometry = WordShapeGeometry.Path; drawn.Paths = []; }
    }

    /// <summary>The shape's fill as .doc reads it from the same properties (see <see cref="WordOfficeArtFill"/>).</summary>
    private static WordOfficeArtFillProperties FillOf(RtfShapePlacement shape, PictureIdOf pictureId)
    {
        var fill = new WordOfficeArtFillProperties
        {
            Type = shape.FillType,
            Color = Css(shape, shape.FillColor ?? 0xFFFFFF, WordOfficeArtColor.FillColor),
            BackColor = Css(shape, shape.FillBackColor ?? 0xFFFFFF, WordOfficeArtColor.FillBackColor),
            HasOpacity = shape.FillOpacity is not null,
            Opacity = (shape.FillOpacity ?? 0x10000) / 65536.0,
            BackOpacity = (shape.FillBackOpacity ?? 0x10000) / 65536.0,
            AngleFix16 = shape.FillAngle, Focus = shape.FillFocus, ToRight = shape.FillToRight, ToBottom = shape.FillToBottom,
            TileWidthPt = shape.FillWidth / EmuPerPoint, TileHeightPt = shape.FillHeight / EmuPerPoint,
        };
        if (shape.FillShadeColors is { } shades)
            fill.ShadeColors = shades.Select(s => (Css(shape, s.color, WordOfficeArtColor.FillColor), 1 - s.position / 65536.0)).ToList();
        if (shape.FillPicture is { } picture)
        {
            if (picture.Extension == ".dib" && fill.Type == 1 && picture.Data is { } dib && WordOfficeArtFill.DibPixels(dib) is { } pattern)
                (fill.PatternPixels, fill.PatternWidth, fill.PatternHeight) = pattern;
            else fill.PictureId = pictureId(picture);
        }
        return fill;
    }

    /// <summary>
    /// The drawn shapes of <paramref name="shape"/> at (x, y, width, height) points in the frame, children of a group placed by
    /// their relLeft.. in the group's groupLeft.. scaled to the group's box (rtfsdrimport.cxx:1036-1053); each text box's HTML
    /// is added to <paramref name="texts"/>, its index on the shape.
    /// </summary>
    public static void Collect(RtfShapePlacement shape, double x, double y, double width, double height, List<WordShape> shapes, List<string> texts, PictureIdOf pictureId)
    {
        if (shape.IsGroup)
        {
            int first = shapes.Count;
            foreach (RtfShapePlacement child in shape.Children)
            {
                if (shape.GroupLeft is { } groupLeft && shape.GroupTop is { } groupTop && shape.GroupRight is { } groupRight && shape.GroupBottom is { } groupBottom
                    && child.RelLeft is { } left && child.RelTop is { } top && child.RelRight is { } right && child.RelBottom is { } bottom
                    && groupRight != groupLeft && groupBottom != groupTop)
                {
                    double scaleX = width / (groupRight - groupLeft), scaleY = height / (groupBottom - groupTop);
                    Collect(child, x + (left - groupLeft) * scaleX, y + (top - groupTop) * scaleY, Math.Abs(right - left) * scaleX, Math.Abs(bottom - top) * scaleY, shapes, texts, pictureId);
                }
                else   // without the coordinates LibreOffice places the child by its own \shpleft.. (rtfsdrimport.cxx:1036 bInShapeGroup false)
                    Collect(child, x + (child.Left - shape.Left) / 20.0, y + (child.Top - shape.Top) / 20.0, Math.Abs(child.Right - child.Left) / 20.0,
                            Math.Abs(child.Bottom - child.Top) / 20.0, shapes, texts, pictureId);
            }
            WordShapeGroups.Turn(shapes, first, x, y, width, height, shape.RotationDegrees, shape.FlipHorizontal, shape.FlipVertical);
            return;
        }
        var drawn = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = shape.RotationDegrees, FlipHorizontal = shape.FlipHorizontal, FlipVertical = shape.FlipVertical,
            FillColor = FillColorCss(shape), LineColor = LineColorCss(shape), LineWidthPt = LineWidthPt(shape),
            Dash = shape.LineDashing is { } dashing ? WordOfficeArtShapeTypes.Dash(dashing) : null,
            FillOpacity = shape.FillOpacity is { } fillOpacity ? fillOpacity / 65536.0 : 1,
            LineOpacity = shape.LineOpacity is { } lineOpacity ? lineOpacity / 65536.0 : 1,
            TextStyle = TextStyle(shape),
        };
        WordOfficeArtShapeTypes.Apply(drawn, shape.ShapeType < 0 ? WordOfficeArtShapeTypes.Rectangle : shape.ShapeType, shape.Adjust);
        if (shape.Vertices is { Count: > 0 } vertices)
        {
            drawn.Geometry = WordShapeGeometry.Path;
            drawn.Paths = WordOfficeArtGeometry.Custom(vertices, shape.Segments ?? [], [], shape.Adjust, shape.GeoLeft, shape.GeoTop, shape.GeoRight, shape.GeoBottom, width, height);
        }
        if (drawn.FillColor is not null) WordOfficeArtFill.Apply(drawn, FillOf(shape, pictureId));
        if (shape.IsTextEffect) TextEffect(drawn, shape);
        WordOfficeArtShapeTypes.ApplyLineEnds(drawn, shape.StartArrowType, shape.StartArrowWidth, shape.StartArrowLength, shape.EndArrowType, shape.EndArrowWidth, shape.EndArrowLength);
        if (shape.Picture is { } picture)
        {
            drawn.Geometry = WordShapeGeometry.Picture;
            drawn.PictureId = pictureId(picture);
        }
        if (shape.HasText && !(shape.IsTextEffect && shape.WordArtText is not null)) { drawn.TextIndex = texts.Count; texts.Add(shape.TextHtml ?? ""); }
        shapes.Add(drawn);
    }
}
