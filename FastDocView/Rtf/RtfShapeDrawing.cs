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
        => (shape.Filled ?? WordOfficeArtShapeTypes.IsFilledByDefault(shape.ShapeType < 0 ? WordOfficeArtShapeTypes.Rectangle : shape.ShapeType)) ? Css(shape.FillColor ?? 0xFFFFFF) : null;

    /// <summary>The line, or null: fLine 0 = no line (rtfsdrimport.cxx:183-190), lineColor defaults to black (rtfsdrimport.cxx:382).</summary>
    public static string? LineColorCss(RtfShapePlacement shape) => shape.HasLine == false ? null : Css(shape.LineColor ?? 0);

    /// <summary>lineWidth in points; Word's default 0.75 pt (rtfsdrimport.cxx:384).</summary>
    public static double LineWidthPt(RtfShapePlacement shape) => shape.LineWidth is { } emu ? emu / EmuPerPoint : 0.75;

    /// <summary>A colour value: 24-bit BGR (rtfsdrimport.cxx:459 BGRToRGB).</summary>
    private static string Css(int bgr) => "#" + (bgr & 0xFF).ToString("x2") + ((bgr >> 8) & 0xFF).ToString("x2") + ((bgr >> 16) & 0xFF).ToString("x2");

    public delegate string? PictureIdOf(in RtfPicture picture);

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
        WordOfficeArtShapeTypes.ApplyLineEnds(drawn, shape.StartArrowType, shape.StartArrowWidth, shape.StartArrowLength, shape.EndArrowType, shape.EndArrowWidth, shape.EndArrowLength);
        if (shape.Picture is { } picture)
        {
            drawn.Geometry = WordShapeGeometry.Picture;
            drawn.PictureId = pictureId(picture);
        }
        if (shape.HasText) { drawn.TextIndex = texts.Count; texts.Add(shape.TextHtml ?? ""); }
        shapes.Add(drawn);
    }
}
