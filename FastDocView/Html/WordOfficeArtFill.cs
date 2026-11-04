using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>An OfficeArt fill as read from a shape (.doc FOPT or RTF {\sp}): colours already resolved to CSS (<see cref="WordOfficeArtColor"/>).</summary>
public sealed class WordOfficeArtFillProperties
{
    /// <summary>fillType (MSO_FillType, msdffdef.hxx:661-675): 0 solid, 1 pattern, 2 texture, 3 picture, 4 shade, 5 shade centre, 6 shade shape, 7 shade scale, 8 shade title, 9 background.</summary>
    public int Type;
    public string Color = "#ffffff", BackColor = "#ffffff";
    public double Opacity = 1, BackOpacity = 1;
    public bool HasOpacity;
    /// <summary>fillAngle (16.16 degrees), fillFocus (percent), fillToRight / fillToBottom (16.16 fractions).</summary>
    public int AngleFix16, Focus, ToRight, ToBottom;
    /// <summary>fillShadeColors: (CSS colour, position 0..1) = (colour, 1 - dist / 65536), msdffimp.cxx:1136-1143; null = none.</summary>
    public List<(string Color, double Position)>? ShadeColors;
    /// <summary>A pattern blip decoded to its pixels (true = not black), with its size.</summary>
    public bool[]? PatternPixels;
    public int PatternWidth, PatternHeight;
    /// <summary>The fill blip's picture id (texture / picture fills), its tile size in points (0 = the picture's own size unknown: one tile = the shape).</summary>
    public string? PictureId;
    public double TileWidthPt, TileHeightPt;
}

/// <summary>
/// Draws an OfficeArt fill the way LibreOffice imports it (filter/source/msfilter/msdffimp.cxx, master read 2026-11-04):
/// ApplyFillAttributes :1313-1468 (solid; pattern / texture / picture as bitmaps; shade centre as a rendered bitmap; the other
/// shades as gradients; background = no fill), ImportGradientColor :2840-2943 (angle, focus and the colour swaps; linear,
/// axial for a focus of 40..60, rectangular for shade shape and shade centre; fill / back opacity as a transparency gradient),
/// ApplyRectangularGradientAsBitmap :1154-1279 (shade centre: distance to the edges toward the fillToRight / fillToBottom
/// focus, coloured from the shade colours, :1146-1151 default back colour at 0 and fill colour at 1) and the pattern bitmap
/// :1385-1418 (black pixels take the back colour, the others the fill colour). Gradient geometry as drawinglayer / basegfx draw
/// it: gradienttools.cxx init1DGradientInfo :52-120 (the box grown to the turned extent, turned counter-clockwise),
/// initRectGradientInfo :196-270 and getRectangularGradientAlpha :694-756 (aspect-corrected distance to the edges), and
/// texture.cxx :249-383 (an axial gradient has the start colour on its axis and the end colour at the edges).
/// Rectangular shades are drawn in 64 steps, as LibreOffice draws gradients in steps.
/// </summary>
public static class WordOfficeArtFill
{
    private const int Steps = 64;

    public static void Apply(WordShape shape, WordOfficeArtFillProperties fill)
    {
        double w = Math.Max(shape.Width, 0.01), h = Math.Max(shape.Height, 0.01);
        switch (fill.Type)
        {
            case 0:
                shape.FillColor = fill.Color;
                if (fill.HasOpacity) shape.FillOpacity = fill.Opacity;
                return;
            case 9:                                    // the slide background: none in a document
                shape.FillColor = null;
                return;
            case 1 when fill.PatternPixels is { } pixels:
                Pattern(shape, fill, pixels, w, h);
                return;
            case 1 or 2 or 3:
                if (fill.PictureId is null) { shape.FillColor = null; return; }
                shape.FillColor = null;
                // a texture tiles at fillWidth / fillHeight (msdffimp.cxx:1420-1426 XFillBmpTile, sizes / 360); a picture stretches (:1428-1431)
                shape.FillPictureId = fill.PictureId;
                bool tiled = fill.Type == 2 && fill.TileWidthPt > 0 && fill.TileHeightPt > 0;
                shape.FillPicture = new WordShapePictureFill(0, 0, 0, 0, tiled, 0, 0, tiled ? fill.TileWidthPt : 0, tiled ? fill.TileHeightPt : 0, null, fill.HasOpacity ? fill.Opacity : 1);
                return;
            case 5:
                ShadeCentre(shape, fill, w, h);
                return;
            default:
                Gradient(shape, fill, w, h);
                return;
        }
    }

    private static string Id(string prefix, string content)
    {
        uint hash = 2166136261;
        foreach (char c in content) { hash ^= c; hash *= 16777619; }
        return prefix + hash.ToString("x8", CultureInfo.InvariantCulture);
    }

    private static string N(double value) => WordShapeHtml.Number(value);

    private static (double r, double g, double b) Rgb(string css)
        => (int.Parse(css.Substring(1, 2), NumberStyles.HexNumber), int.Parse(css.Substring(3, 2), NumberStyles.HexNumber), int.Parse(css.Substring(5, 2), NumberStyles.HexNumber));

    private static string Mix(string a, string b, double t)
    {
        (double r1, double g1, double b1) = Rgb(a);
        (double r2, double g2, double b2) = Rgb(b);
        int R(double x, double y) => (int)Math.Max(0, Math.Min(255, Math.Round(x + (y - x) * t)));
        return "#" + R(r1, r2).ToString("x2") + R(g1, g2).ToString("x2") + R(b1, b2).ToString("x2");
    }

    private static void Gradient(WordShape shape, WordOfficeArtFillProperties fill, double w, double h)
    {
        // ImportGradientColor (msdffimp.cxx:2845-2920)
        int swaps = 0;
        if (fill.AngleFix16 >= 0) swaps ^= 1;
        double fix16Degrees = ((short)(fill.AngleFix16 >> 16) * 100L + (((fill.AngleFix16 & 0xFFFF) * 100L) >> 16)) / 100.0;   // Fix16ToAngle :393-402
        double rotation = fix16Degrees == 0 ? 0 : ((-fix16Degrees % 360) + 360) % 360;
        double angle = ((360 - rotation) % 360 + 360) % 360;            // counter-clockwise, 0..360 (in tenths in LibreOffice)
        angle = Math.Floor(angle * 10) / 10;
        int focus = fill.Focus;
        bool axial = false, rectangle = false;
        double focusX = focus, focusY = focus;
        if (focus == 0) swaps ^= 1;
        else if (focus < 0) { focus = -focus; swaps ^= 1; focusX = focusY = focus; }
        if (focus > 40 && focus < 60) { axial = true; swaps ^= 1; }
        if (fill.Type == 6) { rectangle = true; axial = false; focusX = focusY = 50; swaps ^= 1; }
        string col1 = fill.Color, col2 = fill.BackColor;
        double trans1 = fill.HasOpacity ? fill.Opacity : 1, trans2 = fill.BackOpacity;
        if (swaps != 0) { (col1, col2) = (col2, col1); (trans1, trans2) = (trans2, trans1); }
        // BColorStops(aCol2, aCol1): the start colour is col2
        double radians = angle * Math.PI / 180;
        double cos = Math.Cos(radians), sin = Math.Sin(radians);
        double expandedW = angle == 0 ? w : w * Math.Abs(cos) + h * Math.Abs(sin), expandedH = angle == 0 ? h : h * Math.Abs(cos) + w * Math.Abs(sin);
        if (!rectangle)
        {
            // the gradient runs top to bottom of the expanded box, turned counter-clockwise by the angle about the centre
            double dx = sin, dy = cos, cx = w / 2, cy = h / 2, half = expandedH / 2;
            var stops = new StringBuilder();
            void Stop(double offset, string color, double opacity)
            {
                stops.Append("<stop offset=\"").Append(N(offset)).Append("\" stop-color=\"").Append(color).Append('"');
                if (opacity < 1) stops.Append(" stop-opacity=\"").Append(N(opacity)).Append('"');
                stops.Append("/>");
            }
            if (axial) { Stop(0, col1, trans1); Stop(0.5, col2, trans2); Stop(1, col1, trans1); }
            else { Stop(0, col2, trans2); Stop(1, col1, trans1); }
            string body = "x1=\"" + N(cx - dx * half) + "\" y1=\"" + N(cy - dy * half) + "\" x2=\"" + N(cx + dx * half) + "\" y2=\"" + N(cy + dy * half) + "\">" + stops;
            string id = Id("word-officeart-gradient-", body);
            shape.FillDefinitions = "<linearGradient id=\"" + id + "\" gradientUnits=\"userSpaceOnUse\" " + body + "</linearGradient>";
            shape.FillPaint = "url(#" + id + ")";
            shape.FillColor = col2;
            return;
        }
        // rectangular: t = aspect-corrected distance to the nearest edge of the expanded box (shifted by the focus), 0 at the
        // edge (start colour) to 1 at the centre (end colour)
        var rects = new StringBuilder();
        double offsetX = (focusX / 100 - 0.5) * w, offsetY = (focusY / 100 - 0.5) * h;
        double left = (w - expandedW) / 2 + offsetX, top = (h - expandedH) / 2 + offsetY, smaller = Math.Min(expandedW, expandedH) / 2;
        rects.Append("<rect width=\"").Append(N(w)).Append("\" height=\"").Append(N(h)).Append("\" fill=\"").Append(col2).Append("\"/>");
        rects.Append("<g transform=\"rotate(").Append(N(-angle)).Append(' ').Append(N(w / 2 + offsetX)).Append(' ').Append(N(h / 2 + offsetY)).Append(")\">");
        for (int k = 0; k < Steps; k++)
        {
            double t = (double)k / Steps, inset = t * smaller;
            rects.Append("<rect x=\"").Append(N(left + inset)).Append("\" y=\"").Append(N(top + inset)).Append("\" width=\"").Append(N(Math.Max(0, expandedW - 2 * inset)))
                 .Append("\" height=\"").Append(N(Math.Max(0, expandedH - 2 * inset))).Append("\" fill=\"").Append(Mix(col2, col1, (double)k / (Steps - 1))).Append("\"/>");
        }
        rects.Append("</g>");
        string patternId = Id("word-officeart-shade-", rects.ToString());
        shape.FillDefinitions = "<pattern id=\"" + patternId + "\" patternUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"" + N(w) + "\" height=\"" + N(h) + "\">" + rects + "</pattern>";
        shape.FillPaint = "url(#" + patternId + ")";
        shape.FillColor = col2;
    }

    /// <summary>Shade centre as LibreOffice's bitmap (msdffimp.cxx:1154-1279): each level's rectangle shrinks toward the focus.</summary>
    private static void ShadeCentre(WordShape shape, WordOfficeArtFillProperties fill, double w, double h)
    {
        double fx = fill.ToRight / 65536.0, fy = fill.ToBottom / 65536.0;
        List<(string Color, double Position)> shades = fill.ShadeColors is { Count: > 0 } given ? given : [(fill.BackColor, 0), (fill.Color, 1)];
        string ColorAt(double distance)
        {
            double a = 0, b = 1;
            string colorA = shades[0].Color, colorB = colorA;
            foreach ((string color, double position) in shades)
            {
                if (a <= position && position <= distance) { a = position; colorA = color; }
                if (distance < position && position <= b) { b = position; colorB = color; }
            }
            return b - a != 0 ? Mix(colorA, colorB, (distance - a) / (b - a)) : colorA;
        }
        var rects = new StringBuilder();
        for (int k = 0; k <= Steps; k++)
        {
            double level = (double)k / Steps;
            double x0 = level * fx * w, y0 = level * fy * h, x1 = w - level * (1 - fx) * w, y1 = h - level * (1 - fy) * h;
            rects.Append("<rect x=\"").Append(N(x0)).Append("\" y=\"").Append(N(y0)).Append("\" width=\"").Append(N(Math.Max(0, x1 - x0)))
                 .Append("\" height=\"").Append(N(Math.Max(0, y1 - y0))).Append("\" fill=\"").Append(ColorAt(level)).Append("\"/>");
        }
        string id = Id("word-officeart-shade-", rects.ToString());
        shape.FillDefinitions = "<pattern id=\"" + id + "\" patternUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"" + N(w) + "\" height=\"" + N(h) + "\">" + rects + "</pattern>";
        shape.FillPaint = "url(#" + id + ")";
        shape.FillColor = ColorAt(0);
    }

    /// <summary>A pattern: the blip's pixels, black = back colour, else fill colour (msdffimp.cxx:1394-1416), tiled one pixel = 0.75 pt (96 dpi).</summary>
    private static void Pattern(WordShape shape, WordOfficeArtFillProperties fill, bool[] pixels, double w, double h)
    {
        const double PixelPt = 0.75;
        var cells = new StringBuilder();
        cells.Append("<rect width=\"").Append(N(fill.PatternWidth * PixelPt)).Append("\" height=\"").Append(N(fill.PatternHeight * PixelPt)).Append("\" fill=\"").Append(fill.BackColor).Append("\"/>");
        for (int y = 0; y < fill.PatternHeight; y++)
            for (int x = 0; x < fill.PatternWidth; x++)
                if (pixels[y * fill.PatternWidth + x])
                    cells.Append("<rect x=\"").Append(N(x * PixelPt)).Append("\" y=\"").Append(N(y * PixelPt)).Append("\" width=\"0.75\" height=\"0.75\" fill=\"").Append(fill.Color).Append("\"/>");
        string id = Id("word-officeart-pattern-", cells.ToString());
        shape.FillDefinitions = "<pattern id=\"" + id + "\" patternUnits=\"userSpaceOnUse\" width=\"" + N(fill.PatternWidth * PixelPt) + "\" height=\"" + N(fill.PatternHeight * PixelPt) + "\">" + cells + "</pattern>";
        shape.FillPaint = "url(#" + id + ")";
        shape.FillColor = fill.Color;
    }

    /// <summary>
    /// The pixels of a device-independent bitmap (BITMAPINFOHEADER, palette, rows padded to 4 bytes, bottom-up for a positive
    /// height) as "not black" flags - the pattern blips Word stores (8 x 8, usually 1 bit).
    /// </summary>
    public static (bool[] pixels, int width, int height)? DibPixels(byte[] dib)
    {
        if (dib.Length < 40) return null;
        int headerSize = BitConverter.ToInt32(dib, 0), width = BitConverter.ToInt32(dib, 4), height = BitConverter.ToInt32(dib, 8);
        int bits = BitConverter.ToUInt16(dib, 14), compression = BitConverter.ToInt32(dib, 16), used = BitConverter.ToInt32(dib, 32);
        if (headerSize < 40 || width <= 0 || width > 64 || height == 0 || Math.Abs(height) > 64 || compression != 0) return null;
        bool bottomUp = height > 0;
        height = Math.Abs(height);
        int paletteCount = bits <= 8 ? (used > 0 ? used : 1 << bits) : 0;
        int pixelsAt = headerSize + paletteCount * 4, stride = (width * bits + 31) / 32 * 4;
        if (pixelsAt + stride * height > dib.Length) return null;
        var result = new bool[width * height];
        for (int row = 0; row < height; row++)
        {
            int y = bottomUp ? height - 1 - row : row, at = pixelsAt + row * stride;
            for (int x = 0; x < width; x++)
            {
                int r, g, b;
                if (bits <= 8)
                {
                    int bitOffset = x * bits, value = (dib[at + bitOffset / 8] >> (8 - bits - bitOffset % 8)) & ((1 << bits) - 1);
                    int entry = headerSize + Math.Min(value, paletteCount - 1) * 4;
                    (b, g, r) = (dib[entry], dib[entry + 1], dib[entry + 2]);
                }
                else if (bits is 24 or 32) { int p = at + x * (bits / 8); (b, g, r) = (dib[p], dib[p + 1], dib[p + 2]); }
                else return null;
                result[y * width + x] = r != 0 || g != 0 || b != 0;
            }
        }
        return (result, width, height);
    }
}
