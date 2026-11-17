using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>One colour operation on a picture's pixels, applied in document order (see <see cref="WordPictureEffects"/>).</summary>
public abstract record WordPictureColorEffect
{
    private WordPictureColorEffect() { }

    /// <summary>Grey by luminance (DrawingML a:grayscl).</summary>
    public sealed record Grayscale : WordPictureColorEffect;

    /// <summary>Black where the luminance is below the threshold, white from it up (a:biLevel thresh, 0..1).</summary>
    public sealed record BiLevel(double Threshold) : WordPictureColorEffect;

    /// <summary>Brightness and contrast (a:lum bright / contrast, -1..1).</summary>
    public sealed record Luminance(double Brightness, double Contrast) : WordPictureColorEffect;

    /// <summary>Luminance mapped from the first colour (black) to the second (white) (a:duotone; colours as 0..255 RGB).</summary>
    public sealed record Duotone((int R, int G, int B) Dark, (int R, int G, int B) Light) : WordPictureColorEffect;

    /// <summary>Every pixel of exactly one colour drawn in another (a:clrChange; the new colour's opacity 0..1).</summary>
    public sealed record ColorChange((int R, int G, int B) From, (int R, int G, int B) To, double ToOpacity) : WordPictureColorEffect;
}

/// <summary>
/// A picture's effects as Word 16 prints them (measured on its PDF of test-data/word-picture-effects/effects-xml-v1.docx,
/// 2026-11-17; every pixel value below is from that PDF's images):
/// - grey: Rec. 709 luminance of the sRGB values (0.2126 R + 0.7152 G + 0.0722 B: red 112 -> 23, green 136 -> 97, blue 174 -> 12).
/// - black and white: that luminance against the threshold (thresh 25 / 50 / 75 %: luminance 97 white at 25 %, black at 50 %).
/// - brightness b / contrast c: half the brightness before the contrast and half after (LibreOffice's note on Word's order,
///   oox/source/drawingml/fillproperties.cxx:834-836, matched on Word's pixels: washout b 70 % c -70 % takes 0 -> 205,
///   112 -> 238; b -40 % c 60 % takes 174 -> 65, 249 -> 252): v' = (v + b/2 - 1/2) k + 1/2 + b/2, k = 1 + c for c below 0,
///   1 / (1 - c) from 0 up.
/// - duotone: the luminance mixed linearly between the two colours in sRGB (C00000 / FFFF00: red 112 -> 198, 23, 0).
/// - colour change: only the exact colour (a PNG's 255, 0, 0 turned transparent; 252, 0, 0 kept).
/// - transparency (a:alphaModFix amt): the picture's opacity.
/// - soft edges (a:softEdge rad): the alpha falls off over 1.73 rad from each edge, half way at about 0.9 rad (Word's mask of
///   rad 10 and 20 pt); the two directions multiply (a blurred rectangle is separable).
/// - 3-D rotation (a:scene3d camera rot lat / lon / rev): drawn with an orthographic camera, as a CSS transform without
///   perspective (rot lon 40 degrees: 180 pt shows 139 pt wide, cos 40 = 0.766). The lighting Word adds is not drawn.
/// The colour operations become one SVG filter in sRGB (filter primitives: Filter Effects 1, feColorMatrix /
/// feComponentTransfer / feComposite / feFlood), named by a hash of its contents so equal effects share one id.
/// </summary>
public sealed class WordPictureEffects
{
    public IReadOnlyList<WordPictureColorEffect> ColorEffects { get; init; } = [];

    /// <summary>The picture's opacity (1 = opaque).</summary>
    public double Opacity { get; init; } = 1;

    /// <summary>The soft edge's radius in points (0 = none).</summary>
    public double SoftEdgePt { get; init; }

    /// <summary>The camera rotation in degrees: latitude (about the horizontal axis), longitude (vertical axis), revolution (in the plane).</summary>
    public (double Latitude, double Longitude, double Revolution) Rotation { get; init; }

    public bool IsEmpty => ColorEffects.Count == 0 && Opacity >= 1 && SoftEdgePt <= 0 && Rotation == default;

    private const double RedLuminance = 0.2126, GreenLuminance = 0.7152, BlueLuminance = 0.0722;

    /// <summary>The SVG filter's id and markup (a zero-size svg holding it), or null when no colour operation is set.</summary>
    public (string Id, string Markup)? SvgFilter()
    {
        if (ColorEffects.Count == 0) return null;
        var primitives = new StringBuilder();
        foreach (WordPictureColorEffect effect in ColorEffects)
            switch (effect)
            {
                case WordPictureColorEffect.Grayscale:
                    primitives.Append(LuminanceMatrix(1, 0));
                    break;
                case WordPictureColorEffect.BiLevel biLevel:
                {
                    // luminance, then a step at the threshold: values from it up go to 1, below it to 0 (8-bit steps of 1/255)
                    const double Steep = 2550;
                    double threshold = Math.Max(0, Math.Min(1, biLevel.Threshold));
                    primitives.Append(LuminanceMatrix(1, 0)).Append(Transfer(Steep, -Steep * (threshold - 0.5 / 255)));
                    break;
                }
                case WordPictureColorEffect.Luminance luminance:
                {
                    double brightness = Math.Max(-1, Math.Min(1, luminance.Brightness)), contrast = Math.Max(-1, Math.Min(1, luminance.Contrast));
                    double slope = contrast < 0 ? 1 + contrast : 1 / Math.Max(1 - contrast, 1.0 / 2550);
                    primitives.Append(Transfer(slope, (brightness / 2 - 0.5) * slope + 0.5 + brightness / 2));
                    break;
                }
                case WordPictureColorEffect.Duotone duotone:
                {
                    var matrix = new StringBuilder();
                    foreach ((int dark, int light) in new[] { (duotone.Dark.R, duotone.Light.R), (duotone.Dark.G, duotone.Light.G), (duotone.Dark.B, duotone.Light.B) })
                    {
                        double range = (light - dark) / 255.0;
                        matrix.Append(Number(range * RedLuminance)).Append(' ').Append(Number(range * GreenLuminance)).Append(' ')
                              .Append(Number(range * BlueLuminance)).Append(" 0 ").Append(Number(dark / 255.0)).Append(' ');
                    }
                    primitives.Append("<feColorMatrix type=\"matrix\" values=\"").Append(matrix).Append("0 0 0 1 0\"/>");
                    break;
                }
                case WordPictureColorEffect.ColorChange change:
                    primitives.Append(ColorChangePrimitives(change, primitives.Length));
                    break;
            }
        string body = primitives.ToString();
        string id = "fdv-picture-effect-" + Hash(body).ToString("x8", CultureInfo.InvariantCulture);
        return (id, "<svg class=\"word-picture-effect\" width=\"0\" height=\"0\" aria-hidden=\"true\" style=\"position:absolute;width:0;height:0;overflow:hidden\">"
                    + "<filter id=\"" + id + "\" color-interpolation-filters=\"sRGB\" x=\"0\" y=\"0\" width=\"1\" height=\"1\">" + body + "</filter></svg>");
    }

    /// <summary>The CSS for the picture's element: the filter, opacity, the soft edge's mask and the rotation ("" when none).</summary>
    public string Css(string? filterId)
    {
        var css = new StringBuilder();
        if (filterId is not null) css.Append("filter:url(#").Append(filterId).Append(");");
        if (Opacity < 1) css.Append("opacity:").Append(Number(Math.Max(0, Opacity))).Append(';');
        if (SoftEdgePt > 0)
        {
            string Edge(string direction)
            {
                var gradient = new StringBuilder("linear-gradient(").Append(direction);
                foreach ((double at, double alpha) in SoftEdgeProfile)
                    gradient.Append(",rgba(0,0,0,").Append(Number(alpha)).Append(") ").Append(Number(at * SoftEdgePt)).Append("pt");
                for (int i = SoftEdgeProfile.Length - 1; i >= 0; i--)
                    gradient.Append(",rgba(0,0,0,").Append(Number(SoftEdgeProfile[i].Alpha)).Append(") calc(100% - ").Append(Number(SoftEdgeProfile[i].At * SoftEdgePt)).Append("pt)");
                return gradient.Append(')').ToString();
            }
            css.Append("mask-image:").Append(Edge("to right")).Append(',').Append(Edge("to bottom")).Append(";mask-composite:intersect;mask-repeat:no-repeat;");
        }
        if (Rotation != default)
            css.Append("transform:rotateX(").Append(Number(Rotation.Latitude)).Append("deg) rotateY(").Append(Number(Rotation.Longitude))
               .Append("deg) rotateZ(").Append(Number(Rotation.Revolution)).Append("deg);");
        return css.ToString();
    }

    /// <summary>Word's soft edge alpha by distance from the edge in radii (its PDF masks for rad 10 and 20 pt, averaged).</summary>
    private static readonly (double At, double Alpha)[] SoftEdgeProfile =
        [(0, 0), (0.29, 0.035), (0.58, 0.18), (0.86, 0.47), (1.15, 0.78), (1.44, 0.95), (1.73, 1)];

    private static string LuminanceMatrix(double scale, double offset)
    {
        string row = Number(RedLuminance * scale) + " " + Number(GreenLuminance * scale) + " " + Number(BlueLuminance * scale) + " 0 " + Number(offset) + " ";
        return "<feColorMatrix type=\"matrix\" values=\"" + row + row + row + "0 0 0 1 0\"/>";
    }

    private static string Transfer(double slope, double intercept)
    {
        string function = "type=\"linear\" slope=\"" + Number(slope) + "\" intercept=\"" + Number(intercept) + "\"";
        return "<feComponentTransfer><feFuncR " + function + "/><feFuncG " + function + "/><feFuncB " + function + "/></feComponentTransfer>";
    }

    /// <summary>
    /// The exact colour drawn in another: per channel |v - from| scaled by 255 (the two signs from two matrices, added), summed
    /// into a "differs" alpha; the picture kept where it differs and the new colour (at its opacity) where it matches.
    /// </summary>
    private static string ColorChangePrimitives(WordPictureColorEffect.ColorChange change, int salt)
    {
        string name = "c" + salt.ToString(CultureInfo.InvariantCulture);
        string Difference(double sign)
        {
            var matrix = new StringBuilder();
            foreach ((int channel, int index) in new[] { (change.From.R, 0), (change.From.G, 1), (change.From.B, 2) })
            {
                for (int column = 0; column < 4; column++) matrix.Append(column == index ? Number(sign * 255) : "0").Append(' ');
                matrix.Append(Number(-sign * channel)).Append(' ');
            }
            return matrix.Append("0 0 0 0 1").ToString();
        }
        string source = name + "s", mask = name + "m";
        return "<feOffset result=\"" + source + "\"/>"         // the previous step's result (or the picture), named
             + "<feColorMatrix in=\"" + source + "\" type=\"matrix\" values=\"" + Difference(1) + "\" result=\"" + name + "a\"/>"
             + "<feColorMatrix in=\"" + source + "\" type=\"matrix\" values=\"" + Difference(-1) + "\" result=\"" + name + "b\"/>"
             + "<feComposite in=\"" + name + "a\" in2=\"" + name + "b\" operator=\"arithmetic\" k1=\"0\" k2=\"1\" k3=\"1\" k4=\"0\" result=\"" + name + "d\"/>"
             + "<feColorMatrix in=\"" + name + "d\" type=\"matrix\" values=\"0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 1 1 1 0 0\" result=\"" + mask + "\"/>"
             + "<feComposite in=\"" + source + "\" in2=\"" + mask + "\" operator=\"in\" result=\"" + name + "k\"/>"
             + "<feFlood flood-color=\"rgb(" + change.To.R + "," + change.To.G + "," + change.To.B + ")\" flood-opacity=\"" + Number(change.ToOpacity) + "\" result=\"" + name + "f\"/>"
             + "<feComposite in=\"" + name + "f\" in2=\"" + mask + "\" operator=\"out\" result=\"" + name + "n\"/>"
             + "<feComposite in=\"" + name + "n\" in2=\"" + source + "\" operator=\"in\" result=\"" + name + "t\"/>"
             + "<feMerge><feMergeNode in=\"" + name + "t\"/><feMergeNode in=\"" + name + "k\"/></feMerge>";
    }

    private static uint Hash(string text)
    {
        uint hash = 2166136261;
        foreach (char c in text) { hash ^= c; hash *= 16777619; }
        return hash;
    }

    private static string Number(double value) => Math.Round(value, 4).ToString("0.####", CultureInfo.InvariantCulture);
}
