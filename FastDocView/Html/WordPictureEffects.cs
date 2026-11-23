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
/// - duotone: the luminance mixed linearly between the two colours in sRGB (C00000 / FFFF00: red 112 -> 198, 23, 0); each colour
///   is clipped to RGB before the mix, its satMod unclipped before that (DocxDrawingColor unclampedSaturation).
/// - colour change: only the exact colour (a PNG's 255, 0, 0 turned transparent; 252, 0, 0 kept).
/// - transparency (a:alphaModFix amt): the picture's opacity.
/// - soft edges (a:softEdge rad): a Gaussian edge (see SoftEdgeGradient); the two directions multiply (a blurred rectangle is
///   separable).
/// - 3-D rotation (a:scene3d camera rot lat / lon / rev): orthographic cameras as CSS rotations without perspective (rot lon 40
///   degrees: 180 pt shows 139 pt wide, cos 40 = 0.766); perspective cameras through CSS perspective at Word's camera distance
///   (see CameraHorizontalDistance). Word's light rig: a gain and a lift by the angle, drawn as a
///   linear transfer at the end of the filter (see LightGain).
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

    /// <summary>Whether the picture has a 3-D scene (a:scene3d): Word lights it even when it is not turned.</summary>
    public bool HasScene { get; init; }

    /// <summary>The camera's field of view in degrees for a perspective camera; 0 = orthographic (no perspective).</summary>
    public double CameraFieldOfView { get; init; }

    /// <summary>
    /// Word's camera distance for a 45 degree field of view, in points, fitted on its PDF of effects-3d-v1.docx (pictures of 90 x 60
    /// to 180 x 120 pt, square and tall, fov 45 and 90): it does not scale with the picture. Turning about the vertical axis
    /// (lon) foreshortens as a pinhole camera 645 / tan(fov / 2) pt away (1555 pt at 45 degrees, 655 at 90), turning about the
    /// horizontal axis (lat) as one 452 / tan(fov / 2) pt away (1095 at 45 degrees). Fit error 0.1-1 pt.
    /// </summary>
    private const double CameraHorizontalDistance = 645, CameraVerticalDistance = 452;

    public bool IsEmpty => ColorEffects.Count == 0 && Opacity >= 1 && SoftEdgePt <= 0 && Rotation == default && !HasScene;

    private const double RedLuminance = 0.2126, GreenLuminance = 0.7152, BlueLuminance = 0.0722;

    /// <summary>The SVG filter's id and markup (a zero-size svg holding it), or null when no colour operation is set.</summary>
    public (string Id, string Markup)? SvgFilter()
    {
        if (ColorEffects.Count == 0 && Rotation == default && !HasScene) return null;
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
        if (Rotation != default || HasScene)
        {
            double lift = LightLift(Rotation.Latitude, Rotation.Longitude);
            primitives.Append(Transfer(LightGain, lift / 255));
        }
        string body = primitives.ToString();
        string id = "fdv-picture-effect-" + Hash(body).ToString("x8", CultureInfo.InvariantCulture);
        return (id, "<svg class=\"word-picture-effect\" width=\"0\" height=\"0\" aria-hidden=\"true\" style=\"position:absolute;width:0;height:0;overflow:hidden\">"
                    + "<filter id=\"" + id + "\" color-interpolation-filters=\"sRGB\" x=\"0\" y=\"0\" width=\"1\" height=\"1\">" + body + "</filter></svg>");
    }

    /// <summary>
    /// Word's light rig on a turned flat picture (a:lightRig threePt from the top, the only one its UI writes for pictures):
    /// every channel comes out 1.043 x itself plus a lift that depends on the angle (its PDF of effects-light-v1, orthographic
    /// camera, error under 1 level; a perspective camera adds a few levels of its own across the picture). Measured lifts in
    /// levels, by Word's lat / lon in degrees (negative = 360 - angle); between them straight lines, beyond them the end value.
    /// Both angles turned: Word's lifts on a lat x lon grid (effects-light-two-v1, make_light_two.py, 2026-11-24), read bilinearly
    /// with the one-axis tables as its middle row and column; past 45 degrees the grid's edge. The lift is not the sum of the two
    /// axes (lat -45: lon -15 lifts 35, lon 0 64, lon 15 82).
    /// </summary>
    private const double LightGain = 1.043;

    private static readonly (double Angle, double Lift)[] LatitudeLift =
        [(-60, 44), (-45, 64), (-30, 53), (-15, 24), (-5, 10), (0, 5), (5, 3), (15, 2), (30, 8), (45, 21), (60, 29)];

    private static readonly (double Angle, double Lift)[] LongitudeLift =
        [(-60, 0), (-45, 0), (-30, 0), (-15, 1), (-5, 4), (0, 5), (5, 7), (15, 9), (30, 8), (45, 4), (60, 1)];

    private static double LightLift(double latitude, double longitude)
    {
        static double Signed(double degrees) { degrees %= 360; if (degrees > 180) degrees -= 360; if (degrees <= -180) degrees += 360; return degrees; }
        static double Lookup((double Angle, double Lift)[] table, double angle)
        {
            if (angle <= table[0].Angle) return table[0].Lift;
            for (int i = 1; i < table.Length; i++)
                if (angle <= table[i].Angle)
                    return table[i - 1].Lift + (table[i].Lift - table[i - 1].Lift) * (angle - table[i - 1].Angle) / (table[i].Angle - table[i - 1].Angle);
            return table[table.Length - 1].Lift;
        }
        double lat = Signed(latitude), lon = Signed(longitude);
        if (Math.Abs(lon) < 1e-9) return Lookup(LatitudeLift, lat);
        if (Math.Abs(lat) < 1e-9) return Lookup(LongitudeLift, lon);
        double GridAt(int row, int column)
            => row == 3 ? Lookup(LongitudeLift, GridAngles[column]) : column == 3 ? Lookup(LatitudeLift, GridAngles[row]) : TwoAxisLift[row, column];
        (int Index, double Fraction) Cell(double angle)
        {
            angle = Math.Max(GridAngles[0], Math.Min(GridAngles[GridAngles.Length - 1], angle));
            int i = 0;
            while (i < GridAngles.Length - 2 && angle > GridAngles[i + 1]) i++;
            return (i, (angle - GridAngles[i]) / (GridAngles[i + 1] - GridAngles[i]));
        }
        (int r, double fr) = Cell(lat);
        (int c, double fc) = Cell(lon);
        return (GridAt(r, c) * (1 - fc) + GridAt(r, c + 1) * fc) * (1 - fr) + (GridAt(r + 1, c) * (1 - fc) + GridAt(r + 1, c + 1) * fc) * fr;
    }

    private static readonly double[] GridAngles = [-45, -30, -15, 0, 15, 30, 45];

    /// <summary>Word's lifts with both axes turned (rows lat, columns lon, both -45 .. 45 by 15; row / column 0 from the one-axis tables).</summary>
    private static readonly double[,] TwoAxisLift =
    {
        { 3.4, 13.2, 35.4, 0, 82.4, 80.1, 56.3 },
        { 0.4, 6.0, 23.7, 0, 72.7, 69.6, 45.2 },
        { 0, 1.7, 9.0, 0, 35.6, 33.2, 20.3 },
        { 0, 0, 0, 0, 0, 0, 0 },
        { 1.3, 2.7, 1.6, 0, 0, 0, 0 },
        { 10.2, 13.3, 11.6, 0, 2.7, 0.4, 0 },
        { 26.2, 29.2, 27.5, 0, 12.3, 5.3, 1.4 },
    };

    /// <summary>The CSS for the picture's element: the filter, opacity, the soft edge's mask and the rotation ("" when none).</summary>
    public string Css(string? filterId, double widthPt = 0, double heightPt = 0)
    {
        var css = new StringBuilder();
        if (filterId is not null) css.Append("filter:url(#").Append(filterId).Append(");");
        if (Opacity < 1) css.Append("opacity:").Append(Number(Math.Max(0, Opacity))).Append(';');
        if (SoftEdgePt > 0 && widthPt > 0 && heightPt > 0)
            css.Append("mask-image:").Append(SoftEdgeGradient("to right", widthPt)).Append(',').Append(SoftEdgeGradient("to bottom", heightPt))
               .Append(";mask-composite:intersect;mask-repeat:no-repeat;");
        if (Rotation != default)
        {
            // Word's lat / lon turn the near edge toward the viewer the other way round from CSS rotateX / rotateY (its PDF: lat 30
            // draws the top edge wider, lon 30 the right edge taller). A perspective camera: CSS perspective at the horizontal
            // distance, the vertical axis stretched by their ratio around it so rotateX sees its own distance.
            string turn = "rotateX(" + Number(-Rotation.Latitude) + "deg) rotateY(" + Number(-Rotation.Longitude) + "deg)";
            css.Append("transform:");
            if (CameraFieldOfView > 0 && CameraFieldOfView < 180)
            {
                double tangent = Math.Tan(CameraFieldOfView * Math.PI / 360);
                double stretch = CameraHorizontalDistance / CameraVerticalDistance;
                css.Append("scaleY(").Append(Number(1 / stretch)).Append(") perspective(").Append(Number(CameraHorizontalDistance / tangent)).Append("pt) ")
                   .Append(turn).Append(" scaleY(").Append(Number(stretch)).Append(')');
            }
            else css.Append(turn);
            css.Append(" rotateZ(").Append(Number(Rotation.Revolution)).Append("deg);");
        }
        return css.ToString();
    }

    /// <summary>
    /// One direction of Word's soft edge across a picture side of <paramref name="lengthPt"/>: the picture's rectangle inset by
    /// <see cref="SoftEdgeCentre"/> radii from each edge and blurred with sigma 0.33 rad (a Gaussian edge from both sides:
    /// Phi((x - c) / sigma) - Phi((x - (length - c)) / sigma)). Word never shortens the radius: on a picture smaller than about
    /// four radii both edges overlap and the whole picture stays faint (its PDF of effects-light-v1: rad 20 pt on 40 x 40 pt
    /// peaks at alpha 9 / 255, on 60 x 40 at 30, rad 40 on 180 x 120 at 237). Stops every 0.1 rad from each edge (and the
    /// middle), as percentages of the side so the mask follows the picture when it is drawn smaller; never crossing.
    /// </summary>
    private string SoftEdgeGradient(string direction, double lengthPt)
    {
        const double Sigma = 0.33;
        double centre = SoftEdgeCentre(SoftEdgePt) * SoftEdgePt, sigma = Sigma * SoftEdgePt;
        var positions = new SortedSet<double> { 0, lengthPt / 2, lengthPt };
        for (int i = 1; i <= 20; i++)
        {
            double at = i * 0.1 * SoftEdgePt;
            if (at < lengthPt / 2) { positions.Add(at); positions.Add(lengthPt - at); }
        }
        var gradient = new StringBuilder("linear-gradient(").Append(direction);
        foreach (double at in positions)
        {
            double alpha = at <= 0 || at >= lengthPt ? 0
                : Math.Max(0, NormalDistribution((at - centre) / sigma) - NormalDistribution((at - (lengthPt - centre)) / sigma));
            gradient.Append(",rgba(0,0,0,").Append(Number(alpha)).Append(") ").Append(Number(at / lengthPt * 100)).Append('%');
        }
        return gradient.Append(')').ToString();
    }

    /// <summary>
    /// Where Word's soft edge is half transparent, in radii from the edge: 0.865 at rad 5 pt, 0.845 at 10, 0.927 at 20 (the
    /// mean of the four edges of its PDF masks, each within 0.05), straight lines between and the end values beyond. Not a
    /// constant: Word draws the blur on its own pixel grid.
    /// </summary>
    private static double SoftEdgeCentre(double radiusPt)
    {
        (double Radius, double Centre)[] measured = [(5, 0.865), (10, 0.845), (20, 0.927)];
        if (radiusPt <= measured[0].Radius) return measured[0].Centre;
        for (int i = 1; i < measured.Length; i++)
            if (radiusPt <= measured[i].Radius)
                return measured[i - 1].Centre + (measured[i].Centre - measured[i - 1].Centre) * (radiusPt - measured[i - 1].Radius) / (measured[i].Radius - measured[i - 1].Radius);
        return measured[measured.Length - 1].Centre;
    }

    /// <summary>The standard normal CDF (Abramowitz and Stegun 7.1.26 for erf, error below 1.5e-7).</summary>
    private static double NormalDistribution(double x)
    {
        double z = Math.Abs(x) / Math.Sqrt(2), t = 1 / (1 + 0.3275911 * z);
        double erf = 1 - ((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t * Math.Exp(-z * z);
        return x >= 0 ? (1 + erf) / 2 : (1 - erf) / 2;
    }

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

    private static string Number(double value) => Math.Round(value, 4) is var rounded && rounded == 0 ? "0" : rounded.ToString("0.####", CultureInfo.InvariantCulture);
}
