namespace FastDocView.Html;

/// <summary>
/// A bevel on a 3-D shape's front or back (DrawingML a:sp3d a:bevelT / a:bevelB, ECMA-376 20.1.5.1): its width (inset from the
/// outline) and height in points and its preset profile (ST_BevelPresetType, default circle). Drawn as bands along the outline, one
/// per profile segment (<see cref="FastDocView.Odt.OdtExtrusion"/>). ECMA-376 names the presets but does not define their profiles and
/// LibreOffice draws no bevel (master 2026-11-17: oox/source/drawingml/shape3dproperties.cxx:287-369 only passes bevelT / bevelB / contourW
/// on as grab-bag attributes; scene3dhelper.cxx never reads them), so the profiles are Word's,
/// measured from Word 16's print (test-data/word-shapes-3d/make-shapes-3d-bevels.ps1, 2026-11-17): angle is the straight 45 degree
/// line; every other preset is a table over 48 steps of the width - the slope of the surface (degrees, from the band colours of
/// shapes-3d-bevels-v1 under the threePt rig, inverted through Word's lighting) and its height (fraction of the bevel's height, from
/// the edge-on silhouette of shapes-3d-bevels-edge-v1, the slopes carrying it where the silhouette cannot: falling parts).
/// </summary>
internal sealed record WordShapeBevel(double Width, double Height, string Preset)
{
    /// <summary>(slope degrees for a bevel as wide as high, height at each step's inner end) per preset, 48 steps from the outline in.</summary>
    private static readonly Dictionary<string, (string Degrees, string Heights)> Measured = new()
    {
        ["relaxedInset"] = ("82 82 72 70 68 66 63 61 58 56 55 54 51 49 46 44 42 40 38 36 35 33 31 38 -48 -54 -54 -54 -53 -53 -52 -51 -50 -48 -45 -42 -39 -36 -34 -31 -28 -25 -14 -11 -7 -6 -2 -1", "0.274 0.406 0.447 0.538 0.579 0.609 0.665 0.680 0.726 0.756 0.767 0.812 0.828 0.858 0.868 0.883 0.914 0.929 0.944 0.944 0.959 0.970 0.985 1.000 0.977 0.949 0.921 0.893 0.865 0.837 0.810 0.784 0.760 0.736 0.716 0.696 0.679 0.664 0.650 0.638 0.627 0.617 0.612 0.607 0.605 0.603 0.602 0.602"),
        ["circle"] = ("81 81 69 67 65 63 60 58 56 55 54 54 50 48 46 45 43 42 40 38 37 35 34 33 32 31 30 28 27 27 27 28 28 28 13 12 11 11 10 9 8 6 6 5 3 2 1 0", "0.219 0.320 0.361 0.437 0.462 0.477 0.538 0.564 0.594 0.609 0.640 0.665 0.680 0.726 0.726 0.741 0.767 0.782 0.812 0.812 0.828 0.843 0.858 0.868 0.883 0.883 0.898 0.914 0.914 0.929 0.929 0.944 0.944 0.959 0.959 0.970 0.970 0.985 0.985 0.985 0.985 0.985 0.985 0.985 1.000 1.000 1.000 1.000"),
        ["slope"] = ("82 82 82 82 82 82 72 68 59 46 41 30 14 7 6 4 2 0 0 0 0 0 0 0 -1 -3 -5 -7 -8 -27 -33 -44 -57 -63 -70 -81 -82 -82 -82 -82 -82 -86 -58 -55 -35 -30 -8 -4", "0.244 0.391 0.462 0.609 0.680 0.756 0.868 0.898 0.944 0.959 0.970 0.985 0.985 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 0.999 0.998 0.997 0.994 0.991 0.980 0.967 0.946 0.914 0.874 0.815 0.697 0.578 0.460 0.342 0.224 0.106 0.000 0.000 0.000 0.000 0.000 0.000 0.000"),
        ["cross"] = ("37 37 28 4 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 9 28 38 59 68 56 41 31 7 1 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0", "0.056 0.086 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.102 0.117 0.117 0.132 0.132 0.914 0.959 0.970 0.985 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000"),
        ["softRound"] = ("85 85 84 80 70 65 41 -36 -45 -56 -58 -60 -60 -60 -59 -58 -56 -55 -54 -54 -53 -52 -49 -46 -44 -42 -40 -38 -36 -34 -32 -31 -29 -28 -27 -27 -26 -16 -14 -12 -10 -9 -6 -6 -4 -3 -2 0", "0.437 0.564 0.741 0.812 0.868 0.959 0.985 0.970 0.949 0.919 0.885 0.849 0.813 0.778 0.744 0.710 0.679 0.649 0.620 0.591 0.564 0.538 0.514 0.492 0.472 0.453 0.436 0.419 0.404 0.390 0.376 0.364 0.352 0.341 0.331 0.320 0.310 0.304 0.299 0.294 0.291 0.288 0.285 0.283 0.281 0.280 0.280 0.280"),
        ["convex"] = ("40 40 14 8 1 2 4 8 10 13 15 26 27 28 30 31 32 34 34 36 38 39 40 42 43 44 46 46 49 50 52 53 54 54 55 56 57 58 60 62 64 68 70 66 44 28 9 3", "0.071 0.086 0.102 0.102 0.117 0.117 0.117 0.117 0.117 0.132 0.132 0.132 0.142 0.158 0.173 0.173 0.188 0.203 0.203 0.219 0.234 0.244 0.274 0.274 0.289 0.320 0.335 0.361 0.376 0.391 0.422 0.447 0.477 0.508 0.523 0.564 0.579 0.609 0.665 0.695 0.756 0.797 0.843 0.959 0.985 1.000 1.000 1.000"),
        ["coolSlant"] = ("71 71 69 68 67 66 66 66 65 65 65 64 64 61 58 54 52 46 41 37 31 29 30 10 8 6 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0", "0.188 0.244 0.320 0.361 0.437 0.462 0.492 0.564 0.594 0.665 0.695 0.726 0.782 0.812 0.858 0.868 0.898 0.929 0.944 0.959 0.970 0.970 0.985 0.985 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000"),
        ["divot"] = ("82 82 70 70 68 66 64 63 60 57 55 54 50 46 44 41 40 37 34 32 30 28 25 26 14 7 -39 0 0 0 0 0 0 0 0 0 0 0 0 0 45 45 34 31 12 7 4 1", "0.203 0.259 0.376 0.437 0.508 0.538 0.564 0.640 0.655 0.695 0.726 0.741 0.767 0.782 0.812 0.828 0.843 0.858 0.868 0.883 0.898 0.898 0.914 0.914 0.914 0.929 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.913 0.944 0.970 0.970 0.985 1.000 1.000 1.000 1.000"),
        ["riblet"] = ("70 70 68 66 63 61 60 58 56 54 54 54 54 53 51 50 49 47 46 45 44 42 39 36 32 29 30 29 10 7 5 3 2 1 0 0 0 0 0 0 0 0 29 57 64 31 10 1", "0.188 0.274 0.320 0.376 0.406 0.422 0.477 0.492 0.553 0.564 0.579 0.609 0.640 0.665 0.680 0.711 0.741 0.756 0.767 0.797 0.812 0.843 0.843 0.858 0.868 0.883 0.898 0.898 0.898 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.914 0.944 0.985 1.000 1.000 1.000"),
        ["hardEdge"] = ("82 82 82 82 82 82 82 83 57 32 8 3 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 0 -2 -36 -64 -57 -46 -40 -34 -30 -30 -29 -29 -13 -12 -10 -8 -7 -4 -2 -1", "0.203 0.361 0.437 0.579 0.655 0.726 0.883 0.959 0.985 0.985 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 1.000 0.999 0.984 0.942 0.910 0.889 0.871 0.857 0.845 0.833 0.821 0.810 0.805 0.801 0.797 0.794 0.792 0.790 0.789 0.789"),
        ["artDeco"] = ("59 59 56 54 45 42 37 33 30 26 27 12 10 6 3 37 83 68 66 60 58 56 54 54 49 46 44 41 39 37 35 33 32 30 28 26 27 27 28 13 11 10 8 7 5 4 2 0", "0.132 0.173 0.203 0.234 0.244 0.259 0.289 0.289 0.305 0.320 0.320 0.335 0.335 0.335 0.335 0.361 0.538 0.579 0.609 0.665 0.695 0.726 0.756 0.767 0.797 0.812 0.843 0.858 0.868 0.883 0.898 0.914 0.929 0.929 0.944 0.959 0.959 0.970 0.970 0.985 0.985 0.985 0.985 0.985 1.000 1.000 1.000 1.000"),
    };

    private static readonly Dictionary<(string, int), IReadOnlyList<WordShapeBevelBand>> BandCache = new();

    /// <summary>The profile as at most <paramref name="count"/> bands from the outline in: adjacent steps whose slopes differ by 3 degrees
    /// or less merged into one, the tolerance widened by 3 degrees until the bands fit.</summary>
    public IReadOnlyList<WordShapeBevelBand> BandsAtMost(int count)
    {
        string key = Measured.ContainsKey(Preset) || Preset == "angle" ? Preset : "circle";
        lock (BandCache)
        {
            if (BandCache.TryGetValue((key, count), out var cached)) return cached;
            List<WordShapeBevelBand> bands;
            if (key == "angle") bands = [new(0, 1, 0, 1, 45)];
            else
            {
                int tolerance = 3;
                bands = Merge(Measured[key], tolerance);
                while (bands.Count > count && tolerance < 180) bands = Merge(Measured[key], tolerance += 3);
            }
            BandCache[(key, count)] = bands;
            return bands;
        }
    }

    /// <summary>The profile's slope per step (degrees for a bevel as wide as high), 48 steps from the outline in; angle: one step of 45.</summary>
    public IReadOnlyList<double> StepDegrees()
    {
        string key = Measured.ContainsKey(Preset) || Preset == "angle" ? Preset : "circle";
        return key == "angle" ? [45.0] : Parse(Measured[key].Degrees);
    }

    private static List<WordShapeBevelBand> Merge((string Degrees, string Heights) table, int tolerance)
    {
        double[] degrees = Parse(table.Degrees), heights = Parse(table.Heights);
        int steps = degrees.Length;
        var bands = new List<WordShapeBevelBand>();
        int first = 0;
        double start = 0;
        for (int k = 1; k <= steps; k++)
        {
            if (k < steps && Math.Abs(degrees[k] - degrees[first]) <= tolerance) continue;
            double sum = 0;
            for (int j = first; j < k; j++) sum += degrees[j];
            double end = heights[k - 1];
            bands.Add(new WordShapeBevelBand((double)first / steps, (double)k / steps, start, end, sum / (k - first)));
            start = end;
            first = k;
        }
        return bands;
    }

    private static double[] Parse(string values)
        => values.Split(' ').Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
}

/// <summary>One band of a bevel's profile: from / to as fractions of its width (0 = the outline), the height at each end as a fraction
/// of its height, and the surface's slope in degrees for a bevel as wide as high (Word's shading of the band; positive = rising inward).</summary>
internal readonly record struct WordShapeBevelBand(double From, double To, double Start, double End, double Degrees);
