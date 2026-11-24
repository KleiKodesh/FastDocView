using System.Globalization;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// 3-D shapes against Word 16's PDFs in test-data/word-shapes-3d (make-shapes-3d-fixture.ps1; measured 2026-11-17 at 4 px per pt):
/// shapes-3d-v1.docx (a:scene3d / a:sp3d: legacy oblique and perspective cameras, isometric, perspective and orthographic turns,
/// an extrusion colour, circle / slope / angle bevels, metal, an oblique camera under the balanced rig, a contour) and
/// shapes-3d-on-v1.rtf (Word's RTF save of it with f3D turned on, make_rtf_3d.py: the legacy OfficeArt 3-D properties, with
/// [MS-ODRAW]'s defaults - parallel, key light harsh, fill light soft - and Word's legacy lighting).
/// Each shape's drawn extent against Word's (page points) and the colour at the box's centre against Word's pixel there.
/// </summary>
public static class Shapes3DFidelityTest
{
    private static string Folder => Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-shapes-3d");

    private static readonly (double X, double Y)[] Boxes =
        [(60, 80), (250, 80), (430, 80), (60, 240), (250, 240), (430, 240), (60, 400), (250, 400), (430, 400), (60, 560), (250, 560), (430, 560)];

    // Word's extent of each .docx shape (left, top, right, bottom) and its centre pixel; tolerance in pt / levels (-1: not checked)
    private static readonly (double L, double T, double R, double B, double Outline, string Centre, int Levels)[] Docx =
    [
        (59.8, 66.8, 193.0, 160.2, 1, "4274c7", 6),     // legacyObliqueTopRight, legacy matte
        (249.5, 79.5, 370.2, 160.2, 1, "4779cd", 6),    // legacyPerspectiveFront
        (429.5, 74.8, 550.2, 160.2, 1, "4375c8", 8),    // legacyPerspectiveTop
        (77.2, 214.2, 176.8, 337.2, 1, "4274c7", 8),    // isometricLeftDown
        (249.5, 234.8, 370.2, 314.2, 1, "4a7ccf", 8),   // perspectiveFront turned 30 degrees, an oval
        (422.5, 216.5, 544.8, 327.5, 1, "5587da", 4),   // orthographic, lat 35 lon -25, green extrusion (its side 124 grey as Word's)
        (59.8, 399.5, 180.0, 480.0, 1, "4779cc", 6),    // circle bevel only
        (249.5, 399.5, 370.2, 480.0, 1, "4979cd", 6),   // slope bevel on an oval
        (429.5, 399.5, 550.2, 480.0, 1, "315796", 10),  // angle bevel, legacy metal
        (53.2, 559.0, 174.0, 640.8, 1, "4375c8", 6),    // perspective lon -20, rounded rectangle with a bevel (the bevel's top on the shape's plane)
        (249.5, 552.2, 370.2, 640.2, 1, "5583ce", 4),   // oblique top, balanced rig (half its ambient as white)
        (428.5, 549.0, 551.2, 626.2, 1, "4676c3", 4),   // turned 50 degrees, flat material, 2 pt contour round it
    ];

    // the legacy OfficeArt scene (RTF, f3D on): extents of the shapes whose 3-D style Word wrote as a legacy one (others: -1)
    private static readonly (double L, double T, double R, double B, string Centre)[] Rtf =
    [
        (59.5, 67.0, 192.8, 160.2, "305592"), (249.5, 79.5, 370.2, 160.2, "305592"), (429.5, 74.8, 550.2, 160.2, "305592"), (59.5, 239.5, 180.2, 320.2, "305592"),
        (249.5, 228.0, 370.2, 314.0, "2e528f"), (429.5, 227.0, 562.8, 320.2, "305592"), (59.5, 387.0, 192.8, 480.2, "305592"), (249.5, 387.0, 382.8, 480.2, "305592"),
        (429.5, 387.0, 562.8, 480.2, "305592"), (52.5, 558.5, 175.0, 641.2, "365fa3"), (249.5, 547.2, 382.8, 640.2, "305592"), (429.5, 547.0, 562.8, 640.2, "305592"),
    ];

    public static void Run()
    {
        TestReport.Step("3-D shapes: .docx scene3d / sp3d and RTF legacy 3-D against Word's PDF");
        string docx = Path.Combine(Folder, "shapes-3d-v1.docx"), rtf = Path.Combine(Folder, "shapes-3d-on-v1.rtf"), doc = Path.Combine(Folder, "shapes-3d-v1.doc");
        if (!File.Exists(docx) || !File.Exists(rtf) || !File.Exists(doc)) { TestReport.Check(false, "test-data/word-shapes-3d fixtures present"); return; }
        var shapes = Shapes(docx);
        TestReport.Check(shapes.Count == 12, $".docx: 12 shapes ({shapes.Count})");
        if (shapes.Count != 12) return;
        for (int i = 0; i < 12; i++)
        {
            var (faces, box) = (shapes[i], Boxes[i]);
            var e = Docx[i];
            (double l, double t, double r, double b) = Extent(faces, box);
            double off = new[] { Math.Abs(l - e.L), Math.Abs(t - e.T), Math.Abs(r - e.R), Math.Abs(b - e.B) }.Max();
            TestReport.Check(off <= e.Outline, $".docx shape {i}: extent within {e.Outline} pt of Word's ({off:0.00} pt: {l:0.0} {t:0.0} {r:0.0} {b:0.0})");
            string? centre = ColourAt(faces, 60, 40);
            int levels = centre is null ? 999 : Levels(centre, e.Centre);
            TestReport.Check(levels <= e.Levels, $".docx shape {i}: centre {centre} against Word's {e.Centre} ({levels} levels, allowed {e.Levels})");
        }
        TestReport.Check(shapes[6].Count > 20 && shapes[8].Count >= 5, $"bevels drawn as bands (circle {shapes[6].Count} faces, angle {shapes[8].Count})");

        var legacy = Shapes(rtf);
        TestReport.Check(legacy.Count == 12, $"RTF with f3D: 12 shapes ({legacy.Count})");
        if (legacy.Count == 12)
            for (int i = 0; i < Rtf.Length; i++)
            {
                (double l, double t, double r, double b) = Extent(legacy[i], Boxes[i]);
                double off = new[] { Math.Abs(l - Rtf[i].L), Math.Abs(t - Rtf[i].T), Math.Abs(r - Rtf[i].R), Math.Abs(b - Rtf[i].B) }.Max();
                TestReport.Check(off <= 1, $"RTF legacy 3-D shape {i}: extent within 1 pt of Word's ({off:0.00} pt)");
                string? centre = ColourAt(legacy[i], 60, 40);
                int levels = centre is null ? 999 : Levels(centre, Rtf[i].Centre);
                TestReport.Check(levels <= 3, $"RTF legacy 3-D shape {i}: centre {centre} against Word's {Rtf[i].Centre} ({levels} levels)");
            }
        // Word's .doc / RTF save keeps the 3-D style with f3D off, and Word prints those flat: so do we
        var flat = Shapes(doc);
        TestReport.Check(flat.Count == 12 && flat.All(f => f.Count <= 2), ".doc with f3D off: flat shapes, as Word prints them");
        BevelPresets();
        OvalBevels();
        Parts();
    }

    /// <summary>
    /// Outlines with holes or several parts (shapes-3d-parts-v1.docx: donut, frame, no-symbol, divide sign with a 3 pt contour; bevels on a
    /// donut, a frame, a divide sign; a donut with both): drawn extent against Word's print (4 px per pt), the contour round every
    /// outline and hole, the bevel along every ring.
    /// </summary>
    private static void Parts()
    {
        string docx = Path.Combine(Folder, "shapes-3d-parts-v1.docx");
        if (!File.Exists(docx)) { TestReport.Check(false, "shapes-3d-parts-v1 fixture present"); return; }
        var shapes = Shapes(docx);
        TestReport.Check(shapes.Count == 8, $"parts: 8 shapes ({shapes.Count})");
        if (shapes.Count != 8) return;
        (double X, double Y)[] boxes = [(50, 70), (230, 70), (410, 70), (50, 240), (230, 240), (410, 240), (50, 410), (230, 410)];
        (double L, double T, double R, double B)[] word =
        [
            (48.2, 67.0, 202.0, 168.0), (228.2, 67.0, 382.0, 168.0), (407.5, 68.0, 551.0, 182.0), (67.8, 247.0, 182.0, 328.2),
            (229.5, 239.5, 380.5, 350.5), (409.5, 239.5, 560.5, 350.5), (69.5, 422.5, 180.5, 507.5), (228.2, 408.2, 382.0, 521.8),
        ];
        string[] names = ["donut contour", "frame contour", "no-symbol contour", "divide contour", "donut bevel", "frame bevel", "divide bevel", "donut bevel + contour"];
        for (int i = 0; i < 8; i++)
        {
            (double l, double t, double r, double b) = Extent(shapes[i], boxes[i]);
            double off = new[] { Math.Abs(l - word[i].L), Math.Abs(t - word[i].T), Math.Abs(r - word[i].R), Math.Abs(b - word[i].B) }.Max();
            // the contour is 202020 under the scene's light: every face no channel of which passes 0x60 (the shapes are light blue)
            int contour = shapes[i].Count(f => !f.Fill.StartsWith("g ", StringComparison.Ordinal) && Enumerable.Range(0, 3).All(c => int.Parse(f.Fill.Substring(c * 2, 2), NumberStyles.HexNumber) < 0x60));
            TestReport.Check(off <= 1.5, $"parts {names[i]}: extent within 1.5 pt of Word's ({off:0.00} pt: {l:0.0} {t:0.0} {r:0.0} {b:0.0}; {shapes[i].Count} faces, {contour} contour)");
            if (i is < 4 or 7) TestReport.Check(contour > 0, $"parts {names[i]}: contour drawn ({contour} faces)");
            else if (i == 5)
            {
                // the frame's angle bevel: the cap (outline + hole) and a band per side of the outline and of the hole, each side's
                // pair facing opposite ways, so at least 4 shades besides the cap's
                int shades = shapes[i].Select(f => f.Fill).Distinct().Count();
                TestReport.Check(shapes[i].Count == 10 && shades >= 5, $"parts {names[i]}: bevel along the outline and the hole ({shapes[i].Count} faces, {shades} shades)");
            }
            else TestReport.Check(shapes[i].Count > 20, $"parts {names[i]}: bevel along every ring ({shapes[i].Count} faces)");
        }
    }

    /// <summary>
    /// Every bevel preset (shapes-3d-bevels-v1.docx: 150 x 110 pt rectangles, bevel 24 x 24 pt, threePt rig) against Word's band colours
    /// (shapes-3d-bevels-v1-word-samples.txt, sample_bevels.py: 48 steps across the left, right, top and bottom band, each the mean of 5
    /// pixels at 8 px per pt). The first point of a band (Word's anti-aliased edge) and points at a profile's corner are not judged.
    /// </summary>
    private static void BevelPresets()
    {
        string docx = Path.Combine(Folder, "shapes-3d-bevels-v1.docx"), samples = Path.Combine(Folder, "shapes-3d-bevels-v1-word-samples.txt");
        if (!File.Exists(docx) || !File.Exists(samples)) { TestReport.Check(false, "shapes-3d-bevels-v1 fixtures present"); return; }
        var shapes = Shapes(docx);
        string[] names = ["relaxedInset", "circle", "slope", "cross", "angle", "softRound", "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco"];
        TestReport.Check(shapes.Count == 12, $"bevel presets: 12 shapes ({shapes.Count})");
        if (shapes.Count != 12) return;
        var lines = File.ReadAllLines(samples).Select(l => l.Split(' ')).Where(p => p[1] != "centre").ToList();
        for (int s = 0; s < 12; s++)
        {
            var rows = lines.Where(p => p[0] == names[s]).ToList();
            var levels = new List<int>();
            for (int k = 1; k < rows.Count; k++)
            {
                double d = (k + 0.5) / 2;
                (double X, double Y)[] at = [(d, 55), (150 - d, 55), (75, d), (75, 110 - d)];
                for (int b = 0; b < 4; b++)
                {
                    string? ours = ColourAt(shapes[s], at[b].X, at[b].Y);
                    levels.Add(ours is null ? 999 : Levels(ours, rows[k][2 + b]));
                }
            }
            levels.Sort();
            int median = levels[levels.Count / 2], p90 = levels[levels.Count * 9 / 10];
            TestReport.Check(median <= 6 && p90 <= 12, $"bevel {names[s]}: band colours against Word's, median {median} levels, 90th percentile {p90}, worst {levels[^1]}");
        }
    }

    /// <summary>
    /// Every bevel preset on an oval (shapes-3d-bevels-oval-v1.docx: 150 x 110 pt ovals, bevel 24 x 24 pt, threePt rig; a curved
    /// outline of many edges) against Word's colours (shapes-3d-bevels-oval-v1-word-samples.txt, sample_bevels_oval.py: 48 steps
    /// inward along the outline's normal at 8 directions). The first point (Word's anti-aliased edge) is not judged.
    /// </summary>
    private static void OvalBevels()
    {
        string docx = Path.Combine(Folder, "shapes-3d-bevels-oval-v1.docx"), samples = Path.Combine(Folder, "shapes-3d-bevels-oval-v1-word-samples.txt");
        if (!File.Exists(docx) || !File.Exists(samples)) { TestReport.Check(false, "shapes-3d-bevels-oval-v1 fixtures present"); return; }
        var shapes = Shapes(docx);
        string[] names = ["relaxedInset", "circle", "slope", "cross", "angle", "softRound", "convex", "coolSlant", "divot", "riblet", "hardEdge", "artDeco"];
        TestReport.Check(shapes.Count == 12, $"oval bevels: 12 shapes ({shapes.Count})");
        if (shapes.Count != 12) return;
        var lines = File.ReadAllLines(samples).Select(l => l.Split(' ')).ToList();
        var all = new List<int>();
        for (int s = 0; s < 12; s++)
        {
            var rows = lines.Where(p => p[0] == names[s]).ToList();
            var levels = new List<int>();
            for (int k = 1; k < rows.Count; k++)
            {
                double d = (k + 0.5) / 2;
                for (int j = 0; j < 8; j++)
                {
                    double t = j * Math.PI / 4, ex = 75 + 75 * Math.Cos(t), ey = 55 + 55 * Math.Sin(t);
                    double nx = -Math.Cos(t) / 75, ny = -Math.Sin(t) / 55, length = Math.Sqrt(nx * nx + ny * ny);
                    string? ours = ColourAt(shapes[s], ex + nx / length * d, ey + ny / length * d);
                    levels.Add(ours is null ? 999 : Levels(ours, rows[k][2 + j]));
                }
            }
            levels.Sort();
            all.AddRange(levels);
            int median = levels[levels.Count / 2], p90 = levels[levels.Count * 9 / 10];
            TestReport.Check(median <= 6 && p90 <= 14, $"oval bevel {names[s]}: colours against Word's, median {median} levels, 90th percentile {p90}, worst {levels[^1]}");
        }
        all.Sort();
        Console.WriteLine($"   oval bevels, all presets: median {all[all.Count / 2]}, 90th percentile {all[all.Count * 9 / 10]}, 99th {all[all.Count * 99 / 100]}");
        // the best of five fresh opens (the first pays the JIT)
        long bytes = 0, best = long.MaxValue;
        for (int round = 0; round < 5; round++)
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(docx);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            bytes = document.RenderBlocks(0, document.BlockCount).Sum(block => (long)block.Html.Length);
            best = Math.Min(best, watch.ElapsedMilliseconds);
        }
        Console.WriteLine($"   oval bevels: {bytes / 1024} KB of HTML for 12 shapes, rendered in {best} ms (best of 5 opens)");
        // before the shaded bands (3404625) the flat bands came to 1,331 KB; per-edge gradients took it to 2,383 KB
        TestReport.Check(bytes < 1000 * 1024, $"oval bevels: HTML stays below the flat bands' size ({bytes / 1024} KB < 1,000 KB)");
    }

    private static List<List<(List<(double X, double Y)> Ring, string Fill)>> Shapes(string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        var result = new List<List<(List<(double X, double Y)>, string)>>();
        foreach (Match svg in Regex.Matches(html, "<svg class=\"word-shape-geometry\".*?</svg>", RegexOptions.Singleline))
        {
            double tx = 0, ty = 0;
            Match g = Regex.Match(svg.Value, "<g transform=\"([^\"]*)\"");
            if (g.Success)
                foreach (Match t in Regex.Matches(g.Groups[1].Value, @"translate\(([-\d.]+) ([-\d.]+)\)"))
                { tx += Number(t.Groups[1].Value); ty += Number(t.Groups[2].Value); }
            var faces = new List<(List<(double X, double Y)>, string)>();
            // a bevel band's gradient: "g x1 y1 x2 y2 offset:colour ..." in the shape's own points (ColourAt evaluates it)
            var gradients = new Dictionary<string, string>();
            foreach (Match gradient in Regex.Matches(svg.Value, "<linearGradient id=\"([^\"]+)\" gradientUnits=\"userSpaceOnUse\" x1=\"([-0-9.]+)\" y1=\"([-0-9.]+)\" x2=\"([-0-9.]+)\" y2=\"([-0-9.]+)\">(.*?)</linearGradient>"))
                gradients[gradient.Groups[1].Value] = FormattableString.Invariant($"g {Number(gradient.Groups[2].Value) + tx} {Number(gradient.Groups[3].Value) + ty} {Number(gradient.Groups[4].Value) + tx} {Number(gradient.Groups[5].Value) + ty} ")
                    + string.Join(" ", Regex.Matches(gradient.Groups[6].Value, "offset=\"([0-9.]+)\" stop-color=\"#([0-9a-fA-F]{6})\"").Cast<Match>().Select(stop => stop.Groups[1].Value + ":" + stop.Groups[2].Value.ToLowerInvariant()));
            foreach (Match p in Regex.Matches(svg.Value, "<path d=\"([^\"]*)\"[^>]*fill=\"(?:#([0-9a-fA-F]{6})|url[(]#([^)]+)[)])\""))
                foreach (Match sub in Regex.Matches(p.Groups[1].Value, "M[^M]*"))
                {
                    var ring = Regex.Matches(sub.Value, @"([-\d.]+) ([-\d.]+)").Cast<Match>().Select(m => (Number(m.Groups[1].Value) + tx, Number(m.Groups[2].Value) + ty)).ToList();
                    if (p.Groups[2].Success) faces.Add((ring, p.Groups[2].Value.ToLowerInvariant()));
                    else if (gradients.TryGetValue(p.Groups[3].Value, out string? paint)) faces.Add((ring, paint));
                }
            result.Add(faces);
        }
        result.Reverse();   // the shapes share one anchor: Word's file lists them last first
        return result;
    }

    private static double Number(string s) => double.Parse(s, CultureInfo.InvariantCulture);

    private static (double L, double T, double R, double B) Extent(List<(List<(double X, double Y)> Ring, string Fill)> faces, (double X, double Y) box)
    {
        var points = faces.SelectMany(f => f.Ring).ToList();
        return (points.Min(p => p.X) + box.X, points.Min(p => p.Y) + box.Y, points.Max(p => p.X) + box.X, points.Max(p => p.Y) + box.Y);
    }

    /// <summary>The fill of the last-drawn face containing (x, y) in the shape's own points.</summary>
    private static string? ColourAt(List<(List<(double X, double Y)> Ring, string Fill)> faces, double x, double y)
    {
        string? found = null;
        foreach (var (ring, fill) in faces)
        {
            bool inside = false;
            for (int i = 0, j = ring.Count - 1; i < ring.Count; j = i++)
                if ((ring[i].Y > y) != (ring[j].Y > y) && x < (ring[j].X - ring[i].X) * (y - ring[i].Y) / (ring[j].Y - ring[i].Y) + ring[i].X) inside = !inside;
            if (inside) found = fill;
        }
        if (found is null || !found.StartsWith("g ", StringComparison.Ordinal)) return found;
        // a gradient: the colour at the point's place along x1,y1 -> x2,y2, mixed linearly between the stops around it (as SVG draws it)
        string[] parts = found.Split(' ');
        double x1 = Number(parts[1]), y1 = Number(parts[2]), x2 = Number(parts[3]), y2 = Number(parts[4]);
        double dx = x2 - x1, dy = y2 - y1, t = ((x - x1) * dx + (y - y1) * dy) / Math.Max(1e-9, dx * dx + dy * dy);
        var stops = parts.Skip(5).Select(part => part.Split(':')).Select(stop => (Offset: Number(stop[0]), Colour: stop[1])).ToList();
        if (t <= stops[0].Offset) return stops[0].Colour;
        for (int i = 1; i < stops.Count; i++)
            if (t <= stops[i].Offset)
            {
                double f = (t - stops[i - 1].Offset) / Math.Max(1e-9, stops[i].Offset - stops[i - 1].Offset);
                return string.Concat(Enumerable.Range(0, 3).Select(c =>
                {
                    int a = int.Parse(stops[i - 1].Colour.Substring(c * 2, 2), NumberStyles.HexNumber), b = int.Parse(stops[i].Colour.Substring(c * 2, 2), NumberStyles.HexNumber);
                    return ((int)Math.Round(a + (b - a) * f)).ToString("x2", CultureInfo.InvariantCulture);
                }));
            }
        return stops[^1].Colour;
    }

    private static int Levels(string a, string b)
        => Enumerable.Range(0, 3).Max(i => Math.Abs(int.Parse(a.Substring(i * 2, 2), NumberStyles.HexNumber) - int.Parse(b.Substring(i * 2, 2), NumberStyles.HexNumber)));
}
