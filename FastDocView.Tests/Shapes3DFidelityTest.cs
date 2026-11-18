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
            int contour = shapes[i].Count(f => Enumerable.Range(0, 3).All(c => int.Parse(f.Fill.Substring(c * 2, 2), NumberStyles.HexNumber) < 0x60));
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
            foreach (Match p in Regex.Matches(svg.Value, "<path d=\"([^\"]*)\"[^>]*fill=\"#([0-9a-fA-F]{6})\""))
                foreach (Match sub in Regex.Matches(p.Groups[1].Value, "M[^M]*"))
                {
                    var ring = Regex.Matches(sub.Value, @"([-\d.]+) ([-\d.]+)").Cast<Match>().Select(m => (Number(m.Groups[1].Value) + tx, Number(m.Groups[2].Value) + ty)).ToList();
                    faces.Add((ring, p.Groups[2].Value.ToLowerInvariant()));
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
        return found;
    }

    private static int Levels(string a, string b)
        => Enumerable.Range(0, 3).Max(i => Math.Abs(int.Parse(a.Substring(i * 2, 2), NumberStyles.HexNumber) - int.Parse(b.Substring(i * 2, 2), NumberStyles.HexNumber)));
}
