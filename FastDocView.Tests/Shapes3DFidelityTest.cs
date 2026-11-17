using System.Globalization;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// 3-D shapes against Word 16's PDFs in test-data/word-shapes-3d (make-shapes-3d-fixture.ps1; measured 2026-11-17 at 4 px per pt):
/// shapes-3d-v1.docx (a:scene3d / a:sp3d: legacy oblique and perspective cameras, isometric, perspective and orthographic turns,
/// an extrusion colour, circle / slope / angle bevels, metal, an oblique camera under the balanced rig, a contour) and
/// shapes-3d-on-v1.rtf (Word's RTF save of it with f3D turned on, make_rtf_3d.py: the legacy OfficeArt 3-D properties).
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
        (422.5, 216.5, 544.8, 327.5, 1, "5587da", 16),  // orthographic, lat 35 lon -25, green extrusion: Word lights it brighter
        (59.8, 399.5, 180.0, 480.0, 1, "4779cc", 6),    // circle bevel only
        (249.5, 399.5, 370.2, 480.0, 1, "4979cd", 6),   // slope bevel on an oval
        (429.5, 399.5, 550.2, 480.0, 1, "315796", 10),  // angle bevel, legacy metal
        (53.2, 559.0, 174.0, 640.8, 2.5, "4375c8", 8),  // perspective lat 20 with a bevel, a right triangle: 2.3 pt on its slanted side
        (249.5, 552.2, 370.2, 640.2, 1, "5583ce", 22),  // oblique top, balanced rig: Word lights it 19 levels brighter
        (428.5, 549.0, 551.2, 626.2, 2.5, "4676c3", 8), // turned 50 degrees, flat material, 2 pt contour (not drawn: 2.2 pt)
    ];

    // the legacy OfficeArt scene (RTF, f3D on): extents of the shapes whose 3-D style Word wrote as a legacy one (others: -1)
    private static readonly (double L, double T, double R, double B)[] Rtf =
    [
        (59.5, 67.0, 192.8, 160.2), (249.5, 79.5, 370.2, 160.2), (429.5, 74.8, 550.2, 160.2), (59.5, 239.5, 180.2, 320.2), (249.5, 228.0, 370.2, 314.0),
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
            TestReport.Check(off <= e.Outline, $".docx shape {i}: extent within {e.Outline} pt of Word's ({off:0.00} pt)");
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
            }
        // Word's .doc / RTF save keeps the 3-D style with f3D off, and Word prints those flat: so do we
        var flat = Shapes(doc);
        TestReport.Check(flat.Count == 12 && flat.All(f => f.Count <= 2), ".doc with f3D off: flat shapes, as Word prints them");
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
