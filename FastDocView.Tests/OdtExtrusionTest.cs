using System.Xml.Linq;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// Custom-shape 3-D extrusion (OdtExtrusion, port of LibreOffice EnhancedCustomShape3d::Create3DObject and the 3-D renderer's flat
/// shading). A 100 x 50 pt rectangle, depth 40 pt (fraction 0: all backward), default fill #729fcf = (114, 159, 207).
/// Worked by hand from the cited lines:
/// - object coordinates (x, -y, z), back z = 0, front z = 40; scene matrix translate(-50, 25, -40) then the case's operations;
///   screen = (X + 50, -Y + 25);
/// - lights: L1 = normalize(50000, 0, 10000) = (0.980581, 0, 0.196116), L2 = (-0.980581, 0, 0.196116), level 0.66 as 8 bits
///   = 168/255 = 0.658824, ambient 0.33 = 84/255 = 0.329412; specularity 0 so light 1 becomes non-specular light 4;
///   colour = fill * (ambient + sum level * max(0, L . n)), n = the matrix' linear part times the outward normal, normalised;
///   channels rounded to bytes.
/// </summary>
public static class OdtExtrusionTest
{
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static bool Near(double a, double b) => Math.Abs(a - b) < 1e-3;

    private static List<OdtExtrusionFace> Faces(params (string Name, string Value)[] attributes) => Faces(true, attributes);

    private static List<OdtExtrusionFace> Faces(bool filled, params (string Name, string Value)[] attributes)
    {
        var geometry = new XElement(Draw + "enhanced-geometry", new XAttribute(Draw + "extrusion", "true"), new XAttribute(Draw + "extrusion-depth", "40pt 0"));
        foreach (var (name, value) in attributes) geometry.Add(new XAttribute(Draw + name, value));
        OdtExtrusion extrusion = OdtExtrusion.Read(geometry)!;
        var source = new OdtExtrusionSource { Filled = filled };
        source.Rings.Add((new List<(double X, double Y)> { (0, 0), (100, 0), (100, 50), (0, 50) }, true));
        return extrusion.Project([source], 100, 50, 0, false, false, (114 / 255.0, 159 / 255.0, 207 / 255.0), (114 / 255.0, 159 / 255.0, 207 / 255.0), (0, 0, 0), false);
    }

    private static void Expect(string what, List<OdtExtrusionFace> faces, params (string Kind, string Color, (double X, double Y)[] Points)[] expected)
    {
        bool ok = faces.Count == expected.Length;
        for (int i = 0; ok && i < expected.Length; i++)
        {
            OdtExtrusionFace face = faces[i];
            ok = face.Kind == expected[i].Kind && face.Color == expected[i].Color && face.Rings.Count == 1 && face.Rings[0].Count == expected[i].Points.Length;
            for (int k = 0; ok && k < expected[i].Points.Length; k++)
                ok = Near(face.Rings[0][k].X, expected[i].Points[k].X) && Near(face.Rings[0][k].Y, expected[i].Points[k].Y);
        }
        string dump = string.Join(" | ", faces.Select(f => f.Kind + " " + f.Color + " " + string.Join(" ", f.Rings.Select(r => string.Join(",", r.Select(p => $"({p.X:0.###} {p.Y:0.###})"))))));
        TestReport.Check(ok, $"extrusion: {what} ({dump})");
    }


    // The live rig's fixture (vue-demo/playwright/odtExtrusionCheck.mjs): a 2 x 1 in rectangle custom shape extruded six ways.
    private static readonly (string Name, string Style, string Attributes)[] FixtureShapes =
    [
        ("parallel", "Filled", ""),
        ("perspective", "Filled", "dr3d:projection=\"perspective\" draw:extrusion-viewpoint=\"(1in -1in 10in)\""),
        ("rotated", "Filled", "draw:extrusion-rotation-angle=\"20 30\""),
        ("draft", "Filled", "dr3d:shade-mode=\"draft\""),
        ("unfilled", "Unfilled", ""),
        ("shaded", "Filled", "draw:shade-mode=\"gouraud\" draw:extrusion-specularity=\"80%\" draw:extrusion-shininess=\"80%\" draw:extrusion-metal=\"true\" draw:extrusion-color=\"true\""),
    ];

    /// <summary>Writes TestOutput/fidelity-v1-odt-extrusion.fodt (rewritten only when its text changes) and returns its path.</summary>
    public static string WriteFixture()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-extrusion.fodt");
        string body = string.Concat(FixtureShapes.Select(shape =>
            "<text:p>" + shape.Name + "<draw:custom-shape draw:style-name=\"" + shape.Style + "\" text:anchor-type=\"as-char\" svg:width=\"2in\" svg:height=\"1in\">"
            + "<draw:enhanced-geometry svg:viewBox=\"0 0 21600 21600\" draw:type=\"rectangle\" draw:enhanced-path=\"M 0 0 L 21600 0 21600 21600 0 21600 Z N\" "
            + "draw:extrusion=\"true\" draw:extrusion-depth=\"0.5in 0\" " + shape.Attributes + "/></draw:custom-shape></text:p>"));
        string document = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
            + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
            + "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:dr3d=\"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0\" "
            + "xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles>"
            + "<style:style style:name=\"Filled\" style:family=\"graphic\"><style:graphic-properties draw:fill=\"solid\" draw:fill-color=\"#729fcf\" svg:stroke-color=\"#3465a4\"/></style:style>"
            + "<style:style style:name=\"Unfilled\" style:family=\"graphic\"><style:graphic-properties draw:fill=\"none\" draw:stroke=\"solid\" svg:stroke-color=\"#c9211e\"/></style:style>"
            + "</office:automatic-styles><office:body><office:text>" + body + "</office:text></office:body></office:document>";
        if (!File.Exists(path) || File.ReadAllText(path) != document) File.WriteAllText(path, document, new System.Text.UTF8Encoding(false));
        return path;
    }

    private static void CheckFixture()
    {
        TestReport.Step("ODT custom-shape 3-D extrusion: fixture (parallel, perspective, rotated, draft, unfilled, shaded)");
        List<string> blocks = OdtDrawingFidelityTest.Render(WriteFixture());
        TestReport.Check(blocks.Count == FixtureShapes.Length, $"fixture: {blocks.Count} blocks, one per shape");
        for (int i = 0; i < Math.Min(blocks.Count, FixtureShapes.Length); i++)
        {
            int paths = System.Text.RegularExpressions.Regex.Matches(blocks[i], "<path ").Count;
            TestReport.Check(blocks[i].Contains("<svg") && paths >= 2, $"fixture {FixtureShapes[i].Name}: SVG with {paths} faces");
        }
        if (blocks.Count < FixtureShapes.Length) return;
        static string Geometry(string html) => string.Join("|", System.Text.RegularExpressions.Regex.Matches(html, " d=\"([^\"]*)\"").Select(m => m.Groups[1].Value));
        TestReport.Check(Geometry(blocks[1]) != Geometry(blocks[0]) && Geometry(blocks[2]) != Geometry(blocks[0]), "fixture: perspective and rotated faces differ from parallel");
        TestReport.Check(!System.Text.RegularExpressions.Regex.IsMatch(blocks[3], "<path [^>]*fill=\"#"), "fixture draft: faces drawn as outlines, no fill");
        static List<string> Fills(string html) => System.Text.RegularExpressions.Regex.Matches(html, "fill=\"#([0-9a-fA-F]{6})\"").Select(m => m.Groups[1].Value).ToList();
        static bool Reddish(string hex) => Convert.ToInt32(hex[..2], 16) > Convert.ToInt32(hex[4..], 16);
        List<string> unfilled = Fills(blocks[4]);
        TestReport.Check(unfilled.Count > 0 && unfilled.All(Reddish), $"fixture unfilled: faces lit in the line colour #c9211e, not the fill ({string.Join(",", unfilled)})");
        TestReport.Check(!Fills(blocks[5]).SequenceEqual(Fills(blocks[0])), "fixture shaded: specular / metal / extrusion colour change the face colours");
    }

    public static void Run()
    {
        TestReport.Step("ODT custom-shape 3-D extrusion");

        // 1. parallel, defaults: skew 50 at -135 degrees (GetSkew :62-75) = shearXY(0.5 cos -135, 0.5 sin -135) = (-0.353553, -0.353553).
        //    Front (z' = 0) stays; back (z' = -40) moves x + 14.1421, y(up) + 14.1421 = screen (x + 14.1421, y - 14.1421).
        //    Visible: top and right sides, front (back, bottom, left turn away). Depths: sides -20, front 0.
        //    front n = (-0.353553, -0.353553, 1) / 1.118034: L1 . n < 0; L2 . n = 0.485499 -> 0.329412 + 0.319859 = 0.649271 -> (74, 103, 134)
        //    top n = (0, 1, 0): ambient only -> (38, 52, 68); right n = (1, 0, 0): L1 0.980581 -> 0.975442 -> (111, 155, 202)
        Expect("parallel, default skew", Faces(),
            ("side", "#263444", [(0, 0), (100, 0), (114.1421, -14.1421), (14.1421, -14.1421)]),
            ("side", "#6f9bca", [(100, 0), (100, 50), (114.1421, 35.8579), (114.1421, -14.1421)]),
            ("front", "#4a6786", [(0, 0), (100, 0), (100, 50), (0, 50)]));

        // 2. parallel, rotation angles 0 30 (x 0, y 30; :686-689 rotate y by -30), no skew:
        //    x' = 0.866025 x - 0.5 z, z' = 0.5 x + 0.866025 z. Front x' = +-43.3013; back (z = -40) x' = 0.866025 x + 20.
        //    Visible: front n = (-0.5, 0, 0.866025): L2 0.660132 -> 0.764323 -> (87, 122, 158);
        //    right n = (0.866025, 0, 0.5): L1 0.947266 -> 0.953493 -> (109, 152, 197). Top / bottom are edge-on, left and back away.
        //    Depths: front 0, right (25 + 25 - 9.641 - 9.641) / 4 = 7.68 -> front first.
        Expect("parallel, rotated 0 30", Faces(("extrusion-rotation-angle", "0 30"), ("extrusion-skew", "0 0")),
            ("front", "#577a9e", [(6.6987, 0), (93.3013, 0), (93.3013, 50), (6.6987, 50)]),
            ("side", "#6d98c5", [(93.3013, 0), (93.3013, 50), (113.3013, 50), (113.3013, 0)]));

        // 3. perspective, viewpoint (200pt -100pt 1000pt), origin 0 0 (:716-749): origin offset 0, camera (200, 100, 1000) in y-up scene
        //    coordinates. Front z' = 0: f = 1, unchanged. Back z' = -40: f = 1000 / 1040 = 0.961538;
        //    X = (x' - 200) f + 200: -50 -> -40.3846 (screen 9.6154), 50 -> 55.7692 (105.7692);
        //    Y = (y' - 100) f + 100: 25 -> 27.8846 (screen -2.8846), -25 -> -20.1923 (45.1923). No skew in perspective (:693).
        //    front n = (0, 0, 1): L1 and L2 0.196116 each -> 0.587826 -> (67, 93, 122); top ambient (38, 52, 68); right (111, 155, 202).
        Expect("perspective", Faces(("projection", "perspective"), ("extrusion-viewpoint", "(200pt -100pt 1000pt)"), ("extrusion-origin", "0 0")),
            ("side", "#263444", [(0, 0), (100, 0), (105.7692, -2.8846), (9.6154, -2.8846)]),
            ("side", "#6f9bca", [(100, 0), (100, 50), (105.7692, 45.1923), (105.7692, -2.8846)]),
            ("front", "#435d7a", [(0, 0), (100, 0), (100, 50), (0, 50)]));

        // 4. unfilled shape (draw:fill none), draft: no front / back caps even in draft - the object keeps its fill style (:378-383, :606-613)
        var unfilled = Faces(false, ("shade-mode", "draft"));
        TestReport.Check(unfilled.Count > 0 && unfilled.All(f => f.Kind == "side"), $"extrusion: unfilled draft shape has sides only ({string.Join(",", unfilled.Select(f => f.Kind))})");

        // 5. perspective, viewpoint at z 0 = the front plane (z' = 0): front points have no projection and are skipped (:797-798),
        //    so the front face is not drawn; the back face's points (f = 0) collapse onto the viewpoint, so every side keeps fewer than
        //    3 points: nothing is drawn, no degenerate ring. A viewpoint at z 20pt (front 4 points projected) keeps the front face.
        var atViewpoint = Faces(("projection", "perspective"), ("extrusion-viewpoint", "(0pt 0pt 0pt)"), ("extrusion-origin", "0 0"));
        var behind = Faces(("projection", "perspective"), ("extrusion-viewpoint", "(0pt 0pt -20pt)"), ("extrusion-origin", "0 0"));
        TestReport.Check(behind.Count > 0 && behind.All(f => f.Rings.All(r => r.Count >= 3)), $"extrusion: viewpoint inside the solid: faces drawn, every ring has 3+ points ({behind.Count})");
        TestReport.Check(atViewpoint.Count == 0, $"extrusion: points at the viewpoint's depth are skipped, no degenerate face ({atViewpoint.Count} faces)");

        // 6. no draw:extrusion: drawn flat (Read returns null)
        TestReport.Check(OdtExtrusion.Read(new XElement(Draw + "enhanced-geometry")) is null, "extrusion: absent without draw:extrusion=\"true\"");
        CheckFixture();
    }
}
