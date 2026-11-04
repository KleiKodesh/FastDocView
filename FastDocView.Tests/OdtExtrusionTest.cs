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

    private static List<OdtExtrusionFace> Faces(params (string Name, string Value)[] attributes)
    {
        var geometry = new XElement(Draw + "enhanced-geometry", new XAttribute(Draw + "extrusion", "true"), new XAttribute(Draw + "extrusion-depth", "40pt 0"));
        foreach (var (name, value) in attributes) geometry.Add(new XAttribute(Draw + name, value));
        OdtExtrusion extrusion = OdtExtrusion.Read(geometry)!;
        var source = new OdtExtrusionSource();
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

        // 4. no draw:extrusion: drawn flat (Read returns null)
        TestReport.Check(OdtExtrusion.Read(new XElement(Draw + "enhanced-geometry")) is null, "extrusion: absent without draw:extrusion=\"true\"");
    }
}
