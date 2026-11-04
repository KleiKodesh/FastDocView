using FastDocView.Html;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// draw:connector routing without a stored svg:d (port of LibreOffice SdrEdgeObj::ImpCalcEdgeTrack, svx/source/svdraw/svdoedge.cxx).
/// Expected tracks are worked by hand in LibreOffice's 1/100 mm (1 pt = 2540/72), following the cited source lines.
/// </summary>
public static class OdtConnectorRouteTest
{
    private static double P(long hmm) => hmm * 72.0 / 2540.0;
    private static bool Near(double a, double b) => Math.Abs(a - b) < 1e-9;

    private static string Dump(List<WordShapeSubpath> paths)
    {
        var s = paths[0];
        var text = $"M {s.StartX:0.###} {s.StartY:0.###}";
        foreach (var g in s.Segments)
            text += g.IsCurve ? $" C {g.Control1X:0.###} {g.Control1Y:0.###} {g.Control2X:0.###} {g.Control2Y:0.###} {g.X:0.###} {g.Y:0.###}" : $" L {g.X:0.###} {g.Y:0.###}";
        return text;
    }

    /// <summary>Points in 1/100 mm; a null control pair = a line segment.</summary>
    private static void Expect(string what, List<WordShapeSubpath> paths, (long X, long Y) start, params (long X, long Y, (long X, long Y)? C1, (long X, long Y)? C2)[] segments)
    {
        bool ok = paths.Count == 1 && Near(paths[0].StartX, P(start.X)) && Near(paths[0].StartY, P(start.Y)) && paths[0].Segments.Count == segments.Length;
        for (int i = 0; ok && i < segments.Length; i++)
        {
            var g = paths[0].Segments[i];
            var e = segments[i];
            ok = Near(g.X, P(e.X)) && Near(g.Y, P(e.Y)) && g.IsCurve == (e.C1 != null);
            if (ok && e.C1 is { } c1 && e.C2 is { } c2)
                ok = Near(g.Control1X, P(c1.X)) && Near(g.Control1Y, P(c1.Y)) && Near(g.Control2X, P(c2.X)) && Near(g.Control2Y, P(c2.Y));
        }
        TestReport.Check(ok, $"connector route: {what} ({Dump(paths)})");
    }

    private static (long, long, (long, long)?, (long, long)?) L(long x, long y) => (x, y, null, null);

    public static void Run()
    {
        TestReport.Step("ODT connector routing (LibreOffice ImpCalcEdgeTrack)");
        var none = Array.Empty<double>();

        // "line" = SdrEdgeKind::OneLine: start to end (:912-920).
        Expect("line kind is one straight segment", OdtConnectorRoute.Route("line", (0, 0), (72, 36), none, null, null), (0, 0), L(2540, 1270));

        // "standard" between free points (0,0) -> (2540,2540): the bounds are the points, every escape pair is tried (:827-870).
        // Right/left Z: quality 5080 - 6 = 5074; down/up Z: 5080 - 6 - 1 (vertical start, :1402) = 5073; an L adds 3*0x01000000.
        // The meeting point is the middle of the empty space (:976-983): (1270,1270).
        Expect("standard between free points: vertical Z through the middle", OdtConnectorRoute.Route("standard", (0, 0), (72, 72), none, null, null),
            (0, 0), L(0, 1270), L(2540, 1270), L(2540, 2540));

        // Line skew on the Z: the middle line (index 1) is horizontal (angle 27000, odd index, :112-122) and moves 635 down (:1461-1470).
        Expect("standard with draw:line-skew moves the middle line", OdtConnectorRoute.Route("standard", (0, 0), (72, 72), new[] { 18.0 }, null, null),
            (0, 0), L(0, 1905), L(2540, 1905), L(2540, 2540));

        // "lines" = ThreeLines: with point bounds every pair scores 5080 and the first (right, right) wins; without skew the inner points
        // sit on the ends (:921-955). Skew 18pt 36pt: horizontal exits move x by 635 and 1270 (ImpSetAttrToEdgeInfo :261-282).
        Expect("lines kind without skew", OdtConnectorRoute.Route("lines", (0, 0), (72, 72), none, null, null),
            (0, 0), L(0, 0), L(2540, 2540), L(2540, 2540));
        Expect("lines kind with two skews", OdtConnectorRoute.Route("lines", (0, 0), (72, 72), new[] { 18.0, 36.0 }, null, null),
            (0, 0), L(635, 0), L(3810, 2540), L(2540, 2540));

        // "curve" on the free-point Z: control points move half a first/last line further (:1575-1585):
        // (0,1270)+(0,635) and (2540,1270)+(0,-635).
        Expect("curve kind on a Z: one cubic", OdtConnectorRoute.Route("curve", (0, 0), (72, 72), none, null, null),
            (0, 0), (2540, 2540, (0, 1905), (2540, 635)));

        // Vertex glue points (svdobj.cxx:2563-2582) on a 72pt square: tools::Rectangle right = left + width - 1 (gen.hxx:596),
        // so vertex 1 = (2539, 1269), escape RIGHT by ImpCalcEscAngle (:631-667).
        var right = OdtConnectorRoute.VertexGluePoint((0, 0, 72, 72), 1);
        TestReport.Check(Near(right.X, P(2539)) && Near(right.Y, P(1269)) && right.Escape == OdtGlueEscape.Right, "connector route: vertex glue point 1 = right centre, escapes right");
        var top = OdtConnectorRoute.VertexGluePoint((0, 0, 72, 72), 0);
        TestReport.Check(Near(top.X, P(1269)) && Near(top.Y, 0) && top.Escape == OdtGlueEscape.Top, "connector route: vertex glue point 0 = top centre, escapes up");

        // Two shapes side by side, no glue point given (best vertex, :822-831): right centre of A to left centre of B. The two ends meet
        // on one horizontal line, so the 'I' special case (:1288-1292) leaves one straight segment.
        var shapeA = new OdtConnectorEnd((0, 0, 72, 72), -1);
        var shapeB = new OdtConnectorEnd((144, 0, 72, 72), -1);
        Expect("standard between two shapes picks the facing vertices", OdtConnectorRoute.Route("standard", (0, 0), (0, 0), none, shapeA, shapeB),
            (2539, 1269), L(5080, 1269));

        // Curve from A's right (glue 1) to the top of B at (144,144) (glue 0 = (6349,5080)): case 3.2, an L (:1225-1236);
        // as a curve (:1563-1574): controls (6349-1270, 1269) and (6349, 1269+1270).
        var startA = new OdtConnectorEnd((0, 0, 72, 72), 1);
        var endB = new OdtConnectorEnd((144, 144, 72, 72), 0);
        Expect("standard L from a right side to a top side", OdtConnectorRoute.Route("standard", (0, 0), (0, 0), none, startA, endB),
            (2539, 1269), L(6349, 1269), L(6349, 5080));
        Expect("curve L control points", OdtConnectorRoute.Route("curve", (0, 0), (0, 0), none, startA, endB),
            (2539, 1269), (6349, 5080, (5079, 1269), (6349, 2539)));
    }
}
