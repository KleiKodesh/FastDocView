using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>Escape directions of a glue point (LibreOffice SdrEscapeDirection, include/svx/svdglue.hxx:34-44). Smart = derived from the point's place in the shape.</summary>
[Flags]
internal enum OdtGlueEscape : byte { Smart = 0x00, Left = 0x01, Right = 0x02, Top = 0x04, Bottom = 0x08, Horizontal = Left | Right, Vertical = Top | Bottom, All = 0xff }

/// <summary>A glue point in points (1/72 in) with its escape directions.</summary>
internal readonly record struct OdtConnectorGluePoint(double X, double Y, OdtGlueEscape Escape);

/// <summary>
/// A connector end attached to a shape (draw:start-shape / draw:end-shape). <see cref="SnapRect"/> = the shape's rectangle in points
/// (x, y, width, height), used for the default glue points and smart escape directions; <see cref="BoundRect"/> = the rectangle the
/// connector routes around (LibreOffice: GetCurrentBoundRect, svdoedge.cxx:770; defaults to the snap rectangle).
/// <see cref="GlueIndex"/>: -1 = no draw:start-glue-point (LibreOffice picks the best of the four vertex glue points),
/// 0..3 = vertex glue point, anything else = the user glue point given in <see cref="CustomGlue"/>.
/// </summary>
internal sealed record OdtConnectorEnd(
    (double X, double Y, double Width, double Height) SnapRect,
    int GlueIndex,
    OdtConnectorGluePoint? CustomGlue = null,
    (double X, double Y, double Width, double Height)? BoundRect = null);

/// <summary>
/// The track of a draw:connector when the file stores no svg:d: a port of LibreOffice's SdrEdgeObj::ImpCalcEdgeTrack (both overloads),
/// ImpCalcObjToCenter and ImpCalcEscAngle (svx/source/svdraw/svdoedge.cxx, master), in LibreOffice's integer 1/100 mm.
/// </summary>
internal static class OdtConnectorRoute
{
    // SdrEdgeKind (include/svx/sxekitm.hxx); draw:type -> kind: xmloff/source/draw/sdpropls.cxx:530-536 (enum map) and
    // svx/source/svdraw/svdattr.cxx:1566-1569 (ConnectorType -> SdrEdgeKind); default STANDARD: ximpshap.cxx:1812.
    private enum EdgeKind { OrthoLines, ThreeLines, OneLine, Bezier }

    private const byte FlagNormal = 0, FlagControl = 2, FlagSymmetric = 3;   // PolyFlags (include/tools/poly.hxx)
    private const int NoMiddleLine = 0xFFFF;
    /// <summary>SDRATTR_EDGENODE1/2 HORZ/VERTDIST pool defaults: 500 (svx/source/svdraw/svdattr.cxx:273-276).</summary>
    private const long NodeDistance = 500;

    private struct LP : IEquatable<LP>
    {
        public long X, Y;
        public LP(long x, long y) { X = x; Y = y; }
        public bool Equals(LP o) => X == o.X && Y == o.Y;
        public override bool Equals(object? o) => o is LP p && Equals(p);
        public override int GetHashCode() => (int)(X * 31 + Y);
        public static bool operator ==(LP a, LP b) => a.Equals(b);
        public static bool operator !=(LP a, LP b) => !a.Equals(b);
    }

    /// <summary>tools::Rectangle with inclusive right/bottom (include/tools/gen.hxx:593-603, 773-779).</summary>
    private struct LR
    {
        public long L, T, R, B;
        public LR(long l, long t, long r, long b) { L = l; T = t; R = r; B = b; }
        public LP Center => new((L + R) / 2, (T + B) / 2);
        /// <summary>Rectangle(Point, Size): right = left + (width - 1) (gen.hxx:596).</summary>
        public static LR FromPointSize(long x, long y, long w, long h)
            => new(x, y, w != 0 ? x + w + (w > 0 ? -1 : 1) : x, h != 0 ? y + h + (h > 0 ? -1 : 1) : y);
    }

    private sealed class Track
    {
        public readonly List<LP> P = new();
        public readonly List<byte> F = new();
        public int Count => P.Count;
        public void Add(LP p, byte f = FlagNormal) { P.Add(p); F.Add(f); }
        public void Insert(int i, LP p, byte f) { P.Insert(i, p); F.Insert(i, f); }
        public void RemoveAt(int i) { P.RemoveAt(i); F.RemoveAt(i); }
        public LP this[int i] { get => P[i]; set => P[i] = value; }
        public void AdjustX(int i, long d) { var p = P[i]; p.X += d; P[i] = p; }
        public void AdjustY(int i, long d) { var p = P[i]; p.Y += d; P[i] = p; }
        public void SetX(int i, long v) { var p = P[i]; p.X = v; P[i] = p; }
        public void SetY(int i, long v) { var p = P[i]; p.Y = v; P[i] = p; }
    }

    private enum LineCode { Obj1Line2, Obj1Line3, Obj2Line2, Obj2Line3, MiddleLine }

    /// <summary>SdrEdgeInfoRec (include/svx/svdoedge.hxx:80-107; methods svdoedge.cxx:88-138).</summary>
    private sealed class EdgeInfo
    {
        public LP Obj1Line2, Obj1Line3, Obj2Line2, Obj2Line3, MiddleLine;
        public long Angle1, Angle2;
        public int Obj1Lines, Obj2Lines, MiddleLineIndex = NoMiddleLine;

        public EdgeInfo Clone() => (EdgeInfo)MemberwiseClone();

        private ref LP OffsetPoint(LineCode code)
        {
            switch (code)
            {
                case LineCode.Obj1Line2: return ref Obj1Line2;
                case LineCode.Obj1Line3: return ref Obj1Line3;
                case LineCode.Obj2Line2: return ref Obj2Line2;
                case LineCode.Obj2Line3: return ref Obj2Line3;
                default: return ref MiddleLine;
            }
        }

        // svdoedge.cxx:100-110
        public int PolyIndex(LineCode code, int count) => code switch
        {
            LineCode.Obj1Line2 => 1,
            LineCode.Obj1Line3 => 2,
            LineCode.Obj2Line2 => count - 3,
            LineCode.Obj2Line3 => count - 4,
            _ => MiddleLineIndex,
        };

        // svdoedge.cxx:112-122
        public bool IsHorizontalLine(LineCode code, int count)
        {
            int index = PolyIndex(code, count);
            bool horizontal = Angle1 == 0 || Angle1 == 18000;
            if (code == LineCode.Obj2Line2 || code == LineCode.Obj2Line3)
            {
                index = count - index;
                horizontal = Angle2 == 0 || Angle2 == 18000;
            }
            if ((index & 1) == 1) horizontal = !horizontal;
            return horizontal;
        }

        // svdoedge.cxx:124-129
        public void SetLineOffset(LineCode code, int count, long value)
        {
            ref LP p = ref OffsetPoint(code);
            if (IsHorizontalLine(code, count)) p.Y = value; else p.X = value;
        }

        public LP Offset(LineCode code) => OffsetPoint(code);
    }

    /// <summary>One end as SdrObjConnection sees it.</summary>
    private sealed class Connection
    {
        public bool Connected;
        public LR Snap, Bound;
        public bool BestVertex;
        public int GlueIndex;
        public LP CustomPos;
        public OdtGlueEscape CustomEscape;
    }

    private static long ToHmm(double pt) => (long)Math.Round(pt * 2540.0 / 72.0, MidpointRounding.AwayFromZero);
    private static double ToPt(long hmm) => hmm * 72.0 / 2540.0;

    /// <summary>
    /// The default (vertex) glue point 0 = top centre, 1 = right centre, 2 = bottom centre, 3 = left centre of a shape's rectangle, as
    /// SdrObject::GetVertexGluePoint makes it (svx/source/svdraw/svdobj.cxx:2563-2582: snap rectangle, escape SMART per the
    /// SdrGluePoint(Point) constructor, include/svx/svdglue.hxx:87-91) placed by SdrGluePoint::GetAbsolutePos (svdglue.cxx:48-86).
    /// The smart escape is resolved here as SdrEdgeObj::ImpCalcEscAngle does (svdoedge.cxx:631-667).
    /// </summary>
    internal static OdtConnectorGluePoint VertexGluePoint((double X, double Y, double Width, double Height) rectPt, int index)
    {
        var snap = LR.FromPointSize(ToHmm(rectPt.X), ToHmm(rectPt.Y), ToHmm(rectPt.Width), ToHmm(rectPt.Height));
        var p = VertexGluePointHmm(snap, index);
        return new OdtConnectorGluePoint(ToPt(p.X), ToPt(p.Y), CalcEscAngle(snap, p));
    }

    private static LP VertexGluePointHmm(LR r, int index)
    {
        // svdobj.cxx:2569-2577: position relative to the centre ...
        LP p = index switch
        {
            0 => new LP((r.L + r.R) / 2, r.T),
            1 => new LP(r.R, (r.T + r.B) / 2),
            2 => new LP((r.L + r.R) / 2, r.B),
            3 => new LP(r.L, (r.T + r.B) / 2),
            _ => default,
        };
        var c = r.Center;
        p = new LP(p.X - c.X, p.Y - c.Y);
        // ... made absolute by GetAbsolutePos (svdglue.cxx:55, 80-85: align NONE = centre, no percent, clamp to the rectangle)
        p = new LP(p.X + c.X, p.Y + c.Y);
        if (p.X < r.L) p.X = r.L;
        if (p.X > r.R) p.X = r.R;
        if (p.Y < r.T) p.Y = r.T;
        if (p.Y > r.B) p.Y = r.B;
        return p;
    }

    /// <summary>
    /// Routes a connector. <paramref name="kind"/> = draw:type (standard / lines / line / curve; anything else = standard).
    /// <paramref name="lineDeltasPt"/> = draw:line-skew lengths in points (up to three). Ends that are null are free points.
    /// </summary>
    internal static List<WordShapeSubpath> Route(string kind, (double X, double Y) start, (double X, double Y) end, double[] lineDeltasPt,
        OdtConnectorEnd? startNode, OdtConnectorEnd? endNode)
    {
        var edgeKind = kind switch { "curve" => EdgeKind.Bezier, "line" => EdgeKind.OneLine, "lines" => EdgeKind.ThreeLines, _ => EdgeKind.OrthoLines };
        var con1 = MakeConnection(startNode);
        var con2 = MakeConnection(endNode);
        long[] deltas = new long[3];
        for (int i = 0; i < 3 && lineDeltasPt != null && i < lineDeltasPt.Length; i++) deltas[i] = ToHmm(lineDeltasPt[i]);

        var track0 = new Track();
        track0.Add(new LP(ToHmm(start.X), ToHmm(start.Y)));
        track0.Add(new LP(ToHmm(end.X), ToHmm(end.Y)));

        // A connector imported with the model locked is laid out once, the stored deltas set onto the result, then laid out again
        // (ImpRecalcEdgeTrack, svdoedge.cxx:581-616, #i123048).
        var info = new EdgeInfo();
        var track = CalcEdgeTrack(edgeKind, track0, con1, con2, info);
        SetAttrToEdgeInfo(edgeKind, deltas, info, track);
        track = CalcEdgeTrack(edgeKind, track, con1, con2, info);

        var subpath = new WordShapeSubpath { StartX = ToPt(track[0].X), StartY = ToPt(track[0].Y) };
        for (int i = 1; i < track.Count;)
        {
            if (track.F[i] == FlagControl && i + 2 < track.Count)
            {
                subpath.Segments.Add(new WordShapeSegment(true, ToPt(track[i].X), ToPt(track[i].Y), ToPt(track[i + 1].X), ToPt(track[i + 1].Y),
                    ToPt(track[i + 2].X), ToPt(track[i + 2].Y)));
                i += 3;
            }
            else
            {
                subpath.Segments.Add(WordShapeSegment.Line(ToPt(track[i].X), ToPt(track[i].Y)));
                i++;
            }
        }
        return new List<WordShapeSubpath> { subpath };
    }

    private static Connection MakeConnection(OdtConnectorEnd? node)
    {
        var c = new Connection();
        if (node == null) return c;
        c.Connected = true;
        var s = node.SnapRect;
        c.Snap = LR.FromPointSize(ToHmm(s.X), ToHmm(s.Y), ToHmm(s.Width), ToHmm(s.Height));
        var b = node.BoundRect ?? s;
        c.Bound = LR.FromPointSize(ToHmm(b.X), ToHmm(b.Y), ToHmm(b.Width), ToHmm(b.Height));
        // SdrEdgeObj::setGluePointIndex (svdoedge.cxx:2632-2655): < 0 = best vertex, 0..3 = auto vertex, else a user glue point.
        c.BestVertex = node.GlueIndex < 0;
        c.GlueIndex = node.GlueIndex < 0 ? 0 : node.GlueIndex;
        if (node.GlueIndex > 3)
        {
            var g = node.CustomGlue ?? new OdtConnectorGluePoint(s.X + s.Width / 2, s.Y + s.Height / 2, OdtGlueEscape.Smart);
            c.CustomPos = new LP(ToHmm(g.X), ToHmm(g.Y));
            c.CustomEscape = g.Escape;
        }
        return c;
    }

    /// <summary>SdrObjConnection::TakeGluePoint (svdoedge.cxx:58-85) for a vertex or user glue point.</summary>
    private static (LP Pos, OdtGlueEscape Escape) TakeGluePoint(Connection c, int vertex)
    {
        if (vertex <= 3) return (VertexGluePointHmm(c.Snap, vertex), OdtGlueEscape.Smart);
        return (c.CustomPos, c.CustomEscape);
    }

    /// <summary>ImpSetAttrToEdgeInfo (svdoedge.cxx:213-283): EdgeLine1..3Delta onto the lines the track has.</summary>
    private static void SetAttrToEdgeInfo(EdgeKind kind, long[] values, EdgeInfo info, Track track)
    {
        if (kind == EdgeKind.OrthoLines || kind == EdgeKind.Bezier)
        {
            int n = 0, count = track.Count;
            if (info.Obj1Lines >= 2 && n < 3) info.SetLineOffset(LineCode.Obj1Line2, count, values[n++]);
            if (info.Obj1Lines >= 3 && n < 3) info.SetLineOffset(LineCode.Obj1Line3, count, values[n++]);
            if (info.MiddleLineIndex != NoMiddleLine && n < 3) info.SetLineOffset(LineCode.MiddleLine, count, values[n++]);
            if (info.Obj2Lines >= 3 && n < 3) info.SetLineOffset(LineCode.Obj2Line3, count, values[n++]);
            if (info.Obj2Lines >= 2 && n < 3) info.SetLineOffset(LineCode.Obj2Line2, count, values[n++]);
        }
        else if (kind == EdgeKind.ThreeLines)
        {
            bool hor1 = info.Angle1 == 0 || info.Angle1 == 18000;
            bool hor2 = info.Angle2 == 0 || info.Angle2 == 18000;
            if (hor1) info.Obj1Line2.X = values[0]; else info.Obj1Line2.Y = values[0];
            if (hor2) info.Obj2Line2.X = values[1]; else info.Obj2Line2.Y = values[1];
        }
    }

    /// <summary>SdrEdgeObj::ImpCalcEscAngle (svdoedge.cxx:631-667).</summary>
    private static OdtGlueEscape CalcEscAngle(LR r, LP pt)
    {
        long dxl = pt.X - r.L, dyo = pt.Y - r.T, dxr = r.R - pt.X, dyu = r.B - pt.Y;
        bool xMid = Math.Abs(dxl - dxr) < 2, yMid = Math.Abs(dyo - dyu) < 2;
        long dx = Math.Min(dxl, dxr), dy = Math.Min(dyo, dyu);
        bool diag = Math.Abs(dx - dy) < 2;
        if (xMid && yMid) return OdtGlueEscape.All;
        if (diag)
        {
            var ret = OdtGlueEscape.Smart;
            if (yMid) ret |= OdtGlueEscape.Vertical;
            if (xMid) ret |= OdtGlueEscape.Horizontal;
            if (dxl < dxr) ret |= dyo < dyu ? OdtGlueEscape.Left | OdtGlueEscape.Top : OdtGlueEscape.Left | OdtGlueEscape.Bottom;
            else ret |= dyo < dyu ? OdtGlueEscape.Right | OdtGlueEscape.Top : OdtGlueEscape.Right | OdtGlueEscape.Bottom;
            return ret;
        }
        if (dx < dy)
        {
            if (xMid) return OdtGlueEscape.Horizontal;
            return dxl < dxr ? OdtGlueEscape.Left : OdtGlueEscape.Right;
        }
        if (yMid) return OdtGlueEscape.Vertical;
        return dyo < dyu ? OdtGlueEscape.Top : OdtGlueEscape.Bottom;
    }

    /// <summary>SdrEdgeObj::ImpCalcObjToCenter (svdoedge.cxx:669-726).</summary>
    private static List<LP> CalcObjToCenter(LP start, long escAngle, LR rect, LP meeting)
    {
        var xp = new List<LP> { start };
        bool rts = escAngle == 0, obn = escAngle == 9000, lks = escAngle == 18000, unt = escAngle == 27000;
        LP p1 = start;
        if (lks) p1.X = rect.L;
        if (rts) p1.X = rect.R;
        if (obn) p1.Y = rect.T;
        if (unt) p1.Y = rect.B;
        LP p2 = p1;
        if (lks && meeting.X <= p2.X) p2.X = meeting.X;
        if (rts && meeting.X >= p2.X) p2.X = meeting.X;
        if (obn && meeting.Y <= p2.Y) p2.Y = meeting.Y;
        if (unt && meeting.Y >= p2.Y) p2.Y = meeting.Y;
        xp.Add(p2);
        LP p3 = p2;
        if ((lks && meeting.X > p2.X) || (rts && meeting.X < p2.X))
        {
            if (meeting.Y < p2.Y) { p3.Y = rect.T; if (meeting.Y < p3.Y) p3.Y = meeting.Y; }
            else { p3.Y = rect.B; if (meeting.Y > p3.Y) p3.Y = meeting.Y; }
            xp.Add(p3);
            if (p3.Y != meeting.Y) { p3.X = meeting.X; xp.Add(p3); }
        }
        if ((obn && meeting.Y > p2.Y) || (unt && meeting.Y < p2.Y))
        {
            if (meeting.X < p2.X) { p3.X = rect.L; if (meeting.X < p3.X) p3.X = meeting.X; }
            else { p3.X = rect.R; if (meeting.X > p3.X) p3.X = meeting.X; }
            xp.Add(p3);
            if (p3.X != meeting.X) { p3.Y = meeting.Y; xp.Add(p3); }
        }
        return xp;
    }

    /// <summary>SdrEdgeObj::ImpCalcEdgeTrack(rTrack0, rCon1, rCon2, pInfo) (svdoedge.cxx:728-878): every glue point and escape direction pair, best quality kept.</summary>
    private static Track CalcEdgeTrack(EdgeKind kind, Track track0, Connection con1, Connection con2, EdgeInfo info)
    {
        LP pt1 = track0[0], pt2 = track0[track0.Count - 1];
        var esc1 = OdtGlueEscape.All;
        var esc2 = OdtGlueEscape.All;
        LR bound1, bound2, beware1, beware2;
        if (con1.Connected)
        {
            bound1 = con1.Bound;
            beware1 = new LR(bound1.L - NodeDistance, bound1.T - NodeDistance, bound1.R + NodeDistance, bound1.B + NodeDistance);   // :775-781
        }
        else beware1 = bound1 = new LR(pt1.X, pt1.Y, pt1.X, pt1.Y);   // :783-788
        if (con2.Connected)
        {
            bound2 = con2.Bound;
            beware2 = new LR(bound2.L - NodeDistance, bound2.T - NodeDistance, bound2.R + NodeDistance, bound2.B + NodeDistance);   // :805-811
        }
        else beware2 = bound2 = new LR(pt2.X, pt2.Y, pt2.X, pt2.Y);   // :813-818

        Track? best = null;
        ulong bestQuality = 0xFFFFFFFF;
        EdgeInfo bestInfo = info.Clone();
        bool auto1 = con1.Connected && con1.BestVertex, auto2 = con2.Connected && con2.BestVertex;
        int count1 = auto1 ? 4 : 1, count2 = auto2 ? 4 : 1;
        int bestAuto1 = 0, bestAuto2 = 0;
        for (int num1 = 0; num1 < count1; num1++)
        {
            if (con1.Connected)
            {
                var g = TakeGluePoint(con1, auto1 ? num1 : con1.GlueIndex);
                pt1 = g.Pos; esc1 = g.Escape;
                if (esc1 == OdtGlueEscape.Smart) esc1 = CalcEscAngle(con1.Snap, pt1);
            }
            for (int num2 = 0; num2 < count2; num2++)
            {
                if (con2.Connected)
                {
                    var g = TakeGluePoint(con2, auto2 ? num2 : con2.GlueIndex);
                    pt2 = g.Pos; esc2 = g.Escape;
                    if (esc2 == OdtGlueEscape.Smart) esc2 = CalcEscAngle(con2.Snap, pt2);
                }
                for (long a1 = 0; a1 < 36000; a1 += 9000)
                {
                    var e1 = a1 == 0 ? OdtGlueEscape.Right : a1 == 9000 ? OdtGlueEscape.Top : a1 == 18000 ? OdtGlueEscape.Left : OdtGlueEscape.Bottom;
                    for (long a2 = 0; a2 < 36000; a2 += 9000)
                    {
                        var e2 = a2 == 0 ? OdtGlueEscape.Right : a2 == 9000 ? OdtGlueEscape.Top : a2 == 18000 ? OdtGlueEscape.Left : OdtGlueEscape.Bottom;
                        if ((esc1 & e1) != 0 && (esc2 & e2) != 0)
                        {
                            var tryInfo = info.Clone();
                            var xp = CalcEdgeTrack(kind, pt1, a1, bound1, beware1, pt2, a2, bound2, beware2, out ulong quality, tryInfo);
                            if (quality < bestQuality)
                            {
                                best = xp; bestQuality = quality; bestInfo = tryInfo; bestAuto1 = num1; bestAuto2 = num2;
                            }
                        }
                    }
                }
            }
        }
        if (auto1) con1.GlueIndex = bestAuto1;   // :872-873 (m_bBestVertex stays set: the next layout searches again)
        if (auto2) con2.GlueIndex = bestAuto2;
        CopyInfo(bestInfo, info);
        return best ?? track0;
    }

    private static void CopyInfo(EdgeInfo from, EdgeInfo to)
    {
        to.Obj1Line2 = from.Obj1Line2; to.Obj1Line3 = from.Obj1Line3; to.Obj2Line2 = from.Obj2Line2; to.Obj2Line3 = from.Obj2Line3;
        to.MiddleLine = from.MiddleLine; to.Angle1 = from.Angle1; to.Angle2 = from.Angle2;
        to.Obj1Lines = from.Obj1Lines; to.Obj2Lines = from.Obj2Lines; to.MiddleLineIndex = from.MiddleLineIndex;
    }

    private static long Half(long a, long b) => (a + b + 1) / 2;

    /// <summary>BigMulDiv = BigInt::Scale, rounded (svx/source/svdraw/svdtrans.cxx:551-556, tools/source/generic/bigint.cxx:715-729).</summary>
    private static long BigMulDiv(long value, long mul, long div)
    {
        if (div == 0) return 0x7fffffff;
        decimal v = (decimal)value * mul;
        if ((v < 0) != (div < 0)) v -= div / 2; else v += div / 2;
        return (long)decimal.Truncate(v / div);
    }

    /// <summary>The geometric SdrEdgeObj::ImpCalcEdgeTrack (svdoedge.cxx:880-1606).</summary>
    private static Track CalcEdgeTrack(EdgeKind kind, LP rPt1, long angle1, LR rBound1, LR rBeware1, LP rPt2, long angle2, LR rBound2, LR rBeware2,
        out ulong quality, EdgeInfo info)
    {
        bool rts1 = angle1 == 0, obn1 = angle1 == 9000, lks1 = angle1 == 18000, unt1 = angle1 == 27000;
        bool hor1 = lks1 || rts1, ver1 = obn1 || unt1;
        bool rts2 = angle2 == 0, obn2 = angle2 == 9000, lks2 = angle2 == 18000, unt2 = angle2 == 27000;
        bool hor2 = lks2 || rts2, ver2 = obn2 || unt2;
        info.Angle1 = angle1; info.Angle2 = angle2; info.Obj1Lines = 1; info.Obj2Lines = 1; info.MiddleLineIndex = NoMiddleLine;   // :898-904
        LP pt1 = rPt1, pt2 = rPt2;
        LR bound1 = rBound1, bound2 = rBound2, beware1 = rBeware1, beware2 = rBeware2;
        var meeting = new LP(Half(pt1.X, pt2.X), Half(pt1.Y, pt2.Y));   // :911
        if (kind == EdgeKind.OneLine)   // :912-920
        {
            var t = new Track(); t.Add(rPt1); t.Add(rPt2);
            quality = (ulong)(Math.Abs(rPt1.X - rPt2.X) + Math.Abs(rPt1.Y - rPt2.Y));
            return t;
        }
        if (kind == EdgeKind.ThreeLines)   // :921-955
        {
            var t = new Track(); t.Add(rPt1); t.Add(rPt1); t.Add(rPt2); t.Add(rPt2);
            if (rts1) t.SetX(1, beware1.R);
            if (obn1) t.SetY(1, beware1.T);
            if (lks1) t.SetX(1, beware1.L);
            if (unt1) t.SetY(1, beware1.B);
            if (rts2) t.SetX(2, beware2.R);
            if (obn2) t.SetY(2, beware2.T);
            if (lks2) t.SetX(2, beware2.L);
            if (unt2) t.SetY(2, beware2.B);
            long q = Math.Abs(t[1].X - t[0].X) + Math.Abs(t[1].Y - t[0].Y);
            q += Math.Abs(t[2].X - t[1].X) + Math.Abs(t[2].Y - t[1].Y);
            q += Math.Abs(t[3].X - t[2].X) + Math.Abs(t[3].Y - t[2].Y);
            quality = (ulong)q;
            info.Obj1Lines = 2; info.Obj2Lines = 2;
            if (hor1) t.AdjustX(1, info.Obj1Line2.X); else t.AdjustY(1, info.Obj1Line2.Y);
            if (hor2) t.AdjustX(2, info.Obj2Line2.X); else t.AdjustY(2, info.Obj2Line2.Y);
            return t;
        }
        int intersections = 0;
        {
            // meeting point: :959-987
            LP c1 = beware1.Center, c2 = beware2.Center;
            if (beware1.L <= beware2.R && beware1.R >= beware2.L)
                meeting.X = Half(Math.Max(beware1.L, beware2.L), Math.Min(beware1.R, beware2.R));
            else
                meeting.X = c1.X < c2.X ? Half(beware1.R, beware2.L) : Half(beware1.L, beware2.R);
            if (beware1.T <= beware2.B && beware1.B >= beware2.T)
                meeting.Y = Half(Math.Max(beware1.T, beware2.T), Math.Min(beware1.B, beware2.B));
            else
                meeting.Y = c1.Y < c2.Y ? Half(beware1.B, beware2.T) : Half(beware1.T, beware2.B);
            long xMin = Math.Min(beware1.L, beware2.L), xMax = Math.Max(beware1.R, beware2.R);
            long yMin = Math.Min(beware1.T, beware2.T), yMax = Math.Max(beware1.B, beware2.B);
            bool bewareOverlap = beware1.R > beware2.L && beware1.L < beware2.R && beware1.B > beware2.T && beware1.T < beware2.B;
            int mainCase = 3;
            if (angle1 == angle2) mainCase = 1;
            else if ((hor1 && hor2) || (ver1 && ver2)) mainCase = 2;
            if (mainCase == 1)   // :1001-1024
            {
                if (ver1) meeting.X = Half(pt1.X, pt2.X);
                if (hor1) meeting.Y = Half(pt1.Y, pt2.Y);
                bool x1Ok = pt1.X <= beware2.L || pt1.X >= beware2.R;
                bool x2Ok = pt2.X <= beware1.L || pt2.X >= beware1.R;
                bool y1Ok = pt1.Y <= beware2.T || pt1.Y >= beware2.B;
                bool y2Ok = pt2.Y <= beware1.T || pt2.Y >= beware1.B;
                if (lks1 && (y1Ok || beware1.L < beware2.R) && (y2Ok || beware2.L < beware1.R)) meeting.X = xMin;
                if (rts1 && (y1Ok || beware1.R > beware2.L) && (y2Ok || beware2.R > beware1.L)) meeting.X = xMax;
                if (obn1 && (x1Ok || beware1.T < beware2.B) && (x2Ok || beware2.T < beware1.B)) meeting.Y = yMin;
                if (unt1 && (x1Ok || beware1.B > beware2.T) && (x2Ok || beware2.B > beware1.T)) meeting.Y = yMax;
            }
            else if (mainCase == 2)
            {
                if (hor1)   // :1025-1121
                {
                    LR bewR1 = rts1 ? beware1 : beware2, bewR2 = rts1 ? beware2 : beware1;
                    LR bndR1 = rts1 ? bound1 : bound2, bndR2 = rts1 ? bound2 : bound1;
                    if (bewR1.B > bewR2.T && bewR1.T < bewR2.B && bewR1.R > bewR2.L)
                    {
                        bool case29Direct = false;
                        bool case29 = bewR1.R > bewR2.L;
                        if (bndR1.R <= bndR2.L)
                        {
                            if ((pt1.Y > beware2.T && pt1.Y < beware2.B) || (pt2.Y > beware1.T && pt2.Y < beware1.B)) case29Direct = true;
                        }
                        if (!case29Direct)
                        {
                            bool obenLang = Math.Abs(yMin - meeting.Y) <= Math.Abs(yMax - meeting.Y);
                            meeting.Y = obenLang ? yMin : yMax;
                            if (case29)
                                meeting.X = (bewR1.Center.Y < bewR2.Center.Y) != obenLang ? bewR2.R : bewR1.L;
                        }
                        else
                        {
                            long want1 = bewR1.R - bndR1.R, want2 = bndR2.L - bewR2.L, space = bndR2.L - bndR1.R;
                            long get1 = BigMulDiv(want1, space, want1 + want2), get2 = space - get1;
                            if (rts1) { beware1.R += get1 - want1; beware2.L -= get2 - want2; }
                            else { beware2.R += get1 - want1; beware1.L -= get2 - want2; }
                            intersections++;
                        }
                    }
                }
                else if (ver1)   // :1122-1181
                {
                    LR bewR1 = unt1 ? beware1 : beware2, bewR2 = unt1 ? beware2 : beware1;
                    LR bndR1 = unt1 ? bound1 : bound2, bndR2 = unt1 ? bound2 : bound1;
                    if (bewR1.R > bewR2.L && bewR1.L < bewR2.R && bewR1.B > bewR2.T)
                    {
                        bool case29Direct = false;
                        bool case29 = bewR1.B > bewR2.T;
                        if (bndR1.B <= bndR2.T)
                        {
                            if ((pt1.X > beware2.L && pt1.X < beware2.R) || (pt2.X > beware1.L && pt2.X < beware1.R)) case29Direct = true;
                        }
                        if (!case29Direct)
                        {
                            bool linksLang = Math.Abs(xMin - meeting.X) <= Math.Abs(xMax - meeting.X);
                            meeting.X = linksLang ? xMin : xMax;
                            if (case29)
                                meeting.Y = (bewR1.Center.X < bewR2.Center.X) != linksLang ? bewR2.B : bewR1.T;
                        }
                        else
                        {
                            long want1 = bewR1.B - bndR1.B, want2 = bndR2.T - bewR2.T, space = bndR2.T - bndR1.B;
                            long get1 = BigMulDiv(want1, space, want1 + want2), get2 = space - get1;
                            if (unt1) { beware1.B += get1 - want1; beware2.T -= get2 - want2; }
                            else { beware2.B += get1 - want1; beware1.T -= get2 - want2; }
                            intersections++;
                        }
                    }
                }
            }
            else   // case 3: :1182-1271
            {
                LR tmp1 = beware1, tmp2 = beware2;
                if (bewareOverlap) { tmp1 = bound1; tmp2 = bound2; }
                if ((((rts1 && tmp1.R <= pt2.X) || (lks1 && tmp1.L >= pt2.X)) && ((unt2 && tmp2.B <= pt1.Y) || (obn2 && tmp2.T >= pt1.Y))) ||
                    (((rts2 && tmp2.R <= pt1.X) || (lks2 && tmp2.L >= pt1.X)) && ((unt1 && tmp1.B <= pt2.Y) || (obn1 && tmp1.T >= pt2.Y))))
                {
                    if (hor1) { meeting.X = pt2.X; meeting.Y = pt1.Y; }
                    else { meeting.X = pt1.X; meeting.Y = pt2.Y; }
                    beware1 = tmp1; beware2 = tmp2;
                }
                else if ((((rts1 && beware1.R > beware2.L) || (lks1 && beware1.L < beware2.R)) &&
                          ((unt2 && beware2.B > beware1.T) || (obn2 && beware2.T < beware1.B))) ||
                         (((rts2 && beware2.R > beware1.L) || (lks2 && beware2.L < beware1.R)) &&
                          ((unt1 && beware1.B > beware2.T) || (obn1 && beware1.T < beware2.B))))
                {
                    if (rts1 || rts2) meeting.X = xMax;
                    if (lks1 || lks2) meeting.X = xMin;
                    if (unt1 || unt2) meeting.Y = yMax;
                    if (obn1 || obn2) meeting.Y = yMin;
                }
            }
        }

        // :1274-1324
        var xp1 = CalcObjToCenter(pt1, angle1, beware1, meeting);
        var xp2 = CalcObjToCenter(pt2, angle2, beware2, meeting);
        int n1 = xp1.Count, n2 = xp2.Count;
        info.Obj1Lines = n1 - 1;
        info.Obj2Lines = n2 - 1;
        LP ep1 = xp1[n1 - 1], ep2 = xp2[n2 - 1];
        bool insMeetingPoint = ep1.X != ep2.X && ep1.Y != ep2.Y;
        bool horzE1 = ep1.Y == xp1[n1 - 2].Y, horzE2 = ep2.Y == xp2[n2 - 2].Y;
        if (ep1 == ep2 && ((horzE1 && horzE2 && ep1.Y == ep2.Y) || (!horzE1 && !horzE2 && ep1.X == ep2.X)))
        {
            n1--; xp1.RemoveAt(n1);
            n2--; xp2.RemoveAt(n2);
        }
        if (insMeetingPoint)
        {
            xp1.Add(meeting);
            if (info.Obj1Lines == info.Obj2Lines) { info.Obj1Lines++; info.Obj2Lines++; }
            else if (info.Obj1Lines > info.Obj2Lines) { info.Obj2Lines++; info.MiddleLineIndex = n1 - 1; }
            else { info.Obj1Lines++; info.MiddleLineIndex = n1; }
        }
        else if (ep1 != ep2 && n1 + n2 >= 4)
        {
            info.MiddleLineIndex = n1 - 1;
        }
        int num = xp2.Count;
        if (n1 > 1 && n2 > 1 && xp1[n1 - 1] == xp2[n2 - 1]) num--;
        while (num > 0) { num--; xp1.Add(xp2[num]); }
        int pointCount = xp1.Count;

        // the shape letter: :1326-1366
        char form;
        if (pointCount == 2) form = 'I';
        else if (pointCount == 3) form = 'L';
        else if (pointCount == 4) form = angle1 == angle2 ? 'U' : 'Z';
        else if (pointCount == 6)
        {
            if (angle1 != angle2)
            {
                LP a = xp1[1], b = xp1[2], c = xp1[3], d = xp1[4];
                if (a.Y == b.Y) form = (a.X < b.X) == (c.X < d.X) ? 'S' : 'C';
                else form = (a.Y < b.Y) == (c.Y < d.Y) ? 'S' : 'C';
            }
            else form = '4';
        }
        else form = '?';
        if (form == 'I' || form == 'L' || form == 'Z' || form == 'U')
        {
            info.Obj1Lines = 1; info.Obj2Lines = 1;
            info.MiddleLineIndex = form == 'Z' || form == 'U' ? 1 : NoMiddleLine;
        }
        else if (form == 'S' || form == 'C')
        {
            info.Obj1Lines = 2; info.Obj2Lines = 2; info.MiddleLineIndex = 2;
        }

        // quality: :1371-1459
        {
            ulong q = 0, q0 = q;
            bool overflow = false;
            LP p0 = xp1[0];
            for (int i = 1; i < pointCount; i++)
            {
                LP pb = xp1[i];
                q += (ulong)(Math.Abs(pb.X - p0.X) + Math.Abs(pb.Y - p0.Y));
                if (q < q0) overflow = true;
                q0 = q;
                p0 = pb;
            }
            int tmpCount = pointCount;
            if (form == 'Z')
            {
                tmpCount = 2;
                ulong l1 = (ulong)(Math.Abs(xp1[1].X - xp1[0].X) + Math.Abs(xp1[1].Y - xp1[0].Y));
                ulong l2 = (ulong)(Math.Abs(xp1[2].X - xp1[1].X) + Math.Abs(xp1[2].Y - xp1[1].Y));
                ulong l3 = (ulong)(Math.Abs(xp1[3].X - xp1[2].X) + Math.Abs(xp1[3].Y - xp1[2].Y));
                ulong better = 0;
                l1 += l3;
                l3 = l2 / 4;
                if (l1 >= l2) better = 6;
                else if (l1 >= 3 * l3) better = 4;
                else if (l1 >= 2 * l3) better = 2;
                if (xp1[0].Y != xp1[1].Y) better++;
                if (q > better) q -= better; else q = 0;
            }
            if (tmpCount >= 3)
            {
                q0 = q;
                q += (ulong)tmpCount * 0x01000000;
                if (q < q0 || tmpCount > 15) overflow = true;
            }
            if (pointCount >= 2)
            {
                LP a = new(xp1[1].X - xp1[0].X, xp1[1].Y - xp1[0].Y);
                LP b = new(xp1[pointCount - 2].X - xp1[pointCount - 1].X, xp1[pointCount - 2].Y - xp1[pointCount - 1].Y);
                long ang1 = 0; if (a.X < 0) ang1 = 18000; if (a.Y > 0) ang1 = 27000; if (a.Y < 0) ang1 = 9000; if (a.X != 0 && a.Y != 0) ang1 = 1;
                long ang2 = 0; if (b.X < 0) ang2 = 18000; if (b.Y > 0) ang2 = 27000; if (b.Y < 0) ang2 = 9000; if (b.X != 0 && b.Y != 0) ang2 = 1;
                if (ang1 != angle1) intersections++;
                if (ang2 != angle2) intersections++;
            }
            LR qb1 = rBeware1, qb2 = rBeware2;
            for (int i = 0; i < pointCount; i++)
            {
                LP pb = xp1[i];
                bool b1 = pb.X > qb1.L && pb.X < qb1.R && pb.Y > qb1.T && pb.Y < qb1.B;
                bool b2 = pb.X > qb2.L && pb.X < qb2.R && pb.Y > qb2.T && pb.Y < qb2.B;
                int int0 = intersections;
                if (i == 0 || i == pointCount - 1) { if (b1 && b2) intersections++; }
                else { if (b1) intersections++; if (b2) intersections++; }
                if (i > 0 && int0 == intersections)
                {
                    if (p0.Y == pb.Y)
                    {
                        if (p0.Y > qb1.T && p0.Y < qb1.B && ((p0.X <= qb1.L && pb.X >= qb1.R) || (pb.X <= qb1.L && p0.X >= qb1.R))) intersections++;
                        if (p0.Y > qb2.T && p0.Y < qb2.B && ((p0.X <= qb2.L && pb.X >= qb2.R) || (pb.X <= qb2.L && p0.X >= qb2.R))) intersections++;
                    }
                    else
                    {
                        if (p0.X > qb1.L && p0.X < qb1.R && ((p0.Y <= qb1.T && pb.Y >= qb1.B) || (pb.Y <= qb1.T && p0.Y >= qb1.B))) intersections++;
                        if (p0.X > qb2.L && p0.X < qb2.R && ((p0.Y <= qb2.T && pb.Y >= qb2.B) || (pb.Y <= qb2.T && p0.Y >= qb2.B))) intersections++;
                    }
                }
                p0 = pb;
            }
            if (pointCount <= 1) intersections++;
            q0 = q;
            q += (ulong)intersections * 0x10000000;
            if (q < q0 || intersections > 15) overflow = true;
            if (overflow || q == 0xFFFFFFFF) q = 0xFFFFFFFE;
            quality = q;
        }

        var track = new Track();
        foreach (var p in xp1) track.Add(p);

        // line offsets: :1460-1515
        ApplyOffset(info, track, LineCode.MiddleLine, info.MiddleLineIndex != NoMiddleLine);
        ApplyOffset(info, track, LineCode.Obj1Line2, info.Obj1Lines >= 2);
        ApplyOffset(info, track, LineCode.Obj1Line3, info.Obj1Lines >= 3);
        ApplyOffset(info, track, LineCode.Obj2Line2, info.Obj2Lines >= 2);
        ApplyOffset(info, track, LineCode.Obj2Line3, info.Obj2Lines >= 3);

        // Bezier: :1516-1604 (the OOXML routing branch :1517-1545 is off for ODF: m_bUseOOXMLCurve defaults false, svdoedge.hxx:106)
        if (kind != EdgeKind.Bezier || pointCount <= 2) return track;
        long dx1 = track[1].X - track[0].X, dy1 = track[1].Y - track[0].Y;
        long dx2 = track[pointCount - 2].X - track[pointCount - 1].X, dy2 = track[pointCount - 2].Y - track[pointCount - 1].Y;
        if (form == 'L')
        {
            track.F[1] = FlagControl;
            track.Insert(2, track[1], FlagControl);
            pointCount = track.Count;
            track.AdjustX(1, -(dx1 / 3)); track.AdjustY(1, -(dy1 / 3));
            track.AdjustX(pointCount - 2, -(dx2 / 3)); track.AdjustY(pointCount - 2, -(dy2 / 3));
        }
        else if (pointCount >= 4 && pointCount <= 6)
        {
            track.F[1] = FlagControl;
            track.F[pointCount - 2] = FlagControl;
            track.AdjustX(1, dx1 / 2); track.AdjustY(1, dy1 / 2);
            track.AdjustX(pointCount - 2, dx2 / 2); track.AdjustY(pointCount - 2, dy2 / 2);
            if (pointCount == 5)
            {
                LP center = track[2];
                long dx1b = center.X - track[1].X, dy1b = center.Y - track[1].Y;
                long dx2b = center.X - track[3].X, dy2b = center.Y - track[3].Y;
                track.Insert(2, center, FlagControl);
                track.F[3] = FlagSymmetric;
                track.Insert(4, center, FlagControl);
                track.AdjustX(2, -(dx1b / 2)); track.AdjustY(2, -(dy1b / 2));
                track.AdjustX(3, -((dx1b + dx2b) / 4)); track.AdjustY(3, -((dy1b + dy2b) / 4));
                track.AdjustX(4, -(dx2b / 2)); track.AdjustY(4, -(dy2b / 2));
            }
            if (pointCount == 6)
            {
                LP a = track[2], b = track[3];
                track.Insert(2, a, FlagControl);
                track.Insert(5, b, FlagControl);
                long dx = a.X - b.X, dy = a.Y - b.Y;
                track.AdjustX(3, -(dx / 2)); track.AdjustY(3, -(dy / 2));
                track.F[3] = FlagSymmetric;
                track.RemoveAt(4);
            }
        }
        return track;
    }

    private static void ApplyOffset(EdgeInfo info, Track track, LineCode code, bool applies)
    {
        if (!applies) return;
        int index = info.PolyIndex(code, track.Count);
        LP offset = info.Offset(code);
        if (info.IsHorizontalLine(code, track.Count)) { track.AdjustY(index, offset.Y); track.AdjustY(index + 1, offset.Y); }
        else { track.AdjustX(index, offset.X); track.AdjustX(index + 1, offset.X); }
    }
}
