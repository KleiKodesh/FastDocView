using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.Odt;

/// <summary>One outline object of a custom shape to extrude: its flattened rings (box points), whether it is filled, its line colour.</summary>
internal sealed class OdtExtrusionSource
{
    public readonly List<(List<(double X, double Y)> Points, bool Closed)> Rings = new();
    public bool Filled = true;
}

/// <summary>A projected face of an extruded shape: its rings in the shape's box points (y down), its lit colour (null = the shape's own
/// fill, a bitmap or gradient front / back face), whether it is drawn as an outline only (draft shade mode), and its depth in the scene.</summary>
internal sealed class OdtExtrusionFace
{
    public readonly List<List<(double X, double Y)>> Rings = new();
    public string? Color;
    public bool OutlineOnly;
    public double Depth;
    public string Kind = "";
}

/// <summary>
/// A custom shape's 3-D extrusion (ODF 1.3 part 3 draw:extrusion and its attributes on draw:enhanced-geometry), drawn as LibreOffice draws it
/// (master read 2026-11-04): EnhancedCustomShape3d::Create3DObject (svx/source/customshapes/EnhancedCustomShape3d.cxx:181-1090) builds a
/// 3-D scene of extruded outlines which the drawinglayer 3-D processor renders (svx/source/sdr/contact/viewcontactofe3dscene.cxx:152-249,
/// drawinglayer/source/processor3d/defaultprocessor3d.cxx:330-450). Here the scene becomes 2-D polygons, back to front, each one flat colour:
/// <list type="bullet">
/// <item>attributes as xmloff/source/draw/ximpcustomshape.cxx:920-1060 reads them (percentages only with '%', GetDoublePercentage :136-155;
///   pairs by GetNextParameter, depth = length + fraction :934-974; vectors "(x y z)"; draw:projection present and not "parallel" =
///   perspective :1062-1072, absent = parallel, Create3DObject :343);</item>
/// <item>defaults (Create3DObject helpers :38-160): depth 1270 (1/100 mm) fraction 0, origin (0.5, -0.5), rotation 0 0, skew 50 at -135
///   degrees, viewpoint (3472, -3472, 25000), lights (50000, 0, 10000) and (-50000, 0, 10000) at 66 %, both harsh, brightness 33 %,
///   diffusion 0, specularity 0, shininess 50, shade mode flat;</item>
/// <item>geometry (extrud3d.cxx:53-71 mirrors the outline in y; sdrextrudelathetools3d.cxx:449-520 puts the back at z = 0 and the front at
///   z = depth): every outline object is one extruded body (:412-617) - front and back faces, one side face per edge;</item>
/// <item>scene transform (:643-709): translate (-centre x, centre y, -backward depth), the shape's rotation about z, mirroring, then the
///   rotation angles about the rotation centre (y first, then x, both negated), then for parallel projection the skew
///   shearXY(skew/100 cos a, skew/100 sin a); basegfx applies each operation after the ones before (hommatrixtemplate.hxx:373-391,
///   b3dhommatrix.cxx:125-230);</item>
/// <item>projection: parallel drops z; perspective divides as :796-806 with the viewpoint shifted by the origin (:716-749). The scene is
///   then fitted to the projected bounding volume's rectangle (:772-816), which is the same box the renderer maps its projected content
///   range onto (viewcontactofe3dscene.cxx:179-249), so a vertex lands at its projected point plus the shape's centre (y back down);</item>
/// <item>faces turned away are not drawn (defaultprocessor3d.cxx:348-359) unless double-sided (two fill styles or one-axis mirroring,
///   Create3DObject :399-403, which also turns on two-sided lighting: a back face's normal is reversed, :367-378);</item>
/// <item>flat shading (:418-431): one colour per face from SdrLightingAttribute::solveColorModel (sdrlightingattribute3d.cxx:116-168) with
///   the face's normal transformed by the scene matrix' linear part (b3dvector.cxx:60-77, not its inverse transpose), ambient
///   brightness, light 1 (specular) and 2, soft lights 5 - 8 for non-harsh lights (:870-951), light 1 moved to non-specular light 4 when
///   specularity is 0 (:1072-1080), diffusion scaling the material's HSB brightness (:1038-1051, tools color.cxx:95-193), metal
///   (:995-1068); light colours are 8-bit (tools color.hxx:106-111).</item>
/// </list>
/// Not done - each needs a real 3-D renderer: Gouraud and Phong shading (shade-mode gouraud / phong: per-vertex colours with smoothed
/// normals, defaultprocessor3d.cxx:383-417, interpolated per pixel by the z-buffer raster converter) are drawn with the flat per-face
/// colour; overlapping faces are ordered by their mean depth (painter's order) where LibreOffice resolves them per pixel with a z-buffer
/// (rasterconvertB3DPolyPolygon, defaultprocessor3d.cxx:449), which differs only for concave outlines whose side faces overlap; a bitmap
/// or gradient front face is drawn with the shape's own fill unlit, where LibreOffice maps it as a texture modulated by the light
/// (Svx3DTextureModeItem, Create3DObject :372); a wide or dashed line-only outline is extruded as its centre line, not as its contour
/// (ConvertToContourObj :455-491).
/// </summary>
internal sealed class OdtExtrusion
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw;
    private const string LoExt = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
    /// <summary>One LibreOffice core unit (1/100 mm) in points.</summary>
    internal const double Hundredth = 72 / 2540.0;

    public double Depth = 1270 * Hundredth, DepthFraction;
    public bool Perspective;
    public double RotateXDegrees, RotateYDegrees;
    /// <summary>draw:extrusion-rotation-center: x and y as fractions of the box, z in 1/100 mm (Create3DObject :677-684).</summary>
    public (double X, double Y, double Z) RotationCenter;
    public double SkewAmount = 50, SkewAngleDegrees = -135;
    /// <summary>In points.</summary>
    public (double X, double Y, double Z) ViewPoint = (3472 * Hundredth, -3472 * Hundredth, 25000 * Hundredth);
    public double OriginX = 0.50, OriginY = -0.50;
    public string ShadeMode = "flat";
    public bool UseExtrusionColor;
    public (double X, double Y, double Z) FirstLightDirection = (50000, 0, 10000), SecondLightDirection = (-50000, 0, 10000);
    public double FirstLightLevel = 66, SecondLightLevel = 66;
    public bool FirstLightHarsh = true, SecondLightHarsh = true;
    public double Brightness = 33, Diffusion, Specularity, Shininess = 50;
    public bool Metal, MetalMicrosoftCompatible;

    /// <summary>The extrusion of a draw:enhanced-geometry, or null when it has no draw:extrusion="true".</summary>
    public static OdtExtrusion? Read(XElement? geometry)
    {
        if (geometry is null || (string?)geometry.Attribute(Draw + "extrusion") != "true") return null;
        var e = new OdtExtrusion();
        // LibreOffice matches these attributes by local name only, whatever the prefix (EASGet(getNameFromToken),
        // xmloff/source/draw/EnhancedCustomShapeToken.cxx); ODF puts projection / shade-mode in dr3d:, LibreOffice files use draw:
        string? A(string name) => (string?)geometry.Attribute(Draw + name) ?? (string?)geometry.Attribute(XNamespace.Get("urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0") + name);
        var noNames = new Dictionary<string, int>();
        bool Pair(string? value, out double first, out double second)
        {
            first = second = 0;
            if (value is null) return false;
            int index = 0;
            if (!OdtEnhancedGeometry.NextParameter(value, ref index, noNames, out var a) || !OdtEnhancedGeometry.NextParameter(value, ref index, noNames, out var b)) return false;
            first = a.Value; second = b.Value;
            return true;
        }
        void Percent(string name, ref double target) { if (PercentValue(A(name)) is { } v) target = v; }
        void Bool(string name, ref bool target) { if (A(name) is { } v) target = v == "true"; }
        void Vector(string name, ref (double, double, double) target) { if (Vector3(A(name), null) is { } v) target = v; }

        if (A("extrusion-depth") is { } depthValue)
        {
            // :934-974 - the depth with its unit (none = 1/100 mm), then the fraction; both or nothing
            string[] parts = depthValue.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2 && OdtShapeStyle.Length(parts[0]) is { } d && double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out double f))
            { e.Depth = d; e.DepthFraction = f; }
        }
        Percent("extrusion-brightness", ref e.Brightness);
        Percent("extrusion-diffusion", ref e.Diffusion);
        Percent("extrusion-first-light-level", ref e.FirstLightLevel);
        Percent("extrusion-second-light-level", ref e.SecondLightLevel);
        Percent("extrusion-shininess", ref e.Shininess);
        // loext:extrusion-specularity wins over draw:extrusion-specularity (:1050-1055)
        if (PercentValue((string?)geometry.Attribute(XNamespace.Get(LoExt) + "extrusion-specularity") ?? A("extrusion-specularity")) is { } specularity) e.Specularity = specularity;
        Bool("extrusion-first-light-harsh", ref e.FirstLightHarsh);
        Bool("extrusion-second-light-harsh", ref e.SecondLightHarsh);
        Bool("extrusion-metal", ref e.Metal);
        Bool("extrusion-color", ref e.UseExtrusionColor);
        e.MetalMicrosoftCompatible = ((string?)geometry.Attribute(XNamespace.Get(LoExt) + "extrusion-metal-type") ?? A("extrusion-metal-type")) == "loext:MetalMSCompatible";
        Vector("extrusion-first-light-direction", ref e.FirstLightDirection);
        Vector("extrusion-second-light-direction", ref e.SecondLightDirection);
        Vector("extrusion-rotation-center", ref e.RotationCenter);
        if (Vector3(A("extrusion-viewpoint"), OdtShapeStyle.Length) is { } viewPoint) e.ViewPoint = viewPoint;
        if (A("shade-mode") is { } shade) e.ShadeMode = shade is "phong" or "gouraud" or "draft" ? shade : "flat";
        if (A("projection") is { } projection) e.Perspective = projection != "parallel";
        if (Pair(A("extrusion-rotation-angle"), out double rx, out double ry)) { e.RotateXDegrees = rx; e.RotateYDegrees = ry; }
        if (Pair(A("extrusion-skew"), out double skew, out double skewAngle)) { e.SkewAmount = skew; e.SkewAngleDegrees = skewAngle; }
        if (Pair(A("extrusion-origin"), out double ox, out double oy)) { e.OriginX = ox; e.OriginY = oy; }
        return e;
    }

    /// <summary>GetDoublePercentage (ximpcustomshape.cxx:136-155): only a value with '%' counts; the number itself (33% = 33).</summary>
    private static double? PercentValue(string? value)
    {
        if (value is null) return null;
        string v = value.Trim();
        if (!v.EndsWith("%", StringComparison.Ordinal)) return null;
        return double.TryParse(v.Substring(0, v.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d : null;
    }

    /// <summary>"(x y z)" (convertB3DVector / convertPosition3D); <paramref name="length"/> converts measures, else plain numbers.</summary>
    private static (double, double, double)? Vector3(string? value, Func<string?, double?>? length)
    {
        if (value is null) return null;
        string v = value.Trim();
        if (v.Length < 2 || v[0] != '(' || v[v.Length - 1] != ')') return null;
        string[] parts = v.Substring(1, v.Length - 2).Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 3) return null;
        var n = new double[3];
        for (int i = 0; i < 3; i++)
        {
            double? d = length is not null ? length(parts[i])
                : double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double plain) ? plain : null;
            if (d is null) return null;
            n[i] = d.Value;
        }
        return (n[0], n[1], n[2]);
    }

    // ---- 3-D maths (basegfx conventions: column vectors, each operation applied after the ones before) ----

    private struct Matrix
    {
        public double M00, M01, M02, M03, M10, M11, M12, M13, M20, M21, M22, M23;
        public static Matrix Identity => new() { M00 = 1, M11 = 1, M22 = 1 };
        /// <summary>this = <paramref name="m"/> * this (doMulMatrix, hommatrixtemplate.hxx:373-391).</summary>
        public void Then(Matrix m)
        {
            Matrix a = this;
            M00 = m.M00 * a.M00 + m.M01 * a.M10 + m.M02 * a.M20; M01 = m.M00 * a.M01 + m.M01 * a.M11 + m.M02 * a.M21; M02 = m.M00 * a.M02 + m.M01 * a.M12 + m.M02 * a.M22; M03 = m.M00 * a.M03 + m.M01 * a.M13 + m.M02 * a.M23 + m.M03;
            M10 = m.M10 * a.M00 + m.M11 * a.M10 + m.M12 * a.M20; M11 = m.M10 * a.M01 + m.M11 * a.M11 + m.M12 * a.M21; M12 = m.M10 * a.M02 + m.M11 * a.M12 + m.M12 * a.M22; M13 = m.M10 * a.M03 + m.M11 * a.M13 + m.M12 * a.M23 + m.M13;
            M20 = m.M20 * a.M00 + m.M21 * a.M10 + m.M22 * a.M20; M21 = m.M20 * a.M01 + m.M21 * a.M11 + m.M22 * a.M21; M22 = m.M20 * a.M02 + m.M21 * a.M12 + m.M22 * a.M22; M23 = m.M20 * a.M03 + m.M21 * a.M13 + m.M22 * a.M23 + m.M23;
        }
        public void Translate(double x, double y, double z) => Then(new Matrix { M00 = 1, M11 = 1, M22 = 1, M03 = x, M13 = y, M23 = z });
        public void Scale(double x, double y, double z) => Then(new Matrix { M00 = x, M11 = y, M22 = z });
        /// <summary>b3dhommatrix.cxx:125-171: x, then y, then z.</summary>
        public void Rotate(double ax, double ay, double az)
        {
            if (ax != 0) { double s = Math.Sin(ax), c = Math.Cos(ax); Then(new Matrix { M00 = 1, M11 = c, M22 = c, M21 = s, M12 = -s }); }
            if (ay != 0) { double s = Math.Sin(ay), c = Math.Cos(ay); Then(new Matrix { M00 = c, M22 = c, M02 = s, M20 = -s, M11 = 1 }); }
            if (az != 0) { double s = Math.Sin(az), c = Math.Cos(az); Then(new Matrix { M00 = c, M11 = c, M10 = s, M01 = -s, M22 = 1 }); }
        }
        /// <summary>b3dhommatrix.cxx:218-230.</summary>
        public void ShearXY(double sx, double sy) => Then(new Matrix { M00 = 1, M11 = 1, M22 = 1, M02 = sx, M12 = sy });
        public (double X, double Y, double Z) Point((double X, double Y, double Z) p)
            => (M00 * p.X + M01 * p.Y + M02 * p.Z + M03, M10 * p.X + M11 * p.Y + M12 * p.Z + M13, M20 * p.X + M21 * p.Y + M22 * p.Z + M23);
        /// <summary>B3DVector *= matrix: the linear part only (b3dvector.cxx:60-77).</summary>
        public (double X, double Y, double Z) Vector((double X, double Y, double Z) v)
            => (M00 * v.X + M01 * v.Y + M02 * v.Z, M10 * v.X + M11 * v.Y + M12 * v.Z, M20 * v.X + M21 * v.Y + M22 * v.Z);
    }

    private static (double X, double Y, double Z) Normalized((double X, double Y, double Z) v)
    {
        double length = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
        return length == 0 ? v : (v.X / length, v.Y / length, v.Z / length);
    }

    private static double Dot((double X, double Y, double Z) a, (double X, double Y, double Z) b) => a.X * b.X + a.Y * b.Y + a.Z * b.Z;

    /// <summary>A colour through 8 bits (Color(BColor), color.hxx:106-111) and back (getBColor).</summary>
    private static double Byte(double value) => Math.Round(Math.Max(0, Math.Min(1, value)) * 255.0) / 255.0;

    private sealed record Light((double X, double Y, double Z) Direction, double Level, bool Specular);

    private sealed class Face
    {
        public string Kind = "";
        public readonly List<List<(double X, double Y, double Z)>> Rings = new();   // object coordinates
        public List<(double X, double Y, double Z)> Facing = new();                  // a ring wound so its normal is Normal
        public (double X, double Y, double Z) Normal;                                  // outward, object coordinates
        public (double R, double G, double B) Material;
        public bool DoubleSided, Textured;
    }

    /// <summary>
    /// The projected faces of the extruded <paramref name="sources"/> of a <paramref name="width"/> x <paramref name="height"/> box,
    /// back to front. <paramref name="rotationDegrees"/> is the shape's rotation (counter-clockwise, draw:transform rotate), which
    /// LibreOffice applies inside the scene (Create3DObject :653-658); the faces are around the box's centre and the caller places
    /// that centre without rotating or mirroring again (EnhancedCustomShapeEngine.cxx:262-266).
    /// </summary>
    public List<OdtExtrusionFace> Project(IReadOnlyList<OdtExtrusionSource> sources, double width, double height, double rotationDegrees,
                                          bool mirrorX, bool mirrorY, (double R, double G, double B) fill, (double R, double G, double B) secondaryFill,
                                          (double R, double G, double B) line, bool fillTextured)
    {
        var result = new List<OdtExtrusionFace>();
        // :337-341
        double forward = Depth * DepthFraction, backward = Depth - forward;
        double depth = Math.Max(backward + forward, Hundredth);
        bool draft = ShadeMode == "draft";
        // :379-404
        // an unfilled shape never gets two fill styles: else-if after the FillStyle_NONE branch (:386-389)
        bool shapeFilled = sources.Any(s => s.Filled);
        bool twoFillStyles = !draft && shapeFilled && (fillTextured || UseExtrusionColor);
        bool doubleSidedScene = draft || twoFillStyles || mirrorX != mirrorY;
        bool twoSidedLighting = !draft && (twoFillStyles || mirrorX != mirrorY);
        double cx = width / 2, cy = height / 2;

        // material colours before diffusion (:581-583, :607-614)
        var faces = new List<Face>();
        double minX = double.MaxValue, minY = double.MaxValue, maxX = double.MinValue, maxY = double.MinValue;
        foreach (OdtExtrusionSource source in sources)
        {
            var rings = source.Rings.Where(r => r.Points.Count >= 2).ToList();
            if (rings.Count == 0) continue;
            foreach (var ring in rings) foreach (var p in ring.Points) { minX = Math.Min(minX, p.X); maxX = Math.Max(maxX, p.X); minY = Math.Min(minY, p.Y); maxY = Math.Max(maxY, p.Y); }
            bool caps = source.Filled;   // draft keeps the object's fill style (:378-383)
            var capColour = source.Filled ? fill : line;
            var sideColour = !source.Filled ? line : UseExtrusionColor ? secondaryFill : fill;
            // 3-D object coordinates: y mirrored (extrud3d.cxx:61-64), back at z = 0, front at z = depth
            (double, double, double) At((double X, double Y) p, double z) => (p.X, -p.Y, z);
            if (caps)
            {
                var closed = rings.Where(r => r.Points.Count >= 3).ToList();
                if (closed.Count > 0)
                {
                    // the outer ring (not inside another) wound counter-clockwise in object coordinates gives +z
                    var outer = closed.FirstOrDefault(r => !IsHole(r.Points, closed.Select(c => c.Points).ToList())).Points ?? closed[0].Points;
                    bool ccw = SignedArea(outer) > 0;
                    foreach (bool front in new[] { false, true })
                    {
                        var face = new Face { Kind = front ? "front" : "back", Normal = (0, 0, front ? 1 : -1), Material = capColour,
                                              DoubleSided = doubleSidedScene || !source.Filled, Textured = fillTextured && source.Filled };
                        double z = front ? depth : 0;
                        foreach (var r in closed) face.Rings.Add(r.Points.Select(p => At(p, z)).ToList());
                        var facing = outer.Select(p => At(p, z)).ToList();
                        if (ccw != front) facing.Reverse();
                        face.Facing = facing;
                        faces.Add(face);
                    }
                }
            }
            var allPoints = rings.Select(r => r.Points).ToList();
            foreach (var ring in rings)
            {
                var pts = ring.Points;
                bool ringCcw = SignedArea(pts) > 0;
                bool hole = ring.Closed && IsHole(pts, allPoints);
                int edges = ring.Closed ? pts.Count : pts.Count - 1;
                for (int i = 0; i < edges; i++)
                {
                    var a = pts[i]; var b = pts[(i + 1) % pts.Count];
                    if (a == b) continue;
                    // edge direction in object coordinates (y mirrored); outward = right of the edge for a counter-clockwise outer ring
                    double ex = b.X - a.X, ey = -(b.Y - a.Y);
                    bool outwardRight = ringCcw != hole;
                    var normal = Normalized(outwardRight ? (ey, -ex, 0) : (-ey, ex, 0));
                    var quad = new List<(double X, double Y, double Z)> { At(a, 0), At(b, 0), At(b, depth), At(a, depth) };
                    if (!outwardRight) quad.Reverse();
                    var face = new Face { Kind = "side", Normal = normal, Material = sideColour, DoubleSided = doubleSidedScene || !source.Filled, Facing = quad };
                    face.Rings.Add(quad);
                    faces.Add(face);
                }
            }
        }
        if (faces.Count == 0) return result;

        // scene transform (:643-709)
        Matrix m = Matrix.Identity;
        m.Translate(-cx, cy, -backward);
        double zRotate = rotationDegrees * Math.PI / 180;
        if (zRotate != 0) m.Rotate(0, 0, zRotate);
        if (mirrorX) m.Scale(-1, 1, 1);
        if (mirrorY) m.Scale(1, -1, 1);
        double rcx = RotationCenter.X * width, rcy = RotationCenter.Y * height, rcz = RotationCenter.Z * Hundredth;
        m.Translate(-rcx, rcy, -rcz);
        double yRotate = RotateYDegrees * Math.PI / 180, xRotate = RotateXDegrees * Math.PI / 180;
        if (yRotate != 0) m.Rotate(0, -yRotate, 0);
        if (xRotate != 0) m.Rotate(-xRotate, 0, 0);
        m.Translate(rcx, -rcy, rcz);
        if (!Perspective && SkewAmount != 0)
        {
            double alpha = SkewAngleDegrees * Math.PI / 180, invTanBeta = SkewAmount / 100.0;
            m.ShearXY(invTanBeta * Math.Cos(alpha), invTanBeta * Math.Sin(alpha));
        }

        // perspective viewpoint (:713-749): the origin is relative to the outline's bounding box after the 2-D rotation and mirroring
        double vx = 0, vy = 0, vz = 25000 * Hundredth;
        if (Perspective)
        {
            // the range of the transformed outline itself, aTotalPolyPoly.getB2DRange() (:719-725), not of its box's corners
            var corners = sources.SelectMany(src => src.Rings).Where(r => r.Points.Count >= 2).SelectMany(r => r.Points).Select(p =>
            {
                // aPolyPolyTransform (:651-669): translate(-c), rotate(-z) (y down), mirror, translate(c)
                double x = p.X - cx, y = p.Y - cy;
                double c = Math.Cos(-zRotate), s = Math.Sin(-zRotate);
                (x, y) = (c * x - s * y, s * x + c * y);
                if (mirrorX) x = -x;
                if (mirrorY) y = -y;
                return (X: x + cx, Y: y + cy);
            }).ToList();
            double bl = corners.Min(p => p.X), br = corners.Max(p => p.X), bt = corners.Min(p => p.Y), bb = corners.Max(p => p.Y);
            double originX = OriginX * (br - bl) + (bl + br) / 2 - cx;
            double originY = -(OriginY * (bb - bt) + (bt + bb) / 2) + cy;
            vx = ViewPoint.X + originX;
            vy = -ViewPoint.Y + originY;
            vz = ViewPoint.Z;
        }
        // a point whose line to the viewpoint is parallel to the xy plane has no projection (division by zero). LibreOffice skips such a
        // point only for the 2-D bound volume's corners (EnhancedCustomShape3d.cxx:797-798); its 3-D renderer draws the faces. Here the
        // point is left out of its ring, and a ring or face left with fewer than 3 points is not drawn (our own guard).
        (double X, double Y)? Project2((double X, double Y, double Z) p)
        {
            var q = m.Point(p);
            if (!Perspective) return (q.X, q.Y);
            double div = q.Z - vz;
            if (div == 0) return null;
            double f = -vz / div;
            return ((q.X - vx) * f + vx, (q.Y - vy) * f + vy);
        }

        // lights (:822-1080)
        var lights = Lights(mirrorX != mirrorY, out double ambient, out (double R, double G, double B) specularColour, out int specularIntensity);

        foreach (Face face in faces)
        {
            var projectedFacing = face.Facing.Select(Project2).OfType<(double X, double Y)>().ToList();
            if (projectedFacing.Count < 3) continue;
            // y up here: counter-clockwise = the face's normal points at the viewer
            bool facesViewer = SignedArea(projectedFacing, q => q) > 0;
            if (!face.DoubleSided && !facesViewer) continue;
            var outFace = new OdtExtrusionFace { Kind = face.Kind, OutlineOnly = draft };
            var kept = new List<(double X, double Y, double Z)>();
            foreach (var ring in face.Rings)
            {
                var projected = ring.Where(p => Project2(p) is not null).ToList();
                if (projected.Count < 3) continue;
                kept.AddRange(projected);
                outFace.Rings.Add(projected.Select(p => Project2(p)!.Value).Select(p => (p.X + cx, -p.Y + cy)).ToList());
            }
            if (outFace.Rings.Count == 0) continue;
            outFace.Depth = kept.Average(p => m.Point(p).Z);
            if (draft) outFace.Color = Hex(line);
            else if (face.Textured) outFace.Color = null;
            else
            {
                var normal = m.Vector(face.Normal);
                if (twoSidedLighting && !facesViewer) normal = (-normal.X, -normal.Y, -normal.Z);
                var material = DiffusedMaterial(face.Material);
                var specular = MetalMicrosoftCompatible && Metal ? MetalSpecular(face.Material) : specularColour;
                outFace.Color = Hex(Solve(normal, material, specular, specularIntensity, ambient, lights));
            }
            result.Add(outFace);
        }
        // painter's order: farthest (smallest z, the viewer looks down -z) first; stable for equal depths
        return result.Select((f, i) => (f, i)).OrderBy(t => t.f.Depth).ThenBy(t => t.i).Select(t => t.f).ToList();
    }

    /// <summary>The scene's lights (Create3DObject :822-1080) in the order createNewSdrLightingAttribute lists them (lights 1 - 8,
    /// sdrattributecreator.cxx:1063-1125); only light 1 is specular.</summary>
    private List<Light> Lights(bool singleMirror, out double ambient, out (double R, double G, double B) specularColour, out int specularIntensity)
    {
        (double X, double Y, double Z) Direction((double X, double Y, double Z) d)
        {
            if (d.X == 0 && d.Y == 0 && d.Z == 0) d.Z = 1;
            return Normalized((d.X, -d.Y, d.Z));
        }
        var light1 = Direction(FirstLightDirection);
        var light2 = Direction(SecondLightDirection);
        if (singleMirror) { light1 = (-light1.X, -light1.Y, -light1.Z); light2 = (-light2.X, -light2.Y, -light2.Z); }
        double level1 = FirstLightLevel / 100.0, level2 = SecondLightLevel / 100.0;
        ambient = Brightness / 100.0;
        double level1ForSpecular = level1;
        var soft = new (double X, double Y, double Z)[4];       // lights 5, 6, 7, 8
        double vertical = 0, horizontal = 0;
        bool softLights = false;
        if (!FirstLightHarsh || !SecondLightHarsh)
        {
            if (!SecondLightHarsh && level2 > 0 && (FirstLightHarsh || level1 == 0))
            {
                softLights = true;
                horizontal = vertical = Byte(level2 * 0.40);
                level2 *= 0.2;
                SoftLightDirections(light2, out soft[0], out soft[1], out soft[2], out soft[3]);
            }
            else if (!FirstLightHarsh && level1 > 0 && (SecondLightHarsh || level2 == 0))
            {
                softLights = true;
                horizontal = vertical = Byte(level1 * 0.40);
                level1 *= 0.2;
                SoftLightDirections(light1, out soft[0], out soft[1], out soft[2], out soft[3]);
            }
            else if (!FirstLightHarsh && level1 > 0 && !SecondLightHarsh && level2 > 0)
            {
                softLights = true;
                horizontal = Byte(level1 * 0.8);
                level1 *= 0.4;
                SoftLightDirections(light1, out _, out _, out soft[2], out soft[3]);
                vertical = Byte(level2 * 0.8);
                level2 *= 0.4;
                SoftLightDirections(light2, out soft[0], out soft[1], out _, out _);
            }
        }
        if (level1 > 1.0) ambient += (level1 - 1.0) / 2.0;
        ambient = Byte(ambient);

        double specularity = Specularity / 100.0;
        if (Metal && !MetalMicrosoftCompatible) specularity *= 200.0 / 255.0;
        double shadingFactor = level1ForSpecular * specularity;
        double specular = Byte(shadingFactor);
        specularColour = (specular, specular, specular);
        _shadingFactor = shadingFactor;
        double shininess = Math.Max(0, Math.Min(100, Math.Pow(2, Shininess / 10.0)));
        int intensity = (int)Math.Round(shininess, MidpointRounding.AwayFromZero);
        if (Metal && !MetalMicrosoftCompatible) intensity = Math.Max(0, Math.Min(100, intensity + 15));
        specularIntensity = intensity;

        var lights = new List<Light>();
        // light 1 is the specular one; with no specularity it is replaced by the same light as non-specular light 4 (:1072-1080)
        bool noSpecularity = Math.Abs(specularity) < 1e-9;      // fTools::equalZero
        if (level1 > 0 && !noSpecularity) lights.Add(new Light(light1, Byte(level1), true));
        if (level2 > 0) lights.Add(new Light(light2, Byte(level2), false));
        if (noSpecularity) lights.Add(new Light(light1, Byte(level1), false));
        if (softLights)
        {
            lights.Add(new Light(soft[0], vertical, false));
            lights.Add(new Light(soft[1], vertical, false));
            lights.Add(new Light(soft[2], horizontal, false));
            lights.Add(new Light(soft[3], horizontal, false));
        }
        return lights;
    }

    private double _shadingFactor;

    /// <summary>lcl_SoftLightsDirection (Create3DObject :131-176).</summary>
    private static void SoftLightDirections((double X, double Y, double Z) light, out (double X, double Y, double Z) up, out (double X, double Y, double Z) down,
                                            out (double X, double Y, double Z) right, out (double X, double Y, double Z) left)
    {
        double angle = 60 * Math.PI / 180;
        up = (0, Math.Sin(angle), Math.Cos(angle));
        down = (0, -Math.Sin(angle), Math.Cos(angle));
        right = (Math.Sin(angle), 0, Math.Cos(angle));
        left = (-Math.Sin(angle), 0, Math.Cos(angle));
        Matrix rotate = Matrix.Identity;
        rotate.Rotate(0, 0, Math.PI / 4);
        if (light.X == 0 && light.Z == 0) rotate.Rotate(light.Y >= 0 ? -Math.PI / 2 : Math.PI / 2, 0, 0);
        else
        {
            double azimuth = Math.Atan2(light.X, light.Z);
            double elevation = Math.Atan2(light.Y, Math.Sqrt(light.X * light.X + light.Z * light.Z));
            rotate.Rotate(-elevation, azimuth, 0);
        }
        up = rotate.Vector(up); down = rotate.Vector(down); right = rotate.Vector(right); left = rotate.Vector(left);
    }

    /// <summary>Diffusion (Create3DObject :1038-1051): the material's HSB brightness times the diffusion, unless 0 or 1.</summary>
    private (double R, double G, double B) DiffusedMaterial((double R, double G, double B) colour)
    {
        double diffusion = Diffusion / 100.0;
        if (!(diffusion > 0 && Math.Abs(diffusion) >= 1e-9 && Math.Abs(diffusion - 1) >= 1e-9)) return colour;
        (int h, int s, int b) = ToHsb(colour);
        b = Math.Max(0, Math.Min(100, (int)(b * diffusion)));
        return FromHsb(h, s, b);
    }

    /// <summary>MS-compatible metal (:1054-1064): the specular colour is the material at its brightness times the shading factor.</summary>
    private (double R, double G, double B) MetalSpecular((double R, double G, double B) colour)
    {
        (int h, int s, int b) = ToHsb(colour);
        b = Math.Max(0, Math.Min(100, (int)(b * _shadingFactor)));
        return FromHsb(h, s, b);
    }

    /// <summary>Color::RGBtoHSB (tools/source/generic/color.cxx:95-152), integer arithmetic.</summary>
    internal static (int H, int S, int B) ToHsb((double R, double G, double B) colour)
    {
        int r = (int)Math.Round(colour.R * 255), g = (int)Math.Round(colour.G * 255), bl = (int)Math.Round(colour.B * 255);
        int max = Math.Max(r, Math.Max(g, bl)), min = Math.Min(r, Math.Min(g, bl));
        int bri = max * 100 / 255;
        int delta = max - min;
        int sat = bri > 0 ? delta * 100 / max : 0;
        int hue = 0;
        if (sat != 0)
        {
            double dHue = r == max ? (double)(g - bl) / delta : g == max ? 2.0 + (double)(bl - r) / delta : 4.0 + (double)(r - g) / delta;
            dHue *= 60.0;
            if (dHue < 0) dHue += 360.0;
            hue = (int)dHue;
        }
        return (hue, sat, bri);
    }

    /// <summary>Color::HSBtoRGB (color.cxx:154-193).</summary>
    internal static (double R, double G, double B) FromHsb(int hue, int sat, int bri)
    {
        int nB = (byte)(bri * 255 / 100);
        int r, g, b;
        if (sat == 0) r = g = b = nB;
        else
        {
            double dH = hue;
            if (dH == 360.0) dH = 0;
            dH /= 60.0;
            int n = (int)dH;
            double f = dH - n;
            int a = (byte)(nB * (100 - sat) / 100);
            int bb = (byte)(nB * (100 - sat * f) / 100);
            int c = (byte)(nB * (100 - sat * (1.0 - f)) / 100);
            (r, g, b) = n switch { 0 => (nB, c, a), 1 => (bb, nB, a), 2 => (a, nB, c), 3 => (a, bb, nB), 4 => (c, a, nB), 5 => (nB, a, bb), _ => (0, 0, 0) };
        }
        return (r / 255.0, g / 255.0, b / 255.0);
    }

    /// <summary>SdrLightingAttribute::solveColorModel (sdrlightingattribute3d.cxx:116-168), emission black (svdattr.cxx:374).</summary>
    private static (double R, double G, double B) Solve((double X, double Y, double Z) eyeNormal, (double R, double G, double B) colour,
                                                        (double R, double G, double B) specular, int specularIntensity, double ambient, List<Light> lights)
    {
        double r = ambient * colour.R, g = ambient * colour.G, b = ambient * colour.B;
        if (lights.Count > 0 && !(eyeNormal.X == 0 && eyeNormal.Y == 0 && eyeNormal.Z == 0))
        {
            var n = Normalized(eyeNormal);
            foreach (Light light in lights)
            {
                double cos = Dot(light.Direction, n);
                if (cos <= 0) continue;
                r += light.Level * colour.R * cos; g += light.Level * colour.G * cos; b += light.Level * colour.B * cos;
                if (!light.Specular) continue;
                var half = Normalized((light.Direction.X, light.Direction.Y, light.Direction.Z + 1.0));
                double cos2 = Dot(half, n);
                if (cos2 <= 0) continue;
                cos2 = Math.Pow(cos2, specularIntensity);
                r += specular.R * cos2; g += specular.G * cos2; b += specular.B * cos2;
            }
        }
        return (Math.Max(0, Math.Min(1, r)), Math.Max(0, Math.Min(1, g)), Math.Max(0, Math.Min(1, b)));
    }

    private static string Hex((double R, double G, double B) c)
        => "#" + (((int)Math.Round(c.R * 255) << 16) | ((int)Math.Round(c.G * 255) << 8) | (int)Math.Round(c.B * 255)).ToString("x6", CultureInfo.InvariantCulture);

    /// <summary>Signed area, y as given (positive = counter-clockwise in a y-up frame).</summary>
    private static double SignedArea<T>(List<T> points, Func<T, (double X, double Y)> xy)
    {
        double area = 0;
        for (int i = 0; i < points.Count; i++)
        {
            var a = xy(points[i]); var b = xy(points[(i + 1) % points.Count]);
            area += a.X * b.Y - b.X * a.Y;
        }
        return area / 2;
    }

    /// <summary>Box points (y down) as object coordinates (y up): positive = counter-clockwise there.</summary>
    private static double SignedArea(List<(double X, double Y)> boxPoints) => SignedArea(boxPoints, p => (p.X, -p.Y));

    private static bool IsHole(List<(double X, double Y)> ring, List<List<(double X, double Y)>> all)
    {
        if (ring.Count == 0) return false;
        var probe = ring[0];
        int inside = 0;
        foreach (var other in all)
            if (!ReferenceEquals(other, ring) && other.Count >= 3 && Contains(other, probe)) inside++;
        return (inside & 1) == 1;
    }

    private static bool Contains(List<(double X, double Y)> polygon, (double X, double Y) p)
    {
        bool inside = false;
        for (int i = 0, j = polygon.Count - 1; i < polygon.Count; j = i++)
        {
            var a = polygon[i]; var b = polygon[j];
            if ((a.Y > p.Y) != (b.Y > p.Y) && p.X < (b.X - a.X) * (p.Y - a.Y) / (b.Y - a.Y) + a.X) inside = !inside;
        }
        return inside;
    }
}
