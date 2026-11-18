using System.Globalization;
using System.Xml.Linq;
using FastDocView.Odt;

namespace FastDocView.Html;

/// <summary>
/// 3-D shapes: a shape's DrawingML scene (a:scene3d camera and light rig, a:sp3d extrusion, bevels, material) or its OfficeArt 3-D
/// properties (.doc / RTF c3D*) as the <see cref="OdtExtrusion"/> scene LibreOffice builds for them, drawn by the same projection
/// as an ODF extrusion. Sources (LibreOffice master read 2026-11-17):
/// <list type="bullet">
/// <item>DrawingML: oox/source/drawingml/scene3dhelper.cxx - the camera presets (aPrstCameraValuesArray :69-160: angles, origin,
///   skew, view point), the angle conversion (getAPIAnglesFromOOXAngle :175-226, getAPIAnglesFrom3DProperties :228-286), depth and
///   z position (addExtrusionDepthToMap :299-327), the projection (addProjectionGeometryToMap :329-376: oblique skew angles minus the
///   shape rotation, perspective distance 15976 / tan(fov / 2) 1/100 mm), when a scene is made (setExtrusionProperties :378-444)
///   and the materials (aPrstMaterialArray :940-960). LibreOffice draws no bevel; here a bevel is its own faces (Word is the
///   reference) and makes a scene even for a front camera.</item>
/// <item>Lights: Word's own rigs ([MS-OI29500] values, as scene3dhelper.cxx :475-640 lists them; directions turned into the shape's
///   frame by V = (-Ly, -Lx, Lz), :697-716, and by the rig's dir / rot, :760-810). Measured on Word 16's PDF of
///   test-data/word-shapes-3d: a face is its colour times the sum of each light's level times the cosine to it (threePt on a front
///   face: 1.043, as Word's pictures measure), plus a small lift by material; the lights stay with the shape whatever the camera
///   (LibreOffice turns them with the shape for modern cameras too, :860-870; lights kept with the viewer measured worse).</item>
/// <item>OfficeArt (.doc / RTF): filter/source/msfilter/msdffimp.cxx:1675-1960 (f3D in fc3DLightFace bit 8; depth forward /
///   backward, light levels and directions, harsh flags, rotation angles and centre, skew, view point, origin, projection
///   fc3DParallel, extrusion colour as the secondary fill colour; MS Office defaults where ODF's differ). LibreOffice's RTF import
///   reads no 3-D properties (sw/source/writerfilter/rtftok/rtfsdrimport.cxx); the RTF names are the [MS-ODRAW] property names
///   Word writes.</item>
/// </list>
/// </summary>
internal static class WordShape3D
{
    private readonly record struct Camera(string Name, bool Parallel, double LatX, double LonY, double RevZ, double OriginX, double OriginY,
                                          double SkewAmount, double SkewAngle, double ViewX, double ViewY, double ViewZ);

    // scene3dhelper.cxx:69-160 (angles in 1/60000 degree, view point in 1/100 mm); index 20-37 are the legacy cameras
    private static readonly Camera[] Cameras =
    [
        new("isometricBottomDown", true, 2124000, 18882000, 17988000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricBottomUp", true, 2124000, 2718000, 3612000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricLeftDown", true, 2100000, 2700000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricLeftUp", true, 19500000, 2700000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis1Left", true, 1080000, 3840000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis1Right", true, 1080000, 20040000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis1Top", true, 18078000, 18390000, 3456000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis2Left", true, 1080000, 1560000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis2Right", true, 1080000, 17760000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis2Top", true, 18078000, 3210000, 18144000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis3Bottom", true, 3522000, 18390000, 18144000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis3Left", true, 20520000, 3840000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis3Right", true, 20520000, 20040000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis4Bottom", true, 3522000, 3210000, 3456000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis4Left", true, 20520000, 1560000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricOffAxis4Right", true, 20520000, 17760000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricRightDown", true, 19500000, 18900000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricRightUp", true, 2100000, 18900000, 0, 0, 0, 0, 0, 0, 0, 0),
        new("isometricTopDown", true, 19476000, 2718000, 17988000, 0, 0, 0, 0, 0, 0, 0),
        new("isometricTopUp", true, 19476000, 18882000, 3612000, 0, 0, 0, 0, 0, 0, 0),
        new("legacyObliqueBottom", true, 0, 0, 0, 0, 0.5, 50, 90, 0, 0, 0),
        new("legacyObliqueBottomLeft", true, 0, 0, 0, -0.5, 0.5, 50, 45, 0, 0, 0),
        new("legacyObliqueBottomRight", true, 0, 0, 0, 0.5, 0.5, 50, 135, 0, 0, 0),
        new("legacyObliqueFront", true, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        new("legacyObliqueLeft", true, 0, 0, 0, -0.5, 0, 50, -360, 0, 0, 0),
        new("legacyObliqueRight", true, 0, 0, 0, 0.5, 0, 50, 180, 0, 0, 0),
        new("legacyObliqueTop", true, 0, 0, 0, 0, -0.5, 50, -90, 0, 0, 0),
        new("legacyObliqueTopLeft", true, 0, 0, 0, -0.5, -0.5, 50, -45, 0, 0, 0),
        new("legacyObliqueTopRight", true, 0, 0, 0, 0.5, -0.5, 50, -135, 0, 0, 0),
        new("legacyPerspectiveBottom", false, 0, 0, 0, 0, 0.5, 50, 90, 0, 3472, 25000),
        new("legacyPerspectiveBottomLeft", false, 0, 0, 0, -0.5, 0.5, 50, 45, -3472, 3472, 25000),
        new("legacyPerspectiveBottomRight", false, 0, 0, 0, 0.5, 0.5, 50, 135, 3472, 3472, 25000),
        new("legacyPerspectiveFront", false, 0, 0, 0, 0, 0, 0, 0, 0, 0, 25000),
        new("legacyPerspectiveLeft", false, 0, 0, 0, -0.5, 0, 50, -360, -3472, 0, 25000),
        new("legacyPerspectiveRight", false, 0, 0, 0, 0.5, 0, 50, 180, 3472, 0, 25000),
        new("legacyPerspectiveTop", false, 0, 0, 0, 0, -0.5, 50, -90, 0, -3472, 25000),
        new("legacyPerspectiveTopLeft", false, 0, 0, 0, -0.5, -0.5, 50, -45, -3472, -3472, 25000),
        new("legacyPerspectiveTopRight", false, 0, 0, 0, 0.5, -0.5, 50, -135, 3472, -3472, 25000),
        new("obliqueBottom", true, 0, 0, 0, 0, 0.5, 30, 90, 0, 0, 0),
        new("obliqueBottomLeft", true, 0, 0, 0, -0.5, 0.5, 30, 45, 0, 0, 0),
        new("obliqueBottomRight", true, 0, 0, 0, 0.5, 0.5, 30, 135, 0, 0, 0),
        new("obliqueLeft", true, 0, 0, 0, -0.5, 0, 30, -360, 0, 0, 0),
        new("obliqueRight", true, 0, 0, 0, 0.5, 0, 30, 180, 0, 0, 0),
        new("obliqueTop", true, 0, 0, 0, 0, -0.5, 30, -90, 0, 0, 0),
        new("obliqueTopLeft", true, 0, 0, 0, -0.5, -0.5, 30, -45, 0, 0, 0),
        new("obliqueTopRight", true, 0, 0, 0, 0.5, -0.5, 30, -135, 0, 0, 0),
        new("orthographicFront", true, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        new("perspectiveAbove", false, 20400000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveAboveLeftFacing", false, 2358000, 858000, 20466000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveAboveRightFacing", false, 2358000, 20742000, 1134000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveBelow", false, 1200000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveContrastingLeftFacing", false, 624000, 2634000, 21384000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveContrastingRightFacing", false, 624000, 18966000, 216000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveFront", false, 0, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveHeroicExtremeLeftFacing", false, 486000, 2070000, 21426000, 0, 0, 0, 0, 0, 0, 18981),
        new("perspectiveHeroicExtremeRightFacing", false, 486000, 19530000, 174000, 0, 0, 0, 0, 0, 0, 18981),
        new("perspectiveHeroicLeftFacing", false, 20940000, 858000, 156000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveHeroicRightFacing", false, 20940000, 20742000, 21444000, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveLeft", false, 0, 1200000, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveRelaxed", false, 18576000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveRelaxedModerately", false, 19488000, 0, 0, 0, 0, 0, 0, 0, 0, 38451),
        new("perspectiveRight", false, 0, 20400000, 0, 0, 0, 0, 0, 0, 0, 38451),
    ];

    /// <summary>One light of a rig: its RGB level and its direction ([MS-OI29500] frame).</summary>
    private readonly record struct RigLight(double R, double G, double B, double X, double Y, double Z);

    // scene3dhelper.cxx:475-640: ambient (null = none) and the lights of each preset rig
    private static readonly Dictionary<string, (double Ambient, RigLight[] Lights)> Rigs = new()
    {
        ["balanced"] = (0.13, [new(1.05, 1.05, 1.05, 0.5263, -0.4092, -0.7453), new(1, 1, 1, -0.9386, 0.3426, -0.041), new(0.5, 0.5, 0.5, 0.0934, 0.763, 0.6396)]),
        ["brightRoom"] = (1.5, [new(1, 1, 1, 0, -1, 0), new(1, 1, 1, 0.8227, -0.1882, -0.5364), new(-0.5, -0.5, -0.5, 0, 0, -1), new(0.5, 0.5, 0.5, 0, 1, 0)]),
        ["chilly"] = (0.11, [new(0.31, 0.32, 0.32, 0.6574, -0.7316, -0.1806), new(0.45, 0.45, 0.45, -0.3539, -0.1505, -0.9231), new(1.03, 1.02, 1.15, 0.672, -0.6185, -0.4073), new(0.41, 0.45, 0.48, -0.5781, 0.7976, 0.1722)]),
        ["contrasting"] = (1, [new(1, 1, 1, 0, -1, 0), new(1, 1, 1, 0, 1, 0)]),
        ["flat"] = (1, [new(0.821, 0.821, 0.821, -0.9546, -0.1619, -0.2502), new(2.072, 2.54, 2.91, 0.0009, 0.8605, 0.5095), new(3.843, 3.843, 3.843, 0.6574, -0.7316, -0.1806)]),
        ["flood"] = (0.13, [new(1.1, 1.1, 1.1, 0.5685, -0.7651, -0.3022), new(1.1, 1.1, 1.1, -0.2366, -0.9595, -0.1531), new(0.55, 0.55, 0.55, -0.8982, 0.1386, -0.4171)]),
        ["freezing"] = (0, [new(0.53, 0.567, 0.661, 0.6574, -0.7316, -0.1806), new(0.37, 0.461, 0.461, -0.2781, -0.4509, -0.8482), new(0.649, 0.638, 0.904, 0.672, -0.6185, -0.4073), new(0.971, 1.19, 1.363, -0.1825, 0.968, 0.1722)]),
        ["glow"] = (1, [new(1, 1, 1, 0, -1, 0), new(0.7, 0.7, 0.7, 0, 1, 0)]),
        ["harsh"] = (0.28, [new(0.88, 0.88, 0.88, 0.6689, -0.6755, -0.3104), new(0.88, 0.88, 0.88, -0.592, -0.7371, -0.326)]),
        ["legacyFlat1"] = (0.305, [new(0.58, 0.58, 0.58, 0, 0, -0.2), new(0.58, 0.58, 0.58, 0, 0, -0.2)]),
        ["legacyFlat2"] = (0.305, [new(0.58, 0.58, 0.58, -1, -1, -0.2), new(0.58, 0.58, 0.58, 0, 1, -0.2)]),
        ["legacyFlat3"] = (0.305, [new(0.58, 0.58, 0.58, 0, -1, -0.2), new(0.58, 0.58, 0.58, 0, 1, -0.2)]),
        ["legacyFlat4"] = (0.305, [new(0.58, 0.58, 0.58, 1, -1, -0.2), new(0.58, 0.58, 0.58, 0, 1, -0.2)]),
        ["legacyHarsh1"] = (0.061, [new(0.793, 0.793, 0.793, 0, 0, -0.2), new(0.214, 0.214, 0.214, 0, 0, -0.2)]),
        ["legacyHarsh2"] = (0.061, [new(0.793, 0.793, 0.793, -1, -1, -0.2), new(0.214, 0.214, 0.214, 0, 1, -0.2)]),
        ["legacyHarsh3"] = (0.061, [new(0.793, 0.793, 0.793, 0, -1, -0.2), new(0.214, 0.214, 0.214, 0, 1, -0.2)]),
        ["legacyHarsh4"] = (0.061, [new(0.793, 0.793, 0.793, 1, -1, -0.2), new(0.214, 0.214, 0.214, 0, 1, -0.2)]),
        ["legacyNormal1"] = (0.153, [new(0.671, 0.671, 0.671, 0, 0, -0.2), new(0.366, 0.366, 0.366, 0, 0, -0.2)]),
        ["legacyNormal2"] = (0.153, [new(0.671, 0.671, 0.671, -1, -1, -0.2), new(0.366, 0.366, 0.366, 0, 1, -0.2)]),
        ["legacyNormal3"] = (0.153, [new(0.671, 0.671, 0.671, 0, -1, -0.2), new(0.366, 0.366, 0.366, 0, 1, -0.2)]),
        ["legacyNormal4"] = (0.153, [new(0.671, 0.671, 0.671, 1, -1, -0.2), new(0.366, 0.366, 0.366, 0, 1, -0.2)]),
        ["morning"] = (0, [new(0.669, 0.648, 0.596, 0.6574, -0.7316, -0.1806), new(0.459, 0.454, 0.385, -0.2781, -0.4509, -0.8482), new(0.9, 0.86, 0.83, 0.672, -0.6185, -0.4073), new(0.911, 0.846, 0.728, -0.1825, 0.968, 0.1722)]),
        ["soft"] = (0.3, [new(0.8, 0.8, 0.8, -0.6897, 0.2484, -0.6802)]),
        ["sunrise"] = (0, [new(0.667, 0.63, 0.527, 0.6574, -0.7316, -0.1806), new(0.459, 0.459, 0.371, -0.2781, -0.4509, -0.8482), new(0.826, 0.712, 0.638, 0.672, -0.6185, -0.4073), new(1.511, 1.319, 0.994, -0.1825, 0.968, 0.1722)]),
        ["sunset"] = (0, [new(0.672, 0.169, 0.169, 0.6574, -0.7316, -0.1806), new(0.459, 0.448, 0.327, 0.0922, -0.3551, -0.9303), new(0.775, 0.612, 0.502, 0.672, -0.6185, -0.4073), new(0.761, 0.69, 0.397, -0.424, 0.8891, 0.1722)]),
        ["threePt"] = (0, [new(1.141, 1.141, 1.141, -0.6515, -0.2693, -0.7093), new(0.5, 0.5, 0.5, 0.8482, 0.2469, -0.4686), new(1, 1, 1, 0.5634, -0.2812, 0.7769)]),
        ["twoPt"] = (0.25, [new(0.84, 0.84, 0.84, 0.5266, -0.4089, -0.7454), new(0.3, 0.3, 0.3, -0.8983, 0.2365, -0.3704)]),
    };

    /// <summary>
    /// Per material (ST_PresetMaterialType): the diffuse share (LibreOffice's aPrstMaterialArray :940-960 fDiffusion), a constant lift
    /// in levels, the Phong highlight's strength in levels and its power (LibreOffice's nSpecularPower), fitted on Word 16's PDFs:
    /// warm matte's highlight (power 8, strength 75.7) reproduces the measured picture lifts by camera angle (test-data/word-picture-
    /// effects effects-light-*-v1: RMS 0.5 levels over 17 angles) and the turned shape of shapes-3d-v1 (19 vs 18.5); matte and legacy
    /// matte have none; legacy metal two thirds diffuse and a flat 10. Materials not measured take warm matte's.
    /// </summary>
    private static (double Diffuse, double Lift, double Highlight, double Power) Material(string material) => material switch
    {
        "matte" or "legacyMatte" => (1, 0, 0, 1),
        "legacyMetal" => (0.66699, 10, 0, 1),
        "dkEdge" => (0.7, 0, 75.7, 8),
        _ => (1, 0, 75.7, 8),
    };

    /// <summary>
    /// The scene of a DrawingML shape's a:spPr (a:scene3d + a:sp3d), or null when it has none or draws flat. <paramref name="rotationDegrees"/>
    /// is the shape's own rotation (clockwise); <paramref name="sceneRotation"/> returns the rotation the scene applies instead
    /// (counter-clockwise, the shape's rotation included). <paramref name="color"/> resolves a colour element's parent to "#rrggbb".
    /// </summary>
    public static OdtExtrusion? FromDrawingML(XElement? properties, double rotationDegrees, Func<XElement?, string?> color, out double sceneRotation,
                                              out string? extrusionColor)
    {
        sceneRotation = -rotationDegrees;
        extrusionColor = null;
        XElement? scene = properties?.Elements().FirstOrDefault(e => e.Name.LocalName == "scene3d");
        XElement? shape3d = properties?.Elements().FirstOrDefault(e => e.Name.LocalName == "sp3d");
        XElement? camera = scene?.Elements().FirstOrDefault(e => e.Name.LocalName == "camera");
        if (scene is null || camera is null) return null;
        int index = Array.FindIndex(Cameras, c => c.Name == (string?)camera.Attribute("prst"));
        if (index < 0) return null;
        Camera preset = Cameras[index];
        double Emu(XElement? e, string name) => double.TryParse((string?)e?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0;
        double Number(XElement? e, string name, double fallback) => double.TryParse((string?)e?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : fallback;

        // angles (getAPIAnglesFrom3DProperties :228-286)
        XElement? rot = camera.Elements().FirstOrDefault(e => e.Name.LocalName == "rot");
        double shapeRotation = rotationDegrees * 60000;
        double lat = Number(rot, "lat", 0), lon = Number(rot, "lon", 0), rev = Number(rot, "rev", 0);
        bool legacy = index is >= 20 and <= 37;
        if (index is >= 29 and <= 37) rev = 0;
        else if (index >= 47 && shapeRotation != 0 && lat > 5400000 && lat <= 16200000) rev += 10800000;
        if (rot?.Attribute("rev") is null) { lat = preset.LatX; lon = preset.LonY; rev = preset.RevZ; }
        if (!legacy) rev -= shapeRotation;
        (double fx, double fy, double fz) = Angles(lat, lon, rev);
        if (legacy) fz -= shapeRotation / 60000 * Math.PI / 180;
        sceneRotation = fz * 180 / Math.PI;

        double depth = Emu(shape3d, "extrusionH"), z = Emu(shape3d, "z");
        WordShapeBevel? Bevel(string name)
        {
            XElement? b = shape3d?.Elements().FirstOrDefault(e => e.Name.LocalName == name);
            if (b is null) return null;
            double w = Number(b, "w", 76200) / 12700, h = Number(b, "h", 76200) / 12700;
            string prst = (string?)b.Attribute("prst") ?? "circle";
            return w > 0 && h > 0 ? new WordShapeBevel(w, h, prst) : null;
        }
        WordShapeBevel? top = Bevel("bevelT"), bottom = Bevel("bevelB");
        bool frontType = preset.Name is "legacyObliqueFront" or "orthographicFront";
        bool turned = Math.Abs(fx) > 1e-9 || Math.Abs(fy) > 1e-9;
        bool create = turned || (!preset.Parallel && depth > 0) || (preset.Parallel && !frontType);
        if (!create && top is null && bottom is null) return null;

        var e = new OdtExtrusion
        {
            Perspective = !preset.Parallel,
            RotateXDegrees = fx * 180 / Math.PI, RotateYDegrees = fy * 180 / Math.PI,
            OriginX = preset.OriginX, OriginY = preset.OriginY,
            ShadeMode = "flat",
            BevelTop = top, BevelBottom = bottom,
        };
        if (preset.Parallel)
        {
            e.SkewAmount = preset.SkewAmount;
            e.SkewAngleDegrees = preset.SkewAngle - (index is >= 38 and <= 45 ? shapeRotation / 60000 : 0);
        }
        else
        {
            double viewZ = preset.ViewZ;
            if (camera.Attribute("fov") is not null)
            {
                double fov = Math.Max(0.5, Math.Min(179.5, Number(camera, "fov", 2700000) / 60000));
                viewZ = 15976.0 / Math.Tan(fov / 2 * Math.PI / 180);
            }
            e.ViewPoint = (preset.ViewX * OdtExtrusion.Hundredth, preset.ViewY * OdtExtrusion.Hundredth, viewZ * OdtExtrusion.Hundredth);
        }
        // depth and z position (addExtrusionDepthToMap :299-327); a zero depth with a bevel stays a flat slab
        double depthPt = depth / 12700, zPt = z / 12700;
        if (depthPt == 0) { e.Depth = OdtExtrusion.Hundredth; e.DepthFraction = zPt / OdtExtrusion.Hundredth; if (zPt == 0) e.DepthFraction = 0; }
        else { e.Depth = depthPt; e.DepthFraction = zPt / depthPt; }
        if (shape3d?.Elements().FirstOrDefault(x => x.Name.LocalName == "extrusionClr") is { } extrusion && color(extrusion) is { } extrusionCss)
        {
            e.UseExtrusionColor = true;
            extrusionColor = extrusionCss;
        }
        string material = (string?)shape3d?.Attribute("prstMaterial") ?? "warmMatte";
        e.ContourWidth = Emu(shape3d, "contourW") / 12700;
        if (shape3d?.Elements().FirstOrDefault(x => x.Name.LocalName == "contourClr") is { } contourClr) e.ContourColor = color(contourClr);
        if (material == "legacyWireframe") e.ShadeMode = "draft";
        XElement? rig = scene.Elements().FirstOrDefault(x => x.Name.LocalName == "lightRig");
        // the viewer's direction in the shape's frame (x right, y down, z to the viewer): the camera turned by lon about y after lat about x
        double latRad = lat / 60000 * Math.PI / 180, lonRad = lon / 60000 * Math.PI / 180;
        e.Shader = Shader(rig, material, (Math.Sin(lonRad) * Math.Cos(latRad), -Math.Sin(latRad), Math.Cos(lonRad) * Math.Cos(latRad)));
        return e;
    }

    /// <summary>getAPIAnglesFromOOXAngle (:175-226): Word's lat / lon / rev (y, then x, then z) as LibreOffice's z, y, x order, radians.</summary>
    private static (double X, double Y, double Z) Angles(double lat, double lon, double rev)
    {
        double a = lat / 60000 * Math.PI / 180, b = lon / 60000 * Math.PI / 180, c = rev / 60000 * Math.PI / 180;
        var x = new double[3, 3] { { 1, 0, 0 }, { 0, Math.Cos(a), Math.Sin(a) }, { 0, -Math.Sin(a), Math.Cos(a) } };
        var y = new double[3, 3] { { Math.Cos(b), 0, -Math.Sin(b) }, { 0, 1, 0 }, { Math.Sin(b), 0, Math.Cos(b) } };
        var zm = new double[3, 3] { { Math.Cos(c), Math.Sin(c), 0 }, { -Math.Sin(c), Math.Cos(c), 0 }, { 0, 0, 1 } };
        double[,] Multiply(double[,] p, double[,] q)
        {
            var r = new double[3, 3];
            for (int i = 0; i < 3; i++) for (int j = 0; j < 3; j++) for (int k = 0; k < 3; k++) r[i, j] += p[i, k] * q[k, j];
            return r;
        }
        var t = Multiply(Multiply(zm, x), y);
        double fy = Math.Asin(Math.Max(-1, Math.Min(1, -t[0, 2]))), fx, fz;
        if (Math.Abs(Math.Cos(fy)) < 1e-9) { fx = 0; fz = Math.Atan2(t[2, 1], t[1, 1]); }
        else { fx = Math.Atan2(-t[1, 2] / Math.Cos(fy), t[2, 2] / Math.Cos(fy)); fz = Math.Atan2(t[0, 1] / Math.Cos(fy), t[0, 0] / Math.Cos(fy)); }
        return (fx, fy, fz);
    }

    /// <summary>
    /// Word's lighting of a face: the rig's ambient plus each light's level times the cosine between the face's outward normal and
    /// the direction the light comes from, times the face's colour and the material's diffuse share, plus the material's lift; a flat
    /// material is not lit. The lights stay with the shape whatever the camera: the cosine uses the face's own normal.
    /// </summary>
    private static Func<(double X, double Y, double Z), (double X, double Y, double Z), (double R, double G, double B), (double R, double G, double B)> Shader(
        XElement? rig, string material, (double X, double Y, double Z) view)
    {
        string name = (string?)rig?.Attribute("rig") ?? "threePt";
        if (!Rigs.TryGetValue(name, out var preset)) preset = Rigs["threePt"];
        // lcl_AdaptAndNormalizeLightDirections (:697-716): V = (-Ly, -Lx, Lz) in the shape's frame (x right, y down, z to the viewer)
        var lights = preset.Lights.Select(l =>
        {
            double x = -l.Y, y = -l.X, z = l.Z, length = Math.Sqrt(x * x + y * y + z * z);
            return (Level: (l.R, l.G, l.B), Direction: (X: x / length, Y: y / length, Z: z / length));
        }).ToList();
        // the rig's rotation (lcl_IncorporateRigRotationIntoLightDirections :760-810): dir in 45 degree steps about z
        double turn = ((string?)rig?.Attribute("dir") ?? "t") switch { "tr" => 45, "r" => 90, "br" => 135, "b" => 180, "bl" => -135, "l" => -90, "tl" => -45, _ => 0 };
        if (turn != 0)
        {
            double c = Math.Cos(turn * Math.PI / 180), s = Math.Sin(turn * Math.PI / 180);
            lights = lights.Select(l => (l.Level, Direction: (X: c * l.Direction.X - s * l.Direction.Y, Y: s * l.Direction.X + c * l.Direction.Y, l.Direction.Z))).ToList();
        }
        bool flat = material == "flat";
        (double diffuse, double liftLevels, double highlight, double power) = Material(material);
        // a lit material also takes half the rig's ambient as white (balanced, ambient 0.13: 24.5 levels on a front face where the
        // highlight gives 7.6)
        double lift = (liftLevels + (highlight > 0 ? preset.Ambient * 255 / 2 : 0)) / 255;
        return (own, _, colour) =>
        {
            if (flat) return (Clamp(colour.R + 5 / 255.0), Clamp(colour.G + 5 / 255.0), Clamp(colour.B + 5 / 255.0));   // Word's print: unlit, 5 levels up
            // the lights stay with the shape: its own normal, whatever the camera
            var n = own;
            double r = preset.Ambient, g = preset.Ambient, b = preset.Ambient, sr = 0, sg = 0, sb = 0;
            foreach (var light in lights)
            {
                // the light travels along Direction; a face lit from it faces the other way
                double cos = -(n.X * light.Direction.X + n.Y * light.Direction.Y + n.Z * light.Direction.Z);
                if (cos <= 0) continue;
                r += light.Level.R * cos; g += light.Level.G * cos; b += light.Level.B * cos;
                // Phong: the light reflected about the normal, against the viewer
                double rx = 2 * cos * n.X + light.Direction.X, ry = 2 * cos * n.Y + light.Direction.Y, rz = 2 * cos * n.Z + light.Direction.Z;
                double spot = rx * view.X + ry * view.Y + rz * view.Z;
                if (highlight > 0 && spot > 0) { double h = Math.Pow(spot, power) * highlight / 255; sr += light.Level.R * h; sg += light.Level.G * h; sb += light.Level.B * h; }
            }
            return (Clamp(colour.R * r * diffuse + lift + sr), Clamp(colour.G * g * diffuse + lift + sg), Clamp(colour.B * b * diffuse + lift + sb));
        };
    }

    private static double Clamp(double v) => Math.Max(0, Math.Min(1, v));

    /// <summary>
    /// The scene of an OfficeArt shape's 3-D properties (<paramref name="property"/>: the value of a property id, null = not set), or
    /// null when f3D is off (msdffimp.cxx:1675-1960).
    /// </summary>
    public static OdtExtrusion? FromOfficeArt(Func<int, uint?> property, out uint? extrusionColor)
    {
        extrusionColor = null;
        uint lightFace = property(703) ?? 0, style = property(767) ?? 0;
        // a boolean counts only with its "use" bit (bit + 16); else [MS-ODRAW] 2.3.15.15 / 2.3.16.28's default: f3D off, metal off,
        // parallel on, key light harsh, fill light not harsh, rotation centre not automatic
        bool Flag(uint value, int bit, bool fallback) => (value & (1u << (bit + 16))) != 0 ? (value & (1u << bit)) != 0 : fallback;
        if (!Flag(lightFace, 3, false)) return null;
        bool keyHarsh = Flag(style, 1, true), fillHarsh = Flag(style, 0, false);
        double Fixed(int id, double fallback) => property(id) is { } v ? (int)v / 65536.0 : fallback;
        double Signed(int id, double fallback) => property(id) is { } v ? (int)v : fallback;
        var e = new OdtExtrusion
        {
            Brightness = Signed(722, 20000) / 655.36,
            Diffusion = Signed(641, 65536) / 655.36,
            FirstLightHarsh = keyHarsh, SecondLightHarsh = fillHarsh,
            FirstLightLevel = Signed(726, 38000) / 655.36, SecondLightLevel = Signed(730, 38000) / 655.36,
            Metal = Flag(lightFace, 2, false), MetalMicrosoftCompatible = true,
            ShadeMode = property(713) == 1 ? "draft" : "flat",   // mso_Wireframe
            Perspective = !Flag(style, 2, true),
        };
        if (property(645) is not null || property(644) is not null)
        {
            double back = Signed(645, 1270 * 360) / 360.0, fore = Signed(644, 0) / 360.0, depth = back + fore;
            e.Depth = depth * OdtExtrusion.Hundredth;
            e.DepthFraction = depth != 0 ? fore / depth : 0;
        }
        if (property(723) is not null || property(724) is not null || property(725) is not null)
            e.FirstLightDirection = (Signed(723, 50000), Signed(724, 0), Signed(725, 10000));
        if (property(727) is not null || property(728) is not null || property(729) is not null)
            e.SecondLightDirection = (Signed(727, -50000), Signed(728, 0), Signed(729, 10000));
        if (property(705) is not null || property(704) is not null) { e.RotateXDegrees = Fixed(705, 0); e.RotateYDegrees = Fixed(704, 0); }
        if (!Flag(style, 3, false) && (property(710) is not null || property(711) is not null || property(712) is not null))
            e.RotationCenter = (Fixed(710, 0), Fixed(711, 0), Signed(712, 0) / 360.0);
        if (property(642) is { } shininess) e.Shininess = (int)shininess * 10.0;
        if (property(721) is not null || property(720) is not null) { e.SkewAmount = Signed(721, 50); e.SkewAngleDegrees = Fixed(720, -135); }
        else { e.SkewAmount = 50; e.SkewAngleDegrees = -135; }
        if (property(640) is not null) e.Specularity = Signed(640, 0) / 655.36;
        double viewX = 1250000 / 360.0, viewY = -1250000 / 360.0, viewZ = 9000000 / 360.0;
        if (property(715) is not null || property(716) is not null || property(717) is not null)
        { viewX = Signed(715, 1250000) / 360.0; viewY = Signed(716, -1250000) / 360.0; viewZ = Signed(717, 9000000) / 360.0; }
        e.ViewPoint = (viewX * OdtExtrusion.Hundredth, viewY * OdtExtrusion.Hundredth, viewZ * OdtExtrusion.Hundredth);
        if (property(718) is not null || property(719) is not null) { e.OriginX = Signed(718, 32768) / 65536; e.OriginY = Signed(719, -32768) / 65536; }
        if (property(647) is { } colour) { e.UseExtrusionColor = true; extrusionColor = colour; }
        e.Shader = LegacyShader(e);
        return e;
    }

    /// <summary>
    /// Word's lighting of a legacy (OfficeArt) 3-D shape, fitted on Word 16's print of test-data/word-shapes-3d/shapes-3d-on-v1.rtf
    /// (exact face colours: the PDF draws them as vectors; front 0.766, top 0.594, right side 0.880 of the fill): the ambient
    /// intensity plus each light's level times the cosine to its position for a harsh (directional) light, or times
    /// (1 + cosine) / 2 for a soft one ([MS-ODRAW]: "infinite, uniform, planar"), times the colour; no lift. The lights stay with
    /// the viewer (the shape turns under them, as in LibreOffice).
    /// </summary>
    private static Func<(double X, double Y, double Z), (double X, double Y, double Z), (double R, double G, double B), (double R, double G, double B)> LegacyShader(OdtExtrusion e)
    {
        static (double X, double Y, double Z) Unit((double X, double Y, double Z) v)
        {
            double l = Math.Sqrt(v.X * v.X + v.Y * v.Y + v.Z * v.Z);
            return l == 0 ? (0, 0, 1) : (v.X / l, v.Y / l, v.Z / l);
        }
        var key = Unit(e.FirstLightDirection); var fill = Unit(e.SecondLightDirection);
        double ambient = e.Brightness / 100, keyLevel = e.FirstLightLevel / 100, fillLevel = e.SecondLightLevel / 100, diffuse = e.Diffusion / 100;
        bool keyHarsh = e.FirstLightHarsh, fillHarsh = e.SecondLightHarsh;
        double Light((double X, double Y, double Z) n, (double X, double Y, double Z) l, bool harsh)
        {
            double cos = n.X * l.X + n.Y * l.Y + n.Z * l.Z;
            return harsh ? Math.Max(0, cos) : (1 + cos) / 2;
        }
        return (_, eye, colour) =>
        {
            double k = ambient + keyLevel * Light(eye, key, keyHarsh) + fillLevel * Light(eye, fill, fillHarsh);
            k *= diffuse;
            return (Clamp(colour.R * k), Clamp(colour.G * k), Clamp(colour.B * k));
        };
    }

    /// <summary>A closed ring grown outward by <paramref name="width"/> (each vertex along its two edges' miter, kept to 4 widths).</summary>
    private static List<(double X, double Y)> Grown(List<(double X, double Y)> ring, double width)
    {
        var points = new List<(double X, double Y)>();
        foreach (var p in ring) if (points.Count == 0 || Math.Abs(p.X - points[points.Count - 1].X) + Math.Abs(p.Y - points[points.Count - 1].Y) > 0.01) points.Add(p);
        while (points.Count > 1 && Math.Abs(points[0].X - points[points.Count - 1].X) + Math.Abs(points[0].Y - points[points.Count - 1].Y) <= 0.01) points.RemoveAt(points.Count - 1);
        int n = points.Count;
        if (n < 3) return ring;
        double area = 0;
        for (int i = 0; i < n; i++) area += points[i].X * points[(i + 1) % n].Y - points[(i + 1) % n].X * points[i].Y;
        var outward = new (double X, double Y)[n];
        for (int i = 0; i < n; i++)
        {
            var a = points[i]; var b = points[(i + 1) % n];
            double ex = b.X - a.X, ey = b.Y - a.Y, l = Math.Sqrt(ex * ex + ey * ey);
            // y down: a positive area runs clockwise on screen, its outside is left of each edge
            outward[i] = l == 0 ? (0, 0) : area > 0 ? (ey / l, -ex / l) : (-ey / l, ex / l);
        }
        var grown = new List<(double X, double Y)>(n);
        for (int i = 0; i < n; i++)
        {
            var n1 = outward[(i - 1 + n) % n]; var n2 = outward[i];
            double mx = n1.X + n2.X, my = n1.Y + n2.Y, ml = Math.Sqrt(mx * mx + my * my);
            if (ml < 1e-9) { grown.Add((points[i].X + n2.X * width, points[i].Y + n2.Y * width)); continue; }
            mx /= ml; my /= ml;
            double scale = Math.Min(4, 1 / Math.Max(mx * n2.X + my * n2.Y, 0.25));
            grown.Add((points[i].X + mx * scale * width, points[i].Y + my * scale * width));
        }
        return grown;
    }

    /// <summary>
    /// Replaces <paramref name="shape"/>'s outline (its evaluated <see cref="WordShape.Paths"/>) by the projected faces of
    /// <paramref name="extrusion"/>, each filled with its lit colour, as an ODF extrusion is drawn (OdtShapeBuilder): the scene takes
    /// the rotation (<paramref name="sceneRotation"/>, counter-clockwise) and the flips, the shape is then placed unturned.
    /// </summary>
    public static void Apply(WordShape shape, OdtExtrusion extrusion, double sceneRotation, string? extrusionColor)
    {
        if (shape.Paths is not { Count: > 0 } paths || shape.Geometry == WordShapeGeometry.Picture) return;
        bool filled = shape.FillColor is not null || shape.GradientStops is not null || shape.FillPictureId is not null || shape.FillPaint is not null;
        // lines and open outlines keep their flat drawing and arrows; a mirrored group placement is not split here
        if (!filled && !paths.Any(p => p.Subpaths.Any(s => s.Closed))) return;
        if (shape.Transform is { } placed && placed.A * placed.D - placed.B * placed.C < 0) return;
        var sources = new List<OdtExtrusionSource>();
        foreach (WordShapePath path in paths)
        {
            if (path.Fill == WordShapePathFill.None && !path.Stroke) continue;
            var source = new OdtExtrusionSource { Filled = filled };
            foreach (WordShapeSubpath subpath in path.Subpaths) source.Rings.Add((OdtShapeStyle.Flatten(subpath), filled || subpath.Closed));
            sources.Add(source);
        }
        // the contour (a:sp3d contourW): the outline grown by half its width all round (Word's print: the band straddles the outline), extruded behind the shape in its colour (listed
        // first: at equal depths it is drawn under the shape's own faces)
        if (extrusion.ContourWidth > 0 && filled)
        {
            var contour = new OdtExtrusionSource { Filled = true, Colour = OdtShapeStyle.ParseColor(extrusion.ContourColor) ?? (0, 0, 0),
                                                 Thicken = extrusion.ContourWidth };
            // every outline and hole: an outline grows outward, a hole shrinks (the material grows into it); parts stay separate rings
            foreach (var source in sources)
            {
                var closedRings = source.Rings.Where(r => r.Closed && r.Points.Count >= 3).Select(r => r.Points).ToList();
                foreach (var ring in closedRings)
                    contour.Rings.Add((Grown(ring, (OdtExtrusion.IsHole(ring, closedRings) ? -1 : 1) * extrusion.ContourWidth / 2), true));
            }
            if (contour.Rings.Count > 0) sources.Insert(0, contour);
        }
        var fallback = OdtShapeStyle.ParseColor("#ffffff")!.Value;
        var fill = OdtShapeStyle.ParseColor(shape.FillColor ?? shape.GradientStops?.FirstOrDefault().Color) ?? fallback;
        var line = OdtShapeStyle.ParseColor(shape.LineColor) ?? OdtShapeStyle.ParseColor("#000000")!.Value;
        var secondary = OdtShapeStyle.ParseColor(extrusionColor) ?? fill;
        var faces = extrusion.Project(sources, shape.Width, shape.Height, sceneRotation, shape.FlipHorizontal, shape.FlipVertical, fill, secondary, line, false);
        var drawn = new List<WordShapePath>();
        foreach (OdtExtrusionFace face in faces)
        {
            var path = new WordShapePath { Fill = face.OutlineOnly ? WordShapePathFill.None : WordShapePathFill.Normal, Stroke = face.OutlineOnly, SolidFill = face.OutlineOnly ? null : face.Color };
            foreach (var ring in face.Rings)
            {
                if (ring.Count == 0) continue;
                var subpath = new WordShapeSubpath { StartX = ring[0].X, StartY = ring[0].Y, Closed = true };
                for (int i = 1; i < ring.Count; i++) subpath.Segments.Add(WordShapeSegment.Line(ring[i].X, ring[i].Y));
                path.Subpaths.Add(subpath);
            }
            drawn.Add(path);
        }
        if (drawn.Count == 0) return;
        shape.Paths = drawn;
        if (extrusion.ShadeMode == "draft") { shape.LineColor = OdtShapeStyle.Hex(line); shape.Dash = null; }
        else shape.LineColor = null;
        shape.HeadEnd = shape.TailEnd = default;
        shape.HeadMarker = shape.TailMarker = null;
        shape.Pattern = null; shape.GradientStops = null; shape.FillPictureId = null; shape.FillPaint = null; shape.FillDefinitions = null;
        // the scene already turned and flipped the outline about the box's centre: place the box unturned at the same centre
        if (shape.Transform is { } transform)
        {
            (double cx, double cy) = transform.Apply(shape.Width / 2, shape.Height / 2);
            WordAffine linear = new(transform.A, transform.B, transform.C, transform.D, 0, 0);
            WordAffine unrotated = WordAffine.Then(linear, WordAffine.Rotate(-Math.Atan2(transform.B, transform.A)));
            shape.Transform = WordAffine.Then(WordAffine.Then(WordAffine.Translate(-shape.Width / 2, -shape.Height / 2), unrotated), WordAffine.Translate(cx, cy));
        }
        shape.RotationDegrees = 0;
        shape.FlipHorizontal = shape.FlipVertical = false;
    }
}
