using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// A DrawingML picture's effects (the &lt;effects&gt; copy <see cref="Docx.DocxDrawingStripper"/> keeps in the picture layout)
/// read into <see cref="WordPictureEffects"/>. a:blip children in document order (ECMA-376 Part 1 20.1.8: grayscl, biLevel
/// thresh, lum bright / contrast, duotone, clrChange clrFrom / clrTo, alphaModFix amt; percentages in 1000ths), a:softEdge rad
/// (EMU), a:scene3d a:camera a:rot lat / lon / rev (60000ths of a degree). LibreOffice reads the same attributes
/// (oox/source/drawingml/misccontexts.cxx:337-391 BlipContext) and the colours through its colour model, as
/// <see cref="DocxDrawingColor"/> does here.
/// </summary>
internal static class DocxPictureEffects
{
    private const double Percent = 100000, EmuPerPoint = 12700, PerDegree = 60000;

    public static WordPictureEffects? Read(XElement? effects, DocxTheme? theme)
    {
        if (effects is null) return null;
        var colorEffects = new List<WordPictureColorEffect>();
        double opacity = 1;
        if (effects.Element("blip") is { } blip)
            foreach (XElement effect in blip.Elements())
                switch (effect.Name.LocalName)
                {
                    case "grayscl": colorEffects.Add(new WordPictureColorEffect.Grayscale()); break;
                    case "biLevel": colorEffects.Add(new WordPictureColorEffect.BiLevel(Number(effect, "thresh", 0) / Percent)); break;
                    case "lum":
                        double brightness = Number(effect, "bright", 0), contrast = Number(effect, "contrast", 0);
                        if (brightness != 0 || contrast != 0) colorEffects.Add(new WordPictureColorEffect.Luminance(brightness / Percent, contrast / Percent));
                        break;
                    case "alphaModFix": opacity *= Math.Max(0, Math.Min(1, Number(effect, "amt", Percent) / Percent)); break;
                    case "duotone":
                    {
                        List<XElement> colors = effect.Elements().ToList();
                        if (colors.Count >= 2 && Color(colors[0], theme) is { } dark && Color(colors[1], theme) is { } light)
                            colorEffects.Add(new WordPictureColorEffect.Duotone(dark, light));
                        break;
                    }
                    case "clrChange":
                        if (DocxDrawingColor.Resolve(effect.Element("clrFrom"), theme) is { } from && Rgb(from.Css) is { } fromRgb
                            && DocxDrawingColor.Resolve(effect.Element("clrTo"), theme) is { } to && Rgb(to.Css) is { } toRgb
                            && (fromRgb != toRgb || to.Opacity < 1))
                            colorEffects.Add(new WordPictureColorEffect.ColorChange(fromRgb, toRgb, to.Opacity));
                        break;
                }
        double softEdge = effects.Element("effectLst")?.Element("softEdge") is { } edge ? Number(edge, "rad", 0) / EmuPerPoint : 0;
        (double, double, double) rotation = default;
        if (effects.Element("scene3d")?.Element("camera")?.Element("rot") is { } rot)
            rotation = (Number(rot, "lat", 0) / PerDegree, Number(rot, "lon", 0) / PerDegree, Number(rot, "rev", 0) / PerDegree);
        // a:camera prst (ST_PresetCameraType, 20.1.10.47): the perspective presets (perspective*, legacyPerspective*) look through
        // a:camera fov (60000ths of a degree), else 45 degrees - Word's perspectiveFront and perspectiveRelaxed both draw at
        // 45 (fitted on its PDF, WordPictureEffects.CameraHorizontalDistance); the others are orthographic. A preset's own
        // rotation (used without a:rot) is not known here: such a picture is drawn flat.
        double fieldOfView = 0;
        if (effects.Element("scene3d")?.Element("camera") is { } camera
            && ((string?)camera.Attribute("prst") ?? "") is var preset && (preset.StartsWith("perspective", StringComparison.Ordinal) || preset.StartsWith("legacyPerspective", StringComparison.Ordinal)))
            fieldOfView = camera.Attribute("fov") is not null ? Number(camera, "fov", 0) / PerDegree : 45;
        var result = new WordPictureEffects { ColorEffects = colorEffects, Opacity = opacity, SoftEdgePt = Math.Max(0, softEdge), Rotation = rotation, CameraFieldOfView = fieldOfView,
                                             HasScene = effects.Element("scene3d")?.Element("camera")?.Element("rot") is not null };
        return result.IsEmpty ? null : result;
    }

    /// <summary>One colour element (with its transforms) as RGB; the resolver reads the first colour child of a parent.</summary>
    private static (int R, int G, int B)? Color(XElement color, DocxTheme? theme)
        => DocxDrawingColor.Resolve(new XElement("color", new XElement(color)), theme, unclampedSaturation: true) is { } resolved ? Rgb(resolved.Css) : null;

    private static (int R, int G, int B)? Rgb(string css)
        => css.Length == 7 && css[0] == '#' && int.TryParse(css.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int value)
            ? ((value >> 16) & 0xFF, (value >> 8) & 0xFF, value & 0xFF) : null;

    private static double Number(XElement element, string name, double fallback)
        => double.TryParse((string?)element.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : fallback;
}
