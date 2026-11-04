using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// .docx WordArt mapped onto the shared text-effect drawing (WordOfficeArtGeometry.Preset + WordOfficeArtFontwork.Apply), as
/// LibreOffice imports it:
/// - DrawingML a:prstTxWarp (wps:bodyPr): the preset name to its WordArt shape (oox presetgeometrynames.cxx:28-70 name to
///   fontwork type, svx EnhancedCustomShapeTypeNames.cxx:171-210 fontwork type to MSO_SPT, include/svx/msdffdef.hxx:412-451
///   the numbers), its a:avLst adjust values turned into the WordArt units (oox fontworkhelpers.cxx:95-148: the arch / circle
///   angles to degrees, wave x-handles (v + 50000) x 0.216, pour radii x 0.108, the rest x 0.216); textNoShape = no effect;
/// - VML: a v:shape of type #_x0000_tNNN (o:spt) in 136..175 with v:textpath string, its adj list in the same units.
/// </summary>
internal static class DocxWordArt
{
    private static readonly Dictionary<string, int> WarpShapeTypes = new(StringComparer.Ordinal)
    {
        ["textPlain"] = 136, ["textStop"] = 137, ["textTriangle"] = 138, ["textTriangleInverted"] = 139, ["textChevron"] = 140, ["textChevronInverted"] = 141,
        ["textRingInside"] = 142, ["textRingOutside"] = 143, ["textArchUp"] = 144, ["textArchDown"] = 145, ["textCircle"] = 146, ["textButton"] = 147,
        ["textArchUpPour"] = 148, ["textArchDownPour"] = 149, ["textCirclePour"] = 150, ["textButtonPour"] = 151, ["textCurveUp"] = 152, ["textCurveDown"] = 153,
        ["textCascadeUp"] = 154, ["textCascadeDown"] = 155, ["textWave1"] = 156, ["textWave2"] = 157, ["textDoubleWave1"] = 158, ["textWave4"] = 159,
        ["textInflate"] = 160, ["textDeflate"] = 161, ["textInflateBottom"] = 162, ["textDeflateBottom"] = 163, ["textInflateTop"] = 164, ["textDeflateTop"] = 165,
        ["textDeflateInflate"] = 166, ["textDeflateInflateDeflate"] = 167, ["textFadeRight"] = 168, ["textFadeLeft"] = 169, ["textFadeUp"] = 170, ["textFadeDown"] = 171,
        ["textSlantUp"] = 172, ["textSlantDown"] = 173, ["textCanUp"] = 174, ["textCanDown"] = 175,
    };

    /// <summary>The WordArt shape type and adjust values of a prefix-free a:prstTxWarp; null for textNoShape or an unknown name.</summary>
    public static (int ShapeType, List<int?> Adjust)? FromTextWarp(XElement? warp)
    {
        string? preset = (string?)warp?.Attribute("prst");
        if (preset is null || !WarpShapeTypes.TryGetValue(preset, out int type)) return null;
        var adjust = new List<int?>();
        foreach (XElement guide in warp!.Element("avLst")?.Elements("gd") ?? [])
        {
            string formula = ((string?)guide.Attribute("fmla") ?? "").Trim(), name = (string?)guide.Attribute("name") ?? "";
            if (!formula.StartsWith("val ", StringComparison.Ordinal) || !double.TryParse(formula.Substring(4), NumberStyles.Float, CultureInfo.InvariantCulture, out double value)) { adjust.Add(null); continue; }
            bool polar = type is 144 or 145 or 146 or 147 || (type is 148 or 149 or 150 or 151 && name == "adj1");
            if (polar) { double degrees = value / 60000 % 360; if (degrees < 0) degrees += 360; adjust.Add((int)Math.Round(degrees)); }
            else if (type is 156 or 157 or 158 or 159 && name == "adj2") adjust.Add((int)Math.Round((value + 50000) * 0.216));
            else if (type is 148 or 149 or 150 or 151 && name == "adj2") adjust.Add((int)Math.Round(value * 0.108));
            else adjust.Add((int)Math.Round(value * 0.216));
        }
        return (type, adjust);
    }

    /// <summary>A VML text-effect shape type from v:shape type "#_x0000_tNNN" (or o:spt); null when it is no WordArt type.</summary>
    public static int? VmlShapeType(string? type, string? spt)
    {
        string? digits = spt ?? (type is not null && type.LastIndexOf("_t", StringComparison.Ordinal) is var at and >= 0 ? type.Substring(at + 2) : null);
        return int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) && value is >= 136 and <= 175 ? value : null;
    }

    /// <summary>VML adj: comma-separated integers, empty = the shape's default.</summary>
    public static List<int?> VmlAdjust(string? adj)
        => (adj ?? "").Split(',').Where(_ => adj is { Length: > 0 }).Select(p => int.TryParse(p.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? (int?)v : null).ToList();

    /// <summary>Turns a drawn shape into its WordArt: the preset's outlines at its size, the text laid on them (WordOfficeArtFontwork).</summary>
    public static void Apply(WordShape shape, int shapeType, IReadOnlyList<int?> adjust, string text, string? font, string? color)
    {
        if (WordOfficeArtGeometry.Preset(shapeType, shape.Width, shape.Height, adjust) is not { } paths) return;
        shape.Paths = paths;
        WordOfficeArtFontwork.Apply(shape, text, font);
        if (color is not null) shape.TextOnPathColor = color;
    }
}
