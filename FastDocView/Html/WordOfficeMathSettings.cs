using System.Globalization;
using System.Xml.Linq;

namespace FastDocView.Html;

/// <summary>
/// The document's math settings (ECMA-376 Part 1 22.1.2.62 m:mathPr in settings.xml; RTF {\mmathPr}) that change what an
/// equation prints like. Defaults are the schema's / Word's: mathFont Cambria Math, brkBin before, brkBinSub "--",
/// smallFrac off, dispDef on, no margins, defJc centerGroup, wrapIndent 1440 twips, wrapRight off, intLim subSup, naryLim
/// undOvr (Word 16 writes exactly these in a new document's settings.xml, test-data/word-math/math-v2.docx). A .doc has no
/// place for them (the defaults apply).
/// </summary>
public sealed class WordOfficeMathSettings
{
    public static readonly WordOfficeMathSettings Default = new();

    public string? Font { get; init; }                         // m:mathFont (22.1.2.61)
    public string BreakBinary { get; init; } = "before";       // m:brkBin (22.1.2.15): before, after, repeat
    public string BreakBinarySubtraction { get; init; } = "--";   // m:brkBinSub (22.1.2.16): "--", "-+", "+-" when brkBin is repeat
    public bool SmallFractions { get; init; }                  // m:smallFrac (22.1.2.98)
    public bool DisplayDefaults { get; init; } = true;         // m:dispDef (22.1.2.27)
    public double LeftMarginPt { get; init; }                  // m:lMargin (22.1.2.56), twips
    public double RightMarginPt { get; init; }                 // m:rMargin (22.1.2.89)
    public string DefaultJustification { get; init; } = "centerGroup";   // m:defJc (22.1.2.25)
    public double WrapIndentPt { get; init; } = 72;            // m:wrapIndent (22.1.2.120), 1440 twips
    public bool WrapRight { get; init; }                       // m:wrapRight (22.1.2.121)
    public string IntegralLimits { get; init; } = "subSup";    // m:intLim (22.1.2.49)
    public string NaryLimits { get; init; } = "undOvr";        // m:naryLim (22.1.2.72)

    /// <summary>From settings.xml's m:mathPr (null: the defaults).</summary>
    public static WordOfficeMathSettings FromOoxml(XElement? mathProperties)
    {
        if (mathProperties is null) return Default;
        string? Value(string name) => mathProperties.Elements().FirstOrDefault(e => e.Name.LocalName == name) is { } element
            ? element.Attributes().FirstOrDefault(a => a.Name.LocalName == "val")?.Value ?? "" : null;
        bool? On(string name) => Value(name) is { } value ? value is "" or "1" or "on" or "true" : null;
        double? Twips(string name) => double.TryParse(Value(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out double twips) ? Math.Max(0, Math.Min(twips, 31680)) / 20 : null;
        return new WordOfficeMathSettings
        {
            Font = Value("mathFont") is { Length: > 0 } font ? font : null,
            BreakBinary = Value("brkBin") is "after" or "repeat" or "before" ? Value("brkBin")! : "before",
            BreakBinarySubtraction = Value("brkBinSub") is "-+" or "+-" or "--" ? Value("brkBinSub")! : "--",
            SmallFractions = On("smallFrac") ?? false,
            DisplayDefaults = On("dispDef") ?? true,
            LeftMarginPt = Twips("lMargin") ?? 0,
            RightMarginPt = Twips("rMargin") ?? 0,
            DefaultJustification = Value("defJc") is "left" or "right" or "center" or "centerGroup" ? Value("defJc")! : "centerGroup",
            WrapIndentPt = Twips("wrapIndent") ?? 72,
            WrapRight = On("wrapRight") ?? false,
            IntegralLimits = Value("intLim") == "undOvr" ? "undOvr" : "subSup",
            NaryLimits = Value("naryLim") == "subSup" ? "subSup" : "undOvr",
        };
    }

    /// <summary>
    /// From RTF {\mmathPr} control words (RTF 1.9.1; numbers as Word 16 writes them in math-v2.rtf: \mbrkBin0 = before,
    /// \mbrkBinSub0 = "--", \mdefJc1 = centerGroup, \mintLim0 = subSup, \mnaryLim1 = undOvr): \mbrkBin 0 / 1 / 2 = before /
    /// after / repeat, \mbrkBinSub 0 / 1 / 2 = "--" / "-+" / "+-", \mdefJc 1..4 = centerGroup / center / left / right,
    /// \mintLim / \mnaryLim 0 = subSup, 1 = undOvr, margins and \mwrapIndent in twips, \mmathFontN a font table index.
    /// </summary>
    public static WordOfficeMathSettings FromRtf(IReadOnlyDictionary<string, int> words, Func<int, string?> fontName)
    {
        if (words.Count == 0) return Default;
        int? Get(string name) => words.TryGetValue(name, out int value) ? value : null;
        return new WordOfficeMathSettings
        {
            Font = Get("mmathFont") is { } font ? fontName(font) : null,
            BreakBinary = Get("mbrkBin") switch { 1 => "after", 2 => "repeat", _ => "before" },
            BreakBinarySubtraction = Get("mbrkBinSub") switch { 1 => "-+", 2 => "+-", _ => "--" },
            SmallFractions = Get("msmallFrac") is > 0,
            DisplayDefaults = Get("mdispDef") is not 0,
            LeftMarginPt = Math.Max(0, Get("mlMargin") ?? 0) / 20.0,
            RightMarginPt = Math.Max(0, Get("mrMargin") ?? 0) / 20.0,
            DefaultJustification = Get("mdefJc") switch { 2 => "center", 3 => "left", 4 => "right", _ => "centerGroup" },
            WrapIndentPt = Math.Max(0, Get("mwrapIndent") ?? 1440) / 20.0,
            WrapRight = Get("mwrapRight") is > 0,
            IntegralLimits = Get("mintLim") == 1 ? "undOvr" : "subSup",
            NaryLimits = Get("mnaryLim") == 0 ? "subSup" : "undOvr",
        };
    }
}
