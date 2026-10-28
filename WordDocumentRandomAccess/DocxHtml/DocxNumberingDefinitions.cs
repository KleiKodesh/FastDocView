using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// word/numbering.xml (ECMA-376 §17.9): numId → abstractNumId → per-level number format, start value, label pattern
/// ("%1.", "%1.%2"), restart rule, legal numbering, suffix and label run properties. The level a paragraph uses is looked
/// up the way Open-Xml-PowerTools' ListItemRetriever does it (num/lvlOverride/lvl first — "applied in place of" the
/// abstract level, §17.9.9 — then the abstractNum's lvl); an abstractNum with w:numStyleLink takes its levels from the
/// abstractNum of the num that the linked numbering style's numPr names (§17.9.21). Lists sharing an abstractNum continue
/// one sequence unless a num carries a startOverride. Labels are computed once at index time by
/// <see cref="DocxListCounter"/>, because a windowed renderer never sees the earlier items.
/// </summary>
public sealed class DocxNumberingDefinitions
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    /// <param name="Restart">w:lvlRestart (§17.9.10): null = restart after any higher level; 0 = never; N = when a level 1..N (one-based) is used.</param>
    /// <param name="IsLegal">w:isLgl (§17.9.4): every level shown in this level's text is decimal.</param>
    /// <param name="Suffix">w:suff (§17.9.28): tab (the default), space or nothing after the label.</param>
    /// <param name="RunProperties">The level's w:rPr: the label's own formatting.</param>
    public sealed record LevelDefinition(string Format, string Text, int Start, int? Restart = null, bool IsLegal = false,
                                         string Suffix = "tab", XElement? RunProperties = null);

    private readonly Dictionary<int, int> _abstractIdByNumId = new();
    private readonly Dictionary<int, Dictionary<int, int>> _startOverridesByNumId = new();
    private readonly Dictionary<(int numId, int level), LevelDefinition> _overrideLevels = new();
    private readonly Dictionary<(int abstractId, int level), LevelDefinition> _levels = new();
    private readonly Dictionary<int, string> _numStyleLinkByAbstractId = new();
    private readonly Func<string, int>? _numIdOfStyle;

    /// <param name="numIdOfStyle">A numbering style's numId (its pPr/numPr), for w:numStyleLink; -1 when it has none.</param>
    public DocxNumberingDefinitions(XDocument? numberingDocument, Func<string, int>? numIdOfStyle = null)
    {
        _numIdOfStyle = numIdOfStyle;
        XElement? root = numberingDocument?.Root;
        if (root is null) return;
        foreach (XElement num in root.Elements(W + "num"))
        {
            int? numId = DocxStyleCascade.TwipsAttribute(num, "numId");
            int? abstractId = DocxStyleCascade.TwipsAttribute(num.Element(W + "abstractNumId"), "val");
            if (numId is not { } id || abstractId is not { } abstractNumId) continue;
            _abstractIdByNumId[id] = abstractNumId;
            foreach (XElement levelOverride in num.Elements(W + "lvlOverride"))
            {
                int level = DocxStyleCascade.TwipsAttribute(levelOverride, "ilvl") ?? 0;
                if (DocxStyleCascade.TwipsAttribute(levelOverride.Element(W + "startOverride"), "val") is { } start)
                {
                    if (!_startOverridesByNumId.TryGetValue(id, out var overrides)) _startOverridesByNumId[id] = overrides = new Dictionary<int, int>();
                    overrides[level] = start;
                }
                if (levelOverride.Element(W + "lvl") is { } lvl) _overrideLevels[(id, level)] = ReadLevel(lvl, level);
            }
        }
        foreach (XElement abstractNum in root.Elements(W + "abstractNum"))
        {
            int? abstractId = DocxStyleCascade.TwipsAttribute(abstractNum, "abstractNumId");
            if (abstractId is not { } id) continue;
            if ((string?)abstractNum.Element(W + "numStyleLink")?.Attribute(W + "val") is { } link) _numStyleLinkByAbstractId[id] = link;
            foreach (XElement level in abstractNum.Elements(W + "lvl"))
            {
                int levelIndex = DocxStyleCascade.TwipsAttribute(level, "ilvl") ?? 0;
                _levels[(id, levelIndex)] = ReadLevel(level, levelIndex);
            }
        }
    }

    private static LevelDefinition ReadLevel(XElement level, int levelIndex)
    {
        // A custom number format is written as mc:AlternateContent (Choice: numFmt val="custom", Fallback: a plain numFmt);
        // the fallback is what every reader without that extension uses.
        XElement? numFmt = level.Element(W + "numFmt")
            ?? level.Elements().Where(e => e.Name.LocalName == "AlternateContent").Elements().Where(e => e.Name.LocalName == "Fallback").Elements(W + "numFmt").FirstOrDefault();
        string format = (string?)numFmt?.Attribute(W + "val") ?? "decimal";
        string text = (string?)level.Element(W + "lvlText")?.Attribute(W + "val") ?? "%" + (levelIndex + 1) + ".";
        int start = DocxStyleCascade.TwipsAttribute(level.Element(W + "start"), "val") ?? 1;
        int? restart = DocxStyleCascade.TwipsAttribute(level.Element(W + "lvlRestart"), "val");
        bool isLegal = level.Element(W + "isLgl") is { } legal && DocxStyleCascade.IsOn(legal.Attribute(W + "val"));
        string suffix = (string?)level.Element(W + "suff")?.Attribute(W + "val") ?? "tab";
        return new LevelDefinition(format, text, start, restart, isLegal, suffix, level.Element(W + "rPr"));
    }

    public bool IsEmpty => _abstractIdByNumId.Count == 0;

    /// <summary>The abstractNum whose levels and counters a num uses — through w:numStyleLink when the abstractNum is only a link.</summary>
    public int? AbstractIdOf(int numId)
    {
        if (!_abstractIdByNumId.TryGetValue(numId, out int abstractId)) return null;
        var seen = new HashSet<int> { abstractId };
        while (_numStyleLinkByAbstractId.TryGetValue(abstractId, out string? style))
        {
            int linkedNumId = _numIdOfStyle?.Invoke(style) ?? -1;
            if (!_abstractIdByNumId.TryGetValue(linkedNumId, out int linkedAbstractId) || !seen.Add(linkedAbstractId)) break;
            abstractId = linkedAbstractId;
        }
        return abstractId;
    }

    /// <summary>The level definition a paragraph of <paramref name="numId"/> uses: the num's lvlOverride/lvl, else the abstract level.</summary>
    public LevelDefinition? LevelOf(int numId, int abstractId, int level)
        => _overrideLevels.TryGetValue((numId, level), out LevelDefinition? own) ? own
         : _levels.TryGetValue((abstractId, level), out LevelDefinition? definition) ? definition : null;

    public int? StartOverrideOf(int numId, int level)
        => _startOverridesByNumId.TryGetValue(numId, out var overrides) && overrides.TryGetValue(level, out int start) ? start : null;

    /// <summary>Fills the "%1.%2" pattern of a level with the current counters of all levels; with isLgl every number is decimal.</summary>
    public string FormatLabel(int numId, int abstractId, int level, int[] counters)
    {
        if (LevelOf(numId, abstractId, level) is not { } definition) return "•";
        if (definition.Format == "bullet") return BulletCharacter(definition.Text);
        if (definition.Format == "none") return "";
        var label = new StringBuilder(definition.Text.Length + 8);
        for (int i = 0; i < definition.Text.Length; i++)
        {
            char c = definition.Text[i];
            if (c == '%' && i + 1 < definition.Text.Length && definition.Text[i + 1] is >= '1' and <= '9')
            {
                int placeholderLevel = definition.Text[i + 1] - '1';
                string format = definition.IsLegal ? "decimal"
                              : placeholderLevel == level ? definition.Format : LevelOf(numId, abstractId, placeholderLevel)?.Format ?? "decimal";
                label.Append(FormatNumber(counters[placeholderLevel], format));
                i++;
            }
            else label.Append(c);
        }
        return label.ToString();
    }

    private static string BulletCharacter(string text)
        => text.Length == 1 && text[0] is >= '' and <= '' ? "•" : text.Length == 0 ? "•" : text;

    public static string FormatNumber(int value, string format) => format switch
    {
        "lowerLetter" => Letters(value, 'a'),
        "upperLetter" => Letters(value, 'A'),
        "lowerRoman" => Roman(value).ToLowerInvariant(),
        "upperRoman" => Roman(value),
        "hebrew1" => HebrewNumeral(value),
        "hebrew2" => HebrewLetters(value),
        "decimalZero" => value.ToString("00"),
        "none" => "",
        _ => value.ToString(),
    };

    private static string Letters(int value, char first)
    {
        if (value <= 0) return "";
        int repeat = (value - 1) / 26 + 1;
        return new string((char)(first + (value - 1) % 26), repeat);
    }

    private static string Roman(int value)
    {
        if (value <= 0 || value >= 4000) return value.ToString();
        (int number, string numeral)[] table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")];
        var roman = new StringBuilder();
        foreach ((int number, string numeral) in table) while (value >= number) { roman.Append(numeral); value -= number; }
        return roman.ToString();
    }

    private static readonly string[] HebrewOnes = ["", "א", "ב", "ג", "ד", "ה", "ו", "ז", "ח", "ט"];
    private static readonly string[] HebrewTens = ["", "י", "כ", "ל", "מ", "נ", "ס", "ע", "פ", "צ"];
    private static readonly string[] HebrewHundreds = ["", "ק", "ר", "ש", "ת", "תק", "תר", "תש", "תת", "תתק"];

    /// <summary>
    /// Word's hebrew1: gematria letters (טו / טז for 15 / 16) with no geresh or gershayim, cycling after 392 — what
    /// Word 16 shows (ListFormat.ListString of synthetic lists, 2026-10-28: 11 → יא, 15 → טו, 115 → קטו, 500 → קח,
    /// 1000 → רטז, 5784 → רצו).
    /// </summary>
    public static string HebrewNumeral(int value)
    {
        if (value <= 0) return value.ToString();
        value = (value - 1) % 392 + 1;
        var letters = new StringBuilder();
        letters.Append(HebrewHundreds[value / 100]);
        int tensAndOnes = value % 100;
        if (tensAndOnes == 15) letters.Append("טו");
        else if (tensAndOnes == 16) letters.Append("טז");
        else letters.Append(HebrewTens[tensAndOnes / 10]).Append(HebrewOnes[tensAndOnes % 10]);
        return letters.ToString();
    }

    /// <summary>Word's hebrew2: the 22 letters cycling (א..ת, then אא..).</summary>
    public static string HebrewLetters(int value)
    {
        if (value <= 0) return "";
        const string alphabet = "אבגדהוזחטיכלמנסעפצקרשת";
        int repeat = (value - 1) / alphabet.Length + 1;
        return new string(alphabet[(value - 1) % alphabet.Length], repeat);
    }
}
