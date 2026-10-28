using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// word/numbering.xml (ECMA-376 §17.9): numId → abstractNumId (+ level start overrides) → per-level
/// number format, start value and label pattern ("%1.", "%1.%2"). Lists sharing an abstractNum
/// continue one sequence unless the num carries a startOverride, which is what Word does.
/// Labels are computed once at index time by <see cref="DocxListCounter"/>, because a windowed
/// renderer never sees the earlier items.
/// </summary>
public sealed class DocxNumberingDefinitions
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    public readonly record struct LevelDefinition(string Format, string Text, int Start);

    private readonly Dictionary<int, int> _abstractIdByNumId = new();
    private readonly Dictionary<int, Dictionary<int, int>> _startOverridesByNumId = new();
    private readonly Dictionary<(int abstractId, int level), LevelDefinition> _levels = new();

    public DocxNumberingDefinitions(XDocument? numberingDocument)
    {
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
            }
        }
        foreach (XElement abstractNum in root.Elements(W + "abstractNum"))
        {
            int? abstractId = DocxStyleCascade.TwipsAttribute(abstractNum, "abstractNumId");
            if (abstractId is not { } id) continue;
            foreach (XElement level in abstractNum.Elements(W + "lvl"))
            {
                int levelIndex = DocxStyleCascade.TwipsAttribute(level, "ilvl") ?? 0;
                string format = (string?)level.Element(W + "numFmt")?.Attribute(W + "val") ?? "decimal";
                string text = (string?)level.Element(W + "lvlText")?.Attribute(W + "val") ?? "%" + (levelIndex + 1) + ".";
                int start = DocxStyleCascade.TwipsAttribute(level.Element(W + "start"), "val") ?? 1;
                _levels[(id, levelIndex)] = new LevelDefinition(format, text, start);
            }
        }
    }

    public bool IsEmpty => _abstractIdByNumId.Count == 0;

    public int? AbstractIdOf(int numId) => _abstractIdByNumId.TryGetValue(numId, out int abstractId) ? abstractId : null;

    public LevelDefinition? LevelOf(int abstractId, int level)
        => _levels.TryGetValue((abstractId, level), out LevelDefinition definition) ? definition : null;

    public int? StartOverrideOf(int numId, int level)
        => _startOverridesByNumId.TryGetValue(numId, out var overrides) && overrides.TryGetValue(level, out int start) ? start : null;

    /// <summary>Fills the "%1.%2" pattern of a level with the current counters of all levels.</summary>
    public string FormatLabel(int abstractId, int level, int[] counters)
    {
        if (LevelOf(abstractId, level) is not { } definition) return "•";
        if (definition.Format == "bullet") return BulletCharacter(definition.Text);
        if (definition.Format == "none") return "";
        var label = new StringBuilder(definition.Text.Length + 8);
        for (int i = 0; i < definition.Text.Length; i++)
        {
            char c = definition.Text[i];
            if (c == '%' && i + 1 < definition.Text.Length && definition.Text[i + 1] is >= '1' and <= '9')
            {
                int placeholderLevel = definition.Text[i + 1] - '1';
                string format = placeholderLevel == level ? definition.Format : LevelOf(abstractId, placeholderLevel)?.Format ?? "decimal";
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

    /// <summary>Gematria numerals as Word's hebrew1: טו/טז for 15/16, gershayim before the last letter.</summary>
    public static string HebrewNumeral(int value)
    {
        if (value <= 0 || value >= 10000) return value.ToString();
        var letters = new StringBuilder();
        int thousands = value / 1000;
        if (thousands > 0) letters.Append(HebrewOnes[thousands]).Append('׳');
        int rest = value % 1000;
        letters.Append(HebrewHundreds[rest / 100]);
        int tensAndOnes = rest % 100;
        if (tensAndOnes == 15) letters.Append("טו");
        else if (tensAndOnes == 16) letters.Append("טז");
        else letters.Append(HebrewTens[tensAndOnes / 10]).Append(HebrewOnes[tensAndOnes % 10]);
        string text = letters.ToString();
        if (text.Length == 1) return text + '׳';                                  // geresh after a single letter
        return text.Substring(0, text.Length - 1) + '״' + text[text.Length - 1];  // gershayim before the last letter
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
