using System.Globalization;
using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.DocxHtml;

namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// One text:list-style: per level (1-10) a number or bullet definition. Attribute set and defaults
/// as LibreOffice reads them (xmloff/source/style/xmlnumi.cxx SvxXMLListLevelStyleContext_Impl):
/// text:level, style:num-format (default "1"), style:num-prefix, style:num-suffix,
/// text:start-value (default 1), text:display-levels (default 1), text:bullet-char.
/// </summary>
public sealed class OdtListStyle
{
    public const int MaxLevels = 10;

    public sealed class Level
    {
        public bool IsBullet;
        public string BulletCharacter = "•";
        public string NumberFormat = "1";
        public string Prefix = "";
        public string Suffix = "";
        public int StartValue = 1;
        public int DisplayLevels = 1;
    }

    private readonly Level[] _levels = new Level[MaxLevels];

    public OdtListStyle(XElement listStyle)
    {
        for (int i = 0; i < MaxLevels; i++) _levels[i] = new Level();
        XNamespace text = OdtStyleCascade.Text, style = OdtStyleCascade.Style;
        foreach (XElement element in listStyle.Elements())
        {
            if (!int.TryParse((string?)element.Attribute(text + "level"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int levelNumber)) continue;
            int index = Math.Max(1, Math.Min(MaxLevels, levelNumber)) - 1;
            var level = new Level { Prefix = (string?)element.Attribute(style + "num-prefix") ?? "", Suffix = (string?)element.Attribute(style + "num-suffix") ?? "" };
            if (element.Name == text + "list-level-style-number")
            {
                level.NumberFormat = (string?)element.Attribute(style + "num-format") ?? "1";
                if (int.TryParse((string?)element.Attribute(text + "start-value"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int start)) level.StartValue = start < 0 ? 1 : start;
                if (int.TryParse((string?)element.Attribute(text + "display-levels"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int display)) level.DisplayLevels = Math.Max(1, display);
            }
            else if (element.Name == text + "list-level-style-bullet")
            {
                level.IsBullet = true;
                string bullet = (string?)element.Attribute(text + "bullet-char") ?? "";
                // Private-use symbol-font bullets (Wingdings/Symbol) have no glyph without their font.
                level.BulletCharacter = bullet.Length == 0 || bullet[0] is >= '' and <= '' ? "•" : bullet;
            }
            else { level.IsBullet = true; level.BulletCharacter = "•"; }   // list-level-style-image: a plain bullet
            _levels[index] = level;
        }
    }

    public Level LevelAt(int level) => _levels[Math.Max(0, Math.Min(MaxLevels - 1, level))];

    /// <summary>
    /// The label of an item at <paramref name="level"/> given the current counters of its list:
    /// the last display-levels numbers joined with ".", wrapped in this level's prefix and suffix.
    /// </summary>
    public string FormatLabel(int level, int[] counters)
    {
        Level current = LevelAt(level);
        if (current.IsBullet) return current.Prefix + current.BulletCharacter + current.Suffix;
        var label = new StringBuilder(current.Prefix);
        int first = Math.Max(0, level - current.DisplayLevels + 1);
        for (int i = first; i <= level; i++)
        {
            Level part = LevelAt(i);
            if (part.IsBullet) continue;
            string number = FormatNumber(counters[i] != OdtListNumbering.Unset ? counters[i] : part.StartValue, part.NumberFormat);
            if (number.Length == 0) continue;
            if (label.Length > current.Prefix.Length) label.Append('.');
            label.Append(number);
        }
        return label.Append(current.Suffix).ToString();
    }

    /// <summary>
    /// style:num-format → text. "1 a A i I" and "" per xmloff/source/core/xmluconv.cxx
    /// convertNumFormat; longer names come from the numbering provider
    /// (i18npool defaultnumberingprovider.cxx): "א, י, ק, ..." is NUMBER_HEBREW (gematria).
    /// "א, ב, ג, ..." (CHARS_HEBREW, letters) is unverified. Unknown → decimal, as LibreOffice falls back to ARABIC.
    /// </summary>
    public static string FormatNumber(int value, string format) => format switch
    {
        "" => "",
        "1" => value.ToString(CultureInfo.InvariantCulture),
        "a" => DocxNumberingDefinitions.FormatNumber(value, "lowerLetter"),
        "A" => DocxNumberingDefinitions.FormatNumber(value, "upperLetter"),
        "i" => DocxNumberingDefinitions.FormatNumber(value, "lowerRoman"),
        "I" => DocxNumberingDefinitions.FormatNumber(value, "upperRoman"),
        _ when format.StartsWith("א, י", StringComparison.Ordinal) => DocxNumberingDefinitions.FormatNumber(value, "hebrew1"),
        _ when format.StartsWith("א, ב", StringComparison.Ordinal) => DocxNumberingDefinitions.FormatNumber(value, "hebrew2"),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };
}
