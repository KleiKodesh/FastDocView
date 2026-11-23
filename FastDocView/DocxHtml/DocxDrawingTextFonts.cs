using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// Measures DrawingML paragraphs (a:p with a:r / a:fld runs, namespace-free) in the fonts Word draws them with, for SmartArt
/// autofit. Each character takes a font slot as Office text does: complex-script characters (Hebrew, Arabic, Syriac, Thaana,
/// Thai, ... ) the run's a:cs typeface, else the theme's complex-script font for that script (DocxTheme.ComplexScriptFont: the
/// supplemental a:font, the Office theme's "Hebr" = Arial, else a:cs), else Arial / Times New Roman; East Asian characters (Han,
/// kana, Hangul, CJK punctuation and full-width forms) the run's a:ea typeface, else the theme's East Asian font for the
/// character's script (DocxTheme.EastAsianFont: a:ea, else the supplemental "Jpan" / "Hans" / "Hant" / "Hang" font; Han takes
/// the script of the run's lang, else of w:themeFontLang/@eastAsia, else Jpan), else the faces Windows ships for that script;
/// other letters the run's a:latin typeface, else the theme's minor Latin font, else Calibri ("+mn-lt" / "+mj-cs" / "+mn-ea" ...
/// name the theme's). Spaces, digits and punctuation take the slot of the strong character before them (after them at the
/// start). A face without the character's glyph hands it to the next one in its chain, then to the other slots' chains. A run's
/// b / i pick the bold / italic font file. Consecutive characters of one face are shaped together (WordFontMetrics.ShapedWidth:
/// GSUB for complex scripts, marks without width, kern table or GPOS pair kerning kept apart: Word kerns SmartArt runs from
/// kern="1200"). Japanese and Chinese text may break between any two ideographs or kana (not before closing punctuation, small
/// kana or the prolonged sound mark, not after opening punctuation); Korean breaks at spaces. One instance per layout pass;
/// each paragraph is measured once (words in em at size 1), so fitting a size is cheap.
/// </summary>
internal sealed class DocxDrawingTextFonts(DocxTheme? theme)
{
    /// <summary>A paragraph's words at size 1 (em): advance, the kerning inside it, the space before it, and whether a line break comes before it.</summary>
    internal sealed class MeasuredParagraph(double[] words, double[] kerning, double[] spaces, bool[] breakBefore, double lineHeightEm)
    {
        public readonly double[] Words = words, Kerning = kerning, Spaces = spaces;
        public readonly bool[] BreakBefore = breakBefore;
        public readonly double LineHeightEm = lineHeightEm;
    }

    /// <summary>Test switch: measure East Asian text in the Latin chain, as before (the SmartArt harness proves the fixtures depend on it).</summary>
    internal static bool DisableEastAsianSlot;

    /// <summary>Whether complex-script runs add their pair kerning: Word kerns them (hadassah-hebrew-block-list is 43 pt only kerned); a test switch.</summary>
    internal static bool KernComplexScripts = true;

    /// <summary>
    /// A line holding East Asian text is 1.3 times the font's line height (hhea): Word's drawings of the multi-line Japanese /
    /// Korean / Chinese fixtures (Yu Mincho 1.287 em, Malgun Gothic 1.330, SimSun 1.141) fit only with it.
    /// </summary>
    private const double EastAsianLineFactor = 1.3;

    private readonly Dictionary<XElement, MeasuredParagraph> _measured = [];
    private readonly object _lock = new();

    /// <summary>The paragraph's words in em at size 1 (cached per paragraph element).</summary>
    public MeasuredParagraph Measure(XElement paragraph)
    {
        lock (_lock)
        {
            if (_measured.TryGetValue(paragraph, out MeasuredParagraph? known)) return known;
            MeasuredParagraph measured = MeasureNew(paragraph);
            _measured[paragraph] = measured;
            return measured;
        }
    }

    private enum Slot : byte { Neutral, Latin, Complex, EastAsian, Break }

    private sealed record RunFonts(string? Latin, string? Complex, string? EastAsian, bool Bold, bool Italic, string? Language);

    private MeasuredParagraph MeasureNew(XElement paragraph)
    {
        // the paragraph's characters with their run's fonts
        var codePoints = new List<int>();
        var slots = new List<Slot>();
        var runs = new List<RunFonts>();
        var runOf = new List<int>();
        foreach (XElement run in paragraph.Elements())
        {
            if (run.Name.LocalName == "br") { codePoints.Add('\n'); slots.Add(Slot.Break); runOf.Add(-1); continue; }
            if (run.Name.LocalName is not ("r" or "fld")) continue;
            XElement? properties = run.Element("rPr");
            static bool On(string? value) => value is "1" or "true" or "on";
            runs.Add(new RunFonts((string?)properties?.Element("latin")?.Attribute("typeface"), (string?)properties?.Element("cs")?.Attribute("typeface"),
                                  (string?)properties?.Element("ea")?.Attribute("typeface"), On((string?)properties?.Attribute("b")), On((string?)properties?.Attribute("i")),
                                  (string?)properties?.Attribute("lang")));
            string text = (string?)run.Element("t") ?? "";
            for (int i = 0; i < text.Length; i++)
            {
                int codePoint = text[i];
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) codePoint = char.ConvertToUtf32(text[i], text[++i]);
                codePoints.Add(codePoint);
                slots.Add(ComplexScriptOf(codePoint) is not null ? Slot.Complex
                          : !DisableEastAsianSlot && EastAsianScriptOf(codePoint) is not null ? Slot.EastAsian
                          : IsLetter(codePoint) ? Slot.Latin : Slot.Neutral);
                runOf.Add(runs.Count - 1);
            }
        }
        // neutrals take the slot of the strong character before them, else after them, else Latin
        Slot last = Slot.Neutral;
        for (int i = 0; i < slots.Count; i++)
            if (slots[i] is Slot.Latin or Slot.Complex or Slot.EastAsian) last = slots[i];
            else if (slots[i] == Slot.Neutral && last != Slot.Neutral) slots[i] = last;
        Slot next = Slot.Latin;
        for (int i = slots.Count - 1; i >= 0; i--)
            if (slots[i] is Slot.Latin or Slot.Complex or Slot.EastAsian) next = slots[i];
            else if (slots[i] == Slot.Neutral) slots[i] = next;

        var words = new List<double>();
        var kerning = new List<double>();
        var spaces = new List<double>();
        var breaks = new List<bool>();
        double lineHeight = 0, word = 0, wordKerning = 0, space = 0;
        bool inWord = false, breakPending = false;
        // the characters of the current word drawn in one face: shaped together
        var segment = new List<int>();
        WordFontMetrics? segmentFace = null;
        bool segmentComplex = false;
        string segmentScript = "latn";
        void FlushSegment()
        {
            if (segment.Count == 0) return;
            if (segmentFace is null) word += segment.Count(c => !WordFontMetrics.IsNonSpacingMark(c)) * 0.5;
            else
            {
                word += segmentFace.ShapedWidth(segment, segmentScript, segmentComplex, out double segmentKerning);
                if (!segmentComplex || KernComplexScripts) wordKerning += segmentKerning;
            }
            segment.Clear();
        }
        void EndWord()
        {
            FlushSegment();
            words.Add(word); kerning.Add(wordKerning); inWord = false;
        }
        string? script = null;      // the script of the last complex-script character: neutrals after it use that script's font
        for (int i = 0; i < codePoints.Count; i++)
        {
            int codePoint = codePoints[i];
            if (slots[i] == Slot.Break)
            {
                if (inWord) EndWord();
                else
                {
                    // an empty line (a break at the start or right after another): a word of no width holds its line
                    words.Add(0); kerning.Add(0); spaces.Add(0); breaks.Add(breakPending && words.Count > 1);
                }
                breakPending = true; space = 0;
                continue;
            }
            Slot slot = slots[i];
            if (ComplexScriptOf(codePoint) is { } own) script = own;
            RunFonts fonts = runs[runOf[i]];
            string? eastAsianScript = slot != Slot.EastAsian ? null : EastAsianScriptOf(codePoint) is "Hani" or null ? HanScript(fonts.Language) : EastAsianScriptOf(codePoint);
            WordFontMetrics? face = Face(codePoint, slot, fonts, script, eastAsianScript);
            if (face is not null && !WordFontMetrics.IsNonSpacingMark(codePoint)) lineHeight = Math.Max(lineHeight, slot == Slot.EastAsian ? face.LineHeightEm * EastAsianLineFactor : face.LineHeightEm);
            if (codePoint == ' ')
            {
                double advance = face?.Advance(codePoint) ?? 0.5;
                if (inWord) { EndWord(); space = advance; }
                else if (space == 0 && words.Count > 0) space = advance;
                continue;
            }
            // Japanese / Chinese: a line may break between two characters here, without a space
            if (inWord && IsIdeographicBreak(codePoints[i - 1], codePoint))
            {
                EndWord();
                spaces.Add(0); breaks.Add(false);
                word = wordKerning = 0; inWord = true; segmentFace = null;
            }
            if (!inWord)
            {
                spaces.Add(words.Count == 0 || breakPending ? 0 : space);
                breaks.Add(breakPending && words.Count > 0);
                breakPending = false;
                word = wordKerning = 0; inWord = true; space = 0; segmentFace = null;
            }
            bool complex = slot == Slot.Complex;
            string openTypeScript = OpenTypeScript(complex ? script : null, codePoint, slot);
            if (!ReferenceEquals(face, segmentFace) || complex != segmentComplex || openTypeScript != segmentScript)
            {
                FlushSegment();
                segmentFace = face; segmentComplex = complex; segmentScript = openTypeScript;
            }
            segment.Add(codePoint);
        }
        if (inWord) EndWord();
        if (lineHeight == 0) lineHeight = LatinFaces(null, false, false).FirstOrDefault()?.LineHeightEm ?? 1.2207;
        return new MeasuredParagraph([.. words], [.. kerning], [.. spaces], [.. breaks], lineHeight);
    }

    /// <summary>The face a character is drawn in: the first of its slot's chain that has the glyph, then the other slots' chains, then the slot's first face.</summary>
    private WordFontMetrics? Face(int codePoint, Slot slot, RunFonts fonts, string? script, string? eastAsianScript)
    {
        IEnumerable<WordFontMetrics> latin = LatinFaces(fonts.Latin, fonts.Bold, fonts.Italic);
        IEnumerable<WordFontMetrics> complex = ComplexFaces(fonts.Complex, script, fonts.Bold, fonts.Italic);
        IEnumerable<WordFontMetrics> eastAsian = EastAsianFaces(fonts.EastAsian, eastAsianScript ?? HanScript(fonts.Language), fonts.Bold, fonts.Italic);
        IEnumerable<WordFontMetrics>[] chains = slot switch
        {
            Slot.Complex => [complex, latin, eastAsian],
            Slot.EastAsian => [eastAsian, latin, complex],
            _ => [latin, complex, eastAsian],
        };
        foreach (IEnumerable<WordFontMetrics> chain in chains)
            foreach (WordFontMetrics face in chain)
                if (face.Has(codePoint)) return face;
        return chains[0].FirstOrDefault();
    }

    private IEnumerable<WordFontMetrics> LatinFaces(string? typeface, bool bold, bool italic)
    {
        foreach (string? name in new[] { Resolve(typeface, null), theme?.LatinFont(false), "Calibri" })
            if (WordFontMetrics.Of(name, bold, italic) is { } face) yield return face;
    }

    private IEnumerable<WordFontMetrics> ComplexFaces(string? typeface, string? script, bool bold, bool italic)
    {
        foreach (string? name in new[] { Resolve(typeface, script), theme?.ComplexScriptFont(false, script), "Arial", "Times New Roman" })
            if (WordFontMetrics.Of(name, bold, italic) is { } face) yield return face;
    }

    /// <summary>The run's a:ea, the theme's East Asian font for the script, then the faces Windows ships for it (a theme face that is not installed, such as the Office theme's Jpan Yu Mincho, hands on: Word drew japanese-block-list in Yu Gothic).</summary>
    private IEnumerable<WordFontMetrics> EastAsianFaces(string? typeface, string script, bool bold, bool italic)
    {
        string[] defaults = script switch
        {
            "Hang" => ["Malgun Gothic", "Gulim"],
            "Hans" => ["SimSun", "Microsoft YaHei"],
            "Hant" => ["PMingLiU", "Microsoft JhengHei"],
            _ => ["Yu Gothic", "MS Gothic", "Meiryo"],
        };
        foreach (string? name in new[] { Resolve(typeface, script), theme?.EastAsianFont(false, script) }.Concat(defaults).Concat(["Yu Gothic", "SimSun", "Malgun Gothic"]))
            if (WordFontMetrics.Of(name, bold, italic) is { } face) yield return face;
    }

    /// <summary>A run typeface: a theme reference (+mn-lt / +mj-lt / +mn-cs / +mj-cs / +mn-ea / +mj-ea) resolved, an empty one = none.</summary>
    private string? Resolve(string? typeface, string? script) => typeface switch
    {
        null or "" => null,
        "+mn-lt" => theme?.LatinFont(false),
        "+mj-lt" => theme?.LatinFont(true),
        "+mn-cs" => theme?.ComplexScriptFont(false, script),
        "+mj-cs" => theme?.ComplexScriptFont(true, script),
        "+mn-ea" => theme?.EastAsianFont(false, script),
        "+mj-ea" => theme?.EastAsianFont(true, script),
        _ when typeface.StartsWith("+", StringComparison.Ordinal) => null,
        _ => typeface,
    };

    /// <summary>The script a Han character takes its font by: the run's language (ja Jpan, ko Hang, zh-TW / HK / MO Hant, zh Hans), else w:themeFontLang/@eastAsia's, else Jpan.</summary>
    private string HanScript(string? language)
        => DocxTheme.ScriptTagOfLanguage(language) is { } own && own is "Jpan" or "Hans" or "Hant" or "Hang" ? own : theme?.EastAsianScript ?? "Jpan";

    /// <summary>The OpenType script tag a run of characters is shaped with.</summary>
    private static string OpenTypeScript(string? complexScript, int codePoint, Slot slot) => complexScript switch
    {
        "Hebr" => "hebr", "Arab" => "arab", "Syrc" => "syrc", "Thaa" => "thaa", "Thai" => "thai",
        _ => slot == Slot.EastAsian ? EastAsianScriptOf(codePoint) switch { "Hang" => "hang", "Jpan" => "kana", _ => "hani" } : "latn",
    };

    /// <summary>The ISO 15924 tag of a complex-script character (the theme's supplemental font key); null for any other character.</summary>
    internal static string? ComplexScriptOf(int c) => c switch
    {
        >= 0x0590 and <= 0x05FF or >= 0xFB1D and <= 0xFB4F => "Hebr",
        >= 0x0600 and <= 0x06FF or >= 0x0750 and <= 0x077F or >= 0x08A0 and <= 0x08FF or >= 0xFB50 and <= 0xFDFF or >= 0xFE70 and <= 0xFEFF => "Arab",
        >= 0x0700 and <= 0x074F => "Syrc",
        >= 0x0780 and <= 0x07BF => "Thaa",
        >= 0x0E00 and <= 0x0E7F => "Thai",
        _ => null,
    };

    /// <summary>The script of an East Asian character: "Jpan" (kana), "Hang" (Hangul), "Hani" (Han, CJK punctuation and symbols, full-width forms); null for any other character.</summary>
    internal static string? EastAsianScriptOf(int c) => c switch
    {
        >= 0x3040 and <= 0x30FF or >= 0x31F0 and <= 0x31FF or >= 0xFF66 and <= 0xFF9F => "Jpan",
        >= 0x1100 and <= 0x11FF or >= 0x3130 and <= 0x318F or >= 0xA960 and <= 0xA97F or >= 0xAC00 and <= 0xD7FF or >= 0xFFA0 and <= 0xFFDC => "Hang",
        >= 0x2E80 and <= 0x2FDF or >= 0x3000 and <= 0x303F or >= 0x3190 and <= 0x31EF or >= 0x3200 and <= 0x9FFF or >= 0xF900 and <= 0xFAFF
            or >= 0xFE30 and <= 0xFE4F or >= 0xFF00 and <= 0xFF65 or >= 0xFFE0 and <= 0xFFEE or >= 0x20000 and <= 0x3FFFF => "Hani",
        _ => null,
    };

    /// <summary>
    /// Whether a line may break between two characters of Japanese / Chinese text (one of them Han, kana or CJK punctuation; not
    /// Hangul): never before closing punctuation, small kana, iteration marks or the prolonged sound mark, nor after opening
    /// punctuation (Word's default kinsoku characters, Options > Language > Asian typography).
    /// </summary>
    private static bool IsIdeographicBreak(int before, int after)
    {
        static bool Ideographic(int c) => EastAsianScriptOf(c) is "Jpan" or "Hani";
        if (!Ideographic(before) && !Ideographic(after)) return false;
        const string noBreakBefore = "!%),.:;?]}\u00A2\u00B0\u2019\u201D\u2030\u2032\u2033\u2103\u3001\u3002\u3005\u3009\u300B\u300D\u300F\u3011\u3015\u3017\u3019\u301B\u301F\u3041\u3043\u3045\u3047\u3049\u3063\u3083\u3085\u3087\u308E\u309B\u309C\u309D\u309E\u30A1\u30A3\u30A5\u30A7\u30A9\u30C3\u30E3\u30E5\u30E7\u30EE\u30F5\u30F6\u30FB\u30FC\u30FD\u30FE\uFF01\uFF05\uFF09\uFF0C\uFF0E\uFF1A\uFF1B\uFF1F\uFF3D\uFF5D\uFF61\uFF63\uFF64\uFF65\uFF9E\uFF9F";
        const string noBreakAfter = "$([{\u00A3\u00A5\u2018\u201C\u3008\u300A\u300C\u300E\u3010\u3014\u3016\u3018\u301A\u301D\uFF04\uFF08\uFF3B\uFF5B\uFF62\uFFE1\uFFE5";
        return !(after <= 0xFFFF && noBreakBefore.IndexOf((char)after) >= 0) && !(before <= 0xFFFF && noBreakAfter.IndexOf((char)before) >= 0);
    }

    private static bool IsLetter(int c) => c > 0xFFFF || char.IsLetter((char)c);
}
