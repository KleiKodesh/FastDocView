using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// Measures DrawingML paragraphs (a:p with a:r / a:fld runs, namespace-free) in the fonts Word draws them with, for SmartArt
/// autofit. Each character takes a font slot as Office text does: complex-script characters (Hebrew, Arabic, Syriac, Thaana,
/// Thai, ... ) the run's a:cs typeface, else the theme's complex-script font for that script (DocxTheme.ComplexScriptFont: a:cs,
/// else the supplemental a:font, the Office theme's "Hebr" = Arial), else Arial / Times New Roman; other letters the run's
/// a:latin typeface, else the theme's minor Latin font, else Calibri ("+mn-lt" / "+mj-cs" ... name the theme's). Spaces, digits
/// and punctuation take the slot of the strong character before them (after them at the start). A face without the
/// character's glyph hands it to the next one in its chain, then to the other slot's chain. Combining marks (niqqud,
/// cantillation) add no width. Widths and pair kerning come from the face's own font file (WordFontMetrics); kerning applies
/// between Latin characters of one face (Word kerns SmartArt text, kern="1200", and shapes complex scripts without the kern
/// table). One instance per layout pass; each paragraph is measured once (words in em at size 1), so fitting a size is cheap.
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

    private enum Slot : byte { Neutral, Latin, Complex, Break }

    private MeasuredParagraph MeasureNew(XElement paragraph)
    {
        // the paragraph's characters with their run's typefaces
        var codePoints = new List<int>();
        var slots = new List<Slot>();
        var runs = new List<(string? Latin, string? Complex)>();
        var runOf = new List<int>();
        foreach (XElement run in paragraph.Elements())
        {
            if (run.Name.LocalName == "br") { codePoints.Add('\n'); slots.Add(Slot.Break); runOf.Add(-1); continue; }
            if (run.Name.LocalName is not ("r" or "fld")) continue;
            XElement? properties = run.Element("rPr");
            runs.Add(((string?)properties?.Element("latin")?.Attribute("typeface"), (string?)properties?.Element("cs")?.Attribute("typeface")));
            string text = (string?)run.Element("t") ?? "";
            for (int i = 0; i < text.Length; i++)
            {
                int codePoint = text[i];
                if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) codePoint = char.ConvertToUtf32(text[i], text[++i]);
                codePoints.Add(codePoint);
                slots.Add(ComplexScriptOf(codePoint) is not null ? Slot.Complex : IsLetter(codePoint) ? Slot.Latin : Slot.Neutral);
                runOf.Add(runs.Count - 1);
            }
        }
        // neutrals take the slot of the strong character before them, else after them, else Latin
        Slot last = Slot.Neutral;
        for (int i = 0; i < slots.Count; i++)
            if (slots[i] is Slot.Latin or Slot.Complex) last = slots[i];
            else if (slots[i] == Slot.Neutral && last != Slot.Neutral) slots[i] = last;
        Slot next = Slot.Latin;
        for (int i = slots.Count - 1; i >= 0; i--)
            if (slots[i] is Slot.Latin or Slot.Complex) next = slots[i];
            else if (slots[i] == Slot.Neutral) slots[i] = next;

        var words = new List<double>();
        var kerning = new List<double>();
        var spaces = new List<double>();
        var breaks = new List<bool>();
        double lineHeight = 0, word = 0, wordKerning = 0, space = 0;
        bool inWord = false, breakPending = false;
        WordFontMetrics? previousFace = null;
        int previous = -1;
        string? script = null;      // the script of the last complex-script character: neutrals after it use that script's font
        for (int i = 0; i < codePoints.Count; i++)
        {
            int codePoint = codePoints[i];
            if (slots[i] == Slot.Break)
            {
                if (inWord) { words.Add(word); kerning.Add(wordKerning); inWord = false; }
                else
                {
                    // an empty line (a break at the start or right after another): a word of no width holds its line
                    words.Add(0); kerning.Add(0); spaces.Add(0); breaks.Add(breakPending && words.Count > 1);
                }
                breakPending = true; space = 0; previousFace = null;
                continue;
            }
            bool complex = slots[i] == Slot.Complex;
            if (ComplexScriptOf(codePoint) is { } own) script = own;
            (string? latin, string? complexFace) = runs[runOf[i]];
            WordFontMetrics? face = Face(codePoint, complex, latin, complexFace, script);
            if (face is not null) lineHeight = Math.Max(lineHeight, face.LineHeightEm);
            if (WordFontMetrics.IsNonSpacingMark(codePoint)) continue;
            double advance = face?.Advance(codePoint) ?? 0.5;
            if (codePoint == ' ')
            {
                if (inWord) { words.Add(word); kerning.Add(wordKerning); inWord = false; space = advance; }
                else if (space == 0 && words.Count > 0) space = advance;
                previousFace = null;
                continue;
            }
            if (!inWord)
            {
                spaces.Add(words.Count == 0 || breakPending ? 0 : space);
                breaks.Add(breakPending && words.Count > 0);
                breakPending = false;
                word = wordKerning = 0; inWord = true; space = 0; previousFace = null;
            }
            word += advance;
            if (!complex && face is not null && ReferenceEquals(face, previousFace)) wordKerning += face.Kerning(previous, codePoint);
            previousFace = complex ? null : face;
            previous = codePoint;
        }
        if (inWord) { words.Add(word); kerning.Add(wordKerning); }
        if (lineHeight == 0) lineHeight = LatinFaces(null).FirstOrDefault()?.LineHeightEm ?? 1.2207;
        return new MeasuredParagraph([.. words], [.. kerning], [.. spaces], [.. breaks], lineHeight);
    }

    /// <summary>The face a character is drawn in: the first of its slot's chain that has the glyph, then the other slot's chain, then the slot's first face.</summary>
    private WordFontMetrics? Face(int codePoint, bool complex, string? latin, string? complexFace, string? script)
    {
        IEnumerable<WordFontMetrics> own = complex ? ComplexFaces(complexFace, script) : LatinFaces(latin);
        IEnumerable<WordFontMetrics> other = complex ? LatinFaces(latin) : ComplexFaces(complexFace, script);
        foreach (WordFontMetrics face in own) if (face.Has(codePoint)) return face;
        foreach (WordFontMetrics face in other) if (face.Has(codePoint)) return face;
        return own.FirstOrDefault();
    }

    private IEnumerable<WordFontMetrics> LatinFaces(string? typeface)
    {
        foreach (string? name in new[] { Resolve(typeface, null), theme?.LatinFont(false), "Calibri" })
            if (WordFontMetrics.Of(name) is { } face) yield return face;
    }

    private IEnumerable<WordFontMetrics> ComplexFaces(string? typeface, string? script)
    {
        foreach (string? name in new[] { Resolve(typeface, script), theme?.ComplexScriptFont(false, script), "Arial", "Times New Roman" })
            if (WordFontMetrics.Of(name) is { } face) yield return face;
    }

    /// <summary>A run typeface: a theme reference (+mn-lt / +mj-lt / +mn-cs / +mj-cs) resolved, an empty one = none.</summary>
    private string? Resolve(string? typeface, string? script) => typeface switch
    {
        null or "" => null,
        "+mn-lt" => theme?.LatinFont(false),
        "+mj-lt" => theme?.LatinFont(true),
        "+mn-cs" => theme?.ComplexScriptFont(false, script),
        "+mj-cs" => theme?.ComplexScriptFont(true, script),
        _ when typeface.StartsWith("+", StringComparison.Ordinal) => null,
        _ => typeface,
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

    private static bool IsLetter(int c) => c > 0xFFFF || char.IsLetter((char)c);
}
