using System.Globalization;
using System.Text;
using Glyph = FastDocView.Html.WordFontLayoutTables.Glyph;

namespace FastDocView.Html;

/// <summary>
/// The Indic shaping engine for measuring text (Devanagari, Bengali, Gurmukhi, Gujarati, Oriya, Tamil, Telugu, Kannada,
/// Malayalam), as Microsoft's "Developing OpenType fonts for Devanagari script" (and its sister specs) describes and HarfBuzz
/// implements it for Uniscribe compatibility (hb-ot-shaper-indic.cc, 11.2.1; line numbers below): split matras decomposed,
/// syllables found, each syllable's base consonant found and its characters reordered (reph, pre-base matras, half forms), the
/// masked features (rphf, half, blwf, pstf, pref) given to the glyphs the reordering marks, the basic features applied one
/// stage at a time, then final reordering (pre-base matra and reph moved, init on a word-initial left matra) and the
/// presentation features. Reordering itself moves glyphs only; it changes the advance total through which substitutions
/// apply (a reph or below-base form takes no width, a half form is narrower, presentation ligatures need their glyphs
/// side by side). Not done: dotted circles for broken clusters, Unicode mark reordering. Character classes: HarfBuzz's table generated from Unicode 16
/// IndicSyllabicCategory / IndicPositionalCategory (hb-ot-shaper-indic-table.cc, U+0900..U+0D7F).
/// </summary>
internal static class WordIndicShaper
{
    // syllabic categories (hb-ot-shaper-indic-machine.rl:58-77)
    private const byte X = 0, C = 1, V = 2, N = 3, H = 4, Zwnj = 5, Zwj = 6, M = 7, SM = 8, A = 9, Placeholder = 10, DottedCircle = 11,
                       RS = 12, MPst = 13, Repha = 14, Ra = 15, CM = 16, Symbol = 17, CS = 18;

    // visual positions (hb-ot-shaper-indic.hh:37-62)
    private const byte RaToBecomeReph = 1, PreM = 2, PreC = 3, BaseC = 4, AfterMain = 5, BeforeSub = 7, BelowC = 8, AfterSub = 9,
                       BeforePost = 10, PostC = 11, AfterPost = 12, Smvd = 13, End = 14;

    private enum RephMode { Implicit, Explicit, LogicalRepha }

    /// <summary>A script's configuration (hb-ot-shaper-indic.cc:149-161): OpenType tags (new spec, old spec), virama, where the reph goes, how it is formed, whether blwf also applies before the base.</summary>
    private sealed record Script(string NewTag, string OldTag, int Virama, byte RephPosition, RephMode Reph, bool BelowFormsPostOnly);

    private static readonly Script[] Scripts =
    [
        new("dev2", "deva", 0x094D, BeforePost, RephMode.Implicit, false),
        new("bng2", "beng", 0x09CD, AfterSub, RephMode.Implicit, false),
        new("gur2", "guru", 0x0A4D, BeforeSub, RephMode.Implicit, false),
        new("gjr2", "gujr", 0x0ACD, BeforePost, RephMode.Implicit, false),
        new("ory2", "orya", 0x0B4D, AfterMain, RephMode.Implicit, false),
        new("tml2", "taml", 0x0BCD, AfterPost, RephMode.Implicit, false),
        new("tel2", "telu", 0x0C4D, AfterPost, RephMode.Explicit, true),
        new("knd2", "knda", 0x0CCD, AfterPost, RephMode.Implicit, true),
        new("mlm2", "mlym", 0x0D4D, AfterMain, RephMode.LogicalRepha, false),
    ];

    /// <summary>Whether an OpenType script tag (old-spec form, as <c>DocxDrawingTextFonts</c> passes it) is shaped here.</summary>
    public static bool Handles(string openTypeScript) => Array.FindIndex(Scripts, s => s.OldTag == openTypeScript) >= 0;

    /// <summary>The GPOS script tag for positioning a script's text: the new-spec tag when GPOS has it, else the old one.</summary>
    public static string PositioningTag(WordFontLayoutTables layout, string openTypeScript)
    {
        Script script = Scripts.First(s => s.OldTag == openTypeScript);
        return layout.HasPositioningScript(script.NewTag) ? script.NewTag : script.OldTag;
    }

    /// <summary>
    /// The characters as a shaper decomposes them (hb-ot-shaper-indic.cc:1510-1555 decompose_indic / compose_indic over the
    /// COMPOSED_DIACRITICS_NO_SHORT_CIRCUIT normalizer): split matras (canonical decompositions starting with a mark) and the
    /// composition exclusions (DEVANAGARI LETTER QA ...) are decomposed when the font has the parts; RRA (U+0931, U+09DC,
    /// U+09DD), TAMIL AU and BENGALI YYA stay.
    /// </summary>
    public static List<int> Decompose(IReadOnlyList<int> codePoints, Func<int, bool> fontHas)
    {
        var result = new List<int>(codePoints.Count + 4);
        foreach (int codePoint in codePoints)
        {
            if (codePoint is < 0x0900 or > 0x0D7F or 0x0931 or 0x09DC or 0x09DD or 0x0B94 or 0x09DF) { result.Add(codePoint); continue; }
            string text = char.ConvertFromUtf32(codePoint);
            string decomposed = text.Normalize(NormalizationForm.FormD);
            bool split = decomposed.Length > 1 && IsMark(decomposed[0]);
            bool excluded = decomposed.Length > 1 && decomposed.Normalize(NormalizationForm.FormC) != text;
            if ((split || excluded) && decomposed.All(c => fontHas(c))) foreach (char part in decomposed) result.Add(part);
            else result.Add(codePoint);
        }
        return result;
    }

    private static bool IsMark(char c) => CharUnicodeInfo.GetUnicodeCategory(c) is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark;

    /// <summary>
    /// Shapes a buffer of one Indic script's glyphs (one per character of <see cref="Decompose"/>d text, CodePoint set) through
    /// the font's GSUB in place.
    /// </summary>
    public static void Shape(WordFontLayoutTables layout, List<Glyph> buffer, string openTypeScript, Func<int, int> glyphOf)
    {
        Script script = Scripts.First(s => s.OldTag == openTypeScript);
        // the new-spec tag when the font has it, else the old spec (hb-ot-shaper-indic.cc:325 is_old_spec)
        bool oldSpec = !layout.HasScript(script.NewTag);
        string tag = oldSpec ? script.OldTag : script.NewTag;
        for (int i = 0; i < buffer.Count; i++)
        {
            (byte category, byte position) = Classes(buffer[i].CodePoint);
            buffer[i] = buffer[i] with { Category = category, Position = position, Mask = 0 };
        }
        var types = FindSyllables(buffer);

        layout.Substitute(buffer, tag, [["locl", "ccmp"]]);

        // consonant positions from the font: a consonant with a below-base / post-base form (hb-ot-shaper-indic.cc:365-397, 441-458)
        int virama = glyphOf(script.Virama);
        if (virama != 0)
        {
            var known = new Dictionary<int, byte>();
            for (int i = 0; i < buffer.Count; i++)
                if (buffer[i].Position == BaseC)
                {
                    int consonant = buffer[i].Id;
                    if (!known.TryGetValue(consonant, out byte position))
                        known[consonant] = position = ConsonantPosition(layout, tag, consonant, virama);
                    buffer[i] = buffer[i] with { Position = position };
                }
        }

        bool rephForms = layout.HasFeature(tag, "rphf"), preBaseForms = layout.HasFeature(tag, "pref");
        foreach ((int start, int end) in Syllables(buffer))
            if (types.TryGetValue(buffer[start].Syllable, out SyllableType type) && type is not (SyllableType.Symbol or SyllableType.NonIndic))
                InitialReordering(layout, tag, script, oldSpec, rephForms, preBaseForms, buffer, start, end);

        // the basic features, each a stage of its own (hb-ot-shaper-indic.cc:166-183, 245-262)
        foreach (string feature in new[] { "nukt", "akhn", "rphf", "rkrf", "pref", "blwf", "abvf", "half", "pstf", "vatu", "cjct" })
            layout.Substitute(buffer, tag, [[feature]]);

        foreach ((int start, int end) in Syllables(buffer))
            if (types.TryGetValue(buffer[start].Syllable, out SyllableType type) && type is not (SyllableType.Symbol or SyllableType.NonIndic))
                FinalReordering(script, preBaseForms, buffer, start, end);

        // presentation features with the default ones, 'liga' off (hb-ot-shaper-indic.cc:184-196, 268-272; hb-ot-shape.cc:295-317 common and horizontal features)
        layout.Substitute(buffer, tag, [["init", "pres", "abvs", "blws", "psts", "haln", "rlig", "calt", "clig", "rclt"]]);
    }

    private static byte ConsonantPosition(WordFontLayoutTables layout, string tag, int consonant, int virama)
    {
        // new spec matches virama, consonant; old spec consonant, virama; fonts carry either, so both are tried
        bool Would(string feature) => layout.WouldSubstitute(tag, feature, virama, consonant) || layout.WouldSubstitute(tag, feature, consonant, virama);
        if (Would("blwf") || Would("vatu")) return BelowC;
        if (Would("pstf") || Would("pref")) return PostC;
        return BaseC;
    }

    private static (byte Category, byte Position) Classes(int codePoint)
    {
        if (codePoint is >= 0x0900 and <= 0x0D7F)
        {
            int at = (codePoint - 0x0900) * 2;
            return ((byte)(Table[at] - 'a'), (byte)(Table[at + 1] - 'a'));
        }
        return codePoint switch
        {
            0x200C => (Zwnj, End),
            0x200D => (Zwj, End),
            0x25CC => (DottedCircle, BaseC),
            0x00A0 or (>= 0x0030 and <= 0x0039) or 0x002D or 0x00D7 or (>= 0x2010 and <= 0x2015) => (Placeholder, BaseC),
            _ => (X, End),
        };
    }

    private static bool Is(Glyph glyph, params byte[] categories) => !glyph.Ligated && Array.IndexOf(categories, glyph.Category) >= 0;
    private static bool IsConsonant(Glyph glyph) => Is(glyph, C, CS, Ra, CM, V, Placeholder, DottedCircle);
    private static bool IsJoiner(Glyph glyph) => Is(glyph, Zwj, Zwnj);
    private static bool IsHalant(Glyph glyph) => Is(glyph, H);

    // ── syllables (hb-ot-shaper-indic-machine.rl:80-111) ──────────────────────────

    private enum SyllableType { Consonant, Vowel, Standalone, Symbol, Broken, NonIndic }

    /// <summary>Numbers the syllables (Glyph.Syllable from 1) and returns each one's type.</summary>
    private static Dictionary<int, SyllableType> FindSyllables(List<Glyph> buffer)
    {
        var types = new Dictionary<int, SyllableType>();
        int serial = 0;
        for (int start = 0; start < buffer.Count;)
        {
            (int end, SyllableType type) = Syllable(buffer, start);
            serial++;
            types[serial] = type;
            for (int i = start; i < end; i++) buffer[i] = buffer[i] with { Syllable = serial };
            start = end;
        }
        return types;
    }

    private static IEnumerable<(int Start, int End)> Syllables(List<Glyph> buffer)
    {
        for (int start = 0; start < buffer.Count;)
        {
            int end = start + 1;
            while (end < buffer.Count && buffer[end].Syllable == buffer[start].Syllable) end++;
            yield return (start, end);
            start = end;
        }
    }

    /// <summary>The longest syllable at <paramref name="start"/> (the scanner's alternatives in order; a tie goes to the first).</summary>
    private static (int End, SyllableType Type) Syllable(List<Glyph> buffer, int start)
    {
        int Cat(int i) => i < buffer.Count ? buffer[i].Category : -1;
        bool Joiner(int i) => Cat(i) is Zwj or Zwnj;
        bool Modifier(int i) => Cat(i) is SM;
        // n = ((ZWNJ? RS)? (N N?)?)
        int Modifiers(int i)
        {
            if (Cat(i) == Zwnj && Cat(i + 1) == RS) i += 2;
            else if (Cat(i) == RS) i++;
            if (Cat(i) == N) { i++; if (Cat(i) == N) i++; }
            return i;
        }
        // cn = (C | Ra) ZWJ? n?
        int ConsonantGroup(int i)
        {
            if (Cat(i) is not (C or Ra)) return -1;
            i++;
            if (Cat(i) == Zwj) i++;
            return Modifiers(i);
        }
        // halant_group = z? H (ZWJ N?)?
        int HalantGroup(int i)
        {
            if (Joiner(i)) i++;
            if (Cat(i) != H) return -1;
            i++;
            if (Cat(i) == Zwj) { i++; if (Cat(i) == N) i++; }
            return i;
        }
        // matra_group = z* (M | sm? MPst) N? H?
        int MatraGroup(int i)
        {
            while (Joiner(i)) i++;
            if (Cat(i) == M) i++;
            else
            {
                if (Modifier(i)) i++;
                if (Cat(i) != MPst) return -1;
                i++;
            }
            if (Cat(i) == N) i++;
            if (Cat(i) == H) i++;
            return i;
        }
        // complex_syllable_tail = (halant_group cn)* medial_group halant_or_matra_group syllable_tail
        int Tail(int i)
        {
            while (HalantGroup(i) is var h and >= 0 && ConsonantGroup(h) is var next and >= 0) i = next;
            if (Cat(i) == CM) i++;
            int finalHalant = Math.Max(HalantGroup(i), Cat(i) == H && Cat(i + 1) == Zwnj ? i + 2 : -1);
            if (finalHalant >= 0) i = finalHalant;
            else while (MatraGroup(i) is var m and >= 0 && m > i) i = m;
            return SyllableTail(i);
        }
        // syllable_tail = (z? sm sm? ZWNJ?)? (A | VD)*
        int SyllableTail(int i)
        {
            int j = Joiner(i) ? i + 1 : i;
            if (Modifier(j))
            {
                j++;
                if (Modifier(j)) j++;
                if (Cat(j) == Zwnj) j++;
                i = j;
            }
            while (Cat(i) == A) i++;
            return i;
        }
        int Reph(int i) => Cat(i) == Ra && Cat(i + 1) == H ? i + 2 : Cat(i) == Repha ? i + 1 : i;

        var best = (End: start + 1, Type: SyllableType.NonIndic);
        void Consider(int end, SyllableType type) { if (end > best.End || end == best.End && best.Type == SyllableType.NonIndic && end > start) best = (end, type); }
        // consonant_syllable = (Repha | CS)? cn complex_syllable_tail
        {
            int i = Cat(start) is Repha or CS ? start + 1 : start;
            if (ConsonantGroup(i) is var c and >= 0) Consider(Tail(c), SyllableType.Consonant);
        }
        // vowel_syllable = reph? V n? (ZWJ | complex_syllable_tail)
        {
            int i = Reph(start);
            if (Cat(i) == V)
            {
                i = Modifiers(i + 1);
                Consider(Math.Max(Cat(i) == Zwj ? i + 1 : -1, Tail(i)), SyllableType.Vowel);
            }
        }
        // standalone_cluster = ((Repha | CS)? PLACEHOLDER | reph? DOTTEDCIRCLE) n? complex_syllable_tail
        {
            int i = Cat(start) is Repha or CS ? start + 1 : start;
            int j = Reph(start);
            int at = Cat(i) == Placeholder ? i + 1 : Cat(j) == DottedCircle ? j + 1 : -1;
            if (at >= 0) Consider(Tail(Modifiers(at)), SyllableType.Standalone);
        }
        // symbol_cluster = Symbol N? syllable_tail
        if (Cat(start) == Symbol) Consider(SyllableTail(Cat(start + 1) == N ? start + 2 : start + 1), SyllableType.Symbol);
        // broken_cluster = reph? n? complex_syllable_tail (not empty)
        {
            int end = Tail(Modifiers(Reph(start)));
            if (end > start) Consider(end, SyllableType.Broken);
        }
        return best;
    }

    // ── initial reordering (hb-ot-shaper-indic.cc:466-937) ─────────────────────────

    private static void InitialReordering(WordFontLayoutTables layout, string tag, Script script, bool oldSpec, bool rephForms, bool preBaseForms,
                                          List<Glyph> b, int start, int end)
    {
        void Set(int i, Func<Glyph, Glyph> change) => b[i] = change(b[i]);
        bool kannada = script.OldTag == "knda", devanagari = script.OldTag == "deva";
        // Kannada Ra, H, ZWJ behaves as Ra, ZWJ, H (:476-484)
        if (kannada && start + 3 <= end && Is(b[start], Ra) && Is(b[start + 1], H) && Is(b[start + 2], Zwj))
            (b[start + 1], b[start + 2]) = (b[start + 2], b[start + 1]);

        // 1. the base consonant (:488-600)
        int baseAt = end;
        bool hasReph = false;
        int limit = start;
        if (rephForms && start + 3 <= end
            && (script.Reph == RephMode.Implicit && !IsJoiner(b[start + 2]) || script.Reph == RephMode.Explicit && b[start + 2].Category == Zwj))
        {
            if (layout.WouldSubstitute(tag, "rphf", b[start].Id, b[start + 1].Id)
                || script.Reph == RephMode.Explicit && layout.WouldSubstitute(tag, "rphf", b[start].Id, b[start + 1].Id, b[start + 2].Id))
            {
                limit += 2;
                while (limit < end && IsJoiner(b[limit])) limit++;
                baseAt = start;
                hasReph = true;
            }
        }
        else if (script.Reph == RephMode.LogicalRepha && b[start].Category == Repha)
        {
            limit += 1;
            while (limit < end && IsJoiner(b[limit])) limit++;
            baseAt = start;
            hasReph = true;
        }
        {
            int i = end;
            bool seenBelow = false;
            do
            {
                i--;
                if (IsConsonant(b[i]))
                {
                    if (b[i].Position != BelowC && (b[i].Position != PostC || seenBelow)) { baseAt = i; break; }
                    if (b[i].Position == BelowC) seenBelow = true;
                    baseAt = i;
                }
                else if (start < i && b[i].Category == Zwj && b[i - 1].Category == H) break;
            } while (i > limit);
        }
        if (hasReph && baseAt == start && limit - baseAt <= 2) hasReph = false;

        // reorder characters (:632-833)
        for (int i = start; i < baseAt; i++) Set(i, g => g with { Position = Math.Min(PreC, g.Position) });
        if (baseAt < end) Set(baseAt, g => g with { Position = BaseC });
        if (hasReph) Set(start, g => g with { Position = RaToBecomeReph });
        if (oldSpec)
        {
            // old spec: the first post-base halant moves after the last consonant (:671-690)
            bool noDoubleHalants = kannada;
            for (int i = baseAt + 1; i < end; i++)
                if (b[i].Category == H)
                {
                    int j;
                    for (j = end - 1; j > i; j--)
                        if (IsConsonant(b[j]) || noDoubleHalants && b[j].Category == H) break;
                    if (b[j].Category != H && j > i)
                    {
                        Glyph halant = b[i];
                        b.RemoveAt(i);
                        b.Insert(j, halant);
                    }
                    break;
                }
        }
        // misc marks move with the character before them (:693-723)
        {
            byte lastPosition = 0;
            for (int i = start; i < end; i++)
            {
                if (Array.IndexOf(new[] { Zwj, Zwnj, N, RS, CM, H }, b[i].Category) >= 0)
                {
                    Set(i, g => g with { Position = lastPosition });
                    if (b[i].Category == H && b[i].Position == PreM)
                        for (int j = i; j > start; j--)
                            if (b[j - 1].Position != PreM) { byte p = b[j - 1].Position; Set(i, g => g with { Position = p }); break; }
                }
                else if (b[i].Position != Smvd)
                {
                    if (b[i].Category == MPst && i > start && b[i - 1].Category == SM) { byte p = b[i].Position; Set(i - 1, g => g with { Position = p }); }
                    lastPosition = b[i].Position;
                }
            }
        }
        // post-base consonants own what comes before them since the last consonant or matra (:724-737)
        {
            int last = baseAt;
            for (int i = baseAt + 1; i < end; i++)
                if (IsConsonant(b[i]))
                {
                    for (int j = last + 1; j < i; j++)
                        if (b[j].Position < Smvd) { byte p = b[i].Position; Set(j, g => g with { Position = p }); }
                    last = i;
                }
                else if (b[i].Category is M or MPst) last = i;
        }
        // stable sort by position, find the base again, flip a run of left matras (:739-781)
        {
            List<Glyph> sorted = b.GetRange(start, end - start).Select((g, k) => (g, k)).OrderBy(x => x.g.Position).ThenBy(x => x.k).Select(x => x.g).ToList();
            for (int i = 0; i < sorted.Count; i++) b[start + i] = sorted[i];
            int firstLeft = end, lastLeft = end;
            baseAt = end;
            for (int i = start; i < end; i++)
            {
                if (b[i].Position == BaseC) { baseAt = i; break; }
                if (b[i].Position == PreM)
                {
                    if (firstLeft == end) firstLeft = i;
                    lastLeft = i;
                }
            }
            if (firstLeft < lastLeft)
            {
                b.Reverse(firstLeft, lastLeft + 1 - firstLeft);
                int i = firstLeft;
                for (int j = i; j <= lastLeft; j++)
                    if (b[j].Category is M or MPst)
                    {
                        b.Reverse(i, j + 1 - i);
                        i = j + 1;
                    }
            }
        }

        // masks (:841-933)
        uint rphf = WordFontLayoutTables.FeatureBit("rphf"), half = WordFontLayoutTables.FeatureBit("half"), blwf = WordFontLayoutTables.FeatureBit("blwf"),
             abvf = WordFontLayoutTables.FeatureBit("abvf"), pstf = WordFontLayoutTables.FeatureBit("pstf"), pref = WordFontLayoutTables.FeatureBit("pref");
        for (int i = start; i < end && b[i].Position == RaToBecomeReph; i++) Set(i, g => g with { Mask = g.Mask | rphf });
        uint preBase = half | (!oldSpec && !script.BelowFormsPostOnly ? blwf : 0);
        for (int i = start; i < baseAt; i++) Set(i, g => g with { Mask = g.Mask | preBase });
        for (int i = baseAt + 1; i < end; i++) Set(i, g => g with { Mask = g.Mask | blwf | abvf | pstf });
        if (oldSpec && devanagari)
        {
            // old-spec eyelash Ra: Ra, H before the base takes blwf (:873-899)
            for (int i = start; i + 1 < baseAt; i++)
                if (b[i].Category == Ra && b[i + 1].Category == H && (i + 2 == baseAt || b[i + 2].Category != Zwj))
                {
                    Set(i, g => g with { Mask = g.Mask | blwf });
                    Set(i + 1, g => g with { Mask = g.Mask | blwf });
                }
        }
        if (preBaseForms && baseAt + 2 < end)
        {
            // a Halant, Ra the font forms a pre-base form of (:901-916)
            for (int i = baseAt + 1; i + 1 < end; i++)
                if (layout.WouldSubstitute(tag, "pref", b[i].Id, b[i + 1].Id))
                {
                    Set(i, g => g with { Mask = g.Mask | pref });
                    Set(i + 1, g => g with { Mask = g.Mask | pref });
                    break;
                }
        }
        // a ZWNJ stops the half form before it (:918-937)
        for (int i = start + 1; i < end; i++)
            if (IsJoiner(b[i]) && b[i].Category == Zwnj)
            {
                int j = i;
                do
                {
                    j--;
                    Set(j, g => g with { Mask = g.Mask & ~half });
                } while (j > start && !IsConsonant(b[j]));
            }
    }

    // ── final reordering (hb-ot-shaper-indic.cc:1012-1455) ─────────────────────────

    private static void FinalReordering(Script script, bool preBaseForms, List<Glyph> b, int start, int end)
    {
        bool tamilOrMalayalam = script.OldTag is "taml" or "mlym";
        uint pref = WordFontLayoutTables.FeatureBit("pref");
        bool tryPref = preBaseForms;
        // the base again (:1052-1110)
        int baseAt;
        for (baseAt = start; baseAt < end; baseAt++)
            if (b[baseAt].Position >= BaseC)
            {
                if (tryPref && baseAt + 1 < end)
                {
                    // a pre-base-reordering Ra the font did not form a pre-base glyph of: the consonant after it is the base
                    for (int i = baseAt + 1; i < end; i++)
                        if ((b[i].Mask & pref) != 0)
                        {
                            if (!(b[i].Ligated && !b[i].Multiplied))
                            {
                                baseAt = i;
                                while (baseAt < end && IsHalant(b[baseAt])) baseAt++;
                                if (baseAt < end) b[baseAt] = b[baseAt] with { Position = BaseC };
                                tryPref = false;
                            }
                            break;
                        }
                    if (baseAt == end) break;
                }
                if (script.OldTag == "mlym")
                {
                    for (int i = baseAt + 1; i < end; i++)
                    {
                        while (i < end && IsJoiner(b[i])) i++;
                        if (i == end || !IsHalant(b[i])) break;
                        i++;
                        while (i < end && IsJoiner(b[i])) i++;
                        if (i < end && IsConsonant(b[i]) && b[i].Position == BelowC)
                        {
                            baseAt = i;
                            b[baseAt] = b[baseAt] with { Position = BaseC };
                        }
                    }
                }
                if (start < baseAt && b[baseAt].Position > BaseC) baseAt--;
                break;
            }
        if (baseAt == end && start < baseAt && Is(b[baseAt - 1], Zwj)) baseAt--;
        if (baseAt < end)
            while (start < baseAt && Is(b[baseAt], N, H)) baseAt--;

        // pre-base matras move to just before the base, after the last halant of the half forms (:1113-1222)
        if (start + 1 < end && start < baseAt)
        {
            int newPosition = baseAt == end ? baseAt - 2 : baseAt - 1;
            if (!tamilOrMalayalam)
            {
                while (true)
                {
                    while (newPosition > start && !Is(b[newPosition], M, MPst, H)) newPosition--;
                    if (IsHalant(b[newPosition]) && b[newPosition].Position != PreM)
                    {
                        if (newPosition + 1 < end && b[newPosition + 1].Category == Zwj && newPosition > start)
                        {
                            newPosition--;
                            continue;
                        }
                    }
                    else newPosition = start;
                    break;
                }
            }
            if (start < newPosition && b[newPosition].Position != PreM)
                for (int i = newPosition; i > start; i--)
                    if (b[i - 1].Position == PreM)
                    {
                        int oldPosition = i - 1;
                        if (oldPosition < baseAt && baseAt <= newPosition) baseAt--;
                        Glyph matra = b[oldPosition];
                        b.RemoveAt(oldPosition);
                        b.Insert(newPosition, matra);
                        newPosition--;
                    }
        }

        // the reph moves to its script's position (:1224-1385)
        if (start + 1 < end && b[start].Position == RaToBecomeReph && (b[start].Category == Repha ^ (b[start].Ligated && !b[start].Multiplied)))
        {
            int RephTarget()
            {
                int target;
                // after the first explicit halant before the base
                int AfterHalant()
                {
                    int at = start + 1;
                    while (at < baseAt && !IsHalant(b[at])) at++;
                    if (at < baseAt && IsHalant(b[at]))
                    {
                        if (at + 1 < baseAt && IsJoiner(b[at + 1])) at++;
                        return at;
                    }
                    return -1;
                }
                if (script.RephPosition != AfterPost)
                {
                    if ((target = AfterHalant()) >= 0) return target;
                    if (script.RephPosition == AfterMain)
                    {
                        target = baseAt;
                        while (target + 1 < end && b[target + 1].Position <= AfterMain) target++;
                        if (target < end) return target;
                    }
                    if (script.RephPosition == AfterSub)
                    {
                        target = baseAt;
                        while (target + 1 < end && b[target + 1].Position is not (PostC or AfterPost or Smvd)) target++;
                        if (target < end) return target;
                    }
                }
                if ((target = AfterHalant()) >= 0) return target;
                target = end - 1;
                while (target > start && b[target].Position == Smvd) target--;
                if (IsHalant(b[target]))
                    for (int i = baseAt + 1; i < target; i++)
                        if (b[i].Category is M or MPst) target--;
                return target;
            }
            int to = RephTarget();
            Glyph reph = b[start];
            b.RemoveAt(start);
            b.Insert(to, reph);
            if (start < baseAt && baseAt <= to) baseAt--;
        }

        // a formed pre-base-reordering Ra moves before the base, after the last halant (:1386-1440)
        if (tryPref && baseAt + 1 < end)
            for (int i = baseAt + 1; i < end; i++)
                if ((b[i].Mask & pref) != 0)
                {
                    if (b[i].Ligated && !b[i].Multiplied)
                    {
                        int newPosition = baseAt;
                        if (!tamilOrMalayalam)
                            while (newPosition > start && !Is(b[newPosition - 1], M, MPst, H)) newPosition--;
                        if (newPosition > start && IsHalant(b[newPosition - 1]) && newPosition < end && IsJoiner(b[newPosition])) newPosition++;
                        Glyph moved = b[i];
                        b.RemoveAt(i);
                        b.Insert(newPosition, moved);
                        if (newPosition <= baseAt && baseAt < i) baseAt++;
                    }
                    break;
                }

        // init on a left matra at the start of a word (:1444-1452)
        if (b[start].Position == PreM && (start == 0 || !IsLetterOrMark(b[start - 1].CodePoint)))
            b[start] = b[start] with { Mask = b[start].Mask | WordFontLayoutTables.FeatureBit("init") };
    }

    /// <summary>General categories Cf .. Mn in HarfBuzz's order (format, unassigned, private use, surrogate, letters, marks): a character after which a left matra is not word-initial.</summary>
    private static bool IsLetterOrMark(int codePoint)
    {
        if (codePoint > 0xFFFF) return true;
        return CharUnicodeInfo.GetUnicodeCategory((char)codePoint) is UnicodeCategory.Format or UnicodeCategory.OtherNotAssigned or UnicodeCategory.PrivateUse
            or UnicodeCategory.Surrogate or UnicodeCategory.LowercaseLetter or UnicodeCategory.ModifierLetter or UnicodeCategory.OtherLetter
            or UnicodeCategory.TitlecaseLetter or UnicodeCategory.UppercaseLetter or UnicodeCategory.SpacingCombiningMark
            or UnicodeCategory.EnclosingMark or UnicodeCategory.NonSpacingMark;
    }

    /// <summary>
    /// Syllabic category and position of U+0900..U+0D7F, two letters per character ('a' + value): HarfBuzz
    /// hb-ot-shaper-indic-table.cc (generated by gen-indic-table.py from Unicode 16 IndicSyllabicCategory.txt /
    /// IndicPositionalCategory.txt), categories as hb-ot-shaper-indic-machine.rl numbers them, positions as hb-ot-shaper-indic.hh.
    /// </summary>
    private const string Table =
        "inininincececececececececececececececececebebebebebebebebebebebe" +   // 0900
        "bebebebebebebebebebebebebebebebepebebebebebebebebebehjhjdornhjhc" +   // 0920
        "hjhjhjhjhjhjhjhjhjhjhjhjhjeihchjaojnjnininhjhjhjbebebebebebebebe" +   // 0940
        "cecehjhjaoaokekekekekekekekekekeaoaocecececececebebebebebebebebe" +   // 0960
        "keinininaocecececececececeaoaoceceaoaocecebebebebebebebebebebebe" +   // 0980
        "bebebebebebebebebeaobebebebebebepeaobeaoaoaobebebebeaoaodornhmhc" +   // 09A0
        "hmhjhjhjhjaoaohchcaoaohmhmeibeaoaoaoaoaoaoaoaohmaoaoaoaobebeaobe" +   // 09C0
        "cecehjhjaoaokekekekekekekekekekepebeaoaoaoaoaoaoaoaoaoaokeaoinao" +   // 09E0
        "aoinininaocecececececeaoaoaoaoceceaoaocecebebebebebebebebebebebe" +   // 0A00
        "bebebebebebebebebeaobebebebebebepeaobebeaobebeaobebeaoaodoaohmhc" +   // 0A20
        "nmhmhmaoaoaoaohmhmaoaohmhmeiaoaoaohiaoaoaoaoaoaoaobebebebeaobeao" +   // 0A40
        "aoaoaoaoaoaokekekekekekekekekekeininbebeaoqeaoaoaoaoaoaoaoaoaoao" +   // 0A60
        "aoinininaocececececececececeaocececeaocecebebebebebebebebebebebe" +   // 0A80
        "bebebebebebebebebeaobebebebebebepeaobebeaobebebebebeaoaodornhmhc" +   // 0AA0
        "hmhmhmhmhmhjaohjhjhmaohmhmeiaoaoaoaoaoaoaoaoaoaoaoaoaoaoaoaoaoao" +   // 0AC0
        "cecehmhmaoaokekekekekekekekekekeaoaoaoaoaoaoaoaoaobejndojndododo" +   // 0AE0
        "aoihininaocecececececececeaoaoceceaoaocecebebebebebebebebebebebe" +   // 0B00
        "bebebebebebebebebeaobebebebebebepeaobebeaobebebebebeaoaodornhmhf" +   // 0B20
        "hmhjhjhjhjaoaohchfaoaohmhmeiaoaoaoaoaoaoaodohfhmaoaoaoaobebeaobe" +   // 0B40
        "cecehjhjaoaokekekekekekekekekekeaobeaoaoaoaoaoaoaoaoaoaoaoaoaoao" +   // 0B60
        "aoaoinaoaocecececececeaoaoaocececeaocececebeaoaoaobebeaobeaobebe" +   // 0B80
        "aoaoaobebeaoaoaobebebeaoaoaobebepebebebebebebebebebeaoaoaoaohmhm" +   // 0BA0
        "hjhmhmaoaoaohchchcaohmhmhmegaoaoaoaoaoaoaoaoaohmaoaoaoaoaoaoaoao" +   // 0BC0
        "aoaoaoaoaoaokekekekekekekekekekeaoaoaoaoaoaoaoaoaoaoaoaoaoaoaoao" +   // 0BE0
        "ininininincecececececececeaocececeaocececebebebebebebebebebebebe" +   // 0C00
        "bebebebebebebebebeaobebebebebebepebebebebebebebebebeaoaodornhhhh" +   // 0C20
        "hhhhhhhjhjaohhhhhhaohhhhhhegaoaoaoaoaoaoaohhhhaobebebeaoaobeaoao" +   // 0C40
        "cecehhhhaoaokekekekekekekekekekeaoaoaoaoaoaoaoaoaoaoaoaoaoaoaoao" +   // 0C60
        "keinininaocecececececececeaocececeaocececebebebebebebebebebebebe" +   // 0C80
        "bebebebebebebebebeaobebebebebebepebebebeaobebebebebeaoaodornhhhh" +   // 0CA0
        "hhhhhhhjhjaohhhjhjaohjhjhhegaoaoaoaoaoaoaohjhjaoaoaoaoaoaobebeao" +   // 0CC0
        "cecehhhhaoaokekekekekekekekekekeaoseseinaoaoaoaoaoaoaoaoaoaoaoao" +   // 0CE0
        "ininininkececececececececeaocececeaocececebebebebebebebebebebebe" +   // 0D00
        "bebebebebebebebebebebebebebebebepebebebebebebebebebebehjhjrnhmhm" +   // 0D20
        "hmhmhmhmhmaohchchcaohmhmhmegooaoaoaoaoaobebebehmaoaoaoaoaoaoaoce" +   // 0D40
        "cecehmhmaoaokekekekekekekekekekeaoaoaoaoaoaoaoaoaoaobebebebebebe" +   // 0D60
        "";
}
