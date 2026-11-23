using System.Collections.Concurrent;

namespace FastDocView.Html;

/// <summary>
/// The OpenType layout tables of one face, read only, for measuring shaped text (OpenType 1.9 "GSUB", "GPOS", "GDEF",
/// "OpenType Layout Common Table Formats"): glyph substitution by the features a shaper applies to a script (single, multiple,
/// ligature, context and chained context substitution, through extension lookups; reverse chaining and alternates are not
/// width-changing defaults and are skipped), lookup flags that skip base / ligature / mark glyphs (GDEF glyph classes, mark
/// attachment classes, mark glyph sets) and GPOS pair adjustment (formats 1 and 2, through extension lookups) of the 'kern'
/// feature: the x advance of both value records. Works on the raw table bytes; the lookups of a script and feature set are
/// collected once and kept. A malformed table throws (the caller measures the text unshaped).
/// </summary>
internal sealed class WordFontLayoutTables(byte[]? gsub, byte[]? gpos, byte[]? gdef)
{
    /// <summary>One glyph of the buffer with the positional form feature it takes ("init", "medi", "fina", "isol"; null = any).</summary>
    internal struct Glyph(int id, string? form)
    {
        public int Id = id;
        public string? Form = form;
    }

    /// <summary>Lookups of a table for one script and feature set, in lookup-list order, with the features each serves.</summary>
    private readonly ConcurrentDictionary<string, List<(int Lookup, string[] Features)>> _lookups = new(StringComparer.Ordinal);

    // HarfBuzz caps nested (context) lookups at 64 levels (hb-ot-layout-gsubgpos.hh HB_MAX_NESTING_LEVEL): a font whose
    // lookups call each other in a cycle would recurse forever; real fonts nest one level
    private const int MaximumNesting = 64;

    public bool HasSubstitutions => gsub is not null;

    /// <summary>Whether GDEF gives glyph classes (base / ligature / mark).</summary>
    public bool HasGlyphClasses => gdef is { Length: >= 6 } && U16(gdef, 4) > 0;

    /// <summary>Whether the face kerns through GPOS for the script (a 'kern' feature with lookups).</summary>
    public bool HasPairKerning(string script) => gpos is not null && Lookups(gpos, script, ["kern"]).Count > 0;

    // ── GDEF ──────────────────────────────────────────────────────────────────────

    /// <summary>The GDEF glyph class: 1 base, 2 ligature, 3 mark, 4 component; 0 when the face has no class for it.</summary>
    public int GlyphClass(int glyph) => gdef is not null && U16(gdef, 4) is var at and > 0 ? ClassOf(gdef, at, glyph) : 0;

    private int MarkAttachClass(int glyph) => gdef is not null && U16(gdef, 10) is var at and > 0 ? ClassOf(gdef, at, glyph) : 0;

    private bool InMarkSet(int set, int glyph)
    {
        if (gdef is null || U16(gdef, 2) < 2 || gdef.Length < 14) return false;
        int sets = U16(gdef, 12);
        if (sets == 0 || set >= U16(gdef, sets + 2)) return false;
        return Coverage(gdef, sets + (int)U32(gdef, sets + 4 + set * 4), glyph) >= 0;
    }

    /// <summary>Whether a lookup with these flags skips the glyph (LookupFlag: ignore base 2, ligatures 4, marks 8, mark set 0x10, attachment type 0xFF00).</summary>
    private bool Skips(int glyph, int flag, int markSet)
    {
        if ((flag & 0xFF1E) == 0) return false;
        int glyphClass = GlyphClass(glyph);
        if (glyphClass == 1) return (flag & 2) != 0;
        if (glyphClass == 2) return (flag & 4) != 0;
        if (glyphClass != 3) return false;
        if ((flag & 8) != 0) return true;
        if ((flag & 0x10) != 0) return !InMarkSet(markSet, glyph);
        return flag >> 8 != 0 && MarkAttachClass(glyph) != flag >> 8;
    }

    // ── GSUB ──────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Applies the features of each stage in turn (a stage's lookups in lookup-list order, as shapers do); a positional feature
    /// ("init", "medi", "fina", "isol") applies only to the glyphs whose form names it.
    /// </summary>
    public void Substitute(List<Glyph> buffer, string script, IEnumerable<string[]> stages)
    {
        if (gsub is null) return;
        // a valid font's substitutions grow text by a few glyphs per character at most (decompositions); a cycle of multiple
        // substitutions in a malformed font would grow it without end
        int limit = buffer.Count * 8 + 64;
        foreach (string[] stage in stages)
            foreach ((int lookup, string[] features) in Lookups(gsub, script, stage))
            {
                bool positionalOnly = features.All(IsPositional);
                for (int i = 0; i < buffer.Count; i++)
                {
                    if (positionalOnly && (buffer[i].Form is not { } form || Array.IndexOf(features, form) < 0)) continue;
                    int before = buffer.Count;
                    if (!ApplySubstitution(lookup, buffer, i, 0)) continue;
                    // continue after this lookup's output (a multiple substitution must not run again on the glyphs it made)
                    if (buffer.Count > before) i += buffer.Count - before;
                    if (buffer.Count > limit) throw new InvalidDataException("GSUB output too long");
                }
            }
    }

    private static bool IsPositional(string feature) => feature is "init" or "medi" or "fina" or "isol";

    private (int Type, int Flag, int MarkSet, int[] Subtables) Lookup(byte[] table, int index)
    {
        int list = U16(table, 8);
        if (index >= U16(table, list)) throw new InvalidDataException("lookup index out of range");
        int at = list + U16(table, list + 2 + index * 2);
        int type = U16(table, at), flag = U16(table, at + 2), count = U16(table, at + 4);
        var subtables = new int[count];
        for (int s = 0; s < count; s++) subtables[s] = at + U16(table, at + 6 + s * 2);
        int markSet = (flag & 0x10) != 0 ? U16(table, at + 6 + count * 2) : 0;
        return (type, flag, markSet, subtables);
    }

    /// <summary>Resolves an extension subtable (GSUB 7 / GPOS 9) to its real type and offset.</summary>
    private static (int Type, int At) Resolve(byte[] table, int type, int at, int extensionType)
        => type == extensionType ? (U16(table, at + 2), at + (int)U32(table, at + 4)) : (type, at);

    /// <summary>Applies a GSUB lookup at one buffer position; true when it substituted.</summary>
    private bool ApplySubstitution(int lookupIndex, List<Glyph> buffer, int position, int depth)
    {
        if (depth > MaximumNesting || position >= buffer.Count) return false;
        byte[] table = gsub!;
        (int lookupType, int flag, int markSet, int[] subtables) = Lookup(table, lookupIndex);
        int glyph = buffer[position].Id;
        if (Skips(glyph, flag, markSet)) return false;
        foreach (int subtableAt in subtables)
        {
            (int type, int at) = Resolve(table, lookupType, subtableAt, 7);
            int format = U16(table, at);
            switch (type)
            {
                case 1:
                {
                    int coverage = Coverage(table, at + U16(table, at + 2), glyph);
                    if (coverage < 0) continue;
                    int replacement = format == 1 ? (glyph + S16(table, at + 4)) & 0xFFFF : U16(table, at + 6 + coverage * 2);
                    buffer[position] = new Glyph(replacement, buffer[position].Form);
                    return true;
                }
                case 2:
                {
                    int coverage = Coverage(table, at + U16(table, at + 2), glyph);
                    if (coverage < 0) continue;
                    int sequence = at + U16(table, at + 6 + coverage * 2), count = U16(table, sequence);
                    string? form = buffer[position].Form;
                    buffer.RemoveAt(position);
                    for (int k = count - 1; k >= 0; k--) buffer.Insert(position, new Glyph(U16(table, sequence + 2 + k * 2), form));
                    return true;
                }
                case 4:
                {
                    int coverage = Coverage(table, at + U16(table, at + 2), glyph);
                    if (coverage < 0) continue;
                    int set = at + U16(table, at + 6 + coverage * 2), ligatures = U16(table, set);
                    for (int l = 0; l < ligatures; l++)
                    {
                        int ligature = set + U16(table, set + 2 + l * 2), components = U16(table, ligature + 2);
                        int[]? matched = MatchInput(buffer, position, components, (k, g) => g == U16(table, ligature + 4 + (k - 1) * 2), flag, markSet);
                        if (matched is null) continue;
                        buffer[position] = new Glyph(U16(table, ligature), buffer[position].Form);
                        for (int k = matched.Length - 1; k >= 1; k--) buffer.RemoveAt(matched[k]);
                        return true;
                    }
                    continue;
                }
                case 5 or 6:
                    if (ApplyContext(table, type, format, at, buffer, position, flag, markSet, depth, ApplySubstitution)) return true;
                    continue;
                default:
                    continue;   // 3 alternate (not a default feature), 8 reverse chaining (Nastaliq), unknown
            }
        }
        return false;
    }

    /// <summary>
    /// Context (type 5 / GPOS 7) and chained context (type 6 / GPOS 8) rules, formats 1 (glyphs), 2 (classes) and 3 (coverages):
    /// when the input (and backtrack / lookahead) match at the position, the rule's nested lookups run at their input positions.
    /// </summary>
    private bool ApplyContext(byte[] table, int type, int format, int at, List<Glyph> buffer, int position, int flag, int markSet, int depth,
                              Func<int, List<Glyph>, int, int, bool> apply)
    {
        int glyph = buffer[position].Id;
        bool chained = type == 6;
        if (format == 3)
        {
            if (!chained)
            {
                int count = U16(table, at + 2), records = U16(table, at + 4);
                int Cov(int k) => at + U16(table, at + 6 + k * 2);
                if (Coverage(table, Cov(0), glyph) < 0) return false;
                int[]? matched = MatchInput(buffer, position, count, (k, g) => Coverage(table, Cov(k), g) >= 0, flag, markSet);
                if (matched is null) return false;
                RunRecords(table, at + 6 + count * 2, records, matched, buffer, depth, apply);
                return true;
            }
            int backtrack = U16(table, at + 2);
            int inputAt = at + 4 + backtrack * 2, input = U16(table, inputAt);
            int lookaheadAt = inputAt + 2 + input * 2, lookahead = U16(table, lookaheadAt);
            int recordsAt = lookaheadAt + 2 + lookahead * 2, recordCount = U16(table, recordsAt);
            if (input == 0 || Coverage(table, at + U16(table, inputAt + 2), glyph) < 0) return false;
            int[]? inputMatched = MatchInput(buffer, position, input, (k, g) => Coverage(table, at + U16(table, inputAt + 2 + k * 2), g) >= 0, flag, markSet);
            if (inputMatched is null) return false;
            if (!MatchBacktrack(buffer, position, backtrack, (k, g) => Coverage(table, at + U16(table, at + 4 + k * 2), g) >= 0, flag, markSet)) return false;
            if (!MatchLookahead(buffer, inputMatched[inputMatched.Length - 1], lookahead, (k, g) => Coverage(table, at + U16(table, lookaheadAt + 2 + k * 2), g) >= 0, flag, markSet)) return false;
            RunRecords(table, recordsAt + 2, recordCount, inputMatched, buffer, depth, apply);
            return true;
        }
        int coverage = Coverage(table, at + U16(table, at + 2), glyph);
        if (coverage < 0) return false;
        if (format == 1)
        {
            int setsAt = chained ? at + 4 : at + 4;
            if (coverage >= U16(table, setsAt)) return false;
            int set = at + U16(table, setsAt + 2 + coverage * 2);
            return ApplyRuleSet(table, set, chained, buffer, position, flag, markSet, depth, apply, g => g, g => g, g => g);
        }
        if (format == 2)
        {
            if (!chained)
            {
                int classes = at + U16(table, at + 4);
                int setIndex = ClassOf(table, classes, glyph);
                if (setIndex >= U16(table, at + 6)) return false;
                int offset = U16(table, at + 8 + setIndex * 2);
                if (offset == 0) return false;
                return ApplyRuleSet(table, at + offset, false, buffer, position, flag, markSet, depth, apply, g => ClassOf(table, classes, g), g => 0, g => 0);
            }
            int backtrackClasses = at + U16(table, at + 4), inputClasses = at + U16(table, at + 6), lookaheadClasses = at + U16(table, at + 8);
            int chainSet = ClassOf(table, inputClasses, glyph);
            if (chainSet >= U16(table, at + 10)) return false;
            int chainOffset = U16(table, at + 12 + chainSet * 2);
            if (chainOffset == 0) return false;
            return ApplyRuleSet(table, at + chainOffset, true, buffer, position, flag, markSet, depth, apply,
                                g => ClassOf(table, inputClasses, g), g => ClassOf(table, backtrackClasses, g), g => ClassOf(table, lookaheadClasses, g));
        }
        return false;
    }

    /// <summary>The rules of a format 1 / 2 rule set: values are glyphs (format 1) or classes (format 2, through the class functions).</summary>
    private bool ApplyRuleSet(byte[] table, int set, bool chained, List<Glyph> buffer, int position, int flag, int markSet, int depth,
                              Func<int, List<Glyph>, int, int, bool> apply, Func<int, int> inputValue, Func<int, int> backtrackValue, Func<int, int> lookaheadValue)
    {
        int rules = U16(table, set);
        for (int r = 0; r < rules; r++)
        {
            int rule = set + U16(table, set + 2 + r * 2);
            if (!chained)
            {
                int count = U16(table, rule), records = U16(table, rule + 2);
                int[]? matched = MatchInput(buffer, position, count, (k, g) => inputValue(g) == U16(table, rule + 4 + (k - 1) * 2), flag, markSet);
                if (matched is null) continue;
                RunRecords(table, rule + 4 + (count - 1) * 2, records, matched, buffer, depth, apply);
                return true;
            }
            int backtrack = U16(table, rule);
            int inputAt = rule + 2 + backtrack * 2, input = U16(table, inputAt);
            int lookaheadAt = inputAt + 2 + Math.Max(input - 1, 0) * 2, lookahead = U16(table, lookaheadAt);
            int recordsAt = lookaheadAt + 2 + lookahead * 2, recordCount = U16(table, recordsAt);
            if (input == 0) continue;
            int[]? inputMatched = MatchInput(buffer, position, input, (k, g) => inputValue(g) == U16(table, inputAt + 2 + (k - 1) * 2), flag, markSet);
            if (inputMatched is null) continue;
            if (!MatchBacktrack(buffer, position, backtrack, (k, g) => backtrackValue(g) == U16(table, rule + 2 + k * 2), flag, markSet)) continue;
            if (!MatchLookahead(buffer, inputMatched[inputMatched.Length - 1], lookahead, (k, g) => lookaheadValue(g) == U16(table, lookaheadAt + 2 + k * 2), flag, markSet)) continue;
            RunRecords(table, recordsAt + 2, recordCount, inputMatched, buffer, depth, apply);
            return true;
        }
        return false;
    }

    /// <summary>A rule's (sequence index, lookup index) records, each at its matched input position (later positions follow a change in length).</summary>
    private static void RunRecords(byte[] table, int at, int count, int[] matched, List<Glyph> buffer, int depth, Func<int, List<Glyph>, int, int, bool> apply)
    {
        for (int r = 0; r < count; r++)
        {
            int sequence = U16(table, at + r * 4), lookup = U16(table, at + r * 4 + 2);
            if (sequence >= matched.Length || matched[sequence] >= buffer.Count) continue;
            int before = buffer.Count;
            apply(lookup, buffer, matched[sequence], depth + 1);
            int change = buffer.Count - before;
            if (change != 0) for (int k = sequence + 1; k < matched.Length; k++) matched[k] += change;
        }
    }

    /// <summary>The positions of an input sequence of <paramref name="count"/> glyphs starting at <paramref name="position"/> (skipped glyphs passed over); null when one doesn't match.</summary>
    private int[]? MatchInput(List<Glyph> buffer, int position, int count, Func<int, int, bool> matches, int flag, int markSet)
    {
        var positions = new int[Math.Max(count, 1)];
        positions[0] = position;
        int at = position;
        for (int k = 1; k < count; k++)
        {
            do at++; while (at < buffer.Count && Skips(buffer[at].Id, flag, markSet));
            if (at >= buffer.Count || !matches(k, buffer[at].Id)) return null;
            positions[k] = at;
        }
        return positions;
    }

    private bool MatchBacktrack(List<Glyph> buffer, int position, int count, Func<int, int, bool> matches, int flag, int markSet)
    {
        int at = position;
        for (int k = 0; k < count; k++)
        {
            do at--; while (at >= 0 && Skips(buffer[at].Id, flag, markSet));
            if (at < 0 || !matches(k, buffer[at].Id)) return false;
        }
        return true;
    }

    private bool MatchLookahead(List<Glyph> buffer, int last, int count, Func<int, int, bool> matches, int flag, int markSet)
    {
        int at = last;
        for (int k = 0; k < count; k++)
        {
            do at++; while (at < buffer.Count && Skips(buffer[at].Id, flag, markSet));
            if (at >= buffer.Count || !matches(k, buffer[at].Id)) return false;
        }
        return true;
    }

    // ── GPOS ──────────────────────────────────────────────────────────────────────

    /// <summary>The pair kerning of a glyph sequence in font units: the 'kern' feature's pair adjustment lookups (x advance of both glyphs), each lookup passing over the glyphs its flags skip.</summary>
    public int PairKerning(IReadOnlyList<int> glyphs, string script)
    {
        if (gpos is null || glyphs.Count < 2) return 0;
        byte[] table = gpos;
        int total = 0;
        foreach ((int lookupIndex, _) in Lookups(table, script, ["kern"]))
        {
            (int lookupType, int flag, int markSet, int[] subtables) = Lookup(table, lookupIndex);
            if (lookupType is not (2 or 9)) continue;
            for (int i = 0; i < glyphs.Count; i++)
            {
                if (Skips(glyphs[i], flag, markSet)) continue;
                int j = i + 1;
                while (j < glyphs.Count && Skips(glyphs[j], flag, markSet)) j++;
                if (j >= glyphs.Count) break;
                foreach (int subtableAt in subtables)
                {
                    (int type, int at) = Resolve(table, lookupType, subtableAt, 9);
                    if (type == 2 && PairAdjustment(table, at, glyphs[i], glyphs[j]) is { } adjustment) { total += adjustment; break; }
                }
            }
        }
        return total;
    }

    /// <summary>A PairPos subtable's x advance adjustment (value record 1 + 2) for two glyphs; null when the subtable doesn't cover the pair.</summary>
    private static int? PairAdjustment(byte[] table, int at, int first, int second)
    {
        int format = U16(table, at);
        int coverage = Coverage(table, at + U16(table, at + 2), first);
        if (coverage < 0) return null;
        int format1 = U16(table, at + 4), format2 = U16(table, at + 6);
        int size1 = ValueSize(format1), size2 = ValueSize(format2);
        int XAdvance(int record, int valueFormat) => (valueFormat & 4) != 0 ? S16(table, record + ValueSize(valueFormat & 3)) : 0;
        if (format == 1)
        {
            if (coverage >= U16(table, at + 8)) return null;
            int set = at + U16(table, at + 10 + coverage * 2), count = U16(table, set), recordSize = 2 + size1 + size2;
            int low = 0, high = count - 1;
            while (low <= high)
            {
                int middle = (low + high) / 2, record = set + 2 + middle * recordSize, glyph = U16(table, record);
                if (glyph < second) low = middle + 1;
                else if (glyph > second) high = middle - 1;
                else return XAdvance(record + 2, format1) + XAdvance(record + 2 + size1, format2);
            }
            return null;
        }
        if (format == 2)
        {
            int class1 = ClassOf(table, at + U16(table, at + 8), first), class2 = ClassOf(table, at + U16(table, at + 10), second);
            int class1Count = U16(table, at + 12), class2Count = U16(table, at + 14);
            if (class1 >= class1Count || class2 >= class2Count) return null;
            int record = at + 16 + (class1 * class2Count + class2) * (size1 + size2);
            return XAdvance(record, format1) + XAdvance(record + size1, format2);
        }
        return null;
    }

    /// <summary>A value record's size in bytes: two per set bit of the low byte of its format (device offsets included).</summary>
    private static int ValueSize(int valueFormat)
    {
        int size = 0;
        for (int bit = 0; bit < 8; bit++) if ((valueFormat >> bit & 1) != 0) size += 2;
        return size;
    }

    // ── common ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// The lookups the features of a script use, in lookup-list order: the script's default language system (else DFLT, else
    /// latn), its required feature and each listed feature tag.
    /// </summary>
    private List<(int Lookup, string[] Features)> Lookups(byte[] table, string script, string[] features)
    {
        string key = (ReferenceEquals(table, gsub) ? "S|" : "P|") + script + "|" + string.Join(",", features);
        return _lookups.GetOrAdd(key, _ =>
        {
            var byLookup = new SortedDictionary<int, List<string>>();
            int scripts = U16(table, 4), featureList = U16(table, 6);
            int scriptAt = FindScript(table, scripts, script) ?? FindScript(table, scripts, "DFLT") ?? FindScript(table, scripts, "latn") ?? -1;
            if (scriptAt < 0) return [];
            int langSys = U16(table, scriptAt);
            if (langSys == 0) return [];
            langSys += scriptAt;
            int required = U16(table, langSys + 2), count = U16(table, langSys + 4);
            var featureIndices = new List<int>();
            if (required != 0xFFFF) featureIndices.Add(required);
            for (int f = 0; f < count; f++) featureIndices.Add(U16(table, langSys + 6 + f * 2));
            int featureCount = U16(table, featureList);
            foreach (int index in featureIndices)
            {
                if (index >= featureCount) continue;
                int record = featureList + 2 + index * 6;
                string tag = new string(new[] { (char)table[record], (char)table[record + 1], (char)table[record + 2], (char)table[record + 3] });
                if (Array.IndexOf(features, tag) < 0) continue;
                int feature = featureList + U16(table, record + 4), lookups = U16(table, feature + 2);
                for (int l = 0; l < lookups; l++)
                {
                    int lookup = U16(table, feature + 4 + l * 2);
                    if (!byLookup.TryGetValue(lookup, out List<string>? tags)) byLookup[lookup] = tags = [];
                    if (!tags.Contains(tag)) tags.Add(tag);
                }
            }
            return byLookup.Select(entry => (entry.Key, entry.Value.ToArray())).ToList();
        });
    }

    private static int? FindScript(byte[] table, int scripts, string tag)
    {
        int count = U16(table, scripts);
        for (int s = 0; s < count; s++)
        {
            int record = scripts + 2 + s * 6;
            if (table[record] == tag[0] && table[record + 1] == tag[1] && table[record + 2] == tag[2] && table[record + 3] == tag[3])
                return scripts + U16(table, record + 4);
        }
        return null;
    }

    /// <summary>The coverage index of a glyph (format 1 glyph array, format 2 ranges); -1 when not covered.</summary>
    private static int Coverage(byte[] table, int at, int glyph)
    {
        int format = U16(table, at), count = U16(table, at + 2);
        int low = 0, high = count - 1;
        while (low <= high)
        {
            int middle = (low + high) / 2;
            if (format == 1)
            {
                int value = U16(table, at + 4 + middle * 2);
                if (value < glyph) low = middle + 1;
                else if (value > glyph) high = middle - 1;
                else return middle;
            }
            else if (format == 2)
            {
                int range = at + 4 + middle * 6;
                if (U16(table, range + 2) < glyph) low = middle + 1;
                else if (U16(table, range) > glyph) high = middle - 1;
                else return U16(table, range + 4) + glyph - U16(table, range);
            }
            else return -1;
        }
        return -1;
    }

    /// <summary>The class of a glyph in a class definition (format 1 array, format 2 ranges); 0 when not listed.</summary>
    private static int ClassOf(byte[] table, int at, int glyph)
    {
        int format = U16(table, at);
        if (format == 1)
        {
            int start = U16(table, at + 2), count = U16(table, at + 4);
            return glyph >= start && glyph < start + count ? U16(table, at + 6 + (glyph - start) * 2) : 0;
        }
        if (format != 2) return 0;
        int ranges = U16(table, at + 2), low = 0, high = ranges - 1;
        while (low <= high)
        {
            int middle = (low + high) / 2, range = at + 4 + middle * 6;
            if (U16(table, range + 2) < glyph) low = middle + 1;
            else if (U16(table, range) > glyph) high = middle - 1;
            else return U16(table, range + 4);
        }
        return 0;
    }

    private static int U16(byte[] b, int at) => WordFontFile.U16(b, at);
    private static int S16(byte[] b, int at) => WordFontFile.S16(b, at);
    private static uint U32(byte[] b, int at) => WordFontFile.U32(b, at);
}
