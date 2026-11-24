using System.Collections.Concurrent;
using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// One font face's horizontal metrics for measuring text as Word lays it out: advance widths (hmtx through the Windows
/// Unicode cmap, format 4 / 12), pair kerning (the kern table's horizontal format 0 subtables; a face without one kerns through
/// its GPOS 'kern' pair adjustments, formats 1 and 2), complex-script shaping (GSUB through <see cref="WordFontLayoutTables"/>:
/// Arabic positional forms and ligatures, Hebrew composition) and the line height Word gives a line of the face (hhea
/// ascender - descender + line gap, em). Read once per family and style (regular, bold, italic, bold italic: each its own font
/// file) from the installed font file (<see cref="WordFontFile"/>; only the head / hhea / OS/2 / hmtx / cmap / kern / GSUB /
/// GPOS / GDEF tables, nothing written) and kept for the process; families load independently (no lock is held while one is
/// read). Without a font file (not Windows, WASM, a family not installed) the families Office themes use fall back to embedded
/// tables (<see cref="DocxHtml.DocxChartFontWidths"/>: Calibri / Cambria / Times New Roman / Arial printable ASCII, Arial's
/// Hebrew block, Calibri's kern pairs); other characters there take the face's average width. Immutable once built.
/// </summary>
internal sealed class WordFontMetrics
{
    private static readonly ConcurrentDictionary<string, Lazy<WordFontMetrics?>> ByFamily = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Test switches: turn one measuring step off to prove a fixture depends on it (the SmartArt harness only).</summary>
    internal static bool DisableStyles, DisableGposKerning, DisableShaping, DisableIndicShaping;

    public readonly string Family;
    /// <summary>hhea ascender - descender + line gap in em: the height of one single-spaced line of DrawingML text in the face.</summary>
    public readonly double LineHeightEm;
    private readonly double _averageEm;
    private readonly double[] _advancesEm;          // per glyph
    private readonly ushort[]?[] _glyphPages;       // code point >> 8 -> 256 glyph ids (0 = not in the font)
    private readonly Dictionary<uint, double>? _kerningEm;   // left glyph << 16 | right glyph -> adjustment
    private readonly WordFontLayoutTables? _layout;
    private readonly double _unitsPerEm;

    private WordFontMetrics(string family, double lineHeightEm, double averageEm, double[] advancesEm, ushort[]?[] glyphPages, Dictionary<uint, double>? kerningEm,
                            WordFontLayoutTables? layout = null, double unitsPerEm = 2048)
    {
        Family = family; LineHeightEm = lineHeightEm; _averageEm = averageEm; _advancesEm = advancesEm; _glyphPages = glyphPages; _kerningEm = kerningEm;
        _layout = layout; _unitsPerEm = unitsPerEm;
    }

    /// <summary>Whether the face has a glyph for the code point.</summary>
    public bool Has(int codePoint) => Glyph(codePoint) != 0;

    /// <summary>The advance width in em (the face's average width for a character it has no glyph for).</summary>
    public double Advance(int codePoint)
    {
        int glyph = Glyph(codePoint);
        return glyph == 0 ? _averageEm : _advancesEm[Math.Min(glyph, _advancesEm.Length - 1)];
    }

    /// <summary>The kern table's pair kerning between two characters in em (0 for an unkerned pair or a face without a kern table).</summary>
    public double Kerning(int left, int right)
    {
        if (_kerningEm is null) return 0;
        int a = Glyph(left), b = Glyph(right);
        return a != 0 && b != 0 && _kerningEm.TryGetValue((uint)a << 16 | (uint)b, out double value) ? value : 0;
    }

    private int Glyph(int codePoint)
    {
        int page = codePoint >> 8;
        return page >= 0 && page < _glyphPages.Length && _glyphPages[page] is { } glyphs ? glyphs[codePoint & 0xFF] : 0;
    }

    /// <summary>
    /// The width of a run of characters drawn in this face, in em, and its pair kerning apart (added only to kerned text).
    /// Complex scripts (<paramref name="complex"/>; <paramref name="openTypeScript"/> "arab", "syrc", "mong", "hebr", "thaa",
    /// "thai", or an Indic script "deva" ... "mlym") go through the font's GSUB as a shaper applies it by default: composition
    /// (ccmp, locl), Arabic / Syriac / Mongolian positional forms from cursive joining (isol / fina / medi / init), then required
    /// and standard ligatures and contextual alternates (rlig, calt, liga, clig, mset, rclt; reverse chaining runs from the end);
    /// Indic text through <see cref="WordIndicShaper"/> (reordering, masked features, the font's advances as they are: the
    /// Indic shaper zeroes no marks); other text is measured glyph by glyph (Word draws Latin text without ligatures by
    /// default). Marks add no width (GDEF class 3, else a non-spacing mark character); joiners (ZWJ / ZWNJ) none. Kerning: the
    /// kern table's pairs, else GPOS 'kern' pair adjustments of the script.
    /// </summary>
    public double ShapedWidth(IReadOnlyList<int> codePoints, string openTypeScript, bool complex, out double kerningEm)
    {
        kerningEm = 0;
        bool indic = complex && !DisableIndicShaping && WordIndicShaper.Handles(openTypeScript);
        if (indic) codePoints = WordIndicShaper.Decompose(codePoints, Has);
        var buffer = new List<WordFontLayoutTables.Glyph>(codePoints.Count);
        string?[]? forms = complex && openTypeScript is "arab" or "syrc" or "mong" ? WordArabicJoining.Forms(codePoints) : null;
        for (int i = 0; i < codePoints.Count; i++) buffer.Add(new WordFontLayoutTables.Glyph(Glyph(codePoints[i]), forms?[i], codePoints[i]));
        bool shaped = false;
        if (complex && !DisableShaping && _layout is { HasSubstitutions: true } && buffer.TrueForAll(g => g.Id != 0 || IsJoiner(g.CodePoint)))
        {
            var unshaped = new List<WordFontLayoutTables.Glyph>(buffer);
            try
            {
                if (indic) WordIndicShaper.Shape(_layout, buffer, openTypeScript, Glyph);
                else _layout.Substitute(buffer, openTypeScript, ShapingStages);
                shaped = true;
            }
            catch (Exception exception) when (WordFontFile.IsUnreadable(exception)) { buffer = unshaped; }
        }
        bool classes = _layout is not null && _layout.HasGlyphClasses && !(indic && shaped);
        double width = 0;
        var spacing = new List<int>(buffer.Count);
        for (int i = 0; i < buffer.Count; i++)
        {
            int glyph = buffer[i].Id;
            if (IsJoiner(buffer[i].CodePoint)) continue;
            bool mark = classes ? _layout!.GlyphClass(glyph) == 3 : !shaped && IsNonSpacingMark(buffer[i].CodePoint);
            if (mark) continue;
            width += glyph == 0 ? _averageEm : _advancesEm[Math.Min(glyph, _advancesEm.Length - 1)];
            if (glyph != 0) spacing.Add(glyph);
        }
        if (indic && shaped && !DisableGposKerning)
        {
            try { width += _layout!.Adjustments(spacing, WordIndicShaper.PositioningTag(_layout, openTypeScript), ["dist"]) / _unitsPerEm; }
            catch (Exception exception) when (WordFontFile.IsUnreadable(exception)) { }
        }
        if (spacing.Count > 1)
        {
            if (_kerningEm is not null)
            {
                for (int i = 1; i < spacing.Count; i++)
                    if (_kerningEm.TryGetValue((uint)spacing[i - 1] << 16 | (uint)spacing[i], out double value)) kerningEm += value;
            }
            else if (_layout is not null && !DisableGposKerning)
            {
                try { kerningEm = _layout.PairKerning(spacing, indic ? WordIndicShaper.PositioningTag(_layout, openTypeScript) : openTypeScript) / _unitsPerEm; }
                catch (Exception exception) when (WordFontFile.IsUnreadable(exception)) { kerningEm = 0; }
            }
        }
        return width;
    }

    /// <summary>The GSUB feature stages a shaper applies to complex-script text by default (HarfBuzz hb-ot-shaper-arabic.cc / hb-ot-shaper-hebrew.cc: composition, positional forms, ligatures).</summary>
    private static readonly string[][] ShapingStages = [["ccmp", "locl"], ["isol", "fina", "medi", "init"], ["rlig", "calt", "liga", "clig", "mset", "rclt"]];

    /// <summary>ZERO WIDTH NON-JOINER / JOINER: default ignorables, drawn with no width (HarfBuzz hides them).</summary>
    private static bool IsJoiner(int codePoint) => codePoint is 0x200C or 0x200D;

    /// <summary>The metrics of a family's regular face: its installed font file, else an embedded table for Calibri / Cambria / Times New Roman / Arial; null otherwise.</summary>
    public static WordFontMetrics? Of(string? family) => Of(family, false, false);

    /// <summary>The metrics of a family in a style: its bold / italic font file, else the nearest installed style (the embedded tables are regular only).</summary>
    public static WordFontMetrics? Of(string? family, bool bold, bool italic)
    {
        if (string.IsNullOrWhiteSpace(family)) return null;
        if (DisableStyles) bold = italic = false;
        string key = (bold ? "b" : "") + (italic ? "i" : "") + "|" + family;
        return ByFamily.GetOrAdd(key, _ => new Lazy<WordFontMetrics?>(() => Load(family!, bold, italic), LazyThreadSafetyMode.ExecutionAndPublication)).Value;
    }

    private static WordFontMetrics? Load(string family, bool bold, bool italic)
    {
        WordFontMetrics? metrics = null;
        try
        {
            if (WordFontFile.Find(family, bold, italic) is { } file)
            {
                // no bold face: Word (GDI) synthesizes the bold from the regular / italic face
                bool synthesizedBold = bold && WordFontFile.Find(family, false, italic) is { } plain && plain == file;
                using (FileStream stream = WordFontFile.Open(file.Path)) metrics = Read(family, stream, file.Face, synthesizedBold);
            }
        }
        catch (Exception exception) when (WordFontFile.IsUnreadable(exception)) { metrics = null; }
        return metrics ?? Embedded(family);
    }

    /// <summary>
    /// GDI's synthesized bold of a face whose head.flags bit 4 says instructions may change advance widths and that has no hdmx
    /// table narrows every advance by 1 / 1.02 (Word 16 measures with it: gabriola-bold-block-list draws its advances ~1.5 %
    /// narrower than the regular face's). Measured 2026-11-24 with GDI GetTextExtentPoint32W / GetCharWidth32W at 4096 - 16384
    /// px, weight 700 against 400: Gabriola, MV Boli, Javanese Text, Mongolian Baiti, Microsoft Himalaya, Microsoft Yi Baiti,
    /// Segoe UI Historic, Segoe UI Emoji, Cambria Math 0.9804 (= 1 / 1.02); faces with an hdmx table (Sylfaen, Impact, Lucida
    /// Sans Unicode, Segoe UI Symbol, Webdings) or without bit 4 (Lucida Console, Ink Free) 1.0003 (the regular advances).
    /// </summary>
    internal const double SynthesizedBoldNarrowing = 1 / 1.02;

    private static WordFontMetrics? Read(string family, Stream stream, int face, bool synthesizedBold = false)
    {
        var tables = WordFontFile.Tables(stream, WordFontFile.FaceOffsets(stream)[face]);
        if (!tables.TryGetValue("head", out var headAt) || !tables.TryGetValue("hhea", out var hheaAt) || !tables.TryGetValue("hmtx", out var hmtxAt) || !tables.TryGetValue("cmap", out var cmapAt)) return null;
        byte[] head = WordFontFile.Read(stream, headAt), hhea = WordFontFile.Read(stream, hheaAt), hmtx = WordFontFile.Read(stream, hmtxAt), cmap = WordFontFile.Read(stream, cmapAt);
        double unitsPerEm = WordFontFile.U16(head, 18);
        if (unitsPerEm <= 0) return null;
        int metricCount = WordFontFile.U16(hhea, 34);
        if (metricCount == 0) return null;
        var advances = new double[Math.Min(metricCount, hmtx.Length / 4)];
        for (int i = 0; i < advances.Length; i++) advances[i] = WordFontFile.U16(hmtx, i * 4) / unitsPerEm;
        if (synthesizedBold && (WordFontFile.U16(head, 16) & 0x10) != 0 && !tables.ContainsKey("hdmx"))
            for (int i = 0; i < advances.Length; i++) advances[i] *= SynthesizedBoldNarrowing;
        if (advances.Length == 0) return null;
        // line height: hhea ascender - descender + line gap (what Word gives a line of DrawingML text: Calibri 1.2207, Arial 1.1499)
        double lineHeight = (WordFontFile.S16(hhea, 4) - WordFontFile.S16(hhea, 6) + WordFontFile.S16(hhea, 8)) / unitsPerEm;
        double average = tables.TryGetValue("OS/2", out var os2At) && WordFontFile.Read(stream, os2At) is { Length: >= 4 } os2 ? WordFontFile.S16(os2, 2) / unitsPerEm : 0;
        ushort[]?[] pages = GlyphPages(cmap);
        Dictionary<uint, double>? kerning = tables.TryGetValue("kern", out var kernAt) ? Kerning(WordFontFile.Read(stream, kernAt), unitsPerEm) : null;
        byte[]? Optional(string tag) => tables.TryGetValue(tag, out var at) ? WordFontFile.Read(stream, at) : null;
        byte[]? gsub = Optional("GSUB"), gpos = Optional("GPOS"), gdef = Optional("GDEF");
        WordFontLayoutTables? layout = gsub is null && gpos is null && gdef is null ? null : new WordFontLayoutTables(gsub, gpos, gdef);
        return new WordFontMetrics(family, lineHeight, average > 0 ? average : advances.Average(), advances, pages, kerning, layout, unitsPerEm);
    }


    /// <summary>The Windows Unicode cmap (3,10 format 12, else 3,1 format 4, else 0,* format 4 / 12; a symbol font's 3,0 at U+F000 and its ASCII alias) as glyph pages.</summary>
    private static ushort[]?[] GlyphPages(byte[] cmap)
    {
        var pages = new ushort[]?[0x1100];
        int count = WordFontFile.U16(cmap, 2);
        int best = -1, bestRank = 0;
        bool symbol = false;
        for (int i = 0; i < count; i++)
        {
            int platform = WordFontFile.U16(cmap, 4 + i * 8), encoding = WordFontFile.U16(cmap, 6 + i * 8);
            long offset = WordFontFile.U32(cmap, 8 + i * 8);
            int rank = platform == 3 && encoding == 10 ? 4 : platform == 3 && encoding == 1 ? 3 : platform == 0 ? 2 : platform == 3 && encoding == 0 ? 1 : 0;
            if (rank > bestRank && offset < cmap.Length) { best = (int)offset; bestRank = rank; symbol = rank == 1; }
        }
        if (best < 0) return pages;
        void Map(long codePoint, int glyph)
        {
            if (glyph == 0 || codePoint < 0 || codePoint > 0x10FFFF) return;
            (pages[codePoint >> 8] ??= new ushort[256])[codePoint & 0xFF] = (ushort)glyph;
            // a symbol font's characters sit at U+F000 + byte: Word measures the byte's character with them
            if (symbol && codePoint is >= 0xF000 and <= 0xF0FF) (pages[0] ??= new ushort[256])[codePoint & 0xFF] = (ushort)glyph;
        }
        int format = WordFontFile.U16(cmap, best);
        if (format == 4)
        {
            int segments = WordFontFile.U16(cmap, best + 6) / 2;
            int ends = best + 14, starts = ends + segments * 2 + 2, deltas = starts + segments * 2, ranges = deltas + segments * 2;
            for (int s = 0; s < segments; s++)
            {
                int end = WordFontFile.U16(cmap, ends + s * 2), start = WordFontFile.U16(cmap, starts + s * 2);
                int delta = WordFontFile.S16(cmap, deltas + s * 2), range = WordFontFile.U16(cmap, ranges + s * 2);
                for (int c = start; c <= end && c != 0xFFFF; c++)
                {
                    int glyph;
                    if (range == 0) glyph = (c + delta) & 0xFFFF;
                    else
                    {
                        int at = ranges + s * 2 + range + (c - start) * 2;
                        if (at + 1 >= cmap.Length) break;
                        glyph = WordFontFile.U16(cmap, at);
                        if (glyph != 0) glyph = (glyph + delta) & 0xFFFF;
                    }
                    Map(c, glyph);
                }
            }
        }
        else if (format == 12)
        {
            long groups = WordFontFile.U32(cmap, best + 12), mapped = 0;
            for (long g = 0; g < groups && best + 16 + g * 12 + 12 <= cmap.Length; g++)
            {
                int at = (int)(best + 16 + g * 12);
                long start = WordFontFile.U32(cmap, at), end = Math.Min(WordFontFile.U32(cmap, at + 4), 0x10FFFFu), glyph = WordFontFile.U32(cmap, at + 8);
                if (start > end || mapped + (end - start + 1) > 0x110000) break;   // a malformed table can't make the walk longer than every code point once
                mapped += end - start + 1;
                for (long c = start; c <= end; c++) Map(c, (int)Math.Min(glyph + c - start, 0xFFFF));
            }
        }
        return pages;
    }

    /// <summary>The kern table's horizontal, non-minimum, non-cross-stream format 0 pairs (version 0 header; OpenType "kern").</summary>
    private static Dictionary<uint, double>? Kerning(byte[] kern, double unitsPerEm)
    {
        if (kern.Length < 4 || WordFontFile.U16(kern, 0) != 0) return null;
        var pairs = new Dictionary<uint, double>();
        int tables = WordFontFile.U16(kern, 2), at = 4;
        for (int t = 0; t < tables && at + 6 <= kern.Length; t++)
        {
            int length = WordFontFile.U16(kern, at + 2), coverage = WordFontFile.U16(kern, at + 4);
            if (coverage >> 8 == 0 && (coverage & 0x7) == 1)
            {
                int count = WordFontFile.U16(kern, at + 6);
                for (int i = 0; i < count && at + 14 + i * 6 + 6 <= kern.Length; i++)
                {
                    int p = at + 14 + i * 6;
                    pairs[(uint)WordFontFile.U16(kern, p) << 16 | (uint)WordFontFile.U16(kern, p + 2)] = WordFontFile.S16(kern, p + 4) / unitsPerEm;
                }
                // the 16-bit length overflows for a big format 0 subtable: its pair count gives the true size
                length = 14 + count * 6;
            }
            if (length <= 0) break;
            at += length;
        }
        return pairs.Count > 0 ? pairs : null;
    }

    /// <summary>An embedded table for the families Office themes use, when the font file can't be read: printable ASCII (+ Arial's Hebrew block, Calibri's kern pairs).</summary>
    private static WordFontMetrics? Embedded(string family)
    {
        string name = family.ToLowerInvariant() switch
        {
            "calibri" or "carlito" => "Calibri",
            "cambria" or "caladea" => "Cambria",
            "times new roman" or "liberation serif" or "tinos" => "Times New Roman",
            "arial" or "liberation sans" or "arimo" or "helvetica" => "Arial",
            _ => "",
        };
        if (name.Length == 0) return null;
        var (_, average, _, widths) = DocxHtml.DocxChartFontWidths.Fonts.First(f => f.Name == name);
        double lineHeight = name switch { "Calibri" => 1.2207, "Cambria" => 1.1724, "Times New Roman" => 1.1499, _ => 1.1499 };   // hhea ascender - descender + line gap / 2048 of the Windows files
        var advances = new List<double> { average };
        var pages = new ushort[]?[0x1100];
        void Add(int codePoint, double width) { (pages[codePoint >> 8] ??= new ushort[256])[codePoint & 0xFF] = (ushort)advances.Count; advances.Add(width); }
        for (int c = ' '; c <= '~'; c++) Add(c, widths[c - ' ']);
        if (name == "Arial")
            for (int c = 0x0590; c <= 0x05FF; c++) if (DocxHtml.DocxChartFontWidths.Hebrew[c - 0x0590] is > 0 and var width) Add(c, width);
        Dictionary<uint, double>? kerning = null;
        if (name == "Calibri")
        {
            kerning = [];
            foreach ((char left, char right, double em) in DocxHtml.DocxCalibriKerning.All())
                kerning[(uint)pages[0]![left] << 16 | pages[0]![right]] = em;
        }
        return new WordFontMetrics(name, lineHeight, average, [.. advances], pages, kerning);
    }

    /// <summary>Unicode general category Mn / Me (niqqud, cantillation, accents): no advance of its own once shaped.</summary>
    public static bool IsNonSpacingMark(int codePoint)
    {
        if (codePoint > 0xFFFF) return false;
        UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory((char)codePoint);
        return category is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark;
    }
}
