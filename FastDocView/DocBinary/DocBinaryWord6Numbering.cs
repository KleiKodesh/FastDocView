using System.Text;
using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// Word 6.0 / 95 paragraph numbering, counted at index time like the Word 97 lists: a paragraph's sprmPNLvlAnm (1-9 outline,
/// 10 numbering, 11 bullets / sequence, 12 pause; LibreOffice ww8par2.cxx GetNumType) and sprmPAnld (an ANLD), and the
/// section's sprmSOlstAnm (an OLST) for outline levels (ww8par2.cxx NextAnlLine: the OLST when the section has one - "the
/// ANLDs connected to each outline line contain only nonsense" - else the paragraph's ANLD).
/// Counting follows LibreOffice's StartAnl / NextAnlLine / StopAnlToRestart: a run of numbered paragraphs counts from the
/// ANLD's iStartAt; a paragraph without numbering stops both the numbering and the outline (they restart at iStartAt);
/// moving between numbering and outline stops only the numbering. A label is the ANLV's text before + number + text after
/// (ww8par2.cxx SetBaseAnlv / SetAnlvStrings): nfc 0 arabic, 1 / 2 upper / lower roman, 3 / 4 upper / lower letter,
/// 5 / 7 ordinal (arabic with "."), others arabic; fPrev joins the upper outline levels' numbers with "." and then the
/// texts are dropped (SetIncludeUpperLevels, "there are &lt;= 1 number to show"); a symbol-font ANLV (chs 2) is a bullet,
/// its first character (recoded as the Word 97 symbol bullets are, <see cref="WordSymbolFontBullet"/>).
/// ANLV (ww8struc.hxx WW8_ANLV, 16 bytes): nfc, cbTextBefore, cbTextAfter, flags (0x04 fPrev), ..., ftc at 6, iStartAt at
/// 10. ANLD: an ANLV then rgchAnld[32] at 0x14. OLST: nine ANLVs then rgch[64] at 0x94, each level's texts after the
/// earlier levels' (SetNumOlst nTextOfs).
/// </summary>
internal static class DocBinaryWord6Numbering
{
    public const int AnldLength = 0x14 + 32;
    public const int OlstLength = 0x94 + 64;
    private const int FlagsListId = -2, OutlineListId = -1;      // the counter state's keys (DocBinaryListCounter)

    /// <summary>Advances the counters for one paragraph and returns its label; null when it shows none.</summary>
    public static string? Next(Dictionary<int, int[]> counters, int levelType, byte[]? anld, byte[]? olst, DocBinaryFontTable fonts)
    {
        if (!counters.TryGetValue(FlagsListId, out int[]? flags)) counters[FlagsListId] = flags = new int[DocBinaryListDefinitions.MaxLevels];
        if (!counters.TryGetValue(OutlineListId, out int[]? outline)) counters[OutlineListId] = outline = new int[DocBinaryListDefinitions.MaxLevels];
        const int SequenceActive = 0, SequenceCount = 1, OutlineActive = 2;
        if (levelType == 12) return null;                            // pause: no number, nothing stops
        bool hasAnld = anld is { Length: >= AnldLength };
        bool hasOlst = olst is { Length: >= OlstLength };
        if (levelType is 10 or 11)
        {
            if (!hasAnld) return null;
            if (flags[SequenceActive] == 0) { flags[SequenceActive] = 1; flags[SequenceCount] = StartOf(anld!, 0) - 1; }
            flags[SequenceCount]++;
            // the numbering is the first level of outline levels nested under it: they restart below each number
            outline[0] = flags[SequenceCount];
            for (int deeper = 1; deeper < outline.Length; deeper++) outline[deeper] = (hasOlst ? StartOf(olst!, deeper * 16) : 1) - 1;
            return LabelOf(anld!, 0, anld!, 0x14, 32, flags[SequenceCount], fonts);
        }
        if (levelType is >= 1 and <= 9)
        {
            int level = levelType - 1;
            // tdf150197_anlv2ListFormat.doc (level 10 paragraphs, then level 2 ones) must read 1. 2. 2.1 2.2 2.3 3. 4. 4.1:
            // a deeper outline level nests under a running numbering. LibreOffice halts the numbering instead and shows 4.1
            // (its test says "Did you fix me? I should be 2.1"); a first-level outline paragraph still halts it.
            bool nested = level > 0 && flags[SequenceActive] != 0;
            if (!nested) flags[SequenceActive] = 0;
            if (flags[OutlineActive] == 0)
            {
                flags[OutlineActive] = 1;
                for (int i = nested ? 1 : 0; i < outline.Length; i++) outline[i] = (hasOlst ? StartOf(olst!, i * 16) : 1) - 1;
                if (!hasOlst && hasAnld) outline[level] = StartOf(anld!, 0) - 1;
            }
            outline[level]++;
            for (int deeper = level + 1; deeper < outline.Length; deeper++) outline[deeper] = (hasOlst ? StartOf(olst!, deeper * 16) : 1) - 1;
            if (hasOlst)
            {
                int textOffset = 0;
                for (int i = 0; i < level; i++) textOffset += olst![i * 16 + 1] + olst[i * 16 + 2];
                if ((olst![level * 16 + 3] & 0x04) != 0)            // fPrev: the upper levels' numbers too, no texts
                {
                    var joined = new StringBuilder();
                    for (int i = 0; i <= level; i++)
                    {
                        if (i > 0) joined.Append('.');
                        joined.Append(NumberText(olst[i * 16], outline[i]));
                    }
                    return joined.ToString();
                }
                return LabelOf(olst, level * 16, olst, 0x94 + textOffset, Math.Max(0, 64 - textOffset), outline[level], fonts);
            }
            return hasAnld ? LabelOf(anld!, 0, anld!, 0x14, 32, outline[level], fonts) : null;
        }
        flags[SequenceActive] = 0;                                   // no numbering: both stop (StopAllAnl)
        flags[OutlineActive] = 0;
        return null;
    }

    private static int StartOf(byte[] structure, int anlvAt) => Math.Max(0, (int)BitConverter.ToUInt16(structure, anlvAt + 10));

    private static string LabelOf(byte[] anlvSource, int anlvAt, byte[] textSource, int textAt, int textRoom, int number, DocBinaryFontTable fonts)
    {
        int before = anlvSource[anlvAt + 1], after = anlvSource[anlvAt + 2];
        if (before + after > textRoom) before = after = 0;          // out of range: no texts (SetAnlvStrings ignores them)
        int fontIndex = BitConverter.ToUInt16(anlvSource, anlvAt + 6);
        bool isSymbolFont = fontIndex < fonts.CharacterSets.Count && fonts.CharacterSets[fontIndex] == 2;
        string Text(int from, int count)
        {
            var text = new StringBuilder(count);
            for (int i = 0; i < count; i++) text.Append(DocBinaryTextDecoder.DecodeCompressed(textSource[textAt + from + i]));
            return text.ToString();
        }
        if (isSymbolFont)
        {
            // a bullet: the first character of the texts in the symbol font (SetAnlvStrings: "take only the very first character")
            string bullet = before + after > 0 ? Text(0, 1) : "•";
            return WordSymbolFontBullet.Resolve(bullet, fonts.NameOf(fontIndex)).text;
        }
        int nfc = anlvSource[anlvAt];
        string ordinal = nfc is 5 or 7 ? "." : "";
        string prefixSpace = (anlvSource[anlvAt + 4] & 0x04) != 0 ? " " : "";     // fPrevSpace: "a single prefixing space character"
        return prefixSpace + Text(0, before) + NumberText(nfc, number) + ordinal + Text(before, after);
    }

    /// <summary>
    /// Where the ANLV that draws a paragraph's label lies (render time; the label text itself is counted at index time): the
    /// section's OLST level for outline levels 1-9 when the section has one, else the paragraph's ANLD - the same choice as
    /// <see cref="Next"/>. Null: no numbering.
    /// </summary>
    public static (byte[] source, int at)? LabelLevelOf(int levelType, byte[]? anld, byte[]? olst)
    {
        if (levelType is >= 1 and <= 9 && olst is { Length: >= OlstLength }) return (olst, (levelType - 1) * 16);
        if (levelType is >= 1 and <= 11 && anld is { Length: >= AnldLength }) return (anld, 0);
        return null;
    }

    /// <summary>
    /// The label's character formatting from its ANLV as Word 97 sprms (Microsoft "Word for Windows 6.0 Binary File Format",
    /// ANLD: fSetBold / fSetItalic / fSetSmallCaps / fSetCaps / fSetStrike / fSetKul say that fBold / fItalic / fSmallCaps /
    /// fCaps / fStrike / kul decide; ico "color of autonumber" (0 = auto: the paragraph's colour); ftc "font code of
    /// autonumber" (Word writes 0xFFFF / 0x8001 when the number has no font of its own: only a font in the table is applied);
    /// hps "font half point size (or 0=auto)"). LibreOffice reads only the font, for symbol bullets (ww8par2.cxx:577-578
    /// SetAnlvStrings). Bold / italic / font / size set both the Latin and the complex-script values: the label is drawn in
    /// one of them by its characters (<see cref="WordListLabelScript"/>).
    /// </summary>
    public static byte[] LabelCharacterSprms(byte[] source, int at, int fontCount)
    {
        byte flags1 = source[at + 3], flags2 = source[at + 4], flags3 = source[at + 5];
        var sprms = new List<byte>(40);
        void Toggle(ushort opcode, bool on) => sprms.AddRange([(byte)opcode, (byte)(opcode >> 8), on ? (byte)1 : (byte)0]);
        void Word(ushort opcode, int value) => sprms.AddRange([(byte)opcode, (byte)(opcode >> 8), (byte)value, (byte)(value >> 8)]);
        if ((flags1 & 0x10) != 0) { Toggle(0x0835, (flags2 & 0x08) != 0); Toggle(0x085C, (flags2 & 0x08) != 0); }   // fSetBold: fBold
        if ((flags1 & 0x20) != 0) { Toggle(0x0836, (flags2 & 0x10) != 0); Toggle(0x085D, (flags2 & 0x10) != 0); }   // fSetItalic: fItalic
        if ((flags1 & 0x40) != 0) Toggle(0x083A, (flags2 & 0x20) != 0);                                              // fSetSmallCaps: fSmallCaps
        if ((flags1 & 0x80) != 0) Toggle(0x083B, (flags2 & 0x40) != 0);                                              // fSetCaps: fCaps
        if ((flags2 & 0x01) != 0) Toggle(0x0837, (flags2 & 0x80) != 0);                                              // fSetStrike: fStrike
        if ((flags2 & 0x02) != 0) sprms.AddRange([0x3E, 0x2A, (byte)(flags3 & 0x07)]);                                // fSetKul: kul
        int color = flags3 >> 3;
        if (color != 0) sprms.AddRange([0x42, 0x2A, (byte)color]);                                                   // ico
        int font = BitConverter.ToUInt16(source, at + 6);
        if (font < fontCount) { Word(0x4A4F, font); Word(0x4A5E, font); }                                            // ftc
        int size = BitConverter.ToUInt16(source, at + 8);
        if (size != 0) { Word(0x4A43, size); Word(0x4A61, size); }                                                   // hps
        return sprms.ToArray();
    }

    /// <summary>
    /// The label's spacing from its ANLV: dxaIndent "width of prefix text (same as indent)" and dxaSpace "minimum space between
    /// number and paragraph" (Word 6 binary format, ANLD) - the Word 6 level that Word 97's LVL keeps as fWord6 dxaIndent /
    /// dxaSpace (<see cref="WordListLabelLegacy"/>). fHang "number will be displayed using a hanging indent": without it the
    /// paragraph's hanging indent does not place the text (LibreOffice ww8par2.cxx:546-554 SetBaseAnlv: with fHang the label
    /// hangs by |dxaIndent|, without it dxaIndent is only the distance to the text).
    /// </summary>
    public static WordListLabelLegacy LabelSpacing(byte[] source, int at)
        => new(Math.Max(0, (int)BitConverter.ToInt16(source, at + 14)), Math.Abs((int)BitConverter.ToInt16(source, at + 12)), Hangs: (source[at + 3] & 0x08) != 0);

    private static string NumberText(int nfc, int number) => nfc switch
    {
        1 => Roman(number).ToUpperInvariant(),
        2 => Roman(number),
        3 => Letters(number).ToUpperInvariant(),
        4 => Letters(number),
        _ => number.ToString(System.Globalization.CultureInfo.InvariantCulture),
    };

    private static string Roman(int number)
    {
        if (number <= 0) return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        (int value, string symbol)[] numerals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"), (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")];
        var text = new StringBuilder();
        foreach ((int value, string symbol) in numerals) while (number >= value) { text.Append(symbol); number -= value; }
        return text.ToString();
    }

    /// <summary>a..z, then aa..zz (the letter repeated), as Word numbers letters.</summary>
    private static string Letters(int number)
    {
        if (number <= 0) return number.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return new string((char)('a' + (number - 1) % 26), (number - 1) / 26 + 1);
    }
}
