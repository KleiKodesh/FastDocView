using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// The list tables of a .doc from the table stream: PlfLst ([MS-DOC] §2.9.201: cLst + LSTFs
/// §2.9.147, followed by the LVL array — one LVL per simple list, nine otherwise, which lcbPlfLst
/// does not count; FibRgFcLcb97 fcPlfLst) and PlfLfo (§2.9.200: lfoMac + LFOs §2.9.131 + LFODatas
/// §2.9.132 with their LFOLVL overrides §2.9.133). A paragraph's sprmPIlfo is a 1-based index into
/// rgLfo; the LFO's lsid names the LSTF; sprmPIlvl picks the level (§2.4.6.3). Labels are counted at
/// index time by <see cref="DocBinaryListCounter"/>, like the .docx side.
/// </summary>
public sealed class DocBinaryListDefinitions
{
    public const int MaxLevels = 9;
    private const int LstfSize = 28;      // lsid 4 + tplc 4 + rgistdPara 18 + flags 1 + grfhic 1
    private const int LfoSize = 16;       // lsid 4 + unused 8 + clfolvl 1 + ibstFltAutoNum 1 + grfhic 1 + unused 1
    private const int LvlfSize = 28;
    private const byte BulletFormatCode = 0x17;
    private const byte NoNumberFormatCode = 0xFF;

    /// <summary>One LVL (§2.9.149): its LVLF (§2.9.150) and the number text with level placeholders.</summary>
    public sealed class LevelDefinition
    {
        public int StartAt;
        public byte NumberFormatCode;              // MSONFC ([MS-OSHARED] §2.2.1.3)
        public byte Justification;                 // jc: 0 left, 1 center, 2 right
        public bool IsLegal;                       // fLegal: inherited numbers shown as decimal
        public bool NoRestart;                     // fNoRestart + RestartLimitLevel
        public byte RestartLimitLevel;             // ilvlRestartLim
        public byte[] PlaceholderOffsets = [];     // rgbxchNums: one-based offsets into NumberText, zero-terminated
        public byte FollowCharacter;               // ixchFollow: 0 tab, 1 space, 2 nothing
        public string NumberText = "";             // xst.rgtchar; a placeholder's char value is the zero-based level it stands for
        public int ParagraphGrpprlLength;
        public int CharacterGrpprlLength;

        public bool IsBullet => NumberFormatCode == BulletFormatCode;
        public bool HasNumberSequence => NumberFormatCode is not (BulletFormatCode or NoNumberFormatCode);

        /// <summary>The ECMA number-format name of this level's MSONFC, as <see cref="DocxHtml.DocxNumberingDefinitions.FormatNumber"/> expects.</summary>
        public string NumberFormat => NumberFormatName(NumberFormatCode);
    }

    /// <summary>One LSTF with its LVLs (one for a simple list, nine otherwise).</summary>
    public sealed class ListDefinition
    {
        public int ListId;
        public short[] LinkedStyleIndexes = new short[MaxLevels];   // rgistdPara, 0x0FFF = none
        public bool IsSimpleList;
        public LevelDefinition?[] Levels = new LevelDefinition?[MaxLevels];
    }

    /// <summary>One LFOLVL: a start-at override and/or a whole replacement LVL for one level.</summary>
    public sealed class LevelOverride
    {
        public int Level;
        public bool OverridesStart;
        public int StartAt;
        public LevelDefinition? Formatting;
    }

    /// <summary>One LFO + its LFOData.</summary>
    public sealed class ListOverride
    {
        public int ListId;
        public LevelOverride[] LevelOverrides = [];
    }

    private readonly ListDefinition[] _lists;
    private readonly ListOverride[] _overrides;
    private readonly Dictionary<int, ListDefinition> _listsById = new();

    public static DocBinaryListDefinitions Empty { get; } = new([], []);

    public IReadOnlyList<ListDefinition> Lists => _lists;
    public IReadOnlyList<ListOverride> Overrides => _overrides;
    public bool IsEmpty => _overrides.Length == 0 || _lists.Length == 0;

    private DocBinaryListDefinitions(ListDefinition[] lists, ListOverride[] overrides)
    {
        _lists = lists;
        _overrides = overrides;
        foreach (ListDefinition list in lists) _listsById[list.ListId] = list;      // first LSTF wins on a duplicate lsid
    }

    public static DocBinaryListDefinitions Read(Stream tableStream, (uint offset, uint length) plfLst, (uint offset, uint length) plfLfo)
    {
        if (plfLst.length < 2 || plfLfo.length < 4) return Empty;
        try
        {
            using var reader = new BinaryReader(tableStream, Encoding.Unicode, leaveOpen: true);
            return new DocBinaryListDefinitions(ReadLists(reader, plfLst), ReadOverrides(reader, plfLfo));
        }
        catch (Exception exception) when (exception is EndOfStreamException or IOException or ArgumentOutOfRangeException)
        {
            return Empty;          // a damaged list table costs the labels, never the document
        }
    }

    /// <summary>The LFO a paragraph's sprmPIlfo names (1-based), or null.</summary>
    public ListOverride? OverrideAt(int oneBasedIndex)
        => oneBasedIndex >= 1 && oneBasedIndex <= _overrides.Length ? _overrides[oneBasedIndex - 1] : null;

    public ListDefinition? ListOf(int listId) => _listsById.TryGetValue(listId, out ListDefinition? list) ? list : null;

    /// <summary>The LVL that formats a paragraph (LFOLVL.lvl when fFormatting, else the LSTF's level), or null.</summary>
    public LevelDefinition? LevelFor(ListOverride listOverride, ListDefinition list, int level)
    {
        foreach (LevelOverride levelOverride in listOverride.LevelOverrides)
            if (levelOverride.Level == level && levelOverride.Formatting is { } replaced) return replaced;
        return level >= 0 && level < MaxLevels ? list.Levels[level] : null;
    }

    public static LevelOverride? StartOverrideFor(ListOverride listOverride, int level)
    {
        foreach (LevelOverride levelOverride in listOverride.LevelOverrides)
            if (levelOverride.Level == level && levelOverride.OverridesStart) return levelOverride;
        return null;
    }

    /// <summary>
    /// The number text of a level with its placeholders replaced from the counters (§2.4.6.3 part 2):
    /// a placeholder is the char at each rgbxchNums offset, its value the level it shows; a bullet
    /// xst is its first char, with symbol-font codes (0xF0xx) shown as a plain bullet.
    /// </summary>
    public string FormatLabel(ListOverride listOverride, ListDefinition list, int level, LevelDefinition definition, int[] counters)
    {
        string text = definition.NumberText;
        if (definition.IsBullet)
        {
            if (text.Length == 0) return "•";
            return text[0] is (>= '' and <= '') or < ' ' ? "•" : text.Substring(0, 1);   // symbol-font code (xchBullet & 0xF000)
        }
        if (!definition.HasNumberSequence && definition.PlaceholderOffsets.Length == 0) return text;

        var placeholderAt = new HashSet<int>();
        foreach (byte offset in definition.PlaceholderOffsets)
        {
            if (offset == 0) break;
            placeholderAt.Add(offset - 1);
        }
        var label = new StringBuilder(text.Length + 8);
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (!placeholderAt.Contains(i) || c >= MaxLevels) { if (c >= ' ') label.Append(c); continue; }
            int placeholderLevel = c;
            LevelDefinition? placeholderDefinition = LevelFor(listOverride, list, placeholderLevel);
            if (placeholderDefinition is null || !placeholderDefinition.HasNumberSequence) continue;      // spec: ignored
            string format = placeholderDefinition.NumberFormat;
            if (definition.IsLegal && format != "decimalZero") format = "decimal";                         // fLegal (LVLF)
            label.Append(DocxHtml.DocxNumberingDefinitions.FormatNumber(counters[placeholderLevel], format));
        }
        return label.ToString();
    }

    /// <summary>MSONFC ([MS-OSHARED] §2.2.1.3) → ECMA-376 ST_NumberFormat names the shared formatter knows; the rest fall back to decimal.</summary>
    public static string NumberFormatName(byte code) => code switch
    {
        0x00 => "decimal",
        0x01 => "upperRoman",
        0x02 => "lowerRoman",
        0x03 => "upperLetter",
        0x04 => "lowerLetter",
        0x16 => "decimalZero",
        0x17 => "bullet",
        0x2D => "hebrew1",
        0x2F => "hebrew2",
        0xFF => "none",
        _ => "decimal",
    };

    // ── PlfLst + LVL array ──────────────────────────────────────────────────────────────

    private static ListDefinition[] ReadLists(BinaryReader reader, (uint offset, uint length) plfLst)
    {
        reader.BaseStream.Position = plfLst.offset;
        int count = reader.ReadInt16();
        if (count <= 0) return [];
        count = Math.Min(count, (int)((plfLst.length - 2) / LstfSize));             // bounded by lcbPlfLst
        var lists = new ListDefinition[count];
        for (int i = 0; i < count; i++)
        {
            var list = new ListDefinition { ListId = reader.ReadInt32() };
            reader.ReadInt32();                                                        // tplc
            for (int level = 0; level < MaxLevels; level++) list.LinkedStyleIndexes[level] = reader.ReadInt16();
            byte flags = reader.ReadByte();
            reader.ReadByte();                                                         // grfhic
            list.IsSimpleList = (flags & 0x01) != 0;
            lists[i] = list;
        }
        // The LVL array follows the PlfLst at offset + lcbPlfLst, in LSTF order.
        reader.BaseStream.Position = plfLst.offset + plfLst.length;
        foreach (ListDefinition list in lists)
        {
            int levelCount = list.IsSimpleList ? 1 : MaxLevels;
            for (int level = 0; level < levelCount; level++) list.Levels[level] = ReadLevel(reader);
        }
        return lists;
    }

    /// <summary>LVL: LVLF (28 bytes), grpprlPapx, grpprlChpx, xst (cch + UTF-16 chars).</summary>
    private static LevelDefinition ReadLevel(BinaryReader reader)
    {
        var definition = new LevelDefinition { StartAt = reader.ReadInt32(), NumberFormatCode = reader.ReadByte() };
        byte flags = reader.ReadByte();
        definition.Justification = (byte)(flags & 0x03);
        definition.IsLegal = (flags & 0x04) != 0;
        definition.NoRestart = (flags & 0x08) != 0;
        definition.PlaceholderOffsets = reader.ReadBytes(9);
        definition.FollowCharacter = reader.ReadByte();
        reader.ReadInt32();                                                            // dxaIndentSav
        reader.ReadInt32();                                                            // unused2
        definition.CharacterGrpprlLength = reader.ReadByte();
        definition.ParagraphGrpprlLength = reader.ReadByte();
        definition.RestartLimitLevel = reader.ReadByte();
        reader.ReadByte();                                                             // grfhic
        reader.BaseStream.Position += definition.ParagraphGrpprlLength + definition.CharacterGrpprlLength;   // indent / number-run formatting: not rendered
        int characterCount = reader.ReadUInt16();
        byte[] characters = reader.ReadBytes(characterCount * 2);
        definition.NumberText = Encoding.Unicode.GetString(characters, 0, characters.Length);
        if (definition.StartAt < 0) definition.StartAt = 0;
        return definition;
    }

    // ── PlfLfo ──────────────────────────────────────────────────────────────────────────

    private static ListOverride[] ReadOverrides(BinaryReader reader, (uint offset, uint length) plfLfo)
    {
        reader.BaseStream.Position = plfLfo.offset;
        long end = plfLfo.offset + plfLfo.length;
        int count = reader.ReadInt32();
        if (count <= 0) return [];
        count = Math.Min(count, (int)((plfLfo.length - 4) / LfoSize));
        var overrides = new ListOverride[count];
        var overrideLevelCounts = new int[count];
        for (int i = 0; i < count; i++)
        {
            overrides[i] = new ListOverride { ListId = reader.ReadInt32() };
            reader.ReadInt32();                                                        // unused1
            reader.ReadInt32();                                                        // unused2
            overrideLevelCounts[i] = reader.ReadByte();
            reader.ReadBytes(3);                                                       // ibstFltAutoNum, grfhic, unused3
        }
        for (int i = 0; i < count && reader.BaseStream.Position + 4 <= end; i++)
        {
            // LFOData: cp (4 bytes) then clfolvl LFOLVLs. LibreOffice (ww8par3.cxx) tolerates extra 0xFFFFFFFF
            // dwords before the first LFOLVL; an LFOLVL's iStartAt is never 0xFFFFFFFF (it MUST be >= 0), so skipping them is safe.
            reader.ReadUInt32();
            int levelCount = Math.Min(overrideLevelCounts[i], MaxLevels);
            if (levelCount == 0) continue;                 // the next LFOData's cp may itself be 0xFFFFFFFF: skip nothing here
            while (reader.BaseStream.Position + 4 <= end && PeekUInt32(reader) == 0xFFFFFFFF) reader.ReadUInt32();
            var levelOverrides = new List<LevelOverride>(levelCount);
            for (int j = 0; j < levelCount && reader.BaseStream.Position + 8 <= end; j++)
            {
                int startAt = reader.ReadInt32();
                uint bits = reader.ReadUInt32();
                var levelOverride = new LevelOverride
                {
                    Level = (int)(bits & 0x0F),
                    OverridesStart = (bits & 0x10) != 0,
                    StartAt = Math.Max(0, startAt),
                };
                if ((bits & 0x20) != 0)
                {
                    levelOverride.Formatting = ReadLevel(reader);
                    levelOverride.OverridesStart = false;          // spec: iStartAt is undefined when fFormatting is also set
                }
                levelOverrides.Add(levelOverride);
            }
            overrides[i].LevelOverrides = levelOverrides.ToArray();
        }
        return overrides;
    }

    private static uint PeekUInt32(BinaryReader reader)
    {
        long position = reader.BaseStream.Position;
        uint value = reader.ReadUInt32();
        reader.BaseStream.Position = position;
        return value;
    }
}
