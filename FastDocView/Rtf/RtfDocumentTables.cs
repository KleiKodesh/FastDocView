using System.Globalization;

namespace FastDocView.Rtf;

/// <summary>
/// The RTF header tables a block needs to render or count: font table (name + code page), color table,
/// style sheet (name + outline level, for headings), list table + list override table (for list labels),
/// note numbering formats. Read once by the index pass and stored in the index, so a window never re-reads the header.
/// </summary>
public sealed class RtfDocumentTables
{
    public sealed class FontEntry { public string Name = ""; public int Charset = -1; public int CodePage = -1; }

    public sealed class StyleEntry { public string Name = ""; public int OutlineLevel = -1; }

    public sealed class ListLevelDefinition
    {
        public int NumberFormat;        // \levelnfc (MSONFC numbers, LibreOffice getNumberFormat)
        public int StartAt = 1;
        public string Text = "";        // \leveltext without its length character
        public byte[] NumberPositions = [];   // \levelnumbers: 1-based positions into the \leveltext (length character = position 0)
        public bool NoRestart;
        public bool Legal;
        public int Justification;       // \leveljc: 0 left, 1 center, 2 right (LibreOffice rtfdispatchvalue.cxx reads \leveljc, not \leveljcn)
        public int Follow;              // \levelfollow: 0 tab, 1 space, 2 nothing
        public int Picture = -1;        // \levelpicture: the picture bullet (an index into ListPictures), -1 none
        public int Font = -1;           // \f in the \listlevel group: the label's font (a symbol-font bullet is recoded against it)
        // The label's own character formatting: \af \fs \afs \cf \b \i in the \listlevel group = the level's rPr (LibreOffice
        // rtfdocumentimpl.cxx:2322 / rtfdispatchvalue.cxx:203 put them in the level's table sprms). -1 = not set. Word draws a
        // counted label with the paragraph mark's formatting plus these (Word 16 re-save of list-bullets-nolisttext-v1.rtf).
        public int ComplexFont = -1, FontSize = -1, ComplexFontSize = -1, Color = -1;
        public sbyte Bold = -1, ComplexBold = -1, Italic = -1, ComplexItalic = -1;
        public bool IsLegacy;                  // \levelold: a Word 6 level, drawn with \levelspace / \levelindent (WordListLabelLegacy)
        public int LegacySpace, LegacyIndent;
        public Html.WordListLabelLegacy? Legacy => IsLegacy ? new Html.WordListLabelLegacy(LegacySpace, LegacyIndent) : null;
    }

    /// <summary>A {\pict} of {\*\listpicture}: where its body starts in the file (its bytes are read on demand), its type and goal size.</summary>
    public sealed class ListPictureEntry { public long Offset; public string Extension = ""; public int WidthTwips, HeightTwips; }

    public sealed class ListDefinition { public int ListId; public ListLevelDefinition[] Levels = []; }

    public sealed class ListOverride { public int Index; public int ListId; public Dictionary<int, int> StartAtByLevel = new(); }

    public Dictionary<int, FontEntry> Fonts { get; } = new();
    /// <summary>{\mmathPr}: the document's math settings as control word -> parameter (WordOfficeMathSettings.FromRtf).</summary>
    public Dictionary<string, int> MathProperties { get; } = new(StringComparer.Ordinal);
    public List<string?> Colors { get; } = new();            // "RRGGBB"; null = auto (the empty first entry)
    public Dictionary<int, StyleEntry> Styles { get; } = new();
    public Dictionary<int, ListDefinition> Lists { get; } = new();
    public Dictionary<int, ListOverride> ListOverrides { get; } = new();
    public List<ListPictureEntry> ListPictures { get; } = new();
    public int DefaultFont { get; set; } = -1;
    public int AnsiCodePage { get; set; } = 1252;
    public string FootnoteFormat { get; set; } = "decimal";
    /// <summary>Word reads an RTF without \aftnn* as lower roman endnotes (test-data/word-notes: section 1 has none, Word 16 prints i).</summary>
    public string EndnoteFormat { get; set; } = "lowerRoman";
    public WordDocumentNoteRestart FootnoteRestart { get; set; }
    public WordDocumentNoteRestart EndnoteRestart { get; set; }
    public int FootnoteStart { get; set; } = 1;
    public int EndnoteStart { get; set; } = 1;
    public bool IsRightToLeftDocument { get; set; }
    /// <summary>\hyphauto: the document hyphenates automatically.</summary>
    public bool IsAutoHyphenation { get; set; }
    /// <summary>\deflang: the document default language (LCID), 0 = none.</summary>
    public ushort DefaultLanguage { get; set; }

    /// <summary>The code page text in font <paramref name="font"/> decodes with: its \fcharset / \cpg / name suffix, else \ansicpg.</summary>
    public ushort CodePageOfFont(int font)
    {
        if (Fonts.TryGetValue(font, out FontEntry? entry) && entry.CodePage > 0) return (ushort)entry.CodePage;
        return (ushort)AnsiCodePage;
    }

    public string? FontName(int font) => Fonts.TryGetValue(font, out FontEntry? entry) && entry.Name.Length > 0 ? entry.Name : null;

    public string? ColorHex(int index) => index > 0 && index < Colors.Count ? Colors[index] : null;

    /// <summary>Outline level 0-8 of a paragraph style: its \outlinelevel, else the built-in "heading N" name.</summary>
    public int OutlineLevelOfStyle(int style)
    {
        if (!Styles.TryGetValue(style, out StyleEntry? entry)) return -1;
        if (entry.OutlineLevel is >= 0 and <= 8) return entry.OutlineLevel;
        string name = entry.Name.Trim();
        if (name.StartsWith("heading ", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(name.Substring(8), NumberStyles.Integer, CultureInfo.InvariantCulture, out int level) && level is >= 1 and <= 9)
            return level - 1;
        return -1;
    }

    // ── serialisation (inside the sidecar index) ─────────────────────────────────────

    public void Write(BinaryWriter writer)
    {
        writer.Write(DefaultFont); writer.Write(AnsiCodePage);
        writer.Write(FootnoteFormat); writer.Write(EndnoteFormat); writer.Write(FootnoteStart); writer.Write(EndnoteStart);
        writer.Write((byte)FootnoteRestart); writer.Write((byte)EndnoteRestart);
        writer.Write(IsRightToLeftDocument);
        writer.Write(IsAutoHyphenation); writer.Write(DefaultLanguage);
        writer.Write(Fonts.Count);
        foreach (KeyValuePair<int, FontEntry> font in Fonts) { writer.Write(font.Key); writer.Write(font.Value.Name); writer.Write(font.Value.Charset); writer.Write(font.Value.CodePage); }
        writer.Write(Colors.Count);
        foreach (string? color in Colors) writer.Write(color ?? "");
        writer.Write(MathProperties.Count);
        foreach (KeyValuePair<string, int> property in MathProperties) { writer.Write(property.Key); writer.Write(property.Value); }
        writer.Write(Styles.Count);
        foreach (KeyValuePair<int, StyleEntry> style in Styles) { writer.Write(style.Key); writer.Write(style.Value.Name); writer.Write(style.Value.OutlineLevel); }
        // lists: the compact index counts labels again when it rescans a span
        writer.Write(Lists.Count);
        foreach (KeyValuePair<int, ListDefinition> list in Lists)
        {
            writer.Write(list.Key);
            writer.Write(list.Value.ListId);
            writer.Write(list.Value.Levels.Length);
            foreach (ListLevelDefinition level in list.Value.Levels)
            {
                writer.Write(level.NumberFormat); writer.Write(level.StartAt); writer.Write(level.Text);
                writer.Write(level.NumberPositions.Length); writer.Write(level.NumberPositions);
                writer.Write(level.NoRestart); writer.Write(level.Legal); writer.Write(level.Justification); writer.Write(level.Follow); writer.Write(level.Font); writer.Write(level.Picture);
                writer.Write(level.ComplexFont); writer.Write(level.FontSize); writer.Write(level.ComplexFontSize); writer.Write(level.Color);
                writer.Write(level.Bold); writer.Write(level.ComplexBold); writer.Write(level.Italic); writer.Write(level.ComplexItalic);
                writer.Write(level.IsLegacy); writer.Write(level.LegacySpace); writer.Write(level.LegacyIndent);
            }
        }
        writer.Write(ListOverrides.Count);
        foreach (KeyValuePair<int, ListOverride> listOverride in ListOverrides)
        {
            writer.Write(listOverride.Key);
            writer.Write(listOverride.Value.Index);
            writer.Write(listOverride.Value.ListId);
            writer.Write(listOverride.Value.StartAtByLevel.Count);
            foreach (KeyValuePair<int, int> startAt in listOverride.Value.StartAtByLevel) { writer.Write(startAt.Key); writer.Write(startAt.Value); }
        }
        writer.Write(ListPictures.Count);
        foreach (ListPictureEntry picture in ListPictures) { writer.Write(picture.Offset); writer.Write(picture.Extension); writer.Write(picture.WidthTwips); writer.Write(picture.HeightTwips); }
    }

    public static RtfDocumentTables Read(BinaryReader reader)
    {
        var tables = new RtfDocumentTables
        {
            DefaultFont = reader.ReadInt32(), AnsiCodePage = reader.ReadInt32(),
            FootnoteFormat = reader.ReadString(), EndnoteFormat = reader.ReadString(), FootnoteStart = reader.ReadInt32(), EndnoteStart = reader.ReadInt32(),
            FootnoteRestart = (WordDocumentNoteRestart)reader.ReadByte(), EndnoteRestart = (WordDocumentNoteRestart)reader.ReadByte(),
            IsRightToLeftDocument = reader.ReadBoolean(),
            IsAutoHyphenation = reader.ReadBoolean(),
            DefaultLanguage = reader.ReadUInt16(),
        };
        int fonts = reader.ReadInt32();
        for (int i = 0; i < fonts; i++) tables.Fonts[reader.ReadInt32()] = new FontEntry { Name = reader.ReadString(), Charset = reader.ReadInt32(), CodePage = reader.ReadInt32() };
        int colors = reader.ReadInt32();
        for (int i = 0; i < colors; i++) { string color = reader.ReadString(); tables.Colors.Add(color.Length == 0 ? null : color); }
        int mathProperties = reader.ReadInt32();
        if (mathProperties is < 0 or > 64) throw new InvalidDataException("RTF index: math properties damaged.");
        for (int i = 0; i < mathProperties; i++) tables.MathProperties[reader.ReadString()] = reader.ReadInt32();
        int styles = reader.ReadInt32();
        for (int i = 0; i < styles; i++) tables.Styles[reader.ReadInt32()] = new StyleEntry { Name = reader.ReadString(), OutlineLevel = reader.ReadInt32() };
        int lists = reader.ReadInt32();
        if (lists < 0) throw new InvalidDataException("RTF index: list table damaged.");
        for (int i = 0; i < lists; i++)
        {
            int key = reader.ReadInt32();
            var list = new ListDefinition { ListId = reader.ReadInt32() };
            int levelCount = reader.ReadInt32();
            if (levelCount < 0 || levelCount > 64) throw new InvalidDataException("RTF index: list table damaged.");
            list.Levels = new ListLevelDefinition[levelCount];
            for (int level = 0; level < levelCount; level++)
            {
                var definition = new ListLevelDefinition { NumberFormat = reader.ReadInt32(), StartAt = reader.ReadInt32(), Text = reader.ReadString() };
                int positions = reader.ReadInt32();
                if (positions < 0 || positions > 1024) throw new InvalidDataException("RTF index: list table damaged.");
                definition.NumberPositions = reader.ReadBytes(positions);
                if (definition.NumberPositions.Length != positions) throw new EndOfStreamException("RTF index: list table truncated.");
                definition.NoRestart = reader.ReadBoolean(); definition.Legal = reader.ReadBoolean();
                definition.Justification = reader.ReadInt32(); definition.Follow = reader.ReadInt32(); definition.Font = reader.ReadInt32(); definition.Picture = reader.ReadInt32();
                definition.ComplexFont = reader.ReadInt32(); definition.FontSize = reader.ReadInt32(); definition.ComplexFontSize = reader.ReadInt32(); definition.Color = reader.ReadInt32();
                definition.Bold = reader.ReadSByte(); definition.ComplexBold = reader.ReadSByte(); definition.Italic = reader.ReadSByte(); definition.ComplexItalic = reader.ReadSByte();
                definition.IsLegacy = reader.ReadBoolean(); definition.LegacySpace = reader.ReadInt32(); definition.LegacyIndent = reader.ReadInt32();
                list.Levels[level] = definition;
            }
            tables.Lists[key] = list;
        }
        int overrides = reader.ReadInt32();
        if (overrides < 0) throw new InvalidDataException("RTF index: list override table damaged.");
        for (int i = 0; i < overrides; i++)
        {
            int key = reader.ReadInt32();
            var listOverride = new ListOverride { Index = reader.ReadInt32(), ListId = reader.ReadInt32() };
            int startAts = reader.ReadInt32();
            if (startAts < 0 || startAts > 64) throw new InvalidDataException("RTF index: list override table damaged.");
            for (int level = 0; level < startAts; level++) listOverride.StartAtByLevel[reader.ReadInt32()] = reader.ReadInt32();
            tables.ListOverrides[key] = listOverride;
        }
        int pictures = reader.ReadInt32();
        if (pictures < 0) throw new InvalidDataException("RTF index: list pictures damaged.");
        for (int i = 0; i < pictures; i++)
            tables.ListPictures.Add(new ListPictureEntry { Offset = reader.ReadInt64(), Extension = reader.ReadString(), WidthTwips = reader.ReadInt32(), HeightTwips = reader.ReadInt32() });
        return tables;
    }
}
