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
    }

    public sealed class ListDefinition { public int ListId; public ListLevelDefinition[] Levels = []; }

    public sealed class ListOverride { public int Index; public int ListId; public Dictionary<int, int> StartAtByLevel = new(); }

    public Dictionary<int, FontEntry> Fonts { get; } = new();
    public List<string?> Colors { get; } = new();            // "RRGGBB"; null = auto (the empty first entry)
    public Dictionary<int, StyleEntry> Styles { get; } = new();
    public Dictionary<int, ListDefinition> Lists { get; } = new();
    public Dictionary<int, ListOverride> ListOverrides { get; } = new();
    public int DefaultFont { get; set; } = -1;
    public int AnsiCodePage { get; set; } = 1252;
    public string FootnoteFormat { get; set; } = "decimal";
    public string EndnoteFormat { get; set; } = "decimal";
    public int FootnoteStart { get; set; } = 1;
    public int EndnoteStart { get; set; } = 1;
    public bool IsRightToLeftDocument { get; set; }

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
        writer.Write(IsRightToLeftDocument);
        writer.Write(Fonts.Count);
        foreach (KeyValuePair<int, FontEntry> font in Fonts) { writer.Write(font.Key); writer.Write(font.Value.Name); writer.Write(font.Value.Charset); writer.Write(font.Value.CodePage); }
        writer.Write(Colors.Count);
        foreach (string? color in Colors) writer.Write(color ?? "");
        writer.Write(Styles.Count);
        foreach (KeyValuePair<int, StyleEntry> style in Styles) { writer.Write(style.Key); writer.Write(style.Value.Name); writer.Write(style.Value.OutlineLevel); }
        // lists are needed only by the index pass (labels are stored per block)
    }

    public static RtfDocumentTables Read(BinaryReader reader)
    {
        var tables = new RtfDocumentTables
        {
            DefaultFont = reader.ReadInt32(), AnsiCodePage = reader.ReadInt32(),
            FootnoteFormat = reader.ReadString(), EndnoteFormat = reader.ReadString(), FootnoteStart = reader.ReadInt32(), EndnoteStart = reader.ReadInt32(),
            IsRightToLeftDocument = reader.ReadBoolean(),
        };
        int fonts = reader.ReadInt32();
        for (int i = 0; i < fonts; i++) tables.Fonts[reader.ReadInt32()] = new FontEntry { Name = reader.ReadString(), Charset = reader.ReadInt32(), CodePage = reader.ReadInt32() };
        int colors = reader.ReadInt32();
        for (int i = 0; i < colors; i++) { string color = reader.ReadString(); tables.Colors.Add(color.Length == 0 ? null : color); }
        int styles = reader.ReadInt32();
        for (int i = 0; i < styles; i++) tables.Styles[reader.ReadInt32()] = new StyleEntry { Name = reader.ReadString(), OutlineLevel = reader.ReadInt32() };
        return tables;
    }
}
