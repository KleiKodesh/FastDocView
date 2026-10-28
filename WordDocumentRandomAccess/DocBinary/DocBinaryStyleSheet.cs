using System.Text;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// The STSH style sheet ([MS-DOC] §2.9.271) from the table stream: for every style its name,
/// base style, kind, built-in id (sti 1–9 = Heading 1–9) and its UPX grpprls, resolved through the
/// base chain into <see cref="WordParagraphFormatting"/> on demand (cached).
/// </summary>
public sealed class DocBinaryStyleSheet
{
    public sealed class StyleDefinition
    {
        public int Index;
        public string Name = "";
        public int BuiltInId;                 // sti
        public int BaseStyleIndex = -1;       // istdBase (0xFFF = none)
        public bool IsParagraphStyle;
        public byte[] ParagraphGrpprl = [];
        public byte[] CharacterGrpprl = [];
    }

    private readonly StyleDefinition?[] _styles;
    private readonly Dictionary<int, WordParagraphFormatting> _resolved = new();

    /// <summary>Set once after construction; character sprms in style UPXs name fonts by index.</summary>
    public DocBinaryFontTable Fonts { get; set; } = DocBinaryFontTable.Empty;

    public WordParagraphFormatting DocumentDefaults { get; } = new()
    {
        RunFormatting = new WordRunFormatting { FontSizeHalfPoints = 20, FontName = "Times New Roman" },
    };

    public DocBinaryStyleSheet(Stream tableStream, (uint offset, uint length) location)
    {
        if (location.length < 4) { _styles = []; return; }
        var bytes = new byte[location.length];
        tableStream.Position = location.offset;
        int read = 0;
        while (read < bytes.Length) { int n = tableStream.Read(bytes, read, bytes.Length - read); if (n <= 0) break; read += n; }

        int stshiLength = BitConverter.ToUInt16(bytes, 0);
        int styleCount = BitConverter.ToUInt16(bytes, 2);
        int stdBaseLength = BitConverter.ToUInt16(bytes, 4);
        _styles = new StyleDefinition?[styleCount];

        int position = 2 + stshiLength;
        for (int index = 0; index < styleCount && position + 2 <= bytes.Length; index++)
        {
            int stdLength = BitConverter.ToUInt16(bytes, position);
            position += 2;
            if (stdLength == 0) continue;
            _styles[index] = ParseStyle(bytes, position, stdLength, stdBaseLength, index);
            position += stdLength;
        }
    }

    private static StyleDefinition ParseStyle(byte[] bytes, int start, int length, int stdBaseLength, int index)
    {
        var style = new StyleDefinition { Index = index };
        ushort word0 = BitConverter.ToUInt16(bytes, start);
        ushort word1 = BitConverter.ToUInt16(bytes, start + 2);
        ushort word2 = BitConverter.ToUInt16(bytes, start + 4);
        style.BuiltInId = word0 & 0x0FFF;
        int kind = word1 & 0x000F;
        style.IsParagraphStyle = kind == 1;
        int baseIndex = word1 >> 4;
        style.BaseStyleIndex = baseIndex == 0xFFF ? -1 : baseIndex;
        int upxCount = word2 & 0x000F;

        int position = start + stdBaseLength;
        int nameLength = BitConverter.ToUInt16(bytes, position);
        position += 2;
        if (position + nameLength * 2 <= start + length) style.Name = Encoding.Unicode.GetString(bytes, position, nameLength * 2);
        position += nameLength * 2 + 2;                        // + terminating null
        if ((position - start) % 2 != 0) position++;

        for (int upx = 0; upx < upxCount && position + 2 <= start + length; upx++)
        {
            int upxLength = BitConverter.ToUInt16(bytes, position);
            position += 2;
            int available = Math.Max(0, Math.Min(upxLength, start + length - position));
            var grpprl = new byte[available];
            Buffer.BlockCopy(bytes, position, grpprl, 0, available);
            if (style.IsParagraphStyle && upx == 0) style.ParagraphGrpprl = grpprl;      // starts with istd (2 bytes)
            else style.CharacterGrpprl = grpprl;
            position += upxLength;
            if ((position - start) % 2 != 0) position++;
        }
        return style;
    }

    public StyleDefinition? this[int index] => index >= 0 && index < _styles.Length ? _styles[index] : null;

    public WordParagraphFormatting ResolveParagraphStyle(int styleIndex)
    {
        if (_resolved.TryGetValue(styleIndex, out WordParagraphFormatting? cached)) return cached;
        WordParagraphFormatting formatting = DocumentDefaults.Clone();
        var chain = new List<StyleDefinition>();
        var seen = new HashSet<int>();
        for (StyleDefinition? style = this[styleIndex]; style is not null && seen.Add(style.Index) && chain.Count < 32; style = this[style.BaseStyleIndex])
            chain.Add(style);
        chain.Reverse();
        foreach (StyleDefinition style in chain)
        {
            if (style.ParagraphGrpprl.Length > 2) DocBinaryPropertyModifiers.ApplyParagraphSprms(formatting, style.ParagraphGrpprl, 2, style.ParagraphGrpprl.Length - 2);
            DocBinaryPropertyModifiers.ApplyCharacterSprms(formatting.RunFormatting, style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, fonts: Fonts);
        }
        formatting.StyleKey = styleIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        _resolved[styleIndex] = formatting;
        return formatting;
    }

    /// <summary>Heading level 1-9 for a paragraph style index, 0 otherwise.</summary>
    public int HeadingLevelOf(int styleIndex)
    {
        StyleDefinition? style = this[styleIndex];
        if (style is null) return 0;
        if (style.BuiltInId is >= 1 and <= 9) return style.BuiltInId;
        int outline = ResolveParagraphStyle(styleIndex).OutlineLevel;
        return outline is >= 0 and <= 8 ? outline + 1 : 0;
    }
}
