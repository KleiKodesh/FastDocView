using System.Text;
using FastDocView.Html;

namespace FastDocView.DocBinary;

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
    private const int MinimumStdfBaseLength = 6;     // the three words ParseStyle reads: sti, stk + istdBase, cupx
    private readonly Dictionary<int, WordParagraphFormatting> _resolved = new();

    /// <summary>
    /// Set once after construction; character sprms in style UPXs name fonts by index. Setting it also names the document's
    /// default fonts: the STSHI's ftcAsci (Latin) and ftcBi (complex script; ftcOther when the STSHI is shorter), which every
    /// style without a font of its own uses (LibreOffice ww8scan.cxx:6918-6930 reads them, ww8par2.cxx:3714-3726 Set1StyleDefaults).
    /// </summary>
    public DocBinaryFontTable Fonts
    {
        get => _fonts;
        set
        {
            _fonts = value;
            if (_defaultLatinFont >= 0 && value.NameOf(_defaultLatinFont) is { } latin) DocumentDefaults.RunFormatting.FontName = latin;
            if (_defaultComplexScriptFont >= 0 && value.NameOf(_defaultComplexScriptFont) is { } complex) DocumentDefaults.RunFormatting.ComplexScriptFontName = complex;
        }
    }
    private DocBinaryFontTable _fonts = DocBinaryFontTable.Empty;
    private int _defaultLatinFont = -1, _defaultComplexScriptFont = -1;

    public WordParagraphFormatting DocumentDefaults { get; } = new()
    {
        RunFormatting = new WordRunFormatting { FontSizeHalfPoints = 20, FontName = "Times New Roman", HasComplexScriptBoldItalic = true },
    };

    public DocBinaryStyleSheet(Stream tableStream, (uint offset, uint length) location)
    {
        byte[]? bytes = location.length < 6 ? null : DocBinaryStructureReader.TryRead(tableStream, location);
        if (bytes is null) { _styles = []; return; }       // no style sheet, or one outside the table stream: document defaults only

        int stshiLength = BitConverter.ToUInt16(bytes, 0);
        // Stshif ([MS-DOC] 2.9.271): rgftcStandardChpStsh = ftcAsci, ftcFE, ftcOther at 12..17, then the STSHI's ftcBi at 18
        if (stshiLength >= 14 && bytes.Length >= 2 + 14) _defaultLatinFont = BitConverter.ToUInt16(bytes, 2 + 12);
        if (stshiLength >= 18 && bytes.Length >= 2 + 18) _defaultComplexScriptFont = BitConverter.ToUInt16(bytes, 2 + 16);
        if (stshiLength >= 20 && bytes.Length >= 2 + 20) _defaultComplexScriptFont = BitConverter.ToUInt16(bytes, 2 + 18);
        int styleCount = BitConverter.ToUInt16(bytes, 2);
        int stdBaseLength = BitConverter.ToUInt16(bytes, 4);
        _styles = new StyleDefinition?[styleCount];

        int position = 2 + stshiLength;
        for (int index = 0; index < styleCount && position + 2 <= bytes.Length; index++)
        {
            int stdLength = BitConverter.ToUInt16(bytes, position);
            position += 2;
            if (stdLength == 0) continue;
            // An STD that runs past the style sheet is cut at its end; one too short for its fixed part (StdfBase) is no style.
            int available = Math.Min(stdLength, bytes.Length - position);
            if (available >= MinimumStdfBaseLength) _styles[index] = ParseStyle(bytes, position, available, stdBaseLength, index);
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
        if (position + 2 > start + length) return style;       // no room for the name (xstzName) or UPXs
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
