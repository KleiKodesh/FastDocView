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

    /// <param name="word6Sprms">Word 6/95: style names are 8-bit Pascal strings and the UPX grpprls one-byte sprms (translated);
    /// null for Word 97+.</param>
    public DocBinaryStyleSheet(Stream tableStream, (uint offset, uint length) location, DocBinaryWord6Sprms? word6Sprms = null)
    {
        byte[]? bytes = location.length < 6 ? null : DocBinaryStructureReader.TryRead(tableStream, location);
        if (bytes is null) { _styles = []; return; }       // no style sheet, or one outside the table stream: document defaults only
        if (word6Sprms is { IsWord2: true }) { _styles = ReadWord2Styles(bytes, word6Sprms); return; }

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
            if (available >= MinimumStdfBaseLength) _styles[index] = ParseStyle(bytes, position, available, stdBaseLength, index, word6Sprms);
            position += stdLength;
        }
    }

    private static StyleDefinition ParseStyle(byte[] bytes, int start, int length, int stdBaseLength, int index, DocBinaryWord6Sprms? word6Sprms)
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
        if (word6Sprms is not null)
        {
            // Word 6/95: a Pascal string (length byte, 8-bit characters, a null) read as cp1252 (ww8scan.cxx WW8Style::Read1Style, case 6 / 7)
            int nameBytes = bytes[position];
            if (position + 1 + nameBytes <= start + length)
            {
                var name = new StringBuilder(nameBytes);
                for (int i = 0; i < nameBytes; i++) name.Append(DocBinaryTextDecoder.DecodeCompressed(bytes[position + 1 + i]));
                style.Name = name.ToString();
            }
            position += nameBytes + 2;
        }
        else
        {
            int nameLength = BitConverter.ToUInt16(bytes, position);
            position += 2;
            if (position + nameLength * 2 <= start + length) style.Name = Encoding.Unicode.GetString(bytes, position, nameLength * 2);
            position += nameLength * 2 + 2;                    // + terminating null
        }
        if ((position - start) % 2 != 0) position++;

        for (int upx = 0; upx < upxCount && position + 2 <= start + length; upx++)
        {
            int upxLength = BitConverter.ToUInt16(bytes, position);
            position += 2;
            int available = Math.Max(0, Math.Min(upxLength, start + length - position));
            var grpprl = new byte[available];
            Buffer.BlockCopy(bytes, position, grpprl, 0, available);
            if (word6Sprms is not null)
            {
                // the paragraph UPX starts with the istd (2 bytes, kept); the rest is a Word 6 grpprl (ww8par2.cxx WW8RStyle::ImportUPX)
                int keep = style.IsParagraphStyle && upx == 0 ? Math.Min(2, available) : 0;
                byte[] word97 = word6Sprms.ToWord97(grpprl, keep, available - keep);
                var translated = new byte[keep + word97.Length];
                Buffer.BlockCopy(grpprl, 0, translated, 0, keep);
                Buffer.BlockCopy(word97, 0, translated, keep, word97.Length);
                grpprl = translated;
            }
            if (style.IsParagraphStyle && upx == 0) style.ParagraphGrpprl = grpprl;      // starts with istd (2 bytes)
            else style.CharacterGrpprl = grpprl;
            position += upxLength;
            if ((position - start) % 2 != 0) position++;
        }
        return style;
    }

    /// <summary>
    /// A Word 1/2 style sheet, as LibreOffice's ww8par2.cxx WW8RStyle::ImportOldFormatStyles reads it: cstcStd, then three
    /// length-prefixed lists in style-slot order (names: a length byte, 0xFF = undefined, 0 = the built-in name; CHPXs: a
    /// length byte, 0xFF = none, then a Word2CHPX; PAPXs: a length byte, 0xFF = none, then stc, a 6-byte PHE and the grpprl),
    /// then iMac pairs of (stcNext, stcBase). Slot n is style code (n - cstcStd) &amp; 255 - a paragraph's istd is that code;
    /// a base of 222 (or the style itself) is none. The built-in id comes from the style code (styles.cxx GetCanonicalStiFromStc:
    /// 0 Normal, 254 .. 246 Heading 1 .. 9).
    /// </summary>
    private static StyleDefinition?[] ReadWord2Styles(byte[] bytes, DocBinaryWord6Sprms word2Sprms)
    {
        var styles = new StyleDefinition?[256];
        int standardCount = BitConverter.ToUInt16(bytes, 0);
        int position = 2;
        var names = new List<string?>();
        int listEnd = Math.Min(bytes.Length, position + ReadLength(bytes, position));
        for (int at = position + 2; at < listEnd;)
        {
            int count = bytes[at++];
            string? name = null;
            if (count != 0xFF && count != 0)
            {
                var builder = new StringBuilder(count);
                for (int i = 0; i < count && at + i < bytes.Length; i++) builder.Append(DocBinaryTextDecoder.DecodeCompressed(bytes[at + i]));
                name = builder.ToString();
                at += count;
            }
            names.Add(count == 0xFF ? null : name ?? "");
        }
        position = listEnd;

        var characterProperties = new List<byte[]>();
        listEnd = position + 2 <= bytes.Length ? Math.Min(bytes.Length, position + ReadLength(bytes, position)) : position;
        for (int at = position + 2; at < listEnd && characterProperties.Count < names.Count;)
        {
            int cb = bytes[at++];
            if (cb == 0xFF) { characterProperties.Add([]); continue; }
            int available = Math.Max(0, Math.Min(cb, bytes.Length - at));
            characterProperties.Add(word2Sprms.CharacterPropertiesToWord97(bytes, at, available));
            at += cb;
        }
        position = Math.Max(position, listEnd);

        var paragraphProperties = new List<byte[]?>();
        listEnd = position + 2 <= bytes.Length ? Math.Min(bytes.Length, position + ReadLength(bytes, position)) : position;
        for (int at = position + 2; at < listEnd && paragraphProperties.Count < characterProperties.Count;)
        {
            int cb = bytes[at++];
            if (cb == 0xFF) { paragraphProperties.Add(null); continue; }
            paragraphProperties.Add(cb < 7 ? [] : word2Sprms.ToWord97(bytes, at + 7, Math.Max(0, Math.Min(cb - 7, bytes.Length - at - 7))));
            at += cb;
        }
        position = Math.Max(position, listEnd);

        int definedCount = position + 2 <= bytes.Length ? Math.Min(BitConverter.ToUInt16(bytes, position), names.Count) : 0;
        position += 2;
        for (int slot = 0; slot < definedCount && position + 2 <= bytes.Length; slot++, position += 2)
        {
            int code = (slot - standardCount) & 255;
            int baseCode = bytes[position + 1];
            int builtInId = BuiltInIdOfStyleCode(code);
            if (builtInId < 0 || slot >= paragraphProperties.Count) continue;     // stiNil, or no PAPX slot: LibreOffice skips it
            byte[] papx = paragraphProperties[slot] ?? [];
            var grpprl = new byte[2 + papx.Length];
            grpprl[0] = (byte)code;
            Buffer.BlockCopy(papx, 0, grpprl, 2, papx.Length);
            bool isCharacterStyle = builtInId is CharacterStyleFootnoteReference or CharacterStyleCommentReference or CharacterStyleLineNumber
                                    && papx.Length == 0;   // ww8::StandardStiIsCharStyle with no PAPX
            styles[code] = new StyleDefinition
            {
                Index = code,
                Name = names[slot] is { Length: > 0 } own ? own : $"Style {code}",
                BuiltInId = builtInId,
                BaseStyleIndex = baseCode == 222 || baseCode == code ? -1 : baseCode,
                IsParagraphStyle = !isCharacterStyle,
                ParagraphGrpprl = grpprl,
                CharacterGrpprl = slot < characterProperties.Count ? characterProperties[slot] : [],
            };
        }
        return styles;
    }

    private static int ReadLength(byte[] bytes, int at) => at + 2 <= bytes.Length ? BitConverter.ToUInt16(bytes, at) : 0;

    // sti values (LibreOffice sw/source/filter/inc/wwstyles.hxx:74-76) of the Word 2 built-in character styles (StandardStiIsCharStyle)
    private const int CharacterStyleFootnoteReference = 38, CharacterStyleCommentReference = 39, CharacterStyleLineNumber = 40, UserStyle = 0xFFE;

    /// <summary>styles.cxx GetCanonicalStiFromStc: 0 Normal, 1-221 user styles, 222 none (-1), 223.. the built-ins in its table order.</summary>
    private static int BuiltInIdOfStyleCode(int code)
    {
        if (code == 0) return 0;
        if (code < 222) return UserStyle;
        return code switch
        {
            222 => -1,
            223 => CharacterStyleCommentReference,
            240 => CharacterStyleLineNumber,
            244 => CharacterStyleFootnoteReference,
            >= 246 and <= 254 => 255 - code,                   // stiLev9 .. stiLev1
            _ => UserStyle - 1,                                 // another built-in (TOC, index, header, footer, notes): no heading
        };
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
