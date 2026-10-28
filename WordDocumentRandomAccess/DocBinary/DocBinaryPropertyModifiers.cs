using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Applies a grpprl (a run of Sprm property modifiers, [MS-DOC] §2.6) to the format-neutral
/// paragraph/run formatting. Only the sprms that matter for HTML flow rendering are interpreted;
/// everything else is skipped by operand size. Fonts (ftc → Sttbfffn) are not resolved yet.
/// </summary>
public static class DocBinaryPropertyModifiers
{
    private static readonly string?[] IcoPalette =
    [
        null, "000000", "0000FF", "00FFFF", "00FF00", "FF00FF", "FF0000", "FFFF00", "FFFFFF",
        "000080", "008080", "008000", "800080", "800000", "808000", "808080", "C0C0C0",
    ];

    public static void ApplyParagraphSprms(WordParagraphFormatting formatting, byte[] grpprl, int start, int length,
                                           WordParagraphFormatting? styleBaseline = null)
    {
        foreach ((ushort opcode, int operandAt, int operandLength) in Sprms(grpprl, start, length))
        {
            switch (opcode)
            {
                case 0x2403: case 0x2461:
                    byte jc = grpprl[operandAt];
                    formatting.Alignment = jc switch
                    {
                        1 => WordParagraphAlignment.Center,
                        2 => WordParagraphAlignment.End,
                        3 or 4 or 5 => WordParagraphAlignment.Justify,
                        _ => WordParagraphAlignment.Start,
                    };
                    break;
                case 0x2441: formatting.IsRightToLeft = grpprl[operandAt] != 0; formatting.RunFormatting.IsRightToLeft = formatting.IsRightToLeft; break;
                case 0x840F: case 0x845E: formatting.LeftIndentTwips = BitConverter.ToInt16(grpprl, operandAt); break;
                case 0x840E: case 0x845D: formatting.RightIndentTwips = BitConverter.ToInt16(grpprl, operandAt); break;
                case 0x8411: case 0x8460: formatting.FirstLineIndentTwips = BitConverter.ToInt16(grpprl, operandAt); break;
                case 0xA413: formatting.SpaceBeforeTwips = BitConverter.ToUInt16(grpprl, operandAt); break;
                case 0xA414: formatting.SpaceAfterTwips = BitConverter.ToUInt16(grpprl, operandAt); break;
                case 0x6412:
                    short line = BitConverter.ToInt16(grpprl, operandAt);
                    formatting.IsLineSpacingAuto = BitConverter.ToInt16(grpprl, operandAt + 2) != 0;
                    formatting.LineSpacingTwipsOrAuto = Math.Abs((int)line);
                    break;
                case 0x2416: formatting.IsInTable = grpprl[operandAt] != 0; break;
                case 0x260A: formatting.NumberingLevel = grpprl[operandAt]; break;
                case 0x460B: formatting.NumberingId = BitConverter.ToUInt16(grpprl, operandAt); break;
                case 0x2640: formatting.OutlineLevel = grpprl[operandAt] is var level && level <= 8 ? level : -1; break;
            }
        }
    }

    public static void ApplyCharacterSprms(WordRunFormatting formatting, byte[] grpprl, int start, int length, WordRunFormatting? styleBaseline = null)
    {
        foreach ((ushort opcode, int operandAt, int operandLength) in Sprms(grpprl, start, length))
        {
            switch (opcode)
            {
                case 0x0835: case 0x085C: formatting.IsBold = Toggle(grpprl[operandAt], formatting.IsBold, styleBaseline?.IsBold); break;
                case 0x0836: case 0x085D: formatting.IsItalic = Toggle(grpprl[operandAt], formatting.IsItalic, styleBaseline?.IsItalic); break;
                case 0x0837: formatting.IsStrikethrough = Toggle(grpprl[operandAt], formatting.IsStrikethrough, styleBaseline?.IsStrikethrough); break;
                case 0x083C: formatting.IsHidden = Toggle(grpprl[operandAt], formatting.IsHidden, styleBaseline?.IsHidden); break;
                case 0x085A: formatting.IsRightToLeft = grpprl[operandAt] != 0; break;
                case 0x2A3E: formatting.IsUnderlined = grpprl[operandAt] != 0; break;
                case 0x2A48: formatting.IsSuperscript = grpprl[operandAt] == 1; formatting.IsSubscript = grpprl[operandAt] == 2; break;
                case 0x4A43: formatting.FontSizeHalfPoints = BitConverter.ToUInt16(grpprl, operandAt); break;
                case 0x4A61: formatting.ComplexScriptFontSizeHalfPoints = BitConverter.ToUInt16(grpprl, operandAt); break;
                case 0x2A42: formatting.ColorHex = grpprl[operandAt] < IcoPalette.Length ? IcoPalette[grpprl[operandAt]] : null; break;
                case 0x6870:
                    // COLORREF: red, green, blue, fAuto
                    if (grpprl[operandAt + 3] != 0) formatting.ColorHex = null;
                    else formatting.ColorHex = $"{grpprl[operandAt]:X2}{grpprl[operandAt + 1]:X2}{grpprl[operandAt + 2]:X2}";
                    break;
                case 0x2A0C:
                    byte ico = grpprl[operandAt];
                    formatting.HighlightCssColor = ico != 0 && ico < IcoPalette.Length ? "#" + IcoPalette[ico] : null;
                    break;
            }
        }
    }

    /// <summary>sprmPIstd / sprmCIstd inside a grpprl, when present.</summary>
    public static int? StyleIndexIn(byte[] grpprl, int start, int length)
    {
        foreach ((ushort opcode, int operandAt, _) in Sprms(grpprl, start, length))
            if (opcode is 0x4600 or 0x4A30) return BitConverter.ToUInt16(grpprl, operandAt);
        return null;
    }

    public static bool HasSpecialCharacterFlag(byte[] grpprl, int start, int length)
    {
        foreach ((ushort opcode, int operandAt, _) in Sprms(grpprl, start, length))
            if (opcode == 0x0855) return grpprl[operandAt] == 1;
        return false;
    }

    /// <summary>Toggle operands: 0 off, 1 on, 128 = as the style, 129 = opposite of the style.</summary>
    private static bool Toggle(byte operand, bool current, bool? styleValue) => operand switch
    {
        0 => false,
        1 => true,
        128 => styleValue ?? current,
        129 => !(styleValue ?? current),
        _ => current,
    };

    /// <summary>Walks a grpprl yielding (opcode, operand offset, operand length); operand size comes from the spra bits.</summary>
    public static IEnumerable<(ushort opcode, int operandAt, int operandLength)> Sprms(byte[] grpprl, int start, int length)
    {
        int position = start, end = Math.Min(grpprl.Length, start + length);
        while (position + 2 <= end)
        {
            ushort opcode = BitConverter.ToUInt16(grpprl, position);
            position += 2;
            int operandLength = (opcode >> 13) switch
            {
                0 or 1 => 1,
                2 or 4 or 5 => 2,
                3 => 4,
                7 => 3,
                _ => -1,                                   // 6: variable
            };
            if (operandLength < 0)
            {
                if (position >= end) yield break;
                if (opcode == 0xD608) { operandLength = BitConverter.ToUInt16(grpprl, position); position += 2; }   // sprmTDefTable
                else { operandLength = grpprl[position]; position += 1; }
            }
            if (position + operandLength > end) yield break;
            yield return (opcode, position, operandLength);
            position += operandLength;
        }
    }
}
