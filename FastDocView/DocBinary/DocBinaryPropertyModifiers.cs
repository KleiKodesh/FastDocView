using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// Applies a grpprl (a run of Sprm property modifiers, [MS-DOC] §2.6) to the format-neutral
/// paragraph/run formatting. Only the sprms that matter for HTML flow rendering are interpreted;
/// everything else is skipped by operand size (the audit lists which ones are ignored on purpose:
/// revision ids, language ids, proofing, pagination and frame hints). Fonts resolve through the
/// document's SttbfFfn (<see cref="DocBinaryFontTable"/>).
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
                    formatting.IsAlignmentPhysical = opcode == 0x2403;   // sprmPJc80 is physical, sprmPJc logical ([MS-DOC] 2.6.2)
                    break;
                case 0x2441: formatting.IsRightToLeft = grpprl[operandAt] != 0; formatting.RunFormatting.IsRightToLeft = formatting.IsRightToLeft; break;
                // sprmPDxaLeft80 / sprmPDxaRight80 are physical left / right, sprmPDxaLeft / sprmPDxaRight logical ([MS-DOC] 2.6.2)
                case 0x840F: case 0x845E: formatting.LeftIndentTwips = BitConverter.ToInt16(grpprl, operandAt); formatting.IsLeftIndentPhysical = opcode == 0x840F; break;
                case 0x840E: case 0x845D: formatting.RightIndentTwips = BitConverter.ToInt16(grpprl, operandAt); formatting.IsRightIndentPhysical = opcode == 0x840E; break;
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
                case 0x246D: formatting.IsContextualSpacing = grpprl[operandAt] != 0; break;                     // sprmPFContextualSpacing (sprmids.hxx)
                case 0xC64D: formatting.ShadingHex = ShadingBackground(grpprl, operandAt + 4, operandLength); break;      // SHDOperand: cvFore then cvBack
                case 0x442D: formatting.ShadingHex = Shading80Background(BitConverter.ToUInt16(grpprl, operandAt)); break;
                case 0xC60D: case 0xC615: formatting.FirstTabStopTwips = FirstAddedTabStop(grpprl, operandAt, operandLength, formatting.LeftIndentTwips) ?? formatting.FirstTabStopTwips; break;
                // paragraph borders: Brc80 (Word 97, 4 bytes) and Brc (Word 2000+, 8 bytes); top, left (= start), bottom, right (= end)
                case 0x6424: formatting.BorderTopCss = Brc80Css(grpprl, operandAt); break;
                case 0x6425: formatting.BorderStartCss = Brc80Css(grpprl, operandAt); break;
                case 0x6426: formatting.BorderBottomCss = Brc80Css(grpprl, operandAt); break;
                case 0x6427: formatting.BorderEndCss = Brc80Css(grpprl, operandAt); break;
                case 0xC64E: formatting.BorderTopCss = BrcCss(grpprl, operandAt, operandLength); break;
                case 0xC64F: formatting.BorderStartCss = BrcCss(grpprl, operandAt, operandLength); break;
                case 0xC650: formatting.BorderBottomCss = BrcCss(grpprl, operandAt, operandLength); break;
                case 0xC651: formatting.BorderEndCss = BrcCss(grpprl, operandAt, operandLength); break;
            }
        }
    }

    public static void ApplyCharacterSprms(WordRunFormatting formatting, byte[] grpprl, int start, int length, WordRunFormatting? styleBaseline = null,
                                           DocBinaryFontTable? fonts = null)
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
                case 0xCA71: formatting.HighlightCssColor = ShadingBackground(grpprl, operandAt + 4, operandLength) is { } back ? "#" + back : formatting.HighlightCssColor; break;
                case 0x4866: formatting.HighlightCssColor = Shading80Background(BitConverter.ToUInt16(grpprl, operandAt)) is { } back80 ? "#" + back80 : formatting.HighlightCssColor; break;
                case 0x083A: formatting.IsSmallCaps = Toggle(grpprl[operandAt], formatting.IsSmallCaps, styleBaseline?.IsSmallCaps); break;
                case 0x083B: formatting.IsAllCaps = Toggle(grpprl[operandAt], formatting.IsAllCaps, styleBaseline?.IsAllCaps); break;
                case 0x8840: formatting.CharacterSpacingTwips = BitConverter.ToInt16(grpprl, operandAt); break;
                case 0x4845: formatting.PositionHalfPoints = BitConverter.ToInt16(grpprl, operandAt); break;                            // sprmCHpsPos (LibreOffice sprmids.hxx)
                case 0x0811: formatting.IsHidden = Toggle(grpprl[operandAt], formatting.IsHidden, styleBaseline?.IsHidden); break;     // sprmCFWebHidden (sprmids.hxx; 0x0868 is CFUsePgsuSettings)
                case 0x2A53: formatting.IsStrikethrough = Toggle(grpprl[operandAt], formatting.IsStrikethrough, styleBaseline?.IsStrikethrough); break;   // double strike
                case 0x0838: formatting.IsOutline = Toggle(grpprl[operandAt], formatting.IsOutline, styleBaseline?.IsOutline); break;
                case 0x0839: formatting.IsShadow = Toggle(grpprl[operandAt], formatting.IsShadow, styleBaseline?.IsShadow); break;
                case 0x0858: formatting.IsEmboss = Toggle(grpprl[operandAt], formatting.IsEmboss, styleBaseline?.IsEmboss); break;
                case 0x0854: formatting.IsImprint = Toggle(grpprl[operandAt], formatting.IsImprint, styleBaseline?.IsImprint); break;
                case 0x6865: formatting.BorderCss = Brc80Css(grpprl, operandAt); break;                                                  // sprmCBrc80
                case 0xCA72: formatting.BorderCss = BrcCss(grpprl, operandAt, operandLength); break;                                     // sprmCBrc
                case 0x4A4F: formatting.FontName = fonts?.NameOf(BitConverter.ToUInt16(grpprl, operandAt)) ?? formatting.FontName; break;   // ASCII text
                case 0x4A5E: formatting.ComplexScriptFontName = fonts?.NameOf(BitConverter.ToUInt16(grpprl, operandAt)) ?? formatting.ComplexScriptFontName; break;  // bidi text
            }
        }
    }

    /// <summary>SHDOperand's Shd.cvBack (a COLORREF r,g,b,flag; flag 0xFF = auto) → "RRGGBB", null for auto/white.</summary>
    private static string? ShadingBackground(byte[] grpprl, int colorAt, int operandLength)
    {
        // Sprms() already consumed the operand's cb byte, so the operand IS the Shd and cvBack is at +4.
        int at = colorAt;
        if (operandLength < 10 || at + 4 > grpprl.Length || grpprl[at + 3] == 0xFF) return null;
        string hex = $"{grpprl[at]:X2}{grpprl[at + 1]:X2}{grpprl[at + 2]:X2}";
        return hex == "FFFFFF" ? null : hex;
    }

    /// <summary>Brc80 ([MS-DOC] §2.9.17): dptLineWidth (eighths of a point), brcType, ico, then spacing/flags.</summary>
    private static string? Brc80Css(byte[] grpprl, int at)
    {
        if (at + 4 > grpprl.Length) return null;
        byte width = grpprl[at], type = grpprl[at + 1], ico = grpprl[at + 2];
        if (type is 0 or 0xFF) return "none";
        string color = ico > 0 && ico < IcoPalette.Length ? "#" + IcoPalette[ico] : "#000";
        return BorderCss(width, type, color);
    }

    /// <summary>Brc ([MS-DOC] §2.9.16): cv (COLORREF), dptLineWidth, brcType, then spacing/flags. Sprms() already consumed the operand's cb.</summary>
    private static string? BrcCss(byte[] grpprl, int at, int operandLength)
    {
        if (operandLength < 8 || at + 8 > grpprl.Length) return null;
        byte width = grpprl[at + 4], type = grpprl[at + 5];
        if (type is 0 or 0xFF) return "none";
        string color = grpprl[at + 3] == 0xFF ? "#000" : $"#{grpprl[at]:X2}{grpprl[at + 1]:X2}{grpprl[at + 2]:X2}";
        return BorderCss(width, type, color);
    }

    private static string BorderCss(byte eighthsOfPoint, byte brcType, string color)
    {
        double width = Math.Max(0.25, eighthsOfPoint / 8.0);
        string style = brcType switch { 3 => "double", 6 => "dotted", 7 or 8 or 9 or 22 => "dashed", _ => "solid" };
        return $"{width.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt {style} {color}";
    }

    /// <summary>Shd80: icoFore (5 bits), icoBack (5 bits), ipat (6 bits) → the background ico as "RRGGBB".</summary>
    private static string? Shading80Background(ushort shd80)
    {
        int icoBack = (shd80 >> 5) & 0x1F;
        return icoBack > 0 && icoBack < IcoPalette.Length && IcoPalette[icoBack] != "FFFFFF" ? IcoPalette[icoBack] : null;
    }

    /// <summary>sprmPChgTabsPapx/sprmPChgTabs: cb, cTabsDel + positions, cTabsAdd + positions → the first added stop, relative to the indent.</summary>
    private static int? FirstAddedTabStop(byte[] grpprl, int operandAt, int operandLength, int leftIndentTwips)
    {
        int position = operandAt;
        if (position >= grpprl.Length) return null;
        int deleted = grpprl[position++];
        position += deleted * 2;
        if (position >= grpprl.Length) return null;
        int added = grpprl[position++];
        int first = int.MaxValue;
        for (int i = 0; i < added && position + 2 <= grpprl.Length; i++, position += 2) first = Math.Min(first, BitConverter.ToInt16(grpprl, position));
        return first == int.MaxValue || first <= 0 ? null : Math.Max(120, first - leftIndentTwips);
    }

    /// <summary>sprmPIstd / sprmCIstd inside a grpprl, when present.</summary>
    public static int? StyleIndexIn(byte[] grpprl, int start, int length)
    {
        foreach ((ushort opcode, int operandAt, _) in Sprms(grpprl, start, length))
            if (opcode is 0x4600 or 0x4A30) return BitConverter.ToUInt16(grpprl, operandAt);
        return null;
    }

    /// <summary>sprmCSymbol: ftc (2 bytes) then the character (2 bytes); symbol fonts use the private-use area.</summary>
    public static char? SymbolCharacterIn(byte[] grpprl, int start, int length)
    {
        foreach ((ushort opcode, int operandAt, int operandLength) in Sprms(grpprl, start, length))
        {
            if (opcode != 0x6A09 || operandLength < 4) continue;
            int character = BitConverter.ToUInt16(grpprl, operandAt + 2);
            if (character >= 0xF000) character -= 0xF000;
            return (char)character;
        }
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
                // sprmTDefTable: a 2-byte cb that counts the operand PLUS ONE ([MS-DOC] §2.6.3 TDefTableOperand).
                // Reading it as the plain length shifted every later sprm of the paragraph by a byte.
                if (opcode == 0xD608) { operandLength = Math.Max(0, BitConverter.ToUInt16(grpprl, position) - 1); position += 2; }
                else { operandLength = grpprl[position]; position += 1; }
            }
            if (position + operandLength > end) yield break;
            yield return (opcode, position, operandLength);
            position += operandLength;
        }
    }
}
