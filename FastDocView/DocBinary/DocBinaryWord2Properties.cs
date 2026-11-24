namespace FastDocView.DocBinary;

/// <summary>
/// Word for Windows 1.x / 2.0 properties as Word 6 grpprls, which <see cref="DocBinaryWord6Sprms"/> then reads as for a Word 6
/// file - the version branch LibreOffice takes (eWW1 / eWW2). Sprm ids and operand sizes from ww8scan.cxx GetWW2SprmSearcher,
/// meanings from ww8par6.cxx GetWW2SprmDispatcher (the ids it reads; nullptr ones are dropped here too), the Word 6 ids they
/// become from sprmids.hxx namespace v6. A character run is no grpprl in Word 1/2 but a CHP prefix (Word2CHPX), read and
/// turned into sprms as ww8par2.cxx ReadWord2Chpx / ChpxToSprms do.
/// </summary>
public static class DocBinaryWord2Properties
{
    private const byte Fixed = 0, Variable = 1, Variable2 = 2;

    /// <summary>
    /// The Word 6 id this class writes for Word 2's chp.fBiDi (sprm 86 / the CHPX bit): Word 6 has no right-to-left run sprm
    /// (its 111-116 set only the complex-script bold, italic, font, language, colour and size), so a Word 2 file's run
    /// direction travels on an id no Word 6 file uses (unlisted in GetWW6SprmSearcher: one length byte) and
    /// <see cref="DocBinaryWord6Sprms"/> turns it into sprmCFBiDi (0x085A) for Word 2 files only.
    /// </summary>
    public const byte RightToLeftRun = 250;

    /// <summary>Per Word 2 sprm id: (operand length or extra length, variant), from GetWW2SprmSearcher; unlisted ids are
    /// variable (wwSprmParser::GetSprmInfo falls back to a length byte).</summary>
    private static readonly (byte length, byte variant)[] Sizes = BuildSizes();

    private static (byte, byte)[] BuildSizes()
    {
        var sizes = new (byte, byte)[256];
        for (int i = 0; i < 256; i++) sizes[i] = (0, Variable);
        void Set(int id, int length, byte variant = Fixed) => sizes[id] = ((byte)length, variant);
        void Range(int first, int last, int length) { for (int id = first; id <= last; id++) Set(id, length); }
        Set(0, 0); Set(2, 1); Set(3, 0, Variable); Range(4, 14, 1); Set(15, 0, Variable); Range(16, 22, 2); Set(23, 0, Variable);
        Set(24, 1); Set(25, 1); Range(26, 28, 2); Set(29, 1); Range(30, 36, 2); Set(37, 1); Range(38, 43, 2); Set(44, 1);
        Range(45, 49, 2); Set(50, 1); Set(51, 1); Set(52, 0); Range(53, 55, 1); Set(57, 0, Variable); Set(58, 0);
        Range(60, 67, 1); Set(68, 2); Set(69, 1); Set(70, 3); Set(71, 2); Set(72, 2); Range(73, 77, 1); Set(78, 0, Variable);
        Set(80, 1); Set(81, 1); Set(82, 2); Set(83, 2); Range(84, 87, 1); Set(94, 1); Set(95, 12, Variable); Range(96, 99, 2);
        Set(112, 1); Set(114, 1); Set(115, 2); Set(116, 2); Set(117, 1); Set(118, 1); Set(119, 2); Set(120, 2); Set(121, 1);
        Set(122, 1); Set(123, 2); Set(124, 2); Range(125, 128, 1); Range(129, 132, 2); Set(133, 1); Set(134, 1); Set(135, 2);
        Set(136, 2); Set(137, 1); Set(138, 1); Range(139, 148, 2); Set(149, 1); Set(152, 0, Variable2); Set(153, 2);
        Set(154, 0, Variable2); Set(155, 1, Variable); Set(157, 5); Set(158, 4); Set(159, 2); Set(160, 4); Set(161, 2);
        Set(162, 2); Set(163, 5); Set(164, 4);
        return sizes;
    }

    /// <summary>A Word 1/2 grpprl (PAPX, SEPX, style PAPX, Prm) as a Word 6 one.</summary>
    public static byte[] SprmsToWord6(byte[] grpprl, int start, int length)
    {
        var output = new List<byte>(length + 8);
        int position = start, end = Math.Min(grpprl.Length, start + length);
        while (position < end)
        {
            int id = grpprl[position];
            (int fixedLength, byte variant) = Sizes[id];
            int operandAt, operandLength;
            if (id == 23)
            {
                // sprmPChgTabs: a length byte, or 255 = counted from the deletions / insertions (ww8scan.cxx GetSprmTailLen)
                if (position + 1 >= end) break;
                operandAt = position + 2;
                if (grpprl[position + 1] != 255) operandLength = grpprl[position + 1];
                else
                {
                    int deleted = position + 2 < end ? grpprl[position + 2] : 0;
                    int insertedAt = position + 3 + 4 * deleted;
                    int inserted = insertedAt < end ? grpprl[insertedAt] : 0;
                    operandLength = 2 + 4 * deleted + 3 * inserted;
                }
            }
            else if (variant == Fixed) { operandAt = position + 1; operandLength = fixedLength; }
            else if (variant == Variable)
            {
                if (position + 1 >= end) break;
                operandAt = position + 2;
                operandLength = grpprl[position + 1] + fixedLength;
            }
            else
            {
                if (position + 2 >= end) break;
                operandAt = position + 3;
                operandLength = Math.Max(0, BitConverter.ToUInt16(grpprl, position + 1) - 1) + fixedLength;
            }
            if (operandAt + operandLength > end) break;                 // cut short: stop at the damaged sprm, as LibreOffice
            Convert(output, id, grpprl, operandAt, operandLength, position);
            position = operandAt + operandLength;
        }
        return output.ToArray();
    }

    private static void Convert(List<byte> output, int id, byte[] grpprl, int at, int length, int sprmAt)
    {
        switch (id)
        {
            case 2: output.AddRange([2, grpprl[at], 0]); return;               // sprmPIstd: a one-byte style code (ww8par.cxx Read_StyleCode, eWW2)
            case 3: case 12: case 13: case 18: case 87: return;                 // dropped: 3 / 18 no handler, 12 / 13 one-byte Word 2 forms of the Word 6 numbering sprms, 87 the diacritics' colour flag (not drawn)
            case 20:                                                            // sprmPDyaLine: Word 2 stores only dyaLine (the Word 6 LSPD adds fMultLinespace)
                output.AddRange([20, grpprl[at], grpprl[at + 1], 0, 0]);         // -> LSPD { dyaLine, fMultLinespace 0 }: > 0 at least, < 0 exactly, in twips
                return;
            case 50: output.AddRange([64, 1, grpprl[at]]); return;              // sprmPFBiDi (GetWW2SprmSearcher) -> the Word 6 right-to-left paragraph
            case 80: case 81: output.AddRange([(byte)(id + 31), 1, grpprl[at]]); return;   // sprmCFBoldBi / sprmCFItalicBi -> v6 111 / 112
            case 84: output.AddRange([98, grpprl[at]]); return;                 // sprmCIcoBi: the text colour (ww8par6.cxx:5765 Read_TextColor) -> v6 sprmCIco
            case 82: case 83: output.AddRange([(byte)(id + 31), 2, grpprl[at], grpprl[at + 1]]); return;   // sprmCFtcBi / sprmClidBi -> v6 113 / 114
            case 85: output.AddRange([116, 2, grpprl[at], 0]); return;          // sprmCHpsBi: one byte, as sprmCHps -> v6 116
            case 86: output.AddRange([RightToLeftRun, 1, grpprl[at]]); return;  // sprmCFBiDi
            case 15: case 23:                                                   // tabs: the same operand, with its length byte
                for (int i = sprmAt; i < at + length; i++) output.Add(grpprl[i]);
                return;
            case 55: output.AddRange([67, grpprl[at]]); return;                 // sprmCFFieldVanish -> v6 sprmCFFldVanish
            case 58: output.Add(83); return;                                    // sprmCPlain
            case 74: output.AddRange([99, grpprl[at], 0]); return;              // sprmCHps: one byte in Word 2 (Read_FontSize, eWW2)
            case 76: output.AddRange([101, grpprl[at], (byte)((sbyte)grpprl[at] < 0 ? 0xFF : 0)]); return;   // sprmCHpsPos: a signed byte (Read_SubSuperProp)
            case 154:                                                           // sprmTDefTable -> v6 sprmTDefTable, the same operand (ww8par2.cxx WW8TabBandDesc::ReadDef, IsSevenMinus)
                output.Add(190);
                for (int i = sprmAt + 1; i < at + length; i++) output.Add(grpprl[i]);
                return;
            case 155:                                                           // sprmTDefTableShd -> v6 191, a length byte
                output.Add(191);
                for (int i = sprmAt + 1; i < at + length; i++) output.Add(grpprl[i]);
                return;
        }
        int word6 = id switch
        {
            >= 3 and <= 52 => id,                                               // paragraph sprms keep their ids and sizes
            >= 60 and <= 73 or 75 or 77 => id + 25,                             // sprmCFBold .. sprmCIco, sprmCHpsInc, sprmCHpsPosAdj
            >= 115 and <= 145 and not 138 => id + 25,                           // section sprms (v6 sprmSDmBinFirst .. sprmSDzaGutter)
            >= 146 and <= 148 or 153 or >= 157 and <= 164 => id + 36,           // table sprms (ww8par2.cxx:1706-1730 eWW2 ids)
            _ => -1,
        };
        if (word6 < 0) return;
        output.Add((byte)word6);
        for (int i = 0; i < length; i++) output.Add(grpprl[at + i]);
    }

    /// <summary>
    /// A Word 1/2 CHPX (the first <paramref name="length"/> bytes of a Word2CHPX; the rest stays zero) as Word 6 sprms, as
    /// ww8par2.cxx ChpxToSprms: the bold .. vanish bits as toggles (128 + bit: 128 = as the style, 129 = its opposite), and
    /// the font, underline, language, colour, size and position only when their fs flag is set; the right-to-left values the
    /// same way (below); fSpec and its picture location are kept for the special
    /// characters (ReadWord2Chpx reads fSpec and fcPic).
    /// </summary>
    public static byte[] CharacterPropertiesToWord6(byte[] bytes, int start, int length)
    {
        var chpx = new byte[24];
        int available = Math.Max(0, Math.Min(Math.Min(length, chpx.Length), bytes.Length - start));
        Buffer.BlockCopy(bytes, start, chpx, 0, available);
        byte flags0 = chpx[0], flags1 = chpx[1], present = chpx[2];
        var output = new List<byte>(32)
        {
            85, (byte)(128 + (flags0 & 1)),                 // fBold
            86, (byte)(128 + ((flags0 >> 1) & 1)),          // fItalic
            87, (byte)(128 + ((flags1 >> 2) & 1)),          // fStrike
            88, (byte)(128 + ((flags0 >> 3) & 1)),          // fOutline
            90, (byte)(128 + ((flags0 >> 5) & 1)),          // fSmallCaps
            91, (byte)(128 + ((flags0 >> 6) & 1)),          // fCaps
            92, (byte)(128 + ((flags0 >> 7) & 1)),          // fVanish
        };
        if ((flags0 & 0x10) != 0) output.AddRange([67, 1]);                                     // fFldVanish
        if ((present & 0x02) != 0) output.AddRange([93, chpx[4], chpx[5]]);                    // fsFtc: ftc
        if ((present & 0x08) != 0) output.AddRange([94, (byte)(chpx[9] >> 5)]);                // fsKul: kul
        if ((present & 0x40) != 0) output.AddRange([97, chpx[12], chpx[13]]);                  // fsLid: lid
        if ((present & 0x01) != 0) output.AddRange([98, (byte)(chpx[9] & 0x1F)]);              // fsIco: ico
        if ((present & 0x04) != 0) output.AddRange([99, chpx[6], chpx[7]]);                    // fsHps: hps (a word in the CHPX)
        if ((present & 0x10) != 0) output.AddRange([101, chpx[10], (byte)((sbyte)chpx[10] < 0 ? 0xFF : 0)]);   // fsPos: hpsPos
        // right-to-left values (Word2CHPX, ww8par2.cxx:4181-4183, 4200, 4209-4211, 4246-4262 ReadWord2Chpx): fBoldBi / fItalicBi as
        // toggles like fBold, fBiDi, and icoBi / ftcBi / hpsBi / lidBi when their fs flag is set. LibreOffice's ChpxToSprms
        // (ww8par2.cxx:4090-4127) writes the flag fsFtcBi instead of ftcBi and two bytes for the one-byte sprm 85; read right here.
        output.AddRange([111, 1, (byte)(128 + ((flags1 >> 4) & 1))]);                         // fBoldBi
        output.AddRange([112, 1, (byte)(128 + ((flags1 >> 5) & 1))]);                         // fItalicBi
        if ((flags1 & 0x40) != 0) output.AddRange([RightToLeftRun, 1, 1]);                     // fBiDi
        byte presentBi = chpx[3];
        if ((present & 0x80) != 0 && (flags1 & 0x40) != 0) output.AddRange([98, (byte)(chpx[11] & 0x1F)]);    // fsIcoBi: icoBi, the colour of a right-to-left run (Word 97 has one colour; LibreOffice reads sprm 84 as the colour, ww8par6.cxx:5765)
        if ((presentBi & 0x01) != 0) output.AddRange([113, 2, chpx[14], chpx[15]]);            // fsFtcBi: ftcBi
        if ((presentBi & 0x02) != 0) output.AddRange([116, 2, chpx[16], chpx[17]]);            // fsHpsBi: hpsBi
        if ((presentBi & 0x04) != 0) output.AddRange([114, 2, chpx[18], chpx[19]]);            // fsLidBi: lidBi
        if ((flags1 & 0x02) != 0)
        {
            output.AddRange([117, 1]);                                                          // fSpec
            if (available >= 24) output.AddRange([68, 4, chpx[20], chpx[21], chpx[22], chpx[23]]);   // fcPic -> sprmCPicLocation
        }
        return output.ToArray();
    }
}
