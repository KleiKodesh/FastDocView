namespace FastDocView.DocBinary;

/// <summary>
/// Rewrites a Word 6.0 / 95 grpprl (one-byte sprm ids) as the Word 97 grpprl the rest of the reader understands - the
/// "version branch" LibreOffice takes with its own tables: operand sizes from ww8scan.cxx GetWW6SprmSearcher (fixed,
/// L_VAR = one length byte, L_VAR2 = two; ids it does not list are L_VAR, wwSprmParser::GetSprmInfo), meanings from
/// ww8par6.cxx GetWW6SprmDispatcher, ids from sprmids.hxx namespace v6. Sprms LibreOffice ignores (nullptr in the
/// dispatcher) and ones nothing here reads are dropped. Structures that changed size are widened as LibreOffice reads
/// them: the 2-byte Word 6 BRC becomes a Brc80 (ww8scan.cxx WW8_BRC(WW8_BRCVer6)), a 10-byte Word 6 TC a 20-byte TC80
/// (ww8par2.cxx WW8TabBandDesc::ReadDef, bVer67).
/// </summary>
public sealed class DocBinaryWord6Sprms
{
    private const byte Fixed = 0, Variable = 1, Variable2 = 2;

    /// <summary>Per id: operand length (fixed) or extra length (variable) and the variant, from GetWW6SprmSearcher.</summary>
    private static readonly (byte length, byte variant)[] Sizes = BuildSizes(isFontCodeVariant: false);
    private static readonly (byte length, byte variant)[] FontCodeVariantSizes = BuildSizes(isFontCodeVariant: true);

    private readonly bool _isFontCodeVariant;
    private readonly (byte length, byte variant)[] _sizes;

    /// <param name="isFontCodeVariant">wIdent 0xA697-0xA699: sprms 111-113 are 2-byte font codes (ww8scan.cxx patchCJKVariant,
    /// ww8par6.cxx Read_AmbiguousSPRM); otherwise 111-116 are the right-to-left bold, italic, font, language, colour and size.</param>
    public DocBinaryWord6Sprms(bool isFontCodeVariant)
    {
        _isFontCodeVariant = isFontCodeVariant;
        _sizes = isFontCodeVariant ? FontCodeVariantSizes : Sizes;
    }

    private static (byte, byte)[] BuildSizes(bool isFontCodeVariant)
    {
        var sizes = new (byte, byte)[256];
        for (int i = 0; i < 256; i++) sizes[i] = (0, Variable);          // unknown Word 6/7 sprms: variable (GetSprmInfo)
        void Set(int id, int length, byte variant = Fixed) => sizes[id] = ((byte)length, variant);
        Set(0, 0);
        Set(2, 2); Set(3, 3, Variable); Set(4, 1); Set(5, 1); Set(6, 1); Set(7, 1); Set(8, 1); Set(9, 1); Set(10, 1); Set(11, 1);
        Set(12, 0, Variable); Set(13, 1); Set(14, 1); Set(15, 0, Variable); Set(16, 2); Set(17, 2); Set(18, 2); Set(19, 2);
        Set(20, 4); Set(21, 2); Set(22, 2); Set(23, 0, Variable); Set(24, 1); Set(25, 1); Set(26, 2); Set(27, 2); Set(28, 2);
        Set(29, 1);
        for (int id = 30; id <= 36; id++) Set(id, 2);
        Set(37, 1);
        for (int id = 38; id <= 43; id++) Set(id, 2);
        Set(44, 1); Set(45, 2); Set(46, 2); Set(47, 2); Set(48, 2); Set(49, 2); Set(50, 1); Set(51, 1); Set(52, 0);
        Set(64, 0, Variable);
        Set(65, 1); Set(66, 1); Set(67, 1); Set(68, 0, Variable); Set(69, 2); Set(70, 4); Set(71, 1); Set(72, 2); Set(73, 3);
        Set(74, 0, Variable); Set(75, 1); Set(77, 0, Variable); Set(79, 0, Variable);
        Set(80, 2); Set(81, 0, Variable); Set(82, 0, Variable); Set(83, 0);
        for (int id = 85; id <= 92; id++) Set(id, 1);
        Set(93, 2); Set(94, 1); Set(95, 3); Set(96, 2); Set(97, 2); Set(98, 1); Set(99, 2); Set(100, 1); Set(101, 2); Set(102, 1);
        Set(103, 0, Variable); Set(104, 1); Set(105, 0, Variable); Set(106, 0, Variable); Set(107, 2); Set(108, 0, Variable);
        Set(109, 2); Set(110, 2);
        for (int id = 111; id <= 116; id++) Set(id, 0, Variable);
        if (isFontCodeVariant) for (int id = 111; id <= 113; id++) Set(id, 2);
        Set(117, 1); Set(118, 1); Set(119, 1); Set(120, 12, Variable);
        for (int id = 121; id <= 124; id++) Set(id, 2);
        Set(131, 1); Set(132, 1); Set(133, 0, Variable); Set(136, 3); Set(137, 3); Set(138, 1); Set(139, 1); Set(140, 2); Set(141, 2);
        Set(142, 1); Set(143, 1); Set(144, 2); Set(145, 2); Set(146, 1); Set(147, 1); Set(148, 2); Set(149, 2); Set(150, 1);
        Set(151, 1); Set(152, 1); Set(153, 1); Set(154, 2); Set(155, 2); Set(156, 2); Set(157, 2); Set(158, 1); Set(159, 1);
        Set(160, 2); Set(161, 2); Set(162, 1); Set(163, 0);
        for (int id = 164; id <= 171; id++) Set(id, 2);
        Set(179, 0, Variable); Set(181, 0, Variable);
        Set(182, 2); Set(183, 2); Set(184, 2); Set(185, 1); Set(186, 1); Set(187, 12); Set(188, 0, Variable2); Set(189, 2);
        Set(190, 0, Variable2); Set(191, 1, Variable); Set(192, 4); Set(193, 5); Set(194, 4); Set(195, 2); Set(196, 4);
        Set(197, 2); Set(198, 2); Set(199, 5); Set(200, 4); Set(207, 0, Variable);
        return sizes;
    }

    /// <summary>
    /// Word 6 sprm id -> Word 97 opcode for the sprms whose operand keeps its bytes (sprmids.hxx v6 and Word 97 names; the
    /// operand sizes match the Word 97 spra). 0 = no direct equivalent (converted below, or dropped).
    /// </summary>
    private static readonly ushort[] SameOperand = BuildSameOperand();

    private static ushort[] BuildSameOperand()
    {
        var map = new ushort[256];
        // paragraph
        map[2] = 0x4600;   // sprmPIstd
        map[5] = 0x2403;   // sprmPJc (Word 6 has only the physical one = sprmPJc80)
        map[7] = 0x2405; map[8] = 0x2406; map[9] = 0x2407;   // sprmPFKeep, sprmPFKeepFollow, sprmPFPageBreakBefore
        map[14] = 0x240C;  // sprmPFNoLineNumb
        map[16] = 0x840E; map[17] = 0x840F; map[19] = 0x8411;   // sprmPDxaRight80, sprmPDxaLeft80, sprmPDxaLeft180
        map[20] = 0x6412;  // sprmPDyaLine (LSPD)
        map[21] = 0xA413; map[22] = 0xA414;   // sprmPDyaBefore, sprmPDyaAfter
        map[24] = 0x2416; map[25] = 0x2417;   // sprmPFInTable, sprmPFTtp
        map[26] = 0x8418; map[27] = 0x8419; map[28] = 0x841A; map[29] = 0x261B;   // sprmPDxaAbs, sprmPDyaAbs, sprmPDxaWidth, sprmPPc
        map[37] = 0x2423;  // sprmPWr
        map[44] = 0x242A; map[45] = 0x442B; map[46] = 0x442C; map[47] = 0x442D;   // sprmPFNoAutoHyph, sprmPWHeightAbs, sprmPDcs, sprmPShd80 (same 2-byte SHD: ww8par6.cxx SwWW8Shade)
        map[48] = 0x842E; map[49] = 0x842F; map[51] = 0x2431;   // sprmPDyaFromText, sprmPDxaFromText, sprmPFWidowControl
        // character
        map[65] = 0x0800; map[66] = 0x0801; map[67] = 0x0802;   // sprmCFRMarkDel, sprmCFRMarkIns, sprmCFFldVanish
        map[71] = 0x0806; map[75] = 0x080A;   // sprmCFData, sprmCFOle2
        map[80] = 0x4A30;  // sprmCIstd
        for (int id = 85; id <= 92; id++) map[id] = (ushort)(0x0835 + id - 85);   // sprmCFBold .. sprmCFVanish, same order
        map[93] = 0x4A4F;  // sprmCFtc -> sprmCRgFtc0 (ww8par6.cxx Read_FontCode: RES_CHRATR_FONT)
        map[94] = 0x2A3E; map[96] = 0x8840; map[97] = 0x486D;   // sprmCKul, sprmCDxaSpace, sprmCLid -> sprmCRgLid0_80
        map[98] = 0x2A42; map[99] = 0x4A43; map[101] = 0x4845;  // sprmCIco, sprmCHps, sprmCHpsPos
        map[104] = 0x2A48; map[107] = 0x484B;                    // sprmCIss, sprmCHpsKern
        map[117] = 0x0855; map[118] = 0x0856;                    // sprmCFSpec, sprmCFObj
        // section
        map[136] = 0xF203; map[137] = 0xF204; map[138] = 0x3005;   // sprmSDxaColWidth, sprmSDxaColSpacing, sprmSFEvenlySpaced
        map[142] = 0x3009; map[143] = 0x300A; map[144] = 0x500B; map[145] = 0x900C; map[147] = 0x300E;
        map[150] = 0x3011; map[156] = 0xB017; map[157] = 0xB018; map[158] = 0x3019; map[159] = 0x301A;
        map[161] = 0x501C; map[162] = 0x301D;
        map[164] = 0xB01F; map[165] = 0xB020; map[166] = 0xB021; map[167] = 0xB022; map[168] = 0x9023; map[169] = 0x9024; map[170] = 0xB025;
        // table
        map[182] = 0x5400; map[183] = 0x9601; map[184] = 0x9602; map[185] = 0x3403; map[186] = 0x3404; map[189] = 0x9407;
        map[194] = 0x7621; map[195] = 0x5622; map[196] = 0x7623; map[197] = 0x5624; map[198] = 0x5625;
        return map;
    }

    /// <summary>The Word 97 grpprl for a Word 6 one; empty when nothing in it has a Word 97 meaning here.</summary>
    public byte[] ToWord97(byte[] grpprl, int start, int length)
    {
        var output = new List<byte>(length + length / 2 + 4);
        int position = start, end = Math.Min(grpprl.Length, start + length);
        while (position < end)
        {
            int id = grpprl[position];
            (int fixedLength, byte variant) = _sizes[id];
            int operandAt, operandLength;
            if (id == 23)
            {
                // sprmPChgTabs: a length byte, or 255 = count from the deletions / insertions (ww8scan.cxx GetSprmTailLen)
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
                // sprmTDefTable / sprmTDefTable10: two bytes holding the operand length plus one
                if (position + 2 >= end) break;
                operandAt = position + 3;
                operandLength = Math.Max(0, BitConverter.ToUInt16(grpprl, position + 1) - 1) + fixedLength;
            }
            if (operandAt + operandLength > end) break;                 // cut short: as LibreOffice, stop at the damaged sprm
            Convert(output, id, grpprl, operandAt, operandLength);
            position = operandAt + operandLength;
        }
        return output.ToArray();
    }

    private void Convert(List<byte> output, int id, byte[] grpprl, int at, int length)
    {
        ushort same = SameOperand[id];
        if (same != 0)
        {
            int word97Length = (same >> 13) switch { 0 or 1 => 1, 2 or 4 or 5 => 2, 3 => 4, 7 => 3, _ => -1 };
            if (word97Length < 0) { Emit(output, same, grpprl, at, length, withLengthByte: true); return; }
            if (length < word97Length) return;
            Emit(output, same, grpprl, at, word97Length, withLengthByte: false);
            return;
        }
        switch (id)
        {
            case 12: EmitVariable(output, 0xC63E, grpprl, at, length); break;      // sprmPAnld -> sprmPAnld80 (DocBinaryWord6Numbering reads it)
            case 13: if (length >= 1) Emit(output, 0x240D, grpprl, at, 1, false); break;   // sprmPNLvlAnm -> sprmPNLvlAnm80 ([MS-DOC] 0x240D; ww8par6.cxx Read_ANLevelNo)
            case 15: EmitVariable(output, 0xC60D, grpprl, at, length); break;      // sprmPChgTabsPapx
            case 23: EmitVariable(output, 0xC615, grpprl, at, length); break;      // sprmPChgTabs
            case 38: case 39: case 40: case 41: case 42:                           // sprmPBrcTop .. sprmPBrcBetween (Word 6 BRC) -> Brc80
                if (length >= 2) EmitBrc80(output, (ushort)(0x6424 + id - 38), grpprl, at);
                break;
            case 64: if (length >= 1) output.AddRange([0x41, 0x24, grpprl[at]]); break;   // right-to-left paragraph (ww8par6.cxx Read_ParaBiDi) -> sprmPFBiDi
            case 68: if (length >= 4) Emit(output, 0x6A03, grpprl, at, 4, false); break;  // sprmCPicLocation (ww8par6.cxx Read_PicLoc: 4-byte fc)
            case 74:                                                               // sprmCSymbol: ftc (2), the 8-bit character (1) -> ftc, xchar
                if (length >= 3) output.AddRange([0x09, 0x6A, grpprl[at], grpprl[at + 1], grpprl[at + 2], 0]);
                break;
            case 83: output.AddRange([0x33, 0x2A, 0]); break;                     // sprmCPlain (no operand in Word 6)
            case 111: case 112:
                if (_isFontCodeVariant) { if (length >= 2) Emit(output, id == 111 ? (ushort)0x4A4F : (ushort)0x4A50, grpprl, at, 2, false); }   // font codes (Read_FontCode)
                else if (length >= 1) output.AddRange([id == 111 ? (byte)0x5C : (byte)0x5D, 0x08, grpprl[at]]);   // sprmCFBoldBi / sprmCFItalicBi (Read_BoldBiDiUsw)
                break;
            case 113: if (length >= 2) Emit(output, 0x4A5E, grpprl, at, 2, false); break;   // sprmCFtcBi (Read_FontCode: RES_CHRATR_CTL_FONT)
            case 114: if (length >= 2) Emit(output, 0x485F, grpprl, at, 2, false); break;   // sprmCLidBi
            case 115: if (length >= 1) Emit(output, 0x4A60, [grpprl[at], 0], 0, 2, false); break;   // sprmCIcoBi
            case 116: if (length >= 2) Emit(output, 0x4A61, grpprl, at, 2, false); break;   // sprmCHpsBi
            case 133: EmitVariable(output, 0xD202, grpprl, at, length); break;     // sprmSOlstAnm (Word 6 outline numbering)
            case 187:                                                              // sprmTTableBorders: six Word 6 BRCs -> six Brc80
                if (length >= 12)
                {
                    output.AddRange([0x05, 0xD6, 24]);
                    for (int i = 0; i < 6; i++) AddBrc80(output, grpprl, at + i * 2);
                }
                break;
            case 190: EmitTableDefinition(output, grpprl, at, length); break;      // sprmTDefTable
            case 191: EmitVariable(output, 0xD609, grpprl, at, length); break;     // sprmTDefTableShd -> sprmTDefTableShd80 (same 2-byte SHDs)
            case 193:                                                              // sprmTSetBrc: itcFirst, itcLim, sides, Word 6 BRC -> sprmTSetBrc80
                if (length >= 5)
                {
                    output.AddRange([0x20, 0xD6, 7, grpprl[at], grpprl[at + 1], grpprl[at + 2]]);
                    AddBrc80(output, grpprl, at + 3);
                }
                break;
        }
    }

    private static void Emit(List<byte> output, ushort opcode, byte[] source, int at, int length, bool withLengthByte)
    {
        output.Add((byte)opcode);
        output.Add((byte)(opcode >> 8));
        if (withLengthByte) output.Add((byte)Math.Min(255, length));
        for (int i = 0; i < Math.Min(length, withLengthByte ? 255 : length); i++) output.Add(source[at + i]);
    }

    private static void EmitVariable(List<byte> output, ushort opcode, byte[] source, int at, int length)
    {
        if (length > 255) return;                                          // a Word 97 variable sprm holds at most 255 bytes
        Emit(output, opcode, source, at, length, withLengthByte: true);
    }

    private static void EmitBrc80(List<byte> output, ushort opcode, byte[] source, int at)
    {
        output.Add((byte)opcode);
        output.Add((byte)(opcode >> 8));
        AddBrc80(output, source, at);
    }

    /// <summary>
    /// WW8_BRCVer6 (dxpLineWidth:3 in 0.75 pt - 6 / 7 = dotted / dashed -, brcType:2, fShadow:1, ico:5, dxpSpace:5) as a
    /// Brc80 (dptLineWidth in 1/8 pt, brcType, ico, dptSpace:5 | fShadow:1), ww8scan.cxx WW8_BRC::WW8_BRC(const WW8_BRCVer6&amp;).
    /// </summary>
    private static void AddBrc80(List<byte> output, byte[] source, int at)
    {
        byte low = source[at], high = source[at + 1];
        int width = low & 0x07, type = (low & 0x18) >> 3;
        bool shadow = (low & 0x20) != 0;
        int color = ((low & 0xC0) >> 6) | ((high & 0x07) << 2);
        int space = high >> 3;
        if (width > 5) { type = width; width = 1; }
        output.Add((byte)(width * 6));
        output.Add((byte)type);
        output.Add((byte)color);
        output.Add((byte)((space & 0x1F) | (shadow ? 0x20 : 0)));
    }

    /// <summary>
    /// sprmTDefTable: itcMac, itcMac + 1 cell edges, then Word 6 TCs of 10 bytes (two flag bytes, four BRCs) -> TC80s of 20
    /// (tcgrf, wWidth 0, four Brc80) as ww8par2.cxx WW8TabBandDesc::ReadDef reads them (nFileCols = length / 10 for bVer67).
    /// </summary>
    private static void EmitTableDefinition(List<byte> output, byte[] grpprl, int at, int length)
    {
        if (length < 1) return;
        int cellCount = grpprl[at];
        int edgesLength = 1 + (cellCount + 1) * 2;
        if (length < edgesLength) return;
        int savedCells = Math.Min((length - edgesLength) / 10, (0xFFFE - edgesLength) / 20);   // the Word 97 cb is 2 bytes
        var operand = new List<byte>(edgesLength + savedCells * 20);
        for (int i = 0; i < edgesLength; i++) operand.Add(grpprl[at + i]);
        for (int cell = 0; cell < savedCells; cell++)
        {
            int tc = at + edgesLength + cell * 10;
            operand.Add(grpprl[tc]);
            operand.Add(grpprl[tc + 1]);
            operand.Add(0);
            operand.Add(0);
            for (int side = 0; side < 4; side++) AddBrc80(operand, grpprl, tc + 2 + side * 2);
        }
        int cb = operand.Count + 1;                                        // TDefTableOperand cb counts the operand plus one
        output.AddRange([0x08, 0xD6, (byte)cb, (byte)(cb >> 8)]);
        output.AddRange(operand);
    }
}
