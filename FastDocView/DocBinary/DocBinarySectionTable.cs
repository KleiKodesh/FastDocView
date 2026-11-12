namespace FastDocView.DocBinary;

/// <summary>
/// The page text width of each section of a .doc ([MS-DOC] 2.8.26 PlcfSed: (n + 1) CPs then n
/// 12-byte Sed records; Sed.fcSepx points into the WordDocument stream at a Sepx: a 2-byte cb, then
/// a grpprl of section sprms). Only four sprms matter: sprmSXaPage (0xB01F), sprmSDxaLeft (0xB021),
/// sprmSDxaRight (0xB022), sprmSDzaGutter (0xB025). A section without a Sepx (fcSepx 0xFFFFFFFF), or
/// without one of the sprms, takes the Sep defaults - xaPage 12240, dxaLeft / dxaRight 1800, gutter 0
/// (Apache POI SEPAbstractType constructor; the RTF 1.5 spec \paperw / \margl / \margr defaults agree).
/// Read once at open from the table stream: one small record per section, never the text.
/// </summary>
internal static class DocBinarySectionTable
{
    public const int DefaultPageWidthTwips = 12240;
    public const int DefaultMarginTwips = 1800;
    public static readonly int DefaultTextWidthTwips = WordDocumentPageSetup.TextWidth(DefaultPageWidthTwips, DefaultMarginTwips, DefaultMarginTwips, 0);

    private const ushort SprmSXaPage = 0xB01F, SprmSDxaLeft = 0xB021, SprmSDxaRight = 0xB022, SprmSDzaGutter = 0xB025;
    /// <summary>sprmSFRTLGutter: the gutter is on the right (LibreOffice sprmids.hxx:589 SFRTLGutter 0x322A, one byte).</summary>
    private const ushort SprmSFRightToLeftGutter = 0x322A;
    private const int SedSize = 12;

    /// <summary>The start CP of every section (PlcfSed's sorted CPs, clamped to 0) - the index pass records each one's block.</summary>
    public static int[] ReadStartCps(Stream tableStream, (uint offset, uint length) plcfSed)
        => TryReadPlc(tableStream, plcfSed, out byte[] plc, out int count, out _)
            ? Enumerable.Range(0, count).Select(i => Math.Max(0, BitConverter.ToInt32(plc, i * 4))).ToArray()
            : [];

    /// <param name="firstBlocks">The block each section starts in, in the order of <see cref="ReadStartCps"/> (from the index).</param>
    /// <param name="documentProperties">The Dop, for iGutterPos (gutter at the top - <see cref="IsGutterAtTop"/>).</param>
    public static IReadOnlyList<WordDocumentSection> Read(Stream wordDocument, Stream tableStream, (uint offset, uint length) plcfSed,
                                                         IReadOnlyList<int> firstBlocks, (uint offset, uint length) documentProperties = default)
    {
        bool isGutterAtTop = IsGutterAtTop(tableStream, documentProperties);
        var fallback = new[] { new WordDocumentSection(0, DefaultTextWidthTwips) };
        if (!TryReadPlc(tableStream, plcfSed, out byte[] plc, out int count, out int sedsAt)) return fallback;
        if (firstBlocks.Count != count) throw new InvalidDataException(".doc index: the section count does not match the document.");

        var starts = new List<WordDocumentSection>(count);
        for (int i = 0; i < count; i++)
        {
            uint fcSepx = BitConverter.ToUInt32(plc, sedsAt + i * SedSize + 2);
            (int width, int left) = PageSetupOf(wordDocument, fcSepx, isGutterAtTop);
            starts.Add(new WordDocumentSection(firstBlocks[i], width, left));
        }
        return WordDocumentPageSetup.FromSectionStarts(starts);
    }

    // Note numbering sprms (LibreOffice sw/source/filter/ww8/sprmids.hxx:606-611): sprmSRncFtn 0x303C, sprmSRncEdn 0x303E (one
    // byte: 0 continuous, 1 each section, 2 each page), sprmSNFtn 0x503F, sprmSNEdn 0x5041 (start, 2 bytes), sprmSNfcFtnRef 0x5040,
    // sprmSNfcEdnRef 0x5042 (MSONFC, 2 bytes). LibreOffice ignores them (ww8par6.cxx: SRncFtn has no handler) and numbers by the Dop
    // only (ww8par.cxx:5060-5085); Word numbers by each section's own (test-data/word-notes: Word 16 prints them), which we follow.
    private const ushort SprmSRncFtn = 0x303C, SprmSRncEdn = 0x303E, SprmSNFtn = 0x503F, SprmSNfcFtnRef = 0x5040, SprmSNEdn = 0x5041, SprmSNfcEdnRef = 0x5042;

    /// <summary>Each section's footnote / endnote numbering, keyed by the section's start CP (<see cref="WordDocumentNoteNumbering"/>
    /// positions are main-text CPs here). A section without the sprms takes the Sep defaults: decimal / lower roman, from 1, continuous.</summary>
    public static WordDocumentNoteNumbering ReadNoteNumbering(Stream wordDocument, Stream tableStream, (uint offset, uint length) plcfSed)
    {
        if (!TryReadPlc(tableStream, plcfSed, out byte[] plc, out int count, out int sedsAt)) return WordDocumentNoteNumbering.Defaults;
        var starts = new int[count];
        var footnotes = new WordDocumentNoteNumberRule[count];
        var endnotes = new WordDocumentNoteNumberRule[count];
        for (int i = 0; i < count; i++)
        {
            starts[i] = i == 0 ? 0 : Math.Max(starts[i - 1], BitConverter.ToInt32(plc, i * 4));
            WordDocumentNoteNumberRule footnote = WordDocumentNoteNumberRule.FootnoteDefault, endnote = WordDocumentNoteNumberRule.EndnoteDefault;
            uint fcSepx = BitConverter.ToUInt32(plc, sedsAt + i * SedSize + 2);
            if (ReadSepx(wordDocument, fcSepx) is { } grpprl)
                foreach ((ushort opcode, int at, int length) in DocBinaryPropertyModifiers.Sprms(grpprl, 0, grpprl.Length))
                {
                    if (length == 1 && opcode is SprmSRncFtn or SprmSRncEdn)
                    {
                        var restart = grpprl[at] switch { 1 => WordDocumentNoteRestart.EachSection, 2 => WordDocumentNoteRestart.EachPage, _ => WordDocumentNoteRestart.Continuous };
                        if (opcode == SprmSRncFtn) footnote = footnote with { Restart = restart }; else endnote = endnote with { Restart = restart };
                    }
                    else if (length == 2)
                    {
                        ushort value = BitConverter.ToUInt16(grpprl, at);
                        switch (opcode)
                        {
                            case SprmSNFtn: footnote = footnote with { Start = value }; break;
                            case SprmSNEdn: endnote = endnote with { Start = value }; break;
                            case SprmSNfcFtnRef: footnote = footnote with { Format = DocBinaryListDefinitions.NumberFormatName((byte)Math.Min(value, (ushort)255)) }; break;
                            case SprmSNfcEdnRef: endnote = endnote with { Format = DocBinaryListDefinitions.NumberFormatName((byte)Math.Min(value, (ushort)255)) }; break;
                        }
                    }
                }
            footnotes[i] = footnote; endnotes[i] = endnote;
        }
        return new WordDocumentNoteNumbering(starts, footnotes, endnotes);
    }

    private static byte[]? ReadSepx(Stream wordDocument, uint fcSepx)
    {
        if (fcSepx == 0xFFFFFFFF || DocBinaryStructureReader.TryRead(wordDocument, (fcSepx, 2)) is not { } sizeBytes) return null;
        short cb = BitConverter.ToInt16(sizeBytes, 0);
        return cb > 0 ? DocBinaryStructureReader.TryRead(wordDocument, (fcSepx + 2, (uint)cb)) : null;
    }

    private static bool TryReadPlc(Stream tableStream, (uint offset, uint length) plcfSed, out byte[] plc, out int count, out int sedsAt)
    {
        plc = []; count = 0; sedsAt = 0;
        if (plcfSed.length < 4 + 4 + SedSize) return false;
        if (DocBinaryStructureReader.TryRead(tableStream, plcfSed) is not { } bytes) return false;
        plc = bytes;
        int entries = (int)((plcfSed.length - 4) / (4 + SedSize));
        sedsAt = (entries + 1) * 4;                        // the Seds follow ALL CPs, sorted or not
        // CPs out of order end the PLC (LibreOffice TruncToSortedRange, as for the other PLCs).
        count = DocBinaryStructureReader.SortedEntryCount(i => BitConverter.ToInt32(bytes, i * 4), entries);
        return count > 0;
    }

    /// <summary>
    /// Dop iGutterPos: the gutter is at the top of the page - bit 15 of the 16-bit word at offset 0x52 of the Dop (LibreOffice
    /// ww8scan.cxx:7741-7746 WW8Dop; ww8par.cxx:2117-2119 sets GUTTER_AT_TOP from it). Word 97 and later only: the Dop of a
    /// Word 6 / 95 file is shorter and has other fields there.
    /// </summary>
    private static bool IsGutterAtTop(Stream tableStream, (uint offset, uint length) dop)
    {
        if (dop.length < 0x54 || dop.offset + 0x54L > tableStream.Length) return false;
        return DocBinaryStructureReader.TryRead(tableStream, (dop.offset + 0x52, 2)) is { } word && (word[1] & 0x80) != 0;
    }

    /// <summary>The section's text width and where its column starts (left margin + a left gutter - <see cref="WordDocumentPageSetup.GutterPlacement"/>).</summary>
    private static (int width, int columnLeft) PageSetupOf(Stream wordDocument, uint fcSepx, bool isGutterAtTop)
    {
        int page = DefaultPageWidthTwips, left = DefaultMarginTwips, right = DefaultMarginTwips, gutter = 0;
        bool isRightToLeftGutter = false;
        if (fcSepx != 0xFFFFFFFF && DocBinaryStructureReader.TryRead(wordDocument, (fcSepx, 2)) is { } sizeBytes)
        {
            short cb = BitConverter.ToInt16(sizeBytes, 0);
            if (cb > 0 && DocBinaryStructureReader.TryRead(wordDocument, (fcSepx + 2, (uint)cb)) is { } grpprl)
            {
                foreach ((ushort opcode, int at, int length) in DocBinaryPropertyModifiers.Sprms(grpprl, 0, grpprl.Length))
                {
                    if (opcode == SprmSFRightToLeftGutter && length == 1) { isRightToLeftGutter = grpprl[at] != 0; continue; }
                    if (length != 2) continue;
                    int value = opcode == SprmSXaPage ? BitConverter.ToUInt16(grpprl, at) : BitConverter.ToInt16(grpprl, at);
                    switch (opcode)
                    {
                        case SprmSXaPage: page = value; break;
                        case SprmSDxaLeft: left = value; break;
                        case SprmSDxaRight: right = value; break;
                        case SprmSDzaGutter: gutter = value; break;
                    }
                }
            }
        }
        var placement = WordDocumentPageSetup.GutterPlacement(gutter, isRightToLeftGutter, isGutterAtTop);
        return (WordDocumentPageSetup.TextWidth(page, left, right, placement.WidthGutter), WordDocumentPageSetup.ColumnLeft(left, placement.LeftGutter));
    }
}
