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
    private const int SedSize = 12;

    public static IReadOnlyList<WordDocumentSection> Read(Stream wordDocument, Stream tableStream, (uint offset, uint length) plcfSed,
                                                         Func<int, int> blockAtCp)
    {
        var fallback = new[] { new WordDocumentSection(0, DefaultTextWidthTwips) };
        if (plcfSed.length < 4 + 4 + SedSize) return fallback;
        if (DocBinaryStructureReader.TryRead(tableStream, plcfSed) is not { } plc) return fallback;
        int entries = (int)((plcfSed.length - 4) / (4 + SedSize));
        int sedsAt = (entries + 1) * 4;                    // the Seds follow ALL CPs, sorted or not
        // CPs out of order end the PLC (LibreOffice TruncToSortedRange, as for the other PLCs).
        int count = DocBinaryStructureReader.SortedEntryCount(i => BitConverter.ToInt32(plc, i * 4), entries);
        if (count == 0) return fallback;

        var starts = new List<WordDocumentSection>(count);
        for (int i = 0; i < count; i++)
        {
            int startCp = BitConverter.ToInt32(plc, i * 4);
            uint fcSepx = BitConverter.ToUInt32(plc, sedsAt + i * SedSize + 2);
            starts.Add(new WordDocumentSection(blockAtCp(Math.Max(0, startCp)), TextWidthOf(wordDocument, fcSepx)));
        }
        return WordDocumentPageSetup.FromSectionStarts(starts);
    }

    private static int TextWidthOf(Stream wordDocument, uint fcSepx)
    {
        int page = DefaultPageWidthTwips, left = DefaultMarginTwips, right = DefaultMarginTwips, gutter = 0;
        if (fcSepx != 0xFFFFFFFF && DocBinaryStructureReader.TryRead(wordDocument, (fcSepx, 2)) is { } sizeBytes)
        {
            short cb = BitConverter.ToInt16(sizeBytes, 0);
            if (cb > 0 && DocBinaryStructureReader.TryRead(wordDocument, (fcSepx + 2, (uint)cb)) is { } grpprl)
            {
                foreach ((ushort opcode, int at, int length) in DocBinaryPropertyModifiers.Sprms(grpprl, 0, grpprl.Length))
                {
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
        return WordDocumentPageSetup.TextWidth(page, left, right, gutter);
    }
}
