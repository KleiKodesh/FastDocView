namespace FastDocView.DocBinary;

/// <summary>
/// Looks up paragraph (PAPX) and character (CHPX) properties by file offset through the bin tables
/// (PlcfBtePapx / PlcfBteChpx, [MS-DOC] §2.8.5/§2.8.6) and the 512-byte Formatted Disk Pages they
/// point at (§2.9.179 PapxFkp, §2.9.33 ChpxFkp). Pages are cached, bounded, since consecutive
/// paragraphs share them.
/// </summary>
public sealed class DocBinaryFormattedDiskPages
{
    private const int PageSize = 512;
    private const int CachedPageLimit = 512;
    private const int ParagraphBxSize = 13;     // BxPap: bOffset + 12-byte PHE
    private const int Word6ParagraphBxSize = 7; // Word 6/95: bOffset + 6-byte PHE (ww8scan.cxx WW8PLCFx_Fc_FKP::NewFkp: nItemSize 7 for bVer67)
    private const int CharacterBxSize = 1;      // ChpxFkp rgb: one offset byte

    public readonly record struct PropertyRun(uint FcStart, uint FcEnd, byte[] Page, int GrpprlStart, int GrpprlLength);

    private readonly Stream _wordDocument;
    private readonly uint[] _paragraphBinFcs;
    private readonly uint[] _paragraphBinPages;
    private readonly uint[] _characterBinFcs;
    private readonly uint[] _characterBinPages;
    private readonly Dictionary<uint, LinkedListNode<(uint pageNumber, byte[] page)>> _pageCache = new();
    private readonly LinkedList<(uint pageNumber, byte[] page)> _pageUse = new();     // most recently used first

    private readonly Stream? _dataStream;
    private readonly DocBinaryWord6Sprms? _word6Sprms;
    private readonly int _paragraphBxSize = ParagraphBxSize;
    /// <summary>Word 6/95: each cached page's grpprls translated to Word 97 ones, by offset; dropped with the page.</summary>
    private readonly Dictionary<byte[], Dictionary<(int, int), byte[]>> _translated = new(PageIdentity.Instance);

    /// <param name="dataStream">The compound file's "Data" stream, where oversized PAPXs live (sprmPHugePapx); null when absent.</param>
    public DocBinaryFormattedDiskPages(Stream wordDocument, Stream tableStream, DocBinaryFileInformationBlock fib, Stream? dataStream = null)
    {
        _wordDocument = wordDocument;
        _dataStream = dataStream;
        if (fib.IsWord6)
        {
            _word6Sprms = new DocBinaryWord6Sprms(fib.IsWord6FontCodeVariant);
            _paragraphBxSize = Word6ParagraphBxSize;
            (_paragraphBinFcs, _paragraphBinPages) = ReadWord6BinTable(tableStream, fib.ParagraphBinTable, fib.ParagraphFirstPage, fib.ParagraphBinPageCount);
            (_characterBinFcs, _characterBinPages) = ReadWord6BinTable(tableStream, fib.CharacterBinTable, fib.CharacterFirstPage, fib.CharacterBinPageCount);
            return;
        }
        (_paragraphBinFcs, _paragraphBinPages) = ReadBinTable(tableStream, fib.ParagraphBinTable);
        (_characterBinFcs, _characterBinPages) = ReadBinTable(tableStream, fib.CharacterBinTable);
    }

    /// <summary>
    /// A Word 6/95 bin table: 2-byte page numbers (ww8scan.cxx WW8PLCFx_Fc_FKP: nLenStruct 2 when nVersion &lt; 8). Word 6 may
    /// list fewer pages than the FIB's cpnBte; then the table is rebuilt from the FKPs themselves, consecutive from pnFirst
    /// (WW8PLCF ctor: GeneratePLCF when the PLCF has fewer than ncpN entries - each page's first FC, the last page's last).
    /// </summary>
    private (uint[] fcs, uint[] pages) ReadWord6BinTable(Stream tableStream, (uint offset, uint length) location, int firstPage, int pageCount)
    {
        int listed = location.length >= 4 ? (int)((location.length - 4) / 6) : 0;
        if (listed >= pageCount && listed > 0 && DocBinaryStructureReader.TryRead(tableStream, location) is { } bytes)
        {
            int count = DocBinaryStructureReader.SortedEntryCount(i => BitConverter.ToUInt32(bytes, i * 4), listed);
            int pageNumbersAt = (listed + 1) * 4;
            var fcs = new uint[count + 1];
            var pages = new uint[count];
            for (int i = 0; i <= count; i++) fcs[i] = BitConverter.ToUInt32(bytes, i * 4);
            for (int i = 0; i < count; i++) pages[i] = BitConverter.ToUInt16(bytes, pageNumbersAt + i * 2);
            return (fcs, pages);
        }
        if (pageCount < 1 || firstPage < 0 || firstPage + pageCount > ushort.MaxValue) return ([], []);
        var generatedFcs = new uint[pageCount + 1];
        var generatedPages = new uint[pageCount];
        var page = new byte[PageSize];
        for (int i = 0; i < pageCount; i++)
        {
            if (!ReadPage((uint)(firstPage + i), page)) return ([], []);
            generatedPages[i] = (uint)(firstPage + i);
            generatedFcs[i] = BitConverter.ToUInt32(page, 0);
            if (i == pageCount - 1) generatedFcs[pageCount] = BitConverter.ToUInt32(page, Math.Min((int)page[PageSize - 1], (PageSize - 1) / 4 - 1) * 4);
        }
        int sorted = DocBinaryStructureReader.SortedEntryCount(i => generatedFcs[i], pageCount);
        Array.Resize(ref generatedFcs, sorted + 1);
        Array.Resize(ref generatedPages, sorted);
        return (generatedFcs, generatedPages);
    }

    /// <summary>Pages are told apart by identity (net48 has no ReferenceEqualityComparer).</summary>
    private sealed class PageIdentity : IEqualityComparer<byte[]>
    {
        public static readonly PageIdentity Instance = new();
        public bool Equals(byte[]? x, byte[]? y) => ReferenceEquals(x, y);
        public int GetHashCode(byte[] page) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(page);
    }

    private bool ReadPage(uint pageNumber, byte[] page)
    {
        long at = (long)pageNumber * PageSize;
        if (at + PageSize > _wordDocument.Length) return false;
        _wordDocument.Position = at;
        int read = 0;
        while (read < PageSize) { int n = _wordDocument.Read(page, read, PageSize - read); if (n <= 0) break; read += n; }
        return read == PageSize;
    }

    /// <summary>
    /// A Word 6/95 PAPX / CHPX of an FKP page as a Word 97 one: <paramref name="prefixLength"/> bytes at
    /// <paramref name="prefixAt"/> (the PAPX istd) kept, then the grpprl translated. Cached with the page.
    /// </summary>
    private PropertyRun Translated(uint fcStart, uint fcEnd, byte[] page, int prefixAt, int prefixLength, int grpprlStart, int length)
    {
        if (!_translated.TryGetValue(page, out Dictionary<(int, int), byte[]>? ofPage)) _translated[page] = ofPage = new();
        if (!ofPage.TryGetValue((prefixAt, prefixLength), out byte[]? translated))
        {
            byte[] word97 = _word6Sprms!.ToWord97(page, grpprlStart, length);
            translated = new byte[prefixLength + word97.Length];
            Buffer.BlockCopy(page, prefixAt, translated, 0, prefixLength);
            Buffer.BlockCopy(word97, 0, translated, prefixLength, word97.Length);
            ofPage[(prefixAt, prefixLength)] = translated;
        }
        return new PropertyRun(fcStart, fcEnd, translated, 0, translated.Length);
    }

    /// <summary>The PAPX run containing <paramref name="fc"/> (a paragraph's mark); grpprl starts with istd.</summary>
    public PropertyRun? ParagraphRunAt(uint fc)
    {
        byte[]? page = PageFor(_paragraphBinFcs, _paragraphBinPages, fc);
        if (page is null) return null;
        int runCount = UsableRunCount(page, _paragraphBxSize);
        int run = RunIndex(page, runCount, fc);
        if (run < 0) return null;
        int offset = page[(runCount + 1) * 4 + run * _paragraphBxSize] * 2;
        uint fcStart = BitConverter.ToUInt32(page, run * 4), fcEnd = BitConverter.ToUInt32(page, (run + 1) * 4);
        if (offset == 0 || offset >= PageSize - 1) return new PropertyRun(fcStart, fcEnd, page, 0, 0);   // no PAPX, or one pointing into the crun byte
        if (_word6Sprms is not null)
        {
            // Word 6/95 PAPX: one length byte (in words; never the Word 97 second length byte), the istd, the grpprl
            // (ww8scan.cxx WW8Fkp ctor, case PAP: the extra length byte only IsEightPlus)
            int papxLength = Math.Min(page[offset] * 2, PageSize - 1 - (offset + 1));
            if (papxLength < 2) return new PropertyRun(fcStart, fcEnd, page, 0, 0);
            return Translated(fcStart, fcEnd, page, offset + 1, 2, offset + 3, papxLength - 2);
        }
        int length, grpprlStart;
        if (page[offset] == 0) { length = page[offset + 1] * 2; grpprlStart = offset + 2; }
        else { length = page[offset] * 2 - 1; grpprlStart = offset + 1; }
        return new PropertyRun(fcStart, fcEnd, page, grpprlStart, Math.Max(0, Math.Min(length, PageSize - 1 - grpprlStart)));
    }

    /// <summary>The CHPX run containing <paramref name="fc"/>.</summary>
    public PropertyRun? CharacterRunAt(uint fc)
    {
        byte[]? page = PageFor(_characterBinFcs, _characterBinPages, fc);
        if (page is null) return null;
        int runCount = UsableRunCount(page, CharacterBxSize);
        int run = RunIndex(page, runCount, fc);
        if (run < 0) return null;
        int offset = page[(runCount + 1) * 4 + run * CharacterBxSize] * 2;
        uint fcStart = BitConverter.ToUInt32(page, run * 4), fcEnd = BitConverter.ToUInt32(page, (run + 1) * 4);
        if (offset == 0 || offset >= PageSize - 1) return new PropertyRun(fcStart, fcEnd, page, 0, 0);
        int length = Math.Min((int)page[offset], PageSize - 1 - (offset + 1));
        if (_word6Sprms is not null) return Translated(fcStart, fcEnd, page, offset + 1, 0, offset + 1, length);
        return new PropertyRun(fcStart, fcEnd, page, offset + 1, length);
    }

    /// <summary>
    /// A PAPX grpprl with sprmPHugePapx (0x6646) resolved: that sprm's operand is an offset into
    /// the Data stream holding cb (2 bytes) + the real grpprl ([MS-DOC] §2.6.2). Returns the input
    /// span when there is none — the common case, no allocation.
    /// </summary>
    public (byte[] grpprl, int start, int length) ExpandHugePapx(byte[] grpprl, int start, int length)
    {
        if (_dataStream is null) return (grpprl, start, length);
        foreach ((ushort opcode, int operandAt, _) in DocBinaryPropertyModifiers.Sprms(grpprl, start, length))
        {
            if (opcode != 0x6646) continue;
            uint offset = BitConverter.ToUInt32(grpprl, operandAt);
            if (offset + 2 > _dataStream.Length) break;
            var header = new byte[2];
            _dataStream.Position = offset;
            if (_dataStream.Read(header, 0, 2) < 2) break;
            int size = BitConverter.ToUInt16(header, 0);
            var expanded = new byte[size];
            int read = 0;
            while (read < size) { int n = _dataStream.Read(expanded, read, size - read); if (n <= 0) break; read += n; }
            return (expanded, 0, read);
        }
        return (grpprl, start, length);
    }

    /// <summary>Every PAPX (isParagraph) or CHPX run in the document, page by page — for audits, not for rendering.</summary>
    public IEnumerable<PropertyRun> AllRuns(bool isParagraph)
    {
        uint[] binFcs = isParagraph ? _paragraphBinFcs : _characterBinFcs;
        uint[] binPages = isParagraph ? _paragraphBinPages : _characterBinPages;
        for (int bin = 0; bin < binPages.Length; bin++)
        {
            byte[]? page = PageFor(binFcs, binPages, binFcs[bin]);
            if (page is null) continue;
            int runCount = UsableRunCount(page, isParagraph ? _paragraphBxSize : CharacterBxSize);
            for (int run = 0; run < runCount; run++)
            {
                uint fc = BitConverter.ToUInt32(page, run * 4);
                PropertyRun? found = isParagraph ? ParagraphRunAt(fc) : CharacterRunAt(fc);
                if (found is { } propertyRun) yield return propertyRun;
            }
        }
    }

    /// <summary>
    /// crun (the page's last byte) limited to the runs whose rgfc and rgbx entries fit before that
    /// byte: (crun + 1) 4-byte FCs then crun BX entries. LibreOffice does the same — its WW8Fkp
    /// constructor (ww8scan.cxx) lowers mnIMax when an entry would fall past byte 511. For an intact
    /// page this changes nothing ([MS-DOC] PapxFkp / ChpxFkp: at most 0x1D / 0x65 runs, which is
    /// exactly what fits).
    /// </summary>
    private static int UsableRunCount(byte[] page, int bxSize)
    {
        int runCount = page[PageSize - 1];
        int fitting = (PageSize - 1 - 4) / (4 + bxSize);
        return Math.Min(runCount, fitting);
    }

    private static int RunIndex(byte[] page, int runCount, uint fc)
    {
        for (int run = 0; run < runCount; run++)
            if (BitConverter.ToUInt32(page, run * 4) <= fc && fc < BitConverter.ToUInt32(page, (run + 1) * 4)) return run;
        return runCount > 0 && fc >= BitConverter.ToUInt32(page, runCount * 4) ? runCount - 1 : -1;
    }

    private byte[]? PageFor(uint[] binFcs, uint[] binPages, uint fc)
    {
        if (binPages.Length == 0) return null;
        int low = 0, high = binPages.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (binFcs[middle] <= fc) low = middle; else high = middle - 1;
        }
        uint pageNumber = binPages[low] & 0x3FFFFF;
        if (_pageCache.TryGetValue(pageNumber, out LinkedListNode<(uint pageNumber, byte[] page)>? cached))
        {
            _pageUse.Remove(cached);                       // least recently used last: the one a full cache evicts
            _pageUse.AddFirst(cached);
            return cached.Value.page;
        }
        if (_pageCache.Count >= CachedPageLimit)
        {
            // only the least recently used page goes (its array is not reused: a caller may still be reading it)
            _translated.Remove(_pageUse.Last!.Value.page);
            _pageCache.Remove(_pageUse.Last!.Value.pageNumber);
            _pageUse.RemoveLast();
        }
        var page = new byte[PageSize];
        _wordDocument.Position = (long)pageNumber * PageSize;
        int read = 0;
        while (read < PageSize) { int n = _wordDocument.Read(page, read, PageSize - read); if (n <= 0) break; read += n; }
        _pageCache[pageNumber] = _pageUse.AddFirst((pageNumber, page));
        return page;
    }

    private static (uint[] fcs, uint[] pages) ReadBinTable(Stream tableStream, (uint offset, uint length) location)
    {
        if (location.length < 8) return ([], []);
        byte[]? bytes = DocBinaryStructureReader.TryRead(tableStream, location);
        if (bytes is null) return ([], []);                // outside the table stream: no bin table (LibreOffice: an empty PLCF)
        int count = DocBinaryStructureReader.SortedEntryCount(i => BitConverter.ToUInt32(bytes, i * 4), (int)((location.length - 4) / 8));
        int pageNumbersAt = (int)((location.length - 4) / 8 + 1) * 4;
        var fcs = new uint[count + 1];
        var pages = new uint[count];
        for (int i = 0; i <= count; i++) fcs[i] = BitConverter.ToUInt32(bytes, i * 4);
        for (int i = 0; i < count; i++) pages[i] = BitConverter.ToUInt32(bytes, pageNumbersAt + i * 4);
        return (fcs, pages);
    }
}
