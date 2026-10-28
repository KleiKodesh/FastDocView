namespace WordDocumentRandomAccess.DocBinary;

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

    public readonly record struct PropertyRun(uint FcStart, uint FcEnd, byte[] Page, int GrpprlStart, int GrpprlLength);

    private readonly Stream _wordDocument;
    private readonly uint[] _paragraphBinFcs;
    private readonly uint[] _paragraphBinPages;
    private readonly uint[] _characterBinFcs;
    private readonly uint[] _characterBinPages;
    private readonly Dictionary<uint, byte[]> _pageCache = new();

    public DocBinaryFormattedDiskPages(Stream wordDocument, Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        _wordDocument = wordDocument;
        (_paragraphBinFcs, _paragraphBinPages) = ReadBinTable(tableStream, fib.ParagraphBinTable);
        (_characterBinFcs, _characterBinPages) = ReadBinTable(tableStream, fib.CharacterBinTable);
    }

    /// <summary>The PAPX run containing <paramref name="fc"/> (a paragraph's mark); grpprl starts with istd.</summary>
    public PropertyRun? ParagraphRunAt(uint fc)
    {
        byte[]? page = PageFor(_paragraphBinFcs, _paragraphBinPages, fc);
        if (page is null) return null;
        int runCount = page[PageSize - 1];
        int run = RunIndex(page, runCount, fc);
        if (run < 0) return null;
        int offset = page[(runCount + 1) * 4 + run * 13] * 2;
        uint fcStart = BitConverter.ToUInt32(page, run * 4), fcEnd = BitConverter.ToUInt32(page, (run + 1) * 4);
        if (offset == 0) return new PropertyRun(fcStart, fcEnd, page, 0, 0);
        int length, grpprlStart;
        if (page[offset] == 0) { length = page[offset + 1] * 2; grpprlStart = offset + 2; }
        else { length = page[offset] * 2 - 1; grpprlStart = offset + 1; }
        return new PropertyRun(fcStart, fcEnd, page, grpprlStart, Math.Min(length, PageSize - grpprlStart));
    }

    /// <summary>The CHPX run containing <paramref name="fc"/>.</summary>
    public PropertyRun? CharacterRunAt(uint fc)
    {
        byte[]? page = PageFor(_characterBinFcs, _characterBinPages, fc);
        if (page is null) return null;
        int runCount = page[PageSize - 1];
        int run = RunIndex(page, runCount, fc);
        if (run < 0) return null;
        int offset = page[(runCount + 1) * 4 + run] * 2;
        uint fcStart = BitConverter.ToUInt32(page, run * 4), fcEnd = BitConverter.ToUInt32(page, (run + 1) * 4);
        if (offset == 0) return new PropertyRun(fcStart, fcEnd, page, 0, 0);
        int length = page[offset];
        return new PropertyRun(fcStart, fcEnd, page, offset + 1, Math.Min(length, PageSize - offset - 1));
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
        if (_pageCache.TryGetValue(pageNumber, out byte[]? cached)) return cached;
        if (_pageCache.Count >= CachedPageLimit) _pageCache.Clear();
        var page = new byte[PageSize];
        _wordDocument.Position = (long)pageNumber * PageSize;
        int read = 0;
        while (read < PageSize) { int n = _wordDocument.Read(page, read, PageSize - read); if (n <= 0) break; read += n; }
        _pageCache[pageNumber] = page;
        return page;
    }

    private static (uint[] fcs, uint[] pages) ReadBinTable(Stream tableStream, (uint offset, uint length) location)
    {
        if (location.length < 8) return ([], []);
        var bytes = new byte[location.length];
        tableStream.Position = location.offset;
        int read = 0;
        while (read < bytes.Length) { int n = tableStream.Read(bytes, read, bytes.Length - read); if (n <= 0) break; read += n; }
        int count = (int)((location.length - 4) / 8);
        var fcs = new uint[count + 1];
        var pages = new uint[count];
        for (int i = 0; i <= count; i++) fcs[i] = BitConverter.ToUInt32(bytes, i * 4);
        for (int i = 0; i < count; i++) pages[i] = BitConverter.ToUInt32(bytes, (count + 1) * 4 + i * 4);
        return (fcs, pages);
    }
}
