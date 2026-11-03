namespace FastDocView.Odt;

/// <summary>
/// The search's read path: a search batch needs only the raw bytes of its blocks, and the search scanner finds the block
/// boundaries itself (<see cref="OdtSearchTextScanner.FeedBlocks"/>), so no span scan runs - the anchors give the start
/// offsets and the frames open there. The bytes are read straight from the view in one sequential pass; the tail read
/// past a batch's last block (when the batch does not end at an anchor) is kept for the next batch, which starts there.
/// </summary>
internal sealed partial class OdtCompactBlockTable
{
    private byte[] _searchBuffer = [];
    private long _searchBufferStart = -1;
    private int _searchBufferLength;
    private int _searchResumeBlock = -1;      // a block start the last batch found (where the next batch starts)
    private long _searchResumeOffset;
    private OdtSearchTextScanner.Frame[] _searchResumeFrames = [];

    /// <summary>
    /// Where a byte search can start at <paramref name="fromBlock"/>: an anchor's block (its scan state gives the frames)
    /// or the block the last batch stopped at; null otherwise (the caller reads per-block ranges up to
    /// <see cref="NextSearchStart"/>).
    /// </summary>
    internal OdtSearchTextScanner.Frame[]? SearchStart(int fromBlock, out long startOffset)
    {
        int anchor = _index.AnchorIndexOf(fromBlock);
        if (Anchors[anchor].FirstBlock == fromBlock)
        {
            startOffset = Anchors[anchor].Position;
            return OdtBodyBlockScanner.SearchFramesOf(_index.ScanStateOf(anchor));
        }
        if (_searchResumeBlock == fromBlock) { startOffset = _searchResumeOffset; return _searchResumeFrames; }
        startOffset = -1;
        return null;
    }

    /// <summary>The first anchor block after <paramref name="fromBlock"/> (the block count when none).</summary>
    internal int NextSearchStart(int fromBlock)
    {
        int next = _index.AnchorIndexOf(fromBlock) + 1;
        return next < Anchors.Length ? Anchors[next].FirstBlock : BlockCount;
    }

    /// <summary>
    /// The bytes a search over blocks [from, <paramref name="endBlock"/>) reads, from <paramref name="startOffset"/>: to the end
    /// block's start when an anchor gives it, to the last block's end for the last blocks, else on to the next anchor
    /// (<paramref name="isEndExact"/> false: the scan stops at the end block's start and reports it with
    /// <see cref="RememberSearchResume"/>).
    /// </summary>
    internal (byte[] buffer, int start, int length) ReadSearchBytes(long startOffset, int endBlock, out bool isEndExact)
    {
        long endOffset;
        if (endBlock >= BlockCount) { endOffset = _index.LastBlockEndOffset; isEndExact = true; }
        else
        {
            int endAnchor = _index.AnchorIndexOf(endBlock);
            if (Anchors[endAnchor].FirstBlock == endBlock) { endOffset = Anchors[endAnchor].Position; isEndExact = true; }
            else { endOffset = endAnchor + 1 < Anchors.Length ? Anchors[endAnchor + 1].Position : _index.LastBlockEndOffset; isEndExact = false; }
        }
        int length = checked((int)(endOffset - startOffset));
        // keep what the last read already holds from startOffset on
        int kept = 0;
        if (_searchBufferStart >= 0 && startOffset >= _searchBufferStart && startOffset < _searchBufferStart + _searchBufferLength)
        {
            int from = (int)(startOffset - _searchBufferStart);
            kept = Math.Min(_searchBufferLength - from, length);
            if (from > 0) Buffer.BlockCopy(_searchBuffer, from, _searchBuffer, 0, kept);
        }
        if (_searchBuffer.Length < length)
        {
            var bigger = new byte[Math.Max(length, _searchBuffer.Length * 2)];
            Buffer.BlockCopy(_searchBuffer, 0, bigger, 0, kept);
            _searchBuffer = bigger;
        }
        long readFrom = startOffset + kept;
        if (kept < length && _view.Position != readFrom) _view.Position = readFrom;
        int filled = kept;
        while (filled < length)
        {
            int read = _view.Read(_searchBuffer, filled, length - filled);
            if (read <= 0) throw new EndOfStreamException("OpenDocument body ended inside a search range.");
            filled += read;
        }
        _searchBufferStart = startOffset;
        _searchBufferLength = length;
        return (_searchBuffer, 0, length);
    }

    /// <summary>Where block <paramref name="blockIndex"/> starts and the frames open there, found by a search scan: the next batch starts there without a span scan.</summary>
    internal void RememberSearchResume(int blockIndex, long offset, OdtSearchTextScanner.Frame[] frames)
    {
        _searchResumeBlock = blockIndex;
        _searchResumeOffset = offset;
        _searchResumeFrames = frames;
    }
}
