namespace FastDocView.Docx;

/// <summary>
/// The search's read path: a search batch needs only the raw bytes of its blocks, and the search scanner finds the block
/// boundaries itself (<see cref="DocxSearchTextScanner.FeedBlocks"/>), so no span scan runs - the anchors give the
/// offsets at both ends. The bytes are read straight from the view in one sequential pass; the tail read past a batch's
/// last block (when the batch does not end at an anchor) is kept for the next batch, which starts there.
/// </summary>
internal sealed partial class DocxCompactBlockTable
{
    private byte[] _searchBuffer = [];
    private long _searchBufferStart = -1;
    private int _searchBufferLength;
    private int _searchResumeBlock = -1;      // a block start the last batch found (where the next batch starts)
    private long _searchResumeOffset;

    /// <summary>
    /// The bytes a search over blocks [<paramref name="fromBlock"/>, <paramref name="endBlock"/>) reads: from the first
    /// block's start to the end block's start when an anchor or the body end gives it, else on to the next anchor
    /// (<paramref name="isEndExact"/> false: the scan stops at the end block's start and reports it with
    /// <see cref="RememberSearchResume"/>).
    /// </summary>
    internal (byte[] buffer, int start, int length) ReadSearchBytes(int fromBlock, int endBlock, out long startOffset, out bool isEndExact)
    {
        int fromAnchor = _index.AnchorIndexOf(fromBlock);
        if (Anchors[fromAnchor].FirstBlock == fromBlock) startOffset = Anchors[fromAnchor].Offset;
        else if (_searchResumeBlock == fromBlock) startOffset = _searchResumeOffset;
        else startOffset = BlockStartOffset(fromBlock);
        long endOffset;
        if (endBlock >= BlockCount) { endOffset = _index.BodyEndOffset; isEndExact = true; }
        else
        {
            int endAnchor = _index.AnchorIndexOf(endBlock);
            if (Anchors[endAnchor].FirstBlock == endBlock) { endOffset = Anchors[endAnchor].Offset; isEndExact = true; }
            else { endOffset = endAnchor + 1 < Anchors.Length ? Anchors[endAnchor + 1].Offset : _index.BodyEndOffset; isEndExact = false; }
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
            var bigger = new byte[Math.Max(length, Math.Min(_searchBuffer.Length * 2, WordSearchPiece.NormalBufferBytes))];
            Buffer.BlockCopy(_searchBuffer, 0, bigger, 0, kept);
            _searchBuffer = bigger;
        }
        long readFrom = startOffset + kept;
        if (kept < length && _view.Position != readFrom) _view.Position = readFrom;
        int filled = kept;
        while (filled < length)
        {
            int read = _view.Read(_searchBuffer, filled, length - filled);
            if (read <= 0) throw new EndOfStreamException("Part ended inside a search range.");
            filled += read;
        }
        _searchBufferStart = startOffset;
        _searchBufferLength = length;
        return (_searchBuffer, 0, length);
    }

    /// <summary>After a batch: a buffer that grew past the normal cap for an oversize piece drops back (the kept tail too when it does not fit).</summary>
    internal void ShrinkSearchBuffer()
    {
        if (!WordSearchPiece.Shrink(ref _searchBuffer, _searchBufferLength)) { _searchBufferStart = -1; _searchBufferLength = 0; }
    }

    /// <summary>Where block <paramref name="blockIndex"/> starts, found by a search scan: the next batch starts there without a span scan.</summary>
    internal void RememberSearchResume(int blockIndex, long offset)
    {
        _searchResumeBlock = blockIndex;
        _searchResumeOffset = offset;
    }
}
