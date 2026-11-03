namespace FastDocView.DocBinary;

/// <summary>
/// The search's one pass over a .doc's main text: the piece characters are read once, in 64K-character chunks, and
/// serve both the block boundaries (the same <see cref="DocBinaryBlockScanner"/> the index pass ran, from a block's
/// start) and the matcher (each block's characters handed over as soon as its end is known). The characters read past
/// the last block handed out stay buffered, so the next batch of the same search continues without re-reading them.
/// A buffer holds the text of the block being collected plus the rest of the current chunk - never more than that.
/// </summary>
internal sealed class DocBinarySearchCursor
{
    private const int ChunkCharacters = 1 << 16;

    private readonly DocBinaryPieceTable _pieces;
    private readonly DocBinaryBlockScanner _scanner;
    private readonly int _mainTextCharacterCount;
    private readonly int _blockCount;
    private readonly int _firstBlock;              // the global index of the scanner's block 0
    private int _scannedTo;
    private bool _isFinished;

    // Buffered characters in CP order. A damaged piece table can name CPs whose bytes lie past the stream's end: those
    // CPs have no character (as in DocBinaryPieceTable.ReadText), so the buffer keeps runs of (first CP, buffer offset, length).
    private char[] _buffer = new char[ChunkCharacters];
    private int _bufferLength;
    private readonly List<(int cp, int offset, int length)> _runs = new();

    /// <summary>The next block <see cref="TryNext"/> hands out.</summary>
    public int NextBlock => _firstBlock + _emitted;
    private int _emitted;

    public DocBinarySearchCursor(DocBinaryPieceTable pieces, DocBinaryParagraphHtmlRenderer renderer, DocBinaryStyleSheet styles,
                                 DocBinaryListDefinitions lists, int firstBlock, int firstBlockStartCp, int blockCount, int mainTextCharacterCount)
    {
        _pieces = pieces;
        _scanner = new DocBinaryBlockScanner(renderer, styles, lists, firstBlockStartCp) { CollectsHeadings = false };
        _firstBlock = firstBlock;
        _scannedTo = firstBlockStartCp;
        _blockCount = blockCount;
        _mainTextCharacterCount = mainTextCharacterCount;
    }

    /// <summary>The next block's CP range and characters; false past the last block. The span is valid until the next call.</summary>
    public bool TryNext(out int startCp, out int endCp, out ReadOnlySpan<char> text)
    {
        startCp = endCp = 0;
        text = default;
        if (NextBlock >= _blockCount) return false;
        // a block ends where the next one starts; the last one at the main text's end, which the scan must reach
        bool isLast = NextBlock + 1 >= _blockCount;
        while (!_isFinished && (isLast || _scanner.BlockCount <= _emitted + 1)) ScanMore();
        if (isLast ? _scanner.BlockCount != _emitted + 1 : _scanner.BlockCount <= _emitted + 1)
            throw new InvalidDataException($".doc search: the scan from block {_firstBlock} found {_scanner.BlockCount} blocks, the index {_blockCount - _firstBlock} (the document changed?).");
        startCp = _scanner.BlockStartCps[_emitted];
        endCp = isLast ? _mainTextCharacterCount : _scanner.BlockStartCps[_emitted + 1];
        int from = BufferIndexOf(startCp), to = BufferIndexOf(endCp);
        text = new ReadOnlySpan<char>(_buffer, from, to - from);
        _emitted++;
        return true;
    }

    /// <summary>After a batch: a buffer a huge block grew past 4 MB drops back to what it still holds (the cursor lives with the document).</summary>
    public void ReleaseLargeBuffer()
    {
        const int NormalBufferCharacters = 2 << 20;
        if (_buffer.Length <= NormalBufferCharacters) return;
        DropBefore(_scanner.BlockCount > _emitted ? _scanner.BlockStartCps[_emitted] : _scanner.ScannedEndCp);
        Array.Resize(ref _buffer, Math.Max(ChunkCharacters, _bufferLength));
    }

    private void ScanMore()
    {
        if (_scannedTo >= _mainTextCharacterCount)
        {
            _scanner.Finish(_mainTextCharacterCount);
            _isFinished = true;
            return;
        }
        DropBefore(_scanner.BlockCount > _emitted ? _scanner.BlockStartCps[_emitted] : _scanner.ScannedEndCp);
        int to = (int)Math.Min((long)_scannedTo + ChunkCharacters, _mainTextCharacterCount);
        _pieces.ScanCharacters(_scannedTo, to, (characters, count, chunkCp) =>
        {
            Append(characters, count, chunkCp);
            _scanner.Feed(characters, count, chunkCp);
        }, ChunkCharacters);
        _scannedTo = to;
    }

    private void Append(char[] characters, int count, int chunkCp)
    {
        if (count <= 0) return;
        if (_bufferLength + count > _buffer.Length) Array.Resize(ref _buffer, Math.Max(_buffer.Length * 2, _bufferLength + count));
        Array.Copy(characters, 0, _buffer, _bufferLength, count);
        if (_runs.Count > 0 && _runs[_runs.Count - 1] is var last && last.cp + last.length == chunkCp)
            _runs[_runs.Count - 1] = (last.cp, last.offset, last.length + count);
        else _runs.Add((chunkCp, _bufferLength, count));
        _bufferLength += count;
    }

    /// <summary>The buffer offset of a CP: the number of buffered characters before it.</summary>
    private int BufferIndexOf(int cp)
    {
        int index = 0;
        foreach ((int runCp, int offset, int length) in _runs)
        {
            if (cp < runCp) return offset;
            index = offset + Math.Min(length, cp - runCp);
        }
        return _runs.Count == 0 ? 0 : index;
    }

    /// <summary>Forgets the characters before a CP (blocks already handed out), keeping the rest at the buffer's start.</summary>
    private void DropBefore(int cp)
    {
        int drop = BufferIndexOf(cp);
        if (drop == 0) return;
        Array.Copy(_buffer, drop, _buffer, 0, _bufferLength - drop);
        _bufferLength -= drop;
        for (int i = _runs.Count - 1; i >= 0; i--)
        {
            (int runCp, int offset, int length) = _runs[i];
            if (offset + length <= drop) { _runs.RemoveRange(0, i + 1); break; }
            int cut = Math.Max(0, drop - offset);
            _runs[i] = (runCp + cut, offset + cut - drop, length - cut);
        }
    }
}
