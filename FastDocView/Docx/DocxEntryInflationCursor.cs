using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// The pull form of <see cref="DocxEntryInflation.Run"/>: each <see cref="ReadNextChunk"/> inflates
/// the next chunk of the entry in the caller's thread and returns its length (0 at the end).
/// <see cref="Checkpoints"/> grows as the pass goes, so a view opened over it can already seek
/// anywhere in the part inflated so far. The package stream may be shared with other readers in
/// between calls — the bit reader re-seats the stream position on every refill.
/// </summary>
internal sealed class DocxEntryInflationCursor
{
    private readonly DeflateCheckpointInflater? _inflater;
    private readonly ZipStoredEntryStream? _stored;
    private readonly DeflateCheckpointSelector? _selector;
    private long _nextCheckpointAt;

    /// <summary>With <c>snapToSparseBoundaries</c>: every candidate checkpoint as it is captured (see <see cref="DeflateCheckpointSelector.CandidateCaptured"/>).</summary>
    public event Action<DeflateCheckpoint>? CandidateCaptured;

    public byte[] Chunk { get; } = new byte[DocxEntryInflation.ChunkSize];
    public List<DeflateCheckpoint> Checkpoints { get; }
    public long BytesInflated { get; private set; }
    public bool IsFinished { get; private set; }

    /// <param name="snapToSparseBoundaries">
    /// True: checkpoints are chosen by <see cref="DeflateCheckpointSelector"/> (the block boundary of each region that
    /// needs the fewest window bytes); false: the first boundary after every <paramref name="checkpointSpanBytes"/>.
    /// </param>
    public DocxEntryInflationCursor(Stream package, ZipEntryLocation entry, long checkpointSpanBytes, bool snapToSparseBoundaries = false)
    {
        Checkpoints = new List<DeflateCheckpoint>();
        if (!entry.IsDeflate)
        {
            _stored = new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
            return;
        }
        _inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
        if (checkpointSpanBytes == long.MaxValue) return;
        if (snapToSparseBoundaries)
        {
            _selector = new DeflateCheckpointSelector(_inflater, checkpointSpanBytes);
            _selector.CandidateCaptured += candidate => CandidateCaptured?.Invoke(candidate);
            Checkpoints = _selector.Checkpoints;
            return;
        }
        _inflater.BlockBoundaryReached += source =>
        {
            if (source.TotalOut < _nextCheckpointAt) return;
            Checkpoints.Add(source.CaptureCheckpoint());
            _nextCheckpointAt = source.TotalOut + checkpointSpanBytes;
        };
    }

    /// <summary>Inflates the next chunk into <see cref="Chunk"/>; returns its length, 0 once the entry is exhausted.</summary>
    public int ReadNextChunk()
    {
        if (IsFinished) return 0;
        int read = _inflater is not null ? _inflater.Read(Chunk, 0, Chunk.Length) : _stored!.Read(Chunk, 0, Chunk.Length);
        if (read == 0) { IsFinished = true; _selector?.Finish(); }
        BytesInflated += read;
        return read;
    }
}
