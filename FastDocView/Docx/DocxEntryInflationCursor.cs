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
    private long _nextCheckpointAt;

    public byte[] Chunk { get; } = new byte[DocxEntryInflation.ChunkSize];
    public List<DeflateCheckpoint> Checkpoints { get; } = new();
    public long BytesInflated { get; private set; }
    public bool IsFinished { get; private set; }

    public DocxEntryInflationCursor(Stream package, ZipEntryLocation entry, long checkpointSpanBytes)
    {
        if (!entry.IsDeflate)
        {
            _stored = new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
            return;
        }
        _inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
        if (checkpointSpanBytes != long.MaxValue)
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
        if (read == 0) IsFinished = true;
        BytesInflated += read;
        return read;
    }
}
