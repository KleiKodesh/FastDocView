using SeekableDeflate;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// One sequential pass over a zip entry's uncompressed bytes, chunk by chunk, capturing a deflate
/// checkpoint roughly every <c>checkpointSpanBytes</c> of output (pass <see cref="long.MaxValue"/>
/// for none). Shared by the body and the notes parts so all three get the same random access.
/// </summary>
internal static class DocxEntryInflation
{
    private const int ChunkSize = 1 << 18;

    public static List<DeflateCheckpoint> Run(Stream package, ZipEntryLocation entry, Action<byte[], int> onChunk,
                                              long checkpointSpanBytes, CancellationToken cancellation)
    {
        var checkpoints = new List<DeflateCheckpoint>();
        var chunk = new byte[ChunkSize];
        if (entry.IsDeflate)
        {
            var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
            long nextCheckpointAt = 0;
            if (checkpointSpanBytes != long.MaxValue)
                inflater.BlockBoundaryReached += source =>
                {
                    if (source.TotalOut < nextCheckpointAt) return;
                    checkpoints.Add(source.CaptureCheckpoint());
                    nextCheckpointAt = source.TotalOut + checkpointSpanBytes;
                };
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                int read = inflater.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                onChunk(chunk, read);
            }
        }
        else
        {
            using var stored = new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                int read = stored.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                onChunk(chunk, read);
            }
        }
        return checkpoints;
    }

    /// <summary>A seekable view of the entry's uncompressed bytes over the package (which the view does not own).</summary>
    public static Stream OpenView(Stream package, ZipEntryLocation entry, IReadOnlyList<DeflateCheckpoint> checkpoints)
        => entry.IsDeflate
            ? new DeflateRandomAccessStream(package, entry.DataOffset, entry.CompressedSize, entry.UncompressedSize, checkpoints)
            : new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
}
