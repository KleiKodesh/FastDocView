using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// One sequential pass over a zip entry's uncompressed bytes, chunk by chunk, capturing a deflate
/// checkpoint roughly every <c>checkpointSpanBytes</c> of output (pass <see cref="long.MaxValue"/>
/// for none). Shared by the body and the notes parts so all three get the same random access.
/// The pass can run in one go (<see cref="Run"/>) or be pulled one chunk at a time
/// (<see cref="DocxEntryInflationCursor"/>) — both produce the same chunks and checkpoints.
/// </summary>
internal static class DocxEntryInflation
{
    public const int ChunkSize = 1 << 18;

    public static List<DeflateCheckpoint> Run(Stream package, ZipEntryLocation entry, Action<byte[], int> onChunk,
                                              long checkpointSpanBytes, CancellationToken cancellation)
    {
        var cursor = new DocxEntryInflationCursor(package, entry, checkpointSpanBytes);
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            int read = cursor.ReadNextChunk();
            if (read == 0) break;
            onChunk(cursor.Chunk, read);
        }
        return cursor.Checkpoints;
    }

    /// <summary>A seekable view of the entry's uncompressed bytes over the package (which the view does not own).</summary>
    public static Stream OpenView(Stream package, ZipEntryLocation entry, IReadOnlyList<DeflateCheckpoint> checkpoints)
        => entry.IsDeflate
            ? new DeflateRandomAccessStream(package, entry.DataOffset, entry.CompressedSize, entry.UncompressedSize, checkpoints)
            : new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
}
