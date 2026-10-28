using System.Collections.Concurrent;
using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// One sequential pass over a zip entry's uncompressed bytes, chunk by chunk, capturing a deflate
/// checkpoint roughly every <c>checkpointSpanBytes</c> of output (pass <see cref="long.MaxValue"/>
/// for none). Shared by the body and the notes parts so all three get the same random access.
/// Inflating and consuming run on two threads through a hand-off of a few fixed buffers (memory
/// stays bounded); in the browser, where threads are not available, they run in one.
/// </summary>
internal static class DocxEntryInflation
{
    private const int ChunkSize = 1 << 18;
    private const int BuffersInFlight = 3;

    public static List<DeflateCheckpoint> Run(Stream package, ZipEntryLocation entry, Action<byte[], int> onChunk,
                                              long checkpointSpanBytes, CancellationToken cancellation)
    {
        var checkpoints = new List<DeflateCheckpoint>();
        if (!entry.IsDeflate)
        {
            var chunk = new byte[ChunkSize];
            using var stored = new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                int read = stored.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                onChunk(chunk, read);
            }
            return checkpoints;
        }

        var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
        long nextCheckpointAt = 0;
        if (checkpointSpanBytes != long.MaxValue)
            inflater.BlockBoundaryReached += source =>
            {
                if (source.TotalOut < nextCheckpointAt) return;
                checkpoints.Add(source.CaptureCheckpoint());
                nextCheckpointAt = source.TotalOut + checkpointSpanBytes;
            };

        if (!CanUseBackgroundThread)
        {
            var chunk = new byte[ChunkSize];
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                int read = inflater.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                onChunk(chunk, read);
            }
            return checkpoints;
        }

        // Producer (inflate) on a thread, consumer (scan) here; `free` holds empty buffers, `filled` the inflated ones.
        var free = new BlockingCollection<byte[]>(new ConcurrentQueue<byte[]>());
        var filled = new BlockingCollection<(byte[] buffer, int length)>(new ConcurrentQueue<(byte[], int)>());
        for (int i = 0; i < BuffersInFlight; i++) free.Add(new byte[ChunkSize]);
        Exception? producerError = null;
        var producer = new Thread(() =>
        {
            try
            {
                while (true)
                {
                    byte[] buffer = free.Take(cancellation);
                    int read = inflater.Read(buffer, 0, buffer.Length);
                    if (read == 0) break;
                    filled.Add((buffer, read), cancellation);
                }
            }
            catch (Exception exception) { producerError = exception; }
            finally { filled.CompleteAdding(); }
        })
        { IsBackground = true, Name = "docx-inflate" };
        producer.Start();

        try
        {
            foreach ((byte[] buffer, int length) in filled.GetConsumingEnumerable(cancellation))
            {
                onChunk(buffer, length);
                free.Add(buffer);
            }
        }
        finally
        {
            free.CompleteAdding();
            producer.Join();
        }
        if (producerError is not null and not OperationCanceledException) throw new InvalidDataException("Inflating failed: " + producerError.Message, producerError);
        cancellation.ThrowIfCancellationRequested();
        return checkpoints;
    }

    private static bool CanUseBackgroundThread =>
#if NET5_0_OR_GREATER
        !OperatingSystem.IsBrowser();
#else
        true;
#endif

    /// <summary>A seekable view of the entry's uncompressed bytes over the package (which the view does not own).</summary>
    public static Stream OpenView(Stream package, ZipEntryLocation entry, IReadOnlyList<DeflateCheckpoint> checkpoints)
        => entry.IsDeflate
            ? new DeflateRandomAccessStream(package, entry.DataOffset, entry.CompressedSize, entry.UncompressedSize, checkpoints)
            : new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
}
