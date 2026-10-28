using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The one sequential pass over word/document.xml: inflates it chunk by chunk (never holding more
/// than one chunk), captures a deflate checkpoint roughly every <see cref="CheckpointSpanBytes"/>
/// of output, and feeds every chunk to <see cref="DocxBodyBlockScanner"/>. Memory stays at a few MB
/// regardless of document size, which keeps it inside the project's light-compute rule.
/// </summary>
public static class DocxRandomAccessIndexBuilder
{
    public const string DocumentEntryName = "word/document.xml";
    public const long DefaultCheckpointSpanBytes = 1 << 20;
    private const int ChunkSize = 1 << 18;

    public static DocxRandomAccessIndex Build(string documentPath, IProgress<WordDocumentIndexProgress>? progress = null,
                                              long checkpointSpanBytes = DefaultCheckpointSpanBytes,
                                              CancellationToken cancellation = default)
    {
        var info = new FileInfo(documentPath);
        using FileStream file = new(documentPath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        ZipEntryLocation entry = ZipEntryLocator.Locate(file, DocumentEntryName)
                                 ?? throw new InvalidDataException("Not a Word document (no word/document.xml).");
        if (!entry.IsDeflate && !entry.IsStored)
            throw new InvalidDataException($"Unsupported zip compression method {entry.CompressionMethod}.");

        var scanner = new DocxBodyBlockScanner();
        var checkpoints = new List<DeflateCheckpoint>();
        var chunk = new byte[ChunkSize];
        long totalOut = 0;

        if (entry.IsDeflate)
        {
            var inflater = new DeflateCheckpointInflater(file, entry.DataOffset, entry.CompressedSize);
            long nextCheckpointAt = 0;
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
                scanner.Feed(chunk, read);
                totalOut += read;
                progress?.Report(new WordDocumentIndexProgress(totalOut, entry.UncompressedSize, scanner.BlockStartOffsets.Count));
            }
        }
        else
        {
            using var stored = new ZipStoredEntryStream(file, entry.DataOffset, entry.CompressedSize);
            while (true)
            {
                cancellation.ThrowIfCancellationRequested();
                int read = stored.Read(chunk, 0, chunk.Length);
                if (read == 0) break;
                scanner.Feed(chunk, read);
                totalOut += read;
                progress?.Report(new WordDocumentIndexProgress(totalOut, entry.UncompressedSize, scanner.BlockStartOffsets.Count));
            }
        }

        if (scanner.BodyEndOffset < 0) throw new InvalidDataException("word/document.xml has no closing <w:body>.");

        return new DocxRandomAccessIndex(
            info.Length, info.LastWriteTimeUtc.Ticks, entry, scanner.RootOpenTag, scanner.BodyEndOffset, checkpoints,
            scanner.BlockStartOffsets.ToArray(), scanner.BlockKinds.ToArray(), scanner.BlockTextLengths.ToArray(),
            scanner.BlockStyleIds.ToArray(), scanner.StyleIdTable.ToArray());
    }

    /// <summary>Loads the cached index when it still matches the file, otherwise builds and saves it.</summary>
    public static DocxRandomAccessIndex LoadOrBuild(string documentPath, string? cacheDirectory = null,
                                                    IProgress<WordDocumentIndexProgress>? progress = null,
                                                    CancellationToken cancellation = default)
    {
        string indexPath = WordDocumentIndexCacheLocation.IndexPathFor(documentPath, cacheDirectory);
        DocxRandomAccessIndex? cached = DocxRandomAccessIndex.TryLoad(indexPath, documentPath);
        if (cached is not null) return cached;
        DocxRandomAccessIndex built = Build(documentPath, progress, cancellation: cancellation);
        built.Save(indexPath);
        return built;
    }
}
