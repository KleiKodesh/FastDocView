using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The one sequential pass over word/document.xml: inflates it chunk by chunk, feeds every chunk to
/// <see cref="DocxBodyBlockScanner"/>, and either captures a deflate checkpoint roughly every
/// <see cref="DefaultCheckpointSpanBytes"/> (sidecar mode: memory stays at a few MB whatever the
/// document size) or keeps the whole body in a buffer (in-memory mode for small documents, where
/// a checkpoint index would cost more than it saves — see <see cref="WordIndexPolicy"/>).
/// </summary>
public static class DocxRandomAccessIndexBuilder
{
    public const string DocumentEntryName = "word/document.xml";
    public const long DefaultCheckpointSpanBytes = 1 << 20;
    private const int ChunkSize = 1 << 18;

    /// <summary>Sidecar mode: checkpoints, no body copy.</summary>
    public static DocxRandomAccessIndex Build(Stream package, DocxDocumentParts parts, long sourceLength, long sourceLastWriteTicksUtc,
                                              IProgress<WordDocumentIndexProgress>? progress = null,
                                              long checkpointSpanBytes = DefaultCheckpointSpanBytes,
                                              CancellationToken cancellation = default)
        => Scan(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, checkpointSpanBytes, bodyCopy: null, cancellation);

    /// <summary>In-memory mode: no checkpoints; the uncompressed body is returned alongside the index.</summary>
    public static (DocxRandomAccessIndex index, byte[] body) BuildInMemory(Stream package, DocxDocumentParts parts, long sourceLength,
                                                                           long sourceLastWriteTicksUtc,
                                                                           IProgress<WordDocumentIndexProgress>? progress = null,
                                                                           CancellationToken cancellation = default)
    {
        var body = new MemoryStream();
        DocxRandomAccessIndex index = Scan(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, long.MaxValue, body, cancellation);
        return (index, body.ToArray());
    }

    public static ZipEntryLocation LocateDocumentEntry(Stream package)
    {
        ZipEntryLocation entry = ZipEntryLocator.Locate(package, DocumentEntryName)
                                 ?? throw new InvalidDataException("Not a Word document (no word/document.xml).");
        if (!entry.IsDeflate && !entry.IsStored)
            throw new InvalidDataException($"Unsupported zip compression method {entry.CompressionMethod}.");
        return entry;
    }

    private static DocxRandomAccessIndex Scan(Stream package, DocxDocumentParts parts, long sourceLength, long sourceLastWriteTicksUtc,
                                              IProgress<WordDocumentIndexProgress>? progress, long checkpointSpanBytes,
                                              MemoryStream? bodyCopy, CancellationToken cancellation)
    {
        ZipEntryLocation entry = LocateDocumentEntry(package);
        var scanner = new DocxBodyBlockScanner(parts.CreateListLabelResolver());
        var checkpoints = new List<DeflateCheckpoint>();
        var chunk = new byte[ChunkSize];
        long totalOut = 0;

        void Consume(int read)
        {
            scanner.Feed(chunk, read);
            bodyCopy?.Write(chunk, 0, read);
            totalOut += read;
            progress?.Report(new WordDocumentIndexProgress(totalOut, entry.UncompressedSize, scanner.BlockStartOffsets.Count));
        }

        if (entry.IsDeflate)
        {
            var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
            long nextCheckpointAt = 0;
            if (bodyCopy is null)
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
                Consume(read);
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
                Consume(read);
            }
        }

        if (scanner.BodyEndOffset < 0) throw new InvalidDataException("word/document.xml has no closing <w:body>.");

        return new DocxRandomAccessIndex(
            sourceLength, sourceLastWriteTicksUtc, entry, scanner.RootOpenTag, scanner.BodyEndOffset, checkpoints,
            scanner.BlockStartOffsets.ToArray(), scanner.BlockKinds.ToArray(), scanner.BlockTextLengths.ToArray(),
            scanner.BlockStyleIds.ToArray(), scanner.StyleIdTable.ToArray(),
            scanner.BlockListLabelIds.ToArray(), scanner.ListLabels.ToArray());
    }
}
