using SeekableDeflate;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// The one sequential pass over an OpenDocument body through <see cref="OdtBodyBlockScanner"/>.
/// Package (.odt/.ott): content.xml is inflated chunk by chunk; in sidecar mode deflate checkpoints
/// are captured roughly every <see cref="DefaultCheckpointSpanBytes"/>, in in-memory mode the
/// inflated content.xml is returned instead. Flat (.fodt): the file itself is read chunk by chunk —
/// offsets are file offsets, nothing to checkpoint or keep. The style cascade must already hold
/// styles.xml (package); the scanner adds the sections it captures.
/// </summary>
public static class OdtRandomAccessIndexBuilder
{
    public const long DefaultCheckpointSpanBytes = 1 << 20;
    private const int ChunkSize = 1 << 18;

    public static OdtRandomAccessIndex BuildFromPackage(Stream package, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc,
                                                        IProgress<WordDocumentIndexProgress>? progress, long checkpointSpanBytes,
                                                        MemoryStream? copy, CancellationToken cancellation)
    {
        ZipEntryLocation content = OdtPackage.LocateContentEntry(package);
        var scanner = new OdtBodyBlockScanner(styles);
        long done = 0;
        List<DeflateCheckpoint> checkpoints = DocxEntryInflation.Run(package, content, (chunk, read) =>
        {
            scanner.Feed(chunk, read);
            copy?.Write(chunk, 0, read);
            done += read;
            progress?.Report(new WordDocumentIndexProgress(done, content.UncompressedSize, scanner.BlockStartOffsets.Count));
        }, checkpointSpanBytes, cancellation);
        return Finish(scanner, styles, sourceLength, sourceLastWriteTicksUtc, isFlat: false, content, checkpoints);
    }

    public static OdtRandomAccessIndex BuildFromFlat(Stream file, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc,
                                                     IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        var scanner = new OdtBodyBlockScanner(styles);
        var chunk = new byte[ChunkSize];
        file.Position = 0;
        long done = 0;
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            int read = file.Read(chunk, 0, chunk.Length);
            if (read <= 0) break;
            scanner.Feed(chunk, read);
            done += read;
            progress?.Report(new WordDocumentIndexProgress(done, file.Length, scanner.BlockStartOffsets.Count));
        }
        return Finish(scanner, styles, sourceLength, sourceLastWriteTicksUtc, isFlat: true, null, []);
    }

    private static OdtRandomAccessIndex Finish(OdtBodyBlockScanner scanner, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc, bool isFlat,
                                               ZipEntryLocation? content, List<DeflateCheckpoint> checkpoints)
    {
        if (!scanner.IsBodyComplete && !scanner.KeepCompleteBlocks())
            throw new InvalidDataException("OpenDocument: no complete office:text body.");
        return new OdtRandomAccessIndex(sourceLength, sourceLastWriteTicksUtc, isFlat, content, checkpoints, scanner.RootOpenTag,
                                        styles.SectionsForRendering(), scanner.BlockStartOffsets.ToArray(), scanner.BlockEndOffsets.ToArray(),
                                        scanner.BlockKinds.ToArray(), scanner.BlockTextLengths.ToArray(), scanner.BlockListLevels.ToArray(),
                                        scanner.BlockListLabelIds.ToArray(), scanner.ListLabels.ToArray(), scanner.Headings.ToArray(),
                                        scanner.NoteReferences.ToArray());
    }
}
