using SeekableDeflate;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>
/// The one sequential pass over an OpenDocument body through <see cref="OdtBodyBlockScanner"/>.
/// Package (.odt/.ott): content.xml is inflated chunk by chunk; in sidecar mode deflate checkpoints
/// are captured roughly every <see cref="DefaultCheckpointSpanBytes"/> (snapped to the thriftiest
/// block boundary, as for .docx), in in-memory mode the inflated content.xml is returned instead.
/// Flat (.fodt): the file itself is read chunk by chunk - offsets are file offsets, nothing to
/// checkpoint or keep. The style cascade must already hold styles.xml (package); the scanner adds
/// the sections it captures. An <see cref="OdtIndexAnchorRecorder"/> rides along for the span anchors.
/// </summary>
public static class OdtRandomAccessIndexBuilder
{
    public const long DefaultCheckpointSpanBytes = 2L << 20;
    internal const int ChunkSize = 1 << 18;

    public static OdtRandomAccessIndex BuildFromPackage(Stream package, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc,
                                                        IProgress<WordDocumentIndexProgress>? progress, long checkpointSpanBytes,
                                                        MemoryStream? copy, CancellationToken cancellation)
    {
        ZipEntryLocation content = OdtPackage.LocateContentEntry(package);
        var scanner = new OdtBodyBlockScanner(styles);
        var anchors = new OdtIndexAnchorRecorder(scanner);
        var cursor = new DocxEntryInflationCursor(package, content, checkpointSpanBytes, snapToSparseBoundaries: true);
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            int read = cursor.ReadNextChunk();
            if (read == 0) break;
            scanner.Feed(cursor.Chunk, read);
            copy?.Write(cursor.Chunk, 0, read);
            progress?.Report(new WordDocumentIndexProgress(cursor.BytesInflated, content.UncompressedSize, scanner.BlockStartOffsets.Count));
        }
        return Finish(scanner, anchors, styles, sourceLength, sourceLastWriteTicksUtc, isFlat: false, content, cursor.Checkpoints);
    }

    public static OdtRandomAccessIndex BuildFromFlat(Stream file, OdtStyleCascade styles, long sourceLength, long sourceLastWriteTicksUtc,
                                                     IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        var scanner = new OdtBodyBlockScanner(styles);
        var anchors = new OdtIndexAnchorRecorder(scanner);
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
        return Finish(scanner, anchors, styles, sourceLength, sourceLastWriteTicksUtc, isFlat: true, null, []);
    }

    internal static OdtRandomAccessIndex Finish(OdtBodyBlockScanner scanner, OdtIndexAnchorRecorder anchors, OdtStyleCascade styles,
                                               long sourceLength, long sourceLastWriteTicksUtc, bool isFlat,
                                               ZipEntryLocation? content, List<DeflateCheckpoint> checkpoints)
    {
        if (!scanner.IsBodyComplete && !scanner.KeepCompleteBlocks())
            throw new InvalidDataException("OpenDocument: no complete office:text body.");
        int blockCount = scanner.BlockStartOffsets.Count;
        long lastEnd = blockCount > 0 ? scanner.BlockEndOffsets[blockCount - 1] : 0;
        return new OdtRandomAccessIndex(sourceLength, sourceLastWriteTicksUtc, isFlat, content, checkpoints, scanner.RootOpenTag,
                                        styles.SectionsForRendering(), new Dictionary<string, OdtListStyle>(scanner.UsedListStyles, StringComparer.Ordinal),
                                        blockCount, lastEnd, anchors.Build(blockCount), WordIndexTextLengthGroups.Totals(scanner.BlockTextLengths),
                                        scanner.Headings.ToArray(), scanner.NoteReferences.ToArray())
        {
            SectionFirstBlocks = scanner.SectionStarts.Select(section => section.FirstBlockIndex).ToArray(),
            SectionTextWidths = scanner.SectionStarts.Select(section => section.TextWidthTwips).ToArray(),
            SectionColumnLefts = scanner.SectionStarts.Select(section => section.ColumnLeftTwips + 1).ToArray(),
            Bookmarks = new WordDocumentBookmarkTable(scanner.BookmarkNames.ToArray(), scanner.BookmarkBlocks.ToArray()),
            ConnectedShapeBlocks = scanner.ConnectedShapeBlocks.ToArray(),
            ConnectedShapes = scanner.ConnectedShapeBlocks.Count == 0 ? OdtConnectedShapeTable.Empty : null,
        };
    }
}
