using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// The one sequential pass over a .docx's text parts: the main document part (word/document.xml) through
/// <see cref="DocxBodyBlockScanner"/>, then footnotes.xml and endnotes.xml through
/// <see cref="DocxNotesPartScanner"/>. In sidecar mode each part gets deflate checkpoints roughly
/// every <see cref="DefaultCheckpointSpanBytes"/> and nothing is kept (memory stays at a few MB
/// whatever the size); in in-memory mode — small documents, where checkpoints would cost more
/// than they save (<see cref="WordIndexPolicy"/>) — the inflated parts are returned instead.
/// </summary>
public static class DocxRandomAccessIndexBuilder
{
    public const long DefaultCheckpointSpanBytes = 1 << 20;

    /// <summary>The inflated text parts of a small document, kept for its session (in-memory mode).</summary>
    public sealed record InMemoryParts(byte[] Body, byte[]? Footnotes, byte[]? Endnotes);

    /// <summary>Sidecar mode: checkpoints, no copies.</summary>
    public static DocxRandomAccessIndex Build(Stream package, DocxDocumentParts parts, long sourceLength, long sourceLastWriteTicksUtc,
                                              IProgress<WordDocumentIndexProgress>? progress = null,
                                              long checkpointSpanBytes = DefaultCheckpointSpanBytes,
                                              CancellationToken cancellation = default)
        => Scan(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, checkpointSpanBytes, copies: null, cancellation);

    /// <summary>In-memory mode: no checkpoints; the inflated parts are returned alongside the index.</summary>
    public static (DocxRandomAccessIndex index, InMemoryParts parts) BuildInMemory(Stream package, DocxDocumentParts parts, long sourceLength,
                                                                                    long sourceLastWriteTicksUtc,
                                                                                    IProgress<WordDocumentIndexProgress>? progress = null,
                                                                                    CancellationToken cancellation = default)
    {
        var copies = new MemoryStream?[3];
        DocxRandomAccessIndex index = Scan(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, long.MaxValue, copies, cancellation);
        return (index, new InMemoryParts(copies[0]!.ToArray(), copies[1]?.ToArray(), copies[2]?.ToArray()));
    }

    public static ZipEntryLocation LocateDocumentEntry(Stream package) => LocateDocumentEntry(package, DocxPackagePartNames.Resolve(package));

    public static ZipEntryLocation LocateDocumentEntry(Stream package, DocxPackagePartNames names)
        => LocateTextEntry(package, names.Document) ?? throw new InvalidDataException($"Not a Word document (no main document part {names.Document}).");

    /// <summary>The uncompressed size of everything a reader inflates: body + footnotes + endnotes — what the index policy weighs.</summary>
    public static long TextPartsBytes(Stream package)
    {
        DocxPackagePartNames names = DocxPackagePartNames.Resolve(package);
        return LocateDocumentEntry(package, names).UncompressedSize
             + (LocateTextEntry(package, names.Footnotes)?.UncompressedSize ?? 0)
             + (LocateTextEntry(package, names.Endnotes)?.UncompressedSize ?? 0);
    }

    internal static ZipEntryLocation? LocateTextEntry(Stream package, string? name)
    {
        if (name is null) return null;
        ZipEntryLocation? entry = ZipEntryLocator.Locate(package, name);
        if (entry is not null && !entry.IsDeflate && !entry.IsStored)
            throw new InvalidDataException($"Unsupported zip compression method {entry.CompressionMethod} for {name}.");
        return entry;
    }

    /// <param name="copies">In-memory mode: receives the inflated body, footnotes, endnotes (indexes 0, 1, 2).</param>
    private static DocxRandomAccessIndex Scan(Stream package, DocxDocumentParts parts, long sourceLength, long sourceLastWriteTicksUtc,
                                              IProgress<WordDocumentIndexProgress>? progress, long checkpointSpanBytes,
                                              MemoryStream?[]? copies, CancellationToken cancellation)
    {
        ZipEntryLocation body = LocateDocumentEntry(package, parts.PartNames);
        ZipEntryLocation? footnotes = LocateTextEntry(package, parts.PartNames.Footnotes);
        ZipEntryLocation? endnotes = LocateTextEntry(package, parts.PartNames.Endnotes);
        long totalBytes = body.UncompressedSize + (footnotes?.UncompressedSize ?? 0) + (endnotes?.UncompressedSize ?? 0);
        long done = 0;

        var scanner = new DocxBodyBlockScanner(parts.CreateListLabelResolver(), parts.Styles.HeadingLevelOf);
        if (copies is not null) copies[0] = new MemoryStream();
        List<DeflateCheckpoint> bodyCheckpoints = DocxEntryInflation.Run(package, body, (chunk, read) =>
        {
            scanner.Feed(chunk, read);
            copies?[0]!.Write(chunk, 0, read);
            done += read;
            progress?.Report(new WordDocumentIndexProgress(done, totalBytes, scanner.BlockStartOffsets.Count));
        }, checkpointSpanBytes, cancellation);
        scanner.Finish();
        if (scanner.BodyEndOffset < 0) throw new InvalidDataException($"{parts.PartNames.Document} has no closing <w:body>.");

        DocxNotesPartIndex? ScanNotes(ZipEntryLocation? entry, int copyIndex)
        {
            if (entry is null) return null;
            var notesScanner = new DocxNotesPartScanner();
            if (copies is not null) copies[copyIndex] = new MemoryStream();
            List<DeflateCheckpoint> checkpoints = DocxEntryInflation.Run(package, entry, (chunk, read) =>
            {
                notesScanner.Feed(chunk, read);
                copies?[copyIndex]!.Write(chunk, 0, read);
                done += read;
                progress?.Report(new WordDocumentIndexProgress(done, totalBytes, scanner.BlockStartOffsets.Count));
            }, checkpointSpanBytes, cancellation);
            return CreateNotesPartIndex(entry, notesScanner, checkpoints);
        }

        DocxNotesPartIndex? footnotesPart = ScanNotes(footnotes, 1);
        DocxNotesPartIndex? endnotesPart = ScanNotes(endnotes, 2);

        return CreateIndex(sourceLength, sourceLastWriteTicksUtc, body, scanner, bodyCheckpoints, footnotesPart, endnotesPart);
    }

    /// <summary>The finished notes-part index from its scanner — shared with <see cref="DocxIncrementalIndexBuilder"/> so both passes save the same sidecar.</summary>
    internal static DocxNotesPartIndex CreateNotesPartIndex(ZipEntryLocation entry, DocxNotesPartScanner notesScanner, List<DeflateCheckpoint> checkpoints)
        => new(entry, notesScanner.RootOpenTag, checkpoints,
               notesScanner.NoteIds.ToArray(), notesScanner.NoteStarts.ToArray(), notesScanner.NoteEnds.ToArray());

    /// <summary>The finished index from a body scanner that has seen the whole part — shared with <see cref="DocxIncrementalIndexBuilder"/>.</summary>
    internal static DocxRandomAccessIndex CreateIndex(long sourceLength, long sourceLastWriteTicksUtc, ZipEntryLocation body,
                                                      DocxBodyBlockScanner scanner, List<DeflateCheckpoint> bodyCheckpoints,
                                                      DocxNotesPartIndex? footnotesPart, DocxNotesPartIndex? endnotesPart)
    {
        return new DocxRandomAccessIndex(
            sourceLength, sourceLastWriteTicksUtc, body, scanner.RootOpenTag, scanner.BodyEndOffset, bodyCheckpoints,
            scanner.BlockStartOffsets.ToArray(), scanner.BlockKinds.ToArray(), scanner.BlockTextLengths.ToArray(),
            scanner.BlockStyleIds.ToArray(), scanner.StyleIdTable.ToArray(),
            scanner.BlockListLabelIds.ToArray(), scanner.ListLabels.ToArray(),
            scanner.Headings.ToArray(), scanner.NoteReferences.ToArray(),
            footnotesPart, endnotesPart).WithOpenFields(scanner.BlockOpenFieldIds.ToArray(), scanner.OpenFieldStacks.ToArray())
                                        .WithSections(scanner.SectionLastBlocks.ToArray(), scanner.SectionTextWidths.ToArray())
                                        .WithBookmarks(scanner.BookmarkNames.ToArray(), scanner.BookmarkBlocks.ToArray());
    }
}
