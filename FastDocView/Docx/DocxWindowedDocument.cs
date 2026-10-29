using System.Xml.Linq;
using FastDocView.DocxHtml;
using FastDocView.Html;

namespace FastDocView.Docx;

/// <summary>
/// A .docx opened for windowed rendering. Two ways in:
///   • <see cref="Open"/> — file-system hosts: the policy picks in-memory vs sidecar, the sidecar
///     is cached next to the other indexes;
///   • <see cref="OpenFromStream"/> — any host (WASM): the caller supplies the package stream and
///     an optional previously saved index; nothing touches a file system.
/// A sidecar first open with <see cref="WordIndexPolicy.IndexesOnDemand"/> returns after the first
/// blocks are scanned: <see cref="RenderBlocks"/> extends the index as far as it is asked
/// (<see cref="DocxIncrementalIndexBuilder"/>), and when the scan reaches the end the finished index
/// replaces the builder (and is saved as the sidecar by <see cref="Open"/>).
/// </summary>
public sealed class DocxWindowedDocument : IWordWindowedDocument
{
    /// <summary>Blocks scanned before an on-demand open returns — more than a first screen.</summary>
    public const int OnDemandFirstBlockCount = 100;

    private DocxRandomAccessIndex? _index;                  // null while the on-demand scan is running
    private DocxIncrementalIndexBuilder? _builder;          // null once the index is complete
    private IDocxBlockTable _blocks;
    private readonly string? _sidecarPathToSave;
    private readonly DocxBlockXmlReader _blockReader;
    private readonly DocxStyleCascade _styles;
    private readonly DocxBlockHtmlRenderer _renderer;
    private DocxNoteXmlReader _noteReader;
    private readonly DocxNoteNumberFormats _noteFormats;
    private WordDocumentFileStream? _releasableFile;   // set by Open(path): released after every window
    private bool _isReadingFromFile;                    // false when the whole text was inflated in memory (small .docx)
    private readonly object _readLock = new();          // one package stream is shared by the block and note readers (and the on-demand scan)
    private readonly Dictionary<(WordDocumentNoteKind kind, int id), string> _noteHtmlCache = new();
    private long _noteHtmlCacheCharacters;
    private const long NoteHtmlCacheCharacterLimit = 2_000_000;   // bounded: a few MB of rendered notes, never the part
    private readonly Dictionary<(WordDocumentNoteKind kind, int id), int> _noteNumbers = new();
    private int _noteReferencesRecorded;                // references already counted into _noteNumbers / the counts
    private int _footnoteCount;
    private int _endnoteCount;

    public string FilePath { get; }
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }

    /// <summary>The finished index. Throws while an on-demand open is still scanning (see <see cref="CompleteIndex"/>).</summary>
    public DocxRandomAccessIndex Index => _index ?? throw new InvalidOperationException("The index is still being built on demand; call CompleteIndex() first.");

    public bool IsIndexComplete => _index is not null;

    /// <summary>Blocks whose position is final.</summary>
    public int KnownBlockCount
    {
        get { lock (_readLock) return _builder?.KnownBlockCount ?? _index!.BlockCount; }
    }

    /// <summary>The block count the reader should lay out: exact when complete, the builder's estimate before.</summary>
    public int ScanAhead(int blocks)
    {
        lock (_readLock)
        {
            if (_builder is not { } builder) return _index!.BlockCount;
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                builder.AdvanceUntilKnown((int)Math.Min(int.MaxValue, (long)builder.KnownBlockCount + Math.Max(1, blocks)));
                if (builder.IsBodyScanned) CompleteIndexLocked();
                else RecordNewNoteReferences();
                return _builder?.KnownBlockCount ?? _index!.BlockCount;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    /// <summary>The blocks known so far while an on-demand index is scanning (only grows), exact once complete.</summary>
    public int BlockCount => KnownBlockCount;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy)
    {
        using FileStream file = WordDocumentFileStream.OpenShared(filePath);
        long textBytes = DocxRandomAccessIndexBuilder.TextPartsBytes(file);       // body + footnotes + endnotes
        bool hasCachedIndex = DocxRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, textBytes, hasCachedIndex);
    }

    public static DocxWindowedDocument Open(string filePath, WordIndexPolicy? policy = null,
                                            IProgress<WordDocumentIndexProgress>? progress = null,
                                            CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        DocxRandomAccessIndex? cached = DocxRandomAccessIndex.TryLoad(indexPath, filePath);
        var package = new WordDocumentFileStream(filePath);
        try
        {
            long textBytes = DocxRandomAccessIndexBuilder.TextPartsBytes(package);
            WordIndexDecision decision = policy.Decide(filePath, textBytes, cached is not null);
            DocxWindowedDocument document;
            if (decision.Mode == WordIndexMode.Sidecar && cached is null && policy.IndexesOnDemand)
                document = OpenOnDemand(package, filePath, decision, info.Length, info.LastWriteTimeUtc.Ticks, progress, indexPath);
            else
            {
                document = OpenFromStream(package, filePath, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
                if (decision.Mode == WordIndexMode.Sidecar && cached is null) document.Index.Save(indexPath);
            }
            document._releasableFile = package;
            document._isReadingFromFile = decision.Mode == WordIndexMode.Sidecar || cached is not null;
            package.Release();                                // no handle on the file between windows (Word must be able to save it)
            return document;
        }
        catch
        {
            package.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Host-agnostic entry. <paramref name="package"/> must be seekable and stays open for the
    /// document's lifetime in sidecar mode (it is disposed with the document); in in-memory mode
    /// it is only read during this call. <paramref name="cachedIndex"/> is used when it matches.
    /// <paramref name="indexOnDemand"/>: sidecar mode without a matching cached index returns after
    /// the first blocks and extends the index as blocks are asked for; the host saves
    /// <see cref="Index"/> once <see cref="IsIndexComplete"/>.
    /// </summary>
    public static DocxWindowedDocument OpenFromStream(Stream package, string displayPath, WordIndexDecision decision,
                                                      DocxRandomAccessIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                      IProgress<WordDocumentIndexProgress>? progress = null,
                                                      CancellationToken cancellation = default, bool indexOnDemand = false)
    {
        if (cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc))
            return OverPackage(displayPath, package, cachedIndex, new DocxDocumentParts(package), decision);
        if (decision.Mode == WordIndexMode.Sidecar && indexOnDemand)
            return OpenOnDemand(package, displayPath, decision, sourceLength, sourceLastWriteTicksUtc, progress, sidecarPathToSave: null);

        var parts = new DocxDocumentParts(package);
        if (decision.Mode == WordIndexMode.InMemory)
        {
            (DocxRandomAccessIndex index, DocxRandomAccessIndexBuilder.InMemoryParts inflated) =
                DocxRandomAccessIndexBuilder.BuildInMemory(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
            package.Dispose();
            return new DocxWindowedDocument(displayPath, index, null, DocxBlockXmlReader.OverBody(inflated.Body, index),
                                            DocxNoteXmlReader.OverBytes(inflated.Footnotes, index.FootnotesPart, inflated.Endnotes, index.EndnotesPart),
                                            parts, decision, sidecarPathToSave: null);
        }

        DocxRandomAccessIndex built = DocxRandomAccessIndexBuilder.Build(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation: cancellation);
        return OverPackage(displayPath, package, built, parts, decision);
    }

    private static DocxWindowedDocument OverPackage(string displayPath, Stream package, DocxRandomAccessIndex index, DocxDocumentParts parts, WordIndexDecision decision)
        => new(displayPath, index, null, DocxBlockXmlReader.OverPackage(package, index),
               DocxNoteXmlReader.OverPackage(package, index.FootnotesPart, index.EndnotesPart), parts, decision, sidecarPathToSave: null);

    private static DocxWindowedDocument OpenOnDemand(Stream package, string displayPath, WordIndexDecision decision, long sourceLength,
                                                     long sourceLastWriteTicksUtc, IProgress<WordDocumentIndexProgress>? progress, string? sidecarPathToSave)
    {
        var parts = new DocxDocumentParts(package);
        var builder = new DocxIncrementalIndexBuilder(package, parts, sourceLength, sourceLastWriteTicksUtc, progress);
        builder.AdvanceUntilKnown(OnDemandFirstBlockCount);
        var document = new DocxWindowedDocument(displayPath, null, builder, DocxBlockXmlReader.OverIncrementalBuild(package, builder),
                                                DocxNoteXmlReader.OverIncrementalParts(package, builder.Footnotes, builder.Endnotes),
                                                parts, decision, sidecarPathToSave);
        if (builder.IsBodyScanned) document.CompleteIndexLocked();      // a short body: done already
        return document;
    }

    private DocxWindowedDocument(string filePath, DocxRandomAccessIndex? index, DocxIncrementalIndexBuilder? builder,
                                 DocxBlockXmlReader blockReader, DocxNoteXmlReader noteReader,
                                 DocxDocumentParts parts, WordIndexDecision decision, string? sidecarPathToSave)
    {
        FilePath = filePath;
        _index = index;
        _builder = builder;
        _blocks = (IDocxBlockTable?)index ?? builder!;
        _sidecarPathToSave = sidecarPathToSave;
        _blockReader = blockReader;
        _noteReader = noteReader;
        _styles = parts.Styles;
        _noteFormats = parts.NoteNumberFormats;
        RecordNewNoteReferences();
        _renderer = new DocxBlockHtmlRenderer(_styles, ResolveNoteLabel, parts.HyperlinkTargets, parts.FootnoteHyperlinkTargets, parts.EndnoteHyperlinkTargets);
        Decision = decision;
        IsRightToLeft = _styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(_styles.ResolveParagraphStyle(null).RunFormatting, IsRightToLeft);
    }

    /// <summary>Counts the note references found since the last call into the number lookup and the per-kind counts.</summary>
    private void RecordNewNoteReferences()
    {
        IReadOnlyList<DocxNoteReference> references = (IReadOnlyList<DocxNoteReference>?)_index?.NoteReferences ?? _builder!.KnownNoteReferences;
        for (; _noteReferencesRecorded < references.Count; _noteReferencesRecorded++)
        {
            DocxNoteReference reference = references[_noteReferencesRecorded];
            _noteNumbers[(reference.Kind, reference.NoteId)] = reference.Number;
            if (reference.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
        }
    }

    /// <summary>
    /// Finishes an on-demand index now (scans the rest of the body and the notes parts) and saves
    /// the sidecar when the document was opened from a path. No-op when already complete.
    /// </summary>
    public DocxRandomAccessIndex CompleteIndex()
    {
        lock (_readLock) return CompleteIndexLocked();
    }

    private DocxRandomAccessIndex CompleteIndexLocked()
    {
        if (_index is not null) return _index;
        DocxIncrementalIndexBuilder builder = _builder!;
        DocxRandomAccessIndex index = builder.Complete();
        // Swap the scan's live tables for the finished index (same offsets), and free the scanner.
        _index = index;
        _blocks = index;
        _blockReader.Table = index;
        Stream package = _blockReader.Package!;
        _noteReader.Dispose();
        _noteReader = DocxNoteXmlReader.OverPackage(package, index.FootnotesPart, index.EndnotesPart);
        RecordNewNoteReferences();
        _builder = null;
        _headingsSnapshot = null;
        _noteReferences = null;
        if (_sidecarPathToSave is not null)
        {
            try { index.Save(_sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { /* reading goes on; the next open rebuilds */ }
        }
        return index;
    }

    /// <summary>Past the end (a count read before the index completed) it answers 0 instead of throwing.</summary>
    public int EstimatedTextLength(int blockIndex)
    {
        if (_index is { } index) return blockIndex < index.BlockCount ? index.BlockTextLengths[blockIndex] : 0;
        lock (_readLock)
        {
            if (_builder is not { } builder) return blockIndex < _index!.BlockCount ? _index.BlockTextLengths[blockIndex] : 0;
            return blockIndex < builder.KnownBlockCount ? builder.BlockTextLengthOf(blockIndex) : builder.AverageTextLength;
        }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            // Only documents that still read the file (index mode) are at risk; a stale in-memory body is still consistent.
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                if (_builder is { } builder)
                {
                    builder.AdvanceUntilKnown(firstBlockIndex + count);     // the on-demand scan reopens the released file like any read
                    if (builder.IsBodyScanned) CompleteIndexLocked();
                    else RecordNewNoteReferences();
                }
                return RenderBlocksLocked(firstBlockIndex, count);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        // The window plus its adjacent paragraphs: border groups and deleted paragraph marks reach one block across the edge
        // (and a chain of joined paragraphs back as far as it goes) — still only blocks next to the window are read.
        var elements = new Dictionary<int, XElement>();
        bool IsParagraphBlock(int blockIndex) => blockIndex >= 0 && blockIndex < _blocks.BlockCount && _blocks.BlockKindOf(blockIndex) == WordDocumentBlockKind.Paragraph;
        int readFirst = IsParagraphBlock(firstBlockIndex - 1) ? firstBlockIndex - 1 : firstBlockIndex;      // one read, one inflate pass
        int readEnd = IsParagraphBlock(firstBlockIndex + count) ? firstBlockIndex + count + 1 : firstBlockIndex + count;
        foreach ((int blockIndex, XElement element) in _blockReader.ReadBlocks(readFirst, readEnd - readFirst)) elements[blockIndex] = element;
        XElement? ParagraphAt(int blockIndex)
        {
            if (blockIndex < 0 || blockIndex >= _blocks.BlockCount || _blocks.BlockKindOf(blockIndex) != WordDocumentBlockKind.Paragraph) return null;
            if (!elements.TryGetValue(blockIndex, out XElement? element))
                foreach ((int index, XElement read) in _blockReader.ReadBlocks(blockIndex, 1)) elements[index] = element = read;
            return element is not null && element.Name == DocxStyleCascade.W + "p" ? element : null;
        }
        for (int blockIndex = firstBlockIndex; blockIndex < firstBlockIndex + count && elements.TryGetValue(blockIndex, out XElement? element); blockIndex++)
        {
            int groupStart = blockIndex;
            while (ParagraphAt(groupStart - 1) is { } joined && DocxBlockHtmlRenderer.HasDeletedParagraphMark(joined)) groupStart--;
            XElement[]? mergedBefore = groupStart < blockIndex && element.Name == DocxStyleCascade.W + "p"
                ? Enumerable.Range(groupStart, blockIndex - groupStart).Select(index => elements[index]).ToArray() : null;
            var neighbours = new DocxBlockNeighbours(ParagraphAt(groupStart - 1), ParagraphAt(blockIndex + 1), mergedBefore, IsFirstInContainer: groupStart == 0);
            string? previousStyleKey = blockIndex > 0 ? _blocks.StyleIdOf(blockIndex - 1) ?? "" : null;
            string html = _renderer.Render(element, _blocks.ListLabelOf(blockIndex), previousStyleKey, neighbours with { OpenFields = _blocks.OpenFieldsAt(groupStart) });
            result.Add(new WordDocumentBlockHtml(blockIndex, _blocks.BlockKindOf(blockIndex), html, RenderNotesOf(blockIndex)));
        }
        return result;
    }

    /// <summary>The notes a block references, rendered — what a Word-style notes pane shows for the visible text.</summary>
    private IReadOnlyList<WordDocumentNoteHtml> RenderNotesOf(int blockIndex)
    {
        ReadOnlySpan<DocxNoteReference> references = _blocks.NoteReferencesInBlocks(blockIndex, blockIndex);
        if (references.IsEmpty) return [];
        var notes = new List<WordDocumentNoteHtml>(references.Length);
        foreach (DocxNoteReference reference in references)
        {
            string label = reference.Number > 0 ? _noteFormats.Format(reference.Kind, reference.Number) : "*";
            notes.Add(new WordDocumentNoteHtml(reference.Kind, reference.Number, label, RenderNoteBody(reference, label)));
        }
        return notes;
    }

    /// <summary>One note read on demand from its part (never the whole part), rendered, kept in a size-bounded cache.</summary>
    private string RenderNoteBody(DocxNoteReference reference, string label)
    {
        if (_noteHtmlCache.TryGetValue((reference.Kind, reference.NoteId), out string? cached)) return cached;
        XElement? body = _noteReader.Read(reference.Kind, reference.NoteId);
        string html = body is null ? "" : _renderer.RenderNote(body, reference.Kind, label);
        if (_noteHtmlCacheCharacters + html.Length > NoteHtmlCacheCharacterLimit) { _noteHtmlCache.Clear(); _noteHtmlCacheCharacters = 0; }
        _noteHtmlCache[(reference.Kind, reference.NoteId)] = html;
        _noteHtmlCacheCharacters += html.Length;
        return html;
    }

    private (int number, string label)? ResolveNoteLabel(WordDocumentNoteKind kind, int noteId)
        => _noteNumbers.TryGetValue((kind, noteId), out int number) ? (number, number > 0 ? _noteFormats.Format(kind, number) : "*") : null;

    /// <summary>
    /// From the index — captured during the scan, so a reopen reads no blocks for the TOC. While an
    /// on-demand open is scanning: the headings of the blocks known so far.
    /// </summary>
    public IReadOnlyList<WordDocumentHeading> Headings
    {
        get
        {
            if (_index is { } index) return index.Headings;
            lock (_readLock)
            {
                if (_builder is not { } builder) return _index!.Headings;
                IReadOnlyList<WordDocumentHeading> known = builder.KnownHeadings;
                if (_headingsSnapshot is null || _headingsSnapshot.Length != known.Count) _headingsSnapshot = known.ToArray();
                return _headingsSnapshot;
            }
        }
    }
    private WordDocumentHeading[]? _headingsSnapshot;

    /// <summary>Page text width per section from the index; while scanning, the sections ended so far.</summary>
    public IReadOnlyList<WordDocumentSection> Sections
    {
        get
        {
            if (_index is { } index) return _sections ??= WordDocumentPageSetup.FromSectionEnds(index.SectionLastBlocks, index.SectionTextWidths);
            lock (_readLock)
            {
                if (_builder is not { } builder) return _sections ??= WordDocumentPageSetup.FromSectionEnds(_index!.SectionLastBlocks, _index.SectionTextWidths);
                return WordDocumentPageSetup.FromSectionEnds(builder.KnownSectionLastBlocks.ToArray(), builder.KnownSectionTextWidths.ToArray());
            }
        }
    }
    private IReadOnlyList<WordDocumentSection>? _sections;

    public (int footnotes, int endnotes) NoteCounts
    {
        get { lock (_readLock) return (_footnoteCount, _endnoteCount); }
    }

    /// <summary>Every note reference in body order — while an on-demand open is scanning, those of the blocks known so far.</summary>
    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
    {
        get
        {
            lock (_readLock)
            {
                IReadOnlyList<DocxNoteReference> references = (IReadOnlyList<DocxNoteReference>?)_index?.NoteReferences ?? _builder!.KnownNoteReferences;
                if (_noteReferences is null || _noteReferences.Length != references.Count)
                    _noteReferences = references.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
                return _noteReferences;
            }
        }
    }
    private WordDocumentNoteReference[]? _noteReferences;

    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;

    public void Dispose()
    {
        lock (_readLock)
        {
            _noteReader.Dispose();
            _blockReader.Dispose();
        }
    }
}
