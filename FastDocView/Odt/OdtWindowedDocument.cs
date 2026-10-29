using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// An OpenDocument text (.odt / .ott package, or flat .fodt) opened for windowed rendering, behind
/// the same <see cref="IWordWindowedDocument"/> contract as .docx and .doc.
///   • package, in-memory mode (small): one pass inflates content.xml and keeps it for the session;
///   • package, sidecar mode: content.xml checkpoints + the index are saved; reopen reads no body;
///   • flat: the file is the random-access view; the index (sidecar when big) only saves the pass.
/// A sidecar first open with <see cref="WordIndexPolicy.IndexesOnDemand"/> returns after the first
/// <see cref="OnDemandFirstBlockCount"/> blocks and extends the index as blocks are asked for
/// (<see cref="OdtIncrementalIndexBuilder"/>); the finished index is saved when the scan completes.
/// </summary>
public sealed class OdtWindowedDocument : IWordWindowedDocument
{
    public const int OnDemandFirstBlockCount = 100;

    private OdtRandomAccessIndex? _index;                // null while an on-demand open is still scanning
    private OdtIncrementalIndexBuilder? _builder;
    private IOdtBlockTable _blocks;
    private readonly string? _sidecarPathToSave;
    private readonly OdtBlockXmlReader _reader;
    private readonly OdtStyleCascade _styles;
    private readonly OdtBlockHtmlRenderer _renderer;
    private readonly object _readLock = new();
    private int _footnoteCount, _endnoteCount, _noteReferencesCounted;
    private WordDocumentFileStream? _releasableFile;   // set by Open(path): released after every window
    private bool _isReadingFromFile;                    // false when content.xml was inflated in memory (small package)

    public string FilePath { get; }

    /// <summary>The blocks known so far while an on-demand index is scanning (only grows), exact once complete.</summary>
    public int BlockCount => KnownBlockCount;

    public int ScanAhead(int blocks)
    {
        lock (_readLock)
        {
            if (_builder is not { } builder) return _index!.BlockCount;
            if (_releasableFile is not null && _releasableFile.HasSourceChanged) throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                builder.AdvanceUntilKnown((int)Math.Min(int.MaxValue, (long)builder.KnownBlockCount + Math.Max(1, blocks)));
                if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                else CountNewNoteReferences();
                return _builder?.KnownBlockCount ?? _index!.BlockCount;
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public bool IsIndexComplete => _index is not null;

    public int KnownBlockCount
    {
        get { lock (_readLock) return _builder?.KnownBlockCount ?? _index!.BlockCount; }
    }

    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }

    /// <summary>The finished index. Throws while an on-demand open is still scanning (see <see cref="CompleteIndex"/>).</summary>
    public OdtRandomAccessIndex Index => _index ?? throw new InvalidOperationException("The index is still being built on demand; call CompleteIndex() first.");

    /// <summary>Uncompressed body size: content.xml (package) or the whole file (flat).</summary>
    public static long TextBytes(Stream file, bool isFlat) => isFlat ? file.Length : OdtPackage.LocateContentEntry(file).UncompressedSize;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy, bool isFlat)
    {
        using FileStream file = OpenRead(filePath);
        bool hasCachedIndex = OdtRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, TextBytes(file, isFlat), hasCachedIndex);
    }

    public static OdtWindowedDocument Open(string filePath, bool isFlat, WordIndexPolicy? policy = null,
                                           IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        OdtRandomAccessIndex? cached = OdtRandomAccessIndex.TryLoad(indexPath, filePath);
        var file = new WordDocumentFileStream(filePath);
        try
        {
            WordIndexDecision decision = policy.Decide(filePath, TextBytes(file, isFlat), cached is not null);
            bool isOnDemand = decision.Mode == WordIndexMode.Sidecar && cached is null && policy.IndexesOnDemand;
            OdtWindowedDocument document = isOnDemand
                ? OpenOnDemand(file, filePath, isFlat, decision, info.Length, info.LastWriteTimeUtc.Ticks, progress, indexPath)
                : OpenFromStream(file, filePath, isFlat, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (!isOnDemand && decision.Mode == WordIndexMode.Sidecar && cached is null)
            {
                try { document.Index.Save(indexPath); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache; the open document stands
            }
            document._releasableFile = file;
            document._isReadingFromFile = isFlat || decision.Mode == WordIndexMode.Sidecar
                                          || (cached is not null && cached.MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks));
            file.Release();                                   // no handle on the file between windows (Word / LibreOffice must be able to save it)
            return document;
        }
        catch
        {
            file.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Host-agnostic entry. <paramref name="file"/> must be seekable; the document takes ownership: it
    /// stays open for the document's lifetime and is disposed with it, except in package in-memory
    /// mode, where it is read and disposed during this call.
    /// </summary>
    public static OdtWindowedDocument OpenFromStream(Stream file, string displayPath, bool isFlat, WordIndexDecision decision,
                                                     OdtRandomAccessIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                     IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        if (!isFlat) OdtPackage.ThrowWhenEncrypted(file);
        if (cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc) && cachedIndex.IsFlat == isFlat)
            return OverIndex(displayPath, file, cachedIndex, decision);

        // the index pass needs every style (list styles for the labels); the session keeps only what rendering needs
        OdtStyleCascade indexingStyles = IndexingStyles(file, isFlat);

        if (isFlat)
            return OverIndex(displayPath, file, OdtRandomAccessIndexBuilder.BuildFromFlat(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, cancellation), decision);

        if (decision.Mode == WordIndexMode.InMemory)
        {
            var content = new MemoryStream();
            OdtRandomAccessIndex index = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, long.MaxValue, content, cancellation);
            file.Dispose();
            return new OdtWindowedDocument(displayPath, index, null, OdtBlockXmlReader.OverBytes(content.ToArray(), index),
                                           RenderingStyles(index.RootOpenTag, index.StyleSections), decision, sidecarPathToSave: null);
        }

        OdtRandomAccessIndex built = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress,
                                                                                OdtRandomAccessIndexBuilder.DefaultCheckpointSpanBytes, null, cancellation);
        return OverIndex(displayPath, file, built, decision);
    }

    private static OdtStyleCascade IndexingStyles(Stream file, bool isFlat)
    {
        var styles = new OdtStyleCascade();
        if (!isFlat) foreach (XElement section in OdtPackage.ReadStylesSections(file)) styles.AddSection(section);
        styles.PageSetup = OdtPackage.ReadPageSetup(file, isFlat);
        return styles;
    }

    private static OdtWindowedDocument OverIndex(string displayPath, Stream file, OdtRandomAccessIndex index, WordIndexDecision decision)
        => new(displayPath, index, null, index.IsFlat ? OdtBlockXmlReader.OverFlatFile(file, index) : OdtBlockXmlReader.OverPackage(file, index),
               RenderingStyles(index.RootOpenTag, index.StyleSections), decision, sidecarPathToSave: null);

    /// <summary>
    /// Returns once the first blocks are known; the rest of the index is scanned as blocks are asked
    /// for. The style sections the renderer needs (font faces, automatic styles) precede office:body,
    /// so they are complete by the time the first block is.
    /// </summary>
    private static OdtWindowedDocument OpenOnDemand(Stream file, string displayPath, bool isFlat, WordIndexDecision decision, long sourceLength,
                                                    long sourceLastWriteTicksUtc, IProgress<WordDocumentIndexProgress>? progress, string sidecarPathToSave)
    {
        if (!isFlat) OdtPackage.ThrowWhenEncrypted(file);
        var builder = new OdtIncrementalIndexBuilder(file, isFlat, IndexingStyles(file, isFlat), sourceLength, sourceLastWriteTicksUtc, progress);
        builder.AdvanceUntilKnown(OnDemandFirstBlockCount);
        if (builder.KnownBlockCount == 0)
        {
            // an empty or malformed body: the one-go rules decide (keep the complete blocks, or throw)
            OdtWindowedDocument whole = OverIndex(displayPath, file, builder.Complete(), decision);
            try { whole.Index.Save(sidecarPathToSave); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
            return whole;
        }
        var document = new OdtWindowedDocument(displayPath, null, builder, OdtBlockXmlReader.OverIncrementalBuild(file, builder),
                                               RenderingStyles(builder.RootOpenTag, builder.StyleSectionsForRendering()), decision, sidecarPathToSave);
        if (builder.IsEveryBlockKnown) document.CompleteIndexLocked();   // a short body: done already
        return document;
    }

    /// <summary>The cascade from captured style sections (no list styles) — a reopen never touches styles.xml.</summary>
    private static OdtStyleCascade RenderingStyles(string rootOpenTag, string[] styleSections)
    {
        var styles = new OdtStyleCascade();
        foreach (string section in styleSections) styles.AddSection(OdtBodyBlockScanner.ParseSection(rootOpenTag, section));
        return styles;
    }

    private OdtWindowedDocument(string filePath, OdtRandomAccessIndex? index, OdtIncrementalIndexBuilder? builder, OdtBlockXmlReader reader,
                                OdtStyleCascade styles, WordIndexDecision decision, string? sidecarPathToSave)
    {
        FilePath = filePath;
        _index = index;
        _builder = builder;
        _blocks = (IOdtBlockTable?)index ?? builder!;
        _sidecarPathToSave = sidecarPathToSave;
        _reader = reader;
        _styles = styles;
        _renderer = new OdtBlockHtmlRenderer(styles);
        Decision = decision;
        CountNewNoteReferences();
        IsRightToLeft = styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(styles.DefaultParagraph.Formatting.RunFormatting, IsRightToLeft);
    }

    private IReadOnlyList<OdtNoteReference> KnownNoteReferences => (IReadOnlyList<OdtNoteReference>?)_index?.NoteReferences ?? _builder!.KnownNoteReferences;

    /// <summary>Adds the note references found since the last call to the per-kind counts.</summary>
    private void CountNewNoteReferences()
    {
        IReadOnlyList<OdtNoteReference> references = KnownNoteReferences;
        for (; _noteReferencesCounted < references.Count; _noteReferencesCounted++)
            if (references[_noteReferencesCounted].Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
    }

    /// <summary>
    /// Finishes an on-demand index now (scans the rest of the body) and saves the sidecar when the
    /// document was opened from a path. No-op when already complete.
    /// </summary>
    public OdtRandomAccessIndex CompleteIndex()
    {
        lock (_readLock) return CompleteIndexLocked();
    }

    private OdtRandomAccessIndex CompleteIndexLocked()
    {
        if (_index is not null) return _index;
        OdtRandomAccessIndex index = _builder!.Complete();
        // Swap the scan's live tables for the finished index (same offsets, same checkpoints) and free the scanner.
        _index = index;
        _blocks = index;
        _reader.Table = index;
        _builder = null;
        _noteReferences = null;
        CountNewNoteReferences();
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
        lock (_readLock)
        {
            if (_builder is { } builder) return blockIndex < builder.KnownBlockCount ? builder.BlockTextLengthOf(blockIndex) : builder.AverageTextLength;
            return blockIndex < _index!.BlockCount ? _index.BlockTextLengths[blockIndex] : 0;
        }
    }

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                if (_builder is { } builder)
                {
                    builder.AdvanceUntilKnown(firstBlockIndex + count);     // the on-demand scan reopens the released file like any read
                    if (builder.IsEveryBlockKnown) CompleteIndexLocked();
                    else CountNewNoteReferences();
                }
                return RenderBlocksLocked(firstBlockIndex, count);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(Math.Max(0, count));
        foreach ((int blockIndex, XElement element) in _reader.ReadBlocks(firstBlockIndex, count))
        {
            (string html, IReadOnlyList<WordDocumentNoteHtml> notes) = _renderer.Render(element, _blocks.ListLabelOf(blockIndex), _blocks.ListLevelOf(blockIndex), _blocks.ListLabelLayoutOf(blockIndex),
                                                                                       _blocks.NoteReferencesInBlocks(blockIndex, blockIndex), null);
            result.Add(new WordDocumentBlockHtml(blockIndex, _blocks.BlockKindOf(blockIndex), html, notes));
        }
        return result;
    }

    /// <summary>Headings of the known blocks while an on-demand index is scanning; all of them once complete.</summary>
    public IReadOnlyList<WordDocumentHeading> Headings
    {
        get { lock (_readLock) return _builder is { } builder ? builder.KnownHeadings.ToArray() : _index!.Headings; }
    }

    /// <summary>Page text width per section (master pages, see <see cref="OdtPageSetup"/>); the known ones while scanning.</summary>
    public int? FindBookmark(string name)
    {
        lock (_readLock)
        {
            if (_index is { } index) return index.Bookmarks.BlockOf(name);
            OdtIncrementalIndexBuilder builder = _builder!;
            try
            {
                // scan on only until the bookmark turns up (a target near the start costs nothing more)
                while (true)
                {
                    if (builder.FindKnownBookmark(name) is { } block) return block;
                    if (!builder.AdvanceOne()) break;
                }
                return CompleteIndexLocked().Bookmarks.BlockOf(name);
            }
            finally { _releasableFile?.Release(); }
        }
    }

    public IReadOnlyList<WordDocumentSection> Sections
    {
        get
        {
            lock (_readLock)
            {
                if (_builder is { } builder) return WordDocumentPageSetup.FromSectionStarts(builder.KnownSectionStarts);
                OdtRandomAccessIndex index = _index!;
                return _sections ??= WordDocumentPageSetup.FromSectionStarts(
                    index.SectionFirstBlocks.Select((first, i) => new WordDocumentSection(first, index.SectionTextWidths[i])));
            }
        }
    }
    private IReadOnlyList<WordDocumentSection>? _sections;

    public (int footnotes, int endnotes) NoteCounts
    {
        get { lock (_readLock) return (_footnoteCount, _endnoteCount); }
    }

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
    {
        get
        {
            lock (_readLock)
            {
                if (_builder is not null)
                    return KnownNoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
                return _noteReferences ??= _index!.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
            }
        }
    }
    private WordDocumentNoteReference[]? _noteReferences;

    private static FileStream OpenRead(string path) => WordDocumentFileStream.OpenShared(path);     // Word / LibreOffice may hold the file open

    public void Dispose() => _reader.Dispose();
}
