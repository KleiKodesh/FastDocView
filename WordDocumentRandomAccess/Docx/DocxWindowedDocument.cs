using System.Xml.Linq;
using WordDocumentRandomAccess.DocxHtml;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// A .docx opened for windowed rendering. Two ways in:
///   • <see cref="Open"/> — file-system hosts: the policy picks in-memory vs sidecar, the sidecar
///     is cached next to the other indexes;
///   • <see cref="OpenFromStream"/> — any host (WASM): the caller supplies the package stream and
///     an optional previously saved index; nothing touches a file system.
/// </summary>
public sealed class DocxWindowedDocument : IWordWindowedDocument
{
    private readonly DocxRandomAccessIndex _index;
    private readonly DocxBlockXmlReader _blockReader;
    private readonly DocxStyleCascade _styles;
    private readonly DocxBlockHtmlRenderer _renderer;
    private readonly DocxNoteXmlReader _noteReader;
    private readonly DocxNoteNumberFormats _noteFormats;
    private WordDocumentFileStream? _releasableFile;   // set by Open(path): released after every window
    private bool _isReadingFromFile;                    // false when the whole text was inflated in memory (small .docx)
    private readonly object _readLock = new();          // one package stream is shared by the block and note readers
    private readonly Dictionary<(WordDocumentNoteKind kind, int id), string> _noteHtmlCache = new();
    private long _noteHtmlCacheCharacters;
    private const long NoteHtmlCacheCharacterLimit = 2_000_000;   // bounded: a few MB of rendered notes, never the part
    private readonly Dictionary<(WordDocumentNoteKind kind, int id), int> _noteNumbers = new();
    private readonly int _footnoteCount;
    private readonly int _endnoteCount;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    public DocxRandomAccessIndex Index => _index;

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
            DocxWindowedDocument document = OpenFromStream(package, filePath, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (decision.Mode == WordIndexMode.Sidecar && cached is null) document._index.Save(indexPath);
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
    /// </summary>
    public static DocxWindowedDocument OpenFromStream(Stream package, string displayPath, WordIndexDecision decision,
                                                      DocxRandomAccessIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                      IProgress<WordDocumentIndexProgress>? progress = null,
                                                      CancellationToken cancellation = default)
    {
        var parts = new DocxDocumentParts(package);
        if (cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc))
            return OverPackage(displayPath, package, cachedIndex, parts, decision);

        if (decision.Mode == WordIndexMode.InMemory)
        {
            (DocxRandomAccessIndex index, DocxRandomAccessIndexBuilder.InMemoryParts inflated) =
                DocxRandomAccessIndexBuilder.BuildInMemory(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
            package.Dispose();
            return new DocxWindowedDocument(displayPath, index, DocxBlockXmlReader.OverBody(inflated.Body, index),
                                            DocxNoteXmlReader.OverBytes(inflated.Footnotes, index.FootnotesPart, inflated.Endnotes, index.EndnotesPart),
                                            parts, decision);
        }

        DocxRandomAccessIndex built = DocxRandomAccessIndexBuilder.Build(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation: cancellation);
        return OverPackage(displayPath, package, built, parts, decision);
    }

    private static DocxWindowedDocument OverPackage(string displayPath, Stream package, DocxRandomAccessIndex index, DocxDocumentParts parts, WordIndexDecision decision)
        => new(displayPath, index, DocxBlockXmlReader.OverPackage(package, index),
               DocxNoteXmlReader.OverPackage(package, index.FootnotesPart, index.EndnotesPart), parts, decision);

    private DocxWindowedDocument(string filePath, DocxRandomAccessIndex index, DocxBlockXmlReader blockReader, DocxNoteXmlReader noteReader,
                                 DocxDocumentParts parts, WordIndexDecision decision)
    {
        FilePath = filePath;
        _index = index;
        _blockReader = blockReader;
        _noteReader = noteReader;
        _styles = parts.Styles;
        _noteFormats = parts.NoteNumberFormats;
        foreach (DocxNoteReference reference in index.NoteReferences)
        {
            _noteNumbers[(reference.Kind, reference.NoteId)] = reference.Number;
            if (reference.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
        }
        _renderer = new DocxBlockHtmlRenderer(_styles, ResolveNoteLabel);
        Decision = decision;
        IsRightToLeft = _styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(_styles.ResolveParagraphStyle(null).RunFormatting, IsRightToLeft);
    }

    public int EstimatedTextLength(int blockIndex) => _index.BlockTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            // Only documents that still read the file (index mode) are at risk; a stale in-memory body is still consistent.
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try { return RenderBlocksLocked(firstBlockIndex, count); }
            finally { _releasableFile?.Release(); }
        }
    }

    private List<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        foreach ((int blockIndex, XElement element) in _blockReader.ReadBlocks(firstBlockIndex, count))
        {
            string? previousStyleKey = blockIndex > 0 ? _index.StyleIdOf(blockIndex - 1) ?? "" : null;
            string html = _renderer.Render(element, _index.ListLabelOf(blockIndex), previousStyleKey);
            result.Add(new WordDocumentBlockHtml(blockIndex, _index.BlockKinds[blockIndex], html, RenderNotesOf(blockIndex)));
        }
        return result;
    }

    /// <summary>The notes a block references, rendered — what a Word-style notes pane shows for the visible text.</summary>
    private IReadOnlyList<WordDocumentNoteHtml> RenderNotesOf(int blockIndex)
    {
        ReadOnlySpan<DocxNoteReference> references = _index.NoteReferencesInBlocks(blockIndex, blockIndex);
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

    /// <summary>From the index — captured during the scan, so a reopen reads no blocks for the TOC.</summary>
    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;

    public (int footnotes, int endnotes) NoteCounts => (_footnoteCount, _endnoteCount);

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
        => _noteReferences ??= _index.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
    private WordDocumentNoteReference[]? _noteReferences;


    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;

    public void Dispose()
    {
        _noteReader.Dispose();
        _blockReader.Dispose();
    }
}
