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
    private readonly DocxNoteBodies _notes;
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
        using FileStream file = OpenRead(filePath);
        long bodyBytes = DocxRandomAccessIndexBuilder.LocateDocumentEntry(file).UncompressedSize;
        bool hasCachedIndex = DocxRandomAccessIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, bodyBytes, hasCachedIndex);
    }

    public static DocxWindowedDocument Open(string filePath, WordIndexPolicy? policy = null,
                                            IProgress<WordDocumentIndexProgress>? progress = null,
                                            CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        DocxRandomAccessIndex? cached = DocxRandomAccessIndex.TryLoad(indexPath, filePath);
        FileStream package = OpenRead(filePath);
        try
        {
            long bodyBytes = cached?.DocumentEntry.UncompressedSize ?? DocxRandomAccessIndexBuilder.LocateDocumentEntry(package).UncompressedSize;
            WordIndexDecision decision = policy.Decide(filePath, bodyBytes, cached is not null);
            DocxWindowedDocument document = OpenFromStream(package, filePath, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (decision.Mode == WordIndexMode.Sidecar && cached is null) document._index.Save(indexPath);
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
            return new DocxWindowedDocument(displayPath, cachedIndex, DocxBlockXmlReader.OverPackage(package, cachedIndex), parts, decision);

        if (decision.Mode == WordIndexMode.InMemory)
        {
            (DocxRandomAccessIndex index, byte[] body) = DocxRandomAccessIndexBuilder.BuildInMemory(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);
            package.Dispose();
            return new DocxWindowedDocument(displayPath, index, DocxBlockXmlReader.OverBody(body, index), parts, decision);
        }

        DocxRandomAccessIndex built = DocxRandomAccessIndexBuilder.Build(package, parts, sourceLength, sourceLastWriteTicksUtc, progress, cancellation: cancellation);
        return new DocxWindowedDocument(displayPath, built, DocxBlockXmlReader.OverPackage(package, built), parts, decision);
    }

    private DocxWindowedDocument(string filePath, DocxRandomAccessIndex index, DocxBlockXmlReader blockReader, DocxDocumentParts parts, WordIndexDecision decision)
    {
        FilePath = filePath;
        _index = index;
        _blockReader = blockReader;
        _styles = parts.Styles;
        _notes = parts.Notes;
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
            string label = reference.Number > 0 ? _notes.FormatNumber(reference.Kind, reference.Number) : "*";
            XElement? body = _notes.Find(reference.Kind, reference.NoteId);
            string html = body is null ? "" : _renderer.RenderNote(body, reference.Kind, label);
            notes.Add(new WordDocumentNoteHtml(reference.Kind, reference.Number, label, html));
        }
        return notes;
    }

    private (int number, string label)? ResolveNoteLabel(WordDocumentNoteKind kind, int noteId)
        => _noteNumbers.TryGetValue((kind, noteId), out int number) ? (number, number > 0 ? _notes.FormatNumber(kind, number) : "*") : null;

    /// <summary>From the index — captured during the scan, so a reopen reads no blocks for the TOC.</summary>
    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;

    public (int footnotes, int endnotes) NoteCounts => (_footnoteCount, _endnoteCount);

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
        => _noteReferences ??= _index.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
    private WordDocumentNoteReference[]? _noteReferences;

    private static FileStream OpenRead(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);

    public void Dispose() => _blockReader.Dispose();
}
