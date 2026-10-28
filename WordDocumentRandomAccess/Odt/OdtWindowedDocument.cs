using System.Xml.Linq;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// An OpenDocument text (.odt / .ott package, or flat .fodt) opened for windowed rendering, behind
/// the same <see cref="IWordWindowedDocument"/> contract as .docx and .doc.
///   • package, in-memory mode (small): one pass inflates content.xml and keeps it for the session;
///   • package, sidecar mode: content.xml checkpoints + the index are saved; reopen reads no body;
///   • flat: the file is the random-access view; the index (sidecar when big) only saves the pass.
/// </summary>
public sealed class OdtWindowedDocument : IWordWindowedDocument
{
    private readonly OdtRandomAccessIndex _index;
    private readonly OdtBlockXmlReader _reader;
    private readonly OdtStyleCascade _styles;
    private readonly OdtBlockHtmlRenderer _renderer;
    private readonly object _readLock = new();
    private readonly int _footnoteCount, _endnoteCount;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    public OdtRandomAccessIndex Index => _index;

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
        FileStream file = OpenRead(filePath);
        try
        {
            WordIndexDecision decision = policy.Decide(filePath, TextBytes(file, isFlat), cached is not null);
            OdtWindowedDocument document = OpenFromStream(file, filePath, isFlat, decision, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (decision.Mode == WordIndexMode.Sidecar && cached is null)
            {
                try { document._index.Save(indexPath); }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }   // the sidecar is a cache; the open document stands
            }
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
        var indexingStyles = new OdtStyleCascade();
        if (!isFlat) foreach (XElement section in OdtPackage.ReadStylesSections(file)) indexingStyles.AddSection(section);

        if (isFlat)
            return OverIndex(displayPath, file, OdtRandomAccessIndexBuilder.BuildFromFlat(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, cancellation), decision);

        if (decision.Mode == WordIndexMode.InMemory)
        {
            var content = new MemoryStream();
            OdtRandomAccessIndex index = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress, long.MaxValue, content, cancellation);
            file.Dispose();
            return new OdtWindowedDocument(displayPath, index, OdtBlockXmlReader.OverBytes(content.ToArray(), index), RenderingStyles(index), decision);
        }

        OdtRandomAccessIndex built = OdtRandomAccessIndexBuilder.BuildFromPackage(file, indexingStyles, sourceLength, sourceLastWriteTicksUtc, progress,
                                                                                OdtRandomAccessIndexBuilder.DefaultCheckpointSpanBytes, null, cancellation);
        return OverIndex(displayPath, file, built, decision);
    }

    private static OdtWindowedDocument OverIndex(string displayPath, Stream file, OdtRandomAccessIndex index, WordIndexDecision decision)
        => new(displayPath, index, index.IsFlat ? OdtBlockXmlReader.OverFlatFile(file, index) : OdtBlockXmlReader.OverPackage(file, index), RenderingStyles(index), decision);

    /// <summary>The cascade from the index's style sections (no list styles) — a reopen never touches styles.xml.</summary>
    private static OdtStyleCascade RenderingStyles(OdtRandomAccessIndex index)
    {
        var styles = new OdtStyleCascade();
        foreach (string section in index.StyleSections) styles.AddSection(OdtBodyBlockScanner.ParseSection(index.RootOpenTag, section));
        return styles;
    }

    private OdtWindowedDocument(string filePath, OdtRandomAccessIndex index, OdtBlockXmlReader reader, OdtStyleCascade styles, WordIndexDecision decision)
    {
        FilePath = filePath;
        _index = index;
        _reader = reader;
        _styles = styles;
        _renderer = new OdtBlockHtmlRenderer(styles);
        Decision = decision;
        foreach (OdtNoteReference note in index.NoteReferences)
            if (note.Kind == WordDocumentNoteKind.Footnote) _footnoteCount++; else _endnoteCount++;
        IsRightToLeft = styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(styles.DefaultParagraph.Formatting.RunFormatting, IsRightToLeft);
    }

    public int EstimatedTextLength(int blockIndex) => _index.BlockTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        lock (_readLock)
        {
            var result = new List<WordDocumentBlockHtml>(Math.Max(0, count));
            foreach ((int blockIndex, XElement element) in _reader.ReadBlocks(firstBlockIndex, count))
            {
                (string html, IReadOnlyList<WordDocumentNoteHtml> notes) = _renderer.Render(element, _index.ListLabelOf(blockIndex), _index.ListLevelOf(blockIndex),
                                                                                           _index.NoteReferencesInBlocks(blockIndex, blockIndex), null);
                result.Add(new WordDocumentBlockHtml(blockIndex, _index.BlockKinds[blockIndex], html, notes));
            }
            return result;
        }
    }

    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;

    public (int footnotes, int endnotes) NoteCounts => (_footnoteCount, _endnoteCount);

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
        => _noteReferences ??= _index.NoteReferences.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, r.BlockIndex)).ToArray();
    private WordDocumentNoteReference[]? _noteReferences;

    private static FileStream OpenRead(string path) => new(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);

    public void Dispose() => _reader.Dispose();
}
