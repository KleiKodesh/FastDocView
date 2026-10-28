using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// A binary .doc opened for windowed rendering. The piece table already gives random access to the
/// text; the only pass is the block scan (paragraphs and table rows, headings), kept in memory for
/// small files and saved as a sidecar for big ones (policy). Everything else — PAPX, CHPX, styles,
/// fonts, note bodies — is looked up on demand. <see cref="OpenFromStream"/> is the host-agnostic (WASM) entry.
/// </summary>
public sealed class DocBinaryWindowedDocument : IWordWindowedDocument
{
    private readonly CompoundFileReader _compoundFile;
    private readonly Stream _wordDocument;
    private readonly Stream _tableStream;
    private readonly Stream? _dataStream;
    private readonly DocBinaryBlockIndex _index;
    private readonly DocBinaryParagraphHtmlRenderer _renderer;
    private readonly DocBinaryStyleSheet _styles;
    private readonly DocBinaryNoteReferences _notes;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    public DocBinaryFileInformationBlock FileInformationBlock { get; }
    public DocBinaryPieceTable PieceTable { get; }
    public DocBinaryFormattedDiskPages FormattedDiskPages { get; }
    public DocBinaryStyleSheet StyleSheet => _styles;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy)
    {
        using var compoundFile = new CompoundFileReader(filePath);
        using Stream wordDocument = compoundFile.OpenRootStream("WordDocument") ?? throw new InvalidDataException("No WordDocument stream.");
        var fib = new DocBinaryFileInformationBlock(wordDocument);
        bool hasCachedIndex = DocBinaryBlockIndex.TryLoad(policy.IndexPathFor(filePath), filePath) is not null;
        return policy.Decide(filePath, (long)fib.MainTextCharacterCount * 2, hasCachedIndex);
    }

    public static DocBinaryWindowedDocument Open(string filePath, WordIndexPolicy? policy = null,
                                                 IProgress<WordDocumentIndexProgress>? progress = null,
                                                 CancellationToken cancellation = default)
    {
        policy ??= WordIndexPolicy.Default;
        var info = new FileInfo(filePath);
        string indexPath = policy.IndexPathFor(filePath);
        DocBinaryBlockIndex? cached = DocBinaryBlockIndex.TryLoad(indexPath, filePath);
        var file = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        try
        {
            var document = new DocBinaryWindowedDocument(file, filePath, policy, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (document.Decision.Mode == WordIndexMode.Sidecar && cached is null) document._index.Save(indexPath);
            return document;
        }
        catch
        {
            file.Dispose();
            throw;
        }
    }

    /// <summary>Host-agnostic entry: the stream stays open for the document's lifetime and is disposed with it.</summary>
    public static DocBinaryWindowedDocument OpenFromStream(Stream file, string displayPath, WordIndexPolicy policy,
                                                           DocBinaryBlockIndex? cachedIndex, long sourceLength, long sourceLastWriteTicksUtc,
                                                           IProgress<WordDocumentIndexProgress>? progress = null,
                                                           CancellationToken cancellation = default)
        => new(file, displayPath, policy, cachedIndex, sourceLength, sourceLastWriteTicksUtc, progress, cancellation);

    private DocBinaryWindowedDocument(Stream file, string displayPath, WordIndexPolicy policy, DocBinaryBlockIndex? cachedIndex,
                                      long sourceLength, long sourceLastWriteTicksUtc,
                                      IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        FilePath = displayPath;
        _compoundFile = new CompoundFileReader(file);
        _wordDocument = _compoundFile.OpenRootStream("WordDocument") ?? throw new InvalidDataException("No WordDocument stream — not a Word binary file.");
        FileInformationBlock = new DocBinaryFileInformationBlock(_wordDocument);
        _tableStream = _compoundFile.OpenRootStream(FileInformationBlock.TableStreamName)
                       ?? throw new InvalidDataException($"Table stream {FileInformationBlock.TableStreamName} is missing.");
        PieceTable = new DocBinaryPieceTable(_wordDocument, _tableStream, FileInformationBlock.PieceTable);
        _styles = new DocBinaryStyleSheet(_tableStream, FileInformationBlock.StyleSheet)
        {
            Fonts = DocBinaryFontTable.Read(_tableStream, FileInformationBlock.FontTable),
        };
        _dataStream = _compoundFile.OpenRootStream("Data");
        FormattedDiskPages = new DocBinaryFormattedDiskPages(_wordDocument, _tableStream, FileInformationBlock, _dataStream);
        _notes = new DocBinaryNoteReferences(_tableStream, FileInformationBlock);
        _renderer = new DocBinaryParagraphHtmlRenderer(PieceTable, FormattedDiskPages, _styles, _notes);

        bool isCacheValid = cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc);
        Decision = policy.Decide(displayPath, (long)FileInformationBlock.MainTextCharacterCount * 2, isCacheValid);
        _index = isCacheValid ? cachedIndex!
               : DocBinaryBlockIndex.Build(PieceTable, _renderer, _styles, FileInformationBlock.MainTextCharacterCount,
                                           sourceLength, sourceLastWriteTicksUtc, progress, cancellation);

        IsRightToLeft = DetectRightToLeft();
        Stylesheet = WordDocumentHtmlStylesheet.Build(_styles.ResolveParagraphStyle(0).RunFormatting, IsRightToLeft);
    }

    /// <summary>Normal style bidi, else a Hebrew/Arabic share of the first few thousand characters.</summary>
    private bool DetectRightToLeft()
    {
        if (_styles.ResolveParagraphStyle(0).IsRightToLeft) return true;
        string sample = PieceTable.ReadText(0, Math.Min(4000, FileInformationBlock.MainTextCharacterCount));
        int rightToLeft = 0, letters = 0;
        foreach (char c in sample)
        {
            if (!char.IsLetter(c)) continue;
            letters++;
            if (c is >= '֐' and <= 'ࣿ') rightToLeft++;
        }
        return letters > 0 && rightToLeft * 5 > letters;
    }

    public int EstimatedTextLength(int blockIndex) => _index.BlockTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        int last = Math.Min(BlockCount, firstBlockIndex + count);
        for (int blockIndex = Math.Max(0, firstBlockIndex); blockIndex < last; blockIndex++)
        {
            int start = _index.BlockStartCps[blockIndex], end = _index.BlockEndCp(blockIndex);
            WordDocumentBlockKind kind = _index.BlockKinds[blockIndex];
            string html = kind == WordDocumentBlockKind.TableRow
                ? _renderer.RenderRow(start, end)
                : _renderer.Render(start, end, PreviousStyleKey(blockIndex));
            result.Add(new WordDocumentBlockHtml(blockIndex, kind, html, RenderNotesIn(start, end)));
        }
        return result;
    }

    /// <summary>The style of the paragraph before a block, for contextual spacing (one cached PAPX lookup).</summary>
    private string? PreviousStyleKey(int blockIndex)
    {
        if (blockIndex == 0 || _index.BlockKinds[blockIndex - 1] != WordDocumentBlockKind.Paragraph) return null;
        return _renderer.ResolveParagraph(_index.BlockStartCps[blockIndex - 1], _index.BlockEndCp(blockIndex - 1)).formatting.StyleKey;
    }

    private IReadOnlyList<WordDocumentNoteHtml> RenderNotesIn(int cpStart, int cpEnd)
    {
        ReadOnlySpan<DocBinaryNoteReferences.NoteReference> references = _notes.InRange(cpStart, cpEnd);
        if (references.IsEmpty) return [];
        var notes = new List<WordDocumentNoteHtml>(references.Length);
        foreach (DocBinaryNoteReferences.NoteReference reference in references)
        {
            string label = DocBinaryParagraphHtmlRenderer.NoteLabel(reference);
            notes.Add(new WordDocumentNoteHtml(reference.Kind, reference.Number, label,
                                               _renderer.RenderNote(reference.TextStartCp, reference.TextEndCp, label)));
        }
        return notes;
    }

    /// <summary>From the index — captured during the block scan, so a reopen reads no text for the TOC.</summary>
    public IReadOnlyList<WordDocumentHeading> Headings => _index.Headings;

    public (int footnotes, int endnotes) NoteCounts => (_notes.FootnoteCount, _notes.EndnoteCount);

    public IReadOnlyList<WordDocumentNoteReference> NoteReferences
        => _noteReferences ??= _notes.References.Select(r => new WordDocumentNoteReference(r.Kind, r.Number, _index.BlockAt(r.ReferenceCp))).ToArray();
    private WordDocumentNoteReference[]? _noteReferences;

    public void Dispose()
    {
        _dataStream?.Dispose();
        _tableStream.Dispose();
        _wordDocument.Dispose();
        _compoundFile.Dispose();
    }
}
