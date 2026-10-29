using FastDocView.Html;

namespace FastDocView.DocBinary;

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

    /// <summary>
    /// The block holding a bookmark's start: the bookmark tables (SttbfBkmk + PlcfBkf, small, read at open) give its CP and the
    /// block index maps it - nothing stored in the index, no text read. Bookmarks outside the main text (notes, headers) are null.
    /// </summary>
    public int? FindBookmark(string name)
        => Bookmarks.StartOf(name) is { } cp && cp < FileInformationBlock.MainTextCharacterCount && _index.BlockCount > 0 ? _index.BlockAt(cp) : null;

    public DocBinaryBookmarks Bookmarks { get; }

    /// <summary>Page text width per section, from PlcfSed (read at open: one small record per section).</summary>
    public IReadOnlyList<WordDocumentSection> Sections { get; }
    public bool IsIndexComplete => true;
    public int KnownBlockCount => _index.BlockCount;
    public int ScanAhead(int blocks) => BlockCount;          // always complete
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public WordIndexDecision Decision { get; }
    public DocBinaryFileInformationBlock FileInformationBlock { get; }
    public DocBinaryPieceTable PieceTable { get; }
    public DocBinaryFormattedDiskPages FormattedDiskPages { get; }
    public DocBinaryStyleSheet StyleSheet => _styles;
    public DocBinaryListDefinitions ListDefinitions { get; }
    public DocBinaryBlockIndex BlockIndex => _index;

    public static WordIndexDecision Decide(string filePath, WordIndexPolicy policy)
    {
        using var compoundFile = new CompoundFileReader(filePath);
        ThrowWhenEncryptedPackage(compoundFile);
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
        var file = new WordDocumentFileStream(filePath);
        try
        {
            var document = new DocBinaryWindowedDocument(file, filePath, policy, cached, info.Length, info.LastWriteTimeUtc.Ticks, progress, cancellation);
            if (document.Decision.Mode == WordIndexMode.Sidecar && cached is null) document._index.Save(indexPath);
            document._releasableFile = file;
            document._isReadingFromFile = true;                // .doc text is always read through the piece table
            file.Release();                                   // no handle on the file between windows (Word must be able to save it)
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
        ThrowWhenEncryptedPackage(_compoundFile);
        _wordDocument = _compoundFile.OpenRootStream("WordDocument") ?? throw new InvalidDataException("No WordDocument stream — not a Word binary file.");
        FileInformationBlock = new DocBinaryFileInformationBlock(_wordDocument);
        _tableStream = _compoundFile.OpenRootStream(FileInformationBlock.TableStreamName)
                       ?? throw new InvalidDataException($"Table stream {FileInformationBlock.TableStreamName} is missing.");
        PieceTable = new DocBinaryPieceTable(_wordDocument, _tableStream, FileInformationBlock);
        _styles = new DocBinaryStyleSheet(_tableStream, FileInformationBlock.StyleSheet)
        {
            Fonts = DocBinaryFontTable.Read(_tableStream, FileInformationBlock.FontTable),
        };
        _dataStream = _compoundFile.OpenRootStream("Data");
        FormattedDiskPages = new DocBinaryFormattedDiskPages(_wordDocument, _tableStream, FileInformationBlock, _dataStream);
        _notes = new DocBinaryNoteReferences(_tableStream, FileInformationBlock);
        _renderer = new DocBinaryParagraphHtmlRenderer(PieceTable, FormattedDiskPages, _styles, _notes);
        ListDefinitions = DocBinaryListDefinitions.Read(_tableStream, FileInformationBlock.ListFormats, FileInformationBlock.ListFormatOverrides);
        _renderer.Lists = ListDefinitions;
        _renderer.DefaultTabStopTwips = ReadDefaultTabStop(_tableStream, FileInformationBlock.DocumentProperties);
        Bookmarks = DocBinaryBookmarks.Read(_tableStream, FileInformationBlock);
        _renderer.Bookmarks = Bookmarks;

        bool isCacheValid = cachedIndex is not null && cachedIndex.MatchesSource(sourceLength, sourceLastWriteTicksUtc);
        Decision = policy.Decide(displayPath, (long)FileInformationBlock.MainTextCharacterCount * 2, isCacheValid);
        _index = isCacheValid ? cachedIndex!
               : DocBinaryBlockIndex.Build(PieceTable, _renderer, _styles, ListDefinitions, FileInformationBlock.MainTextCharacterCount,
                                           sourceLength, sourceLastWriteTicksUtc, progress, cancellation);

        Sections = DocBinarySectionTable.Read(_wordDocument, _tableStream, FileInformationBlock.SectionDescriptors, _index.BlockAt);
        IsRightToLeft = DetectRightToLeft();
        Stylesheet = WordDocumentHtmlStylesheet.Build(_styles.ResolveParagraphStyle(0).RunFormatting, IsRightToLeft);
    }

    /// <summary>
    /// A password-protected .docx/.docm is an OLE compound file holding EncryptionInfo and
    /// EncryptedPackage streams ([MS-OFFCRYPTO] §2.3.4 / §2.3.4.4), often with a decoy WordDocument stream.
    /// </summary>
    internal static void ThrowWhenEncryptedPackage(CompoundFileReader compoundFile)
    {
        if (compoundFile.HasRootStream("EncryptedPackage") && compoundFile.HasRootStream("EncryptionInfo"))
            throw new NotSupportedException("Encrypted (password-protected) Word 2007+ documents are not supported.");
    }

    /// <summary>Normal style bidi, else a Hebrew/Arabic share of the first few thousand characters.</summary>
    /// <summary>Dop dxaTab: the default tab stop interval, a 16-bit value at offset 10 of the Dop (LibreOffice ww8scan.cxx WW8Dop); 720 when absent.</summary>
    private static int ReadDefaultTabStop(Stream tableStream, (uint offset, uint length) dop)
    {
        if (dop.length < 12 || dop.offset + 12 > tableStream.Length) return WordListLabelHtml.DefaultTabStopTwips;
        tableStream.Position = dop.offset + 10;
        int low = tableStream.ReadByte(), high = tableStream.ReadByte();
        int value = (short)(low | high << 8);
        return value > 0 ? value : WordListLabelHtml.DefaultTabStopTwips;
    }

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

    private WordDocumentFileStream? _releasableFile;   // set by Open(path): released after every window
    private bool _isReadingFromFile;                    // false when the whole text was inflated in memory (small .docx)
    private readonly object _readLock = new();

    public bool HasSourceChanged => _releasableFile?.HasSourceChanged ?? false;

    /// <summary>
    /// Scans blocks [fromBlock, fromBlock + maxBlocksToScan): each block's characters straight from the piece table
    /// (no FKP formatting, no rendering), and with <see cref="WordDocumentSearchOptions.IncludeNotes"/> the bodies of the
    /// notes it references. The index of a .doc is always complete.
    /// </summary>
    public WordDocumentSearchBatch Search(string query, WordDocumentSearchOptions options, int fromBlock, int maxBlocksToScan)
    {
        var matcher = new WordDocumentSearchMatcher(query, options);
        if (matcher.IsEmpty) return WordDocumentSearchBatch.Nothing;
        fromBlock = Math.Max(0, fromBlock);
        lock (_readLock)
        {
            if (_releasableFile is not null && _isReadingFromFile && _releasableFile.HasSourceChanged)
                throw new WordDocumentSourceChangedException(FilePath);
            try
            {
                int end = (int)Math.Min(BlockCount, (long)fromBlock + Math.Max(1, maxBlocksToScan));
                var hits = new List<int>();
                for (int blockIndex = fromBlock; blockIndex < end; blockIndex++)
                {
                    int start = _index.BlockStartCps[blockIndex], blockEnd = _index.BlockEndCp(blockIndex);
                    matcher.Reset();
                    _renderer.FeedSearchText(start, blockEnd, matcher);
                    if (!matcher.Found && options.IncludeNotes)
                        foreach (DocBinaryNoteReferences.NoteReference reference in _notes.InRange(start, blockEnd))
                        {
                            matcher.Separator();
                            _renderer.FeedSearchText(reference.TextStartCp, reference.TextEndCp, matcher);
                            if (matcher.Found) break;
                        }
                    if (matcher.Found) hits.Add(blockIndex);
                }
                return WordDocumentSearchBatch.Finish(hits, Math.Max(end, fromBlock), isIndexComplete: true, BlockCount);
            }
            finally { _releasableFile?.Release(); }
        }
    }

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

    private IReadOnlyList<WordDocumentBlockHtml> RenderBlocksLocked(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        int last = Math.Min(BlockCount, firstBlockIndex + count);
        for (int blockIndex = Math.Max(0, firstBlockIndex); blockIndex < last; blockIndex++)
        {
            int start = _index.BlockStartCps[blockIndex], end = _index.BlockEndCp(blockIndex);
            WordDocumentBlockKind kind = _index.BlockKinds[blockIndex];
            string html = kind == WordDocumentBlockKind.TableRow
                ? _renderer.RenderRow(start, end)
                : _renderer.Render(start, end, PreviousStyleKey(blockIndex), listLabel: _index.ListLabelOf(blockIndex));
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
