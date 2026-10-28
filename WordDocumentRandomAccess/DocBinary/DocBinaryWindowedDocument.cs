using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// A binary .doc opened for windowed rendering. The piece table already gives random access to the
/// text; the only pass is the paragraph-boundary scan (cached as a sidecar). Everything else — PAPX,
/// CHPX, styles — is looked up on demand through the bin tables.
/// </summary>
public sealed class DocBinaryWindowedDocument : IWordWindowedDocument
{
    private readonly CompoundFileReader _compoundFile;
    private readonly Stream _wordDocument;
    private readonly Stream _tableStream;
    private readonly DocBinaryParagraphIndex _index;
    private readonly DocBinaryParagraphHtmlRenderer _renderer;
    private readonly DocBinaryStyleSheet _styles;
    private IReadOnlyList<WordDocumentHeading>? _headings;

    public string FilePath { get; }
    public int BlockCount => _index.ParagraphCount;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public DocBinaryFileInformationBlock FileInformationBlock { get; }
    public DocBinaryPieceTable PieceTable { get; }

    public static DocBinaryWindowedDocument Open(string filePath, string? cacheDirectory = null,
                                                 IProgress<WordDocumentIndexProgress>? progress = null,
                                                 CancellationToken cancellation = default)
        => new(filePath, cacheDirectory, progress, cancellation);

    private DocBinaryWindowedDocument(string filePath, string? cacheDirectory, IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        FilePath = filePath;
        _compoundFile = new CompoundFileReader(filePath);
        _wordDocument = _compoundFile.OpenRootStream("WordDocument") ?? throw new InvalidDataException("No WordDocument stream — not a Word binary file.");
        FileInformationBlock = new DocBinaryFileInformationBlock(_wordDocument);
        _tableStream = _compoundFile.OpenRootStream(FileInformationBlock.TableStreamName)
                       ?? throw new InvalidDataException($"Table stream {FileInformationBlock.TableStreamName} is missing.");
        PieceTable = new DocBinaryPieceTable(_wordDocument, _tableStream, FileInformationBlock.PieceTable);
        _styles = new DocBinaryStyleSheet(_tableStream, FileInformationBlock.StyleSheet);
        var pages = new DocBinaryFormattedDiskPages(_wordDocument, _tableStream, FileInformationBlock);
        _renderer = new DocBinaryParagraphHtmlRenderer(PieceTable, pages, _styles);

        string indexPath = WordDocumentIndexCacheLocation.IndexPathFor(filePath, cacheDirectory);
        DocBinaryParagraphIndex? cached = DocBinaryParagraphIndex.TryLoad(indexPath, filePath);
        if (cached is null)
        {
            cached = DocBinaryParagraphIndex.Build(filePath, PieceTable, FileInformationBlock.MainTextCharacterCount, progress, cancellation);
            cached.Save(indexPath);
        }
        _index = cached;

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

    public int EstimatedTextLength(int blockIndex) => _index.ParagraphTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        int last = Math.Min(BlockCount, firstBlockIndex + count);
        for (int blockIndex = Math.Max(0, firstBlockIndex); blockIndex < last; blockIndex++)
            result.Add(new WordDocumentBlockHtml(blockIndex, WordDocumentBlockKind.Paragraph,
                                                 _renderer.Render(_index.ParagraphStartCps[blockIndex], _index.ParagraphEndCp(blockIndex))));
        return result;
    }

    public IReadOnlyList<WordDocumentHeading> Headings => _headings ??= CollectHeadings();

    private List<WordDocumentHeading> CollectHeadings()
    {
        // Heading detection needs each paragraph's PAPX; that is one page lookup per paragraph,
        // cached per 512-byte page, so a full pass stays cheap even for large documents.
        var headings = new List<WordDocumentHeading>();
        for (int blockIndex = 0; blockIndex < BlockCount; blockIndex++)
        {
            int cpStart = _index.ParagraphStartCps[blockIndex], cpEnd = _index.ParagraphEndCp(blockIndex);
            (_, int styleIndex) = _renderer.ResolveParagraph(cpStart, cpEnd);
            int level = _styles.HeadingLevelOf(styleIndex);
            if (level == 0) continue;
            string text = _renderer.PlainText(cpStart, cpEnd);
            if (text.Length > 0) headings.Add(new WordDocumentHeading(blockIndex, level, text));
        }
        return headings;
    }

    public void Dispose()
    {
        _tableStream.Dispose();
        _wordDocument.Dispose();
        _compoundFile.Dispose();
    }
}
