using System.IO.Compression;
using System.Xml.Linq;
using WordDocumentRandomAccess.DocxHtml;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// A .docx opened for windowed rendering: the (cached or freshly built) random-access index, the
/// small parts loaded whole (styles, numbering), and a block reader over the compressed body.
/// </summary>
public sealed class DocxWindowedDocument : IWordWindowedDocument
{
    private readonly DocxRandomAccessIndex _index;
    private readonly DocxBlockXmlReader _blockReader;
    private readonly DocxStyleCascade _styles;
    private readonly DocxBlockHtmlRenderer _renderer;
    private IReadOnlyList<WordDocumentHeading>? _headings;

    public string FilePath { get; }
    public int BlockCount => _index.BlockCount;
    public bool IsRightToLeft { get; }
    public string Stylesheet { get; }
    public DocxRandomAccessIndex Index => _index;

    public static DocxWindowedDocument Open(string filePath, string? cacheDirectory = null,
                                            IProgress<WordDocumentIndexProgress>? progress = null,
                                            CancellationToken cancellation = default)
    {
        DocxRandomAccessIndex index = DocxRandomAccessIndexBuilder.LoadOrBuild(filePath, cacheDirectory, progress, cancellation);
        return new DocxWindowedDocument(filePath, index);
    }

    public DocxWindowedDocument(string filePath, DocxRandomAccessIndex index)
    {
        FilePath = filePath;
        _index = index;
        using (var file = File.OpenRead(filePath))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Read))
        {
            _styles = new DocxStyleCascade(LoadPart(zip, "word/styles.xml"));
            _renderer = new DocxBlockHtmlRenderer(_styles, new DocxNumberingDefinitions(LoadPart(zip, "word/numbering.xml")));
        }
        IsRightToLeft = _styles.IsDocumentRightToLeft;
        Stylesheet = WordDocumentHtmlStylesheet.Build(_styles.ResolveParagraphStyle(null).RunFormatting, IsRightToLeft);
        _blockReader = new DocxBlockXmlReader(filePath, index);
    }

    public int EstimatedTextLength(int blockIndex) => _index.BlockTextLengths[blockIndex];

    public IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count)
    {
        var result = new List<WordDocumentBlockHtml>(count);
        foreach ((int blockIndex, XElement element) in _blockReader.ReadBlocks(firstBlockIndex, count))
            result.Add(new WordDocumentBlockHtml(blockIndex, _index.BlockKinds[blockIndex], _renderer.Render(element)));
        return result;
    }

    public IReadOnlyList<WordDocumentHeading> Headings => _headings ??= CollectHeadings();

    private List<WordDocumentHeading> CollectHeadings()
    {
        var headings = new List<WordDocumentHeading>();
        var levelByStyleId = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string styleId in _index.StyleIdTable) levelByStyleId[styleId] = _styles.HeadingLevelOf(styleId);

        for (int blockIndex = 0; blockIndex < _index.BlockCount; blockIndex++)
        {
            string? styleId = _index.StyleIdOf(blockIndex);
            if (styleId is null || levelByStyleId[styleId] == 0) continue;
            foreach ((_, XElement element) in _blockReader.ReadBlocks(blockIndex, 1))
            {
                string text = DocxBlockHtmlRenderer.PlainText(element);
                if (text.Length > 0) headings.Add(new WordDocumentHeading(blockIndex, levelByStyleId[styleId], text));
            }
        }
        return headings;
    }

    private static XDocument? LoadPart(ZipArchive zip, string name)
    {
        ZipArchiveEntry? entry = zip.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }

    public void Dispose() => _blockReader.Dispose();
}
