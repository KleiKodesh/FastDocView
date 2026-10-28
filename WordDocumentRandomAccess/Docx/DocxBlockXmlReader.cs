using System.Text;
using System.Xml.Linq;
using SeekableDeflate;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Reads the XML of a run of body-level blocks out of a seekable view of document.xml using the
/// index: one seek to the first block, one contiguous read to the last block's end, then each
/// block's bytes are wrapped in the root element (for its namespace declarations) and parsed.
/// The view is a <see cref="DeflateRandomAccessStream"/> over the compressed entry (sidecar mode),
/// a <see cref="ZipStoredEntryStream"/> (stored entry) or a <see cref="MemoryStream"/> (in-memory mode).
/// </summary>
public sealed class DocxBlockXmlReader : IDisposable
{
    private readonly Stream _documentXml;
    private readonly IDisposable? _owned;
    private byte[]? _rootOpenTagBytes;            // resolved on first read: an incremental table knows it only after a block
    private string? _rootElementName;

    /// <summary>Where block ranges come from: the finished index, or the incremental builder's known blocks (swapped for the index when it completes).</summary>
    internal IDocxBlockTable Table { get; set; }

    /// <summary>The package stream in sidecar mode (the note reader shares it), null in in-memory mode.</summary>
    public Stream? Package => _owned as Stream;

    /// <summary>Sidecar mode over the package stream (which this reader then owns).</summary>
    public static DocxBlockXmlReader OverPackage(Stream package, DocxRandomAccessIndex index)
    {
        ZipEntryLocation entry = index.DocumentEntry;
        Stream view = DocxEntryInflation.OpenView(package, entry, index.Checkpoints);
        return new DocxBlockXmlReader(view, index, package);
    }

    /// <summary>
    /// Incremental first open over the package stream (owned): the view seeks through the builder's
    /// checkpoint list, which grows as the scan advances.
    /// </summary>
    internal static DocxBlockXmlReader OverIncrementalBuild(Stream package, DocxIncrementalIndexBuilder builder)
        => new(DocxEntryInflation.OpenView(package, builder.BodyEntry, builder.BodyCheckpoints), builder, package);

    /// <summary>In-memory mode over the already-inflated body.</summary>
    public static DocxBlockXmlReader OverBody(byte[] body, DocxRandomAccessIndex index)
        => new(new MemoryStream(body, writable: false), index, null);

    private DocxBlockXmlReader(Stream documentXml, IDocxBlockTable table, IDisposable? owned)
    {
        _documentXml = documentXml;
        Table = table;
        _owned = owned;
    }

    /// <summary>Parsed body children for blocks [first, first + count), clamped to the blocks the table knows.</summary>
    public List<(int blockIndex, XElement element)> ReadBlocks(int firstBlockIndex, int count)
    {
        var result = new List<(int, XElement)>();
        IDocxBlockTable table = Table;
        int last = Math.Min(table.BlockCount, firstBlockIndex + count) - 1;
        if (firstBlockIndex < 0 || firstBlockIndex > last) return result;
        if (_rootOpenTagBytes is null)
        {
            _rootOpenTagBytes = Encoding.UTF8.GetBytes(table.RootOpenTag);
            _rootElementName = DocxXmlFragment.RootElementName(table.RootOpenTag);
        }

        long rangeStart = table.BlockStartOffset(firstBlockIndex);
        long rangeEnd = table.BlockEndOffset(last);
        byte[] bytes = DocxXmlFragment.ReadRange(_documentXml, rangeStart, rangeEnd);

        for (int blockIndex = firstBlockIndex; blockIndex <= last; blockIndex++)
        {
            int start = (int)(table.BlockStartOffset(blockIndex) - rangeStart);
            int end = (int)(table.BlockEndOffset(blockIndex) - rangeStart);
            result.Add((blockIndex, DocxXmlFragment.Parse(_rootOpenTagBytes, _rootElementName!, bytes, start, end - start, $"Body block {blockIndex}")));
        }
        return result;
    }

    public void Dispose()
    {
        _documentXml.Dispose();
        _owned?.Dispose();
    }
}