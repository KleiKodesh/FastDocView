using System.Text;
using System.Xml.Linq;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>
/// Reads a run of blocks out of a seekable view of the body using the index: one seek to the first
/// block, one contiguous read to the last block's end, then each block's own byte range (a list
/// paragraph is only its text:p, not the list around it) is wrapped in the root tag and parsed.
/// The view is a deflate random-access stream over content.xml (sidecar mode), the inflated
/// content.xml (in-memory mode), or the .fodt file itself.
/// </summary>
public sealed class OdtBlockXmlReader : IDisposable
{
    private readonly Stream _view;
    private readonly IDisposable? _owned;
    private byte[]? _rootOpenTagBytes;
    private string? _rootElementName;

    /// <summary>The block table read from: the builder while an on-demand index is scanning, then the finished index (same offsets).</summary>
    internal IOdtBlockTable Table { get; set; }

    public static OdtBlockXmlReader OverPackage(Stream package, OdtRandomAccessIndex index)
        => new(DocxEntryInflation.OpenView(package, index.ContentEntry!, index.Checkpoints), index, package);

    public static OdtBlockXmlReader OverBytes(byte[] content, OdtRandomAccessIndex index)
        => new(new MemoryStream(content, writable: false), index, null);

    /// <summary>A .fodt: the file is the view (owned).</summary>
    public static OdtBlockXmlReader OverFlatFile(Stream file, OdtRandomAccessIndex index) => new(file, index, null);

    /// <summary>An on-demand open: content.xml through the builder's live checkpoint list (it seeks anywhere scanned so far), or the .fodt file.</summary>
    internal static OdtBlockXmlReader OverIncrementalBuild(Stream file, OdtIncrementalIndexBuilder builder)
        => builder.ContentEntry is { } content
            ? new(DocxEntryInflation.OpenView(file, content, builder.Checkpoints), builder, file)
            : new(file, builder, null);

    private OdtBlockXmlReader(Stream view, IOdtBlockTable table, IDisposable? owned)
    {
        _view = view;
        Table = table;
        _owned = owned;
    }

    public List<(int blockIndex, XElement element)> ReadBlocks(int firstBlockIndex, int count)
    {
        var result = new List<(int, XElement)>();
        IOdtBlockTable table = Table;
        int last = Math.Min(table.BlockCount, firstBlockIndex + count) - 1;
        if (firstBlockIndex < 0 || firstBlockIndex > last) return result;
        _rootOpenTagBytes ??= Encoding.UTF8.GetBytes(table.RootOpenTag);
        _rootElementName ??= OdtXmlFragment.RootElementName(table.RootOpenTag);
        long rangeStart = table.BlockStartOffset(firstBlockIndex);
        long rangeEnd = table.BlockEndOffset(last);
        byte[] bytes = OdtXmlFragment.ReadRange(_view, rangeStart, rangeEnd);
        for (int blockIndex = firstBlockIndex; blockIndex <= last; blockIndex++)
        {
            int start = (int)(table.BlockStartOffset(blockIndex) - rangeStart);
            int end = (int)(table.BlockEndOffset(blockIndex) - rangeStart);
            result.Add((blockIndex, OdtXmlFragment.Parse(_rootOpenTagBytes, _rootElementName, bytes, start, end - start)));
        }
        return result;
    }

    /// <summary>
    /// The raw bytes of blocks [first, first + count) (clamped to the known blocks) in one contiguous read, and each
    /// block's (start, length) inside them - for the search, which reads text straight from the bytes.
    /// </summary>
    public (byte[] bytes, List<(int blockIndex, int start, int length)> blocks) ReadBlockBytes(int firstBlockIndex, int count)
    {
        var blocks = new List<(int, int, int)>();
        IOdtBlockTable table = Table;
        int last = Math.Min(table.BlockCount, firstBlockIndex + count) - 1;
        if (firstBlockIndex < 0 || firstBlockIndex > last) return ([], blocks);
        long rangeStart = table.BlockStartOffset(firstBlockIndex);
        byte[] bytes = OdtXmlFragment.ReadRange(_view, rangeStart, table.BlockEndOffset(last));
        for (int blockIndex = firstBlockIndex; blockIndex <= last; blockIndex++)
        {
            int start = (int)(table.BlockStartOffset(blockIndex) - rangeStart);
            blocks.Add((blockIndex, start, (int)(table.BlockEndOffset(blockIndex) - rangeStart) - start));
        }
        return (bytes, blocks);
    }

    public void Dispose()
    {
        _view.Dispose();
        _owned?.Dispose();
    }
}
