using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Odt;

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
    private readonly OdtRandomAccessIndex _index;
    private readonly byte[] _rootOpenTagBytes;
    private readonly string _rootElementName;

    public static OdtBlockXmlReader OverPackage(Stream package, OdtRandomAccessIndex index)
        => new(DocxEntryInflation.OpenView(package, index.ContentEntry!, index.Checkpoints), index, package);

    public static OdtBlockXmlReader OverBytes(byte[] content, OdtRandomAccessIndex index)
        => new(new MemoryStream(content, writable: false), index, null);

    /// <summary>A .fodt: the file is the view (owned).</summary>
    public static OdtBlockXmlReader OverFlatFile(Stream file, OdtRandomAccessIndex index) => new(file, index, null);

    private OdtBlockXmlReader(Stream view, OdtRandomAccessIndex index, IDisposable? owned)
    {
        _view = view;
        _index = index;
        _owned = owned;
        _rootOpenTagBytes = Encoding.UTF8.GetBytes(index.RootOpenTag);
        _rootElementName = OdtXmlFragment.RootElementName(index.RootOpenTag);
    }

    public List<(int blockIndex, XElement element)> ReadBlocks(int firstBlockIndex, int count)
    {
        var result = new List<(int, XElement)>();
        int last = Math.Min(_index.BlockCount, firstBlockIndex + count) - 1;
        if (firstBlockIndex < 0 || firstBlockIndex > last) return result;
        long rangeStart = _index.BlockStartOffsets[firstBlockIndex];
        long rangeEnd = _index.BlockEndOffsets[last];
        byte[] bytes = OdtXmlFragment.ReadRange(_view, rangeStart, rangeEnd);
        for (int blockIndex = firstBlockIndex; blockIndex <= last; blockIndex++)
        {
            int start = (int)(_index.BlockStartOffsets[blockIndex] - rangeStart);
            int end = (int)(_index.BlockEndOffsets[blockIndex] - rangeStart);
            result.Add((blockIndex, OdtXmlFragment.Parse(_rootOpenTagBytes, _rootElementName, bytes, start, end - start)));
        }
        return result;
    }

    public void Dispose()
    {
        _view.Dispose();
        _owned?.Dispose();
    }
}
