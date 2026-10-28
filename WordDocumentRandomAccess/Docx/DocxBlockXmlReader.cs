using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

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
    private readonly DocxRandomAccessIndex _index;
    private readonly byte[] _rootOpenTagBytes;
    private static readonly byte[] RootCloseTagBytes = Encoding.UTF8.GetBytes("</w:document>");

    /// <summary>Sidecar mode over the package stream (which this reader then owns).</summary>
    public static DocxBlockXmlReader OverPackage(Stream package, DocxRandomAccessIndex index)
    {
        ZipEntryLocation entry = index.DocumentEntry;
        Stream view = entry.IsDeflate
            ? new DeflateRandomAccessStream(package, entry.DataOffset, entry.CompressedSize, entry.UncompressedSize, index.Checkpoints)
            : new ZipStoredEntryStream(package, entry.DataOffset, entry.CompressedSize);
        return new DocxBlockXmlReader(view, index, package);
    }

    /// <summary>In-memory mode over the already-inflated body.</summary>
    public static DocxBlockXmlReader OverBody(byte[] body, DocxRandomAccessIndex index)
        => new(new MemoryStream(body, writable: false), index, null);

    private DocxBlockXmlReader(Stream documentXml, DocxRandomAccessIndex index, IDisposable? owned)
    {
        _documentXml = documentXml;
        _index = index;
        _owned = owned;
        _rootOpenTagBytes = Encoding.UTF8.GetBytes(index.RootOpenTag);
    }

    /// <summary>Parsed body children for blocks [first, first + count), clamped to the document.</summary>
    public List<(int blockIndex, XElement element)> ReadBlocks(int firstBlockIndex, int count)
    {
        var result = new List<(int, XElement)>();
        int last = Math.Min(_index.BlockCount, firstBlockIndex + count) - 1;
        if (firstBlockIndex < 0 || firstBlockIndex > last) return result;

        long rangeStart = _index.BlockStartOffsets[firstBlockIndex];
        long rangeEnd = _index.BlockEndOffset(last);
        byte[] bytes = ReadRange(rangeStart, rangeEnd);

        for (int blockIndex = firstBlockIndex; blockIndex <= last; blockIndex++)
        {
            int start = (int)(_index.BlockStartOffsets[blockIndex] - rangeStart);
            int end = (int)(_index.BlockEndOffset(blockIndex) - rangeStart);
            result.Add((blockIndex, ParseWrapped(bytes, start, end - start)));
        }
        return result;
    }

    private byte[] ReadRange(long start, long end)
    {
        var bytes = new byte[end - start];
        _documentXml.Position = start;
        int filled = 0;
        while (filled < bytes.Length)
        {
            int read = _documentXml.Read(bytes, filled, bytes.Length - filled);
            if (read <= 0) throw new EndOfStreamException("document.xml ended inside a block.");
            filled += read;
        }
        return bytes;
    }

    private XElement ParseWrapped(byte[] bytes, int start, int length)
    {
        ArraySegment<byte> block = DocxDrawingStripper.Strip(bytes, start, length);
        using var wrapped = new MemoryStream(_rootOpenTagBytes.Length + block.Count + RootCloseTagBytes.Length);
        wrapped.Write(_rootOpenTagBytes, 0, _rootOpenTagBytes.Length);
        wrapped.Write(block.Array!, block.Offset, block.Count);
        wrapped.Write(RootCloseTagBytes, 0, RootCloseTagBytes.Length);
        wrapped.Position = 0;
        XElement root = XElement.Load(wrapped);
        return root.Elements().FirstOrDefault() ?? new XElement("empty");
    }

    public void Dispose()
    {
        _documentXml.Dispose();
        _owned?.Dispose();
    }
}
