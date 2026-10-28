using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Deflate;
using WordDocumentRandomAccess.Zip;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Reads the XML of a run of body-level blocks straight out of the compressed document using the
/// index: one seek to the first block, one contiguous read to the last block's end, then each
/// block's bytes are wrapped in the root element (for its namespace declarations) and parsed.
/// Keeps the file and the inflater open so consecutive requests continue without re-seeking.
/// </summary>
public sealed class DocxBlockXmlReader : IDisposable
{
    private readonly FileStream _file;
    private readonly Stream _documentXml;
    private readonly DocxRandomAccessIndex _index;
    private readonly byte[] _rootOpenTagBytes;
    private static readonly byte[] RootCloseTagBytes = Encoding.UTF8.GetBytes("</w:document>");

    public DocxBlockXmlReader(string documentPath, DocxRandomAccessIndex index)
    {
        _index = index;
        _file = new FileStream(documentPath, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
        ZipEntryLocation entry = index.DocumentEntry;
        _documentXml = entry.IsDeflate
            ? new DeflateRandomAccessStream(_file, entry.DataOffset, entry.CompressedSize, entry.UncompressedSize, index.Checkpoints)
            : new ZipStoredEntryStream(_file, entry.DataOffset, entry.CompressedSize);
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
        using var wrapped = new MemoryStream(_rootOpenTagBytes.Length + length + RootCloseTagBytes.Length);
        wrapped.Write(_rootOpenTagBytes, 0, _rootOpenTagBytes.Length);
        wrapped.Write(bytes, start, length);
        wrapped.Write(RootCloseTagBytes, 0, RootCloseTagBytes.Length);
        wrapped.Position = 0;
        XElement root = XElement.Load(wrapped);
        return root.Elements().FirstOrDefault() ?? new XElement("empty");
    }

    public void Dispose()
    {
        _documentXml.Dispose();
        _file.Dispose();
    }
}
