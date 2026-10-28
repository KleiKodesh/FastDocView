using System.Text;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// The piece table ([MS-DOC] §2.8.35 Clx / PlcPcd): maps character positions (CP) to file offsets
/// (FC) in the WordDocument stream, piece by piece. Each piece is either 8-bit "compressed" text
/// (cp1252 with Word's special mappings) or UTF-16LE. This is the .doc format's native random access.
/// </summary>
public sealed class DocBinaryPieceTable
{
    public readonly record struct Piece(int CpStart, int CpEnd, uint FileOffset, bool IsCompressed)
    {
        public int Length => CpEnd - CpStart;
        public int BytesPerCharacter => IsCompressed ? 1 : 2;
        public uint FileOffsetOf(int cp) => FileOffset + (uint)((cp - CpStart) * BytesPerCharacter);
    }

    private readonly Piece[] _pieces;
    private readonly Stream _wordDocument;

    public IReadOnlyList<Piece> Pieces => _pieces;
    public int TotalCharacterCount => _pieces.Length == 0 ? 0 : _pieces[_pieces.Length - 1].CpEnd;

    public DocBinaryPieceTable(Stream wordDocument, Stream tableStream, (uint offset, uint length) clx)
    {
        _wordDocument = wordDocument;
        if (clx.length == 0) throw new InvalidDataException("Word document has no piece table (Clx).");
        var bytes = new byte[clx.length];
        tableStream.Position = clx.offset;
        int read = 0;
        while (read < bytes.Length) { int n = tableStream.Read(bytes, read, bytes.Length - read); if (n <= 0) break; read += n; }

        // Clx = Prc* (clxt 0x01: cb + grpprl) followed by one Pcdt (clxt 0x02: lcb + PlcPcd).
        int position = 0;
        while (position < bytes.Length && bytes[position] == 0x01)
        {
            int cb = BitConverter.ToUInt16(bytes, position + 1);
            position += 3 + cb;
        }
        if (position >= bytes.Length || bytes[position] != 0x02) throw new InvalidDataException("Word piece table: Pcdt not found.");
        int plcLength = BitConverter.ToInt32(bytes, position + 1);
        int plcStart = position + 5;
        int pieceCount = (plcLength - 4) / 12;
        _pieces = new Piece[pieceCount];
        int pcdStart = plcStart + 4 * (pieceCount + 1);
        for (int i = 0; i < pieceCount; i++)
        {
            int cpStart = BitConverter.ToInt32(bytes, plcStart + i * 4);
            int cpEnd = BitConverter.ToInt32(bytes, plcStart + (i + 1) * 4);
            uint fc = BitConverter.ToUInt32(bytes, pcdStart + i * 8 + 2);
            bool isCompressed = (fc & 0x40000000) != 0;
            fc &= 0x3FFFFFFF;
            if (isCompressed) fc /= 2;
            _pieces[i] = new Piece(cpStart, cpEnd, fc, isCompressed);
        }
    }

    public int PieceIndexOf(int cp)
    {
        int low = 0, high = _pieces.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (_pieces[middle].CpStart <= cp) low = middle; else high = middle - 1;
        }
        return low;
    }

    public uint FileOffsetOf(int cp) => _pieces[PieceIndexOf(cp)].FileOffsetOf(cp);

    /// <summary>Splits a CP range into (piece, cpStart, cpEnd) segments, each inside one piece.</summary>
    public IEnumerable<(Piece piece, int cpStart, int cpEnd)> SegmentsOf(int cpStart, int cpEnd)
    {
        for (int i = PieceIndexOf(cpStart); i < _pieces.Length && _pieces[i].CpStart < cpEnd; i++)
        {
            Piece piece = _pieces[i];
            yield return (piece, Math.Max(cpStart, piece.CpStart), Math.Min(cpEnd, piece.CpEnd));
        }
    }

    /// <summary>Reads the text of a CP range (control characters included, as Word stores them).</summary>
    public string ReadText(int cpStart, int cpEnd)
    {
        var text = new StringBuilder(Math.Max(0, cpEnd - cpStart));
        foreach ((Piece piece, int start, int end) in SegmentsOf(cpStart, cpEnd))
        {
            int byteCount = (end - start) * piece.BytesPerCharacter;
            var bytes = new byte[byteCount];
            _wordDocument.Position = piece.FileOffsetOf(start);
            int read = 0;
            while (read < byteCount) { int n = _wordDocument.Read(bytes, read, byteCount - read); if (n <= 0) break; read += n; }
            DocBinaryTextDecoder.Append(text, bytes, read, piece.IsCompressed);
        }
        return text.ToString();
    }

    /// <summary>Streams the raw characters of a CP range in chunks, for the paragraph-boundary scan.</summary>
    public void ScanCharacters(int cpStart, int cpEnd, Action<char[], int, int> onChunk, int chunkCharacters = 1 << 16)
    {
        var characters = new char[chunkCharacters];
        var bytes = new byte[chunkCharacters * 2];
        foreach ((Piece piece, int start, int end) in SegmentsOf(cpStart, cpEnd))
        {
            for (int cp = start; cp < end; cp += chunkCharacters)
            {
                int count = Math.Min(chunkCharacters, end - cp);
                int byteCount = count * piece.BytesPerCharacter;
                _wordDocument.Position = piece.FileOffsetOf(cp);
                int read = 0;
                while (read < byteCount) { int n = _wordDocument.Read(bytes, read, byteCount - read); if (n <= 0) break; read += n; }
                int decoded = DocBinaryTextDecoder.Decode(bytes, read, piece.IsCompressed, characters);
                onChunk(characters, decoded, cp);
            }
        }
    }
}
