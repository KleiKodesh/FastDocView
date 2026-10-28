using System.Text;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// The piece table ([MS-DOC] §2.8.35 Clx / PlcPcd): maps character positions (CP) to file offsets
/// (FC) in the WordDocument stream, piece by piece. Each piece is either 8-bit "compressed" text
/// (cp1252 with Word's special mappings) or UTF-16LE. This is the .doc format's native random access.
/// </summary>
public sealed class DocBinaryPieceTable
{
    /// <param name="PrmGrpprl">
    /// The piece's Prm ([MS-DOC] §2.9.214) as a grpprl, applied after the PAPX/CHPX of every paragraph
    /// and run inside the piece (§2.4.6.1 / §2.4.6.2): a Prm1's Prc grpprl from the Clx, or a Prm0's
    /// single compressed sprm expanded to opcode + 1-byte operand. Null when the piece has none.
    /// </param>
    public readonly record struct Piece(int CpStart, int CpEnd, uint FileOffset, bool IsCompressed, byte[]? PrmGrpprl = null)
    {
        public int Length => CpEnd - CpStart;
        public int BytesPerCharacter => IsCompressed ? 1 : 2;
        public uint FileOffsetOf(int cp) => FileOffset + (uint)((cp - CpStart) * BytesPerCharacter);
    }

    private readonly Piece[] _pieces;
    private readonly Stream _wordDocument;

    public IReadOnlyList<Piece> Pieces => _pieces;
    public int TotalCharacterCount => _pieces.Length == 0 ? 0 : _pieces[_pieces.Length - 1].CpEnd;

    public DocBinaryPieceTable(Stream wordDocument, Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        _wordDocument = wordDocument;
        (uint offset, uint length) clx = fib.PieceTable;
        if (clx.length == 0)
        {
            // A non-complex file has no Clx: the text is one contiguous run from fcMin ([MS-DOC] §2.5.1; LibreOffice ww8scan does the same).
            _pieces = [ContiguousPiece(fib)];
            return;
        }
        var bytes = new byte[clx.length];
        tableStream.Position = clx.offset;
        int read = 0;
        while (read < bytes.Length) { int n = tableStream.Read(bytes, read, bytes.Length - read); if (n <= 0) break; read += n; }

        // Clx = Prc* (clxt 0x01: PrcData = cbGrpprl + grpprl, §2.9.209/§2.9.210) followed by one Pcdt (clxt 0x02: lcb + PlcPcd).
        // The Prc grpprls are kept: a Prm1 names one by index (fast-saved files).
        var prcGrpprls = new List<byte[]>();
        int position = 0;
        while (position + 3 <= bytes.Length && bytes[position] == 0x01)
        {
            int cb = BitConverter.ToUInt16(bytes, position + 1);
            int available = Math.Max(0, Math.Min(cb, bytes.Length - position - 3));
            var grpprl = new byte[available];
            Buffer.BlockCopy(bytes, position + 3, grpprl, 0, available);
            prcGrpprls.Add(grpprl);
            position += 3 + cb;
        }
        PrcGrpprls = prcGrpprls;
        if (position >= bytes.Length || bytes[position] != 0x02 || !IsPlausiblePcdt(bytes, position))
        {
            position = ScanForPcdt(bytes);           // tolerate junk before the Pcdt, as LibreOffice does
            prcGrpprls.Clear();                      // whatever the Prc walk collected was junk too
        }
        if (position < 0) { _pieces = [ContiguousPiece(fib)]; return; }
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
            ushort prm = BitConverter.ToUInt16(bytes, pcdStart + i * 8 + 6);
            _pieces[i] = new Piece(cpStart, cpEnd, fc, isCompressed, PrmGrpprlOf(prm, prcGrpprls));
        }
    }

    /// <summary>The Prc grpprls of the Clx (in order); a Prm1.igrpprl indexes this list.</summary>
    public IReadOnlyList<byte[]> PrcGrpprls { get; } = [];

    /// <summary>
    /// Prm ([MS-DOC] §2.9.214): bit 0 fComplex. Prm1 (§2.9.216): igrpprl = prm >> 1, a Prc index.
    /// Prm0 (§2.9.215): isprm = bits 1-7, val = bits 8-15; isprm names a 1-byte-operand sprm through the
    /// spec's table (below), which becomes a 3-byte grpprl. isprm 0 with val 0 has no effect.
    /// </summary>
    private static byte[]? PrmGrpprlOf(ushort prm, List<byte[]> prcGrpprls)
    {
        if (prm == 0) return null;
        if ((prm & 1) != 0)
        {
            int index = prm >> 1;
            return index < prcGrpprls.Count ? prcGrpprls[index] : null;
        }
        int isprm = (prm >> 1) & 0x7F;
        byte value = (byte)(prm >> 8);
        ushort opcode = Prm0SprmOpcodes[isprm];
        if (opcode == 0 || (isprm == 0 && value == 0)) return null;
        return [(byte)(opcode & 0xFF), (byte)(opcode >> 8), value];
    }

    /// <summary>
    /// Prm0.isprm → sprm opcode, from the table in [MS-DOC] §2.9.215 with each name resolved against
    /// §2.6.1 (character) / §2.6.2 (paragraph). Entries the spec leaves out are 0 (no sprm). All are
    /// 1-byte-operand sprms, so a 3-byte grpprl holds them.
    /// </summary>
    private static readonly ushort[] Prm0SprmOpcodes = BuildPrm0SprmOpcodes();

    private static ushort[] BuildPrm0SprmOpcodes()
    {
        var table = new ushort[0x80];
        (byte isprm, ushort opcode)[] entries =
        [
            (0x00, 0x2879), // sprmCLbcCRJ
            (0x04, 0x2602), // sprmPIncLvl
            (0x05, 0x2461), // sprmPJc
            (0x07, 0x2405), // sprmPFKeep
            (0x08, 0x2406), // sprmPFKeepFollow
            (0x09, 0x2407), // sprmPFPageBreakBefore
            (0x0C, 0x260A), // sprmPIlvl
            (0x0D, 0x2470), // sprmPFMirrorIndents
            (0x0E, 0x240C), // sprmPFNoLineNumb
            (0x0F, 0x2471), // sprmPTtwo
            (0x18, 0x2416), // sprmPFInTable
            (0x19, 0x2417), // sprmPFTtp
            (0x1D, 0x261B), // sprmPPc
            (0x25, 0x2423), // sprmPWr
            (0x2C, 0x242A), // sprmPFNoAutoHyph
            (0x32, 0x2430), // sprmPFLocked
            (0x33, 0x2431), // sprmPFWidowControl
            (0x35, 0x2433), // sprmPFKinsoku
            (0x36, 0x2434), // sprmPFWordWrap
            (0x37, 0x2435), // sprmPFOverflowPunct
            (0x38, 0x2436), // sprmPFTopLinePunct
            (0x39, 0x2437), // sprmPFAutoSpaceDE
            (0x3A, 0x2438), // sprmPFAutoSpaceDN
            (0x41, 0x0800), // sprmCFRMarkDel
            (0x42, 0x0801), // sprmCFRMarkIns
            (0x43, 0x0802), // sprmCFFldVanish
            (0x47, 0x0806), // sprmCFData
            (0x4B, 0x080A), // sprmCFOle2
            (0x4D, 0x2A0C), // sprmCHighlight
            (0x4E, 0x0858), // sprmCFEmboss
            (0x4F, 0x2859), // sprmCSfxText
            (0x50, 0x0811), // sprmCFWebHidden
            (0x51, 0x0818), // sprmCFSpecVanish
            (0x53, 0x2A33), // sprmCPlain
            (0x55, 0x0835), // sprmCFBold
            (0x56, 0x0836), // sprmCFItalic
            (0x57, 0x0837), // sprmCFStrike
            (0x58, 0x0838), // sprmCFOutline
            (0x59, 0x0839), // sprmCFShadow
            (0x5A, 0x083A), // sprmCFSmallCaps
            (0x5B, 0x083B), // sprmCFCaps
            (0x5C, 0x083C), // sprmCFVanish
            (0x5E, 0x2A3E), // sprmCKul
            (0x62, 0x2A42), // sprmCIco
            (0x68, 0x2A48), // sprmCIss
            (0x73, 0x2A53), // sprmCFDStrike
            (0x74, 0x0854), // sprmCFImprint
            (0x75, 0x0855), // sprmCFSpec
            (0x76, 0x0856), // sprmCFObj
            (0x78, 0x2640), // sprmPOutLvl
            (0x7B, 0x2A90), // sprmCFSdtVanish
            (0x7C, 0x2A86), // sprmCNeedFontFixup
            (0x7E, 0x2443), // sprmPFNumRMIns
        ];
        foreach ((byte isprm, ushort opcode) in entries) table[isprm] = opcode;
        return table;
    }

    /// <summary>The piece containing a CP.</summary>
    public Piece PieceOf(int cp) => _pieces[PieceIndexOf(cp)];

    /// <summary>A 0x02 whose lcb fits the remaining bytes and describes a whole PlcPcd ((n + 1) CPs + n 8-byte PCDs).</summary>
    private static bool IsPlausiblePcdt(byte[] bytes, int position)
    {
        if (position + 5 > bytes.Length) return false;
        int length = BitConverter.ToInt32(bytes, position + 1);
        return length >= 16 && position + 5 + length <= bytes.Length && (length - 4) % 12 == 0;
    }

    private static int ScanForPcdt(byte[] bytes)
    {
        for (int position = 0; position + 5 <= bytes.Length; position++)
            if (bytes[position] == 0x02 && IsPlausiblePcdt(bytes, position)) return position;
        return -1;
    }

    /// <summary>The whole text as one piece from fcMin to fcMac: 8-bit unless the FIB says extended (Unicode) characters.</summary>
    private static Piece ContiguousPiece(DocBinaryFileInformationBlock fib)
    {
        bool isCompressed = !fib.HasExtendedCharacters;
        int bytesPerCharacter = isCompressed ? 1 : 2;
        int characterCount = (int)Math.Max(0, (fib.TextEndFileOffset - fib.TextStartFileOffset) / bytesPerCharacter);
        return new Piece(0, characterCount, fib.TextStartFileOffset, isCompressed);
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
