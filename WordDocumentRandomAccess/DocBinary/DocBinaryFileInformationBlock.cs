namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// The parts of the File Information Block ([MS-DOC] §2.5.1, the first bytes of the WordDocument
/// stream) that random access needs: which table stream to use, the text piece counts, and the
/// (offset, length) of the piece table (Clx), style sheet and the paragraph/character bin tables.
/// Word 97 and later only; Word 6/95 and encrypted files are rejected.
/// </summary>
public sealed class DocBinaryFileInformationBlock
{
    private const ushort WordIdentifier = 0xA5EC;
    private const ushort FirstWord97Version = 0x00C1;
    private const int FibRgFcLcbStart = 0x9A;    // after FibBase(32) + csw(2) + FibRgW97(28) + cslw(2) + FibRgLw97(88) + cbRgFcLcb(2)

    public ushort Version { get; }
    public bool IsComplex { get; }
    public bool IsEncrypted { get; }
    public bool HasExtendedCharacters { get; }      // fExtChar: text is UTF-16 (only matters for non-complex files)
    public uint TextStartFileOffset { get; }        // fcMin
    public uint TextEndFileOffset { get; }          // fcMac
    public string TableStreamName { get; }
    public int MainTextCharacterCount { get; }      // ccpText
    public int FootnoteCharacterCount { get; }      // ccpFtn
    public int HeaderCharacterCount { get; }        // ccpHdd
    public int CommentCharacterCount { get; }       // ccpAtn
    public int EndnoteCharacterCount { get; }       // ccpEdn
    public (uint offset, uint length) StyleSheet { get; }          // Stshf
    public (uint offset, uint length) PieceTable { get; }          // Clx
    public (uint offset, uint length) CharacterBinTable { get; }   // PlcfBteChpx
    public (uint offset, uint length) ParagraphBinTable { get; }   // PlcfBtePapx
    public (uint offset, uint length) FootnoteReferences { get; }  // PlcffndRef
    public (uint offset, uint length) FootnoteTexts { get; }       // PlcffndTxt
    public (uint offset, uint length) EndnoteReferences { get; }   // PlcfendRef
    public (uint offset, uint length) EndnoteTexts { get; }        // PlcfendTxt
    public (uint offset, uint length) FontTable { get; }           // SttbfFfn
    public (uint offset, uint length) ListFormats { get; }         // PlfLst (FibRgFcLcb97 index 73; the LVL array follows, not counted by lcb)
    public (uint offset, uint length) ListFormatOverrides { get; } // PlfLfo (index 74)

    /// <summary>Where each story starts in the CP space: main text, footnotes, headers, comments, endnotes ([MS-DOC] §2.3).</summary>
    public int FootnoteStoryStart => MainTextCharacterCount;
    public int EndnoteStoryStart => MainTextCharacterCount + FootnoteCharacterCount + HeaderCharacterCount + CommentCharacterCount;

    public DocBinaryFileInformationBlock(Stream wordDocumentStream)
    {
        var fib = new byte[0x600];                     // FibRgFcLcb2007 at most ends near 0x5B6
        wordDocumentStream.Position = 0;
        int read = 0;
        while (read < fib.Length) { int n = wordDocumentStream.Read(fib, read, fib.Length - read); if (n <= 0) break; read += n; }
        if (read < 0x1AA || BitConverter.ToUInt16(fib, 0) != WordIdentifier) throw new InvalidDataException("Not a Word binary document (bad FIB).");

        Version = BitConverter.ToUInt16(fib, 2);
        if (Version < FirstWord97Version) throw new NotSupportedException("Word 6/95 documents are not supported (only Word 97 and later).");
        ushort flags = BitConverter.ToUInt16(fib, 0x0A);
        IsComplex = (flags & 0x0004) != 0;
        IsEncrypted = (flags & 0x0100) != 0;
        HasExtendedCharacters = (flags & 0x1000) != 0;
        TextStartFileOffset = BitConverter.ToUInt32(fib, 0x18);
        TextEndFileOffset = BitConverter.ToUInt32(fib, 0x1C);
        if (IsEncrypted) throw new NotSupportedException("Encrypted (password-protected) Word documents are not supported.");
        TableStreamName = (flags & 0x0200) != 0 ? "1Table" : "0Table";

        MainTextCharacterCount = BitConverter.ToInt32(fib, 0x4C);
        FootnoteCharacterCount = BitConverter.ToInt32(fib, 0x50);
        HeaderCharacterCount = BitConverter.ToInt32(fib, 0x54);
        CommentCharacterCount = BitConverter.ToInt32(fib, 0x5C);
        EndnoteCharacterCount = BitConverter.ToInt32(fib, 0x60);

        _fcLcbCount = BitConverter.ToUInt16(fib, 0x98);
        StyleSheet = FcLcb(fib, 1);
        FootnoteReferences = FcLcb(fib, 2);
        FootnoteTexts = FcLcb(fib, 3);
        CharacterBinTable = FcLcb(fib, 12);
        ParagraphBinTable = FcLcb(fib, 13);
        PieceTable = FcLcb(fib, 33);
        FontTable = FcLcb(fib, 15);
        EndnoteReferences = FcLcb(fib, 46);
        EndnoteTexts = FcLcb(fib, 47);
        ListFormats = FcLcb(fib, 73);
        ListFormatOverrides = FcLcb(fib, 74);
    }

    private readonly int _fcLcbCount;

    /// <summary>An (offset, length) pair of FibRgFcLcb; (0, 0) when this file's FIB is too short to have it.</summary>
    private (uint, uint) FcLcb(byte[] fib, int index)
    {
        int at = FibRgFcLcbStart + index * 8;
        if (index >= _fcLcbCount || at + 8 > fib.Length) return (0, 0);
        return (BitConverter.ToUInt32(fib, at), BitConverter.ToUInt32(fib, at + 4));
    }
}
