namespace FastDocView.DocBinary;

/// <summary>
/// The parts of the File Information Block ([MS-DOC] §2.5.1, the first bytes of the WordDocument
/// stream) that random access needs: which table stream to use, the text piece counts, and the
/// (offset, length) of the piece table (Clx), style sheet and the paragraph/character bin tables.
/// Word 97 and later, and Word 6.0 / 95 (its shorter FIB, LibreOffice ww8scan.cxx WW8Fib); encrypted files are rejected.
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
    /// <summary>ccpMcr: a story between the headers and the comments (ww8scan.cxx GetBaseCp MAN_AND adds it); Word 97+ writes 0.</summary>
    public int MacroCharacterCount { get; }
    public int EndnoteCharacterCount { get; }       // ccpEdn
    public int TextBoxCharacterCount { get; }
    public int HeaderTextBoxCharacterCount { get; } // ccpHdrTxbx (after ccpTxbx, ww8scan.cxx:5976)       // ccpTxbx (FibRgLw97 after ccpEdn; LibreOffice ww8scan.cxx:5969-5976 read order)
    public (uint offset, uint length) StyleSheet { get; }          // Stshf
    public (uint offset, uint length) PieceTable { get; }          // Clx
    public (uint offset, uint length) CharacterBinTable { get; }   // PlcfBteChpx
    public (uint offset, uint length) ParagraphBinTable { get; }   // PlcfBtePapx
    public (uint offset, uint length) FootnoteReferences { get; }  // PlcffndRef
    public (uint offset, uint length) FootnoteTexts { get; }       // PlcffndTxt
    public (uint offset, uint length) SectionDescriptors { get; }  // PlcfSed (FibRgFcLcb97 index 6)
    public (uint offset, uint length) EndnoteReferences { get; }   // PlcfendRef
    public (uint offset, uint length) EndnoteTexts { get; }        // PlcfendTxt
    public (uint offset, uint length) FontTable { get; }           // SttbfFfn
    public (uint offset, uint length) ListFormats { get; }         // PlfLst (FibRgFcLcb97 index 73; the LVL array follows, not counted by lcb)
    public (uint offset, uint length) ListFormatOverrides { get; } // PlfLfo (index 74)
    public (uint offset, uint length) DocumentProperties { get; }  // Dop (index 31)
    public (uint offset, uint length) BookmarkNames { get; }       // SttbfBkmk (index 21; LibreOffice ww8scan.cxx WW8Fib read order)
    public (uint offset, uint length) BookmarkStarts { get; }      // PlcfBkf (index 22)
    public (uint offset, uint length) BookmarkEnds { get; }        // PlcfBkl (index 23; ww8scan.cxx WW8Fib read order: Plcfbkf, Plcfbkl)
    public (uint offset, uint length) MainShapeAnchors { get; }    // PlcSpaMom (index 40; LibreOffice ww8scan.cxx:6117 read order)
    public (uint offset, uint length) DrawingGroup { get; }        // OfficeArtContent at fcDggInfo (index 50; ww8scan.cxx:6138)
    public (uint offset, uint length) TextBoxTexts { get; }        // PlcftxbxTxt (index 56; ww8scan.cxx:6138-6150 read order)
    public (uint offset, uint length) TextBoxBreaks { get; }
    public (uint offset, uint length) HeaderShapeAnchors { get; }  // PlcSpaHdr (index 41; ww8scan.cxx:6117-6120 read order)
    public (uint offset, uint length) HeaderTextBoxTexts { get; }  // PlcfHdrtxbxTxt (index 58; ww8scan.cxx:6150-6154)
    public (uint offset, uint length) HeaderTextBoxBreaks { get; } // PlcfHdrtxbxBkd (index 76; ww8scan.cxx:6216-6217)       // PlcftxbxBkd (index 75; ww8scan.cxx:6208-6216, at 0x2da + 3 * 8)

    /// <summary>Where each story starts in the CP space: main text, footnotes, headers, comments, endnotes ([MS-DOC] §2.3).</summary>
    public int FootnoteStoryStart => MainTextCharacterCount;
    public int EndnoteStoryStart => MainTextCharacterCount + FootnoteCharacterCount + HeaderCharacterCount + MacroCharacterCount + CommentCharacterCount;
    /// <summary>The main text-box story starts after the endnotes (ww8scan.cxx GetBaseCp MAN_TXBX).</summary>
    public int TextBoxStoryStart => EndnoteStoryStart + EndnoteCharacterCount;
    /// <summary>The header text-box story follows the main one (ww8scan.cxx GetBaseCp MAN_TXBX_HDFT adds ccpTxbx).</summary>
    public int HeaderTextBoxStoryStart => TextBoxStoryStart + TextBoxCharacterCount;
    /// <summary>The header / footer story starts after the footnotes ([MS-DOC] §2.3; ww8scan.cxx GetBaseCp MAN_HDFT).</summary>
    public int HeaderStoryStart => MainTextCharacterCount + FootnoteCharacterCount;

    public DocBinaryFileInformationBlock(Stream wordDocumentStream)
    {
        var fib = new byte[0x600];                     // FibRgFcLcb2007 at most ends near 0x5B6
        wordDocumentStream.Position = 0;
        int read = 0;
        while (read < fib.Length) { int n = wordDocumentStream.Read(fib, read, fib.Length - read); if (n <= 0) break; read += n; }
        ushort identifier = read >= 4 ? BitConverter.ToUInt16(fib, 0) : (ushort)0;
        // Word for Windows 1.x (wIdent 0xA59B, 0xA59C = OS/2 "PM Word") and 2.0 (0xA5DB): no compound file, the FIB at the
        // file's start (LibreOffice ww8scan.cxx GetFIBVersion: eWW1 / eWW2). They take the Word 6 branches plus their own.
        IsWord1 = identifier is 0xA59B or 0xA59C;
        IsWord2 = IsWord1 || identifier == 0xA5DB;
        // Word 6.0 / 95 FIBs: wIdent 0xA5DC, or 0xA697-0xA699 (LibreOffice ww8par.cxx accepts them; GetWW6SprmSearcher treats
        // the 0xA697-0xA699 ones as the far-east variant, sprms 111-113 = font codes); nFib 101-105 (ww8scan.cxx WW8Fib ctor).
        IsWord6 = IsWord2 || identifier is 0xA5DC or >= 0xA697 and <= 0xA699;
        IsWord6FontCodeVariant = identifier is >= 0xA697 and <= 0xA699;
        if (IsWord2 && read < Word2FibLength)
            throw new InvalidDataException($"Damaged Word for Windows {(IsWord1 ? "1.x" : "2.0")} document: the file ({read} bytes) is shorter than its FIB.");
        if (IsWord6 && !IsWord2 && read < 0x192 + 22 * 8)
            throw new InvalidDataException($"Damaged Word 6/95 document: the WordDocument stream ({read} bytes) is shorter than its FIB.");
        if (!IsWord6 && (read < 0x1AA || identifier != WordIdentifier))
            throw new InvalidDataException($"Not a Word 97+ binary document: the WordDocument stream starts with 0x{identifier:X4}, not the FIB identifier 0xA5EC.");

        Version = BitConverter.ToUInt16(fib, 2);
        if (!IsWord6 && Version < FirstWord97Version) throw new NotSupportedException($"Word 6/95 documents with a Word 97 FIB identifier are not supported (nFib {Version}).");
        ushort flags = BitConverter.ToUInt16(fib, 0x0A);
        IsComplex = (flags & 0x0004) != 0;
        IsEncrypted = (flags & 0x0100) != 0;
        HasExtendedCharacters = !IsWord2 && (flags & 0x1000) != 0;     // Word 6/95: every piece is UTF-16 when set, else 8-bit (ww8scan.cxx WW8Cp2Fc: m_fExtChar for nVersion < 8)
        TextStartFileOffset = BitConverter.ToUInt32(fib, 0x18);
        TextEndFileOffset = BitConverter.ToUInt32(fib, 0x1C);
        if (IsEncrypted) throw new WordDocumentEncryptedException("Encrypted (password-protected) Word documents are not supported.");
        // Word 6/95 has no table stream: the tables live in the WordDocument stream (ww8par.cxx: pTableStream = pStrm for bVer67).
        TableStreamName = IsWord6 ? "WordDocument" : (flags & 0x0200) != 0 ? "1Table" : "0Table";
        Language = BitConverter.ToUInt16(fib, 6);

        if (IsWord6)
        {
            // Word 6/95 rglw (ww8scan.cxx WW8Fib ctor, IsSevenMinus branches): cbMac at 0x20, four longs skipped, then ccpText,
            // ccpFtn, ccpHdr, ccpMcr, ccpAtn, ccpEdn, ccpTxbx, ccpHdrTxbx from 0x34; one long skipped; the fc/lcb pairs from 0x58.
            MainTextCharacterCount = BitConverter.ToInt32(fib, 0x34);
            FootnoteCharacterCount = BitConverter.ToInt32(fib, 0x38);
            HeaderCharacterCount = BitConverter.ToInt32(fib, 0x3C);
            MacroCharacterCount = Math.Max(0, BitConverter.ToInt32(fib, 0x40));
            CommentCharacterCount = BitConverter.ToInt32(fib, 0x44);
            EndnoteCharacterCount = BitConverter.ToInt32(fib, 0x48);
            TextBoxCharacterCount = Math.Max(0, BitConverter.ToInt32(fib, 0x4C));
            HeaderTextBoxCharacterCount = Math.Max(0, BitConverter.ToInt32(fib, 0x50));
            if (IsWord1)
            {
                // Word 1: 36 six-byte pairs end at 0x130; a long and a short skipped; cpnBteChp, cpnBtePap (no pnFirst: 0)
                CharacterBinPageCount = BitConverter.ToInt16(fib, 0x136);
                ParagraphBinPageCount = BitConverter.ToInt16(fib, 0x138);
            }
            else if (IsWord2)
            {
                // Word 2: 38 six-byte pairs end at 0x13C; a short, pnChpFirst, pnPapFirst, cpnBteChp, cpnBtePap
                CharacterFirstPage = BitConverter.ToInt16(fib, 0x13E);
                ParagraphFirstPage = BitConverter.ToInt16(fib, 0x140);
                CharacterBinPageCount = BitConverter.ToInt16(fib, 0x142);
                ParagraphBinPageCount = BitConverter.ToInt16(fib, 0x144);
            }
            else
            {
                // after pair 37 (fcSttbfAtnbkmk): a short, pnChpFirst, pnPapFirst, cpnBteChp, cpnBtePap (0x188..0x191), then pairs 38.. at 0x192
                CharacterFirstPage = BitConverter.ToInt16(fib, 0x18A);
                ParagraphFirstPage = BitConverter.ToInt16(fib, 0x18C);
                CharacterBinPageCount = BitConverter.ToInt16(fib, 0x18E);
                ParagraphBinPageCount = BitConverter.ToInt16(fib, 0x190);
            }
        }
        else
        {
            MainTextCharacterCount = BitConverter.ToInt32(fib, 0x4C);
            FootnoteCharacterCount = BitConverter.ToInt32(fib, 0x50);
            HeaderCharacterCount = BitConverter.ToInt32(fib, 0x54);
            CommentCharacterCount = BitConverter.ToInt32(fib, 0x5C);
            EndnoteCharacterCount = BitConverter.ToInt32(fib, 0x60);
            TextBoxCharacterCount = Math.Max(0, BitConverter.ToInt32(fib, 0x64));
            HeaderTextBoxCharacterCount = Math.Max(0, BitConverter.ToInt32(fib, 0x68));
        }
        if (MainTextCharacterCount < 0 || FootnoteCharacterCount < 0 || HeaderCharacterCount < 0 || CommentCharacterCount < 0 || EndnoteCharacterCount < 0)
            throw new InvalidDataException("Damaged Word document: the FIB gives a negative character count.");

        _isWord6 = IsWord6;
        _fcLcbCount = IsWord2 ? Word2FcLcbCount : IsWord6 ? Word6FcLcbCount : BitConverter.ToUInt16(fib, 0x98);
        StyleSheet = FcLcb(fib, 1);
        FootnoteReferences = FcLcb(fib, 2);
        FootnoteTexts = FcLcb(fib, 3);
        SectionDescriptors = FcLcb(fib, 6);
        CharacterBinTable = FcLcb(fib, 12);
        ParagraphBinTable = FcLcb(fib, 13);
        PieceTable = FcLcb(fib, 33);
        FontTable = FcLcb(fib, 15);
        EndnoteReferences = FcLcb(fib, 46);
        EndnoteTexts = FcLcb(fib, 47);
        ListFormats = FcLcb(fib, 73);
        ListFormatOverrides = FcLcb(fib, 74);
        DocumentProperties = FcLcb(fib, 31);
        BookmarkNames = FcLcb(fib, 21);
        BookmarkStarts = FcLcb(fib, 22);
        BookmarkEnds = FcLcb(fib, 23);
        MainShapeAnchors = FcLcb(fib, 40);
        DrawingGroup = FcLcb(fib, 50);
        TextBoxTexts = FcLcb(fib, 56);
        TextBoxBreaks = FcLcb(fib, 75);
        HeaderShapeAnchors = FcLcb(fib, 41);
        HeaderTextBoxTexts = FcLcb(fib, 58);
        HeaderTextBoxBreaks = FcLcb(fib, 76);
        foreach (int pair in (int[])[16, 17, 18, 48, 57, 59]) _fieldTables[pair] = FcLcb(fib, pair);
    }

    private readonly Dictionary<int, (uint, uint)> _fieldTables = new();

    /// <summary>A PlcFld (FibRgFcLcb97 pair 16 main, 17 headers, 18 footnotes, 48 endnotes, 57 text boxes, 59 header text boxes - see DocBinaryFields); (0, 0) when absent.</summary>
    public (uint offset, uint length) FieldTable(int pair) => _fieldTables.TryGetValue(pair, out var table) ? table : (0, 0);

    private readonly int _fcLcbCount;
    private readonly bool _isWord6;
    private const int Word2FcLcbCount = 38;    // pairs 0..37 (through fcSttbfAtnbkmk); Word 1 stores 36 of them
    private const int Word2FibLength = 0x146;  // through cpnBtePap (Word 1 ends at 0x13A)
    private const int Word6FcLcbCount = 62;    // pairs 0..61 (through fcSttbttmbd), ww8scan.cxx WW8Fib ctor read order for eWW6 / eWW7

    /// <summary>A Word for Windows 1.x / 2.0 file (wIdent 0xA59B / 0xA59C / 0xA5DB): <see cref="IsWord6"/> is set too.</summary>
    public bool IsWord2 { get; }
    /// <summary>Word for Windows 1.x (wIdent 0xA59B, 0xA59C): a shorter FIB than Word 2's.</summary>
    public bool IsWord1 { get; }
    /// <summary>A Word 6.0 / 95 file (wIdent 0xA5DC or 0xA697-0xA699) or an older one (<see cref="IsWord2"/>): one-byte sprms, 8-bit text, no table stream.</summary>
    public bool IsWord6 { get; }
    /// <summary>wIdent 0xA697-0xA699: sprms 111-113 are font codes, not the right-to-left bold / italic / font (ww8scan.cxx patchCJKVariant).</summary>
    public bool IsWord6FontCodeVariant { get; }
    /// <summary>The FIB's lid (the language of the Word that saved the file); picks the 8-bit code page of Word 6/95 text when no font names one.</summary>
    public ushort Language { get; }
    /// <summary>Word 6/95: the first FKP page and the FKP page counts the bin tables may list only partly (ww8scan.cxx WW8PLCF::GeneratePLCF).</summary>
    public int CharacterFirstPage { get; }
    public int ParagraphFirstPage { get; }
    public int CharacterBinPageCount { get; }
    public int ParagraphBinPageCount { get; }

    /// <summary>An (offset, length) pair of FibRgFcLcb; (0, 0) when this file's FIB is too short to have it.</summary>
    private (uint, uint) FcLcb(byte[] fib, int index)
    {
        if (IsWord2)
        {
            // Word 1/2: a 4-byte fc and a 2-byte lcb (ww8scan.cxx Readcb, eVer <= eWW2) from 0x58, in the Word 97 order up to
            // fcSttbfAtnbkmk; Word 1 has no fcPrDrvr / fcPrEnvLand (only fcPrEnvPort), so from fcPrEnvPort on it is two pairs behind.
            if (index >= Word2FcLcbCount || IsWord1 && index is 27 or 29) return (0, 0);
            int pair = IsWord1 && index > 27 ? index - (index > 29 ? 2 : 1) : index;
            int word2At = 0x58 + pair * 6;
            uint lcb = BitConverter.ToUInt16(fib, word2At + 4);
            return lcb == 0 ? (0, 0) : (BitConverter.ToUInt32(fib, word2At), lcb);
        }
        int at = !_isWord6 ? FibRgFcLcbStart + index * 8 : index < 38 ? 0x58 + index * 8 : 0x192 + (index - 38) * 8;
        if (index >= _fcLcbCount || at + 8 > fib.Length) return (0, 0);
        return (BitConverter.ToUInt32(fib, at), BitConverter.ToUInt32(fib, at + 4));
    }
}
