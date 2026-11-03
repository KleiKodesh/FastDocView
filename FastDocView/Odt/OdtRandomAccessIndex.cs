using System.Text;
using SeekableDeflate;
using FastDocView.Docx;

namespace FastDocView.Odt;

/// <summary>
/// The compact index of one OpenDocument text. It lists no block: it holds where content.xml sits in the zip and its
/// deflate checkpoints (a .fodt has neither - offsets are file offsets), the root open tag and the captured style
/// sections, the list styles that labelled a paragraph (a reopen never reads styles.xml, so the rescan counts from
/// these), and one <see cref="WordIndexSpanAnchor"/> per ~256 KB of body: that block's index and offset, the block
/// before it, and the scanner's frames + list-numbering state there (as a delta). A block's byte range, kind, text
/// length, list level / label / layout are recovered by rescanning its span with the same
/// <see cref="OdtBodyBlockScanner"/> the index pass ran (<see cref="OdtCompactBlockTable"/>). Index-time lists a
/// reader needs without a scan stay: headings, note references, sections, bookmarks,
/// text-length totals per 64 blocks.
/// </summary>
public sealed class OdtRandomAccessIndex
{
    private const string Magic = "KHODTIDX";
    private const int FormatVersion = 7;     // 5: compact per-block tables; 6: no per-block tables - span anchors + scan state, used list styles, one window store, Elias-Fano lists; 7: graphic styles kept (frame wrap / position)
    /// <summary>Body bytes between span anchors: a random block rescans at most this much on top of the inflate from its checkpoint.</summary>
    public const long AnchorSpanBytes = 256 << 10;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public bool IsFlat { get; }
    public ZipEntryLocation? ContentEntry { get; }           // null for .fodt
    public IReadOnlyList<DeflateCheckpoint> Checkpoints { get; }
    public string RootOpenTag { get; }
    public string[] StyleSections { get; }
    /// <summary>The list styles that labelled a paragraph during the index pass, by name - what a span rescan counts labels from.</summary>
    public IReadOnlyDictionary<string, OdtListStyle> UsedListStyles { get; }
    public int BlockCount { get; }
    /// <summary>Where the last block ends (the last span's rescan stops there).</summary>
    public long LastBlockEndOffset { get; }
    public WordIndexSpanAnchor[] Anchors { get; }
    public int[] TextLengthGroupTotals { get; }
    public WordDocumentHeading[] Headings { get; }
    public OdtNoteReference[] NoteReferences { get; }        // body order
    /// <summary>Per section: its first block and its page text width in twips (see OdtPageSetup).</summary>
    public int[] SectionFirstBlocks { get; init; } = [];
    public int[] SectionTextWidths { get; init; } = [];
    /// <summary>Bookmarks and reference marks (name -> block), see OdtBodyBlockScanner.BookmarkNames.</summary>
    public WordDocumentBookmarkTable Bookmarks { get; init; } = WordDocumentBookmarkTable.Empty;
    private readonly WordListCounterState.DeltaChain _scanStates;

    public OdtRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, bool isFlat, ZipEntryLocation? contentEntry,
                                IReadOnlyList<DeflateCheckpoint> checkpoints, string rootOpenTag, string[] styleSections,
                                IReadOnlyDictionary<string, OdtListStyle> usedListStyles, int blockCount, long lastBlockEndOffset,
                                WordIndexSpanAnchor[] anchors, int[] textLengthGroupTotals,
                                WordDocumentHeading[] headings, OdtNoteReference[] noteReferences)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        IsFlat = isFlat;
        ContentEntry = contentEntry;
        Checkpoints = checkpoints;
        RootOpenTag = rootOpenTag;
        StyleSections = styleSections;
        UsedListStyles = usedListStyles;
        BlockCount = blockCount;
        LastBlockEndOffset = lastBlockEndOffset;
        Anchors = anchors;
        TextLengthGroupTotals = textLengthGroupTotals;
        Headings = headings;
        NoteReferences = noteReferences;
        _scanStates = new WordListCounterState.DeltaChain(anchors.Select(anchor => anchor.ListStateDelta).ToArray(), OdtBodyBlockScanner.ApplyScanStateDelta);
    }

    /// <summary>The scanner state (frames + numbering) at an anchor: what its span's rescan starts from.</summary>
    public byte[] ScanStateOf(int anchorIndex) => _scanStates.StateAt(anchorIndex);

    public int AnchorIndexOf(int blockIndex) => WordIndexSpanAnchor.IndexOfBlock(Anchors, blockIndex);

    /// <summary>The average text length of the block's 64-block group - the scroll weight before its span is scanned.</summary>
    public int EstimatedTextLength(int blockIndex) => WordIndexTextLengthGroups.Average(TextLengthGroupTotals, BlockCount, blockIndex);

    /// <summary>The note references inside blocks [first, last] - a slice of the body-ordered array.</summary>
    public ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
    {
        int start = LowerBound(firstBlockIndex), end = LowerBound(lastBlockIndex + 1);
        return NoteReferences.AsSpan(start, end - start);
    }

    private int LowerBound(int blockIndex)
    {
        int low = 0, high = NoteReferences.Length;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (NoteReferences[middle].BlockIndex < blockIndex) low = middle + 1; else high = middle;
        }
        return low;
    }

    public bool MatchesSource(long fileLength, long lastWriteTicksUtc)
        => fileLength == SourceFileLength && lastWriteTicksUtc == SourceLastWriteTicksUtc;

    public bool MatchesSource(string documentPath)
    {
        var info = new FileInfo(documentPath);
        return info.Exists && MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks);
    }

    // -- serialisation ------------------------------------------------------------------

    /// <summary>Bytes per section of the serialised index - for size reports.</summary>
    public IReadOnlyList<(string section, long bytes)> SizeBreakdown()
    {
        var sizes = new List<(string, long)>();
        using var buffer = new MemoryStream();
        Write(buffer, sizes);
        return sizes;
    }

    public void Write(Stream destination) => Write(destination, null);

    private void Write(Stream destination, List<(string, long)>? sizes)
    {
        var writer = new BinaryWriter(destination, Encoding.UTF8, leaveOpen: true);
        long mark = destination.Position;
        void Section(string name) { writer.Flush(); sizes?.Add((name, destination.Position - mark)); mark = destination.Position; }
        writer.Write(Magic);
        writer.Write(FormatVersion);
        writer.Write(SourceFileLength);
        writer.Write(SourceLastWriteTicksUtc);
        writer.Write(IsFlat);
        writer.Write(ContentEntry is not null);
        if (ContentEntry is not null) DocxRandomAccessIndex.WriteEntry(writer, ContentEntry);
        writer.Write(BlockCount);
        writer.Write(LastBlockEndOffset);
        Section("header");
        DeflateCheckpointTable.Write(writer, Checkpoints);
        Section($"checkpoints + windows ({Checkpoints.Count})");
        IndexBinaryArrays.WriteCompressedString(writer, RootOpenTag);
        IndexBinaryArrays.WriteCompressedStrings(writer, StyleSections);
        Section("root tag + style sections");
        string[] styleNames = UsedListStyles.Keys.OrderBy(name => name, StringComparer.Ordinal).ToArray();
        IndexBinaryArrays.WriteCompressedStrings(writer, styleNames);
        using (var raw = new MemoryStream())
        {
            using (var styleWriter = new BinaryWriter(raw, Encoding.UTF8, leaveOpen: true))
                foreach (string name in styleNames) UsedListStyles[name].Write(styleWriter);
            WordIndexDeflatedBlob.Write(writer, raw.ToArray());
        }
        Section($"list styles ({styleNames.Length})");
        WordIndexSpanAnchor.WriteAll(writer, Anchors);
        Section($"anchors ({Anchors.Length}, scan state)");
        IndexBinaryArrays.WriteVarints(writer, TextLengthGroupTotals);
        Section("text-length groups");
        new EliasFanoSequence(Headings.Select(h => h.BlockIndex).ToArray()).Write(writer);
        foreach (WordDocumentHeading heading in Headings) writer.Write((byte)heading.Level);
        IndexBinaryArrays.WriteCompressedStrings(writer, Headings.Select(h => h.Text).ToArray());
        Section($"headings ({Headings.Length})");
        // note references: blocks as Elias-Fano; kind + number as varints (numbers delta-coded per kind); labels deflated. The bodies' byte
        // ranges are not stored: nothing reads them (a note is rendered inside its block), so a loaded index reports -1 for them.
        new EliasFanoSequence(NoteReferences.Select(r => r.BlockIndex).ToArray()).Write(writer);
        var noteBytes = new List<byte>(NoteReferences.Length * 3);
        var previousNumber = new int[2];
        foreach (OdtNoteReference note in NoteReferences)
        {
            int kind = (int)note.Kind;
            noteBytes.Add((byte)kind);
            IndexBinaryArrays.AppendVarint(noteBytes, WordListCounterState.Zigzag(note.Number - previousNumber[kind]));
            previousNumber[kind] = note.Number;
        }
        writer.Write(noteBytes.Count);
        writer.Write(noteBytes.ToArray());
        IndexBinaryArrays.WriteCompressedStrings(writer, NoteReferences.Select(r => r.Label).ToArray());
        Section($"note references ({NoteReferences.Length})");
        IndexBinaryArrays.WriteDeltaVarints(writer, SectionFirstBlocks);
        IndexBinaryArrays.WriteVarints(writer, SectionTextWidths);
        Section($"sections ({SectionFirstBlocks.Length})");
        IndexBinaryArrays.WriteCompressedStrings(writer, Bookmarks.Names);
        new EliasFanoSequence(Bookmarks.Blocks).Write(writer);
        Section($"bookmarks ({Bookmarks.Names.Length})");
        writer.Flush();
    }

    public static OdtRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not an OpenDocument index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("OpenDocument index: unsupported version.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        bool isFlat = reader.ReadBoolean();
        ZipEntryLocation? entry = reader.ReadBoolean() ? DocxRandomAccessIndex.ReadEntry(reader) : null;
        int blockCount = reader.ReadInt32();
        long lastBlockEnd = reader.ReadInt64();
        if (blockCount < 0 || lastBlockEnd < 0) throw new InvalidDataException("OpenDocument index: negative count.");
        List<DeflateCheckpoint> checkpoints = DeflateCheckpointTable.Read(reader);
        string rootOpenTag = IndexBinaryArrays.ReadCompressedString(reader);
        string[] styleSections = IndexBinaryArrays.ReadCompressedStrings(reader);
        string[] styleNames = IndexBinaryArrays.ReadCompressedStrings(reader);
        var usedListStyles = new Dictionary<string, OdtListStyle>(StringComparer.Ordinal);
        using (var styleReader = new BinaryReader(new MemoryStream(WordIndexDeflatedBlob.Read(reader)), Encoding.UTF8))
            foreach (string name in styleNames) usedListStyles[name] = OdtListStyle.Read(styleReader);
        WordIndexSpanAnchor[] anchors = WordIndexSpanAnchor.ReadAll(reader, blockCount, lastBlockEnd);
        int[] groupTotals = IndexBinaryArrays.ReadVarints(reader);
        WordIndexTextLengthGroups.Validate(groupTotals, blockCount);
        int[] headingBlocks = EliasFanoSequence.Read(reader).ToArray();
        var levels = new byte[headingBlocks.Length];
        for (int i = 0; i < levels.Length; i++) levels[i] = reader.ReadByte();
        string[] headingTexts = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (headingTexts.Length != headingBlocks.Length || headingBlocks.Any(block => block >= blockCount)) throw new InvalidDataException("OpenDocument index: headings do not match the blocks.");
        var headings = new WordDocumentHeading[headingBlocks.Length];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(headingBlocks[i], levels[i], headingTexts[i]);
        int[] noteBlocks = EliasFanoSequence.Read(reader).ToArray();
        int noteByteCount = reader.ReadInt32();
        if (noteByteCount < 0 || noteByteCount > reader.BaseStream.Length - reader.BaseStream.Position) throw new InvalidDataException("OpenDocument index: note references damaged.");
        byte[] noteBytes = reader.ReadBytes(noteByteCount);
        if (noteBytes.Length != noteByteCount) throw new EndOfStreamException("OpenDocument index: note references truncated.");
        string[] noteLabels = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (noteLabels.Length != noteBlocks.Length || noteBlocks.Any(block => block >= blockCount)) throw new InvalidDataException("OpenDocument index: note references do not match the blocks.");
        var notes = new OdtNoteReference[noteBlocks.Length];
        {
            int position = 0;
            var previousNumber = new int[2];
            for (int i = 0; i < notes.Length; i++)
            {
                if (position >= noteBytes.Length) throw new EndOfStreamException("OpenDocument index: note references truncated.");
                int kind = noteBytes[position++];
                if (kind > 1) throw new InvalidDataException("OpenDocument index: bad note kind.");
                previousNumber[kind] += WordListCounterState.Unzigzag(IndexBinaryArrays.ReadVarint(noteBytes, ref position));
                notes[i] = new OdtNoteReference(noteBlocks[i], (WordDocumentNoteKind)kind, previousNumber[kind], noteLabels[i], -1, -1);
            }
        }
        int[] sectionFirstBlocks = IndexBinaryArrays.ReadDeltaVarintsInt32(reader);
        int[] sectionTextWidths = IndexBinaryArrays.ReadVarints(reader);
        if (sectionFirstBlocks.Length != sectionTextWidths.Length || sectionFirstBlocks.Any(block => (uint)block >= (uint)Math.Max(1, blockCount)))
            throw new InvalidDataException("OpenDocument index: section table does not match the blocks.");
        string[] bookmarkNames = IndexBinaryArrays.ReadCompressedStrings(reader);
        int[] bookmarkBlocks = EliasFanoSequence.Read(reader).ToArray();
        if (bookmarkBlocks.Length != bookmarkNames.Length || bookmarkBlocks.Any(block => block < 0 || block >= Math.Max(1, blockCount)))
            throw new InvalidDataException("OpenDocument index: bookmark table does not match the blocks.");
        return new OdtRandomAccessIndex(sourceLength, sourceTicks, isFlat, entry, checkpoints, rootOpenTag, styleSections, usedListStyles,
                                        blockCount, lastBlockEnd, anchors, groupTotals, headings, notes)
        {
            SectionFirstBlocks = sectionFirstBlocks,
            SectionTextWidths = sectionTextWidths,
            Bookmarks = new WordDocumentBookmarkTable(bookmarkNames, bookmarkBlocks),
        };
    }

    // -- sidecar file helpers -----------------------------------------------------------

    /// <summary>Saves the sidecar atomically with its record (<see cref="WordDocumentIndexCache.Save"/>); a failure leaves the old one or none - it is a cache.</summary>
    public void Save(string indexPath, string? documentPath = null) => WordDocumentIndexCache.Save(indexPath, Write, documentPath);

    public static OdtRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader? reader = IndexBinaryArrays.TryOpenInMemory(indexPath);
            if (reader is null) return null;                  // gone, held or pending delete this instant: a cache miss
            OdtRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? WordDocumentIndexCache.MarkUsed(indexPath, index) : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException or OverflowException or ArgumentException or FormatException)
        {
            return null;
        }
    }
}
