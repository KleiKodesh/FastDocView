using System.Text;

namespace FastDocView.Rtf;

/// <summary>A note referenced in a block: kind, number among the auto-numbered notes of its kind (0 = custom mark), label.</summary>
public readonly record struct RtfNoteReference(int BlockIndex, WordDocumentNoteKind Kind, int Number, string Label);

/// <summary>
/// The compact index of one RTF file. RTF is plain bytes, so a block is a byte range of the file itself - no
/// checkpoints - and the index lists no block: it keeps the header tables rendering needs and one
/// <see cref="WordIndexSpanAnchor"/> per ~256 KB of file (that block's index and offset, the block before it, the
/// parser's group-state snapshot there and the list-numbering state as a delta). A block's byte range, kind, text
/// length, start snapshot and list label are recovered by rescanning its span with the same block rules the index
/// pass ran (<see cref="RtfRandomAccessIndexBuilder.ScanSpan"/>, served by <see cref="RtfCompactBlockTable"/>).
/// Index-time lists a reader needs without a scan stay: headings, note references, sections, bookmarks, text-length
/// totals per 64 blocks.
/// </summary>
public sealed class RtfRandomAccessIndex
{
    private const string Magic = "KHRTFIDX";
    private const int FormatVersion = 13;    // 13: list level fonts; 12: drop caps, hyphenation, language (group state + header); 11: gutter placement (rtlgutter, gutterprl) in section widths; 10: group state carries borders, shading, underline colour, emphasis; 9: section column lefts; 8: XE / TA entries;     // 7: anchor heads hold the open fields + field counters (RtfFieldTracker)
     // 4: compact block tables; 5: no per-block tables - span anchors + parser snapshot + list state, Elias-Fano lists; 6: row indent + alignment in the snapshot
    /// <summary>File bytes between span anchors: a random block rescans at most this much.</summary>
    public const long AnchorSpanBytes = 256 << 10;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public RtfDocumentTables Tables { get; }
    public bool IsRightToLeft { get; }
    public int BlockCount { get; }
    /// <summary>Where the last block ends (the last span's rescan stops there).</summary>
    public long LastBlockEndOffset { get; }
    public WordIndexSpanAnchor[] Anchors { get; }
    public int[] TextLengthGroupTotals { get; }
    public WordDocumentHeading[] Headings { get; }
    public RtfNoteReference[] NoteReferences { get; }         // body order
    /// <summary>Per section: its last block (inclusive) and its page text width in twips.</summary>
    public int[] SectionLastBlocks { get; init; } = [];
    public int[] SectionTextWidths { get; init; } = [];
    public int[] SectionColumnLefts { get; init; } = [];
    /// <summary>Bookmarks ({\*\bkmkstart name}) and the block each marks.</summary>
    public WordDocumentBookmarkTable Bookmarks { get; init; } = WordDocumentBookmarkTable.Empty;

    /// <summary>XE / TA field codes with their blocks ("block U+0001 code"): an INDEX / TOA saved without a result is built from them.</summary>
    public string[] IndexEntries { get; init; } = [];
    private readonly WordListCounterState.DeltaChain _scanStates;

    public RtfRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, RtfDocumentTables tables, bool isRightToLeft,
                                int blockCount, long lastBlockEndOffset, WordIndexSpanAnchor[] anchors, int[] textLengthGroupTotals,
                                WordDocumentHeading[] headings, RtfNoteReference[] noteReferences)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        Tables = tables;
        IsRightToLeft = isRightToLeft;
        BlockCount = blockCount;
        LastBlockEndOffset = lastBlockEndOffset;
        Anchors = anchors;
        TextLengthGroupTotals = textLengthGroupTotals;
        Headings = headings;
        NoteReferences = noteReferences;
        _scanStates = new WordListCounterState.DeltaChain(anchors.Select(anchor => anchor.ListStateDelta).ToArray(), WordIndexScanState.Apply);
    }

    /// <summary>The scan state (parser snapshot + list numbering) at an anchor: what its span's rescan starts from.</summary>
    public byte[] ScanStateOf(int anchorIndex) => _scanStates.StateAt(anchorIndex);

    public int AnchorIndexOf(int blockIndex) => WordIndexSpanAnchor.IndexOfBlock(Anchors, blockIndex);

    /// <summary>The average text length of the block's 64-block group - the scroll weight before its span is scanned.</summary>
    public int EstimatedTextLength(int blockIndex) => WordIndexTextLengthGroups.Average(TextLengthGroupTotals, BlockCount, blockIndex);

    public ReadOnlySpan<RtfNoteReference> NoteReferencesInBlock(int blockIndex)
    {
        int start = LowerBound(blockIndex), end = LowerBound(blockIndex + 1);
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
        writer.Write(RtfGroupState.Size);                     // the anchors' snapshots are raw RtfGroupState bytes
        writer.Write(SourceFileLength);
        writer.Write(SourceLastWriteTicksUtc);
        writer.Write(IsRightToLeft);
        writer.Write(BlockCount);
        writer.Write(LastBlockEndOffset);
        Section("header");
        Tables.Write(writer);
        Section("header tables (fonts, colours, styles, lists)");
        WordIndexSpanAnchor.WriteAll(writer, Anchors);
        Section($"anchors ({Anchors.Length}, parser snapshot + list state)");
        IndexBinaryArrays.WriteVarints(writer, TextLengthGroupTotals);
        Section("text-length groups");
        new EliasFanoSequence(Headings.Select(h => h.BlockIndex).ToArray()).Write(writer);
        foreach (WordDocumentHeading heading in Headings) writer.Write((byte)heading.Level);
        IndexBinaryArrays.WriteCompressedStrings(writer, Headings.Select(h => h.Text).ToArray());
        Section($"headings ({Headings.Length})");
        new EliasFanoSequence(NoteReferences.Select(r => r.BlockIndex).ToArray()).Write(writer);
        var noteBytes = new List<byte>(NoteReferences.Length * 3);
        var previousNumber = new int[2];
        foreach (RtfNoteReference note in NoteReferences)
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
        IndexBinaryArrays.WriteDeltaVarints(writer, SectionLastBlocks);
        IndexBinaryArrays.WriteVarints(writer, SectionTextWidths);
        IndexBinaryArrays.WriteVarints(writer, SectionColumnLefts);
        Section($"sections ({SectionLastBlocks.Length})");
        IndexBinaryArrays.WriteCompressedStrings(writer, Bookmarks.Names);
        new EliasFanoSequence(Bookmarks.Blocks).Write(writer);
        Section($"bookmarks ({Bookmarks.Names.Length})");
        IndexBinaryArrays.WriteCompressedStrings(writer, IndexEntries);
        writer.Flush();
    }

    public static RtfRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not an RTF index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("RTF index: unsupported version.");
        if (reader.ReadInt32() != RtfGroupState.Size) throw new InvalidDataException("RTF index: group-state layout changed.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        bool isRightToLeft = reader.ReadBoolean();
        int blockCount = reader.ReadInt32();
        long lastBlockEnd = reader.ReadInt64();
        if (blockCount < 0 || lastBlockEnd < 0) throw new InvalidDataException("RTF index: negative count.");
        RtfDocumentTables tables = RtfDocumentTables.Read(reader);
        WordIndexSpanAnchor[] anchors = WordIndexSpanAnchor.ReadAll(reader, blockCount, lastBlockEnd);
        int[] groupTotals = IndexBinaryArrays.ReadVarints(reader);
        WordIndexTextLengthGroups.Validate(groupTotals, blockCount);
        int[] headingBlocks = EliasFanoSequence.Read(reader).ToArray();
        var levels = new byte[headingBlocks.Length];
        for (int i = 0; i < levels.Length; i++) levels[i] = reader.ReadByte();
        string[] headingTexts = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (headingTexts.Length != headingBlocks.Length || headingBlocks.Any(block => block >= blockCount)) throw new InvalidDataException("RTF index: headings do not match the blocks.");
        var headings = new WordDocumentHeading[headingBlocks.Length];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(headingBlocks[i], levels[i], headingTexts[i]);
        int[] noteBlocks = EliasFanoSequence.Read(reader).ToArray();
        int noteByteCount = reader.ReadInt32();
        if (noteByteCount < 0 || noteByteCount > reader.BaseStream.Length - reader.BaseStream.Position) throw new InvalidDataException("RTF index: note references damaged.");
        byte[] noteBytes = reader.ReadBytes(noteByteCount);
        if (noteBytes.Length != noteByteCount) throw new EndOfStreamException("RTF index: note references truncated.");
        string[] noteLabels = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (noteLabels.Length != noteBlocks.Length || noteBlocks.Any(block => block >= Math.Max(1, blockCount))) throw new InvalidDataException("RTF index: note references do not match the blocks.");
        var notes = new RtfNoteReference[noteBlocks.Length];
        {
            int position = 0;
            var previousNumber = new int[2];
            for (int i = 0; i < notes.Length; i++)
            {
                if (position >= noteBytes.Length) throw new EndOfStreamException("RTF index: note references truncated.");
                int kind = noteBytes[position++];
                if (kind > 1) throw new InvalidDataException("RTF index: bad note kind.");
                previousNumber[kind] += WordListCounterState.Unzigzag(IndexBinaryArrays.ReadVarint(noteBytes, ref position));
                notes[i] = new RtfNoteReference(noteBlocks[i], (WordDocumentNoteKind)kind, previousNumber[kind], noteLabels[i]);
            }
        }
        int[] sectionLastBlocks = IndexBinaryArrays.ReadDeltaVarintsInt32(reader);
        int[] sectionTextWidths = IndexBinaryArrays.ReadVarints(reader);
        int[] sectionColumnLefts = IndexBinaryArrays.ReadVarints(reader);
        if (sectionLastBlocks.Length != sectionTextWidths.Length || sectionColumnLefts.Length != sectionTextWidths.Length || sectionLastBlocks.Any(block => (uint)block >= (uint)blockCount))
            throw new InvalidDataException("RTF index: section table does not match the blocks.");
        string[] bookmarkNames = IndexBinaryArrays.ReadCompressedStrings(reader);
        int[] bookmarkBlocks = EliasFanoSequence.Read(reader).ToArray();
        if (bookmarkBlocks.Length != bookmarkNames.Length || bookmarkBlocks.Any(block => block < 0 || block >= Math.Max(1, blockCount)))
            throw new InvalidDataException("RTF index: bookmark table does not match the blocks.");
        string[] indexEntries = IndexBinaryArrays.ReadCompressedStrings(reader);
        return new RtfRandomAccessIndex(sourceLength, sourceTicks, tables, isRightToLeft, blockCount, lastBlockEnd, anchors, groupTotals, headings, notes)
        {
            SectionLastBlocks = sectionLastBlocks,
            SectionTextWidths = sectionTextWidths,
            SectionColumnLefts = sectionColumnLefts,
            Bookmarks = new WordDocumentBookmarkTable(bookmarkNames, bookmarkBlocks),
            IndexEntries = indexEntries,
        };
    }

    /// <summary>Saves the sidecar atomically with its record (<see cref="WordDocumentIndexCache.Save"/>); a failure leaves the old one or none - it is a cache.</summary>
    public void Save(string indexPath, string? documentPath = null) => WordDocumentIndexCache.Save(indexPath, Write, documentPath);

    public static RtfRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader? reader = IndexBinaryArrays.TryOpenInMemory(indexPath);
            if (reader is null) return null;                  // gone, held or pending delete this instant: a cache miss
            RtfRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? WordDocumentIndexCache.MarkUsed(indexPath, index) : null;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            return null;                                     // a corrupt sidecar is only a cache miss
        }
    }
}
