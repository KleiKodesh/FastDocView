using System.Text;
using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// The compact index of one .docx body. It no longer lists every block: it holds the deflate checkpoints
/// (positions + all sparse windows in one store), and per <b>span</b> (the blocks from the first block at
/// or after a checkpoint up to the next span's first block) one <see cref="DocxIndexAnchor"/>: that block's
/// index and offset, the block before it (so a window starting there needs no rescan of the previous
/// span), and the list-numbering state at that point. A block's position, kind, style, list label and
/// text length are recovered by rescanning its span from its anchor with the same
/// <see cref="DocxBodyBlockScanner"/> that built the index (<see cref="DocxCompactBlockTable"/>) — a few
/// ms, and the same bytes the renderer is about to read anyway. Index-time lists that a reader needs
/// without any scan stay: headings, note references, sections, bookmarks, the fields open at a block,
/// text-length totals per 64 blocks (scroll weights). Measured on a 91 MB body: ~45 KB, where the
/// per-block layout took 1.17 MB.
/// </summary>
public sealed class DocxRandomAccessIndex
{
    private const string Magic = "KHWORDIDX";
    private const int FormatVersion = 22;     // 22: per-section footnote / endnote numbering (numFmt / numStart / numRestart); 21: a symbol bullet resolved against the level's Latin font only (not w:cs); 20: list label CSS for both the Latin and the complex-script set; 19: picture bullets (numPicBullet); 18: list labels in every Word number format, symbol-font bullets resolved; 17: gutter placement (w:rtlGutter / w:gutterAtTop) in section widths; 16: section column lefts; 15: XE / TA entries;     // 14: field counters (SEQ / AUTONUM / LISTNUM) in the open-field stacks; 12: compact per-block tables; 13: no per-block tables — span anchors + list counter state, one window store, Elias-Fano block lists
    public const int TextLengthGroupSize = 64;
    /// <summary>Output bytes between span anchors: a random block rescans at most this much (~1-2 ms) on top of the inflate from its checkpoint.</summary>
    public const long AnchorSpanBytes = 256 << 10;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public ZipEntryLocation DocumentEntry { get; }
    public string RootOpenTag { get; }
    public long BodyEndOffset { get; }
    public IReadOnlyList<DeflateCheckpoint> Checkpoints { get; }
    public int BlockCount { get; }
    public DocxIndexAnchor[] Anchors { get; }
    /// <summary>Sum of the text lengths of blocks [64 i, 64 i + 64).</summary>
    public int[] TextLengthGroupTotals { get; }
    public WordDocumentHeading[] Headings { get; }
    public DocxNoteReference[] NoteReferences { get; }      // in body order
    public DocxNotesPartIndex? FootnotesPart { get; }
    public DocxNotesPartIndex? EndnotesPart { get; }
    /// <summary>Blocks (ascending) where a complex field is open at the start, and the index into <see cref="OpenFieldStacks"/> for each.</summary>
    public int[] OpenFieldBlocks { get; private set; } = [];
    public int[] OpenFieldIds { get; private set; } = [];
    public string[] OpenFieldStacks { get; private set; } = [];
    /// <summary>Per section: its last block (inclusive) and its page text width in twips (see DocxBodyBlockScanner).</summary>
    public int[] SectionLastBlocks { get; private set; } = [];
    public int[] SectionTextWidths { get; private set; } = [];
    public int[] SectionColumnLefts { get; private set; } = [];
    /// <summary>Per section: its footnote and endnote numbering rules (<see cref="WordDocumentNoteNumberRule.Serialize"/>).</summary>
    public WordDocumentNoteNumberRule[] SectionFootnoteRules { get; private set; } = [];
    public WordDocumentNoteNumberRule[] SectionEndnoteRules { get; private set; } = [];
    /// <summary>Bookmarks in body order: name and the block it marks (see DocxBodyBlockScanner.BookmarkNames).</summary>
    public string[] BookmarkNames { get; private set; } = [];
    public int[] BookmarkBlocks { get; private set; } = [];

    /// <summary>XE / TA field codes with their blocks ("block U+0001 code", DocxBodyBlockScanner.IndexEntries).</summary>
    public string[] IndexEntries { get; private set; } = [];

    internal DocxRandomAccessIndex WithIndexEntries(string[] entries)
    {
        IndexEntries = entries;
        return this;
    }
    private Dictionary<string, int>? _bookmarkLookup;

    public DocxRandomAccessIndex(long sourceFileLength, long sourceLastWriteTicksUtc, ZipEntryLocation documentEntry,
                                 string rootOpenTag, long bodyEndOffset, IReadOnlyList<DeflateCheckpoint> checkpoints,
                                 int blockCount, DocxIndexAnchor[] anchors, int[] textLengthGroupTotals,
                                 WordDocumentHeading[] headings, DocxNoteReference[] noteReferences,
                                 DocxNotesPartIndex? footnotesPart, DocxNotesPartIndex? endnotesPart)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        DocumentEntry = documentEntry;
        RootOpenTag = rootOpenTag;
        BodyEndOffset = bodyEndOffset;
        Checkpoints = checkpoints;
        BlockCount = blockCount;
        Anchors = anchors;
        TextLengthGroupTotals = textLengthGroupTotals;
        Headings = headings;
        NoteReferences = noteReferences;
        FootnotesPart = footnotesPart;
        EndnotesPart = endnotesPart;
    }

    internal DocxRandomAccessIndex WithBookmarks(string[] names, int[] blocks)
    {
        BookmarkNames = names;
        BookmarkBlocks = blocks;
        _bookmarkLookup = null;
        return this;
    }

    internal DocxRandomAccessIndex WithSections(int[] sectionLastBlocks, int[] sectionTextWidths, int[] sectionColumnLefts,
                                                WordDocumentNoteNumberRule[] footnoteRules, WordDocumentNoteNumberRule[] endnoteRules)
    {
        SectionFootnoteRules = footnoteRules;
        SectionEndnoteRules = endnoteRules;
        SectionLastBlocks = sectionLastBlocks;
        SectionTextWidths = sectionTextWidths;
        SectionColumnLefts = sectionColumnLefts;
        return this;
    }

    /// <summary>From the scanner's per-block ids (-1 = none): only the set blocks are kept.</summary>
    internal DocxRandomAccessIndex WithOpenFields(IReadOnlyList<int> blockOpenFieldIds, string[] openFieldStacks)
    {
        var blocks = new List<int>();
        var ids = new List<int>();
        for (int block = 0; block < blockOpenFieldIds.Count; block++)
            if (blockOpenFieldIds[block] >= 0) { blocks.Add(block); ids.Add(blockOpenFieldIds[block]); }
        OpenFieldBlocks = blocks.ToArray();
        OpenFieldIds = ids.ToArray();
        OpenFieldStacks = openFieldStacks;
        return this;
    }

    private DocxRandomAccessIndex WithOpenFieldPairs(int[] blocks, int[] ids, string[] stacks)
    {
        OpenFieldBlocks = blocks;
        OpenFieldIds = ids;
        OpenFieldStacks = stacks;
        return this;
    }

    /// <summary>The fields open where a block starts (see DocxBodyBlockScanner.OpenFieldStacks), or null.</summary>
    public string? OpenFieldsAt(int blockIndex)
    {
        int position = Array.BinarySearch(OpenFieldBlocks, blockIndex);
        return position >= 0 ? OpenFieldStacks[OpenFieldIds[position]] : null;
    }

    /// <summary>The block a bookmark marks (the first one of that name, as Word resolves duplicates), or null.</summary>
    public int? BlockOfBookmark(string name)
    {
        if (_bookmarkLookup is null)
        {
            var lookup = new Dictionary<string, int>(BookmarkNames.Length, StringComparer.Ordinal);
            for (int i = 0; i < BookmarkNames.Length; i++)
                if (!lookup.ContainsKey(BookmarkNames[i])) lookup[BookmarkNames[i]] = Math.Min(BookmarkBlocks[i], Math.Max(0, BlockCount - 1));
            _bookmarkLookup = lookup;
        }
        return _bookmarkLookup.TryGetValue(name, out int block) ? block : null;
    }

    private byte[]?[]? _anchorStates;

    /// <summary>
    /// The list-numbering state at anchor <paramref name="anchorIndex"/> (what the span's scan starts counting from): the
    /// deltas applied from the start, each anchor's full state kept once computed (a few hundred bytes per anchor).
    /// </summary>
    public byte[] ListStateOf(int anchorIndex)
    {
        _anchorStates ??= new byte[]?[Anchors.Length];
        if (_anchorStates[anchorIndex] is { } known) return known;
        int from = anchorIndex;
        while (from > 0 && _anchorStates[from - 1] is null) from--;
        byte[] current = from > 0 ? _anchorStates[from - 1]! : [];
        for (int i = from; i <= anchorIndex; i++) _anchorStates[i] = current = DocxHtml.DocxListCounter.ApplyDelta(current, Anchors[i].ListStateDelta);
        return current;
    }

    /// <summary>The anchor (span) a block belongs to: the last anchor whose first block is ≤ the block.</summary>
    public int AnchorIndexOf(int blockIndex)
    {
        int low = 0, high = Anchors.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (Anchors[middle].FirstBlock <= blockIndex) low = middle; else high = middle - 1;
        }
        return low;
    }

    /// <summary>The average text length of the block's 64-block group — the scroll weight before the span is scanned.</summary>
    public int EstimatedTextLength(int blockIndex)
    {
        if ((uint)blockIndex >= (uint)BlockCount) return 0;
        int group = blockIndex / TextLengthGroupSize;
        int members = Math.Min(TextLengthGroupSize, BlockCount - group * TextLengthGroupSize);
        return TextLengthGroupTotals[group] / members;
    }

    /// <summary>The note references inside blocks [first, last] — a slice of the body-ordered array.</summary>
    public ReadOnlySpan<DocxNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex)
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

    // ── serialisation ─────────────────────────────────────────────────────────────────

    /// <summary>Bytes per section of the serialised index — for size reports.</summary>
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
        WriteEntry(writer, DocumentEntry);
        IndexBinaryArrays.WriteCompressedString(writer, RootOpenTag);
        writer.Write(BodyEndOffset);
        writer.Write(BlockCount);
        Section("header");
        DeflateCheckpointTable.Write(writer, Checkpoints);
        Section("checkpoints + windows");
        DocxIndexAnchor.WriteAll(writer, Anchors);
        Section("anchors (list state)");
        IndexBinaryArrays.WriteVarints(writer, TextLengthGroupTotals);
        Section("text-length groups");

        // headings: blocks as Elias-Fano, then level + text each
        new EliasFanoSequence(Headings.Select(h => h.BlockIndex).ToArray()).Write(writer);
        foreach (WordDocumentHeading heading in Headings) { writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        Section("headings");
        // note references: blocks as Elias-Fano; kind, id and number as varints (ids / numbers delta-coded per kind: Word numbers in order)
        new EliasFanoSequence(NoteReferences.Select(r => r.BlockIndex).ToArray()).Write(writer);
        var noteBytes = new List<byte>(NoteReferences.Length * 3);
        var previousId = new int[2];
        var previousNumber = new int[2];
        foreach (DocxNoteReference reference in NoteReferences)
        {
            int kind = (int)reference.Kind;
            noteBytes.Add((byte)kind);
            IndexBinaryArrays.AppendVarint(noteBytes, Zigzag(reference.NoteId - previousId[kind]));
            IndexBinaryArrays.AppendVarint(noteBytes, Zigzag(reference.Number - previousNumber[kind]));
            previousId[kind] = reference.NoteId;
            previousNumber[kind] = reference.Number;
        }
        writer.Write(noteBytes.Count);
        writer.Write(noteBytes.ToArray());
        Section("note references");
        foreach (DocxNotesPartIndex? part in new[] { FootnotesPart, EndnotesPart })
        {
            writer.Write(part is not null);
            part?.Write(writer);
            Section(part is null ? "no notes part" : $"notes part {part.Entry.Name} ({part.NoteIds.Length} notes, {part.Checkpoints.Count} checkpoints)");
        }
        IndexBinaryArrays.WriteCompressedStrings(writer, OpenFieldStacks);
        IndexBinaryArrays.WriteDeltaVarints(writer, OpenFieldBlocks);
        IndexBinaryArrays.WriteVarints(writer, OpenFieldIds);
        Section($"open fields ({OpenFieldBlocks.Length} blocks)");
        IndexBinaryArrays.WriteDeltaVarints(writer, SectionLastBlocks);
        // widths: a few distinct values over many sections -> a table + byte ids
        int[] widthTable = SectionTextWidths.Distinct().OrderBy(w => w).ToArray();
        IndexBinaryArrays.WriteVarints(writer, widthTable);
        IndexBinaryArrays.WriteByteIds(writer, SectionTextWidths.Select(w => (short)Array.IndexOf(widthTable, w)).ToArray());
        IndexBinaryArrays.WriteVarints(writer, SectionColumnLefts);
        IndexBinaryArrays.WriteCompressedStrings(writer, SectionFootnoteRules.Select(rule => rule.Serialize()).ToArray());
        IndexBinaryArrays.WriteCompressedStrings(writer, SectionEndnoteRules.Select(rule => rule.Serialize()).ToArray());
        Section("sections");
        IndexBinaryArrays.WriteCompressedStrings(writer, BookmarkNames);
        new EliasFanoSequence(BookmarkBlocks).Write(writer);
        Section($"bookmarks ({BookmarkNames.Length})");
        IndexBinaryArrays.WriteCompressedStrings(writer, IndexEntries);
        Section($"index entries ({IndexEntries.Length})");
        writer.Flush();
    }

    public static DocxRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not a word index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("Word index: unsupported version.");
        long sourceLength = reader.ReadInt64();
        long sourceTicks = reader.ReadInt64();
        ZipEntryLocation entry = ReadEntry(reader);
        string rootOpenTag = IndexBinaryArrays.ReadCompressedString(reader);
        long bodyEnd = reader.ReadInt64();
        int blockCount = reader.ReadInt32();
        if (blockCount < 0) throw new InvalidDataException("Word index: negative block count.");
        List<DeflateCheckpoint> checkpoints = DeflateCheckpointTable.Read(reader);
        DocxIndexAnchor[] anchors = DocxIndexAnchor.ReadAll(reader, blockCount);
        int[] groupTotals = IndexBinaryArrays.ReadVarints(reader);
        if (groupTotals.Length != (blockCount + TextLengthGroupSize - 1) / TextLengthGroupSize) throw new InvalidDataException("Word index: text-length groups do not match the blocks.");

        int[] headingBlocks = EliasFanoSequence.Read(reader).ToArray();
        var headings = new WordDocumentHeading[headingBlocks.Length];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(headingBlocks[i], reader.ReadByte(), reader.ReadString());
        int[] noteBlocks = EliasFanoSequence.Read(reader).ToArray();
        byte[] noteBytes = reader.ReadBytes(reader.ReadInt32());
        var noteReferences = new DocxNoteReference[noteBlocks.Length];
        {
            int position = 0;
            var previousId = new int[2];
            var previousNumber = new int[2];
            for (int i = 0; i < noteReferences.Length; i++)
            {
                if (position >= noteBytes.Length) throw new EndOfStreamException("Word index: note references truncated.");
                int kind = noteBytes[position++];
                if (kind > 1) throw new InvalidDataException("Word index: bad note kind.");
                previousId[kind] += Unzigzag(IndexBinaryArrays.ReadVarint(noteBytes, ref position));
                previousNumber[kind] += Unzigzag(IndexBinaryArrays.ReadVarint(noteBytes, ref position));
                noteReferences[i] = new DocxNoteReference(noteBlocks[i], (WordDocumentNoteKind)kind, previousId[kind], previousNumber[kind]);
            }
        }
        DocxNotesPartIndex? footnotesPart = reader.ReadBoolean() ? DocxNotesPartIndex.Read(reader) : null;
        DocxNotesPartIndex? endnotesPart = reader.ReadBoolean() ? DocxNotesPartIndex.Read(reader) : null;
        string[] openFieldStacks = IndexBinaryArrays.ReadCompressedStrings(reader);
        int[] openFieldBlocks = IndexBinaryArrays.ReadDeltaVarintsInt32(reader);
        int[] openFieldIds = IndexBinaryArrays.ReadVarints(reader);
        if (openFieldIds.Length != openFieldBlocks.Length || openFieldIds.Any(id => id >= openFieldStacks.Length) || openFieldBlocks.Any(block => block >= blockCount))
            throw new InvalidDataException("Word index: open-field table does not match the blocks.");
        int[] sectionLastBlocks = IndexBinaryArrays.ReadDeltaVarintsInt32(reader);
        int[] widthTable = IndexBinaryArrays.ReadVarints(reader);
        short[] widthIds = IndexBinaryArrays.ReadByteIds(reader);
        if (widthIds.Any(id => id < 0 || id >= widthTable.Length)) throw new InvalidDataException("Word index: section widths out of range.");
        int[] sectionTextWidths = widthIds.Select(id => widthTable[id]).ToArray();
        int[] sectionColumnLefts = IndexBinaryArrays.ReadVarints(reader);
        WordDocumentNoteNumberRule[] footnoteRules = IndexBinaryArrays.ReadCompressedStrings(reader).Select(WordDocumentNoteNumberRule.Parse).ToArray();
        WordDocumentNoteNumberRule[] endnoteRules = IndexBinaryArrays.ReadCompressedStrings(reader).Select(WordDocumentNoteNumberRule.Parse).ToArray();
        if (sectionLastBlocks.Length != sectionTextWidths.Length || sectionColumnLefts.Length != sectionTextWidths.Length || sectionLastBlocks.Any(block => (uint)block >= (uint)blockCount)
            || footnoteRules.Length != sectionLastBlocks.Length || endnoteRules.Length != sectionLastBlocks.Length)
            throw new InvalidDataException("Word index: section table does not match the blocks.");
        string[] bookmarkNames = IndexBinaryArrays.ReadCompressedStrings(reader);
        int[] bookmarkBlocks = EliasFanoSequence.Read(reader).ToArray();
        if (bookmarkBlocks.Length != bookmarkNames.Length || bookmarkBlocks.Any(block => block < 0 || block > blockCount))
            throw new InvalidDataException("Word index: bookmark table does not match the blocks.");
        string[] indexEntries = IndexBinaryArrays.ReadCompressedStrings(reader);

        return new DocxRandomAccessIndex(sourceLength, sourceTicks, entry, rootOpenTag, bodyEnd, checkpoints, blockCount, anchors, groupTotals,
                                         headings, noteReferences, footnotesPart, endnotesPart)
            .WithOpenFieldPairs(openFieldBlocks, openFieldIds, openFieldStacks)
            .WithSections(sectionLastBlocks, sectionTextWidths, sectionColumnLefts, footnoteRules, endnoteRules)
            .WithBookmarks(bookmarkNames, bookmarkBlocks)
            .WithIndexEntries(indexEntries);
    }

    internal static ulong Zigzag(int value) => (ulong)(uint)((value << 1) ^ (value >> 31));
    internal static int Unzigzag(ulong value) => (int)((uint)value >> 1) ^ -(int)((uint)value & 1);

    // ── shared record layouts (also used by the .odt index) ───────────────────────────

    internal static void WriteEntry(BinaryWriter writer, ZipEntryLocation entry)
    {
        writer.Write(entry.Name);
        writer.Write(entry.DataOffset);
        writer.Write(entry.CompressedSize);
        writer.Write(entry.UncompressedSize);
        writer.Write(entry.CompressionMethod);
        writer.Write(entry.Crc32);
    }

    internal static ZipEntryLocation ReadEntry(BinaryReader reader)
        => new(reader.ReadString(), reader.ReadInt64(), reader.ReadInt64(), reader.ReadInt64(), reader.ReadUInt16(), reader.ReadUInt32());

    /// <summary>The per-checkpoint layout (each window deflated on its own), still used by the .odt index.</summary>
    internal static void WriteCheckpoints(BinaryWriter writer, IReadOnlyList<DeflateCheckpoint> checkpoints)
    {
        writer.Write(checkpoints.Count);
        foreach (DeflateCheckpoint checkpoint in checkpoints)
        {
            writer.Write(checkpoint.CompressedBytePosition);
            writer.Write((byte)checkpoint.LeftoverBitCount);
            writer.Write((byte)checkpoint.LeftoverBits);
            writer.Write(checkpoint.UncompressedOffset);
            byte[] blob = checkpoint.CompressedWindow;                         // sparse + deflated window (see DeflateCheckpoint)
            var lengths = new List<byte>(8);
            IndexBinaryArrays.AppendVarint(lengths, (uint)checkpoint.WindowLength);
            IndexBinaryArrays.AppendVarint(lengths, (uint)blob.Length);
            writer.Write(lengths.ToArray());
            writer.Write(blob);
        }
    }

    internal static List<DeflateCheckpoint> ReadCheckpoints(BinaryReader reader)
    {
        int count = reader.ReadInt32();
        var checkpoints = new List<DeflateCheckpoint>(count);
        for (int i = 0; i < count; i++)
        {
            long compressedPosition = reader.ReadInt64();
            int leftoverBitCount = reader.ReadByte();
            int leftoverBits = reader.ReadByte();
            long uncompressedOffset = reader.ReadInt64();
            int windowLength = (int)ReadVarint(reader);
            int blobLength = (int)ReadVarint(reader);
            if (windowLength > DeflateCheckpoint.WindowSize || blobLength < 0) throw new InvalidDataException("Word index: checkpoint window out of range.");
            byte[] window = reader.ReadBytes(blobLength);
            if (window.Length != blobLength) throw new EndOfStreamException("Word index: checkpoint truncated.");
            checkpoints.Add(new DeflateCheckpoint(compressedPosition, leftoverBitCount, leftoverBits, uncompressedOffset, windowLength, window));
        }
        return checkpoints;
    }

    internal static ulong ReadVarint(BinaryReader reader)
    {
        ulong value = 0;
        int shift = 0;
        while (true)
        {
            byte b = reader.ReadByte();
            value |= (ulong)(b & 0x7F) << shift;
            if (b < 0x80) return value;
            shift += 7;
            if (shift > 63) throw new InvalidDataException("Word index: varint too long.");
        }
    }

    // ── sidecar file helpers (file-system hosts) ───────────────────────────────────────

    /// <summary>Saves the sidecar atomically with its record (<see cref="WordDocumentIndexCache.Save"/>); a failure leaves the old one or none - it is a cache.</summary>
    public void Save(string indexPath, string? documentPath = null) => WordDocumentIndexCache.Save(indexPath, Write, documentPath);

    public static DocxRandomAccessIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader? reader = IndexBinaryArrays.TryOpenInMemory(indexPath);
            if (reader is null) return null;                  // gone, held or pending delete this instant: a cache miss
            DocxRandomAccessIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? WordDocumentIndexCache.MarkUsed(indexPath, index) : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException)
        {
            return null;
        }
    }
}

/// <summary>
/// One span of the compact index: where the rescan of blocks [FirstBlock, next anchor's FirstBlock) starts
/// (<see cref="Offset"/> = that block's start, at or after a deflate checkpoint), the block just before it
/// (offset, kind, paragraph style — a window that starts at this span's first block reads its predecessor as a
/// neighbour without scanning the previous span), and the list numbering state the scan continues from.
/// </summary>
public sealed record DocxIndexAnchor(int FirstBlock, long Offset, long PreviousBlockOffset, WordDocumentBlockKind PreviousBlockKind,
                                     string? PreviousStyleId, byte[] ListStateDelta)
{
    internal static void WriteAll(BinaryWriter writer, DocxIndexAnchor[] anchors)
    {
        writer.Write(anchors.Length);
        int previousBlock = 0;
        long previousOffset = 0;
        foreach (DocxIndexAnchor anchor in anchors)
        {
            byte[] delta = anchor.ListStateDelta;
            var bytes = new List<byte>(24);
            IndexBinaryArrays.AppendVarint(bytes, (uint)(anchor.FirstBlock - previousBlock));
            IndexBinaryArrays.AppendVarint(bytes, (ulong)(anchor.Offset - previousOffset));
            IndexBinaryArrays.AppendVarint(bytes, anchor.PreviousBlockOffset < 0 ? 0UL : (ulong)(anchor.Offset - anchor.PreviousBlockOffset));   // 0 = no previous block
            bytes.Add((byte)anchor.PreviousBlockKind);
            writer.Write(bytes.ToArray());
            writer.Write(anchor.PreviousStyleId ?? "");
            writer.Write(delta.Length);
            writer.Write(delta);
            previousBlock = anchor.FirstBlock;
            previousOffset = anchor.Offset;
        }
    }

    internal static DocxIndexAnchor[] ReadAll(BinaryReader reader, int blockCount)
    {
        int count = reader.ReadInt32();
        if (count < 0 || (count == 0) != (blockCount == 0)) throw new InvalidDataException("Word index: anchors do not match the blocks.");
        var anchors = new DocxIndexAnchor[count];
        int previousBlock = 0;
        long previousOffset = 0;
        for (int i = 0; i < count; i++)
        {
            int firstBlock = previousBlock + (int)DocxRandomAccessIndex.ReadVarint(reader);
            long offset = previousOffset + (long)DocxRandomAccessIndex.ReadVarint(reader);
            long back = (long)DocxRandomAccessIndex.ReadVarint(reader);
            var kind = (WordDocumentBlockKind)reader.ReadByte();
            string style = reader.ReadString();
            byte[] state = reader.ReadBytes(reader.ReadInt32());
            if (firstBlock >= blockCount || (i > 0 && firstBlock <= previousBlock) || (i == 0 && firstBlock != 0)) throw new InvalidDataException("Word index: anchor blocks out of order.");
            anchors[i] = new DocxIndexAnchor(firstBlock, offset, back == 0 ? -1 : offset - back, kind, style.Length == 0 ? null : style, state);
            previousBlock = firstBlock;
            previousOffset = offset;
        }
        return anchors;
    }
}
