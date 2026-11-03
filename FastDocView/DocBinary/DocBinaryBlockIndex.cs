using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// The compact index of a .doc's main text. It lists no block: the text is random-access through the piece table,
/// so it keeps one <see cref="WordIndexSpanAnchor"/> per ~128K characters (the first block starting past that CP,
/// the block before it, the list-numbering state there) and recovers a block's CP range, kind and list label by
/// rescanning its span with the same <see cref="DocBinaryBlockScanner"/> the index pass ran
/// (<see cref="DocBinaryCompactBlockTable"/>). What a reader needs without any scan stays: the headings (level +
/// text, blocks as Elias-Fano), the block of every note reference, text-length totals per 64 blocks (scroll weights).
/// Persisted as a sidecar for big documents (policy); the text itself is never copied.
/// </summary>
public sealed class DocBinaryBlockIndex
{
    private const string Magic = "KHDOCIDX";
    private const int FormatVersion = 7;        // 5: compact per-block tables; 6: no per-block tables — span anchors + list counter state, Elias-Fano lists; 7: section first blocks
    /// <summary>Characters between span anchors (256 KB of UTF-16 text): a random block rescans at most this much of the piece text.</summary>
    public const int AnchorSpanCharacters = 128 << 10;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public int MainTextCharacterCount { get; }
    public int BlockCount { get; }
    public WordIndexSpanAnchor[] Anchors { get; }
    /// <summary>Sum of the text lengths of blocks [64 i, 64 i + 64).</summary>
    public int[] TextLengthGroupTotals { get; }
    public WordDocumentHeading[] Headings { get; }
    /// <summary>The block of every note reference, in the order of <see cref="DocBinaryNoteReferences.References"/>.</summary>
    public int[] NoteReferenceBlocks { get; }
    /// <summary>The block every section starts in, in the order of <see cref="DocBinarySectionTable.ReadStartCps"/> - a reopen rescans no span for them.</summary>
    public int[] SectionFirstBlocks { get; }
    private readonly WordListCounterState.DeltaChain _listStates;

    private DocBinaryBlockIndex(long sourceFileLength, long sourceLastWriteTicksUtc, int mainTextCharacterCount, int blockCount,
                                WordIndexSpanAnchor[] anchors, int[] textLengthGroupTotals, WordDocumentHeading[] headings, int[] noteReferenceBlocks,
                                int[] sectionFirstBlocks)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        MainTextCharacterCount = mainTextCharacterCount;
        BlockCount = blockCount;
        Anchors = anchors;
        TextLengthGroupTotals = textLengthGroupTotals;
        Headings = headings;
        NoteReferenceBlocks = noteReferenceBlocks;
        SectionFirstBlocks = sectionFirstBlocks;
        _listStates = new WordListCounterState.DeltaChain(anchors.Select(anchor => anchor.ListStateDelta).ToArray());
    }

    /// <summary>The list-numbering state at an anchor (what its span's scan starts counting from).</summary>
    public byte[] ListStateOf(int anchorIndex) => _listStates.StateAt(anchorIndex);

    public int AnchorIndexOf(int blockIndex) => WordIndexSpanAnchor.IndexOfBlock(Anchors, blockIndex);

    /// <summary>The average text length of the block's 64-block group — the scroll weight before its span is scanned.</summary>
    public int EstimatedTextLength(int blockIndex) => WordIndexTextLengthGroups.Average(TextLengthGroupTotals, BlockCount, blockIndex);

    /// <param name="lists">The document's list tables; <see cref="DocBinaryListDefinitions.Empty"/> when it has none.</param>
    /// <param name="noteReferenceCps">The CP of every note reference (body order), so each one's block is recorded.</param>
    /// <param name="sectionStartCps">The start CP of every section (<see cref="DocBinarySectionTable.ReadStartCps"/>), so each one's block is recorded.</param>
    public static DocBinaryBlockIndex Build(DocBinaryPieceTable pieces, DocBinaryParagraphHtmlRenderer renderer, DocBinaryStyleSheet styles,
                                            DocBinaryListDefinitions lists, IReadOnlyList<int> noteReferenceCps, IReadOnlyList<int> sectionStartCps,
                                            int mainTextCharacterCount, long sourceLength, long sourceLastWriteTicksUtc,
                                            IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        var scanner = new DocBinaryBlockScanner(renderer, styles, lists);
        var anchorBlocks = new List<int>();
        var anchorStates = new List<byte[]>();
        scanner.CaptureStateAtOrAfterCp = 0;
        scanner.BlockAdded += (blockIndex, startCp) =>
        {
            if (scanner.StateOfLastBlock is not { } state && blockIndex != 0) return;
            anchorBlocks.Add(blockIndex);
            anchorStates.Add(scanner.StateOfLastBlock ?? []);
            scanner.CaptureStateAtOrAfterCp = startCp + AnchorSpanCharacters;
        };
        pieces.ScanCharacters(0, mainTextCharacterCount, (characters, count, chunkCp) =>
        {
            cancellation.ThrowIfCancellationRequested();
            scanner.Feed(characters, count, chunkCp);
            progress?.Report(new WordDocumentIndexProgress(chunkCp + count, mainTextCharacterCount, scanner.BlockCount));
        });
        scanner.Finish(mainTextCharacterCount);

        List<int> starts = scanner.BlockStartCps;
        int EndOf(int block) => block + 1 < starts.Count ? starts[block + 1] : mainTextCharacterCount;
        byte[][] deltas = WordListCounterState.Deltas(anchorStates);
        var anchors = new WordIndexSpanAnchor[anchorBlocks.Count];
        for (int i = 0; i < anchors.Length; i++)
        {
            int block = anchorBlocks[i];
            anchors[i] = new WordIndexSpanAnchor(block, starts[block], block > 0 ? starts[block - 1] : -1,
                                                 block > 0 ? scanner.BlockKinds[block - 1] : WordDocumentBlockKind.Other, deltas[i]);
        }
        var lengths = new int[starts.Count];
        for (int i = 0; i < lengths.Length; i++) lengths[i] = EndOf(i) - starts[i];
        var noteBlocks = new int[noteReferenceCps.Count];
        for (int i = 0; i < noteBlocks.Length; i++) noteBlocks[i] = BlockAt(starts, noteReferenceCps[i]);
        var sectionBlocks = new int[sectionStartCps.Count];
        for (int i = 0; i < sectionBlocks.Length; i++) sectionBlocks[i] = BlockAt(starts, Math.Min(sectionStartCps[i], mainTextCharacterCount));
        return new DocBinaryBlockIndex(sourceLength, sourceLastWriteTicksUtc, mainTextCharacterCount, starts.Count, anchors,
                                       WordIndexTextLengthGroups.Totals(lengths), scanner.Headings.ToArray(), noteBlocks, sectionBlocks);
    }

    /// <summary>The block containing a CP, among sorted block starts.</summary>
    internal static int BlockAt(IReadOnlyList<int> starts, int cp)
    {
        int low = 0, high = starts.Count - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (starts[middle] <= cp) low = middle; else high = middle - 1;
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
        writer.Write(MainTextCharacterCount);
        writer.Write(BlockCount);
        Section("header");
        WordIndexSpanAnchor.WriteAll(writer, Anchors);
        Section($"anchors ({Anchors.Length}, list state)");
        IndexBinaryArrays.WriteVarints(writer, TextLengthGroupTotals);
        Section("text-length groups");
        new EliasFanoSequence(Headings.Select(h => h.BlockIndex).ToArray()).Write(writer);
        foreach (WordDocumentHeading heading in Headings) writer.Write((byte)heading.Level);
        IndexBinaryArrays.WriteCompressedStrings(writer, Headings.Select(h => h.Text).ToArray());
        Section($"headings ({Headings.Length})");
        new EliasFanoSequence(NoteReferenceBlocks).Write(writer);
        Section($"note references ({NoteReferenceBlocks.Length})");
        new EliasFanoSequence(SectionFirstBlocks).Write(writer);
        Section($"section first blocks ({SectionFirstBlocks.Length})");
        writer.Flush();
    }

    public static DocBinaryBlockIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not a .doc index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException(".doc index: unsupported version.");
        long length = reader.ReadInt64(), ticks = reader.ReadInt64();
        int mainText = reader.ReadInt32();
        int blockCount = reader.ReadInt32();
        if (blockCount < 0 || mainText < 0) throw new InvalidDataException(".doc index: negative count.");
        WordIndexSpanAnchor[] anchors = WordIndexSpanAnchor.ReadAll(reader, blockCount, mainText);
        int[] groupTotals = IndexBinaryArrays.ReadVarints(reader);
        WordIndexTextLengthGroups.Validate(groupTotals, blockCount);
        int[] headingBlocks = EliasFanoSequence.Read(reader).ToArray();
        var levels = new byte[headingBlocks.Length];
        for (int i = 0; i < levels.Length; i++) levels[i] = reader.ReadByte();
        string[] texts = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (texts.Length != headingBlocks.Length || headingBlocks.Any(block => block >= blockCount)) throw new InvalidDataException(".doc index: headings do not match the blocks.");
        var headings = new WordDocumentHeading[headingBlocks.Length];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(headingBlocks[i], levels[i], texts[i]);
        int[] noteBlocks = EliasFanoSequence.Read(reader).ToArray();
        if (noteBlocks.Any(block => block >= Math.Max(1, blockCount))) throw new InvalidDataException(".doc index: note references do not match the blocks.");
        int[] sectionBlocks = EliasFanoSequence.Read(reader).ToArray();
        if (sectionBlocks.Any(block => block >= Math.Max(1, blockCount))) throw new InvalidDataException(".doc index: section starts do not match the blocks.");
        return new DocBinaryBlockIndex(length, ticks, mainText, blockCount, anchors, groupTotals, headings, noteBlocks, sectionBlocks);
    }

    /// <summary>Saves the sidecar atomically with its record (<see cref="WordDocumentIndexCache.Save"/>); a failure leaves the old one or none - it is a cache.</summary>
    public void Save(string indexPath, string? documentPath = null) => WordDocumentIndexCache.Save(indexPath, Write, documentPath);

    public static DocBinaryBlockIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader? reader = IndexBinaryArrays.TryOpenInMemory(indexPath);
            if (reader is null) return null;                  // gone, held or pending delete this instant: a cache miss
            DocBinaryBlockIndex index = Read(reader.BaseStream);
            return index.MatchesSource(documentPath) ? WordDocumentIndexCache.MarkUsed(indexPath, index) : null;
        }
        catch (Exception exception) when (exception is IOException or EndOfStreamException or InvalidDataException or OverflowException or ArgumentException or FormatException)
        {
            return null;
        }
    }
}
