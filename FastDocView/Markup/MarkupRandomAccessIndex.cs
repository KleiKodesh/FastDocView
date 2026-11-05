using System.Text;

namespace FastDocView.Markup;

public enum MarkupSourceKind : byte { PlainText = 0, WebArchive = 1 } // 2 was loose .html; dropped: real web pages need scripts and live resources we do not run

/// <summary>
/// One span of a compact markup index: where the rescan of blocks [<see cref="FirstBlock"/>, next anchor's FirstBlock)
/// starts. Anchor 0 has no <see cref="State"/>: its rescan is a fresh scanner from the text's start. Every other anchor
/// is a <see cref="MarkupBlockScanner"/> span point (<see cref="MarkupScanState"/>, taken between two fed pieces) at
/// <see cref="ScanPosition"/> (text coordinates). MHTML: <see cref="ScanResume"/> restarts the transfer decoder at
/// ScanPosition, <see cref="BlockResume"/> at the open block's start (where a window reading that block begins).
/// </summary>
public sealed record MarkupSpanAnchor(MarkupScanState? State, long ScanPosition, long ScanResume, long BlockResume)
{
    public int FirstBlock => State?.Block ?? 0;
    public long FirstBlockStart => State?.BlockStart ?? ScanPosition;
}

/// <summary>
/// The index of a plain-text (.txt read as HTML) or web-archive (MHTML) document - compact (format 6): no per-block
/// table. Text offsets are "text coordinates": file offsets for .txt, decoded offsets of the HTML part for MHTML.
/// It keeps one <see cref="MarkupSpanAnchor"/> per ~64 KB of text (MHTML: ~16 KB), text-length totals per 64 blocks (scroll weights),
/// the interned element stacks, the headings and the head's CSS. A block's start, end, text length, open-element stack
/// (and MHTML decoder resume point) are recovered by rescanning its span (<see cref="MarkupCompactBlockTable"/>) with
/// the same scanner the index pass ran, so every value equals the one the pass found.
/// </summary>
public sealed class MarkupRandomAccessIndex
{
    private const string Magic = "KHMKPIDX";
    private const int FormatVersion = 11; // 11: blank lines dropped unless PutNewLinesAsIs=1, tilde headings return when the contents would be empty; 10: (also covers 8, the path rule: an empty head, no head block) PutNewLinesAsIs=0 books drop blank lines, LastLevelIndex 0 hides tilde headings; 9: heading levels 1-4 (dollar = 1), label lines merge with the next line; 8: Torat Emet without line-1 metadata (empty head) and without a title line; 7: Torat Emet head (metadata + where it ends; heading levels from line marks); 6: compact (anchors + rescans, no per-block tables); 5: compact block tables; 2: .txt encoding by TextEncodingDetection (Notepad++ parity); 3: its Improved mode; 4: .html kind
    public const int AnchorSpanBytes = 64 * 1024;
    /// <summary>
    /// MHTML anchors are 4x denser: decoding plus tag-heavy HTML rescans at ~25 MB/s, so a 64 KB span cost a random block
    /// ~2.5 ms; at 16 KB the anchors of a 300k-block archive are still under 2 KB. (A .txt anchor stays at a 64 KB piece
    /// boundary; its rescan is a vectorised run search.) An index with other spacing still reads: the anchors are stored.
    /// </summary>
    public const int WebArchiveAnchorSpanBytes = 16 * 1024;

    public MarkupSourceKind Kind { get; }
    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public int CodePage { get; }
    public string EncodingReason { get; }
    public MimeTransferEncoding TransferEncoding { get; }
    public long TextEnd { get; }                      // end of the last block, text coordinates
    public long EncodedEnd { get; }                   // MHTML: end of the HTML part body in the file
    public bool IsRightToLeft { get; }
    public string DocumentCss { get; }                // MHTML: the head's style text, scoped
    /// <summary>.txt: the first text byte (after a BOM). MHTML: the decoder resume point of the HTML part's first byte.</summary>
    public long TextStart { get; }
    /// <summary>.txt: where the index pass's first fed piece ended (the detection sample); later pieces are <see cref="MarkupRandomAccessIndexBuilder.ChunkSize"/> each - a rescan feeds the same pieces.</summary>
    public long FirstFeedEnd { get; }
    public int BlockCount { get; }
    public MarkupSpanAnchor[] Anchors { get; }
    public int[] TextLengthGroupTotals { get; }
    public List<string> Stacks { get; }
    public WordDocumentHeading[] Headings { get; }
    /// <summary>.txt in the Torat Emet book format: its metadata head (one hidden block, block 0) and the keys; null for every other text.</summary>
    public ToratEmetHead? ToratEmet { get; }
    /// <summary>Where the Torat Emet head ends (text coordinates), -1 for any other text: what a scanner of this index is told.</summary>
    /// <summary>
    /// The index of the blocks known so far while an on-demand first open is still scanning (never saved): its last span
    /// ends at <see cref="TextEnd"/>, where the pass had read to, past the start of the block still being read.
    /// </summary>
    public bool IsPartial { get; init; }

    private Dictionary<string, int>? _stackIds;
    /// <summary>The interned stacks by text (what a span rescan's scanner shares), built on first use.</summary>
    internal Dictionary<string, int> StackIds
    {
        get
        {
            if (_stackIds is null)
            {
                var ids = new Dictionary<string, int>(Stacks.Count, StringComparer.Ordinal);
                for (int i = 0; i < Stacks.Count; i++) ids[Stacks[i]] = i;
                _stackIds = ids;
            }
            return _stackIds;
        }
    }

    public MarkupRandomAccessIndex(MarkupSourceKind kind, long sourceFileLength, long sourceLastWriteTicksUtc, int codePage, string encodingReason,
                                   MimeTransferEncoding transferEncoding, long textEnd, long encodedEnd, bool isRightToLeft, string documentCss,
                                   long textStart, long firstFeedEnd, int blockCount, MarkupSpanAnchor[] anchors, int[] textLengthGroupTotals,
                                   List<string> stacks, WordDocumentHeading[] headings, ToratEmetHead? toratEmet = null)
    {
        ToratEmet = toratEmet;
        Kind = kind; SourceFileLength = sourceFileLength; SourceLastWriteTicksUtc = sourceLastWriteTicksUtc; CodePage = codePage;
        EncodingReason = encodingReason; TransferEncoding = transferEncoding; TextEnd = textEnd; EncodedEnd = encodedEnd;
        IsRightToLeft = isRightToLeft; DocumentCss = documentCss; TextStart = textStart; FirstFeedEnd = firstFeedEnd;
        BlockCount = blockCount; Anchors = anchors; TextLengthGroupTotals = textLengthGroupTotals; Stacks = stacks; Headings = headings;
    }

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

    public int EstimatedTextLength(int blockIndex) => WordIndexTextLengthGroups.Average(TextLengthGroupTotals, BlockCount, blockIndex);

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

    private static ulong ZigZag(long value) => (ulong)((value << 1) ^ (value >> 63));
    private static long UnZigZag(ulong value) => (long)(value >> 1) ^ -(long)(value & 1);

    private void Write(Stream destination, List<(string, long)>? sizes)
    {
        var writer = new BinaryWriter(destination, Encoding.UTF8, leaveOpen: true);
        long mark = destination.Position;
        void Section(string name) { writer.Flush(); sizes?.Add((name, destination.Position - mark)); mark = destination.Position; }
        writer.Write(Magic);
        writer.Write(FormatVersion);
        writer.Write((byte)Kind);
        writer.Write(SourceFileLength);
        writer.Write(SourceLastWriteTicksUtc);
        writer.Write(CodePage);
        writer.Write(EncodingReason);
        writer.Write((byte)TransferEncoding);
        writer.Write(TextEnd);
        writer.Write(EncodedEnd);
        writer.Write(IsRightToLeft);
        writer.Write(TextStart);
        writer.Write(FirstFeedEnd);
        writer.Write(BlockCount);
        Section("header");
        WriteStrings(writer, [DocumentCss]);
        Section("document CSS");

        var raw = new List<byte>(Anchors.Length * 24);
        IndexBinaryArrays.AppendVarint(raw, (uint)Anchors.Length);
        MarkupSpanAnchor? previous = null;
        foreach (MarkupSpanAnchor anchor in Anchors)
        {
            IndexBinaryArrays.AppendVarint(raw, ZigZag(anchor.ScanPosition - (previous?.ScanPosition ?? 0)));
            IndexBinaryArrays.AppendVarint(raw, ZigZag(anchor.ScanResume - (previous?.ScanResume ?? 0)));
            IndexBinaryArrays.AppendVarint(raw, ZigZag(anchor.ScanResume - anchor.BlockResume));
            MarkupScanState? state = anchor.State;
            raw.Add((byte)(state is null ? 0 : 1 | (state.BlockHasContent ? 2 : 0) | (state.IsDoubleByteTrailNext ? 4 : 0)));
            if (state is not null)
            {
                IndexBinaryArrays.AppendVarint(raw, (uint)(state.Block - (previous?.FirstBlock ?? 0)));
                IndexBinaryArrays.AppendVarint(raw, (ulong)(anchor.ScanPosition - state.BlockStart));
                IndexBinaryArrays.AppendVarint(raw, (uint)state.BlockStackId);
                IndexBinaryArrays.AppendVarint(raw, (uint)state.BlockTextLength);
                IndexBinaryArrays.AppendVarint(raw, (uint)state.LiveStackId);
                IndexBinaryArrays.AppendVarint(raw, (uint)state.TagLengths.Length);
                foreach (int length in state.TagLengths) IndexBinaryArrays.AppendVarint(raw, (uint)length);
                IndexBinaryArrays.AppendVarint(raw, (uint)state.UnpushedDepth);
                IndexBinaryArrays.AppendVarint(raw, (uint)state.TableDepth);
            }
            previous = anchor;
        }
        WordIndexDeflatedBlob.Write(writer, raw.ToArray());
        Section($"anchors ({Anchors.Length}, scanner snapshot)");

        var weights = new List<byte>(TextLengthGroupTotals.Length * 3);
        foreach (int total in TextLengthGroupTotals) IndexBinaryArrays.AppendVarint(weights, (uint)total);
        WordIndexDeflatedBlob.Write(writer, weights.ToArray());
        Section("text-length groups");
        WriteStrings(writer, Stacks);
        Section($"element stacks ({Stacks.Count})");
        new EliasFanoSequence(Headings.Select(h => h.BlockIndex).ToArray()).Write(writer);
        WordIndexDeflatedBlob.Write(writer, Headings.Select(h => (byte)h.Level).ToArray());
        WriteStrings(writer, Headings.Select(h => h.Text).ToArray());
        Section($"headings ({Headings.Length})");
        var head = new List<string>();
        if (ToratEmet is { } toratEmet)
        {
            head.Add(toratEmet.HeadEndOffset.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (KeyValuePair<string, string> pair in toratEmet.Metadata) { head.Add(pair.Key); head.Add(pair.Value); }
        }
        WriteStrings(writer, head);
        Section($"Torat Emet head ({head.Count / 2} keys)");
        writer.Flush();
    }

    public static MarkupRandomAccessIndex Read(Stream source)
    {
        var reader = new BinaryReader(source, Encoding.UTF8, leaveOpen: true);
        if (reader.ReadString() != Magic) throw new InvalidDataException("Not a text / web archive index.");
        if (reader.ReadInt32() != FormatVersion) throw new InvalidDataException("Text / web archive index: unsupported version.");
        var kind = (MarkupSourceKind)reader.ReadByte();
        long length = reader.ReadInt64(), ticks = reader.ReadInt64();
        int codePage = reader.ReadInt32();
        string reason = reader.ReadString();
        var transfer = (MimeTransferEncoding)reader.ReadByte();
        long textEnd = reader.ReadInt64(), encodedEnd = reader.ReadInt64();
        bool rtl = reader.ReadBoolean();
        long textStart = reader.ReadInt64(), firstFeedEnd = reader.ReadInt64();
        int blockCount = reader.ReadInt32();
        if (blockCount < 0 || textEnd < 0) throw new InvalidDataException("Text / web archive index: negative size.");
        string[] cssStrings = ReadStrings(reader);
        if (cssStrings.Length != 1) throw new InvalidDataException("Text / web archive index: CSS damaged.");
        string css = cssStrings[0];

        byte[] raw = WordIndexDeflatedBlob.Read(reader);
        int at = 0;
        int count = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
        if (count < 0 || (count == 0) != (blockCount == 0) || count > raw.Length) throw new InvalidDataException("Text / web archive index: anchors do not match the blocks.");
        var anchors = new MarkupSpanAnchor[count];
        MarkupSpanAnchor? previous = null;
        for (int i = 0; i < count; i++)
        {
            long position = (previous?.ScanPosition ?? 0) + UnZigZag(IndexBinaryArrays.ReadVarint(raw, ref at));
            long scanResume = (previous?.ScanResume ?? 0) + UnZigZag(IndexBinaryArrays.ReadVarint(raw, ref at));
            long blockResume = scanResume - UnZigZag(IndexBinaryArrays.ReadVarint(raw, ref at));
            if (at >= raw.Length) throw new EndOfStreamException("Text / web archive index: anchors truncated.");
            byte flags = raw[at++];
            MarkupScanState? state = null;
            if ((flags & 1) != 0)
            {
                int block = (previous?.FirstBlock ?? 0) + (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                long blockStart = position - (long)IndexBinaryArrays.ReadVarint(raw, ref at);
                int blockStackId = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                int textLength = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                int liveStackId = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                int tagCount = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                if (tagCount < 0 || tagCount > raw.Length - at) throw new InvalidDataException("Text / web archive index: anchor stack damaged.");
                var tagLengths = new int[tagCount];
                for (int k = 0; k < tagCount; k++) tagLengths[k] = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                int unpushed = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                int tableDepth = (int)IndexBinaryArrays.ReadVarint(raw, ref at);
                state = new MarkupScanState(block, blockStart, blockStackId, textLength, liveStackId, tagLengths, unpushed, tableDepth,
                                            (flags & 2) != 0, (flags & 4) != 0);
            }
            var anchor = new MarkupSpanAnchor(state, position, scanResume, blockResume);
            if ((i == 0) != (state is null) || anchor.FirstBlock >= blockCount || (i > 0 && anchor.FirstBlock <= previous!.FirstBlock)
                || position < 0 || position > textEnd || anchor.FirstBlockStart > position || (i > 0 && position <= previous!.ScanPosition))
                throw new InvalidDataException("Text / web archive index: anchors out of order.");
            anchors[i] = anchor;
            previous = anchor;
        }

        byte[] weightBytes = WordIndexDeflatedBlob.Read(reader);
        var totals = new List<int>();
        for (int position = 0; position < weightBytes.Length;) totals.Add((int)IndexBinaryArrays.ReadVarint(weightBytes, ref position));
        int[] groupTotals = totals.ToArray();
        WordIndexTextLengthGroups.Validate(groupTotals, blockCount);
        var stacks = new List<string>(ReadStrings(reader));
        if (stacks.Count == 0 || stacks[0] != "") throw new InvalidDataException("Text / web archive index: stacks damaged.");
        foreach (MarkupSpanAnchor anchor in anchors)
            if (anchor.State is { } s && ((uint)s.BlockStackId >= (uint)stacks.Count || (uint)s.LiveStackId >= (uint)stacks.Count))
                throw new InvalidDataException("Text / web archive index: anchor stack out of range.");
        int[] headingBlocks = EliasFanoSequence.Read(reader).ToArray();
        byte[] levels = WordIndexDeflatedBlob.Read(reader);
        string[] headingTexts = ReadStrings(reader);
        if (levels.Length != headingBlocks.Length || headingTexts.Length != headingBlocks.Length) throw new InvalidDataException("Text / web archive index: headings do not match.");
        var headings = new WordDocumentHeading[headingBlocks.Length];
        for (int i = 0; i < headings.Length; i++)
        {
            if ((uint)headingBlocks[i] >= (uint)blockCount) throw new InvalidDataException("Text / web archive index: heading past the blocks.");
            headings[i] = new WordDocumentHeading(headingBlocks[i], levels[i], headingTexts[i]);
        }
        string[] head = ReadStrings(reader);
        ToratEmetHead? toratEmet = null;
        if (head.Length > 0)
        {
            if (head.Length % 2 == 0 || !long.TryParse(head[0], System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out long headEnd) || headEnd >= textEnd)
                throw new InvalidDataException("Text / web archive index: Torat Emet head damaged.");
            var pairs = new List<KeyValuePair<string, string>>((head.Length - 1) / 2);
            for (int i = 1; i < head.Length; i += 2) pairs.Add(new(head[i], head[i + 1]));
            toratEmet = new ToratEmetHead(headEnd, pairs);
        }
        return new MarkupRandomAccessIndex(kind, length, ticks, codePage, reason, transfer, textEnd, encodedEnd, rtl, css, textStart, firstFeedEnd,
                                           blockCount, anchors, groupTotals, stacks, headings, toratEmet);
    }

    /// <summary>Strings as varint UTF-8 lengths + bytes, deflated together (any character allowed).</summary>
    private static void WriteStrings(BinaryWriter writer, IReadOnlyList<string> values)
    {
        var raw = new List<byte>();
        IndexBinaryArrays.AppendVarint(raw, (uint)values.Count);
        foreach (string value in values)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(value);
            IndexBinaryArrays.AppendVarint(raw, (uint)bytes.Length);
            raw.AddRange(bytes);
        }
        WordIndexDeflatedBlob.Write(writer, raw.ToArray());
    }

    private static string[] ReadStrings(BinaryReader reader)
    {
        byte[] raw = WordIndexDeflatedBlob.Read(reader);
        int at = 0;
        ulong count = IndexBinaryArrays.ReadVarint(raw, ref at);
        if (count > (ulong)raw.Length) throw new InvalidDataException("Text / web archive index: strings damaged.");
        var values = new string[count];
        for (int i = 0; i < values.Length; i++)
        {
            ulong length = IndexBinaryArrays.ReadVarint(raw, ref at);
            if (length > (ulong)(raw.Length - at)) throw new InvalidDataException("Text / web archive index: strings truncated.");
            values[i] = Encoding.UTF8.GetString(raw, at, (int)length);
            at += (int)length;
        }
        return values;
    }

    /// <summary>Saves the sidecar atomically with its record (<see cref="WordDocumentIndexCache.Save"/>); a failure leaves the old one or none - it is a cache.</summary>
    public void Save(string indexPath, string? documentPath = null) => WordDocumentIndexCache.Save(indexPath, Write, documentPath);

    public static MarkupRandomAccessIndex? TryLoad(string indexPath, string documentPath, MarkupSourceKind kind)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader? reader = IndexBinaryArrays.TryOpenInMemory(indexPath);
            if (reader is null) return null;                  // gone, held or pending delete this instant: a cache miss
            MarkupRandomAccessIndex index = Read(reader.BaseStream);
            return index.Kind == kind && index.MatchesSource(documentPath) ? WordDocumentIndexCache.MarkUsed(indexPath, index) : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException
                                           or OverflowException or ArgumentException or OutOfMemoryException)   // a damaged sidecar is rebuilt
        {
            return null;
        }
    }
}
