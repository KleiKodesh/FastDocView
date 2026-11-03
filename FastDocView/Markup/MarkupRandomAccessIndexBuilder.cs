using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// The one streaming pass for .txt and MHTML. .txt: the encoding from
/// the first 128 KB (<see cref="MarkupTextEncoding.DetectPlainText"/>), then the file is fed raw to the scanner in
/// <see cref="MarkupBlockMode.Lines"/>. MHTML: the HTML part is located from the MIME headers, its body is decoded
/// (<see cref="MimeTransferDecoder"/>) and fed to the scanner in <see cref="MarkupBlockMode.Elements"/>; block starts
/// are mapped to their decoder resume points. Reading stops at the end of the HTML part (or at &lt;/body&gt;).
/// The pass keeps no per-block table: finished blocks are handed off into the text-length groups, and at the first
/// span point past every <see cref="MarkupRandomAccessIndex.AnchorSpanBytes"/> (MHTML: <see cref="MarkupRandomAccessIndex.WebArchiveAnchorSpanBytes"/>) of text a scanner snapshot is kept
/// (<see cref="MarkupSpanAnchor"/>). A rescan (<see cref="MarkupSpanRescan"/>) feeds the scanner exactly the pieces the pass fed (.txt: the same
/// piece boundaries - a long line's split point depends on them), so it finds the same blocks.
/// </summary>
public static class MarkupRandomAccessIndexBuilder
{
    public const int ChunkSize = 1 << 16;                 // = the anchor spacing: a .txt anchor sits at a piece boundary
    private const int RightToLeftSampleBytes = 64 * 1024;
    private const int HandOffBlocks = 4096;                 // finished blocks kept before they are folded into the groups

    /// <summary>Folds finished blocks into the text-length groups and drops them from the scanner.</summary>
    private sealed class BlockHandOff
    {
        public readonly List<int> GroupTotals = [];

        public void HandOff(MarkupBlockScanner scanner, int count, List<long>? resumes = null)
        {
            if (count <= 0) return;
            for (int i = 0; i < count; i++)
            {
                int group = (scanner.BlockIndexBase + i) / WordIndexTextLengthGroups.GroupSize;
                while (GroupTotals.Count <= group) GroupTotals.Add(0);
                GroupTotals[group] += scanner.BlockTextLengths[i];
            }
            scanner.RemoveFirstBlocks(count);
            resumes?.RemoveRange(0, count);
        }

        public void HandOffAll(MarkupBlockScanner scanner)
        {
            for (int i = 0; i < scanner.BlockTextLengths.Count; i++)
            {
                int group = (scanner.BlockIndexBase + i) / WordIndexTextLengthGroups.GroupSize;
                while (GroupTotals.Count <= group) GroupTotals.Add(0);
                GroupTotals[group] += scanner.BlockTextLengths[i];
            }
        }
    }

    private static int CurrentBlock(MarkupBlockScanner scanner) => scanner.BlockIndexBase + scanner.BlockStartOffsets.Count - 1;

    private static MarkupSpanAnchor[] FinishAnchors(List<MarkupSpanAnchor> anchors, int blockCount)
    {
        if (blockCount == 0) return [];
        while (anchors.Count > 1 && anchors[anchors.Count - 1].FirstBlock >= blockCount) anchors.RemoveAt(anchors.Count - 1);
        return anchors.ToArray();
    }

    private static int[] FinishGroups(BlockHandOff handOff, int blockCount)
    {
        int groups = (blockCount + WordIndexTextLengthGroups.GroupSize - 1) / WordIndexTextLengthGroups.GroupSize;
        while (handOff.GroupTotals.Count < groups) handOff.GroupTotals.Add(0);
        return handOff.GroupTotals.GetRange(0, groups).ToArray();
    }

    // -- plain text -----------------------------------------------------------------------

    internal static (int codePage, int unitSize, Encoding encoding) PlainTextEncoding(int codePage)
    {
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        return (codePage, unitSize, MarkupTextEncoding.FromCodePage(codePage));
    }

    public static MarkupRandomAccessIndex BuildPlainText(Stream file, long sourceLength, long sourceTicks,
                                                         IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        file.Position = 0;
        var chunk = new byte[Math.Max(ChunkSize, MarkupTextEncoding.DetectionSampleBytes)];
        int sampleLength = ReadFully(file, chunk, MarkupTextEncoding.DetectionSampleBytes);
        (int codePage, int bomLength, string reason) = MarkupTextEncoding.DetectPlainText(chunk.AsSpan(0, sampleLength),
            sampleLength < MarkupTextEncoding.DetectionSampleBytes ? sampleLength : file.Length); // a short read (file shrank) = the file ends there
        (_, int unitSize, Encoding encoding) = PlainTextEncoding(codePage);
        int sampleText = Math.Min(sampleLength - bomLength, RightToLeftSampleBytes) & ~(unitSize - 1);
        bool isRightToLeft = MarkupTextEncoding.ContainsHebrew(encoding.GetString(chunk, bomLength, Math.Max(0, sampleText)));

        var scanner = new MarkupBlockScanner(MarkupBlockMode.Lines, unitSize, codePage == MarkupTextEncoding.Utf16BigEndian, encoding);
        var handOff = new BlockHandOff();
        var anchors = new List<MarkupSpanAnchor> { new(null, bomLength, 0, 0) };
        scanner.Begin(bomLength);
        scanner.Feed(chunk, bomLength, sampleLength - bomLength, bomLength);
        long done = sampleLength;
        void AfterPiece()
        {
            MarkupSpanAnchor last = anchors[anchors.Count - 1];
            if (done - last.ScanPosition >= MarkupRandomAccessIndex.AnchorSpanBytes && scanner.IsAtSpanPoint && CurrentBlock(scanner) > last.FirstBlock)
                anchors.Add(new MarkupSpanAnchor(scanner.ExportSpanState(), done, 0, 0));
            if (scanner.BlockStartOffsets.Count > HandOffBlocks) handOff.HandOff(scanner, scanner.BlockStartOffsets.Count - 1);
            progress?.Report(new WordDocumentIndexProgress(done, file.Length, CurrentBlock(scanner) + 1));
        }
        AfterPiece();
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            int read = ReadFully(file, chunk, ChunkSize);            // full pieces: a rescan feeds the same boundaries
            if (read <= 0) break;
            scanner.Feed(chunk, 0, read, done);
            done += read;
            AfterPiece();
        }
        scanner.Finish(done);
        int blockCount = scanner.BlockIndexBase + scanner.BlockStartOffsets.Count;
        handOff.HandOffAll(scanner);
        return new MarkupRandomAccessIndex(MarkupSourceKind.PlainText, sourceLength, sourceTicks, codePage, reason, MimeTransferEncoding.Identity,
                                           done, done, isRightToLeft, "", bomLength, sampleLength, blockCount, FinishAnchors(anchors, blockCount),
                                           FinishGroups(handOff, blockCount), scanner.Stacks, scanner.Headings.ToArray());
    }

    internal static int ReadFully(Stream stream, byte[] buffer, int count)
    {
        int total = 0;
        while (total < count)
        {
            int read = stream.Read(buffer, total, count - total);
            if (read <= 0) break;
            total += read;
        }
        return total;
    }

    // -- web archive ----------------------------------------------------------------------

    internal const int ResumeRingSize = 1 << 14;             // > MarkupBlockScanner.MaximumTagBytes + one fed piece
    internal const int FeedPieceBytes = 4096;

    private static (int codePage, string reason, int unitSize, Encoding encoding) WebArchiveEncoding(string? charset)
    {
        int codePage = (charset is null ? null : MarkupTextEncoding.CodePageOfCharsetName(charset)) ?? MarkupTextEncoding.HebrewWindows;
        string reason = charset is null ? "no charset parameter → windows-1255" : $"charset={charset}";
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        return (codePage, reason, unitSize, MarkupTextEncoding.FromCodePage(codePage));
    }

    public static MarkupRandomAccessIndex BuildWebArchive(Stream file, long sourceLength, long sourceTicks,
                                                          IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        WebArchiveHtmlPart part = WebArchiveContainer.LocateHtmlPart(file);
        (int codePage, string reason, int unitSize, Encoding encoding) = WebArchiveEncoding(part.Charset);
        var scanner = new MarkupBlockScanner(MarkupBlockMode.Elements, unitSize, codePage == MarkupTextEncoding.Utf16BigEndian, encoding);
        var decoder = new MimeTransferDecoder(part.TransferEncoding);
        var handOff = new BlockHandOff();
        long startResume = MimeTransferDecoder.ResumePoint(part.BodyStart, 0);
        var anchors = new List<MarkupSpanAnchor> { new(null, 0, startResume, startResume) };
        (MarkupScanState state, long position)? pending = null;

        var resumeRing = new long[ResumeRingSize];
        var resumes = new List<long>();                     // aligned with the scanner's (not yet handed off) blocks
        var decoded = new byte[FeedPieceBytes + 4];
        var resumePiece = new long[FeedPieceBytes + 4];
        var bodySample = new List<byte>();
        long decodedTotal = 0;
        long encodedEnd = file.Length;
        int pieces = 0;

        void FinishPending()
        {
            if (pending is not { } p || decodedTotal <= p.position) return;
            int local = p.state.Block - scanner.BlockIndexBase;
            anchors.Add(new MarkupSpanAnchor(p.state, p.position, resumeRing[p.position & (ResumeRingSize - 1)], resumes[local]));
            handOff.HandOff(scanner, local, resumes);
            pending = null;
        }

        file.Position = part.BodyStart;
        var lines = new MarkupLineReader(file);
        while (!scanner.IsBodyEnded && lines.ReadLine(out ReadOnlySpan<byte> line, out long lineOffset))
        {
            cancellation.ThrowIfCancellationRequested();
            if (part.Boundary is not null && WebArchiveContainer.IsBoundaryLine(line, part.Boundary, out _)) { encodedEnd = lineOffset; break; }
            byte[] lineBytes = line.ToArray();
            for (int at = 0; at < lineBytes.Length; at += FeedPieceBytes)
            {
                int count = Math.Min(FeedPieceBytes, lineBytes.Length - at), produced = 0;
                decoder.Feed(lineBytes, at, count, lineOffset + at, decoded, resumePiece, ref produced);
                if (produced == 0) continue;
                for (int k = 0; k < produced; k++) resumeRing[(decodedTotal + k) & (ResumeRingSize - 1)] = resumePiece[k];
                if (scanner.IsBodyStarted && bodySample.Count < RightToLeftSampleBytes)
                    for (int k = 0; k < produced && bodySample.Count < RightToLeftSampleBytes; k++) bodySample.Add(decoded[k]);
                scanner.Feed(decoded, 0, produced, decodedTotal);
                decodedTotal += produced;
                MapResumePoints(scanner, resumes, resumeRing, decodedTotal);
                FinishPending();
                MarkupSpanAnchor last = anchors[anchors.Count - 1];
                if (pending is null && decodedTotal - last.ScanPosition >= MarkupRandomAccessIndex.WebArchiveAnchorSpanBytes && scanner.IsAtSpanPoint
                    && CurrentBlock(scanner) > last.FirstBlock)
                    pending = (scanner.ExportSpanState(), decodedTotal);
                if (pending is null && resumes.Count > HandOffBlocks) handOff.HandOff(scanner, Math.Min(resumes.Count, scanner.BlockStartOffsets.Count - 1), resumes);
            }
            if (progress is not null && (++pieces & 255) == 0)
                progress.Report(new WordDocumentIndexProgress(lineOffset - part.BodyStart, file.Length - part.BodyStart, CurrentBlock(scanner) + 1));
        }
        if (!scanner.IsBodyStarted) throw new InvalidDataException("Web archive: the HTML part has no <body>.");
        long textEnd = scanner.IsBodyEnded ? scanner.BodyEndTagOffset : decodedTotal;
        scanner.Finish(textEnd);
        while (resumes.Count > scanner.BlockStartOffsets.Count) resumes.RemoveAt(resumes.Count - 1);
        if (resumes.Count < scanner.BlockStartOffsets.Count) throw new InvalidDataException("Web archive: block start past the decoded text.");
        int blockCount = scanner.BlockIndexBase + scanner.BlockStartOffsets.Count;
        if (pending is { } open && open.state.Block < blockCount && decodedTotal > open.position)
            anchors.Add(new MarkupSpanAnchor(open.state, open.position, resumeRing[open.position & (ResumeRingSize - 1)], resumes[open.state.Block - scanner.BlockIndexBase]));
        handOff.HandOffAll(scanner);
        progress?.Report(new WordDocumentIndexProgress(file.Length - part.BodyStart, file.Length - part.BodyStart, blockCount));

        int sampleLength = bodySample.Count & ~(unitSize - 1);
        bool isRightToLeft = MarkupTextEncoding.ContainsHebrew(encoding.GetString(bodySample.ToArray(), 0, sampleLength));
        return new MarkupRandomAccessIndex(MarkupSourceKind.WebArchive, sourceLength, sourceTicks, codePage, reason, part.TransferEncoding,
                                           textEnd, encodedEnd, isRightToLeft, WebArchiveStylesheet.Scope(scanner.StyleText), startResume, 0,
                                           blockCount, FinishAnchors(anchors, blockCount), FinishGroups(handOff, blockCount), scanner.Stacks,
                                           scanner.Headings.ToArray());
    }

    /// <summary>Block starts whose first decoded byte now exists get its resume point (block starts lie at most one tag back).</summary>
    internal static void MapResumePoints(MarkupBlockScanner scanner, List<long> resumes, long[] ring, long decodedTotal)
    {
        while (resumes.Count < scanner.BlockStartOffsets.Count)
        {
            long start = scanner.BlockStartOffsets[resumes.Count];
            if (start >= decodedTotal) return;
            if (decodedTotal - start > ResumeRingSize) throw new InvalidDataException("Web archive: block start fell out of the resume window.");
            resumes.Add(ring[start & (ResumeRingSize - 1)]);
        }
    }
}
