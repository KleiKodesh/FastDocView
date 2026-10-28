using System.Text;

namespace WordDocumentRandomAccess.Markup;

/// <summary>
/// The one streaming pass for .txt and MHTML. .txt: the encoding from the first 128 KB
/// (<see cref="MarkupTextEncoding.DetectPlainText"/>), then the file is fed raw to the scanner in
/// <see cref="MarkupBlockMode.Lines"/>. MHTML: the HTML part is located from the MIME headers, its body
/// is decoded line by line (<see cref="MimeTransferDecoder"/>) and fed to the scanner in
/// <see cref="MarkupBlockMode.Elements"/>; each block start is mapped to its decoder resume point.
/// Reading stops at the end of the HTML part (or at &lt;/body&gt;).
/// </summary>
public static class MarkupRandomAccessIndexBuilder
{
    private const int ChunkSize = 1 << 18;
    private const int RightToLeftSampleBytes = 64 * 1024;

    public static MarkupRandomAccessIndex BuildPlainText(Stream file, long sourceLength, long sourceTicks,
                                                         IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        file.Position = 0;
        var chunk = new byte[Math.Max(ChunkSize, MarkupTextEncoding.DetectionSampleBytes)];
        int sampleLength = ReadFully(file, chunk, MarkupTextEncoding.DetectionSampleBytes);
        (int codePage, int bomLength, string reason) = MarkupTextEncoding.DetectPlainText(chunk.AsSpan(0, sampleLength),
            sampleLength < MarkupTextEncoding.DetectionSampleBytes ? sampleLength : file.Length); // a short read (file shrank) = the file ends there
        Encoding encoding = MarkupTextEncoding.FromCodePage(codePage);
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        int sampleText = Math.Min(sampleLength - bomLength, RightToLeftSampleBytes) & ~(unitSize - 1);
        bool isRightToLeft = MarkupTextEncoding.ContainsHebrew(encoding.GetString(chunk, bomLength, Math.Max(0, sampleText)));

        var scanner = new MarkupBlockScanner(MarkupBlockMode.Lines, unitSize, codePage == MarkupTextEncoding.Utf16BigEndian, encoding);
        scanner.Begin(bomLength);
        scanner.Feed(chunk, bomLength, sampleLength - bomLength, bomLength);
        long done = sampleLength;
        progress?.Report(new WordDocumentIndexProgress(done, file.Length, scanner.BlockStartOffsets.Count));
        while (true)
        {
            cancellation.ThrowIfCancellationRequested();
            int read = file.Read(chunk, 0, chunk.Length);
            if (read <= 0) break;
            scanner.Feed(chunk, 0, read, done);
            done += read;
            progress?.Report(new WordDocumentIndexProgress(done, file.Length, scanner.BlockStartOffsets.Count));
        }
        scanner.Finish(done);
        return new MarkupRandomAccessIndex(MarkupSourceKind.PlainText, sourceLength, sourceTicks, codePage, reason, MimeTransferEncoding.Identity,
                                           done, done, isRightToLeft, "", scanner.BlockStartOffsets.ToArray(), [],
                                           scanner.BlockTextLengths.ToArray(), scanner.BlockStackIds.ToArray(), scanner.Stacks.ToArray(), scanner.Headings.ToArray());
    }

    private static int ReadFully(Stream stream, byte[] buffer, int count)
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

    // ── web archive ──────────────────────────────────────────────────────────────────

    private const int ResumeRingSize = 1 << 14;             // > MarkupBlockScanner.MaximumTagBytes + one fed piece
    private const int FeedPieceBytes = 4096;

    public static MarkupRandomAccessIndex BuildWebArchive(Stream file, long sourceLength, long sourceTicks,
                                                          IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        WebArchiveHtmlPart part = WebArchiveContainer.LocateHtmlPart(file);
        int codePage = (part.Charset is null ? null : MarkupTextEncoding.CodePageOfCharsetName(part.Charset)) ?? MarkupTextEncoding.HebrewWindows;
        string reason = part.Charset is null ? "no charset parameter → windows-1255" : $"charset={part.Charset}";
        Encoding encoding = MarkupTextEncoding.FromCodePage(codePage);
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        var scanner = new MarkupBlockScanner(MarkupBlockMode.Elements, unitSize, codePage == MarkupTextEncoding.Utf16BigEndian, encoding);
        var decoder = new MimeTransferDecoder(part.TransferEncoding);

        var resumeRing = new long[ResumeRingSize];
        var resumes = new List<long>();
        var decoded = new byte[FeedPieceBytes + 4];
        var resumePiece = new long[FeedPieceBytes + 4];
        var bodySample = new List<byte>();
        long decodedTotal = 0;
        long encodedEnd = file.Length;

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
            }
            if (progress is not null && (resumes.Count & 255) == 0)
                progress.Report(new WordDocumentIndexProgress(lineOffset - part.BodyStart, file.Length - part.BodyStart, scanner.BlockStartOffsets.Count));
        }
        if (!scanner.IsBodyStarted) throw new InvalidDataException("Web archive: the HTML part has no <body>.");
        long textEnd = scanner.IsBodyEnded ? scanner.BodyEndTagOffset : decodedTotal;
        scanner.Finish(textEnd);
        while (resumes.Count > scanner.BlockStartOffsets.Count) resumes.RemoveAt(resumes.Count - 1);
        if (resumes.Count < scanner.BlockStartOffsets.Count) throw new InvalidDataException("Web archive: block start past the decoded text.");
        progress?.Report(new WordDocumentIndexProgress(file.Length - part.BodyStart, file.Length - part.BodyStart, scanner.BlockStartOffsets.Count));

        int sampleLength = bodySample.Count & ~(unitSize - 1);
        bool isRightToLeft = MarkupTextEncoding.ContainsHebrew(encoding.GetString(bodySample.ToArray(), 0, sampleLength));
        return new MarkupRandomAccessIndex(MarkupSourceKind.WebArchive, sourceLength, sourceTicks, codePage, reason, part.TransferEncoding,
                                           textEnd, encodedEnd, isRightToLeft, WebArchiveStylesheet.Scope(scanner.StyleText),
                                           scanner.BlockStartOffsets.ToArray(), resumes.ToArray(), scanner.BlockTextLengths.ToArray(),
                                           scanner.BlockStackIds.ToArray(), scanner.Stacks.ToArray(), scanner.Headings.ToArray());
    }

    /// <summary>Block starts whose first decoded byte now exists get its resume point (block starts lie at most one tag back).</summary>
    private static void MapResumePoints(MarkupBlockScanner scanner, List<long> resumes, long[] ring, long decodedTotal)
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
