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
    // -- plain text -----------------------------------------------------------------------

    internal static (int codePage, int unitSize, Encoding encoding) PlainTextEncoding(int codePage)
    {
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        return (codePage, unitSize, MarkupTextEncoding.FromCodePage(codePage));
    }

    public static MarkupRandomAccessIndex BuildPlainText(Stream file, long sourceLength, long sourceTicks,
                                                         IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
        => new MarkupIncrementalIndexBuilder(MarkupSourceKind.PlainText, file, sourceLength, sourceTicks, progress, cancellation).Complete();

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

    internal static (int codePage, string reason, int unitSize, Encoding encoding) WebArchiveEncoding(string? charset)
    {
        int codePage = (charset is null ? null : MarkupTextEncoding.CodePageOfCharsetName(charset)) ?? MarkupTextEncoding.HebrewWindows;
        string reason = charset is null ? "no charset parameter → windows-1255" : $"charset={charset}";
        int unitSize = codePage is MarkupTextEncoding.Utf16LittleEndian or MarkupTextEncoding.Utf16BigEndian ? 2 : 1;
        return (codePage, reason, unitSize, MarkupTextEncoding.FromCodePage(codePage));
    }

    public static MarkupRandomAccessIndex BuildWebArchive(Stream file, long sourceLength, long sourceTicks,
                                                          IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
        => new MarkupIncrementalIndexBuilder(MarkupSourceKind.WebArchive, file, sourceLength, sourceTicks, progress, cancellation).Complete();

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
