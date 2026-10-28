using System.Text;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Paragraph start CPs of the main text, found by one streamed scan for paragraph marks (0x0D),
/// cell/row marks (0x07) and section marks (0x0C). Persisted as a sidecar so re-opening is instant;
/// the scan itself is linear in the text size and needs no decompression.
/// </summary>
public sealed class DocBinaryParagraphIndex
{
    private const string Magic = "KHDOCIDX";
    private const int FormatVersion = 1;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public int[] ParagraphStartCps { get; }
    public int[] ParagraphTextLengths { get; }
    public int MainTextCharacterCount { get; }

    public int ParagraphCount => ParagraphStartCps.Length;

    public int ParagraphEndCp(int paragraphIndex)
        => paragraphIndex + 1 < ParagraphStartCps.Length ? ParagraphStartCps[paragraphIndex + 1] : MainTextCharacterCount;

    private DocBinaryParagraphIndex(long sourceFileLength, long sourceLastWriteTicksUtc, int[] starts, int[] lengths, int mainTextCharacterCount)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        ParagraphStartCps = starts;
        ParagraphTextLengths = lengths;
        MainTextCharacterCount = mainTextCharacterCount;
    }

    public static DocBinaryParagraphIndex Build(DocBinaryPieceTable pieces, int mainTextCharacterCount, long sourceLength, long sourceLastWriteTicksUtc,
                                                IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        var starts = new List<int> { 0 };
        var lengths = new List<int>();
        int currentLength = 0;
        pieces.ScanCharacters(0, mainTextCharacterCount, (characters, count, chunkCp) =>
        {
            cancellation.ThrowIfCancellationRequested();
            for (int i = 0; i < count; i++)
            {
                char c = characters[i];
                currentLength++;
                if (c is '\r' or '\u0007' or '\u000C')
                {
                    int next = chunkCp + i + 1;
                    lengths.Add(currentLength);
                    currentLength = 0;
                    if (next < mainTextCharacterCount) starts.Add(next);
                }
            }
            progress?.Report(new WordDocumentIndexProgress(chunkCp + count, mainTextCharacterCount, starts.Count));
        });
        while (lengths.Count < starts.Count) lengths.Add(currentLength);
        return new DocBinaryParagraphIndex(sourceLength, sourceLastWriteTicksUtc, starts.ToArray(), lengths.ToArray(), mainTextCharacterCount);
    }

    public bool MatchesSource(long fileLength, long lastWriteTicksUtc)
        => fileLength == SourceFileLength && lastWriteTicksUtc == SourceLastWriteTicksUtc;

    public bool MatchesSource(string documentPath)
    {
        var info = new FileInfo(documentPath);
        return info.Exists && MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks);
    }

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var writer = new BinaryWriter(File.Create(temporaryPath), Encoding.UTF8))
        {
            writer.Write(Magic);
            writer.Write(FormatVersion);
            writer.Write(SourceFileLength);
            writer.Write(SourceLastWriteTicksUtc);
            writer.Write(MainTextCharacterCount);
            writer.Write(ParagraphStartCps.Length);
            foreach (int cp in ParagraphStartCps) writer.Write(cp);
            foreach (int length in ParagraphTextLengths) writer.Write(length);
        }
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static DocBinaryParagraphIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using var reader = new BinaryReader(File.OpenRead(indexPath), Encoding.UTF8);
            if (reader.ReadString() != Magic || reader.ReadInt32() != FormatVersion) return null;
            long length = reader.ReadInt64(), ticks = reader.ReadInt64();
            int mainText = reader.ReadInt32();
            int count = reader.ReadInt32();
            var starts = new int[count];
            var lengths = new int[count];
            for (int i = 0; i < count; i++) starts[i] = reader.ReadInt32();
            for (int i = 0; i < count; i++) lengths[i] = reader.ReadInt32();
            var index = new DocBinaryParagraphIndex(length, ticks, starts, lengths, mainText);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or EndOfStreamException)
        {
            return null;
        }
    }
}
