using System.Text;

namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// The blocks of a .doc's main text, found by one streamed scan: a block is a paragraph, or a
/// whole table row (its cells' paragraphs up to the row-end mark) — rendering each cell mark as its
/// own row produced runs of empty, zero-height rows the windowed scroller could not land on. Each
/// paragraph's PAPX is looked up once to tell table paragraphs and row ends apart, and heading
/// paragraphs keep their level and text here so a reopen reads no text for the TOC. Persisted as a
/// sidecar for big documents (policy); the text itself is never copied.
/// </summary>
public sealed class DocBinaryBlockIndex
{
    private const string Magic = "KHDOCIDX";
    private const int FormatVersion = 3;
    private const int MaxHeadingTextLength = 300;

    public long SourceFileLength { get; }
    public long SourceLastWriteTicksUtc { get; }
    public int MainTextCharacterCount { get; }
    public int[] BlockStartCps { get; }
    public WordDocumentBlockKind[] BlockKinds { get; }
    public int[] BlockTextLengths { get; }
    public WordDocumentHeading[] Headings { get; }

    public int BlockCount => BlockStartCps.Length;

    public int BlockEndCp(int blockIndex)
        => blockIndex + 1 < BlockStartCps.Length ? BlockStartCps[blockIndex + 1] : MainTextCharacterCount;

    /// <summary>The block containing a main-text CP.</summary>
    public int BlockAt(int cp)
    {
        int low = 0, high = BlockStartCps.Length - 1;
        while (low < high)
        {
            int middle = (low + high + 1) / 2;
            if (BlockStartCps[middle] <= cp) low = middle; else high = middle - 1;
        }
        return low;
    }

    private DocBinaryBlockIndex(long sourceFileLength, long sourceLastWriteTicksUtc, int mainTextCharacterCount,
                                int[] starts, WordDocumentBlockKind[] kinds, int[] lengths, WordDocumentHeading[] headings)
    {
        SourceFileLength = sourceFileLength;
        SourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        MainTextCharacterCount = mainTextCharacterCount;
        BlockStartCps = starts;
        BlockKinds = kinds;
        BlockTextLengths = lengths;
        Headings = headings;
    }

    public static DocBinaryBlockIndex Build(DocBinaryPieceTable pieces, DocBinaryParagraphHtmlRenderer renderer, DocBinaryStyleSheet styles,
                                            int mainTextCharacterCount, long sourceLength, long sourceLastWriteTicksUtc,
                                            IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        var starts = new List<int>();
        var kinds = new List<WordDocumentBlockKind>();
        var lengths = new List<int>();
        var headings = new List<WordDocumentHeading>();
        int paragraphStart = 0;
        int rowStart = -1;                     // first CP of the table row being collected, -1 outside a row

        void AddBlock(int start, int end, WordDocumentBlockKind kind)
        {
            starts.Add(start);
            kinds.Add(kind);
            lengths.Add(end - start);
        }

        void OnParagraphEnd(int end, char mark)
        {
            int start = paragraphStart;
            paragraphStart = end;
            DocBinaryParagraphHtmlRenderer.TableParagraphInfo table = renderer.ReadTableInfo(start, end);
            if (table.IsInTable || mark == '\u0007')
            {
                if (rowStart < 0) rowStart = start;
                if (table.IsRowEnd) { AddBlock(rowStart, end, WordDocumentBlockKind.TableRow); rowStart = -1; }
                return;
            }
            if (rowStart >= 0) { AddBlock(rowStart, start, WordDocumentBlockKind.TableRow); rowStart = -1; }   // a row without its end mark
            AddBlock(start, end, WordDocumentBlockKind.Paragraph);

            (_, int styleIndex) = renderer.ResolveParagraph(start, end);
            int level = styles.HeadingLevelOf(styleIndex);
            if (level > 0)
            {
                string text = renderer.PlainText(start, Math.Min(end, start + MaxHeadingTextLength));
                if (text.Length > 0) headings.Add(new WordDocumentHeading(starts.Count - 1, level, text));
            }
        }

        pieces.ScanCharacters(0, mainTextCharacterCount, (characters, count, chunkCp) =>
        {
            cancellation.ThrowIfCancellationRequested();
            for (int i = 0; i < count; i++)
                if (characters[i] is '\r' or '\u0007' or '\u000C') OnParagraphEnd(chunkCp + i + 1, characters[i]);
            progress?.Report(new WordDocumentIndexProgress(chunkCp + count, mainTextCharacterCount, starts.Count));
        });
        if (rowStart >= 0) AddBlock(rowStart, paragraphStart, WordDocumentBlockKind.TableRow);
        if (paragraphStart < mainTextCharacterCount) AddBlock(paragraphStart, mainTextCharacterCount, WordDocumentBlockKind.Paragraph);
        if (starts.Count == 0) AddBlock(0, mainTextCharacterCount, WordDocumentBlockKind.Paragraph);

        return new DocBinaryBlockIndex(sourceLength, sourceLastWriteTicksUtc, mainTextCharacterCount,
                                       starts.ToArray(), kinds.ToArray(), lengths.ToArray(), headings.ToArray());
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
            IndexBinaryArrays.Write(writer, BlockStartCps);
            IndexBinaryArrays.Write(writer, BlockKinds);
            IndexBinaryArrays.Write(writer, BlockTextLengths);
            writer.Write(Headings.Length);
            foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
        }
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static DocBinaryBlockIndex? TryLoad(string indexPath, string documentPath)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader reader = IndexBinaryArrays.OpenInMemory(indexPath);
            if (reader.ReadString() != Magic || reader.ReadInt32() != FormatVersion) return null;
            long length = reader.ReadInt64(), ticks = reader.ReadInt64();
            int mainText = reader.ReadInt32();
            int[] starts = IndexBinaryArrays.Read<int>(reader);
            WordDocumentBlockKind[] kinds = IndexBinaryArrays.Read<WordDocumentBlockKind>(reader);
            int[] lengths = IndexBinaryArrays.Read<int>(reader);
            var headings = new WordDocumentHeading[reader.ReadInt32()];
            for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
            var index = new DocBinaryBlockIndex(length, ticks, mainText, starts, kinds, lengths, headings);
            return index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or EndOfStreamException or InvalidDataException)
        {
            return null;
        }
    }
}
