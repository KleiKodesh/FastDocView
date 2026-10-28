using System.Text;

namespace WordDocumentRandomAccess.Markup;

public enum MarkupSourceKind : byte { PlainText = 0, WebArchive = 1 }

/// <summary>
/// The index of a plain-text (.txt read as HTML) or web-archive (MHTML) document. Text offsets are
/// "text coordinates": file offsets for .txt, decoded offsets of the HTML part for MHTML, where
/// <see cref="BlockResumePoints"/> says where to restart the transfer decoder for each block
/// (<see cref="MimeTransferDecoder.ResumePoint"/>). Per block: text start, text length, the id of the
/// element stack open at its start (<see cref="Stacks"/>). Serialised to/from a <see cref="Stream"/>
/// with its own magic.
/// </summary>
public sealed class MarkupRandomAccessIndex
{
    private const string Magic = "KHMKPIDX";
    private const int FormatVersion = 1;

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
    public long[] BlockStarts { get; }
    public long[] BlockResumePoints { get; }          // empty for .txt
    public int[] BlockTextLengths { get; }
    public int[] BlockStackIds { get; }
    public string[] Stacks { get; }
    public WordDocumentHeading[] Headings { get; }

    public int BlockCount => BlockStarts.Length;

    public MarkupRandomAccessIndex(MarkupSourceKind kind, long sourceFileLength, long sourceLastWriteTicksUtc, int codePage, string encodingReason,
                                   MimeTransferEncoding transferEncoding, long textEnd, long encodedEnd, bool isRightToLeft, string documentCss,
                                   long[] blockStarts, long[] blockResumePoints, int[] blockTextLengths, int[] blockStackIds, string[] stacks,
                                   WordDocumentHeading[] headings)
    {
        Kind = kind; SourceFileLength = sourceFileLength; SourceLastWriteTicksUtc = sourceLastWriteTicksUtc; CodePage = codePage;
        EncodingReason = encodingReason; TransferEncoding = transferEncoding; TextEnd = textEnd; EncodedEnd = encodedEnd;
        IsRightToLeft = isRightToLeft; DocumentCss = documentCss; BlockStarts = blockStarts; BlockResumePoints = blockResumePoints;
        BlockTextLengths = blockTextLengths; BlockStackIds = blockStackIds; Stacks = stacks; Headings = headings;
    }

    public long BlockEnd(int blockIndex) => blockIndex + 1 < BlockStarts.Length ? BlockStarts[blockIndex + 1] : TextEnd;

    public bool MatchesSource(long fileLength, long lastWriteTicksUtc)
        => fileLength == SourceFileLength && lastWriteTicksUtc == SourceLastWriteTicksUtc;

    public bool MatchesSource(string documentPath)
    {
        var info = new FileInfo(documentPath);
        return info.Exists && MatchesSource(info.Length, info.LastWriteTimeUtc.Ticks);
    }

    public void Write(Stream destination)
    {
        var writer = new BinaryWriter(destination, Encoding.UTF8, leaveOpen: true);
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
        writer.Write(DocumentCss);
        IndexBinaryArrays.Write(writer, BlockStarts);
        IndexBinaryArrays.Write(writer, BlockResumePoints);
        IndexBinaryArrays.Write(writer, BlockTextLengths);
        IndexBinaryArrays.Write(writer, BlockStackIds);
        writer.Write(Stacks.Length);
        foreach (string stack in Stacks) writer.Write(stack);
        writer.Write(Headings.Length);
        foreach (WordDocumentHeading heading in Headings) { writer.Write(heading.BlockIndex); writer.Write((byte)heading.Level); writer.Write(heading.Text); }
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
        string css = reader.ReadString();
        long[] starts = IndexBinaryArrays.Read<long>(reader);
        long[] resumes = IndexBinaryArrays.Read<long>(reader);
        int[] lengths = IndexBinaryArrays.Read<int>(reader);
        int[] stackIds = IndexBinaryArrays.Read<int>(reader);
        var stacks = new string[reader.ReadInt32()];
        for (int i = 0; i < stacks.Length; i++) stacks[i] = reader.ReadString();
        var headings = new WordDocumentHeading[reader.ReadInt32()];
        for (int i = 0; i < headings.Length; i++) headings[i] = new WordDocumentHeading(reader.ReadInt32(), reader.ReadByte(), reader.ReadString());
        return new MarkupRandomAccessIndex(kind, length, ticks, codePage, reason, transfer, textEnd, encodedEnd, rtl, css, starts, resumes, lengths, stackIds, stacks, headings);
    }

    public void Save(string indexPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(indexPath)!);
        string temporaryPath = indexPath + ".tmp";
        using (var file = File.Create(temporaryPath)) Write(file);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        File.Move(temporaryPath, indexPath);
    }

    public static MarkupRandomAccessIndex? TryLoad(string indexPath, string documentPath, MarkupSourceKind kind)
    {
        if (!File.Exists(indexPath)) return null;
        try
        {
            using BinaryReader reader = IndexBinaryArrays.OpenInMemory(indexPath);
            MarkupRandomAccessIndex index = Read(reader.BaseStream);
            return index.Kind == kind && index.MatchesSource(documentPath) ? index : null;
        }
        catch (Exception exception) when (exception is IOException or InvalidDataException or EndOfStreamException
                                           or OverflowException or ArgumentException or OutOfMemoryException)   // a damaged sidecar is rebuilt
        {
            return null;
        }
    }
}
