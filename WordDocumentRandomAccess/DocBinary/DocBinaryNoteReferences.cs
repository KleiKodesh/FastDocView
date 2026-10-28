namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Footnotes and endnotes of a .doc ([MS-DOC] §2.8.10 PlcffndRef / §2.8.11 PlcffndTxt and the
/// endnote twins): where each reference sits in the main text (CP), whether it is auto-numbered
/// (FRD.nAuto ≠ 0) and its number, and the CP range of its text in the footnote or endnote story.
/// The tables are small (8 bytes per note), so they are read whole at open.
/// </summary>
public sealed class DocBinaryNoteReferences
{
    public readonly record struct NoteReference(int ReferenceCp, WordDocumentNoteKind Kind, int Number, int TextStartCp, int TextEndCp);

    /// <summary>All references of both kinds, sorted by <see cref="NoteReference.ReferenceCp"/>.</summary>
    public NoteReference[] References { get; }

    public int FootnoteCount { get; }
    public int EndnoteCount { get; }

    public DocBinaryNoteReferences(Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        var references = new List<NoteReference>();
        FootnoteCount = Read(tableStream, fib.FootnoteReferences, fib.FootnoteTexts, fib.FootnoteStoryStart, WordDocumentNoteKind.Footnote, references);
        EndnoteCount = Read(tableStream, fib.EndnoteReferences, fib.EndnoteTexts, fib.EndnoteStoryStart, WordDocumentNoteKind.Endnote, references);
        references.Sort((a, b) => a.ReferenceCp.CompareTo(b.ReferenceCp));
        References = references.ToArray();
    }

    /// <summary>The reference at exactly this main-text CP, if any.</summary>
    public NoteReference? At(int cp)
    {
        int index = LowerBound(cp);
        return index < References.Length && References[index].ReferenceCp == cp ? References[index] : null;
    }

    /// <summary>References with ReferenceCp in [cpStart, cpEnd).</summary>
    public ReadOnlySpan<NoteReference> InRange(int cpStart, int cpEnd)
    {
        int first = LowerBound(cpStart), last = LowerBound(cpEnd);
        return References.AsSpan(first, last - first);
    }

    private int LowerBound(int cp)
    {
        int low = 0, high = References.Length;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (References[middle].ReferenceCp < cp) low = middle + 1; else high = middle;
        }
        return low;
    }

    private static int Read(Stream tableStream, (uint offset, uint length) referencePlc, (uint offset, uint length) textPlc,
                            int storyStart, WordDocumentNoteKind kind, List<NoteReference> into)
    {
        // PlcffndRef: (n + 1) CPs then n FRDs (2 bytes) → n = (lcb - 4) / 6.
        if (referencePlc.length < 10 || textPlc.length < 8) return 0;
        byte[] referenceBytes = ReadBytes(tableStream, referencePlc);
        byte[] textBytes = ReadBytes(tableStream, textPlc);
        int count = (int)((referencePlc.length - 4) / 6);
        int textCpCount = textBytes.Length / 4;
        int autoNumber = 0;
        for (int i = 0; i < count; i++)
        {
            int referenceCp = BitConverter.ToInt32(referenceBytes, i * 4);
            short nAuto = BitConverter.ToInt16(referenceBytes, (count + 1) * 4 + i * 2);
            int number = nAuto != 0 ? ++autoNumber : 0;
            int textStart = i < textCpCount ? BitConverter.ToInt32(textBytes, i * 4) : 0;
            int textEnd = i + 1 < textCpCount ? BitConverter.ToInt32(textBytes, (i + 1) * 4) : textStart;
            into.Add(new NoteReference(referenceCp, kind, number, storyStart + textStart, storyStart + textEnd));
        }
        return count;
    }

    private static byte[] ReadBytes(Stream stream, (uint offset, uint length) location)
    {
        var bytes = new byte[location.length];
        stream.Position = location.offset;
        int read = 0;
        while (read < bytes.Length) { int n = stream.Read(bytes, read, bytes.Length - read); if (n <= 0) break; read += n; }
        return bytes;
    }
}
