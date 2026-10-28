using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Reads one footnote or endnote on demand: looks its byte range up in the part's index, reads just
/// that range through a seekable view of the part (a <see cref="SeekableDeflate.DeflateRandomAccessStream"/>
/// over the package in sidecar mode, the inflated bytes in in-memory mode) and parses it alone.
/// The notes parts are never held whole.
/// </summary>
public sealed class DocxNoteXmlReader : IDisposable
{
    private readonly PartView? _footnotes;
    private readonly PartView? _endnotes;

    /// <summary>
    /// One part: how to find a note's range (the finished part index, or the incremental scan that
    /// advances until the note is found), the view to read it from, and the root tag (known once
    /// the part's first bytes are scanned — resolved on first read).
    /// </summary>
    private sealed class PartView(Func<int, (long start, long end)?> rangeOf, Stream view, Func<string> rootOpenTag)
    {
        public Func<int, (long start, long end)?> RangeOf { get; } = rangeOf;
        public Stream View { get; } = view;
        private byte[]? _rootOpenTagBytes;
        private string? _rootElementName;

        public XElement? Read(int noteId)
        {
            if (RangeOf(noteId) is not { } range) return null;
            if (_rootOpenTagBytes is null)
            {
                string tag = rootOpenTag();
                _rootOpenTagBytes = Encoding.UTF8.GetBytes(tag);
                _rootElementName = DocxXmlFragment.RootElementName(tag);
            }
            byte[] bytes = DocxXmlFragment.ReadRange(View, range.start, range.end);
            return DocxXmlFragment.Parse(_rootOpenTagBytes, _rootElementName!, bytes, 0, bytes.Length, $"Note {noteId}");
        }
    }

    /// <summary>Sidecar mode: views over the package stream (owned by the body reader, not by this one).</summary>
    public static DocxNoteXmlReader OverPackage(Stream package, DocxNotesPartIndex? footnotes, DocxNotesPartIndex? endnotes)
        => new(footnotes is null ? null : OverIndex(DocxEntryInflation.OpenView(package, footnotes.Entry, footnotes.Checkpoints), footnotes),
               endnotes is null ? null : OverIndex(DocxEntryInflation.OpenView(package, endnotes.Entry, endnotes.Checkpoints), endnotes));

    /// <summary>In-memory mode: over the parts inflated during the open pass.</summary>
    public static DocxNoteXmlReader OverBytes(byte[]? footnoteBytes, DocxNotesPartIndex? footnotes, byte[]? endnoteBytes, DocxNotesPartIndex? endnotes)
        => new(footnoteBytes is null || footnotes is null ? null : OverIndex(new MemoryStream(footnoteBytes, writable: false), footnotes),
               endnoteBytes is null || endnotes is null ? null : OverIndex(new MemoryStream(endnoteBytes, writable: false), endnotes));

    /// <summary>Incremental first open: each lookup advances the part's scan only as far as the note (views over the package, not owned).</summary>
    internal static DocxNoteXmlReader OverIncrementalParts(Stream package, DocxIncrementalNotesPart? footnotes, DocxIncrementalNotesPart? endnotes)
        => new(footnotes is null ? null : OverIncrementalPart(package, footnotes),
               endnotes is null ? null : OverIncrementalPart(package, endnotes));

    private static PartView OverIndex(Stream view, DocxNotesPartIndex index)
        => new(index.RangeOf, view, () => index.RootOpenTag);

    private static PartView OverIncrementalPart(Stream package, DocxIncrementalNotesPart part)
        => new(part.RangeOf, DocxEntryInflation.OpenView(package, part.Entry, part.Checkpoints), () => part.RootOpenTag);

    private DocxNoteXmlReader(PartView? footnotes, PartView? endnotes)
    {
        _footnotes = footnotes;
        _endnotes = endnotes;
    }

    public XElement? Read(WordDocumentNoteKind kind, int noteId)
        => (kind == WordDocumentNoteKind.Footnote ? _footnotes : _endnotes)?.Read(noteId);

    /// <summary>Disposes the views only; the package stream belongs to the body reader.</summary>
    public void Dispose()
    {
        _footnotes?.View.Dispose();
        _endnotes?.View.Dispose();
    }
}