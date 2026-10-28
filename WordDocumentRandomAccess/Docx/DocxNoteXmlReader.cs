using System.Text;
using System.Xml.Linq;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Reads one footnote or endnote on demand: looks its byte range up in the part's index, reads just
/// that range through a seekable view of the part (a <see cref="Deflate.DeflateRandomAccessStream"/>
/// over the package in sidecar mode, the inflated bytes in in-memory mode) and parses it alone.
/// The notes parts are never held whole.
/// </summary>
public sealed class DocxNoteXmlReader : IDisposable
{
    private readonly PartView? _footnotes;
    private readonly PartView? _endnotes;

    private sealed record PartView(DocxNotesPartIndex Index, Stream View, byte[] RootOpenTagBytes, string RootElementName);

    /// <summary>Sidecar mode: views over the package stream (owned by the body reader, not by this one).</summary>
    public static DocxNoteXmlReader OverPackage(Stream package, DocxNotesPartIndex? footnotes, DocxNotesPartIndex? endnotes)
        => new(footnotes is null ? null : DocxEntryInflation.OpenView(package, footnotes.Entry, footnotes.Checkpoints),
               endnotes is null ? null : DocxEntryInflation.OpenView(package, endnotes.Entry, endnotes.Checkpoints),
               footnotes, endnotes);

    /// <summary>In-memory mode: over the parts inflated during the open pass.</summary>
    public static DocxNoteXmlReader OverBytes(byte[]? footnoteBytes, DocxNotesPartIndex? footnotes, byte[]? endnoteBytes, DocxNotesPartIndex? endnotes)
        => new(footnoteBytes is null ? null : new MemoryStream(footnoteBytes, writable: false),
               endnoteBytes is null ? null : new MemoryStream(endnoteBytes, writable: false),
               footnotes, endnotes);

    private DocxNoteXmlReader(Stream? footnoteView, Stream? endnoteView, DocxNotesPartIndex? footnotes, DocxNotesPartIndex? endnotes)
    {
        _footnotes = footnotes is null || footnoteView is null ? null : CreateView(footnotes, footnoteView);
        _endnotes = endnotes is null || endnoteView is null ? null : CreateView(endnotes, endnoteView);
    }

    private static PartView CreateView(DocxNotesPartIndex index, Stream view)
        => new(index, view, Encoding.UTF8.GetBytes(index.RootOpenTag), DocxXmlFragment.RootElementName(index.RootOpenTag));

    public XElement? Read(WordDocumentNoteKind kind, int noteId)
    {
        PartView? part = kind == WordDocumentNoteKind.Footnote ? _footnotes : _endnotes;
        if (part?.Index.RangeOf(noteId) is not { } range) return null;
        byte[] bytes = DocxXmlFragment.ReadRange(part.View, range.start, range.end);
        return DocxXmlFragment.Parse(part.RootOpenTagBytes, part.RootElementName, bytes, 0, bytes.Length);
    }

    /// <summary>Disposes the views only; the package stream belongs to the body reader.</summary>
    public void Dispose()
    {
        _footnotes?.View.Dispose();
        _endnotes?.View.Dispose();
    }
}
