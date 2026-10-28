namespace WordDocumentRandomAccess;

/// <summary>
/// One body-level block of a Word document rendered to HTML. A block is a paragraph, a table
/// (.docx) or a table row (.doc), or another body child; it is the unit the windowed viewer asks
/// for. <see cref="Notes"/> holds the footnotes/endnotes this block references, already rendered,
/// so a notes pane can follow the visible text without a second round trip (Word's draft view).
/// </summary>
public sealed record WordDocumentBlockHtml(int BlockIndex, WordDocumentBlockKind Kind, string Html, IReadOnlyList<WordDocumentNoteHtml> Notes);

public enum WordDocumentBlockKind : byte
{
    Paragraph = 0,
    Table = 1,
    ContentControl = 2,
    Other = 3,
    TableRow = 4,
}

public enum WordDocumentNoteKind : byte
{
    Footnote = 0,
    Endnote = 1,
}

/// <summary>
/// A footnote or endnote body. <see cref="Label"/> is what the reference shows ("3", "iv", or a
/// custom mark); <see cref="Number"/> is its position among the auto-numbered notes of its kind
/// (0 for a custom mark). References in block HTML carry data-note-kind / data-note-number.
/// </summary>
public sealed record WordDocumentNoteHtml(WordDocumentNoteKind Kind, int Number, string Label, string Html);

/// <summary>
/// A Word document opened for random access. Implemented by <see cref="Docx.DocxWindowedDocument"/>
/// (.docx) and <see cref="DocBinary.DocBinaryWindowedDocument"/> (.doc). Both share the same
/// contract so the viewer and, later, the book-view lines adapter need no format branch.
/// </summary>
public interface IWordWindowedDocument : IDisposable
{
    string FilePath { get; }

    /// <summary>
    /// Blocks in the document. While the index is still being built on demand (<see cref="IsIndexComplete"/>
    /// false, .docx first open) this is an estimate that the reader corrects as the scan advances.
    /// </summary>
    int BlockCount { get; }

    /// <summary>False while a .docx first open is still extending its index on demand; .doc and cached/in-memory opens are always complete.</summary>
    bool IsIndexComplete { get; }

    /// <summary>Blocks whose position is known for sure (equals <see cref="BlockCount"/> once the index is complete).</summary>
    int KnownBlockCount { get; }

    /// <summary>True when the document reads right-to-left by default (Hebrew documents).</summary>
    bool IsRightToLeft { get; }

    /// <summary>CSS the rendered block HTML relies on (fonts, paragraph spacing, table borders).</summary>
    string Stylesheet { get; }

    /// <summary>
    /// Approximate text length per block — only for the viewer's scroll-height estimate before a
    /// block has been rendered. Index-time data; no file access.
    /// </summary>
    int EstimatedTextLength(int blockIndex);

    /// <summary>Renders blocks [firstBlockIndex, firstBlockIndex + count) — clamped to the document.</summary>
    IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count);

    /// <summary>Heading blocks for a table of contents: (blockIndex, level 1-9, text). Index-time data; no file access.</summary>
    IReadOnlyList<WordDocumentHeading> Headings { get; }

    /// <summary>How many footnotes and endnotes the main text references (index-time counts).</summary>
    (int footnotes, int endnotes) NoteCounts { get; }

    /// <summary>Every note reference in body order with the block it sits in — for "next footnote" navigation. No file access.</summary>
    IReadOnlyList<WordDocumentNoteReference> NoteReferences { get; }
}

public sealed record WordDocumentNoteReference(WordDocumentNoteKind Kind, int Number, int BlockIndex);

public sealed record WordDocumentHeading(int BlockIndex, int Level, string Text);
