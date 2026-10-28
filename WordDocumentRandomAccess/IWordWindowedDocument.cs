namespace WordDocumentRandomAccess;

/// <summary>
/// One body-level block of a Word document rendered to HTML. A block is a paragraph, a table or
/// another body child; it is the unit the windowed viewer asks for.
/// </summary>
public sealed record WordDocumentBlockHtml(int BlockIndex, WordDocumentBlockKind Kind, string Html);

public enum WordDocumentBlockKind : byte
{
    Paragraph = 0,
    Table = 1,
    ContentControl = 2,
    Other = 3,
}

/// <summary>
/// A Word document opened for random access. Implemented by <see cref="Docx.DocxWindowedDocument"/>
/// (.docx) and <see cref="DocBinary.DocBinaryWindowedDocument"/> (.doc). Both share the same
/// contract so the viewer and, later, the book-view lines adapter need no format branch.
/// </summary>
public interface IWordWindowedDocument : IDisposable
{
    string FilePath { get; }

    int BlockCount { get; }

    /// <summary>True when the document reads right-to-left by default (Hebrew documents).</summary>
    bool IsRightToLeft { get; }

    /// <summary>CSS the rendered block HTML relies on (fonts, paragraph spacing, table borders).</summary>
    string Stylesheet { get; }

    /// <summary>
    /// Approximate text length in bytes/characters per block — only for the viewer's scroll-height
    /// estimate before a block has been rendered. Index-time data; no file access.
    /// </summary>
    int EstimatedTextLength(int blockIndex);

    /// <summary>Renders blocks [firstBlockIndex, firstBlockIndex + count) — clamped to the document.</summary>
    IReadOnlyList<WordDocumentBlockHtml> RenderBlocks(int firstBlockIndex, int count);

    /// <summary>Heading blocks for a table of contents: (blockIndex, level 1-9, text).</summary>
    IReadOnlyList<WordDocumentHeading> Headings { get; }
}

public sealed record WordDocumentHeading(int BlockIndex, int Level, string Text);
