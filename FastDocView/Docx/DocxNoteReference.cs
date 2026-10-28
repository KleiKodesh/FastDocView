namespace FastDocView.Docx;

/// <summary>
/// One footnote/endnote reference in the body, found by the index pass: the block it sits in, the
/// note's id in footnotes.xml / endnotes.xml, and its auto-number among notes of its kind
/// (0 when the reference uses a custom mark).
/// </summary>
public readonly record struct DocxNoteReference(int BlockIndex, WordDocumentNoteKind Kind, int NoteId, int Number);
