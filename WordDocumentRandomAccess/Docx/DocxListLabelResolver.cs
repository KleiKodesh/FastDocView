namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// The index pass's hook for list numbering: given a body-level paragraph's style id and its direct
/// numPr (numId -1 when absent), returns the next label of that list ("3.", "ב.") or null when the
/// paragraph is not a list item. See <see cref="DocxDocumentParts.CreateListLabelResolver"/>.
/// </summary>
public delegate string? DocxListLabelResolver(string? styleId, int numId, int level);
