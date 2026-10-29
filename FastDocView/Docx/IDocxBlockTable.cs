namespace FastDocView.Docx;

/// <summary>
/// What the block reader and the renderer need per body block: its byte range, kind, style, list
/// label, text length and note references. Served by the finished <see cref="DocxRandomAccessIndex"/>
/// or, while the first open is still scanning, by <see cref="DocxIncrementalIndexBuilder"/> over the
/// blocks known so far.
/// </summary>
internal interface IDocxBlockTable
{
    /// <summary>Blocks whose byte range and metadata are final.</summary>
    int BlockCount { get; }
    string RootOpenTag { get; }
    long BlockStartOffset(int blockIndex);
    long BlockEndOffset(int blockIndex);
    WordDocumentBlockKind BlockKindOf(int blockIndex);
    int BlockTextLengthOf(int blockIndex);
    string? StyleIdOf(int blockIndex);
    DocxListLabel? ListLabelOf(int blockIndex);
    /// <summary>The complex fields open where the block starts (a link that runs across paragraphs), or null.</summary>
    string? OpenFieldsAt(int blockIndex);
    ReadOnlySpan<DocxNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex);
}
