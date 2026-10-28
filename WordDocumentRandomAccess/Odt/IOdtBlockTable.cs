namespace WordDocumentRandomAccess.Odt;

/// <summary>
/// What the block reader and the renderer read per block. The finished <see cref="OdtRandomAccessIndex"/>
/// serves it, and so does <see cref="OdtIncrementalIndexBuilder"/> over the blocks it knows so far,
/// so an on-demand open renders through the same code as a complete index.
/// </summary>
internal interface IOdtBlockTable
{
    int BlockCount { get; }
    string RootOpenTag { get; }
    long BlockStartOffset(int blockIndex);
    long BlockEndOffset(int blockIndex);
    WordDocumentBlockKind BlockKindOf(int blockIndex);
    int BlockTextLengthOf(int blockIndex);
    string? ListLabelOf(int blockIndex);
    int ListLevelOf(int blockIndex);
    ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex);
}
