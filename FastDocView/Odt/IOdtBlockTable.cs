namespace FastDocView.Odt;

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
    /// <summary>The label's level geometry ("" when none / not label-alignment): <see cref="OdtListStyle.Level.Layout"/>.</summary>
    string ListLabelLayoutOf(int blockIndex);
    int ListLevelOf(int blockIndex);
    ReadOnlySpan<OdtNoteReference> NoteReferencesInBlocks(int firstBlockIndex, int lastBlockIndex);
    /// <summary>
    /// The body's bytes [start, end) - from the table's own recently scanned bytes when it has them (the compact table reads a
    /// window's bytes while locating its blocks), else from <paramref name="view"/>.
    /// </summary>
    byte[] ReadRange(Stream view, long start, long end);
    /// <summary>A list style kept by the index (body lists, and lists inside text boxes and shapes); null when unknown.</summary>
    OdtListStyle? ListStyleNamed(string name);
}
