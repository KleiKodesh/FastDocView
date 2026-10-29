namespace FastDocView;

/// <summary>
/// <see cref="IWordWindowedDocument.RenderNotes"/> for readers whose notes are inline in the blocks (.odt, .rtf): renders
/// only the distinct blocks holding the window's notes (runs of adjacent blocks in one call), and picks each note by its
/// ordinal within its block. Sparse notes never pull in the blocks between them.
/// </summary>
internal static class WordDocumentNoteWindow
{
    /// <summary>The clamped window [first, end) over <paramref name="noteCount"/> notes.</summary>
    public static (int first, int end) Clamp(int firstNoteIndex, int count, int noteCount)
    {
        int first = Math.Max(0, Math.Min(firstNoteIndex, noteCount));
        int end = (int)Math.Min(noteCount, (long)first + Math.Max(0, count));
        return (first, end);
    }

    public static IReadOnlyList<WordDocumentNoteHtml> RenderFromBlocks(IReadOnlyList<WordDocumentNoteReference> references, int firstNoteIndex, int count,
                                                                       Func<int, int, IReadOnlyList<WordDocumentBlockHtml>> renderBlocks)
    {
        (int first, int end) = Clamp(firstNoteIndex, count, references.Count);
        if (first >= end) return [];
        // the ordinal of the window's first note inside its block
        int firstBlock = references[first].BlockIndex;
        int ordinal = 0;
        for (int i = first - 1; i >= 0 && references[i].BlockIndex == firstBlock; i--) ordinal++;

        var result = new List<WordDocumentNoteHtml>(end - first);
        int index = first;
        while (index < end)
        {
            int runStart = references[index].BlockIndex, runEnd = runStart;
            int scan = index;
            while (scan < end && references[scan].BlockIndex <= runEnd + 1) { runEnd = references[scan].BlockIndex; scan++; }
            foreach (WordDocumentBlockHtml block in renderBlocks(runStart, runEnd - runStart + 1))
            {
                int skip = block.BlockIndex == firstBlock ? ordinal : 0;
                for (int i = skip; i < block.Notes.Count && index < end && references[index].BlockIndex == block.BlockIndex; i++, index++)
                    result.Add(block.Notes[i]);
                // a block that rendered fewer notes than the index lists: empty stand-ins keep note i at position i
                for (; index < end && references[index].BlockIndex == block.BlockIndex; index++)
                    result.Add(new WordDocumentNoteHtml(references[index].Kind, references[index].Number, references[index].Number > 0 ? references[index].Number.ToString() : "*", ""));
            }
            for (; index < scan; index++)     // blocks the reader did not return (past the end): stand-ins as well
                result.Add(new WordDocumentNoteHtml(references[index].Kind, references[index].Number, references[index].Number > 0 ? references[index].Number.ToString() : "*", ""));
        }
        return result;
    }
}
