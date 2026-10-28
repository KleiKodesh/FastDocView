using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// footnotes.xml or endnotes.xml scanned only as far as a reader has asked: <see cref="RangeOf"/>
/// pulls chunks through <see cref="DocxNotesPartScanner"/> until the note's byte range is known (or
/// the part ends). Word writes notes in reference order, so the early blocks' notes sit near the
/// start of the part and each lookup continues where the last stopped — the part is still inflated
/// once in total, never held. <see cref="Complete"/> scans the rest and yields the same
/// <see cref="DocxNotesPartIndex"/> a one-go build produces (same chunks, same checkpoints).
/// </summary>
internal sealed class DocxIncrementalNotesPart
{
    private readonly ZipEntryLocation _entry;
    private readonly DocxEntryInflationCursor _cursor;
    private readonly DocxNotesPartScanner _scanner = new();
    private readonly Dictionary<int, (long start, long end)> _rangeById = new();
    private int _notesRecorded;                      // scanner notes already copied into _rangeById

    public DocxIncrementalNotesPart(Stream package, ZipEntryLocation entry, long checkpointSpanBytes)
    {
        _entry = entry;
        _cursor = new DocxEntryInflationCursor(package, entry, checkpointSpanBytes);
    }

    public ZipEntryLocation Entry => _entry;
    public IReadOnlyList<DeflateCheckpoint> Checkpoints => _cursor.Checkpoints;
    public string RootOpenTag => _scanner.RootOpenTag;
    public long BytesInflated => _cursor.BytesInflated;

    public (long start, long end)? RangeOf(int noteId)
    {
        while (true)
        {
            if (_rangeById.TryGetValue(noteId, out (long, long) range)) return range;
            if (!AdvanceOneChunk()) return null;
        }
    }

    public DocxNotesPartIndex Complete()
    {
        while (AdvanceOneChunk()) { }
        return DocxRandomAccessIndexBuilder.CreateNotesPartIndex(_entry, _scanner, _cursor.Checkpoints);
    }

    private bool AdvanceOneChunk()
    {
        int read = _cursor.ReadNextChunk();
        if (read == 0) return false;
        _scanner.Feed(_cursor.Chunk, read);
        for (; _notesRecorded < _scanner.NoteIds.Count; _notesRecorded++)
            _rangeById[_scanner.NoteIds[_notesRecorded]] = (_scanner.NoteStarts[_notesRecorded], _scanner.NoteEnds[_notesRecorded]);
        return true;
    }
}
