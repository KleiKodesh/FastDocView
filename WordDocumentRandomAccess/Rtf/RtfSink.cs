namespace WordDocumentRandomAccess.Rtf;

/// <summary>What a picture-like group was (it is never decoded: a placeholder is shown).</summary>
internal enum RtfEmbeddedKind : byte { Picture, Object, Equation }

/// <summary>
/// Receives the content the <see cref="RtfParser"/> finds. The index pass and the HTML renderer are two sinks over
/// the same parser, so a block is segmented and formatted identically by both.
/// </summary>
internal abstract class RtfSink
{
    /// <summary>False: the parser reports text lengths only and skips decoding (the index pass, outside headings).</summary>
    public virtual bool WantsText => true;

    /// <summary>The audit counts control words the reader does not act on.</summary>
    public virtual bool WantsUnhandledWords => false;

    /// <summary>The document group '{' opened (the parser is just after it).</summary>
    public virtual void DocumentStarted() { }
    public virtual void Text(ReadOnlySpan<char> text, in RtfGroupState state) { }
    public virtual void TextLength(int length, in RtfGroupState state) { }
    /// <summary>\par (or \sect when <paramref name="isSectionEnd"/>).</summary>
    public virtual void ParagraphEnd(in RtfGroupState state, bool isSectionEnd) { }
    public virtual void CellEnd(in RtfGroupState state, bool isNested) { }
    public virtual void RowEnd(in RtfGroupState state, bool isNested) { }
    /// <summary>\page, \line, \column.</summary>
    public virtual void Break(RtfWord word, in RtfGroupState state) { }
    public virtual void FootnoteStart() { }
    public virtual void FootnoteEnd(bool isEndnote) { }
    /// <summary>\chftn — the automatic note number (in the text: the reference; in the note: its mark).</summary>
    public virtual void FootnoteMark(in RtfGroupState state) { }
    public virtual void ListTextStart() { }
    public virtual void ListTextEnd() { }
    public virtual void Embedded(RtfEmbeddedKind kind, in RtfGroupState state) { }
    /// <summary>A group closed; <paramref name="depth"/> is the depth after the close (0 = the document group).</summary>
    public virtual void GroupClosed(int depth) { }
    public virtual void UnhandledWord(ReadOnlySpan<byte> name) { }
    /// <summary>A destination group was skipped whole (header, info, bookmarks, ...).</summary>
    public virtual void DestinationSkipped(ReadOnlySpan<byte> name) { }
}
