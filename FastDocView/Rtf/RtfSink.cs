namespace FastDocView.Rtf;

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
    /// <summary>True: the parser reads \pict and \shp groups into <see cref="Picture"/> (the HTML renderer); false: a placeholder (index, search).</summary>
    public virtual bool WantsPictures => false;
    /// <summary>True: <see cref="RtfPicture.Data"/> is decoded (reading one picture's bytes); false: only its size and type.</summary>
    public virtual bool WantsPictureData => false;
    /// <summary>A {\pict} group, or a {\shp} holding one (then <see cref="RtfPicture.Placement"/> is set).</summary>
    public virtual void Picture(in RtfPicture picture, in RtfGroupState state) { }
    /// <summary>A group closed; <paramref name="depth"/> is the depth after the close (0 = the document group).</summary>
    public virtual void GroupClosed(int depth) { }
    public virtual void UnhandledWord(ReadOnlySpan<byte> name) { }
    /// <summary>A page-setup word (\paperw \margl \margr \gutter, \sectd, \pgwsxn \marglsxn \margrsxn \guttersxn) in the main text.</summary>
    public virtual void PageSetup(RtfWord word, int value) { }
    /// <summary>A destination group was skipped whole (header, info, bookmarks, ...).</summary>
    public virtual void DestinationSkipped(ReadOnlySpan<byte> name) { }
    /// <summary>A \field group opened at <paramref name="depth"/> (its group closes when GroupClosed reports a smaller depth).</summary>
    public virtual void FieldStart(int depth) { }
    /// <summary>The field's {\*\fldinst ...} text, when its group closed.</summary>
    public virtual void FieldInstruction(string instruction) { }
    /// <summary>The field's {\fldrslt ...} group opened: what follows is the shown result.</summary>
    public virtual void FieldResultStart() { }
    /// <summary>{\*\bkmkstart name} closed: a bookmark starts here.</summary>
    public virtual void BookmarkStart(string name) { }
}

/// <summary>A picture: shown size in twips (0 = unknown), the blip type's file extension ("" = a type a browser cannot show),
/// its bytes when the sink asked for them, and for a floating shape its placement.</summary>
internal readonly record struct RtfPicture(int WidthTwips, int HeightTwips, string Extension, byte[]? Data, RtfShapePlacement? Placement = null);

/// <summary>A {\shp}'s placement words and properties, as the parser read them (meanings: RtfParser.ReadShape).</summary>
internal sealed class RtfShapePlacement
{
    public int Left, Top, Right, Bottom;
    /// <summary>\shpwr (0 = not given).</summary>
    public int Wrap;
    /// <summary>\shpwrk (0 both sides, 1 left, 2 right, 3 largest).</summary>
    public int WrapSide;
    /// <summary>posh (0 = not given: placed at \shpleft).</summary>
    public int HorizontalPosition;
    /// <summary>dxWrapDistLeft, dyWrapDistTop, dxWrapDistRight, dyWrapDistBottom in EMU.</summary>
    public long WrapLeft, WrapTop, WrapRight, WrapBottom;

    public void SetProperty(string name, string value)
    {
        if (!long.TryParse(value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long number)) return;
        switch (name)
        {
            case "posh": HorizontalPosition = (int)Math.Max(0, Math.Min(5, number)); break;
            case "dxWrapDistLeft": WrapLeft = number; break;
            case "dyWrapDistTop": WrapTop = number; break;
            case "dxWrapDistRight": WrapRight = number; break;
            case "dyWrapDistBottom": WrapBottom = number; break;
        }
    }
}
