using System.Text;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Collapses every &lt;w:drawing&gt;, &lt;w:pict&gt; and &lt;w:object&gt; subtree in a block's raw XML
/// bytes before parsing, keeping only the text boxes inside it (&lt;w:txbxContent&gt;). A single
/// drawing can carry megabytes of DrawingML that the flow renderer shows as one placeholder, but
/// the text of a text box is document text and must survive. Matching is depth-aware: drawings DO
/// nest (a picture inside a text box inside a shape), and cutting at the first close tag broke the
/// XML of such blocks.
/// </summary>
internal static class DocxDrawingStripper
{
    private static readonly string[] CollapsedElementNames = ["w:drawing", "w:pict", "w:object"];
    private const string TextBoxContentName = "w:txbxContent";

    /// <summary>The bytes of [start, start+length) with drawings collapsed; the input range untouched when there are none.</summary>
    public static ArraySegment<byte> Strip(byte[] bytes, int start, int length)
    {
        var span = new ReadOnlySpan<byte>(bytes, start, length);
        bool hasAny = false;
        foreach (string name in CollapsedElementNames) if (IndexOfOpenTag(span, 0, name) >= 0) { hasAny = true; break; }
        if (!hasAny) return new ArraySegment<byte>(bytes, start, length);

        var output = new MemoryStream(length);
        int position = 0;
        while (position < span.Length)
        {
            (int elementStart, string? name) = NextCollapsedElement(span, position);
            if (name is null) { output.Write(bytes, start + position, span.Length - position); break; }
            output.Write(bytes, start + position, elementStart - position);
            int elementEnd = ElementEnd(span, elementStart, name);
            if (elementEnd < 0) { output.Write(bytes, start + elementStart, span.Length - elementStart); break; }

            // Keep the element (so the renderer still sees "a drawing here") with only its text boxes inside.
            WriteAscii(output, "<" + name + ">");
            CopyTextBoxes(output, bytes, start, span, elementStart, elementEnd);
            WriteAscii(output, "</" + name + ">");
            position = elementEnd;
        }
        return new ArraySegment<byte>(output.GetBuffer(), 0, (int)output.Length);
    }

    private static (int index, string? name) NextCollapsedElement(ReadOnlySpan<byte> span, int from)
    {
        int best = -1;
        string? bestName = null;
        foreach (string name in CollapsedElementNames)
        {
            int found = IndexOfOpenTag(span, from, name);
            if (found >= 0 && (best < 0 || found < best)) { best = found; bestName = name; }
        }
        return (best, bestName);
    }

    /// <summary>Copies every outermost &lt;w:txbxContent&gt;…&lt;/w:txbxContent&gt; inside [from, to).</summary>
    private static void CopyTextBoxes(MemoryStream output, byte[] bytes, int start, ReadOnlySpan<byte> span, int from, int to)
    {
        int position = from;
        while (position < to)
        {
            int textBox = IndexOfOpenTag(span, position, TextBoxContentName);
            if (textBox < 0 || textBox >= to) return;
            int textBoxEnd = ElementEnd(span, textBox, TextBoxContentName);
            if (textBoxEnd < 0 || textBoxEnd > to) return;
            // A text box's own drawings are collapsed the same way, recursively.
            ArraySegment<byte> inner = Strip(bytes, start + textBox, textBoxEnd - textBox);
            output.Write(inner.Array!, inner.Offset, inner.Count);
            position = textBoxEnd;
        }
    }

    /// <summary>End (exclusive) of the element whose open tag starts at <paramref name="openAt"/>, counting nested same-name elements; -1 if unterminated.</summary>
    private static int ElementEnd(ReadOnlySpan<byte> span, int openAt, string name)
    {
        int tagEnd = span.Slice(openAt).IndexOf((byte)'>');
        if (tagEnd < 0) return -1;
        if (span[openAt + tagEnd - 1] == (byte)'/') return openAt + tagEnd + 1;     // self-closing
        byte[] close = Encoding.ASCII.GetBytes("</" + name + ">");
        int depth = 1, position = openAt + tagEnd + 1;
        while (depth > 0)
        {
            int nextOpen = IndexOfOpenTag(span, position, name);
            int nextClose = span.Slice(position).IndexOf(close);
            if (nextClose < 0) return -1;
            nextClose += position;
            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                int openEnd = span.Slice(nextOpen).IndexOf((byte)'>');
                if (openEnd < 0) return -1;
                if (span[nextOpen + openEnd - 1] != (byte)'/') depth++;
                position = nextOpen + openEnd + 1;
                continue;
            }
            depth--;
            position = nextClose + close.Length;
        }
        return position;
    }

    /// <summary>Index of "&lt;name" followed by a delimiter (so &lt;w:pict does not match &lt;w:pictureX), at or after <paramref name="from"/>.</summary>
    private static int IndexOfOpenTag(ReadOnlySpan<byte> span, int from, string name)
    {
        byte[] open = Encoding.ASCII.GetBytes("<" + name);
        int offset = from;
        while (offset < span.Length)
        {
            int found = span.Slice(offset).IndexOf(open);
            if (found < 0) return -1;
            int after = offset + found + open.Length;
            if (after >= span.Length || span[after] is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n') return offset + found;
            offset = after;
        }
        return -1;
    }

    private static void WriteAscii(MemoryStream output, string text)
    {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }
}
