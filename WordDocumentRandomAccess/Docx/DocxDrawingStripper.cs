using System.Text;

namespace WordDocumentRandomAccess.Docx;

/// <summary>
/// Replaces every &lt;w:drawing&gt;…&lt;/w:drawing&gt;, &lt;w:pict&gt;…&lt;/w:pict&gt; and
/// &lt;w:object&gt;…&lt;/w:object&gt; subtree in a block's raw XML bytes with an empty element
/// before parsing. A single drawing can carry megabytes of DrawingML (shapes, charts, embedded
/// SmartArt) that the flow renderer shows as one placeholder anyway; parsing it into a DOM was the
/// cost behind the slow windows of drawing-heavy documents. These elements do not nest in
/// themselves, so a byte search for the matching close tag is exact.
/// </summary>
internal static class DocxDrawingStripper
{
    private static readonly (byte[] open, byte[] close, byte[] empty)[] Elements =
    [
        Element("w:drawing"),
        Element("w:pict"),
        Element("w:object"),
    ];

    private static (byte[] open, byte[] close, byte[] empty) Element(string name)
        => (Encoding.ASCII.GetBytes("<" + name), Encoding.ASCII.GetBytes("</" + name + ">"), Encoding.ASCII.GetBytes("<" + name + "/>"));

    /// <summary>The bytes of [start, start+length) with drawings collapsed; returns the input range untouched when there are none.</summary>
    public static ArraySegment<byte> Strip(byte[] bytes, int start, int length)
    {
        var span = new ReadOnlySpan<byte>(bytes, start, length);
        bool hasAny = false;
        foreach ((byte[] open, _, _) in Elements) if (span.IndexOf(open) >= 0) { hasAny = true; break; }
        if (!hasAny) return new ArraySegment<byte>(bytes, start, length);

        var output = new MemoryStream(length);
        int position = 0;
        while (position < span.Length)
        {
            int nextStart = -1, which = -1;
            for (int i = 0; i < Elements.Length; i++)
            {
                int found = IndexOfElement(span.Slice(position), Elements[i].open);
                if (found >= 0 && (nextStart < 0 || found < nextStart)) { nextStart = found; which = i; }
            }
            if (nextStart < 0) { output.Write(bytes, start + position, span.Length - position); break; }

            output.Write(bytes, start + position, nextStart);
            int elementStart = position + nextStart;
            int tagEnd = span.Slice(elementStart).IndexOf((byte)'>');
            if (tagEnd < 0) { output.Write(bytes, start + elementStart, span.Length - elementStart); break; }
            bool isSelfClosing = span[elementStart + tagEnd - 1] == (byte)'/';
            int elementEnd;
            if (isSelfClosing) elementEnd = elementStart + tagEnd + 1;
            else
            {
                int close = span.Slice(elementStart).IndexOf(Elements[which].close);
                if (close < 0) { output.Write(bytes, start + elementStart, span.Length - elementStart); break; }
                elementEnd = elementStart + close + Elements[which].close.Length;
            }
            output.Write(Elements[which].empty, 0, Elements[which].empty.Length);
            position = elementEnd;
        }
        return new ArraySegment<byte>(output.GetBuffer(), 0, (int)output.Length);
    }

    /// <summary>Index of "&lt;w:name" followed by a delimiter (so &lt;w:pict does not match &lt;w:pictureX).</summary>
    private static int IndexOfElement(ReadOnlySpan<byte> span, byte[] open)
    {
        int offset = 0;
        while (true)
        {
            int found = span.Slice(offset).IndexOf(open);
            if (found < 0) return -1;
            int after = offset + found + open.Length;
            if (after >= span.Length || span[after] is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n') return offset + found;
            offset = after;
        }
    }
}
