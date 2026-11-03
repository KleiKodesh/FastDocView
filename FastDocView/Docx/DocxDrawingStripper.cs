using System.Text;

namespace FastDocView.Docx;

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
            if (name == "w:drawing") WritePictureLayout(output, span, elementStart, elementEnd);
            else WriteVmlPictureLayout(output, span, elementStart, elementEnd);
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

    /// <summary>The element the renderer reads a picture's layout from (no namespace: it is ours, not WordprocessingML).</summary>
    public const string PictureLayoutElementName = "fastDocViewPicture";

    /// <summary>
    /// A DrawingML picture's layout, kept as one small element: wp:inline / wp:anchor (ECMA-376 20.4.2.8 / 20.4.2.3) with
    /// its dist*, behindDoc, wp:extent (20.4.2.7), the wrap type (20.4.2.15-20.4.2.18: wrapSquare / wrapTight / wrapThrough /
    /// wrapTopAndBottom / wrapNone) and its wrapText, wp:positionH (20.4.2.10: relativeFrom, align or posOffset) and the
    /// first a:blip r:embed (20.1.8.13). Only the drawing's first wp:inline / wp:anchor is read; the prefixes are the ones
    /// Word writes (wp:, a:, r:), as w: is for the drawing itself. The picture bytes are not here: the renderer reads them
    /// from the package by the relationship.
    /// </summary>
    private static void WritePictureLayout(MemoryStream output, ReadOnlySpan<byte> span, int from, int to)
    {
        ReadOnlySpan<byte> drawing = span.Slice(from, to - from);
        int frame = IndexOfOpenTag(drawing, 0, "wp:anchor");
        string kind = "anchor";
        int inline = IndexOfOpenTag(drawing, 0, "wp:inline");
        if (frame < 0 || (inline >= 0 && inline < frame)) { frame = inline; kind = "inline"; }
        if (frame < 0) return;
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"" + kind + "\"");
        void Copy(string attribute, ReadOnlySpan<byte> tag, string name)
        {
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(attribute).Append("=\"").Append(value).Append('"');
        }
        ReadOnlySpan<byte> frameTag = OpenTagAt(drawing, frame);
        foreach (string name in new[] { "distT", "distB", "distL", "distR", "behindDoc" }) Copy(name, frameTag, name);
        if (IndexOfOpenTag(drawing, frame, "wp:extent") is var extent and >= 0)
        {
            Copy("cx", OpenTagAt(drawing, extent), "cx");
            Copy("cy", OpenTagAt(drawing, extent), "cy");
        }
        if (kind == "anchor")
        {
            foreach (string wrap in new[] { "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "wrapNone" })
                if (IndexOfOpenTag(drawing, frame, "wp:" + wrap) is var at and >= 0)
                {
                    layout.Append(" wrap=\"").Append(wrap).Append('"');
                    Copy("wrapText", OpenTagAt(drawing, at), "wrapText");
                    break;
                }
            if (IndexOfOpenTag(drawing, frame, "wp:positionH") is var horizontal and >= 0)
            {
                Copy("relativeFrom", OpenTagAt(drawing, horizontal), "relativeFrom");
                int end = drawing.Slice(horizontal).IndexOf("</wp:positionH>"u8);
                ReadOnlySpan<byte> position = end < 0 ? default : drawing.Slice(horizontal, end);
                foreach ((string element, string attribute) in new[] { ("wp:align", "align"), ("wp:posOffset", "posOffset") })
                    if (IndexOfOpenTag(position, 0, element) is var at and >= 0 && position.Slice(at).IndexOf((byte)'>') is var open and >= 0
                        && position.Slice(at + open + 1).IndexOf((byte)'<') is var length and >= 0)
                        layout.Append(' ').Append(attribute).Append("=\"").Append(Encoding.UTF8.GetString(position.Slice(at + open + 1, length).ToArray()).Trim().Replace("\"", "&quot;")).Append('"');
            }
        }
        if (IndexOfOpenTag(drawing, frame, "a:blip") is var blip and >= 0) Copy("embed", OpenTagAt(drawing, blip), "r:embed");
        layout.Append("/>");
        WriteUtf8(output, layout.ToString());
    }

    /// <summary>
    /// An old-style VML picture (w:pict / w:object, ECMA-376 Part 4 VML v:imagedata) kept as the same small element
    /// with kind="vml": the r:id of the first v:imagedata (o:relid when that is what it carries - LibreOffice vmlshapecontext.cxx
    /// imagedata: "shapes in docx use r:id"), the style attribute of the shape that holds it (v:shape / v:rect / v:image:
    /// width, height, position, mso-position-horizontal, margin-left, mso-wrap-distance-*; vmlshapecontext.cxx setStyle) and
    /// w10:wrap type / side (vmlshapecontext.cxx NMSP_vmlWord | XML_wrap). A w:pict without an image (a shape, a text box)
    /// writes nothing - its text boxes are kept as before.
    /// </summary>
    private static void WriteVmlPictureLayout(MemoryStream output, ReadOnlySpan<byte> span, int from, int to)
    {
        ReadOnlySpan<byte> picture = span.Slice(from, to - from);
        int imageData = IndexOfOpenTag(picture, 0, "v:imagedata");
        if (imageData < 0) return;
        ReadOnlySpan<byte> imageTag = OpenTagAt(picture, imageData);
        string? embed = AttributeValue(imageTag, "r:id") ?? AttributeValue(imageTag, "o:relid");
        if (embed is null) return;
        // the shape holding the image: the last v:shape / v:rect / v:image opened before it
        int shape = -1;
        foreach (string name in new[] { "v:shape", "v:rect", "v:image" })
        {
            for (int at = IndexOfOpenTag(picture, 0, name); at >= 0 && at < imageData; at = IndexOfOpenTag(picture, at + 1, name))
                if (at > shape) shape = at;
        }
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"vml\" embed=\"" + embed + "\"");
        if (shape >= 0 && AttributeValue(OpenTagAt(picture, shape), "style") is { } style) layout.Append(" style=\"").Append(style).Append('"');
        if (IndexOfOpenTag(picture, 0, "w10:wrap") is var wrap and >= 0)
        {
            ReadOnlySpan<byte> wrapTag = OpenTagAt(picture, wrap);
            if (AttributeValue(wrapTag, "type") is { } type) layout.Append(" wrap=\"").Append(type).Append('"');
            if (AttributeValue(wrapTag, "side") is { } side) layout.Append(" wrapText=\"").Append(side).Append('"');
        }
        layout.Append("/>");
        WriteUtf8(output, layout.ToString());
    }

    private static void WriteUtf8(MemoryStream output, string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }

    /// <summary>The open tag starting at <paramref name="at"/> (up to its '&gt;').</summary>
    private static ReadOnlySpan<byte> OpenTagAt(ReadOnlySpan<byte> span, int at)
    {
        int end = span.Slice(at).IndexOf((byte)'>');
        return end < 0 ? default : span.Slice(at, end);
    }

    /// <summary>An attribute's raw value (still XML-escaped, so it can be written back as is), null when absent.</summary>
    private static string? AttributeValue(ReadOnlySpan<byte> tag, string name)
    {
        byte[] key = Encoding.ASCII.GetBytes(name + "=");
        int offset = 0;
        while (offset < tag.Length)
        {
            int found = tag.Slice(offset).IndexOf(key);
            if (found < 0) return null;
            int at = offset + found;
            int valueAt = at + key.Length;
            if (at > 0 && tag[at - 1] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' && valueAt < tag.Length && tag[valueAt] is (byte)'"' or (byte)'\'')
            {
                byte quote = tag[valueAt];
                int length = tag.Slice(valueAt + 1).IndexOf(quote);
                if (length < 0) return null;
                string value = Encoding.UTF8.GetString(tag.Slice(valueAt + 1, length).ToArray());
                return quote == (byte)'"' ? value : value.Replace("\"", "&quot;");
            }
            offset = valueAt;
        }
        return null;
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
