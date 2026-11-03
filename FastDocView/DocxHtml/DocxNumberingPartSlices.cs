using System.Globalization;
using System.Text;

namespace FastDocView.DocxHtml;

/// <summary>
/// numbering.xml cut into its top-level elements without parsing it: the byte range of every w:num and
/// w:abstractNum (with the id its opening tag carries) and the root's opening tag, so a definition is parsed
/// only when a list of the window asks for it (<see cref="DocxNumberingDefinitions"/>) instead of the whole
/// part on the first window — a big document's numbering.xml is a megabyte of levels for hundreds of lists,
/// of which a window uses a few. A part whose shape the walk cannot follow (another prefix for the main
/// namespace, not UTF-8, unbalanced tags) gives null: the caller parses it whole, as before.
/// </summary>
internal sealed class DocxNumberingPartSlices
{
    public readonly record struct Slice(int Start, int Length);

    private static readonly byte[] MainNamespace = Encoding.ASCII.GetBytes("xmlns:w=\"" + DocxStyleCascade.W.NamespaceName + "\"");

    public byte[] Bytes { get; }
    public byte[] RootOpenTag { get; }
    /// <summary>Every w:num / w:abstractNum with that id, in document order (a duplicated id is read in order, as the whole-part parse layers it).</summary>
    public Dictionary<int, List<Slice>> NumsById { get; } = new();
    public Dictionary<int, List<Slice>> AbstractNumsById { get; } = new();

    private DocxNumberingPartSlices(byte[] bytes, byte[] rootOpenTag)
    {
        Bytes = bytes;
        RootOpenTag = rootOpenTag;
    }

    public static DocxNumberingPartSlices? TryCut(byte[] bytes)
    {
        var span = new ReadOnlySpan<byte>(bytes);
        int position = 0;
        // prolog: declaration, comments, processing instructions, whitespace — up to the root element
        while (true)
        {
            int open = span.Slice(position).IndexOf((byte)'<');
            if (open < 0) return null;
            position += open;
            if (!TrySkipMarkup(span, ref position, out bool isElement)) return null;
            if (isElement) break;
        }
        int rootStart = position;
        if (!TryReadTag(span, ref position, out ReadOnlySpan<byte> rootName, out bool rootSelfClosing)) return null;
        if (!rootName.SequenceEqual("w:numbering"u8) || rootSelfClosing) return null;
        ReadOnlySpan<byte> rootTag = span.Slice(rootStart, position - rootStart);
        if (rootTag.IndexOf(MainNamespace) < 0) return null;         // the w: prefix must be the main namespace for the byte-level names to mean what the parse sees
        var slices = new DocxNumberingPartSlices(bytes, rootTag.ToArray());

        int depth = 1;
        while (depth > 0)
        {
            int open = span.Slice(position).IndexOf((byte)'<');
            if (open < 0) return null;
            position += open;
            if (position + 1 < span.Length && span[position + 1] == (byte)'/')
            {
                int close = span.Slice(position).IndexOf((byte)'>');
                if (close < 0) return null;
                position += close + 1;
                depth--;
                continue;
            }
            if (!TrySkipMarkup(span, ref position, out bool isElement)) return null;
            if (!isElement) continue;
            int elementStart = position;
            if (!TryReadTag(span, ref position, out ReadOnlySpan<byte> name, out bool selfClosing)) return null;
            if (depth == 1 && (name.SequenceEqual("w:num"u8) || name.SequenceEqual("w:abstractNum"u8)))
            {
                bool isAbstract = name.Length > 5;
                int? id = AttributeValue(span.Slice(elementStart, position - elementStart), isAbstract ? "w:abstractNumId="u8 : "w:numId="u8);
                if (!selfClosing && !TrySkipElement(span, ref position)) return null;
                if (id is { } known)
                {
                    Dictionary<int, List<Slice>> table = isAbstract ? slices.AbstractNumsById : slices.NumsById;
                    if (!table.TryGetValue(known, out List<Slice>? list)) table[known] = list = new List<Slice>();
                    list.Add(new Slice(elementStart, position - elementStart));
                }
                continue;
            }
            if (!selfClosing) depth++;
        }
        return slices;
    }

    /// <summary>At '&lt;': skips a comment, CDATA section or processing instruction; an element start is left for <see cref="TryReadTag"/>.</summary>
    private static bool TrySkipMarkup(ReadOnlySpan<byte> span, ref int position, out bool isElement)
    {
        isElement = false;
        if (position + 1 >= span.Length) return false;
        byte next = span[position + 1];
        if (next == (byte)'?') return SkipPast(span, ref position, "?>"u8);
        if (next != (byte)'!') { isElement = true; return true; }
        if (span.Slice(position).StartsWith("<!--"u8)) return SkipPast(span, ref position, "-->"u8);
        if (span.Slice(position).StartsWith("<![CDATA["u8)) return SkipPast(span, ref position, "]]>"u8);
        return false;                                                   // a DOCTYPE: not worth following here
    }

    private static bool SkipPast(ReadOnlySpan<byte> span, ref int position, ReadOnlySpan<byte> marker)
    {
        int at = span.Slice(position).IndexOf(marker);
        if (at < 0) return false;
        position += at + marker.Length;
        return true;
    }

    /// <summary>Reads one start tag at '&lt;' (quotes honoured: a '&gt;' inside an attribute value does not end it); leaves the position after '&gt;'.</summary>
    private static bool TryReadTag(ReadOnlySpan<byte> span, ref int position, out ReadOnlySpan<byte> name, out bool selfClosing)
    {
        name = default;
        selfClosing = false;
        int nameStart = position + 1, nameEnd = nameStart;
        while (nameEnd < span.Length && span[nameEnd] > (byte)' ' && span[nameEnd] != (byte)'>' && span[nameEnd] != (byte)'/') nameEnd++;
        if (nameEnd == nameStart) return false;
        name = span.Slice(nameStart, nameEnd - nameStart);
        int at = nameEnd;
        byte quote = 0;
        while (at < span.Length)
        {
            byte b = span[at++];
            if (quote != 0) { if (b == quote) quote = 0; continue; }
            if (b is (byte)'"' or (byte)'\'') quote = b;
            else if (b == (byte)'>') { selfClosing = span[at - 2] == (byte)'/'; position = at; return true; }
        }
        return false;
    }

    /// <summary>After an element's start tag: skips its content and end tag, nested elements of any name included.</summary>
    private static bool TrySkipElement(ReadOnlySpan<byte> span, ref int position)
    {
        int depth = 1;
        while (depth > 0)
        {
            int open = span.Slice(position).IndexOf((byte)'<');
            if (open < 0) return false;
            position += open;
            if (position + 1 < span.Length && span[position + 1] == (byte)'/')
            {
                int close = span.Slice(position).IndexOf((byte)'>');
                if (close < 0) return false;
                position += close + 1;
                depth--;
                continue;
            }
            if (!TrySkipMarkup(span, ref position, out bool isElement)) return false;
            if (!isElement) continue;
            if (!TryReadTag(span, ref position, out _, out bool selfClosing)) return false;
            if (!selfClosing) depth++;
        }
        return true;
    }

    /// <summary>The integer value of an attribute of the start tag (either quote), as the parse reads it; null when absent or not a number.</summary>
    private static int? AttributeValue(ReadOnlySpan<byte> tag, ReadOnlySpan<byte> attributeWithEquals)
    {
        int at = tag.IndexOf(attributeWithEquals);
        if (at < 0) return null;
        int valueStart = at + attributeWithEquals.Length + 1;              // past the quote
        if (valueStart > tag.Length) return null;
        int valueEnd = valueStart;
        while (valueEnd < tag.Length && tag[valueEnd] != (byte)'"' && tag[valueEnd] != (byte)'\'') valueEnd++;
        string text = Encoding.UTF8.GetString(tag.Slice(valueStart, valueEnd - valueStart).ToArray());
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) ? value : null;
    }
}
