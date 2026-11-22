using System.Text;
using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// styles.xml cut into its top-level elements without parsing it: the byte range of w:docDefaults and of every w:style (with
/// the w:styleId, w:type and w:default its start tag carries), so a style is parsed only when a window uses it
/// (<see cref="DocxStyleCascade"/>) - a converted or long-lived document's styles.xml holds thousands of styles (the sample:
/// 850 KB), of which a window uses a few. The same walk as <see cref="DocxNumberingPartSlices"/>: a part whose shape it cannot
/// follow (another prefix for the main namespace, a top-level element under another prefix, a style without a w:styleId,
/// an encoding other than UTF-8, unbalanced tags, an entity in a style id) gives null, and the caller parses
/// the part whole, as before.
/// </summary>
internal sealed class DocxStylesPartSlices
{
    public readonly record struct Slice(int Start, int Length);

    private static readonly byte[] MainNamespace = Encoding.ASCII.GetBytes("xmlns:w=\"" + DocxStyleCascade.W.NamespaceName + "\"");

    public byte[] Bytes { get; }
    public byte[] RootOpenTag { get; }
    public Slice? DocumentDefaults { get; private set; }
    /// <summary>Every w:style by w:styleId (a repeated id: the last one, as the whole-part parse keeps it).</summary>
    public Dictionary<string, Slice> StylesById { get; } = new(StringComparer.Ordinal);
    /// <summary>The ids of the last w:default paragraph and character styles, as the whole-part parse finds them.</summary>
    public string? DefaultParagraphStyleId { get; private set; }
    public string? DefaultCharacterStyleId { get; private set; }

    private DocxStylesPartSlices(byte[] bytes, byte[] rootOpenTag)
    {
        Bytes = bytes;
        RootOpenTag = rootOpenTag;
    }

    public static DocxStylesPartSlices? TryCut(byte[] bytes)
    {
        var span = new ReadOnlySpan<byte>(bytes);
        if (!IsUtf8Declared(span)) return null;                                 // the slices are parsed as UTF-8 (no declaration of their own)
        int position = 0;
        while (true)
        {
            int open = span.Slice(position).IndexOf((byte)'<');
            if (open < 0) return null;
            position += open;
            if (!DocxNumberingPartSlices.TrySkipMarkup(span, ref position, out bool isElement)) return null;
            if (isElement) break;
        }
        int rootStart = position;
        if (!DocxNumberingPartSlices.TryReadTag(span, ref position, out ReadOnlySpan<byte> rootName, out bool rootSelfClosing)) return null;
        if (!rootName.SequenceEqual("w:styles"u8) || rootSelfClosing) return null;
        ReadOnlySpan<byte> rootTag = span.Slice(rootStart, position - rootStart);
        if (rootTag.IndexOf(MainNamespace) < 0) return null;
        var slices = new DocxStylesPartSlices(bytes, rootTag.ToArray());

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
            if (!DocxNumberingPartSlices.TrySkipMarkup(span, ref position, out bool isElement)) return null;
            if (!isElement) continue;
            int elementStart = position;
            if (!DocxNumberingPartSlices.TryReadTag(span, ref position, out ReadOnlySpan<byte> name, out bool selfClosing)) return null;
            // a top-level element under another prefix (or the default namespace) may still be a w:style the whole parse reads
            if (depth == 1 && !name.StartsWith("w:"u8)) return null;
            bool isStyle = depth == 1 && name.SequenceEqual("w:style"u8);
            bool isDefaults = depth == 1 && name.SequenceEqual("w:docDefaults"u8);
            if (!isStyle && !isDefaults)
            {
                if (!selfClosing) depth++;
                continue;
            }
            ReadOnlySpan<byte> tag = span.Slice(elementStart, position - elementStart);
            if (!selfClosing && !DocxNumberingPartSlices.TrySkipElement(span, ref position)) return null;
            var slice = new Slice(elementStart, position - elementStart);
            if (isDefaults) { slices.DocumentDefaults ??= slice; continue; }      // the whole-part parse reads the first
            if (!TryAttribute(tag, "w:styleId"u8, out string? id)) return null;
            if (id is null) return null;                                        // no w:styleId: maybe the id under another prefix
            slices.StylesById[id] = slice;
            if (!TryAttribute(tag, "w:default"u8, out string? isDefaultText) || !TryAttribute(tag, "w:type"u8, out string? type)) return null;
            // DocxStyleCascade.IsOn: a missing w:default means NOT default; present, "1" / "true" / "on" is on
            bool isDefault = isDefaultText is "1" or "true" or "on";
            if (isDefault && type == "paragraph") slices.DefaultParagraphStyleId = id;
            if (isDefault && type == "character") slices.DefaultCharacterStyleId = id;
        }
        return slices;
    }

    /// <summary>The style element of a slice, parsed inside the root start tag (its namespace declarations), as the whole-part parse reads it.</summary>
    public XElement Parse(Slice slice)
    {
        byte[] close = "</w:styles>"u8.ToArray();
        using var wrapped = new MemoryStream(RootOpenTag.Length + slice.Length + close.Length);
        wrapped.Write(RootOpenTag, 0, RootOpenTag.Length);
        wrapped.Write(Bytes, slice.Start, slice.Length);
        wrapped.Write(close, 0, close.Length);
        wrapped.Position = 0;
        return WordDocumentXml.ReadFirstChildElement(wrapped, LoadOptions.None, tolerateUndeclaredPrefixes: true, "styles", keepContentBeforeFirstError: true);
    }

    /// <summary>
    /// An attribute's value in a start tag (either quote): null when absent; false when it cannot be read as plain text here
    /// (an entity or character reference: the caller parses the part whole).
    /// </summary>
    private static bool TryAttribute(ReadOnlySpan<byte> tag, ReadOnlySpan<byte> attributeName, out string? value)
    {
        value = null;
        int from = 0;
        while (true)
        {
            int at = tag.Slice(from).IndexOf(attributeName);
            if (at < 0) return true;
            at += from;
            from = at + attributeName.Length;
            if (at == 0 || !IsSpace(tag[at - 1])) continue;                    // the end of a longer name
            int position = at + attributeName.Length;
            while (position < tag.Length && IsSpace(tag[position])) position++;
            if (position >= tag.Length || tag[position] != (byte)'=') continue; // a longer name that starts the same
            position++;
            while (position < tag.Length && IsSpace(tag[position])) position++;
            if (position >= tag.Length || tag[position] is not ((byte)'"' or (byte)'\'')) return false;
            byte quote = tag[position];
            int end = tag.Slice(position + 1).IndexOf(quote);
            if (end < 0) return false;
            ReadOnlySpan<byte> raw = tag.Slice(position + 1, end);
            if (raw.IndexOf((byte)'&') >= 0 || raw.IndexOfAny((byte)'\t', (byte)'\r', (byte)'\n') >= 0) return false;   // the parser would decode or normalise these
            value = Encoding.UTF8.GetString(raw.ToArray());
            return true;
        }
    }

    /// <summary>False when the XML declaration names an encoding other than UTF-8 (a UTF-8 byte order mark or no declaration: UTF-8).</summary>
    private static bool IsUtf8Declared(ReadOnlySpan<byte> span)
    {
        if (span.Length >= 2 && (span[0] is 0xFF or 0xFE || span[1] == 0)) return false;   // UTF-16 (a byte order mark, or '<' then a zero byte)
        int start = span.StartsWith(new byte[] { 0xEF, 0xBB, 0xBF }) ? 3 : 0;
        if (!span.Slice(start).StartsWith("<?xml"u8)) return true;
        int end = span.Slice(start).IndexOf("?>"u8);
        if (end < 0) return false;
        ReadOnlySpan<byte> declaration = span.Slice(start, end);
        if (!TryAttribute(declaration, "encoding"u8, out string? encoding)) return false;
        return encoding is null || string.Equals(encoding, "UTF-8", StringComparison.OrdinalIgnoreCase) || string.Equals(encoding, "UTF8", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSpace(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';
}
