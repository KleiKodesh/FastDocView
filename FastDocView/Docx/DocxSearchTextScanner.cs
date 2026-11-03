using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;

namespace FastDocView.Docx;

/// <summary>
/// The search text of .docx body blocks straight from their UTF-8 XML bytes: a hand-written forward scan (no
/// <see cref="XmlReader"/>, no strings, no per-node allocation) that feeds the matcher exactly what
/// <see cref="WordXmlSearchText.FeedOoxml(byte[], int, int, WordDocumentSearchMatcher)"/> feeds it - that reader
/// stays as the reference the golden check compares with:
/// - text = character content of w:t (any depth inside it), entities and character references decoded, CDATA kept;
///   everything else (w:delText, w:instrText, attributes, comments) is not text;
/// - w:tab = a space; w:br / w:cr, an empty w:p and the end of every w:p = separators;
/// - w:pPr / w:rPr and mc:Fallback are skipped whole;
/// - element names are resolved by namespace (the part root's declarations plus any declared inside the block), not by
///   prefix, and only for the few local names that matter;
/// - an undefined entity or unterminated markup ends the block there (the reference reader throws there and keeps
///   what precedes).
/// Markup inside a word (runs, proofErr, bookmarks, comment ranges) produces no character, so the word still matches.
/// It can also find the block boundaries itself (<see cref="FeedBlocks"/>), so a search batch needs no per-block ranges.
/// </summary>
internal sealed class DocxSearchTextScanner
{
    private enum NamespaceKind : byte { Other, Wordprocessing, MarkupCompatibility }

    private readonly List<(byte[] prefix, NamespaceKind kind)> _rootNamespaces = new();
    private readonly List<(byte[] prefix, NamespaceKind kind, int depth)> _blockNamespaces = new();   // declared inside the block, innermost last

    public DocxSearchTextScanner(string rootOpenTag)
    {
        // the root tag's own declarations, read by a parser (never by hand)
        string rootName = DocxXmlFragment.RootElementName(rootOpenTag);
        using XmlReader root = XmlReader.Create(new StringReader(rootOpenTag + "</" + rootName + ">"), new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
        root.MoveToContent();
        if (root.MoveToFirstAttribute())
            do
            {
                if (root.Prefix == "xmlns") _rootNamespaces.Add((Encoding.UTF8.GetBytes(root.LocalName), KindOf(root.Value)));
                else if (root.Name == "xmlns") _rootNamespaces.Add(([], KindOf(root.Value)));
            } while (root.MoveToNextAttribute());
    }

    private static NamespaceKind KindOf(string uri) => uri switch
    {
        WordXmlSearchText.WordprocessingNamespace or WordXmlSearchText.WordprocessingStrictNamespace => NamespaceKind.Wordprocessing,
        WordXmlSearchText.MarkupCompatibilityNamespace => NamespaceKind.MarkupCompatibility,
        _ => NamespaceKind.Other,
    };

    private static NamespaceKind KindOf(ReadOnlySpan<byte> uri)   // a declaration inside a block: rare
        => KindOf(uri.IndexOf((byte)'&') >= 0 ? DecodeAttribute(uri) : Encoding.UTF8.GetString(uri.ToArray()));

    private static string DecodeAttribute(ReadOnlySpan<byte> raw)
    {
        // a namespace URI written with entities - decoded by a parser
        string text = Encoding.UTF8.GetString(raw.ToArray());
        try
        {
            using XmlReader reader = XmlReader.Create(new StringReader("<a b=\"" + text.Replace("\"", "&quot;") + "\"/>"));
            reader.MoveToContent();
            return reader.GetAttribute("b") ?? "";
        }
        catch (XmlException) { return text; }
    }

    private NamespaceKind Resolve(ReadOnlySpan<byte> prefix)
    {
        for (int i = _blockNamespaces.Count - 1; i >= 0; i--)
            if (prefix.SequenceEqual(_blockNamespaces[i].prefix)) return _blockNamespaces[i].kind;
        foreach ((byte[] declared, NamespaceKind kind) in _rootNamespaces)
            if (prefix.SequenceEqual(declared)) return kind;
        return NamespaceKind.Other;
    }

    private static bool IsNameEnd(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/';

    private static bool IsSpace(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';

    private static bool Is(ReadOnlySpan<byte> name, string ascii)
    {
        if (name.Length != ascii.Length) return false;
        for (int i = 0; i < name.Length; i++) if (name[i] != ascii[i]) return false;
        return true;
    }

    /// <summary>One block [start, start + length) of <paramref name="bytes"/> (a block whose range comes from the index).</summary>
    public void Feed(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher)
        => Scan(bytes.AsSpan(start, length), matcher, 0, -1, null, out _);

    /// <summary>The outcome of <see cref="FeedBlocks"/>.</summary>
    public enum BlocksResult { Done, Stopped, Failed }

    /// <summary>
    /// A run of consecutive body blocks with no per-block ranges: <paramref name="xml"/> starts at block
    /// <paramref name="firstBlock"/>'s start, and every element that starts at the body level begins the next block (the
    /// index pass's rule: every child of w:body is a block). After each block <paramref name="blockEnded"/> gets its
    /// index (the matcher holds its count; the caller resets it). Stops at the start of block
    /// <paramref name="stopBlock"/> (<see cref="BlocksResult.Stopped"/>, <paramref name="stoppedAt"/> = its offset in
    /// <paramref name="xml"/>) or at the end of the bytes (<see cref="BlocksResult.Done"/>: only when the block count
    /// comes out at exactly <paramref name="stopBlock"/>). <see cref="BlocksResult.Failed"/>: a block the reference reader
    /// would stop inside (an undefined entity, unterminated markup) or a layout that is not the index's -
    /// <paramref name="stoppedAt"/> = the failing block's index; the blocks before it were reported, the caller reads the
    /// rest block by block.
    /// </summary>
    public BlocksResult FeedBlocks(ReadOnlySpan<byte> xml, int firstBlock, int stopBlock, WordDocumentSearchMatcher matcher, Action<int> blockEnded, out long stoppedAt)
        => Scan(xml, matcher, firstBlock, stopBlock, blockEnded, out stoppedAt);

#if NET
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]   // optimised from the first call: a search runs in short bursts
#endif
    private BlocksResult Scan(ReadOnlySpan<byte> xml, WordDocumentSearchMatcher matcher, int firstBlock, int stopBlock, Action<int>? blockEnded, out long stoppedAt)
    {
        _blockNamespaces.Clear();
        bool segmented = blockEnded is not null;
        int block = firstBlock;
        bool blockStarted = false;
        int position = 0, end = xml.Length;
        int depth = 0;          // open elements; an element's XmlReader depth = depth before it opens
        int textDepth = -1;     // depth of the open w:t
        int skipDepth = -1;     // depth of the open skipped element (w:pPr, w:rPr, mc:Fallback)
        stoppedAt = 0;
        while (position < end)
        {
            int next = xml.Slice(position).IndexOfAny((byte)'<', (byte)'&');
            int stop = next < 0 ? end : position + next;
            if (stop > position && textDepth >= 0 && skipDepth < 0) matcher.FeedUtf8(xml.Slice(position, stop - position));
            if (next < 0) break;
            position = stop;
            if (xml[position] == (byte)'&')
            {
                int semicolon = xml.Slice(position).IndexOf((byte)';');
                if (semicolon < 0 || !TryDecodeEntity(xml.Slice(position + 1, semicolon - 1), out int codePoint)) return Fail(block, out stoppedAt);
                if (textDepth >= 0 && skipDepth < 0) matcher.FeedCodePoint(codePoint);
                position += semicolon + 1;
                continue;
            }
            // '<'
            if (position + 1 >= end) return Fail(block, out stoppedAt);
            byte kind = xml[position + 1];
            if (kind == (byte)'!')
            {
                ReadOnlySpan<byte> rest = xml.Slice(position);
                if (rest.Length >= 4 && rest[2] == (byte)'-' && rest[3] == (byte)'-')
                {
                    int close = rest.Slice(4).IndexOf("-->"u8);
                    if (close < 0) return Fail(block, out stoppedAt);
                    position += 4 + close + 3;
                    continue;
                }
                if (rest.StartsWith("<![CDATA["u8))
                {
                    int close = rest.Slice(9).IndexOf("]]>"u8);
                    if (close < 0) return Fail(block, out stoppedAt);
                    if (textDepth >= 0 && skipDepth < 0) matcher.FeedUtf8(rest.Slice(9, close));
                    position += 9 + close + 3;
                    continue;
                }
                return Fail(block, out stoppedAt);     // a DOCTYPE inside a fragment: the reader stops there
            }
            if (kind == (byte)'?')
            {
                int close = xml.Slice(position + 2).IndexOf("?>"u8);
                if (close < 0) return Fail(block, out stoppedAt);
                position += 2 + close + 2;
                continue;
            }
            if (kind == (byte)'/')
            {
                int nameStart = position + 2;
                int close = xml.Slice(nameStart).IndexOf((byte)'>');
                if (close < 0) return Fail(block, out stoppedAt);
                position = nameStart + close + 1;
                depth--;
                if (depth < 0) return Fail(block, out stoppedAt);     // an element closed that the range did not open: not the index's layout
                if (skipDepth >= 0)
                {
                    if (depth == skipDepth) skipDepth = -1;
                }
                else if (depth == textDepth) textDepth = -1;
                else
                {
                    int nameEnd = nameStart;
                    while (nameEnd < position - 1 && !IsSpace(xml[nameEnd])) nameEnd++;
                    ReadOnlySpan<byte> prefix = PrefixOf(xml.Slice(nameStart, nameEnd - nameStart), out ReadOnlySpan<byte> local);
                    if (local.Length == 1 && local[0] == (byte)'p' && Resolve(prefix) == NamespaceKind.Wordprocessing) matcher.Separator();
                }
                if (_blockNamespaces.Count > 0) PopNamespaces(depth);
                continue;
            }
            // a start tag; at the body level it begins a block
            if (segmented && depth == 0)
            {
                if (blockStarted)
                {
                    blockEnded!(block);
                    block++;
                    if (block == stopBlock) { stoppedAt = position; return BlocksResult.Stopped; }
                }
                blockStarted = true;
            }
            int elementNameStart = position + 1, elementNameEnd = elementNameStart;
            while (elementNameEnd < end && !IsNameEnd(xml[elementNameEnd])) elementNameEnd++;
            // the tag's end: the first '>' outside a quoted attribute value
            int cursor = elementNameEnd;
            bool hasAttributes = false;
            while (true)
            {
                int found = xml.Slice(cursor).IndexOfAny((byte)'>', (byte)'"', (byte)'\'');
                if (found < 0) return Fail(block, out stoppedAt);
                cursor += found;
                if (xml[cursor] == (byte)'>') break;
                hasAttributes = true;
                int closeQuote = xml.Slice(cursor + 1).IndexOf(xml[cursor]);
                if (closeQuote < 0) return Fail(block, out stoppedAt);
                cursor += 1 + closeQuote + 1;
            }
            bool isEmpty = cursor - 1 >= elementNameEnd && xml[cursor - 1] == (byte)'/';
            int elementDepth = depth;
            if (hasAttributes)
            {
                ReadOnlySpan<byte> attributes = xml.Slice(elementNameEnd, cursor - elementNameEnd);
                if (attributes.IndexOf("xmlns"u8) >= 0 && !ReadNamespaceDeclarations(attributes, elementDepth)) return Fail(block, out stoppedAt);
            }
            position = cursor + 1;
            if (skipDepth < 0)
            {
                ReadOnlySpan<byte> elementPrefix = PrefixOf(xml.Slice(elementNameStart, elementNameEnd - elementNameStart), out ReadOnlySpan<byte> local);
                switch (local.Length)
                {
                    case 1 when local[0] == (byte)'t':
                        if (!isEmpty && Resolve(elementPrefix) == NamespaceKind.Wordprocessing) textDepth = elementDepth;
                        break;
                    case 1 when local[0] == (byte)'p':
                        if (isEmpty && Resolve(elementPrefix) == NamespaceKind.Wordprocessing) matcher.Separator();
                        break;
                    case 2 when (Is(local, "br") || Is(local, "cr")) && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:
                        matcher.Separator();
                        break;
                    case 3 when Is(local, "tab") && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:
                        matcher.Feed(' ');
                        break;
                    case 3 when (Is(local, "pPr") || Is(local, "rPr")) && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:
                        if (!isEmpty) skipDepth = elementDepth;
                        break;
                    case 8 when Is(local, "Fallback") && Resolve(elementPrefix) == NamespaceKind.MarkupCompatibility:
                        if (!isEmpty) skipDepth = elementDepth;
                        break;
                }
            }
            if (isEmpty) { if (_blockNamespaces.Count > 0) PopNamespaces(depth); }
            else depth++;
        }
        if (segmented)
        {
            if (blockStarted) { blockEnded!(block); block++; }
            if (block != stopBlock) return Fail(block, out stoppedAt);
        }
        stoppedAt = end;
        return BlocksResult.Done;
    }

    private static BlocksResult Fail(int block, out long stoppedAt) { stoppedAt = block; return BlocksResult.Failed; }

    /// <summary>The xmlns / xmlns:p attributes of one start tag's attribute text; false when it is not well formed.</summary>
    private bool ReadNamespaceDeclarations(ReadOnlySpan<byte> attributes, int elementDepth)
    {
        int cursor = 0, end = attributes.Length;
        while (cursor < end)
        {
            byte b = attributes[cursor];
            if (IsSpace(b) || b == (byte)'/') { cursor++; continue; }
            int equals = attributes.Slice(cursor).IndexOf((byte)'=');
            if (equals < 0) return false;
            ReadOnlySpan<byte> attributeName = attributes.Slice(cursor, equals);
            int trimmed = attributeName.Length;
            while (trimmed > 0 && IsSpace(attributeName[trimmed - 1])) trimmed--;
            attributeName = attributeName.Slice(0, trimmed);
            cursor += equals + 1;
            while (cursor < end && IsSpace(attributes[cursor])) cursor++;
            if (cursor >= end) return false;
            byte quote = attributes[cursor];
            if (quote is not ((byte)'"' or (byte)'\'')) return false;
            int valueEnd = attributes.Slice(cursor + 1).IndexOf(quote);
            if (valueEnd < 0) return false;
            ReadOnlySpan<byte> value = attributes.Slice(cursor + 1, valueEnd);
            cursor += 1 + valueEnd + 1;
            if (attributeName.Length >= 5 && attributeName.StartsWith("xmlns"u8))
            {
                if (attributeName.Length == 5) _blockNamespaces.Add(([], KindOf(value), elementDepth));
                else if (attributeName[5] == (byte)':') _blockNamespaces.Add((attributeName.Slice(6).ToArray(), KindOf(value), elementDepth));
            }
        }
        return true;
    }

    private void PopNamespaces(int depth)
    {
        while (_blockNamespaces.Count > 0 && _blockNamespaces[_blockNamespaces.Count - 1].depth >= depth) _blockNamespaces.RemoveAt(_blockNamespaces.Count - 1);
    }

    private static ReadOnlySpan<byte> PrefixOf(ReadOnlySpan<byte> name, out ReadOnlySpan<byte> local)
    {
        int colon = name.IndexOf((byte)':');
        if (colon < 0) { local = name; return []; }
        local = name.Slice(colon + 1);
        return name.Slice(0, colon);
    }

    /// <summary>The five predefined entities and decimal / hexadecimal character references (XML 1.0 4.1, 4.6).</summary>
    private static bool TryDecodeEntity(ReadOnlySpan<byte> name, out int codePoint)
    {
        codePoint = 0;
        if (name.Length == 0) return false;
        if (name[0] == (byte)'#')
        {
            bool hex = name.Length > 1 && name[1] == (byte)'x';
            int digits = hex ? 2 : 1;
            if (digits >= name.Length) return false;
            long value = 0;
            for (int i = digits; i < name.Length; i++)
            {
                int b = name[i], digit;
                if (b is >= '0' and <= '9') digit = b - '0';
                else if (hex && b is >= 'a' and <= 'f') digit = b - 'a' + 10;
                else if (hex && b is >= 'A' and <= 'F') digit = b - 'A' + 10;
                else return false;
                value = value * (hex ? 16 : 10) + digit;
                if (value > 0x10FFFF) return false;
            }
            if (value is >= 0xD800 and <= 0xDFFF) return false;
            codePoint = (int)value;
            return true;
        }
        if (Is(name, "amp")) codePoint = '&';
        else if (Is(name, "lt")) codePoint = '<';
        else if (Is(name, "gt")) codePoint = '>';
        else if (Is(name, "quot")) codePoint = '"';
        else if (Is(name, "apos")) codePoint = '\'';
        else return false;
        return true;
    }
}
