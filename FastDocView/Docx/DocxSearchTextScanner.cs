using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using static FastDocView.WordXmlBytes;

namespace FastDocView.Docx;

/// <summary>
/// The search text of .docx body blocks straight from their UTF-8 XML bytes: a hand-written forward scan (no
/// <see cref="XmlReader"/>, no strings, no per-node allocation) that feeds the matcher exactly what
/// <see cref="WordXmlSearchText.FeedOoxml(byte[], int, int, WordDocumentSearchMatcher)"/> feeds it - that reader
/// stays as the reference the golden check compares with:
/// - text = character content of w:t (any depth inside it), entities and character references decoded, CDATA kept;
///   everything else (w:delText, w:instrText, attributes, comments) is not text;
/// - w:tab = a space; w:br / w:cr, an empty w:p and the end of every w:p = separators;
/// - w:pPr / w:rPr, w:moveFrom (a tracked move's source), w:rt (a ruby's guide text) and mc:Fallback are skipped whole;
/// - element names are resolved by namespace (the part root's declarations plus any declared inside the block), not by
///   prefix, and only for the few local names that matter;
/// - an undefined entity or unterminated markup ends the block there (the reference reader throws there and keeps
///   what precedes).
/// Markup inside a word (runs, proofErr, bookmarks, comment ranges) produces no character, so the word still matches.
/// It can also find the block boundaries itself (<see cref="FeedBlocks"/>), so a search batch needs no per-block ranges.
/// </summary>
internal sealed class DocxSearchTextScanner
{
    private enum NamespaceKind : byte { Other, Wordprocessing, MarkupCompatibility, Diagram, Chart, Vml }

    /// <summary>
    /// Fields (null: field characters are not followed - the text of codes still is not, results are read as plain text).
    /// The caller restores it at the first block it feeds (<see cref="DocxSearchFields.Restore"/>); it runs on across the
    /// blocks of a <see cref="FeedBlocks"/> call as the index pass did.
    /// </summary>
    public DocxSearchFields? Fields { get; set; }

    /// <summary>The r:ids of the SmartArt (dgm:relIds r:dm) and chart (c:chart r:id) drawings met since the caller last cleared it (outside mc:Fallback), as the reference reader reports them.</summary>
    public List<string> DrawingRelationshipIds { get; } = new();

    private readonly WordXmlNamespaceScope _namespaces;

    public DocxSearchTextScanner(string rootOpenTag) => _namespaces = new WordXmlNamespaceScope(rootOpenTag, uri => (byte)KindOf(uri));

    private static NamespaceKind KindOf(string uri) => uri switch
    {
        WordXmlSearchText.WordprocessingNamespace or WordXmlSearchText.WordprocessingStrictNamespace => NamespaceKind.Wordprocessing,
        WordXmlSearchText.MarkupCompatibilityNamespace => NamespaceKind.MarkupCompatibility,
        WordXmlSearchText.DiagramNamespace => NamespaceKind.Diagram,
        WordXmlSearchText.VmlNamespace => NamespaceKind.Vml,
        WordXmlSearchText.ChartNamespace or WordXmlSearchText.ChartExNamespace => NamespaceKind.Chart,
        _ => NamespaceKind.Other,
    };

    private NamespaceKind Resolve(ReadOnlySpan<byte> prefix) => (NamespaceKind)_namespaces.Resolve(prefix);

    /// <summary>One block [start, start + length) of <paramref name="bytes"/> (a block whose range comes from the index).</summary>
    public void Feed(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher)
        => Scan(bytes.AsSpan(start, length), matcher, 0, -1, null, out _);

    /// <summary>
    /// A run of consecutive body blocks with no per-block ranges: <paramref name="xml"/> starts at block
    /// <paramref name="firstBlock"/>'s start, and every element that starts at the body level begins the next block (the
    /// index pass's rule: every child of w:body is a block). After each block <paramref name="blockEnded"/> gets its
    /// index (the matcher holds its count; the caller resets it). Stops at the start of block
    /// <paramref name="stopBlock"/> (<see cref="WordSearchBlocksResult.Stopped"/>, <paramref name="stoppedAt"/> = its offset in
    /// <paramref name="xml"/>) or at the end of the bytes (<see cref="WordSearchBlocksResult.Done"/>: only when the block count
    /// comes out at exactly <paramref name="stopBlock"/>). <see cref="WordSearchBlocksResult.Failed"/>: a block the reference reader
    /// would stop inside (an undefined entity, unterminated markup) or a layout that is not the index's -
    /// <paramref name="stoppedAt"/> = the failing block's index (the piece's first block when the layout is not the index's); the blocks before it were reported, the caller drops any hit from it on and reads the
    /// rest block by block.
    /// </summary>
    public WordSearchBlocksResult FeedBlocks(ReadOnlySpan<byte> xml, int firstBlock, int stopBlock, WordDocumentSearchMatcher matcher, Action<int> blockEnded, out long stoppedAt)
        => Scan(xml, matcher, firstBlock, stopBlock, blockEnded, out stoppedAt);

#if NET
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]   // optimised from the first call: a search runs in short bursts
#endif
    private WordSearchBlocksResult Scan(ReadOnlySpan<byte> xml, WordDocumentSearchMatcher matcher, int firstBlock, int stopBlock, Action<int>? blockEnded, out long stoppedAt)
    {
        _namespaces.ClearBlock();
        CurrentBlockStart = 0;
        bool segmented = blockEnded is not null;
        int block = firstBlock;
        bool blockStarted = false;
        int position = 0, end = xml.Length;
        int depth = 0;          // open elements; an element's XmlReader depth = depth before it opens
        int textDepth = -1;     // depth of the open w:t
        int skipDepth = -1;     // depth of the open skipped element (w:pPr, w:rPr, mc:Fallback)
        int instructionDepth = -1;     // depth of the open w:instrText
        int formDepth = -1;            // depth of the open w:fldChar whose w:ffData is being read
        string? pendingFieldType = null;   // that w:fldChar's type, dirty and lock flags
        bool pendingDirty = false, pendingLocked = false;
        DocxSearchFields.FormDataReader? form = null;
        string? formParent = null;
        List<(int depth, bool separated)>? simpleFields = null;   // open w:fldSimple elements
        DocxSearchFields? fields = Fields;
        stoppedAt = 0;
        while (position < end)
        {
            int next = xml.Slice(position).IndexOfAny((byte)'<', (byte)'&');
            int stop = next < 0 ? end : position + next;
            if (stop > position && skipDepth < 0)
            {
                if (instructionDepth >= 0) fields!.Instruction(Encoding.UTF8.GetString(xml.Slice(position, stop - position).ToArray()));
                else if (textDepth >= 0)
                {
                    if (fields is null || fields.IsIdle) matcher.FeedUtf8(xml.Slice(position, stop - position));
                    else if (fields.Shows(Encoding.UTF8.GetString(xml.Slice(position, stop - position).ToArray()))) matcher.FeedUtf8(xml.Slice(position, stop - position));
                }
            }
            if (next < 0) break;
            position = stop;
            if (xml[position] == (byte)'&')
            {
                int semicolon = xml.Slice(position).IndexOf((byte)';');
                if (semicolon < 0 || !TryDecodeEntity(xml.Slice(position + 1, semicolon - 1), out int codePoint)) return Fail(block, out stoppedAt);
                if (instructionDepth >= 0 && skipDepth < 0) fields!.Instruction(char.ConvertFromUtf32(codePoint));
                else if (textDepth >= 0 && skipDepth < 0 && (fields is null || fields.IsIdle || fields.Shows(char.ConvertFromUtf32(codePoint)))) matcher.FeedCodePoint(codePoint);
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
                    if (instructionDepth >= 0 && skipDepth < 0) fields!.Instruction(Encoding.UTF8.GetString(rest.Slice(9, close).ToArray()));
                    else if (textDepth >= 0 && skipDepth < 0 && (fields is null || fields.IsIdle || fields.Shows(Encoding.UTF8.GetString(rest.Slice(9, close).ToArray())))) matcher.FeedUtf8(rest.Slice(9, close));
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
                if (depth < 0) return Fail(segmented ? firstBlock : block, out stoppedAt);     // an element closed that the range did not open: not the index's layout (no block of the piece can be trusted)
                if (fields is not null && skipDepth < 0)
                {
                    if (depth == instructionDepth) instructionDepth = -1;
                    if (depth == formDepth)
                    {
                        FieldCharacter(fields, pendingFieldType, pendingDirty, pendingLocked, form!.State(), matcher);
                        formDepth = -1;
                        form = null;
                    }
                    if (simpleFields is { Count: > 0 } && simpleFields[simpleFields.Count - 1].depth == depth)
                    {
                        simpleFields.RemoveAt(simpleFields.Count - 1);
                        fields.End(matcher);
                    }
                    fields.Closed(depth);
                }
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
                if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth);
                continue;
            }
            // a start tag; at the body level it begins a block
            if (segmented && depth == 0)
            {
                if (blockStarted)
                {
                    blockEnded!(block);
                    block++;
                    if (block == stopBlock) { stoppedAt = position; return WordSearchBlocksResult.Stopped; }
                }
                blockStarted = true;
                CurrentBlockStart = position;
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
                if (attributes.IndexOf("xmlns"u8) >= 0 && !_namespaces.ReadDeclarations(attributes, elementDepth)) return Fail(block, out stoppedAt);
            }
            position = cursor + 1;
            if (skipDepth < 0)
            {
                ReadOnlySpan<byte> elementPrefix = PrefixOf(xml.Slice(elementNameStart, elementNameEnd - elementNameStart), out ReadOnlySpan<byte> local);
                if (fields is not null && IsFieldMarkup(local) && Resolve(elementPrefix) == NamespaceKind.Wordprocessing)
                {
                    ReadOnlySpan<byte> attributes = xml.Slice(elementNameEnd, cursor - elementNameEnd);
                    // a w:fldSimple's result starts at its first run-level child (the renderer separates before rendering any)
                    if (simpleFields is { Count: > 0 } && !simpleFields[simpleFields.Count - 1].separated && elementDepth == simpleFields[simpleFields.Count - 1].depth + 1
                        && (Is(local, "r") || Is(local, "hyperlink") || Is(local, "fldSimple") || Is(local, "smartTag") || Is(local, "ins") || Is(local, "sdt") || Is(local, "customXml")))
                    {
                        simpleFields[simpleFields.Count - 1] = (simpleFields[simpleFields.Count - 1].depth, true);
                        fields.Separate(matcher);
                    }
                    if (formDepth >= 0)
                    {
                        string name = Encoding.UTF8.GetString(local.ToArray());
                        form!.Element(name, AttributeValue(attributes, "val"u8) is { } v ? System.Net.WebUtility.HtmlDecode(v) : null, formParent);
                        if (name is "checkBox" or "ddList" or "textInput") formParent = name;
                    }
                    else if (Is(local, "fldChar"))
                    {
                        pendingFieldType = AttributeValue(attributes, "fldCharType"u8);
                        pendingDirty = IsOn(AttributeValue(attributes, "dirty"u8));
                        pendingLocked = IsOn(AttributeValue(attributes, "fldLock"u8));
                        if (isEmpty) FieldCharacter(fields, pendingFieldType, pendingDirty, pendingLocked, null, matcher);
                        else { formDepth = elementDepth; form = new DocxSearchFields.FormDataReader(); formParent = null; }
                    }
                    else if (Is(local, "instrText")) { if (!isEmpty) instructionDepth = elementDepth; }
                    else if (Is(local, "fldSimple"))
                    {
                        fields.Begin(IsOn(AttributeValue(attributes, "dirty"u8)), IsOn(AttributeValue(attributes, "fldLock"u8)), null);
                        fields.Instruction(AttributeValue(attributes, "instr"u8) is { } instruction ? System.Net.WebUtility.HtmlDecode(instruction) : "");
                        if (isEmpty) fields.End(matcher);
                        else (simpleFields ??= new()).Add((elementDepth, false));
                    }
                    else if (Is(local, "txbxContent")) { if (!isEmpty) fields.EnterTextBox(elementDepth); }
                }
                if (fields is not null && !fields.IsIdle && Resolve(elementPrefix) == NamespaceKind.Wordprocessing
                    && (Is(local, "tab") && !fields.Shows("\t") || (Is(local, "br") || Is(local, "cr")) && !fields.Shows(null)))
                    local = default;                                      // inside a code or a replaced result: not shown
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
                    case 2 when Is(local, "rt") && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:          // a ruby's guide text: shown above its base, not searched
                    case 8 when Is(local, "moveFrom") && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:    // a tracked move's source: hidden in Word's final view
                    case 3 when (Is(local, "pPr") || Is(local, "rPr")) && Resolve(elementPrefix) == NamespaceKind.Wordprocessing:
                        if (!isEmpty) skipDepth = elementDepth;
                        break;
                    case 8 when Is(local, "textpath") && Resolve(elementPrefix) == NamespaceKind.Vml && hasAttributes:
                        // VML WordArt: the v:textpath string is the text it shows (the reference reader feeds it the same way)
                        if (UnprefixedAttributeValue(xml.Slice(elementNameEnd, cursor - elementNameEnd), "string"u8) is { Length: > 0 } wordArt) { matcher.Separator(); matcher.Feed(wordArt.AsSpan()); matcher.Separator(); }
                        break;
                    case 6 when Is(local, "relIds") && Resolve(elementPrefix) == NamespaceKind.Diagram && hasAttributes:
                        if (AttributeValue(xml.Slice(elementNameEnd, cursor - elementNameEnd), "dm"u8) is { } diagram) DrawingRelationshipIds.Add(diagram);
                        break;
                    case 5 when Is(local, "chart") && Resolve(elementPrefix) == NamespaceKind.Chart && hasAttributes:
                        if (AttributeValue(xml.Slice(elementNameEnd, cursor - elementNameEnd), "id"u8) is { } chart) DrawingRelationshipIds.Add(chart);
                        break;
                    case 8 when Is(local, "Fallback") && Resolve(elementPrefix) == NamespaceKind.MarkupCompatibility:
                        if (!isEmpty) skipDepth = elementDepth;
                        break;
                }
            }
            if (isEmpty) { if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth); }
            else depth++;
        }
        if (segmented)
        {
            if (blockStarted) { blockEnded!(block); block++; }
            if (block != stopBlock) return Fail(firstBlock, out stoppedAt);     // not the index's count: no block of the piece can be trusted
        }
        stoppedAt = end;
        return WordSearchBlocksResult.Done;
    }

    /// <summary>
    /// Where the scanner's current block starts in the bytes of the last <see cref="FeedBlocks"/>: inside the block-ended
    /// callback the block just ended, after a <see cref="WordSearchBlocksResult.Failed"/> the failing block - the caller
    /// checks it against the index's offset before it trusts the scanner's block numbering.
    /// </summary>
    public long CurrentBlockStart { get; private set; }

    /// <summary>The value of the prefixed attribute whose local name is <paramref name="localName"/> (r:dm, r:id) in an open tag's attribute text; null when none.</summary>
    private static string? AttributeValue(ReadOnlySpan<byte> attributes, ReadOnlySpan<byte> localName)
    {
        int position = 0;
        while (position < attributes.Length)
        {
            int equals = attributes.Slice(position).IndexOf((byte)'=');
            if (equals < 0) return null;
            ReadOnlySpan<byte> name = attributes.Slice(position, equals);
            while (name.Length > 0 && name[0] <= (byte)' ') name = name.Slice(1);
            while (name.Length > 0 && name[name.Length - 1] <= (byte)' ') name = name.Slice(0, name.Length - 1);
            int quoteAt = position + equals + 1;
            while (quoteAt < attributes.Length && attributes[quoteAt] is not ((byte)'"' or (byte)'\'')) quoteAt++;
            if (quoteAt >= attributes.Length) return null;
            int close = attributes.Slice(quoteAt + 1).IndexOf(attributes[quoteAt]);
            if (close < 0) return null;
            int colon = name.IndexOf((byte)':');
            if (colon > 0 && name.Slice(colon + 1).SequenceEqual(localName))
                return Encoding.UTF8.GetString(attributes.Slice(quoteAt + 1, close).ToArray());
            position = quoteAt + 1 + close + 1;
        }
        return null;
    }

    /// <summary>The value of the unprefixed attribute <paramref name="name"/> in an open tag's attribute text, entities decoded; null when none.</summary>
    private static string? UnprefixedAttributeValue(ReadOnlySpan<byte> attributes, ReadOnlySpan<byte> name)
    {
        int position = 0;
        while (position < attributes.Length)
        {
            int equals = attributes.Slice(position).IndexOf((byte)'=');
            if (equals < 0) return null;
            ReadOnlySpan<byte> attribute = attributes.Slice(position, equals);
            while (attribute.Length > 0 && attribute[0] <= (byte)' ') attribute = attribute.Slice(1);
            while (attribute.Length > 0 && attribute[attribute.Length - 1] <= (byte)' ') attribute = attribute.Slice(0, attribute.Length - 1);
            int quoteAt = position + equals + 1;
            while (quoteAt < attributes.Length && attributes[quoteAt] is not ((byte)'"' or (byte)'\'')) quoteAt++;
            if (quoteAt >= attributes.Length) return null;
            int close = attributes.Slice(quoteAt + 1).IndexOf(attributes[quoteAt]);
            if (close < 0) return null;
            if (attribute.SequenceEqual(name))
                return System.Net.WebUtility.HtmlDecode(Encoding.UTF8.GetString(attributes.Slice(quoteAt + 1, close).ToArray()));
            position = quoteAt + 1 + close + 1;
        }
        return null;
    }

    private static bool IsOn(string? value) => value is "1" or "true" or "on";

    /// <summary>Local names the field reading looks at (cheap length filter first).</summary>
    private static bool IsFieldMarkup(ReadOnlySpan<byte> local) => local.Length switch
    {
        1 => local[0] == (byte)'r',
        3 => Is(local, "ins") || Is(local, "sdt"),
        6 => Is(local, "result") || Is(local, "ddList"),
        7 => Is(local, "fldChar") || Is(local, "default") || Is(local, "checked"),
        8 => Is(local, "checkBox") || Is(local, "smartTag"),
        9 => Is(local, "instrText") || Is(local, "fldSimple") || Is(local, "hyperlink") || Is(local, "customXml") || Is(local, "listEntry") || Is(local, "textInput"),
        11 => Is(local, "txbxContent"),
        _ => false,
    };

    private static void FieldCharacter(DocxSearchFields fields, string? type, bool dirty, bool locked, Fields.WordFormFieldState? form, WordDocumentSearchMatcher matcher)
    {
        if (type == "begin") fields.Begin(dirty, locked, form);
        else if (type == "separate") fields.Separate(matcher);
        else if (type == "end") fields.End(matcher);
    }

    private static WordSearchBlocksResult Fail(int block, out long stoppedAt) { stoppedAt = block; return WordSearchBlocksResult.Failed; }
}
