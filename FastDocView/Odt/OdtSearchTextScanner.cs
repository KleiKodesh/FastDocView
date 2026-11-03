using System.Runtime.CompilerServices;
using System.Text;
using static FastDocView.WordXmlBytes;

namespace FastDocView.Odt;

/// <summary>
/// The search text of .odt / .fodt body blocks straight from their UTF-8 XML bytes: a hand-written forward scan (no
/// XmlReader, no strings, no per-node allocation) that feeds the matcher exactly what
/// <see cref="WordXmlSearchText.FeedOdf"/> feeds it - that reader stays as the reference the golden check compares with:
/// - text = character data inside text:p / text:h (any depth, nested paragraphs counted), entities and character
///   references decoded, CDATA kept; white space (space, tab, CR, LF) collapses to one space (ODF 1.3 part 3 6.1.2),
///   the "last was a space" state carried through the whole block as the reference carries it;
/// - text:s = its text:c spaces (default 1), text:tab a space, text:line-break and every text:p / text:h start and end
///   separators;
/// - text:note-citation, text:tracked-changes, office:annotation and (unless notes are included) text:note-body are
///   skipped whole;
/// - element names are resolved by namespace (the root's declarations plus any declared inside the block);
/// - an undefined entity or unterminated markup ends the block there.
/// Markup inside a word (text:span, text:soft-page-break, bookmarks, reference marks, niqqud in its own span) adds no
/// character, so the word still matches.
/// It can also find the block boundaries itself (<see cref="FeedBlocks"/>), with the body scanner's rule
/// (<see cref="OdtBodyBlockScanner"/>: a text:p / text:h / table:table that starts at block level - in office:text, a
/// section, an index body or a list item - is a block; lists, sections and everything else between blocks are not
/// text), so a search batch needs no per-block ranges.
/// </summary>
internal sealed class OdtSearchTextScanner
{
    /// <summary>An element open around the blocks, as far as finding the next block start is concerned.</summary>
    internal enum Frame : byte
    {
        /// <summary>Above office:text (the root, office:body): no block starts here.</summary>
        Outer,
        /// <summary>office:text, text:section, text:numbered-paragraph, text:index-body: a p / h / table here is a block.</summary>
        Container,
        /// <summary>An index element (text:table-of-content...): only its text:index-body holds blocks.</summary>
        Index,
        /// <summary>text:list: only list items.</summary>
        List,
        /// <summary>text:list-item / text:list-header: a p / h / table here is a block.</summary>
        ListItem,
        /// <summary>Anything else: not content.</summary>
        Skip,
    }

    private enum NamespaceKind : byte { Other, Text, Office }

    private enum Outside : byte { Other, BlockElement, List, ListItem, Container, IndexElement, IndexBody }

    private readonly WordXmlNamespaceScope _namespaces;
    private readonly (byte[] name, Outside kind)[] _outsideNames;
    private readonly List<Frame> _frames = new();

    public OdtSearchTextScanner(string rootOpenTag)
    {
        _namespaces = new WordXmlNamespaceScope(rootOpenTag, uri => (byte)(uri switch
        {
            WordXmlSearchText.OdfTextNamespace => NamespaceKind.Text,
            WordXmlSearchText.OdfOfficeNamespace => NamespaceKind.Office,
            _ => NamespaceKind.Other,
        }));
        // between blocks names are classified as the body scanner classifies them: by the prefixes the root declares
        Func<string, string> qualified = OdtBodyBlockScanner.QualifiedNameFunction(rootOpenTag);
        var names = new List<(byte[], Outside)>();
        void Add(string canonical, Outside kind) => names.Add((Encoding.UTF8.GetBytes(qualified(canonical)), kind));
        Add("text:p", Outside.BlockElement); Add("text:h", Outside.BlockElement); Add("table:table", Outside.BlockElement);
        Add("text:list", Outside.List); Add("text:list-item", Outside.ListItem); Add("text:list-header", Outside.ListItem);
        Add("text:section", Outside.Container); Add("text:numbered-paragraph", Outside.Container); Add("text:index-body", Outside.IndexBody);
        foreach (string index in OdtBodyBlockScanner.IndexElementNames) Add(index, Outside.IndexElement);
        _outsideNames = names.ToArray();
    }

    /// <summary>The frames open at the start of the block a <see cref="FeedBlocks"/> run stopped at (bottom first).</summary>
    public Frame[] StoppedFrames() => _frames.ToArray();

    /// <summary>One block [start, start + length) of <paramref name="bytes"/> (a block whose range comes from the index).</summary>
    public void Feed(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher, bool includeNotes)
        => Scan(bytes.AsSpan(start, length), matcher, includeNotes, 0, -1, null, null, out _);

    /// <summary>
    /// A run of consecutive body blocks with no per-block ranges: <paramref name="xml"/> starts at block
    /// <paramref name="firstBlock"/>'s start tag, with <paramref name="frames"/> open around it (from an index anchor's scan
    /// state, or the frames a previous run stopped with). After each block <paramref name="blockEnded"/> gets its index
    /// (the matcher holds its count; the caller resets it). Stops at the start of block <paramref name="stopBlock"/>
    /// (<see cref="WordSearchBlocksResult.Stopped"/>, <paramref name="stoppedAt"/> = its offset, <see cref="StoppedFrames"/>
    /// the frames there) or at the end of the bytes (<see cref="WordSearchBlocksResult.Done"/>: only when the bytes end
    /// between blocks and the count comes out at exactly <paramref name="stopBlock"/>). <see cref="WordSearchBlocksResult.Failed"/>:
    /// <paramref name="stoppedAt"/> = a block the reference reader would stop inside (an undefined entity, unterminated
    /// markup), or <paramref name="firstBlock"/> when the layout is not the index's; the caller drops the hits from there
    /// on and reads the rest block by block.
    /// </summary>
    public WordSearchBlocksResult FeedBlocks(ReadOnlySpan<byte> xml, int firstBlock, int stopBlock, IReadOnlyList<Frame> frames,
                                             WordDocumentSearchMatcher matcher, bool includeNotes, Action<int> blockEnded, out long stoppedAt)
        => Scan(xml, matcher, includeNotes, firstBlock, stopBlock, frames, blockEnded, out stoppedAt);

    private Outside ClassifyOutside(ReadOnlySpan<byte> qualifiedName)
    {
        foreach ((byte[] name, Outside kind) in _outsideNames)
            if (qualifiedName.SequenceEqual(name)) return kind;
        return Outside.Other;
    }

    private bool _lastWasSpace;

#if NET
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]   // optimised from the first call: a search runs in short bursts
#endif
    private WordSearchBlocksResult Scan(ReadOnlySpan<byte> xml, WordDocumentSearchMatcher matcher, bool includeNotes, int firstBlock, int stopBlock,
                                        IReadOnlyList<Frame>? frames, Action<int>? blockEnded, out long stoppedAt)
    {
        bool segmented = blockEnded is not null;
        _frames.Clear();
        if (frames is not null) _frames.AddRange(frames);
        _namespaces.ClearBlock();
        int block = firstBlock;
        bool inBlock = !segmented;      // a per-block range is the block element itself
        int position = 0, end = xml.Length;
        int depth = 0;                  // open elements inside the block; an element's XmlReader depth = depth before it opens
        int skipDepth = -1;             // depth of the open skipped element
        int paragraphDepth = 0;         // open text:p / text:h
        _lastWasSpace = false;
        stoppedAt = 0;
        while (position < end)
        {
            int next = xml.Slice(position).IndexOfAny((byte)'<', (byte)'&');
            int stop = next < 0 ? end : position + next;
            if (stop > position && inBlock && paragraphDepth > 0 && skipDepth < 0) FeedText(xml.Slice(position, stop - position), matcher);
            if (next < 0) break;
            position = stop;
            if (xml[position] == (byte)'&')
            {
                int semicolon = xml.Slice(position).IndexOf((byte)';');
                if (!inBlock) { position++; continue; }   // between blocks: no reader sees it
                if (semicolon < 0 || !TryDecodeEntity(xml.Slice(position + 1, semicolon - 1), out int codePoint)) return Fail(segmented, block, out stoppedAt);
                if (paragraphDepth > 0 && skipDepth < 0) FeedCodePoint(codePoint, matcher);
                position += semicolon + 1;
                continue;
            }
            // '<'
            if (position + 1 >= end) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
            byte kind = xml[position + 1];
            if (kind == (byte)'!')
            {
                ReadOnlySpan<byte> rest = xml.Slice(position);
                if (rest.Length >= 4 && rest[2] == (byte)'-' && rest[3] == (byte)'-')
                {
                    int close = rest.Slice(4).IndexOf("-->"u8);
                    if (close < 0) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
                    position += 4 + close + 3;
                    continue;
                }
                if (rest.StartsWith("<![CDATA["u8))
                {
                    int close = rest.Slice(9).IndexOf("]]>"u8);
                    if (close < 0) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
                    if (inBlock && paragraphDepth > 0 && skipDepth < 0) FeedText(rest.Slice(9, close), matcher);
                    position += 9 + close + 3;
                    continue;
                }
                return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);     // a DOCTYPE inside the body
            }
            if (kind == (byte)'?')
            {
                int close = xml.Slice(position + 2).IndexOf("?>"u8);
                if (close < 0) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
                position += 2 + close + 2;
                continue;
            }
            if (kind == (byte)'/')
            {
                int nameStart = position + 2;
                int close = xml.Slice(nameStart).IndexOf((byte)'>');
                if (close < 0) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
                position = nameStart + close + 1;
                if (!inBlock)
                {
                    // an element around the blocks closed; the bytes never reach past office:text
                    if (_frames.Count == 0) return Fail(segmented, firstBlock, out stoppedAt);
                    _frames.RemoveAt(_frames.Count - 1);
                    continue;
                }
                depth--;
                if (depth < 0)
                {
                    if (!segmented) break;               // per-block: past the block's own end
                    return Fail(segmented, firstBlock, out stoppedAt);
                }
                if (skipDepth >= 0)
                {
                    if (depth == skipDepth) skipDepth = -1;
                }
                else
                {
                    int nameEnd = nameStart;
                    while (nameEnd < position - 1 && !IsSpace(xml[nameEnd]) && xml[nameEnd] != (byte)'>') nameEnd++;
                    ReadOnlySpan<byte> prefix = PrefixOf(xml.Slice(nameStart, nameEnd - nameStart), out ReadOnlySpan<byte> local);
                    if (local.Length == 1 && (local[0] is (byte)'p' or (byte)'h') && Resolve(prefix) == NamespaceKind.Text) { paragraphDepth--; matcher.Separator(); }
                }
                if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth);
                if (segmented && depth == 0)
                {
                    blockEnded!(block);
                    block++;
                    inBlock = false;
                }
                continue;
            }
            // a start tag
            int elementNameStart = position + 1, elementNameEnd = elementNameStart;
            while (elementNameEnd < end && !IsNameEnd(xml[elementNameEnd])) elementNameEnd++;
            int tagEnd = TagEnd(xml, elementNameEnd, out bool hasAttributes);
            if (tagEnd < 0) return Fail(segmented, inBlock ? block : firstBlock, out stoppedAt);
            bool isEmpty = tagEnd - 1 >= elementNameEnd && xml[tagEnd - 1] == (byte)'/';
            ReadOnlySpan<byte> elementName = xml.Slice(elementNameStart, elementNameEnd - elementNameStart);
            if (!inBlock)
            {
                Frame parent = _frames.Count == 0 ? Frame.Outer : _frames[_frames.Count - 1];
                Outside outside = parent == Frame.Skip ? Outside.Other : ClassifyOutside(elementName);
                Frame frame;
                bool startsBlock = false;
                switch (parent)
                {
                    case Frame.Container or Frame.ListItem:
                        switch (outside)
                        {
                            case Outside.BlockElement: startsBlock = true; frame = Frame.Skip; break;
                            case Outside.List: frame = Frame.List; break;
                            case Outside.Container: frame = Frame.Container; break;
                            case Outside.IndexElement: frame = Frame.Index; break;
                            default: frame = Frame.Skip; break;
                        }
                        break;
                    case Frame.Index: frame = outside == Outside.IndexBody ? Frame.Container : Frame.Skip; break;
                    case Frame.List: frame = outside == Outside.ListItem ? Frame.ListItem : Frame.Skip; break;
                    case Frame.Skip: frame = Frame.Skip; break;
                    default: return Fail(segmented, firstBlock, out stoppedAt);     // above office:text: not where the index's blocks are
                }
                if (!startsBlock)
                {
                    if (!isEmpty) _frames.Add(frame);
                    position = tagEnd + 1;
                    continue;
                }
                if (block == stopBlock) { stoppedAt = position; return WordSearchBlocksResult.Stopped; }
                inBlock = true;
                depth = 0;
                skipDepth = -1;
                paragraphDepth = 0;
                _lastWasSpace = false;
                _namespaces.ClearBlock();     // the reference reads the block's range alone: only the root's declarations hold
            }
            int elementDepth = depth;
            if (hasAttributes)
            {
                ReadOnlySpan<byte> attributes = xml.Slice(elementNameEnd, tagEnd - elementNameEnd);
                if (attributes.IndexOf("xmlns"u8) >= 0 && !_namespaces.ReadDeclarations(attributes, elementDepth)) return Fail(segmented, block, out stoppedAt);
            }
            if (skipDepth < 0)
            {
                ReadOnlySpan<byte> elementPrefix = PrefixOf(elementName, out ReadOnlySpan<byte> local);
                NamespaceKind namespaceKind = local.Length is 1 or 3 or 9 or 10 or 13 or 15 ? Resolve(elementPrefix) : NamespaceKind.Other;   // the lengths of the names below
                if (namespaceKind == NamespaceKind.Text)
                {
                    if (local.Length == 1 && (local[0] is (byte)'p' or (byte)'h'))
                    {
                        matcher.Separator();
                        if (!isEmpty) paragraphDepth++;
                    }
                    else if (local.Length == 1 && local[0] == (byte)'s')
                    {
                        int spaces = hasAttributes ? SpaceCount(xml.Slice(elementNameEnd, tagEnd - elementNameEnd)) : 1;
                        for (int i = 0; i < spaces; i++) matcher.Feed(' ');
                    }
                    else if (Is(local, "tab")) matcher.Feed(' ');
                    else if (Is(local, "line-break")) matcher.Separator();
                    else if (Is(local, "note-citation") || Is(local, "tracked-changes") || (!includeNotes && Is(local, "note-body"))) { if (!isEmpty) skipDepth = elementDepth; }
                }
                else if (namespaceKind == NamespaceKind.Office && Is(local, "annotation")) { if (!isEmpty) skipDepth = elementDepth; }
            }
            position = tagEnd + 1;
            if (isEmpty)
            {
                if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth);
                if (segmented && depth == 0)
                {
                    // an empty block element (<text:p/>)
                    blockEnded!(block);
                    block++;
                    inBlock = false;
                }
            }
            else depth++;
        }
        if (segmented)
        {
            if (inBlock) return Fail(segmented, block, out stoppedAt);       // the bytes end inside a block
            if (block != stopBlock) return Fail(segmented, firstBlock, out stoppedAt);
        }
        stoppedAt = end;
        return WordSearchBlocksResult.Done;
    }

    private NamespaceKind Resolve(ReadOnlySpan<byte> prefix) => (NamespaceKind)_namespaces.Resolve(prefix);

    private static WordSearchBlocksResult Fail(bool segmented, int block, out long stoppedAt)
    {
        // per block: the reference reader keeps what precedes the error - so does this scan (nothing more is fed)
        stoppedAt = segmented ? block : 0;
        return segmented ? WordSearchBlocksResult.Failed : WordSearchBlocksResult.Done;
    }

    /// <summary>text:s's text:c (ODF 1.3 part 3 6.1.3): the count when it parses as a positive integer, else 1 - as the reference's int.TryParse reads it.</summary>
    private int SpaceCount(ReadOnlySpan<byte> attributes)
    {
        int cursor = 0, end = attributes.Length;
        while (cursor < end)
        {
            byte b = attributes[cursor];
            if (IsSpace(b) || b == (byte)'/') { cursor++; continue; }
            int equals = attributes.Slice(cursor).IndexOf((byte)'=');
            if (equals < 0) return 1;
            ReadOnlySpan<byte> attributeName = attributes.Slice(cursor, equals);
            int trimmed = attributeName.Length;
            while (trimmed > 0 && IsSpace(attributeName[trimmed - 1])) trimmed--;
            attributeName = attributeName.Slice(0, trimmed);
            cursor += equals + 1;
            while (cursor < end && IsSpace(attributes[cursor])) cursor++;
            if (cursor >= end) return 1;
            byte quote = attributes[cursor];
            if (quote is not ((byte)'"' or (byte)'\'')) return 1;
            int valueEnd = attributes.Slice(cursor + 1).IndexOf(quote);
            if (valueEnd < 0) return 1;
            ReadOnlySpan<byte> value = attributes.Slice(cursor + 1, valueEnd);
            cursor += 1 + valueEnd + 1;
            ReadOnlySpan<byte> prefix = PrefixOf(attributeName, out ReadOnlySpan<byte> local);
            if (prefix.Length == 0 || local.Length != 1 || local[0] != (byte)'c' || Resolve(prefix) != NamespaceKind.Text) continue;
            int count = 0;
            bool plainDigits = value.Length is > 0 and <= 9;
            foreach (byte digit in value)
            {
                if (digit is < (byte)'0' or > (byte)'9') { plainDigits = false; break; }
                count = count * 10 + (digit - '0');
            }
            if (!plainDigits && !int.TryParse(DecodeAttribute(value), out count)) return 1;
            return count > 0 ? count : 1;
        }
        return 1;
    }

    /// <summary>Paragraph text as UTF-8: white space collapsed to one space, every other character fed as the reference's string holds it.</summary>
    private void FeedText(ReadOnlySpan<byte> text, WordDocumentSearchMatcher matcher)
    {
        int length = text.Length;
        for (int i = 0; i < length;)
        {
            int b = text[i];
            if (b < 0x80)
            {
                if (b is ' ' or '\t' or '\r' or '\n') { if (!_lastWasSpace) matcher.Feed(' '); _lastWasSpace = true; }
                else { matcher.Feed((char)b); _lastWasSpace = false; }
                i++;
                continue;
            }
            int codePoint;
            if ((b & 0xE0) == 0xC0 && i + 1 < length && (text[i + 1] & 0xC0) == 0x80)
            { codePoint = ((b & 0x1F) << 6) | (text[i + 1] & 0x3F); i += 2; }
            else if ((b & 0xF0) == 0xE0 && i + 2 < length && (text[i + 1] & 0xC0) == 0x80 && (text[i + 2] & 0xC0) == 0x80)
            { codePoint = ((b & 0x0F) << 12) | ((text[i + 1] & 0x3F) << 6) | (text[i + 2] & 0x3F); i += 3; }
            else if ((b & 0xF8) == 0xF0 && i + 3 < length && (text[i + 1] & 0xC0) == 0x80 && (text[i + 2] & 0xC0) == 0x80 && (text[i + 3] & 0xC0) == 0x80)
            { codePoint = ((b & 0x07) << 18) | ((text[i + 1] & 0x3F) << 12) | ((text[i + 2] & 0x3F) << 6) | (text[i + 3] & 0x3F); i += 4; }
            else { codePoint = 0xFFFD; i++; }
            FeedCodePoint(codePoint, matcher);
        }
    }

    private void FeedCodePoint(int codePoint, WordDocumentSearchMatcher matcher)
    {
        if (codePoint is ' ' or '\t' or '\r' or '\n') { if (!_lastWasSpace) matcher.Feed(' '); _lastWasSpace = true; return; }
        _lastWasSpace = false;
        if (codePoint < 0x10000) { matcher.Feed((char)codePoint); return; }
        codePoint -= 0x10000;
        matcher.Feed((char)(0xD800 + (codePoint >> 10)));
        matcher.Feed((char)(0xDC00 + (codePoint & 0x3FF)));
    }
}
