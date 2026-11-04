using System.Runtime.CompilerServices;
using static FastDocView.WordXmlBytes;

namespace FastDocView.Docx;

/// <summary>
/// The search text of one .docx footnote / endnote straight from its UTF-8 XML bytes (the w:footnote / w:endnote element
/// the notes index gives the range of): no XElement, no strings. It feeds the matcher exactly what
/// <see cref="WordXmlSearchText.FeedOoxml(System.Xml.Linq.XElement, WordDocumentSearchMatcher)"/> feeds it from the parsed
/// note - that path stays as the reference:
/// - per w:p: the text of its w:t (entities and character references decoded, CDATA kept; a whitespace-only text node
///   dropped unless xml:space="preserve" is in scope, as the note parser drops it), w:tab inside a run a space, w:br / w:cr
///   separators, the end of the w:p a separator; text outside paragraphs is not text;
/// - names resolved by namespace (the part root's declarations plus any declared in the note).
/// <see cref="TryFeed"/> declines (returns false, the caller rolls the matcher back and parses the note) where the
/// reference does something a forward scan cannot copy cheaply: a paragraph inside a paragraph (a text box - the
/// reference reads its text twice), w:drawing / w:pict / w:object (the parser strips them first), whitespace written only
/// as character references, and anything malformed (the parser keeps what precedes the error).
/// </summary>
internal sealed class DocxNoteSearchTextScanner
{
    private enum NamespaceKind : byte { Other, Wordprocessing, Xml }

    private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";

    private readonly WordXmlNamespaceScope _namespaces;
    private readonly List<bool> _isRun = new();              // per open element: is it a run (local name "r", any namespace)
    private readonly List<(int depth, bool preserve)> _spaces = new();   // xml:space in scope, innermost last

    public DocxNoteSearchTextScanner(string rootOpenTag)
        => _namespaces = new WordXmlNamespaceScope(rootOpenTag, uri => (byte)(uri switch
        {
            WordXmlSearchText.WordprocessingNamespace or WordXmlSearchText.WordprocessingStrictNamespace => NamespaceKind.Wordprocessing,
            _ => NamespaceKind.Other,
        }));

    private NamespaceKind Resolve(ReadOnlySpan<byte> prefix)
        => prefix.Length == 3 && Is(prefix, "xml") ? NamespaceKind.Xml : (NamespaceKind)_namespaces.Resolve(prefix);

    /// <summary>Feeds the note's text; false (with possibly part of it fed - the caller restores the matcher) when the reference must read it.</summary>
#if NET
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
#endif
    public bool TryFeed(ReadOnlySpan<byte> xml, WordDocumentSearchMatcher matcher)
    {
        _namespaces.ClearBlock();
        _isRun.Clear();
        _spaces.Clear();
        int position = 0, end = xml.Length;
        int depth = 0;
        int paragraphDepth = -1;     // depth of the open w:p
        int textDepth = -1;          // depth of the open w:t
        while (position < end)
        {
            int next = xml.Slice(position).IndexOf((byte)'<');
            int stop = next < 0 ? end : position + next;
            if (stop > position && textDepth >= 0 && !FeedTextNode(xml.Slice(position, stop - position), matcher)) return false;
            if (next < 0) break;
            position = stop;
            if (position + 1 >= end) return false;
            byte kind = xml[position + 1];
            if (kind == (byte)'!')
            {
                ReadOnlySpan<byte> rest = xml.Slice(position);
                if (rest.Length >= 4 && rest[2] == (byte)'-' && rest[3] == (byte)'-')
                {
                    int close = rest.Slice(4).IndexOf("-->"u8);
                    if (close < 0) return false;
                    position += 4 + close + 3;
                    continue;
                }
                if (rest.StartsWith("<![CDATA["u8))
                {
                    int close = rest.Slice(9).IndexOf("]]>"u8);
                    if (close < 0) return false;
                    if (textDepth >= 0) matcher.FeedUtf8(rest.Slice(9, close));
                    position += 9 + close + 3;
                    continue;
                }
                return false;
            }
            if (kind == (byte)'?')
            {
                int close = xml.Slice(position + 2).IndexOf("?>"u8);
                if (close < 0) return false;
                position += 2 + close + 2;
                continue;
            }
            if (kind == (byte)'/')
            {
                int nameStart = position + 2;
                int close = xml.Slice(nameStart).IndexOf((byte)'>');
                if (close < 0) return false;
                position = nameStart + close + 1;
                depth--;
                if (depth < 0) return false;
                _isRun.RemoveAt(_isRun.Count - 1);
                if (depth == textDepth) textDepth = -1;
                else if (depth == paragraphDepth) { paragraphDepth = -1; matcher.Separator(); }
                if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth);
                while (_spaces.Count > 0 && _spaces[_spaces.Count - 1].depth >= depth) _spaces.RemoveAt(_spaces.Count - 1);
                continue;
            }
            int nameStartIndex = position + 1, nameEnd = nameStartIndex;
            while (nameEnd < end && !IsNameEnd(xml[nameEnd])) nameEnd++;
            int tagEnd = TagEnd(xml, nameEnd, out bool hasAttributes);
            if (tagEnd < 0) return false;
            bool isEmpty = tagEnd - 1 >= nameEnd && xml[tagEnd - 1] == (byte)'/';
            int elementDepth = depth;
            if (hasAttributes)
            {
                ReadOnlySpan<byte> attributes = xml.Slice(nameEnd, tagEnd - nameEnd);
                if (attributes.IndexOf("xmlns"u8) >= 0 && !_namespaces.ReadDeclarations(attributes, elementDepth)) return false;
                if (!isEmpty && attributes.IndexOf("space"u8) >= 0 && !ReadSpace(attributes, elementDepth)) return false;
            }
            ReadOnlySpan<byte> prefix = PrefixOf(xml.Slice(nameStartIndex, nameEnd - nameStartIndex), out ReadOnlySpan<byte> local);
            bool parentIsRun = _isRun.Count > 0 && _isRun[_isRun.Count - 1];
            if (Is(local, "drawing") || Is(local, "pict") || Is(local, "object")) return false;    // the parser strips these first
            if (Is(local, "fldChar") || Is(local, "fldSimple") || Is(local, "instrText")) return false;   // fields: read through the field stack by the parsed path
            if (local.Length is 1 or 2 or 3 && Resolve(prefix) == NamespaceKind.Wordprocessing)
            {
                if (local.Length == 1 && local[0] == (byte)'p')
                {
                    if (paragraphDepth >= 0) return false;    // a paragraph inside a paragraph: the reference reads it twice
                    if (isEmpty) matcher.Separator(); else paragraphDepth = elementDepth;
                }
                else if (paragraphDepth >= 0)
                {
                    if (local.Length == 1 && local[0] == (byte)'t') { if (!isEmpty) textDepth = elementDepth; }
                    else if (Is(local, "tab")) { if (parentIsRun) matcher.Feed(' '); }
                    else if (Is(local, "br") || Is(local, "cr")) matcher.Separator();
                }
            }
            position = tagEnd + 1;
            if (isEmpty) { if (_namespaces.HasBlockDeclarations) _namespaces.Pop(depth); }
            else { depth++; _isRun.Add(local.Length == 1 && local[0] == (byte)'r'); }
        }
        return depth == 0;
    }

    /// <summary>
    /// One text node inside a w:t: whitespace-only is dropped unless xml:space="preserve" is in scope (the parser ignores
    /// whitespace); false when it is whitespace written with character references (a case left to the reference).
    /// </summary>
    private bool FeedTextNode(ReadOnlySpan<byte> text, WordDocumentSearchMatcher matcher)
    {
        bool preserve = _spaces.Count > 0 && _spaces[_spaces.Count - 1].preserve;
        int ampersand = text.IndexOf((byte)'&');
        if (!preserve)
        {
            // content = a raw byte that is not white space, outside the references
            bool hasContent = false;
            for (int i = 0; i < text.Length; i++)
            {
                byte b = text[i];
                if (b == (byte)'&') { int semicolon = text.Slice(i).IndexOf((byte)';'); if (semicolon < 0) return false; i += semicolon; continue; }
                if (!IsSpace(b)) { hasContent = true; break; }
            }
            if (!hasContent) return ampersand < 0;     // whitespace only: dropped; with references: left to the reference
        }
        if (ampersand < 0) { matcher.FeedUtf8(text); return true; }
        int position = 0;
        while (position < text.Length)
        {
            int next = text.Slice(position).IndexOf((byte)'&');
            if (next < 0) { matcher.FeedUtf8(text.Slice(position)); break; }
            if (next > 0) matcher.FeedUtf8(text.Slice(position, next));
            position += next;
            int semicolon = text.Slice(position).IndexOf((byte)';');
            if (semicolon < 0 || !TryDecodeEntity(text.Slice(position + 1, semicolon - 1), out int codePoint)) return false;
            matcher.FeedCodePoint(codePoint);
            position += semicolon + 1;
        }
        return true;
    }

    /// <summary>An xml:space attribute of a start tag: pushes its scope; false when the tag's attributes are not well formed.</summary>
    private bool ReadSpace(ReadOnlySpan<byte> attributes, int elementDepth)
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
            if (Is(attributeName, "xml:space"))
            {
                if (Is(value, "preserve")) _spaces.Add((elementDepth, true));
                else if (Is(value, "default")) _spaces.Add((elementDepth, false));
                else return false;       // the parser rejects any other value
            }
        }
        return true;
    }
}
