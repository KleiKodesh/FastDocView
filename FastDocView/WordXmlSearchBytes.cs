using System.Text;
using System.Xml;

namespace FastDocView;

/// <summary>The outcome of a byte scanner's run over consecutive blocks (.docx <c>DocxSearchTextScanner</c>, .odt <c>OdtSearchTextScanner</c>).</summary>
internal enum WordSearchBlocksResult
{
    /// <summary>The bytes ended and the block count came out at exactly the stop block.</summary>
    Done,
    /// <summary>The stop block's start was found; the stopped-at value is its offset in the bytes.</summary>
    Stopped,
    /// <summary>A block the reference reader would stop inside, or a layout that is not the index's; the stopped-at value is the failing block.</summary>
    Failed,
}

/// <summary>What a search batch does with one piece the byte scanner read.</summary>
internal enum WordSearchPieceVerdict
{
    /// <summary>The piece's hits stand.</summary>
    Accept,
    /// <summary>The hits before the failing block stand; the rest is read block by block from the index's ranges.</summary>
    FallBackFromFailedBlock,
    /// <summary>The scanner's blocks are not the index's: drop the piece's hits and read it block by block.</summary>
    RedoPiece,
}

internal static class WordSearchPiece
{
    /// <summary>
    /// Whether a piece's scan agrees with the index. <paramref name="isEndExact"/>: the bytes end exactly at the piece's
    /// end block start (an anchor or the body end), so the scan must come out Done there - Done on bytes that run on to
    /// the next anchor, or Stopped anywhere but the very end of exact bytes, means the scanner counted blocks the index
    /// does not have, and every hit of the piece could be numbered wrong.
    /// </summary>
    public static WordSearchPieceVerdict Judge(WordSearchBlocksResult result, bool isEndExact, long stoppedAt, int length) => result switch
    {
        WordSearchBlocksResult.Failed => WordSearchPieceVerdict.FallBackFromFailedBlock,
        WordSearchBlocksResult.Done when !isEndExact => WordSearchPieceVerdict.RedoPiece,
        WordSearchBlocksResult.Stopped when isEndExact && stoppedAt != length => WordSearchPieceVerdict.RedoPiece,
        _ => WordSearchPieceVerdict.Accept,
    };
}

/// <summary>Byte-level XML helpers shared by the search scanners: names, entities, attribute text.</summary>
internal static class WordXmlBytes
{
    public static bool IsNameEnd(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' or (byte)'>' or (byte)'/';

    public static bool IsSpace(byte b) => b is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n';

    public static bool Is(ReadOnlySpan<byte> name, string ascii)
    {
        if (name.Length != ascii.Length) return false;
        for (int i = 0; i < name.Length; i++) if (name[i] != ascii[i]) return false;
        return true;
    }

    public static ReadOnlySpan<byte> PrefixOf(ReadOnlySpan<byte> name, out ReadOnlySpan<byte> local)
    {
        int colon = name.IndexOf((byte)':');
        if (colon < 0) { local = name; return []; }
        local = name.Slice(colon + 1);
        return name.Slice(0, colon);
    }

    /// <summary>The offset of a start tag's closing '&gt;' (the first one outside a quoted attribute value) at or after <paramref name="from"/>; -1 when unterminated.</summary>
    public static int TagEnd(ReadOnlySpan<byte> xml, int from, out bool hasAttributes)
    {
        hasAttributes = false;
        int cursor = from;
        while (true)
        {
            int found = xml.Slice(cursor).IndexOfAny((byte)'>', (byte)'"', (byte)'\'');
            if (found < 0) return -1;
            cursor += found;
            if (xml[cursor] == (byte)'>') return cursor;
            hasAttributes = true;
            int closeQuote = xml.Slice(cursor + 1).IndexOf(xml[cursor]);
            if (closeQuote < 0) return -1;
            cursor += 1 + closeQuote + 1;
        }
    }

    /// <summary>An attribute value written with entities, decoded by a parser (rare: the callers take the plain bytes when there is no '&amp;').</summary>
    public static string DecodeAttribute(ReadOnlySpan<byte> raw)
    {
        string text = Encoding.UTF8.GetString(raw.ToArray());
        if (text.IndexOf('&') < 0) return text;
        try
        {
            using XmlReader reader = XmlReader.Create(new StringReader("<a b=\"" + text.Replace("\"", "&quot;") + "\"/>"));
            reader.MoveToContent();
            return reader.GetAttribute("b") ?? "";
        }
        catch (XmlException) { return text; }
    }

    /// <summary>The five predefined entities and decimal / hexadecimal character references (XML 1.0 4.1, 4.6).</summary>
    public static bool TryDecodeEntity(ReadOnlySpan<byte> name, out int codePoint)
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

/// <summary>
/// Element prefixes resolved by namespace, as a parser resolves them: the part root's declarations plus any declared
/// inside the block being scanned (innermost last). A namespace is reduced to a small kind the scanner chooses
/// (0 = not one it cares about).
/// </summary>
internal sealed class WordXmlNamespaceScope
{
    private readonly Func<string, byte> _kindOf;
    private readonly List<(byte[] prefix, byte kind)> _root = new();
    private readonly List<(byte[] prefix, byte kind, int depth)> _block = new();

    public WordXmlNamespaceScope(string rootOpenTag, Func<string, byte> kindOf)
    {
        _kindOf = kindOf;
        // the root tag's own declarations, read by a parser (never by hand)
        string rootName = Docx.DocxXmlFragment.RootElementName(rootOpenTag);
        using XmlReader root = XmlReader.Create(new StringReader(rootOpenTag + "</" + rootName + ">"), new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
        root.MoveToContent();
        if (root.MoveToFirstAttribute())
            do
            {
                if (root.Prefix == "xmlns") _root.Add((Encoding.UTF8.GetBytes(root.LocalName), kindOf(root.Value)));
                else if (root.Name == "xmlns") _root.Add(([], kindOf(root.Value)));
            } while (root.MoveToNextAttribute());
    }

    public bool HasBlockDeclarations => _block.Count > 0;

    public void ClearBlock() => _block.Clear();

    public byte Resolve(ReadOnlySpan<byte> prefix)
    {
        for (int i = _block.Count - 1; i >= 0; i--)
            if (prefix.SequenceEqual(_block[i].prefix)) return _block[i].kind;
        foreach ((byte[] declared, byte kind) in _root)
            if (prefix.SequenceEqual(declared)) return kind;
        return 0;
    }

    /// <summary>Drops the declarations of elements at <paramref name="depth"/> and deeper (they closed).</summary>
    public void Pop(int depth)
    {
        while (_block.Count > 0 && _block[_block.Count - 1].depth >= depth) _block.RemoveAt(_block.Count - 1);
    }

    /// <summary>The xmlns / xmlns:p attributes of one start tag's attribute text; false when it is not well formed.</summary>
    public bool ReadDeclarations(ReadOnlySpan<byte> attributes, int elementDepth)
    {
        int cursor = 0, end = attributes.Length;
        while (cursor < end)
        {
            byte b = attributes[cursor];
            if (WordXmlBytes.IsSpace(b) || b == (byte)'/') { cursor++; continue; }
            int equals = attributes.Slice(cursor).IndexOf((byte)'=');
            if (equals < 0) return false;
            ReadOnlySpan<byte> attributeName = attributes.Slice(cursor, equals);
            int trimmed = attributeName.Length;
            while (trimmed > 0 && WordXmlBytes.IsSpace(attributeName[trimmed - 1])) trimmed--;
            attributeName = attributeName.Slice(0, trimmed);
            cursor += equals + 1;
            while (cursor < end && WordXmlBytes.IsSpace(attributes[cursor])) cursor++;
            if (cursor >= end) return false;
            byte quote = attributes[cursor];
            if (quote is not ((byte)'"' or (byte)'\'')) return false;
            int valueEnd = attributes.Slice(cursor + 1).IndexOf(quote);
            if (valueEnd < 0) return false;
            ReadOnlySpan<byte> value = attributes.Slice(cursor + 1, valueEnd);
            cursor += 1 + valueEnd + 1;
            if (attributeName.Length >= 5 && attributeName.StartsWith("xmlns"u8))
            {
                if (attributeName.Length == 5) _block.Add(([], _kindOf(WordXmlBytes.DecodeAttribute(value)), elementDepth));
                else if (attributeName[5] == (byte)':') _block.Add((attributeName.Slice(6).ToArray(), _kindOf(WordXmlBytes.DecodeAttribute(value)), elementDepth));
            }
        }
        return true;
    }
}
