using System.Xml;
using System.Xml.Linq;

namespace FastDocView;

/// <summary>
/// In-document search options. <see cref="MatchCase"/> true (the default) compares letters exactly, as the app's
/// book-view search does (Hebrew has no case); false folds both sides with <see cref="char.ToLowerInvariant"/>.
/// <see cref="IncludeNotes"/>: a footnote / endnote body that holds the query makes the block that references it a hit.
/// </summary>
public sealed record WordDocumentSearchOptions(bool MatchCase = true, bool IncludeNotes = true)
{
    public static readonly WordDocumentSearchOptions Default = new();
}

/// <summary>
/// One bounded step of an in-document search. <see cref="HitBlocks"/>: the blocks of the scanned range whose text holds
/// the query, ascending. <see cref="NextBlock"/>: where the next call continues; -1 once <see cref="ReachedEnd"/>
/// (the index is complete and every block was scanned). Only block indexes are produced - no text is kept.
/// </summary>
public sealed record WordDocumentSearchBatch(int[] HitBlocks, int NextBlock, bool ReachedEnd)
{
    internal static WordDocumentSearchBatch Finish(List<int> hits, int scannedEnd, bool isIndexComplete, int blockCount)
    {
        bool reachedEnd = isIndexComplete && scannedEnd >= blockCount;
        return new WordDocumentSearchBatch(hits.ToArray(), reachedEnd ? -1 : scannedEnd, reachedEnd);
    }

    internal static readonly WordDocumentSearchBatch Nothing = new([], -1, true);
}

/// <summary>
/// The search's text comparison - the SAME normalisation as the KitveiHakodesh app's book-view search
/// (Core SeforimDbLineTextScan.NormalizeQuery / CountIn and vue-frontend utils/hebrewTextProcessing.ts
/// removeDiacriticsForSearch): Hebrew marks U+0591-U+05C7 except the maqaf are dropped; the maqaf U+05BE is a
/// space; every double-quote glyph (" U+05F4 U+201C U+201D U+201E) is '"'; every single-quote glyph
/// (' U+05F3 U+2018 U+2019 U+201A `) is '\''; the query is trimmed; letters compare exactly. Final letters are NOT folded
/// (the app does not fold them either). One KMP automaton fed character by character: no normalised copy is built.
/// <see cref="Separator"/> ends a paragraph / cell so a match never joins two of them.
/// </summary>
public sealed class WordDocumentSearchMatcher
{
    private readonly string _needle;
    private readonly int[] _fail;
    private readonly bool _foldCase;
    private int _state;

    public WordDocumentSearchMatcher(string query, WordDocumentSearchOptions? options = null)
    {
        _foldCase = !(options ?? WordDocumentSearchOptions.Default).MatchCase;
        _needle = Normalize(query ?? "", _foldCase);
        _fail = new int[_needle.Length];
        for (int i = 1, k = 0; i < _needle.Length; i++)
        {
            while (k > 0 && _needle[i] != _needle[k]) k = _fail[k - 1];
            if (_needle[i] == _needle[k]) k++;
            _fail[i] = k;
        }
    }

    public bool IsEmpty => _needle.Length == 0;

    /// <summary>True once the needle was seen since the last <see cref="Reset"/>.</summary>
    public bool Found { get; private set; }

    public void Reset() { _state = 0; Found = false; }

    /// <summary>A paragraph / cell / line boundary: the partial match is dropped.</summary>
    public void Separator() => _state = 0;

    public static bool IsIgnoredMark(char c) => c >= '֑' && c <= 'ׇ' && c != '־';

    public static char Fold(char c) => c switch
    {
        '־' => ' ',
        '"' or '״' or '“' or '”' or '„' => '"',
        '\'' or '׳' or '‘' or '’' or '‚' or '`' => '\'',
        _ => c,
    };

    private static string Normalize(string query, bool foldCase)
    {
        var text = new System.Text.StringBuilder(query.Length);
        foreach (char c in query.Trim())
        {
            if (IsIgnoredMark(c)) continue;
            char folded = Fold(c);
            text.Append(foldCase ? char.ToLowerInvariant(folded) : folded);
        }
        return text.ToString();
    }

    /// <summary>One character of document text (raw: the marks and quotes are normalised here).</summary>
    public void Feed(char c)
    {
        if (Found || IsIgnoredMark(c)) return;
        c = Fold(c);
        if (_foldCase) c = char.ToLowerInvariant(c);
        while (_state > 0 && _needle[_state] != c) _state = _fail[_state - 1];
        if (_needle[_state] == c) _state++;
        if (_state == _needle.Length) Found = true;
    }

    public void Feed(ReadOnlySpan<char> text)
    {
        foreach (char c in text)
        {
            if (Found) return;
            if (c is '\r' or '\n' or '\u0007' or '\u000B' or '\u000C') Separator(); else Feed(c);
        }
    }

    /// <summary>
    /// An HTML fragment walked exactly as the app's search walks a line's HTML (SeforimDbLineTextScan.CountIn): tags
    /// skipped, a stray '>' dropped, a quote entity becomes its quote, any other entity one sentinel character.
    /// </summary>
    public void FeedHtml(string html)
    {
        int i = 0, length = html.Length;
        while (i < length && !Found)
        {
            char ch = html[i];
            if (ch == '<')
            {
                int close = html.IndexOf('>', i + 1);
                if (close < 0) break;
                i = close + 1;
                continue;
            }
            if (ch == '>') { i++; continue; }
            if (ch == '&')
            {
                int entityEnd = -1;
                for (int j = i + 1; j < length && j <= i + 12; j++)
                {
                    char c = html[j];
                    if (c == ';') { entityEnd = j; break; }
                    if (c is ' ' or '\t' or '\n' or '<') break;
                }
                if (entityEnd >= 0)
                {
                    Feed(html.Substring(i, entityEnd - i + 1).ToLowerInvariant() switch
                    {
                        "&quot;" or "&#34;" or "&#x22;" => '"',
                        "&apos;" or "&#39;" or "&#x27;" => '\'',
                        _ => '\0',
                    });
                    i = entityEnd + 1;
                    continue;
                }
            }
            Feed(ch);
            i++;
        }
    }
}

/// <summary>
/// Plain text of one block's XML for the search, straight from its bytes (no XElement tree, no rendering): a
/// forward-only <see cref="XmlReader"/> over the block's byte range, with the part root's namespace declarations as
/// context. Paragraph ends are separators. A malformed block feeds what precedes its first error.
/// </summary>
internal sealed class WordXmlSearchText
{
    public const string WordprocessingNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    public const string WordprocessingStrictNamespace = "http://purl.oclc.org/ooxml/wordprocessingml/main";
    public const string OdfTextNamespace = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    public const string MarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    public const string OdfOfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

    private readonly XmlParserContext _context;
    private readonly XmlReaderSettings _settings;

    public WordXmlSearchText(string rootOpenTag)
    {
        var names = new NameTable();
        var namespaces = new XmlNamespaceManager(names);
        // the root tag's own declarations, read by a parser (never by hand)
        string rootName = Docx.DocxXmlFragment.RootElementName(rootOpenTag);
        using (var root = XmlReader.Create(new StringReader(rootOpenTag + "</" + rootName + ">"), new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
        {
            root.MoveToContent();
            if (root.MoveToFirstAttribute())
                do
                {
                    if (root.Prefix == "xmlns") namespaces.AddNamespace(root.LocalName, root.Value);
                    else if (root.Name == "xmlns") namespaces.AddNamespace("", root.Value);
                } while (root.MoveToNextAttribute());
        }
        _context = new XmlParserContext(names, namespaces, null, XmlSpace.None);
        _settings = new XmlReaderSettings
        {
            ConformanceLevel = ConformanceLevel.Fragment, DtdProcessing = DtdProcessing.Ignore, CheckCharacters = false,
            IgnoreComments = true, IgnoreProcessingInstructions = true, NameTable = names,
        };
    }

    private XmlReader Open(byte[] bytes, int start, int length) => XmlReader.Create(new MemoryStream(bytes, start, length, writable: false), _settings, _context);

    private static bool IsWordprocessing(string ns) => ns is WordprocessingNamespace or WordprocessingStrictNamespace;

    /// <summary>
    /// A .docx body block: the character content of w:t (w:delText, field codes w:instrText and everything else are not
    /// text Word shows); w:tab a space; w:br / w:cr and the end of each w:p separators; w:pPr / w:rPr and mc:Fallback skipped.
    /// </summary>
    public void FeedOoxml(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher)
    {
        try
        {
            using XmlReader reader = Open(bytes, start, length);
            int textDepth = -1;
            bool skipped = false;                  // Skip() already moved to the next node
            while (!matcher.Found && (skipped ? !reader.EOF : reader.Read()))
            {
                skipped = false;
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        // properties (w:pPr tab stops) and the fallback copy of mc:AlternateContent are not text
                        if (reader.LocalName == "Fallback" && reader.NamespaceURI == MarkupCompatibilityNamespace) { reader.Skip(); skipped = true; continue; }
                        if (!IsWordprocessing(reader.NamespaceURI)) break;
                        if (reader.LocalName is "pPr" or "rPr") { reader.Skip(); skipped = true; continue; }
                        if (reader.LocalName == "t" && !reader.IsEmptyElement) textDepth = reader.Depth;
                        else if (reader.LocalName == "tab") matcher.Feed(' ');
                        else if (reader.LocalName is "br" or "cr") matcher.Separator();
                        else if (reader.LocalName == "p" && reader.IsEmptyElement) matcher.Separator();
                        break;
                    case XmlNodeType.EndElement:
                        if (reader.Depth == textDepth) textDepth = -1;
                        else if (reader.LocalName == "p" && IsWordprocessing(reader.NamespaceURI)) matcher.Separator();
                        break;
                    case XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace:
                        if (textDepth >= 0) matcher.Feed(reader.Value.AsSpan());
                        break;
                }
            }
        }
        catch (XmlException) { /* keep what precedes the error */ }
    }

    /// <summary>A .docx note body (already parsed by the note reader): its w:t text, paragraphs separated.</summary>
    public static void FeedOoxml(XElement note, WordDocumentSearchMatcher matcher)
    {
        foreach (XElement paragraph in note.DescendantsAndSelf().Where(e => e.Name.LocalName == "p" && IsWordprocessing(e.Name.NamespaceName)))
        {
            foreach (XElement element in paragraph.Descendants())
            {
                if (!IsWordprocessing(element.Name.NamespaceName)) continue;
                if (element.Name.LocalName == "t") matcher.Feed(element.Value.AsSpan());
                else if (element.Name.LocalName == "tab" && element.Parent?.Name.LocalName == "r") matcher.Feed(' ');
                else if (element.Name.LocalName is "br" or "cr") matcher.Separator();
                else if (element.Name.LocalName == "p") matcher.Separator();     // a nested paragraph (text box)
                if (matcher.Found) return;
            }
            matcher.Separator();
        }
    }

    /// <summary>
    /// An .odt / .fodt body block: character data inside text:p / text:h; text:s = its text:c spaces, text:tab a space,
    /// text:line-break and paragraph ends separators. Skipped: text:note-citation (the note's label), office:annotation
    /// (comments) and, unless <paramref name="includeNotes"/>, text:note-body.
    /// </summary>
    public void FeedOdf(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher, bool includeNotes)
    {
        try
        {
            using XmlReader reader = Open(bytes, start, length);
            int paragraphDepth = 0;
            bool lastWasSpace = false;
            bool skipped = false;                  // Skip() already moved to the next node
            while (!matcher.Found && (skipped ? !reader.EOF : reader.Read()))
            {
                skipped = false;
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        if (reader.NamespaceURI == OdfTextNamespace)
                        {
                            switch (reader.LocalName)
                            {
                                case "p" or "h":
                                    matcher.Separator();
                                    if (!reader.IsEmptyElement) paragraphDepth++;
                                    break;
                                case "s":
                                    int spaces = int.TryParse(reader.GetAttribute("c", OdfTextNamespace), out int c) && c > 0 ? c : 1;
                                    for (int i = 0; i < spaces; i++) matcher.Feed(' ');
                                    break;
                                case "tab": matcher.Feed(' '); break;
                                case "line-break": matcher.Separator(); break;
                                case "note-citation" or "tracked-changes": reader.Skip(); skipped = true; continue;
                                case "note-body" when !includeNotes: reader.Skip(); skipped = true; continue;
                            }
                        }
                        else if (reader.NamespaceURI == OdfOfficeNamespace && reader.LocalName == "annotation") { reader.Skip(); skipped = true; continue; }
                        break;
                    case XmlNodeType.EndElement:
                        if (reader.NamespaceURI == OdfTextNamespace && reader.LocalName is "p" or "h") { paragraphDepth--; matcher.Separator(); }
                        break;
                    case XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace:
                        // ODF 1.3 part 3 6.1.2: white space in paragraph text (space, tab, CR, LF) collapses to one space
                        if (paragraphDepth > 0)
                            foreach (char c in reader.Value)
                            {
                                if (c is ' ' or '\t' or '\r' or '\n') { if (!lastWasSpace) matcher.Feed(' '); lastWasSpace = true; }
                                else { matcher.Feed(c); lastWasSpace = false; }
                            }
                        break;
                    case XmlNodeType.Whitespace:
                        if (paragraphDepth > 0 && !lastWasSpace) { matcher.Feed(' '); lastWasSpace = true; }
                        break;
                }
            }
        }
        catch (XmlException) { /* keep what precedes the error */ }
    }
}
