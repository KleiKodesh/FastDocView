using System.Runtime.CompilerServices;
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
/// the query, ascending; <see cref="HitCounts"/>, parallel to it: how many times the query occurs in each (the block's
/// text first, then its notes when they are searched - the app's search steps match by match, as the book view does per
/// line). Occurrences never overlap (leftmost first, as the app's highlighter marks them). <see cref="NextBlock"/>:
/// where the next call continues; -1 once <see cref="ReachedEnd"/> (the index is complete and every block was scanned).
/// Only numbers are produced - no text is kept.
/// </summary>
public sealed record WordDocumentSearchBatch(int[] HitBlocks, int[] HitCounts, int NextBlock, bool ReachedEnd)
{
    internal static WordDocumentSearchBatch Finish(List<int> hits, List<int> counts, int scannedEnd, bool isIndexComplete, int blockCount)
    {
        bool reachedEnd = isIndexComplete && scannedEnd >= blockCount;
        return new WordDocumentSearchBatch(hits.ToArray(), counts.ToArray(), reachedEnd ? -1 : scannedEnd, reachedEnd);
    }

    internal static readonly WordDocumentSearchBatch Nothing = new([], [], -1, true);
}

/// <summary>
/// The search's text comparison - the SAME normalisation as the KitveiHakodesh app's book-view search
/// (Core SeforimDbLineTextScan.NormalizeQuery / CountIn and vue-frontend utils/hebrewTextProcessing.ts
/// removeDiacriticsForSearch): Hebrew marks U+0591-U+05C7 except the maqaf are dropped; the maqaf U+05BE is a
/// space; every double-quote glyph (" U+05F4 U+201C U+201D U+201E) is '"'; every single-quote glyph
/// (' U+05F3 U+2018 U+2019 U+201A `) is '\''; the query is trimmed; letters compare exactly. Final letters are NOT folded
/// (the app does not fold them either). One KMP automaton fed character by character: no normalised copy is built.
/// <see cref="Separator"/> ends a paragraph / cell so a match never joins two of them. Every occurrence is counted
/// (<see cref="Count"/>); occurrences never overlap - after a match the automaton starts over, which is the leftmost
/// non-overlapping set the app's highlighter marks.
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

    /// <summary>How many (non-overlapping) occurrences of the needle were seen since the last <see cref="Reset"/>.</summary>
    public int Count { get; private set; }

    /// <summary>True once the needle was seen since the last <see cref="Reset"/>.</summary>
    public bool Found => Count > 0;

    public void Reset() { _state = 0; Count = 0; }

    /// <summary>The partial match and count, to undo a feed that turned out to need another reader (<see cref="RestoreState"/>).</summary>
    internal (int state, int count) SaveState() => (_state, Count);

    internal void RestoreState((int state, int count) saved) { _state = saved.state; Count = saved.count; }

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
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void Feed(char c)
    {
        if (IsIgnoredMark(c)) return;
        c = Fold(c);
        if (_foldCase) c = char.ToLowerInvariant(c);
        while (_state > 0 && _needle[_state] != c) _state = _fail[_state - 1];
        if (_needle[_state] == c) _state++;
        if (_state == _needle.Length) { Count++; _state = 0; }
    }

    public void Feed(ReadOnlySpan<char> text)
    {
        foreach (char c in text) FeedTextCharacter(c);
    }

    /// <summary>One character of a text run: a line / paragraph control character is a separator, anything else <see cref="Feed(char)"/>.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public void FeedTextCharacter(char c)
    {
        if (c is '\r' or '\n' or '\u0007' or '\u000B' or '\u000C') Separator(); else Feed(c);
    }

    /// <summary>
    /// Already-decoded text as UTF-8 bytes (markup and entities resolved by the caller), decoded here character by
    /// character - the same characters <see cref="Feed(ReadOnlySpan{char})"/> gets from the decoded string, with no
    /// string built. A malformed sequence feeds U+FFFD per bad byte.
    /// </summary>
#if NET
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]   // optimised from the first call: a search runs in short bursts
#endif
    public void FeedUtf8(ReadOnlySpan<byte> text)
    {
        int length = text.Length;
        for (int i = 0; i < length;)
        {
            int b = text[i];
            if (b < 0x80) { FeedTextCharacter((char)b); i++; continue; }
            int codePoint;
            if ((b & 0xE0) == 0xC0 && i + 1 < length && (text[i + 1] & 0xC0) == 0x80)
            { codePoint = ((b & 0x1F) << 6) | (text[i + 1] & 0x3F); i += 2; }
            else if ((b & 0xF0) == 0xE0 && i + 2 < length && (text[i + 1] & 0xC0) == 0x80 && (text[i + 2] & 0xC0) == 0x80)
            { codePoint = ((b & 0x0F) << 12) | ((text[i + 1] & 0x3F) << 6) | (text[i + 2] & 0x3F); i += 3; }
            else if ((b & 0xF8) == 0xF0 && i + 3 < length && (text[i + 1] & 0xC0) == 0x80 && (text[i + 2] & 0xC0) == 0x80 && (text[i + 3] & 0xC0) == 0x80)
            { codePoint = ((b & 0x07) << 18) | ((text[i + 1] & 0x3F) << 12) | ((text[i + 2] & 0x3F) << 6) | (text[i + 3] & 0x3F); i += 4; }
            else { codePoint = 0xFFFD; i++; }
            FeedCodePoint(codePoint);
        }
    }

    /// <summary>One Unicode scalar: a supplementary one is fed as its two UTF-16 halves, as a string holds it.</summary>
    public void FeedCodePoint(int codePoint)
    {
        if (codePoint < 0x10000) { FeedTextCharacter((char)codePoint); return; }
        codePoint -= 0x10000;
        Feed((char)(0xD800 + (codePoint >> 10)));
        Feed((char)(0xDC00 + (codePoint & 0x3FF)));
    }

    /// <summary>
    /// An HTML fragment walked exactly as the app's search walks a line's HTML (SeforimDbLineTextScan.CountIn): tags
    /// skipped, a stray '>' dropped, a quote entity becomes its quote, any other entity one sentinel character.
    /// </summary>
    public void FeedHtml(string html)
    {
        int i = 0, length = html.Length;
        while (i < length)
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
    public const string DiagramNamespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    public const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    public const string ChartExNamespace = "http://schemas.microsoft.com/office/drawing/2014/chartex";
    public const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    public const string RelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    public const string OdfOfficeNamespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    public const string OdfDrawingNamespace = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    public const string OdfTableNamespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";

    /// <summary>
    /// Where an .odt element sits, as far as the renderer (<c>OdtBlockHtmlRenderer</c>) shows its text: paragraph content,
    /// a draw:frame's children, a draw:text-box's children, or a note body (blocks again).
    /// </summary>
    internal enum OdfTextContext : byte { Paragraph, Frame, TextBox, NoteBody, Shape, ShapeGroup, BoxContent }

    /// <summary>
    /// Whether the renderer shows nothing of an element that opens in <paramref name="context"/> (it is skipped whole), and
    /// the context it opens (null: none). In a paragraph: only the text namespace and draw:a / draw:frame render
    /// (svg:title / svg:desc, shapes, images, controls and every other namespace render nothing), and text:ruby-text /
    /// text:hidden-text / text:hidden-paragraph are hidden; of a frame only its draw:text-box, of a text box only its
    /// text:p / text:h; of a drawn shape (draw:custom-shape / draw:rect / draw:ellipse / draw:line) only its text:p / text:h;
    /// of a draw:g only the shapes the renderer draws. Outside a paragraph (a note body, between
    /// blocks) nothing extra is skipped.
    /// </summary>
    /// <summary>An element's local name as bytes for <see cref="IsHiddenOdf"/>; the rules' names are short ASCII, so a longer or non-ASCII name is none of them.</summary>
    private static ReadOnlySpan<byte> AsciiName(string name, byte[] buffer)
    {
        if (name.Length > buffer.Length) return "?"u8;
        for (int i = 0; i < name.Length; i++)
        {
            if (name[i] > 127) return "?"u8;
            buffer[i] = (byte)name[i];
        }
        return buffer.AsSpan(0, name.Length);
    }

    internal static bool IsHiddenOdf(OdfTextContext? context, bool isText, bool isDrawing, ReadOnlySpan<byte> localName, out OdfTextContext? opens, bool isTable = false)
    {
        opens = null;
        if (isText && (WordXmlBytes.Is(localName, "p") || WordXmlBytes.Is(localName, "h"))) { opens = OdfTextContext.Paragraph; return context is OdfTextContext.Frame or OdfTextContext.ShapeGroup; }
        if (isText && WordXmlBytes.Is(localName, "note-body")) opens = OdfTextContext.NoteBody;
        // a block that is a drawing (a page-anchored frame or shape between paragraphs) is read like one inside a paragraph
        if (context is null && isDrawing) context = OdfTextContext.Paragraph;
        switch (context)
        {
            case OdfTextContext.Paragraph:
                if (isText) return WordXmlBytes.Is(localName, "ruby-text") || WordXmlBytes.Is(localName, "hidden-text") || WordXmlBytes.Is(localName, "hidden-paragraph");
                if (isDrawing && WordXmlBytes.Is(localName, "a")) return false;
                if (isDrawing && WordXmlBytes.Is(localName, "frame")) { opens = OdfTextContext.Frame; return false; }
                return !IsDrawnShapeOdf(isDrawing, localName, out opens);
            case OdfTextContext.ShapeGroup:
                return !IsDrawnShapeOdf(isDrawing, localName, out opens);
            case OdfTextContext.Shape:
                // text:p / text:h returned above; a shape's lists show (ODF 1.3 part 3 10.3: shapes hold text:p and text:list);
                // draw:enhanced-geometry, svg:title / svg:desc, glue points ... hidden
                if (isText && WordXmlBytes.Is(localName, "list")) { opens = OdfTextContext.BoxContent; return false; }
                return true;
            case OdfTextContext.Frame:
                if (isDrawing && WordXmlBytes.Is(localName, "text-box")) { opens = OdfTextContext.TextBox; return false; }
                return true;
            case OdfTextContext.TextBox:
                // text:p / text:h returned above; lists, tables and sections in a text box are drawn (OdtBlockHtmlRenderer.RenderBoxBlock)
                if ((isText && (WordXmlBytes.Is(localName, "list") || WordXmlBytes.Is(localName, "section"))) || (isTable && WordXmlBytes.Is(localName, "table")))
                { opens = OdfTextContext.BoxContent; return false; }
                return true;
            default:
                return false;
        }
    }

    /// <summary>A shape <see cref="Odt.OdtBlockHtmlRenderer"/> draws, and the context its children are read in (null: none of its text shows).</summary>
    private static bool IsDrawnShapeOdf(bool isDrawing, ReadOnlySpan<byte> localName, out OdfTextContext? opens)
    {
        opens = null;
        if (!isDrawing) return false;
        if (WordXmlBytes.Is(localName, "custom-shape") || WordXmlBytes.Is(localName, "rect") || WordXmlBytes.Is(localName, "ellipse")
            || WordXmlBytes.Is(localName, "circle") || WordXmlBytes.Is(localName, "line") || WordXmlBytes.Is(localName, "polygon")
            || WordXmlBytes.Is(localName, "polyline") || WordXmlBytes.Is(localName, "path") || WordXmlBytes.Is(localName, "connector")
            || WordXmlBytes.Is(localName, "measure") || WordXmlBytes.Is(localName, "caption") || WordXmlBytes.Is(localName, "regular-polygon"))
        { opens = OdfTextContext.Shape; return true; }
        if (WordXmlBytes.Is(localName, "g")) { opens = OdfTextContext.ShapeGroup; return true; }
        return false;
    }

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
    public void FeedOoxml(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher, List<string>? drawingRelationshipIds = null,
                          Docx.DocxSearchFields? fields = null)
    {
        try
        {
            using XmlReader reader = Open(bytes, start, length);
            int textDepth = -1;
            int instructionDepth = -1;             // the open w:instrText
            var simpleFields = new List<(int depth, bool separated)>();   // open w:fldSimple elements
            bool skipped = false;                  // Skip() already moved to the next node
            while (skipped ? !reader.EOF : reader.Read())
            {
                skipped = false;
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        // properties (w:pPr tab stops) and the fallback copy of mc:AlternateContent are not text
                        if (reader.LocalName == "Fallback" && reader.NamespaceURI == MarkupCompatibilityNamespace) { reader.Skip(); skipped = true; continue; }
                        // VML WordArt: the v:textpath string is the text it shows
                        if (reader.LocalName == "textpath" && reader.NamespaceURI == VmlNamespace && reader.GetAttribute("string") is { Length: > 0 } wordArt) { matcher.Separator(); matcher.Feed(wordArt.AsSpan()); matcher.Separator(); }
                        // a SmartArt (dgm:relIds r:dm) or chart (c:chart r:id) drawing: its part's text is fed after the block
                        if (drawingRelationshipIds is not null && reader.LocalName == "relIds" && reader.NamespaceURI == DiagramNamespace && reader.GetAttribute("dm", RelationshipsNamespace) is { } diagram) drawingRelationshipIds.Add(diagram);
                        if (drawingRelationshipIds is not null && reader.LocalName == "chart" && reader.NamespaceURI is ChartNamespace or ChartExNamespace && reader.GetAttribute("id", RelationshipsNamespace) is { } chart) drawingRelationshipIds.Add(chart);
                        if (!IsWordprocessing(reader.NamespaceURI)) break;
                        // properties, the source of a tracked move (Word's final view hides it) and a ruby's guide text (shown above
                        // its base, not searched - as text:ruby-text in .odt) are not text
                        if (reader.LocalName is "pPr" or "rPr" or "moveFrom" or "rt") { reader.Skip(); skipped = true; continue; }
                        if (fields is not null)
                        {
                            // a w:fldSimple's result starts at its first run-level child (the renderer separates before rendering any)
                            if (simpleFields.Count > 0 && !simpleFields[simpleFields.Count - 1].separated && reader.Depth == simpleFields[simpleFields.Count - 1].depth + 1
                                && reader.LocalName is "r" or "hyperlink" or "fldSimple" or "smartTag" or "ins" or "sdt" or "customXml")
                            {
                                simpleFields[simpleFields.Count - 1] = (simpleFields[simpleFields.Count - 1].depth, true);
                                fields.Separate(matcher);
                            }
                            switch (reader.LocalName)
                            {
                                case "fldChar":
                                {
                                    string? type = reader.GetAttribute("fldCharType", reader.NamespaceURI);
                                    bool dirty = IsOn(reader.GetAttribute("dirty", reader.NamespaceURI)), locked = IsOn(reader.GetAttribute("fldLock", reader.NamespaceURI));
                                    Fields.WordFormFieldState? form = null;
                                    if (!reader.IsEmptyElement)
                                    {
                                        form = DocxHtml.DocxBlockHtmlRenderer.FormFieldState((XElement)XNode.ReadFrom(reader));
                                        skipped = true;
                                    }
                                    if (type == "begin") fields.Begin(dirty, locked, form);
                                    else if (type == "separate") fields.Separate(matcher);
                                    else if (type == "end") fields.End(matcher);
                                    continue;
                                }
                                case "instrText":
                                    if (!reader.IsEmptyElement) instructionDepth = reader.Depth;
                                    continue;
                                case "fldSimple":
                                    fields.Begin(IsOn(reader.GetAttribute("dirty", reader.NamespaceURI)), IsOn(reader.GetAttribute("fldLock", reader.NamespaceURI)), null);
                                    fields.Instruction(reader.GetAttribute("instr", reader.NamespaceURI) ?? "");
                                    if (reader.IsEmptyElement) fields.End(matcher);
                                    else simpleFields.Add((reader.Depth, false));
                                    continue;
                                case "txbxContent":
                                    if (!reader.IsEmptyElement) fields.EnterTextBox(reader.Depth);
                                    continue;
                            }
                        }
                        if (reader.LocalName == "t" && !reader.IsEmptyElement) textDepth = reader.Depth;
                        else if (reader.LocalName == "tab") { if (fields is null || fields.Shows("\t")) matcher.Feed(' '); }
                        else if (reader.LocalName is "br" or "cr") { if (fields is null || fields.Shows(null)) matcher.Separator(); }
                        else if (reader.LocalName == "p" && reader.IsEmptyElement) matcher.Separator();
                        break;
                    case XmlNodeType.EndElement:
                        if (reader.Depth == textDepth) textDepth = -1;
                        else if (reader.Depth == instructionDepth) instructionDepth = -1;
                        else if (reader.LocalName == "p" && IsWordprocessing(reader.NamespaceURI)) matcher.Separator();
                        else if (fields is not null && simpleFields.Count > 0 && reader.Depth == simpleFields[simpleFields.Count - 1].depth)
                        {
                            simpleFields.RemoveAt(simpleFields.Count - 1);
                            fields.End(matcher);
                        }
                        fields?.Closed(reader.Depth);
                        break;
                    case XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace:
                        if (instructionDepth >= 0) fields?.Instruction(reader.Value);
                        else if (textDepth >= 0 && (fields is null || fields.IsIdle || fields.Shows(reader.Value))) matcher.Feed(reader.Value.AsSpan());
                        break;
                }
            }
        }
        catch (XmlException) { /* keep what precedes the error */ }
    }

    private static bool IsOn(string? value) => value is "1" or "true" or "on";

    /// <summary>A .docx note body (already parsed by the note reader): its w:t text, paragraphs separated.</summary>
    public static void FeedOoxml(XElement note, WordDocumentSearchMatcher matcher, Fields.WordFieldContext? fieldContext = null)
    {
        if (note.Descendants().Any(e => e.Name.LocalName is "fldChar" or "fldSimple" or "instrText"))
        {
            // fields: one pass in document order through the renderer's field stack (DocxBlockHtmlRenderer.RenderNote: fresh, no counters)
            FeedFieldNodes(note, matcher, new Docx.DocxSearchFields(fieldContext, counts: false), new List<(XElement, bool)>());
            return;
        }
        foreach (XElement paragraph in note.DescendantsAndSelf().Where(e => e.Name.LocalName == "p" && IsWordprocessing(e.Name.NamespaceName)))
        {
            foreach (XElement element in paragraph.Descendants())
            {
                if (element.Name.NamespaceName == Html.WordOfficeMathHtml.MathNamespace || element.Ancestors().TakeWhile(a => a != paragraph).Any(Html.WordOfficeMathHtml.IsMathZone))
                {
                    // an equation: its text as drawn, once (at its outermost zone)
                    if (Html.WordOfficeMathHtml.IsMathZone(element) && !element.Ancestors().TakeWhile(a => a != paragraph).Any(Html.WordOfficeMathHtml.IsMathZone)
                        && !element.Ancestors().TakeWhile(a => a != paragraph).Any(a => IsWordprocessing(a.Name.NamespaceName) && a.Name.LocalName is "moveFrom" or "del" or "rt"))
                        matcher.Feed(Html.WordOfficeMathHtml.SearchText(element).AsSpan());
                    continue;
                }
                if (!IsWordprocessing(element.Name.NamespaceName)) continue;
                if (element.Ancestors().TakeWhile(a => a != paragraph).Any(a => IsWordprocessing(a.Name.NamespaceName) && a.Name.LocalName is "moveFrom" or "rt" or "del")) continue;   // not shown
                if (element.Name.LocalName == "t") matcher.Feed(element.Value.AsSpan());
                else if (element.Name.LocalName == "tab" && element.Parent?.Name.LocalName == "r") matcher.Feed(' ');
                else if (element.Name.LocalName is "br" or "cr") matcher.Separator();
                else if (element.Name.LocalName == "p") matcher.Separator();     // a nested paragraph (text box)
            }
            matcher.Separator();
        }
    }

    /// <summary>A note's elements in document order with fields routed (the same rules as the byte reader's).</summary>
    private static void FeedFieldNodes(XElement element, WordDocumentSearchMatcher matcher, Docx.DocxSearchFields fields, List<(XElement simple, bool separated)> simpleFields)
    {
        foreach (XElement child in element.Elements())
        {
            if (!IsWordprocessing(child.Name.NamespaceName))
            {
                if (Html.WordOfficeMathHtml.IsMathZone(child))
                {
                    string equation = Html.WordOfficeMathHtml.SearchText(child);     // an equation's text as drawn
                    if (fields.IsIdle || fields.Shows(equation)) matcher.Feed(equation.AsSpan());
                    continue;
                }
                if (child.Name.LocalName != "Fallback") FeedFieldNodes(child, matcher, fields, simpleFields);
                continue;
            }
            string name = child.Name.LocalName;
            if (name is "pPr" or "rPr" or "del" or "moveFrom" or "rt") continue;
            if (simpleFields.Count > 0 && !simpleFields[simpleFields.Count - 1].separated && child.Parent == simpleFields[simpleFields.Count - 1].simple
                && name is "r" or "hyperlink" or "fldSimple" or "smartTag" or "ins" or "sdt" or "customXml")
            {
                simpleFields[simpleFields.Count - 1] = (simpleFields[simpleFields.Count - 1].simple, true);
                fields.Separate(matcher);
            }
            XNamespace w = child.Name.Namespace;
            switch (name)
            {
                case "fldChar":
                {
                    string? type = (string?)child.Attribute(w + "fldCharType");
                    if (type == "begin") fields.Begin(IsOn((string?)child.Attribute(w + "dirty")), IsOn((string?)child.Attribute(w + "fldLock")), DocxHtml.DocxBlockHtmlRenderer.FormFieldState(child));
                    else if (type == "separate") fields.Separate(matcher);
                    else if (type == "end") fields.End(matcher);
                    continue;
                }
                case "instrText": fields.Instruction(child.Value); continue;
                case "fldSimple":
                    fields.Begin(IsOn((string?)child.Attribute(w + "dirty")), IsOn((string?)child.Attribute(w + "fldLock")), null);
                    fields.Instruction((string?)child.Attribute(w + "instr") ?? "");
                    simpleFields.Add((child, false));
                    FeedFieldNodes(child, matcher, fields, simpleFields);
                    simpleFields.RemoveAt(simpleFields.Count - 1);
                    fields.End(matcher);
                    continue;
                case "txbxContent":
                    fields.EnterTextBox(0);
                    FeedFieldNodes(child, matcher, fields, simpleFields);
                    fields.Closed(0);
                    continue;
                case "t": if (fields.IsIdle || fields.Shows(child.Value)) matcher.Feed(child.Value.AsSpan()); continue;
                case "tab": if (child.Parent?.Name.LocalName == "r" && fields.Shows("\t")) matcher.Feed(' '); continue;
                case "br" or "cr": if (fields.Shows(null)) matcher.Separator(); continue;
            }
            FeedFieldNodes(child, matcher, fields, simpleFields);
            if (name == "p") matcher.Separator();
        }
    }

    /// <summary>
    /// An .odt / .fodt body block: character data inside text:p / text:h; text:s = its text:c spaces, text:tab a space,
    /// text:line-break and paragraph ends separators. Skipped: text:note-citation (the note's label), office:annotation
    /// (comments), unless <paramref name="includeNotes"/> text:note-body, and what the renderer does not show
    /// (<see cref="IsHiddenOdf"/>: svg:title / svg:desc, shapes, images, other namespaces inside a paragraph...).
    /// </summary>
    public void FeedOdf(byte[] bytes, int start, int length, WordDocumentSearchMatcher matcher, bool includeNotes)
    {
        try
        {
            using XmlReader reader = Open(bytes, start, length);
            int paragraphDepth = 0;
            bool lastWasSpace = false;
            bool skipped = false;                  // Skip() already moved to the next node
            var contexts = new List<(int depth, OdfTextContext context)>();   // what the renderer shows inside the open elements
            byte[] localBytes = new byte[32];
            while (skipped ? !reader.EOF : reader.Read())
            {
                skipped = false;
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                    {
                        OdfTextContext? context = contexts.Count == 0 ? null : contexts[contexts.Count - 1].context;
                        ReadOnlySpan<byte> local = AsciiName(reader.LocalName, localBytes);
                        if (IsHiddenOdf(context, reader.NamespaceURI == OdfTextNamespace, reader.NamespaceURI == OdfDrawingNamespace, local, out OdfTextContext? opens, reader.NamespaceURI == OdfTableNamespace))
                        { reader.Skip(); skipped = true; continue; }
                        if (opens is { } opened && !reader.IsEmptyElement) contexts.Add((reader.Depth, opened));
                    }
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
                                case "note-body" when !includeNotes:
                                    if (contexts.Count > 0 && contexts[contexts.Count - 1].depth == reader.Depth) contexts.RemoveAt(contexts.Count - 1);   // never closes: skipped whole
                                    reader.Skip(); skipped = true; continue;
                            }
                        }
                        else if (reader.NamespaceURI == OdfOfficeNamespace && reader.LocalName == "annotation") { reader.Skip(); skipped = true; continue; }
                        break;
                    case XmlNodeType.EndElement:
                        if (contexts.Count > 0 && contexts[contexts.Count - 1].depth == reader.Depth) contexts.RemoveAt(contexts.Count - 1);
                        if (reader.NamespaceURI == OdfTextNamespace && reader.LocalName is "p" or "h") { paragraphDepth--; matcher.Separator(); }
                        break;
                    case XmlNodeType.Text or XmlNodeType.CDATA or XmlNodeType.SignificantWhitespace:
                        // ODF 1.3 part 3 6.1.2: white space in paragraph text (space, tab, CR, LF) collapses to one space;
                        // only paragraph content is text (not white space between a frame's / text box's / note body's children)
                        if (paragraphDepth > 0 && contexts.Count > 0 && contexts[contexts.Count - 1].context == OdfTextContext.Paragraph)
                            foreach (char c in reader.Value)
                            {
                                if (c is ' ' or '\t' or '\r' or '\n') { if (!lastWasSpace) matcher.Feed(' '); lastWasSpace = true; }
                                else { matcher.Feed(c); lastWasSpace = false; }
                            }
                        break;
                    case XmlNodeType.Whitespace:
                        if (paragraphDepth > 0 && contexts.Count > 0 && contexts[contexts.Count - 1].context == OdfTextContext.Paragraph && !lastWasSpace) { matcher.Feed(' '); lastWasSpace = true; }
                        break;
                }
            }
        }
        catch (XmlException) { /* keep what precedes the error */ }
    }
}
