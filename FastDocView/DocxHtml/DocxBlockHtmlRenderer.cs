using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Docx;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// Renders one body-level block element (w:p, w:tbl, w:sdt, …) or a footnote/endnote body to HTML
/// with resolved formatting. No page layout: paragraphs flow, page and section breaks become thin
/// rules, tabs are real tab characters sized by the paragraph's first tab stop. Note references
/// show the number counted at index time and carry data-note-kind / data-note-number so the host
/// can pair them with the notes delivered alongside the block. Images are placeholders.
/// </summary>
public sealed class DocxBlockHtmlRenderer
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    /// <summary>(kind, note id in footnotes.xml/endnotes.xml) → its number and label, or null when unknown.</summary>
    public delegate (int number, string label)? NoteLabelResolver(WordDocumentNoteKind kind, int noteId);

    private readonly DocxStyleCascade _styles;
    private readonly NoteLabelResolver? _noteLabels;
    private DocxListLabel? _listLabel;    // the index-time label of the block being rendered (its first paragraph)
    private string? _previousStyleKey;    // the style of the paragraph rendered just before (contextual spacing)
    private string? _currentNoteLabel;    // inside a note body: the label its w:footnoteRef/w:endnoteRef shows
    private WordDocumentNoteKind _currentNoteKind;
    private string? _tableStyleId;        // inside a table: its style, layered under paragraph styles
    private DocxTableConditions _tableConditions;   // inside a table cell: the table style's conditional parts that apply to it
    private int _textBoxDepth;

    /// <summary>settings.xml compatibility option w:ulTrailSpace: underline the spaces that end a paragraph too.</summary>
    private bool UnderlinesTrailingSpaces => _styles.Compatibility.UnderlinesTrailingSpaces;            // inside a text box paragraphs become block spans (a <p> may not hold a <p>)

    private readonly IReadOnlyDictionary<string, string>? _hyperlinkTargets;   // the main part's r:id -> URL
    private readonly IReadOnlyDictionary<string, string>? _footnoteHyperlinkTargets;   // footnotes.xml.rels (a note's r:ids point there)
    private readonly IReadOnlyDictionary<string, string>? _endnoteHyperlinkTargets;    // endnotes.xml.rels
    private IReadOnlyDictionary<string, string>? _currentHyperlinkTargets;      // the relationships of the part being rendered
    private string? _activeLinkOpenTag;                                         // inside a w:hyperlink or a HYPERLINK field result: its <a> tag
    private readonly List<DocxField> _openFields = new();                       // complex fields begun in this block (w:fldChar begin … end)

    private sealed class DocxField
    {
        public readonly StringBuilder Instruction = new();
        public bool IsSeparated;
        public string? LinkOpenTag;
    }

    public DocxBlockHtmlRenderer(DocxStyleCascade styles, NoteLabelResolver? noteLabels = null, IReadOnlyDictionary<string, string>? hyperlinkTargets = null,
                                 IReadOnlyDictionary<string, string>? footnoteHyperlinkTargets = null, IReadOnlyDictionary<string, string>? endnoteHyperlinkTargets = null)
    {
        _styles = styles;
        _noteLabels = noteLabels;
        _hyperlinkTargets = hyperlinkTargets;
        _footnoteHyperlinkTargets = footnoteHyperlinkTargets;
        _endnoteHyperlinkTargets = endnoteHyperlinkTargets;
    }

    /// <param name="listLabel">The label counted at index time ("3.", "ב."), or null when the block is not a list item.</param>
    /// <param name="previousStyleKey">The paragraph style of the block before (for contextual spacing), from the index.</param>
    public string Render(XElement block, DocxListLabel? listLabel = null, string? previousStyleKey = null) => Render(block, listLabel, previousStyleKey, default);

    /// <param name="neighbours">The paragraphs around a body block (read with the window): border groups and deleted paragraph marks reach across blocks.</param>
    public string Render(XElement block, DocxListLabel? listLabel, string? previousStyleKey, DocxBlockNeighbours neighbours)
    {
        var html = new StringBuilder(1024);
        if (block.Name == W + "p" && neighbours.Next is not null && HasDeletedParagraphMark(block))
            return "";                                           // its text is drawn at the start of the next block (MergedBefore there)
        _blockNeighbours = neighbours;
        _listLabel = listLabel;
        _previousStyleKey = previousStyleKey;
        _currentHyperlinkTargets = _hyperlinkTargets;
        _openFields.Clear();
        // Fields begun in an earlier block (recorded at index time): a HYPERLINK result that runs on into this block stays a link.
        if (neighbours.OpenFields is { } openFields)
            foreach (string open in openFields.Split('\u0001'))
            {
                var field = new DocxField { IsSeparated = open.Length > 0 && open[0] == 'S' };
                field.Instruction.Append(open, Math.Min(1, open.Length), Math.Max(0, open.Length - 1));
                if (field.IsSeparated) field.LinkOpenTag = HyperlinkFieldTag(field.Instruction.ToString());
                _openFields.Add(field);
            }
        Run(RenderBlock(block, html));
        _blockNeighbours = default;
        _activeLinkOpenTag = null;
        _listLabel = null;
        _paragraphEndOf = null;                                    // keep no element of the rendered block alive
        _paragraphEndTexts.Clear();
        return html.ToString();
    }

    /// <summary>A footnote/endnote body (the w:footnote / w:endnote element); its w:footnoteRef shows <paramref name="label"/>.</summary>
    public string RenderNote(XElement note, WordDocumentNoteKind kind, string label)
    {
        var html = new StringBuilder(512);
        _currentNoteLabel = label;
        _currentNoteKind = kind;
        _previousStyleKey = null;
        // a note's r:ids point into its own part's relationships (footnotes.xml.rels / endnotes.xml.rels)
        _currentHyperlinkTargets = kind == WordDocumentNoteKind.Footnote ? _footnoteHyperlinkTargets : _endnoteHyperlinkTargets;
        _openFields.Clear();
        Run(RenderBlocks(note.Elements(), html));
        _activeLinkOpenTag = null;
        _currentNoteLabel = null;
        return html.ToString();
    }

    /// <summary>Plain text of a block.</summary>
    public static string PlainText(XElement block)
    {
        var text = new StringBuilder();
        foreach (XElement element in block.Descendants())
        {
            if (element.Name == W + "t") text.Append(element.Value);
            else if (element.Name == W + "tab") text.Append(' ');
        }
        return text.ToString().Trim();
    }

    // ── the walk ─────────────────────────────────────────────────────────────────────
    // Block content nests without bound (tables in cells, text boxes in runs, tables in text boxes…), so
    // the walk keeps its own stack instead of the call stack: each render step is an iterator that
    // writes its HTML in document order and yields a nested step where the nested content goes; Run
    // finishes that step before resuming the outer one. Any nesting depth renders in order.

    private readonly record struct RenderStep(IEnumerable<RenderStep> Steps);

    private static void Run(IEnumerable<RenderStep> root)
    {
        var stack = new Stack<IEnumerator<RenderStep>>();
        stack.Push(root.GetEnumerator());
        while (stack.Count > 0)
        {
            IEnumerator<RenderStep> top = stack.Peek();
            if (top.MoveNext()) stack.Push(top.Current.Steps.GetEnumerator());
            else stack.Pop().Dispose();
        }
    }

    /// <summary>
    /// Sibling blocks (a cell, a text box, a note, a content control): consecutive paragraphs share border groups, and a
    /// paragraph whose mark is deleted joins the paragraph after it.
    /// </summary>
    /// <param name="isAutoSpacingContainer">A table cell or a text box: Word drops the automatic spacing before its first and after its last paragraph.</param>
    private IEnumerable<RenderStep> RenderBlocks(IEnumerable<XElement> blocks, StringBuilder html, bool isAutoSpacingContainer = false)
    {
        List<XElement> siblings = blocks.ToList();
        var merged = new List<XElement>();
        for (int i = 0; i < siblings.Count; i++)
        {
            XElement block = siblings[i];
            XElement? next = i + 1 < siblings.Count && siblings[i + 1].Name == W + "p" ? siblings[i + 1] : null;
            if (block.Name == W + "p" && next is not null && HasDeletedParagraphMark(block)) { merged.Add(block); continue; }
            DocxBlockNeighbours outer = _blockNeighbours;
            int firstOfGroup = i - merged.Count;                       // the group starts at the first joined paragraph
            XElement? before = firstOfGroup > 0 && siblings[firstOfGroup - 1].Name == W + "p" ? siblings[firstOfGroup - 1] : null;
            _blockNeighbours = new DocxBlockNeighbours(before, next, merged.Count > 0 ? merged.ToArray() : null,
                                                       IsFirstInContainer: isAutoSpacingContainer && firstOfGroup == 0,
                                                       IsLastInContainer: isAutoSpacingContainer && i == siblings.Count - 1);
            merged.Clear();
            yield return new(RenderBlock(block, html));
            _blockNeighbours = outer;
        }
    }

    private DocxBlockNeighbours _blockNeighbours;      // the paragraphs around the paragraph being rendered
    private int _cellDepth;                              // > 0 inside a table cell
    private bool _isTrimmingTrailingSpaces;              // the current paragraph drops its trailing spaces (table cell, compatibility mode 11-14)

    /// <summary>
    /// Automatic spacing (w:spacing beforeAutospacing / afterAutospacing), resolved as LibreOffice does: 280 twips, or 100
    /// with the compatibility option doNotUseHTMLParagraphAutoSpacing (DomainMapper.cxx); unless that option is set, none
    /// before the first paragraph of the document, a cell or a text box and none after the last one of a cell or text box
    /// (DomainMapper_Impl finishParagraph / setCellLastParaAfterAutospacing); and none between list paragraphs of the same
    /// list (tdf#133363 there).
    /// </summary>
    private void ResolveAutoSpacing(WordParagraphFormatting formatting, DocxBlockNeighbours neighbours)
    {
        if (!formatting.IsSpaceBeforeAuto && !formatting.IsSpaceAfterAuto) return;
        DocxCompatibilitySettings compatibility = _styles.Compatibility;
        int automatic = compatibility.DoesNotUseHtmlParagraphAutoSpacing ? 100 : 280;
        bool adjusts = !compatibility.DoesNotUseHtmlParagraphAutoSpacing;
        if (formatting.IsSpaceBeforeAuto)
        {
            bool isFirst = neighbours.IsFirstInContainer;
            formatting.SpaceBeforeTwips = adjusts && isFirst || IsSameList(formatting, neighbours.Previous) ? 0 : automatic;
        }
        if (formatting.IsSpaceAfterAuto)
            formatting.SpaceAfterTwips = adjusts && neighbours.IsLastInContainer || IsSameList(formatting, neighbours.Next) ? 0 : automatic;
    }

    private bool IsSameList(WordParagraphFormatting formatting, XElement? neighbour)
    {
        if (formatting.NumberingId <= 0 || neighbour is null) return false;
        XElement? pPr = neighbour.Element(W + "pPr");
        return _styles.ResolveParagraph((string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val"), pPr, _tableStyleId, _tableConditions).NumberingId == formatting.NumberingId;
    }

    /// <summary>Whether the paragraph's first content is a page or column break (w:br type="page" / "column").</summary>
    private static bool StartsWithPageOrColumnBreak(XElement paragraph)
    {
        foreach (XElement element in paragraph.Descendants())
        {
            string name = element.Name.LocalName;
            if (name == "br") return (string?)element.Attribute(W + "type") is "page" or "column";
            if (name is "t" or "tab" or "ptab" or "cr" or "sym" or "drawing" or "pict" or "object" or "footnoteReference" or "endnoteReference" or "noBreakHyphen")
                return false;
        }
        return false;
    }

    /// <summary>
    /// A paragraph whose mark is a tracked deletion or move-away (pPr/rPr/w:del or w:moveFrom): in Word's final view its
    /// text joins the paragraph after it, which keeps its own paragraph properties (Open-Xml-PowerTools
    /// AcceptDeletedAndMoveFromParagraphMarks: the merged paragraph takes the last paragraph's pPr).
    /// </summary>
    public static bool HasDeletedParagraphMark(XElement paragraph)
        => paragraph.Element(W + "pPr")?.Element(W + "rPr") is { } markProperties
           && (markProperties.Element(W + "del") is not null || markProperties.Element(W + "moveFrom") is not null);

    /// <summary>The border group key of a neighbouring paragraph (resolved in the same table context), or null.</summary>
    private string? BorderGroupKeyOf(XElement? paragraph)
    {
        if (paragraph is null) return null;
        XElement? pPr = paragraph.Element(W + "pPr");
        return _styles.ResolveParagraph((string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val"), pPr, _tableStyleId, _tableConditions).BorderGroupKey;
    }

    private IEnumerable<RenderStep> RenderBlock(XElement block, StringBuilder html)
    {
        if (block.Name == W + "p") yield return new(RenderParagraph(block, html));
        else if (block.Name == W + "tbl") yield return new(RenderTable(block, html));
        else if (block.Name == W + "sdt")
        {
            XElement? content = block.Element(W + "sdtContent");
            if (content is not null) yield return new(RenderBlocks(content.Elements(), html));
        }
        // w:sectPr, bookmarks, proofing marks and other body children render nothing.
    }

    private IEnumerable<RenderStep> RenderParagraph(XElement paragraph, StringBuilder html)
    {
        XElement? pPr = paragraph.Element(W + "pPr");
        string? styleId = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val");
        WordParagraphFormatting formatting = _styles.ResolveParagraph(styleId, pPr, _tableStyleId, _tableConditions);
        DocxBlockNeighbours neighbours = _blockNeighbours;
        _blockNeighbours = default;                                  // nested blocks (text boxes) start fresh
        ResolveAutoSpacing(formatting, neighbours);
        // Compatibility mode 15+: a paragraph that starts with a page / column break has no space before (LibreOffice
        // DomainMapper.cxx: deferred break on the first run → top margin 0 when GetWordCompatibilityMode() > 14).
        if (_styles.Compatibility.CompatibilityMode > 14 && _styles.Compatibility.HasCompatibilityMode && StartsWithPageOrColumnBreak(paragraph))
            formatting.SpaceBeforeTwips = 0;
        // Compatibility mode 11-14 (stated in settings.xml): trailing spaces, tabs and no-break spaces of a table-cell
        // paragraph are removed (LibreOffice DomainMapper_Impl, tdf#77417: 0 < mode <= 14 inside a table).
        bool outerTrimming = _isTrimmingTrailingSpaces;
        _isTrimmingTrailingSpaces = _cellDepth > 0 && _styles.Compatibility.HasCompatibilityMode && _styles.Compatibility.CompatibilityMode <= 14;
        if (formatting.BorderGroupKey is { } borderGroup)
        {
            formatting.IsBorderJoinedAbove = BorderGroupKeyOf(neighbours.Previous) == borderGroup;
            formatting.IsBorderJoinedBelow = BorderGroupKeyOf(neighbours.Next) == borderGroup;
        }
        int headingLevel = formatting.OutlineLevel is >= 0 and <= 8 ? formatting.OutlineLevel + 1 : _styles.HeadingLevelOf(styleId);
        string tag = _textBoxDepth > 0 ? "span" : headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";
        string css = formatting.ToCss(_previousStyleKey) + (_textBoxDepth > 0 ? "display:block;" : "");
        // The paragraph mark's own run properties (pPr/rPr) format the list label and give an empty paragraph its height.
        // Resolved only when used: for a label, or when the paragraph is empty.
        bool hasLabel = _listLabel is { Text.Length: > 0 };
        bool isEmpty = pPr?.Element(W + "rPr") is not null && !HasVisibleContent(paragraph);
        WordRunFormatting? mark = hasLabel || isEmpty ? _styles.ResolveRun(formatting.RunFormatting, pPr?.Element(W + "rPr")) : null;
        int markSize = isEmpty ? mark!.EffectiveFontSizeHalfPoints : 0;
        if (markSize > 0 && markSize != formatting.RunFormatting.EffectiveFontSizeHalfPoints)
            css += "font-size:" + (markSize / 2.0).ToString("0.#", CultureInfo.InvariantCulture) + "pt;";

        // A last tab that goes to an end-aligned stop (a table of contents line: "text ...... 12") is laid out with flex:
        // what precedes the tab, a filler drawn with the stop's leader, then the rest pushed to the line's end.
        XElement? outerLeaderTab = _leaderTab;
        string? outerLeaderCss = _leaderCss;
        int outerInlineDepth = _openInline.Count;
        (_leaderTab, _leaderCss) = _textBoxDepth == 0 ? FindEndAlignedLastTab(paragraph, formatting.TabStops) : (null, null);
        if (_leaderTab is not null) css += "display:flex;align-items:baseline;";
        // Word stretches a justified line that ends in a manual line break (LibreOffice: DoNotJustifyLinesWithManualBreak is
        // off unless settings.xml has doNotExpandShiftReturn). CSS justifies a line before a forced break by text-align-last,
        // which would also stretch the paragraph's real last line — so the lines up to the last manual break go in a block of
        // their own with text-align-last:justify, and the rest stays an ordinary last line.
        XElement? outerJustifiedBreak = _justifiedBreak;
        _justifiedBreak = _leaderTab is null && _textBoxDepth == 0 && formatting.Alignment == WordParagraphAlignment.Justify
                          && !_styles.Compatibility.DoesNotExpandShiftReturn ? LastManualBreak(paragraph) : null;

        WordHtmlText.AppendOpenTag(html, tag, css, _textBoxDepth > 0 ? "text-box-paragraph" : null, formatting.IsRightToLeft);
        _isParagraphRightToLeft = formatting.IsRightToLeft;
        if (_leaderTab is not null) html.Append("<span class=\"tab-before\">");
        if (_justifiedBreak is not null) html.Append("<span class=\"justified-lines\" style=\"display:block;text-align-last:justify\">");
        if (_listLabel is { Text.Length: > 0 } label)
        {
            // Word draws the number with the mark's formatting plus the level's rPr (§17.9.6), then the level's suffix (§17.9.28).
            string labelCss = mark!.ToCss(formatting.RunFormatting) + label.LevelCss;
            html.Append("<span class=\"list-marker\"");
            if (labelCss.Length > 0) html.Append(" style=\"").Append(labelCss).Append('"');
            html.Append('>').Append(WordHtmlText.Encode(label.Text)).Append("</span>");
            html.Append(label.Suffix switch { DocxListLabelSuffix.Tab => "\t", DocxListLabelSuffix.Space => " ", _ => "" });
        }
        if (_listLabel is not null) _listLabel = null;              // only the block's first paragraph carries it
        foreach (XElement joined in neighbours.MergedBefore ?? [])          // paragraphs whose deleted marks joined them to this one
            yield return new(RenderInlineChildren(joined, html, formatting.RunFormatting));
        yield return new(RenderInlineChildren(paragraph, html, formatting.RunFormatting));
        if (_leaderTab is not null) html.Append("</span>");          // tab-after (or tab-before when the tab was not reached)
        if (_justifiedBreak is not null) html.Append("</span>");      // after-justified-lines (or justified-lines when the break was not reached)
        _justifiedBreak = outerJustifiedBreak;
        _isTrimmingTrailingSpaces = outerTrimming;
        _leaderTab = outerLeaderTab;
        _leaderCss = outerLeaderCss;
        _openInline.RemoveRange(outerInlineDepth, _openInline.Count - outerInlineDepth);
        html.Append("</").Append(tag).Append('>');
        _previousStyleKey = formatting.StyleKey;

        // A w:sectPr inside pPr ends a section: in a flow view that is a rule, not a new page.
        if (pPr?.Element(W + "sectPr") is not null && _textBoxDepth == 0) html.Append("<hr class=\"section-break\">");
    }

    /// <summary>Whether a paragraph draws anything besides its mark (text, tabs, breaks, symbols, drawings, note references).</summary>
    private static bool HasVisibleContent(XElement paragraph)
    {
        foreach (XElement element in paragraph.Descendants())
        {
            string name = element.Name.LocalName;
            if (name is "t" or "tab" or "ptab" or "br" or "cr" or "sym" or "drawing" or "pict" or "object" or "footnoteReference" or "endnoteReference" or "noBreakHyphen")
                return true;
        }
        return false;
    }

    private IEnumerable<RenderStep> RenderInlineChildren(XElement container, StringBuilder html, WordRunFormatting paragraphBase)
    {
        foreach (XElement child in container.Elements())
        {
            string name = child.Name.LocalName;
            switch (name)
            {
                case "r": yield return new(RenderRun(child, html, paragraphBase)); break;
                case "hyperlink" or "fldSimple" when LinkOpenTagOf(child) is { } linkTag:
                {
                    string? outerLink = _activeLinkOpenTag;
                    _activeLinkOpenTag = linkTag;
                    yield return new(RenderInlineChildren(child, html, paragraphBase));
                    _activeLinkOpenTag = outerLink;
                    break;
                }
                case "bookmarkStart" when (string?)child.Attribute(W + "name") is { Length: > 0 } bookmark && !bookmark.StartsWith("_GoBack", StringComparison.Ordinal):
                    html.Append("<span class=\"bookmark\" id=\"").Append(WordHtmlText.Encode(BookmarkElementId(bookmark))).Append("\"></span>");
                    break;
                case "hyperlink": case "ins": case "moveTo": case "smartTag": case "sdt": case "customXml": case "fldSimple": case "dir": case "bdo":
                    XElement inner = name == "sdt" ? child.Element(W + "sdtContent") ?? child : child;
                    yield return new(RenderInlineChildren(inner, html, paragraphBase));
                    break;
                case "AlternateContent":
                    if (ChosenAlternative(child) is { } chosen) yield return new(RenderInlineChildren(chosen, html, paragraphBase));
                    break;
                // Tracked changes show as Word's "final" view: w:del / w:moveFrom hidden, w:ins / w:moveTo shown.
                // w:pPr, bookmarks, proofErr, commentRange, move ranges… render nothing.
            }
        }
    }

    private bool _isParagraphRightToLeft;   // the paragraph being rendered: a right-to-left run in a left-to-right one is embedded

    private IEnumerable<RenderStep> RenderRun(XElement run, StringBuilder html, WordRunFormatting paragraphBase)
    {
        XElement? rPr = run.Element(W + "rPr");
        WordRunFormatting formatting = _styles.ResolveRun(paragraphBase, rPr);
        // The underline is drawn on the text pieces inside the run (not under word-only spaces, not under the paragraph's
        // trailing spaces); the strikethrough stays on the run's span, so the two never have to share one style.
        string css = formatting.ToCss(paragraphBase, WordDecorationOutput.StrikeOnly) + WordHtmlText.RunDirectionCss(formatting.IsRightToLeft, _isParagraphRightToLeft);
        bool hasSpan = css.Length > 0;
        string openTag = "<span style=\"" + css + "\">";
        if (hasSpan) { html.Append(openTag); _openInline.Add((openTag, "</span>")); }
        // A link is opened inside each run of the link (not around them), so it never crosses a run's span.
        string? linkTag = _activeLinkOpenTag ?? ActiveFieldLink();
        if (linkTag is not null) { html.Append(linkTag); _openInline.Add((linkTag, "</a>")); }
        yield return new(RenderRunChildren(run, html, formatting));
        if (linkTag is not null) { html.Append("</a>"); _openInline.RemoveAt(_openInline.Count - 1); }
        if (hasSpan) { html.Append("</span>"); _openInline.RemoveAt(_openInline.Count - 1); }
    }

    private XElement? _justifiedBreak;                      // a justified paragraph's last manual line break (lines before it are stretched)

    /// <summary>The paragraph's last manual line break (w:br without a type or textWrapping, w:cr) outside deletions.</summary>
    private static XElement? LastManualBreak(XElement paragraph)
    {
        XElement? last = null;
        foreach (XElement element in paragraph.Descendants())
        {
            bool isBreak = element.Name == W + "cr" || element.Name == W + "br" && (string?)element.Attribute(W + "type") is null or "textWrapping";
            if (!isBreak || element.Parent?.Name != W + "r") continue;
            XElement? owner = element.Parent.Parent;
            while (owner is not null && owner.Name != W + "p" && owner.Name != W + "del" && owner.Name != W + "moveFrom") owner = owner.Parent;
            if (owner == paragraph) last = element;
        }
        return last;
    }

    private XElement? _leaderTab;                           // the paragraph's last tab when it goes to an end-aligned stop
    private string? _leaderCss;                             // the filler's CSS (the stop's leader)
    private readonly List<(string open, string close)> _openInline = new();   // inline tags around the current position (closed and reopened at the leader tab)

    /// <summary>The innermost HYPERLINK field whose result is being rendered, as an &lt;a&gt; tag.</summary>
    private string? ActiveFieldLink()
    {
        for (int i = _openFields.Count - 1; i >= 0; i--)
            if (_openFields[i].IsSeparated && _openFields[i].LinkOpenTag is { } tag) return tag;
        return null;
    }

    /// <summary>w:hyperlink (r:id → the relationship's target, w:anchor → a bookmark, w:tooltip) or a w:fldSimple HYPERLINK, as an &lt;a&gt; tag; null when it has no safe target.</summary>
    private string? LinkOpenTagOf(XElement element)
    {
        if (element.Name.LocalName == "fldSimple") return HyperlinkFieldTag((string?)element.Attribute(W + "instr") ?? "");
        string? target = null;
        XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        if ((string?)element.Attribute(relationships + "id") is { } id && _currentHyperlinkTargets is not null) _currentHyperlinkTargets.TryGetValue(id, out target);
        if ((string?)element.Attribute(W + "anchor") is { Length: > 0 } anchor) target = (target ?? "") + "#" + anchor;
        return LinkTag(target, (string?)element.Attribute(W + "tooltip"));
    }

    /// <summary>
    /// A HYPERLINK field instruction, parsed as LibreOffice does (DomainMapper_Impl FIELD_HYPERLINK): the first plain part
    /// is the URL (a doubled backslash is one backslash), \l "name" appends "#name", \o "text" is the tooltip, \t / \m /
    /// \n / \h are ignored here.
    /// </summary>
    private static string? HyperlinkFieldTag(string instruction)
    {
        List<string> parts = FieldInstructionParts(instruction);
        if (parts.Count == 0) return null;
        // Cross-references: REF / PAGEREF / NOTEREF <bookmark> with \h link to the bookmark (ECMA-376 REF field, docx4java
        // mirror: "\h Creates a hyperlink to the bookmarked paragraph"; Microsoft's PAGEREF help says the same of its \h);
        // without \h Word shows the result but it is not a link.
        if (parts[0].ToUpperInvariant() is "REF" or "PAGEREF" or "NOTEREF")
            return parts.Count > 1 && !parts[1].StartsWith("\\", StringComparison.Ordinal) && parts.Skip(2).Any(part => part.Equals("\\h", StringComparison.OrdinalIgnoreCase))
                ? LinkTag("#" + parts[1], null)
                : null;
        if (!parts[0].Equals("HYPERLINK", StringComparison.OrdinalIgnoreCase)) return null;
        string url = "", anchors = "", tooltip = "";
        for (int i = 1; i < parts.Count; i++)
        {
            switch (parts[i])
            {
                case "\\l": if (++i < parts.Count) anchors += "#" + parts[i]; break;
                case "\\o": if (++i < parts.Count) tooltip = parts[i]; break;
                case "\\t": i++; break;
                case "\\m": case "\\n": case "\\h": break;
                default: url = parts[i].Replace("\\\\", "\\"); break;
            }
        }
        return LinkTag(url + anchors, tooltip);
    }

    /// <summary>A field instruction split into words; a quoted part is one word (quotes removed).</summary>
    private static List<string> FieldInstructionParts(string instruction)
    {
        var parts = new List<string>();
        int i = 0;
        while (i < instruction.Length)
        {
            while (i < instruction.Length && char.IsWhiteSpace(instruction[i])) i++;
            if (i >= instruction.Length) break;
            if (instruction[i] == '"')
            {
                int close = instruction.IndexOf('"', i + 1);
                if (close < 0) close = instruction.Length;
                parts.Add(instruction.Substring(i + 1, close - i - 1));
                i = close + 1;
            }
            else
            {
                int start = i;
                while (i < instruction.Length && !char.IsWhiteSpace(instruction[i])) i++;
                parts.Add(instruction.Substring(start, i - start));
            }
        }
        return parts;
    }

    /// <summary>
    /// An &lt;a&gt; tag for a link target, or null when the target is not safe to follow. Only http, https and mailto URLs and
    /// in-document anchors (#bookmark → the bookmark's element id) become links; anything else (javascript:, file:, data:,
    /// relative paths) renders as plain text.
    /// </summary>
    internal static string? LinkTag(string? target, string? tooltip)
    {
        if (string.IsNullOrWhiteSpace(target)) return null;
        target = target!.Trim();
        string href;
        bool isAppLink = false;
        if (target.StartsWith("#", StringComparison.Ordinal)) href = "#" + BookmarkElementId(target.Substring(1));
        else if (Uri.TryCreate(target, UriKind.Absolute, out Uri? uri) && uri.Scheme is "http" or "https" or "mailto") href = uri.AbsoluteUri;
        // links into a library app (Otzaria, Zayit, KitveiHakodesh's own): the host routes them itself, in place
        else if (Uri.TryCreate(target, UriKind.Absolute, out Uri? appUri) && appUri.Scheme is "otzaria" or "zayit" or "kitveihakodeshapp" or "seforimapp")
        {
            href = target;
            isAppLink = true;
        }
        else return null;
        var tag = new StringBuilder(isAppLink ? "<a class=\"link app-link\" data-app-link=\"1\" href=\"" : "<a class=\"link\" href=\"").Append(WordHtmlText.Encode(href)).Append('"');
        if (!string.IsNullOrEmpty(tooltip)) tag.Append(" title=\"").Append(WordHtmlText.Encode(tooltip!)).Append('"');
        if (!href.StartsWith("#", StringComparison.Ordinal) && !isAppLink) tag.Append(" rel=\"noopener noreferrer\" target=\"_blank\"");
        return tag.Append('>').ToString();
    }

    /// <summary>The HTML id a bookmark gets; links to it use "#" + this id.</summary>
    internal static string BookmarkElementId(string bookmarkName) => "bookmark-" + Uri.EscapeDataString(bookmarkName);

    /// <summary>
    /// The paragraph's last tab (w:tab, or w:ptab with alignment right) when it goes to an end-aligned stop, and the CSS of
    /// the leader that fills up to it. With no font metrics, the last tab is taken to reach the paragraph's last custom
    /// stop; that stop must be right / end aligned (ST_TabJc). Leaders (ST_TabTlc): dot and middleDot → dotted, hyphen →
    /// dashed, underscore → solid, heavy → a thick solid line ("a heavy solid line, or an underscore"), none → blank.
    /// </summary>
    private static (XElement? tab, string? css) FindEndAlignedLastTab(XElement paragraph, WordTabStop[] stops)
    {
        XElement? last = null;
        foreach (XElement element in paragraph.Descendants())
        {
            if (element.Name != W + "tab" && element.Name != W + "ptab") continue;
            if (element.Parent?.Name != W + "r") continue;                      // w:tabs/w:tab are stop definitions, not tab characters
            XElement? owner = element.Parent.Parent;
            while (owner is not null && owner.Name != W + "p" && owner.Name != W + "del" && owner.Name != W + "moveFrom") owner = owner.Parent;
            if (owner == paragraph) last = element;
        }
        if (last is null) return (null, null);
        string leader;
        if (last.Name == W + "ptab")
        {
            if ((string?)last.Attribute(W + "alignment") != "right") return (null, null);
            leader = (string?)last.Attribute(W + "leader") ?? "none";
        }
        else
        {
            if (stops.Length == 0 || stops[stops.Length - 1].Alignment is not ("right" or "end")) return (null, null);
            leader = stops[stops.Length - 1].Leader;
        }
        string line = leader switch
        {
            "dot" or "middleDot" => "border-bottom:0.15em dotted currentColor;",
            "hyphen" => "border-bottom:1px dashed currentColor;",
            "underscore" => "border-bottom:1px solid currentColor;",
            "heavy" => "border-bottom:2px solid currentColor;",
            _ => "",
        };
        return (last, "flex:1 1 0;min-width:0.5em;margin:0 0.2em;" + line);
    }

    /// <summary>At the leader tab: close the inline tags, end the part before the tab, draw the filler, start the part after it and reopen the tags.</summary>
    private void AppendLeaderTab(StringBuilder html)
    {
        for (int i = _openInline.Count - 1; i >= 0; i--) html.Append(_openInline[i].close);
        html.Append("</span><span class=\"tab-leader\" style=\"").Append(_leaderCss).Append("\"></span><span class=\"tab-after\" style=\"flex:none\">");
        foreach ((string openTag, _) in _openInline) html.Append(openTag);
        _leaderTab = LeaderTabReached;                        // reached: no other tab matches; the paragraph's closing </span> ends tab-after
    }

    private static readonly XElement LeaderTabReached = new(W + "tab");

    /// <summary>
    /// Run text with the run's underline drawn the way Word draws it (Word 16 PDF export, 2026-10-28): under spaces
    /// between words and under tabs, but not under the spaces that end the paragraph (unless the compatibility option
    /// w:ulTrailSpace asks for them); with w:u words only under the non-space characters (ST_Underline "words": no line
    /// under breaking or non-breaking spaces).
    /// </summary>
    private void AppendRunText(StringBuilder html, XElement textElement, WordRunFormatting formatting)
    {
        string text = textElement.Value;
        if (_isTrimmingTrailingSpaces && text.Length > 0 && IsTrimmedSpace(text[text.Length - 1]) && IsInParagraphEnd(textElement))
            text = text.TrimEnd(' ', '\t', '\u00A0');
        if (!formatting.IsUnderlined) { AppendScriptText(html, text, formatting); return; }
        int contentEnd = text.Length;
        if (contentEnd > 0 && char.IsWhiteSpace(text[contentEnd - 1]) && !UnderlinesTrailingSpaces && IsInParagraphEnd(textElement))
            while (contentEnd > 0 && char.IsWhiteSpace(text[contentEnd - 1])) contentEnd--;
        string underline = formatting.UnderlineCss();
        if (!formatting.IsUnderlineWordsOnly)
        {
            if (contentEnd > 0) AppendUnderlined(html, text.Substring(0, contentEnd), formatting, underline);
        }
        else
        {
            for (int start = 0; start < contentEnd;)
            {
                bool isSpace = char.IsWhiteSpace(text[start]);
                int end = start + 1;
                while (end < contentEnd && char.IsWhiteSpace(text[end]) == isSpace) end++;
                if (isSpace) AppendScriptText(html, text.Substring(start, end - start), formatting);
                else AppendUnderlined(html, text.Substring(start, end - start), formatting, underline);
                start = end;
            }
        }
        if (contentEnd < text.Length) AppendScriptText(html, text.Substring(contentEnd), formatting);
    }

    private static bool IsTrimmedSpace(char c) => c is ' ' or '\t' or '\u00A0';

    private static void AppendUnderlined(StringBuilder html, string text, WordRunFormatting formatting, string underlineCss)
    {
        html.Append("<span style=\"").Append(underlineCss).Append("\">");
        AppendScriptText(html, text, formatting);
        html.Append("</span>");
    }

    /// <summary>A non-text run child (tab, symbol, hyphen) under the run's underline; a word-only underline skips tabs.</summary>
    private static void AppendUnderlinedMarkup(StringBuilder html, string markup, WordRunFormatting formatting, bool isSpace)
    {
        if (!formatting.IsUnderlined || (isSpace && formatting.IsUnderlineWordsOnly)) { html.Append(markup); return; }
        html.Append("<span style=\"").Append(formatting.UnderlineCss()).Append("\">").Append(markup).Append("</span>");
    }

    private XElement? _paragraphEndOf;                              // the paragraph _paragraphEndTexts was computed for
    private readonly HashSet<XElement> _paragraphEndTexts = new();   // its w:t elements that nothing visible follows

    /// <summary>
    /// Whether nothing visible follows <paramref name="textElement"/> in its paragraph (only spaces, properties, deletions).
    /// One backward walk per paragraph, on first need: the w:t elements from the end back to the last visible content.
    /// </summary>
    private bool IsInParagraphEnd(XElement textElement)
    {
        XElement? paragraph = textElement.Parent;
        while (paragraph is not null && paragraph.Name != W + "p") paragraph = paragraph.Parent;
        if (paragraph is null) return false;
        if (_paragraphEndOf != paragraph)
        {
            _paragraphEndOf = paragraph;
            _paragraphEndTexts.Clear();
            // Backward document order without materialising the paragraph: last child first, descending into the last child.
            var pending = new Stack<XElement>();
            for (XNode? node = paragraph.LastNode; node is not null; node = node.PreviousNode) if (node is XElement child) pending.Push(child);
            var reversed = new Stack<XElement>(pending);           // reversed order: last child on top
            pending.Clear();
            while (reversed.Count > 0)
            {
                XElement element = reversed.Pop();
                string name = element.Name.LocalName;
                if (element.Name.Namespace == W && name is "del" or "moveFrom" or "rPr" or "pPr" or "instrText") continue;   // not shown: skip the subtree
                if (name == "t")
                {
                    _paragraphEndTexts.Add(element);
                    if (!string.IsNullOrWhiteSpace(element.Value)) break;
                    continue;
                }
                if (name is "tab" or "ptab" or "br" or "cr" or "sym" or "drawing" or "pict" or "object" or "noBreakHyphen"
                    or "footnoteReference" or "endnoteReference") break;
                for (XNode? node = element.FirstNode; node is not null; node = node.NextNode) if (node is XElement child) reversed.Push(child);
            }
        }
        return _paragraphEndTexts.Contains(textElement);
    }

    /// <summary>Text; CJK characters switch to the run's eastAsia face when it differs (see <see cref="WordRunFormatting.EastAsianFontSwitch"/>).</summary>
    private static void AppendScriptText(StringBuilder html, string text, WordRunFormatting formatting)
    {
        string? eastAsianFont = formatting.EastAsianFontSwitch;
        if (eastAsianFont is null) { html.Append(WordHtmlText.Encode(text)); return; }
        for (int start = 0; start < text.Length;)
        {
            bool isEastAsian = WordRunFormatting.IsEastAsianCharacter(text[start]);
            int end = start + 1;
            while (end < text.Length && WordRunFormatting.IsEastAsianCharacter(text[end]) == isEastAsian) end++;
            string piece = WordHtmlText.Encode(text.Substring(start, end - start));
            if (isEastAsian) html.Append("<span style=\"font-family:'").Append(WordHtmlText.CssFontName(eastAsianFont)).Append("'\">").Append(piece).Append("</span>");
            else html.Append(piece);
            start = end;
        }
    }

    private IEnumerable<RenderStep> RenderRunChildren(XElement run, StringBuilder html, WordRunFormatting formatting)
    {
        foreach (XElement child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t": case "delText": AppendRunText(html, child, formatting); break;
                case "tab": case "ptab":
                    if (child == _leaderTab) AppendLeaderTab(html);
                    else AppendUnderlinedMarkup(html, "\t", formatting, isSpace: true);
                    break;
                case "br" or "cr" when child == _justifiedBreak:
                    // the last manual break: close the justified block here (the break itself ends its last line)
                    for (int i = _openInline.Count - 1; i >= 0; i--) html.Append(_openInline[i].close);
                    // the rest is an ordinary last line; its own block keeps it from taking the paragraph's first-line indent again
                    html.Append("</span><span class=\"after-justified-lines\" style=\"display:block;text-indent:0\">");
                    foreach ((string openTag, _) in _openInline) html.Append(openTag);
                    _justifiedBreak = LeaderTabReached;               // reached: the paragraph's closing </span> ends after-justified-lines
                    break;
                case "br":
                    html.Append(((string?)child.Attribute(W + "type")) == "page" ? "<span class=\"page-break\"></span>" : "<br>");
                    break;
                case "cr": html.Append("<br>"); break;
                case "noBreakHyphen": AppendUnderlinedMarkup(html, "‑", formatting, isSpace: false); break;
                case "softHyphen": html.Append('­'); break;
                case "sym": AppendUnderlinedMarkup(html, RenderSymbol(child), formatting, isSpace: false); break;
                case "footnoteReference": AppendNoteReference(html, WordDocumentNoteKind.Footnote, child); break;
                case "endnoteReference": AppendNoteReference(html, WordDocumentNoteKind.Endnote, child); break;
                case "footnoteRef": case "endnoteRef":
                    // the number at the start of a note body
                    if (_currentNoteLabel is not null)
                        html.Append("<span class=\"note-mark\">").Append(WordHtmlText.Encode(_currentNoteLabel)).Append("</span>");
                    break;
                case "drawing": case "pict": case "object":
                    yield return new(AppendDrawing(html, child));
                    break;
                case "AlternateContent":
                    // mc:AlternateContent: the first choice this renderer can show, else the fallback — never both.
                    // The chosen branch's children are run children: walked in place, not copied (a copy clones recursively).
                    if (ChosenAlternative(child) is { } chosen) yield return new(RenderRunChildren(chosen, html, formatting));
                    break;
                case "pgNum": case "commentReference": case "annotationRef": break;   // no pages; comments are not shown in a reading view
                // separator lines of the notes part, field codes, layout hints: nothing to show
                case "fldChar":
                    switch ((string?)child.Attribute(W + "fldCharType"))
                    {
                        case "begin": _openFields.Add(new DocxField()); break;
                        case "separate" when _openFields.Count > 0:
                            DocxField field = _openFields[_openFields.Count - 1];
                            field.IsSeparated = true;
                            field.LinkOpenTag = HyperlinkFieldTag(field.Instruction.ToString());
                            break;
                        case "end" when _openFields.Count > 0: _openFields.RemoveAt(_openFields.Count - 1); break;
                    }
                    break;
                case "instrText":
                    if (_openFields.Count > 0 && !_openFields[_openFields.Count - 1].IsSeparated) _openFields[_openFields.Count - 1].Instruction.Append(child.Value);
                    break;
                case "separator": case "continuationSeparator": case "rPr": case "lastRenderedPageBreak": break;
            }
        }
    }

    /// <summary>
    /// A drawing's text boxes (kept by <see cref="Docx.DocxDrawingStripper"/>) render as framed inline
    /// blocks; a drawing without text is an image placeholder.
    /// </summary>
    private IEnumerable<RenderStep> AppendDrawing(StringBuilder html, XElement drawing)
    {
        List<XElement> textBoxes = OutermostTextBoxes(drawing);
        if (textBoxes.Count == 0) { html.Append("<span class=\"unsupported\">[image]</span>"); yield break; }
        string? previous = _previousStyleKey;
        // A text box inside a link is not part of the link (no <a> inside <a>, no blocks inside <a>); its own fields start fresh.
        string? outerLink = _activeLinkOpenTag;
        List<DocxField> outerFields = new(_openFields);
        _activeLinkOpenTag = null;
        _openFields.Clear();
        bool outerIsRightToLeft = _isParagraphRightToLeft;   // the text box's own paragraphs set it; the outer runs need theirs back
        _textBoxDepth++;
        foreach (XElement textBox in textBoxes)
        {
            html.Append("<span class=\"text-box\">");
            _previousStyleKey = null;
            yield return new(RenderBlocks(textBox.Elements(), html, isAutoSpacingContainer: true));
            html.Append("</span>");
        }
        _activeLinkOpenTag = outerLink;
        _openFields.Clear();
        _openFields.AddRange(outerFields);
        _textBoxDepth--;
        _isParagraphRightToLeft = outerIsRightToLeft;
        _previousStyleKey = previous;
    }

    /// <summary>
    /// The drawing's w:txbxContent elements that are not inside another one (nested ones render with their
    /// outer box). One walk with its own stack that does not enter a text box: linear in the drawing's size.
    /// </summary>
    private static List<XElement> OutermostTextBoxes(XElement drawing)
    {
        var textBoxes = new List<XElement>();
        var pending = new Stack<XElement>();
        pending.Push(drawing);
        while (pending.Count > 0)
        {
            XElement element = pending.Pop();
            if (element.Name == W + "txbxContent") { textBoxes.Add(element); continue; }
            // pushed last child first so boxes come out in document order
            for (XNode? node = element.LastNode; node is not null; node = node.PreviousNode)
                if (node is XElement child) pending.Push(child);
        }
        return textBoxes;
    }

    /// <summary>mc:AlternateContent → the first mc:Choice (all its requirements are WordprocessingML extensions we degrade gracefully), else mc:Fallback.</summary>
    private static XElement? ChosenAlternative(XElement alternateContent)
        => alternateContent.Elements().FirstOrDefault(e => e.Name.LocalName == "Choice")
        ?? alternateContent.Elements().FirstOrDefault(e => e.Name.LocalName == "Fallback");

    private void AppendNoteReference(StringBuilder html, WordDocumentNoteKind kind, XElement reference)
    {
        int.TryParse((string?)reference.Attribute(W + "id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int noteId);
        (int number, string label)? resolved = _noteLabels?.Invoke(kind, noteId);
        string kindName = kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote";
        bool isCustomMark = reference.Attribute(W + "customMarkFollows") is { } custom && DocxStyleCascade.IsOn(custom);
        html.Append("<sup class=\"note-reference\" data-note-kind=\"").Append(kindName)
            .Append("\" data-note-number=\"").Append(resolved?.number ?? 0).Append("\" data-note-id=\"").Append(noteId).Append("\">");
        if (!isCustomMark) html.Append(WordHtmlText.Encode(resolved?.label ?? "*"));     // a custom mark is the run text that follows
        html.Append("</sup>");
    }

    /// <summary>
    /// w:sym (§17.3.3.30): the character in its own font, as LibreOffice does (DomainMapper LN_EG_RunInnerContent_sym: the run
    /// font becomes w:font with the symbol character set and the character is kept as written). Symbol fonts map their glyphs
    /// at U+F020-F0FF, so a code written without the F0 high byte is moved there. Without w:font the private-use offset is
    /// dropped (the plain character is the best guess).
    /// </summary>
    private static string RenderSymbol(XElement symbol)
    {
        string? code = (string?)symbol.Attribute(W + "char");
        if (code is not { Length: > 0 and <= 4 } || !int.TryParse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int value) || value == 0) return "";
        string? font = (string?)symbol.Attribute(W + "font");
        if (string.IsNullOrEmpty(font))
        {
            if (value >= 0xF000) value -= 0xF000;
            return WordHtmlText.Encode(((char)value).ToString());
        }
        if (value < 0x100) value += 0xF000;
        return "<span style=\"font-family:'" + WordHtmlText.CssFontName(font!) + "'\">" + WordHtmlText.Encode(((char)value).ToString()) + "</span>";
    }

    // ── tables ───────────────────────────────────────────────────────────────────────

    private IEnumerable<RenderStep> RenderTable(XElement table, StringBuilder html)
    {
        XElement? tblPr = table.Element(W + "tblPr");
        string? tableStyleId = (string?)tblPr?.Element(W + "tblStyle")?.Attribute(W + "val");
        // The table style's tblPr first, the table's own tblPr last: the later one wins per property.
        var tablePropertyLayers = _styles.TableStyleProperties(tableStyleId, "tblPr").ToList();
        if (tblPr is not null) tablePropertyLayers.Add(tblPr);
        XElement? Last(string name) => LastTableProperty(tablePropertyLayers, name);

        bool isRightToLeft = Last("bidiVisual") is { } bidi && DocxStyleCascade.IsOn(bidi.Attribute(W + "val"));
        var tableCss = new StringBuilder();
        if (DocxTableFormatting.WidthCss(Last("tblW")) is { } width) tableCss.Append("width:").Append(width).Append(';');
        if (DocxStyleCascade.TwipsAttribute(Last("tblInd"), "w") is { } indent and not 0)
            tableCss.Append("margin-inline-start:").Append(WordParagraphFormatting.TwipsToPt(indent)).Append("pt;");
        if (DocxStyleCascade.TwipsAttribute(Last("tblCellSpacing"), "w") is { } spacing and > 0)
            tableCss.Append("border-collapse:separate;border-spacing:").Append(WordParagraphFormatting.TwipsToPt(spacing * 2)).Append("pt;");
        if ((string?)Last("tblLayout")?.Attribute(W + "type") == "fixed") tableCss.Append("table-layout:fixed;");
        DocxTableFormatting.Borders tableBorders = DocxTableFormatting.ReadBorders(Last("tblBorders"), _styles.Theme);
        string cellPadding = DocxTableFormatting.CellMarginCss(Last("tblCellMar"));
        string? outerTableStyle = _tableStyleId;
        DocxTableConditions outerConditions = _tableConditions;
        _tableStyleId = tableStyleId;
        int look = DocxTableConditionOrder.LookBits(Last("tblLook"));
        // LibreOffice (DomainMapperTableHandler): a look bit whose conditional part the style lacks is dropped (it changes band counting)
        foreach ((int bit, DocxTableConditions condition) in new[] { (0x020, DocxTableConditions.FirstRow), (0x040, DocxTableConditions.LastRow),
                                                                      (0x080, DocxTableConditions.FirstColumn), (0x100, DocxTableConditions.LastColumn) })
            if ((look & bit) != 0 && !_styles.HasTableConditionalPart(tableStyleId, condition)) look &= ~bit;
        int rowBandSize = Math.Max(1, DocxStyleCascade.TwipsAttribute(Last("tblStyleRowBandSize"), "val") ?? 1);
        int columnBandSize = Math.Max(1, DocxStyleCascade.TwipsAttribute(Last("tblStyleColBandSize"), "val") ?? 1);

        WordHtmlText.AppendOpenTag(html, "table", tableCss.ToString(), null, isRightToLeft);
        List<XElement> rows = table.Elements(W + "tr").ToList();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            XElement? trPr = rows[rowIndex].Element(W + "trPr");
            // A deleted row (trPr/w:del) is gone in Word's final view (Open-Xml-PowerTools RevisionProcessor accepts it by removing the row).
            if (trPr?.Element(W + "del") is not null) continue;
            // Row height: a minimum unless hRule="exact" (LibreOffice MeasureHandler: only "exact" is fixed;
            // absent or "auto"/"atLeast" is a minimum). A table row's CSS height is already a minimum.
            int? rowHeight = DocxStyleCascade.TwipsAttribute(trPr?.Element(W + "trHeight"), "val");
            bool isExactHeight = rowHeight is > 0 && (string?)trPr?.Element(W + "trHeight")?.Attribute(W + "hRule") == "exact";
            html.Append(rowHeight is > 0 ? $"<tr style=\"height:{WordParagraphFormatting.TwipsToPt(rowHeight.Value)}pt\">" : "<tr>");
            // w:gridBefore / w:gridAfter: grid columns the row leaves empty before its first / after its last cell.
            int gridBefore = GridSkip(trPr, "gridBefore");
            if (gridBefore > 0) AppendGridGap(html, gridBefore, trPr?.Element(W + "wBefore"));
            int column = gridBefore;
            DocxTableConditions rowConditions = RowConditions(rowIndex, rows.Count, trPr, look, rowBandSize);
            List<XElement> cells = rows[rowIndex].Elements(W + "tc").ToList();
            // LibreOffice counts cells, not grid columns, and the gridBefore / gridAfter gaps are cells of their own.
            int cellCount = cells.Count + (gridBefore > 0 ? 1 : 0) + (GridSkip(trPr, "gridAfter") > 0 ? 1 : 0);
            int cellIndex = gridBefore > 0 ? 1 : 0;
            foreach (XElement cell in cells)
            {
                XElement? tcPr = cell.Element(W + "tcPr");
                int span = DocxStyleCascade.TwipsAttribute(tcPr?.Element(W + "gridSpan"), "val") ?? 1;
                string? verticalMerge = tcPr?.Element(W + "vMerge") is { } merge ? ((string?)merge.Attribute(W + "val") ?? "continue") : null;
                DocxTableConditions conditions = CellConditions(rowConditions, cellIndex++, cellCount, look, columnBandSize);
                if (tcPr?.Element(W + "cellDel") is not null) continue;      // a deleted cell (AcceptDeletedCellsTransform removes it)
                if (verticalMerge == "continue") { column += span; continue; }
                _tableConditions = conditions;

                int rowSpan = verticalMerge == "restart" ? CountVerticalMergeSpan(rows, rowIndex, column) : 1;
                html.Append("<td");
                if (span > 1) html.Append(" colspan=\"").Append(span).Append('"');
                if (rowSpan > 1) html.Append(" rowspan=\"").Append(rowSpan).Append('"');
                string cellCss = DocxTableFormatting.CellCss(tcPr, tableBorders, _styles.Theme, _styles.TableStyleCellProperties(tableStyleId, conditions))
                               + (DocxTableFormatting.CellMarginCss(tcPr?.Element(W + "tcMar")) is { Length: > 0 } ownMargins ? ownMargins : cellPadding)
                               + (tcPr?.Element(W + "noWrap") is { } noWrap && DocxStyleCascade.IsOn(noWrap.Attribute(W + "val")) ? "white-space:nowrap;" : "");
                if (cellCss.Length > 0) html.Append(" style=\"").Append(cellCss).Append('"');
                html.Append('>');
                string? previous = _previousStyleKey;
                _previousStyleKey = null;
                // An exact row height clips what does not fit, as Word does. A cell merged down over several
                // rows is as tall as all of them, so only a single-row cell is clipped to this row's height.
                bool isClipped = isExactHeight && rowSpan == 1;
                if (isClipped) html.Append("<div style=\"max-height:").Append(WordParagraphFormatting.TwipsToPt(rowHeight!.Value)).Append("pt;overflow:hidden\">");
                _cellDepth++;
                yield return new(RenderBlocks(cell.Elements(), html, isAutoSpacingContainer: true));
                _cellDepth--;
                if (isClipped) html.Append("</div>");
                _previousStyleKey = previous;
                html.Append("</td>");
                column += span;
            }
            int gridAfter = GridSkip(trPr, "gridAfter");
            if (gridAfter > 0) AppendGridGap(html, gridAfter, trPr?.Element(W + "wAfter"));
            html.Append("</tr>");
        }
        html.Append("</table>");
        _tableStyleId = outerTableStyle;
        _tableConditions = outerConditions;
    }

    /// <summary>
    /// The row's conditional parts, as LibreOffice decides them (DomainMapperTableHandler::endTableGetCellProperties): a header
    /// row (w:tblHeader) or the first row with the first-row look → firstRow; the last row with the last-row look → lastRow;
    /// any other row is banded unless noHBand — odd bands are band1, counted after the first row when it is styled. The band
    /// size w:tblStyleRowBandSize (§17.7.6.7) groups rows per band (LibreOffice ignores it; 1 by default).
    /// </summary>
    private static DocxTableConditions RowConditions(int rowIndex, int rowCount, XElement? trPr, int look, int bandSize)
    {
        if (trPr?.Element(W + "tblHeader") is { } header && DocxStyleCascade.IsOn(header.Attribute(W + "val"))) return DocxTableConditions.FirstRow;
        if (rowIndex == 0 && (look & 0x020) != 0) return DocxTableConditions.FirstRow;
        if (rowIndex == rowCount - 1 && rowIndex > 0 && (look & 0x040) != 0) return DocxTableConditions.LastRow;
        if ((look & 0x200) != 0) return DocxTableConditions.None;
        int band = (rowIndex - ((look & 0x020) != 0 ? 1 : 0)) / bandSize;
        return band % 2 == 0 ? DocxTableConditions.Band1Horizontal : DocxTableConditions.Band2Horizontal;
    }

    /// <summary>The cell's column part (first / last column or a vertical band) added to its row's, plus the corner cell parts.</summary>
    private static DocxTableConditions CellConditions(DocxTableConditions row, int cellIndex, int cellCount, int look, int bandSize)
    {
        DocxTableConditions column = DocxTableConditions.None;
        if (cellIndex == 0 && (look & 0x080) != 0) column = DocxTableConditions.FirstColumn;
        else if (cellIndex == cellCount - 1 && cellIndex > 0 && (look & 0x100) != 0) column = DocxTableConditions.LastColumn;
        else if ((look & 0x400) == 0)
        {
            int band = (cellIndex - ((look & 0x080) != 0 ? 1 : 0)) / bandSize;
            column = band % 2 == 0 ? DocxTableConditions.Band1Vertical : DocxTableConditions.Band2Vertical;
        }
        DocxTableConditions conditions = row | column;
        if (conditions == (DocxTableConditions.FirstColumn | DocxTableConditions.FirstRow)) conditions |= DocxTableConditions.NorthWestCell;
        else if (conditions == (DocxTableConditions.FirstColumn | DocxTableConditions.LastRow)) conditions |= DocxTableConditions.SouthWestCell;
        else if (conditions == (DocxTableConditions.LastColumn | DocxTableConditions.FirstRow)) conditions |= DocxTableConditions.NorthEastCell;
        else if (conditions == (DocxTableConditions.LastColumn | DocxTableConditions.LastRow)) conditions |= DocxTableConditions.SouthEastCell;
        return conditions;
    }

    private static int GridSkip(XElement? trPr, string name) => Math.Max(0, DocxStyleCascade.TwipsAttribute(trPr?.Element(W + name), "val") ?? 0);

    /// <summary>The empty, borderless grid columns a row skips (w:gridBefore / w:gridAfter), w:wBefore / w:wAfter wide.</summary>
    private static void AppendGridGap(StringBuilder html, int columns, XElement? width)
    {
        html.Append("<td class=\"grid-gap\"");
        if (columns > 1) html.Append(" colspan=\"").Append(columns).Append('"');
        html.Append(" style=\"border:none;padding:0;");
        if (DocxTableFormatting.WidthCss(width) is { } css) html.Append("width:").Append(css).Append(';');
        html.Append("\"></td>");
    }

    private static XElement? LastTableProperty(List<XElement> layers, string name)
        => layers.Select(layer => layer.Element(W + name)).LastOrDefault(element => element is not null);

    private static int CountVerticalMergeSpan(List<XElement> rows, int startRow, int targetColumn)
    {
        int span = 1;
        for (int rowIndex = startRow + 1; rowIndex < rows.Count; rowIndex++)
        {
            int column = GridSkip(rows[rowIndex].Element(W + "trPr"), "gridBefore");
            bool continues = false;
            foreach (XElement cell in rows[rowIndex].Elements(W + "tc"))
            {
                XElement? tcPr = cell.Element(W + "tcPr");
                int cellSpan = DocxStyleCascade.TwipsAttribute(tcPr?.Element(W + "gridSpan"), "val") ?? 1;
                if (column == targetColumn)
                {
                    continues = tcPr?.Element(W + "vMerge") is { } merge && ((string?)merge.Attribute(W + "val") ?? "continue") == "continue";
                    break;
                }
                column += cellSpan;
            }
            if (!continues) break;
            span++;
        }
        return span;
    }
}

/// <summary>
/// The paragraphs next to a block that change how it is drawn: the paragraph before and after (border groups) and the
/// paragraphs before it whose deleted marks join them to it. A body block gets them from the window reader (one adjacent
/// block each side, plus the chain of joined paragraphs); blocks inside a cell or a text box from their siblings.
/// </summary>
public readonly record struct DocxBlockNeighbours(XElement? Previous, XElement? Next, XElement[]? MergedBefore,
                                                  bool IsFirstInContainer = false, bool IsLastInContainer = false, string? OpenFields = null);
