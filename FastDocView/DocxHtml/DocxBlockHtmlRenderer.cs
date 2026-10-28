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

    public DocxBlockHtmlRenderer(DocxStyleCascade styles, NoteLabelResolver? noteLabels = null)
    {
        _styles = styles;
        _noteLabels = noteLabels;
    }

    /// <param name="listLabel">The label counted at index time ("3.", "ב."), or null when the block is not a list item.</param>
    /// <param name="previousStyleKey">The paragraph style of the block before (for contextual spacing), from the index.</param>
    public string Render(XElement block, DocxListLabel? listLabel = null, string? previousStyleKey = null)
    {
        var html = new StringBuilder(1024);
        _listLabel = listLabel;
        _previousStyleKey = previousStyleKey;
        Run(RenderBlock(block, html));
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
        Run(RenderBlocks(note.Elements(), html));
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

    private IEnumerable<RenderStep> RenderBlocks(IEnumerable<XElement> blocks, StringBuilder html)
    {
        foreach (XElement block in blocks) yield return new(RenderBlock(block, html));
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

        WordHtmlText.AppendOpenTag(html, tag, css, _textBoxDepth > 0 ? "text-box-paragraph" : null, formatting.IsRightToLeft);
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
        yield return new(RenderInlineChildren(paragraph, html, formatting.RunFormatting));
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

    private IEnumerable<RenderStep> RenderRun(XElement run, StringBuilder html, WordRunFormatting paragraphBase)
    {
        XElement? rPr = run.Element(W + "rPr");
        WordRunFormatting formatting = _styles.ResolveRun(paragraphBase, rPr);
        // The underline is drawn on the text pieces inside the run (not under word-only spaces, not under the paragraph's
        // trailing spaces); the strikethrough stays on the run's span, so the two never have to share one style.
        string css = formatting.ToCss(paragraphBase, WordDecorationOutput.StrikeOnly);
        bool hasSpan = css.Length > 0;
        if (hasSpan) html.Append("<span style=\"").Append(css).Append("\">");
        yield return new(RenderRunChildren(run, html, formatting));
        if (hasSpan) html.Append("</span>");
    }

    /// <summary>
    /// Run text with the run's underline drawn the way Word draws it (Word 16 PDF export, 2026-10-28): under spaces
    /// between words and under tabs, but not under the spaces that end the paragraph (unless the compatibility option
    /// w:ulTrailSpace asks for them); with w:u words only under the non-space characters (ST_Underline "words": no line
    /// under breaking or non-breaking spaces).
    /// </summary>
    private void AppendRunText(StringBuilder html, XElement textElement, WordRunFormatting formatting)
    {
        string text = textElement.Value;
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
                case "tab": case "ptab": AppendUnderlinedMarkup(html, "\t", formatting, isSpace: true); break;
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
                case "separator": case "continuationSeparator": case "fldChar": case "instrText": case "rPr": case "lastRenderedPageBreak": break;
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
        _textBoxDepth++;
        foreach (XElement textBox in textBoxes)
        {
            html.Append("<span class=\"text-box\">");
            _previousStyleKey = null;
            yield return new(RenderBlocks(textBox.Elements(), html));
            html.Append("</span>");
        }
        _textBoxDepth--;
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
                yield return new(RenderBlocks(cell.Elements(), html));
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
