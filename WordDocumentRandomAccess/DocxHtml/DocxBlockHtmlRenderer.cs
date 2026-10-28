using System.Globalization;
using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocxHtml;

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
    private string? _listLabel;           // the index-time label of the block being rendered (its first paragraph)
    private string? _previousStyleKey;    // the style of the paragraph rendered just before (contextual spacing)
    private string? _currentNoteLabel;    // inside a note body: the label its w:footnoteRef/w:endnoteRef shows
    private WordDocumentNoteKind _currentNoteKind;
    private string? _tableStyleId;        // inside a table: its style, layered under paragraph styles
    private int _textBoxDepth;            // inside a text box paragraphs become block spans (a <p> may not hold a <p>)

    public DocxBlockHtmlRenderer(DocxStyleCascade styles, NoteLabelResolver? noteLabels = null)
    {
        _styles = styles;
        _noteLabels = noteLabels;
    }

    /// <param name="listLabel">The label counted at index time ("3.", "ב."), or null when the block is not a list item.</param>
    /// <param name="previousStyleKey">The paragraph style of the block before (for contextual spacing), from the index.</param>
    public string Render(XElement block, string? listLabel = null, string? previousStyleKey = null)
    {
        var html = new StringBuilder(1024);
        _listLabel = listLabel;
        _previousStyleKey = previousStyleKey;
        Run(RenderBlock(block, html));
        _listLabel = null;
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
        WordParagraphFormatting formatting = _styles.ResolveParagraph(styleId, pPr, _tableStyleId);
        int headingLevel = formatting.OutlineLevel is >= 0 and <= 8 ? formatting.OutlineLevel + 1 : _styles.HeadingLevelOf(styleId);
        string tag = _textBoxDepth > 0 ? "span" : headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";
        string css = formatting.ToCss(_previousStyleKey) + (_textBoxDepth > 0 ? "display:block;" : "");

        WordHtmlText.AppendOpenTag(html, tag, css, _textBoxDepth > 0 ? "text-box-paragraph" : null, formatting.IsRightToLeft);
        if (_listLabel is { Length: > 0 } label)
        {
            html.Append("<span class=\"list-marker\">").Append(WordHtmlText.Encode(label)).Append("</span> ");
            _listLabel = null;                                      // only the block's first paragraph carries it
        }
        yield return new(RenderInlineChildren(paragraph, html, formatting.RunFormatting));
        html.Append("</").Append(tag).Append('>');
        _previousStyleKey = formatting.StyleKey;

        // A w:sectPr inside pPr ends a section: in a flow view that is a rule, not a new page.
        if (pPr?.Element(W + "sectPr") is not null && _textBoxDepth == 0) html.Append("<hr class=\"section-break\">");
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
        string css = formatting.ToCss(paragraphBase);
        bool hasSpan = css.Length > 0;
        if (hasSpan) html.Append("<span style=\"").Append(css).Append("\">");
        yield return new(RenderRunChildren(run, html));
        if (hasSpan) html.Append("</span>");
    }

    private IEnumerable<RenderStep> RenderRunChildren(XElement run, StringBuilder html)
    {
        foreach (XElement child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t": case "delText": html.Append(WordHtmlText.Encode(child.Value)); break;
                case "tab": case "ptab": html.Append('\t'); break;
                case "br":
                    html.Append(((string?)child.Attribute(W + "type")) == "page" ? "<span class=\"page-break\"></span>" : "<br>");
                    break;
                case "cr": html.Append("<br>"); break;
                case "noBreakHyphen": html.Append('‑'); break;
                case "softHyphen": html.Append('­'); break;
                case "sym": html.Append(RenderSymbol(child)); break;
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
                    if (ChosenAlternative(child) is { } chosen) yield return new(RenderRunChildren(chosen, html));
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

    private static string RenderSymbol(XElement symbol)
    {
        string? code = (string?)symbol.Attribute(W + "char");
        if (code is { Length: 4 } && int.TryParse(code, NumberStyles.HexNumber, null, out int value))
        {
            if (value >= 0xF000) value -= 0xF000;          // symbol fonts live in the private-use area
            return WordHtmlText.Encode(((char)value).ToString());
        }
        return "";
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
        _tableStyleId = tableStyleId;

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
            foreach (XElement cell in rows[rowIndex].Elements(W + "tc"))
            {
                XElement? tcPr = cell.Element(W + "tcPr");
                int span = DocxStyleCascade.TwipsAttribute(tcPr?.Element(W + "gridSpan"), "val") ?? 1;
                string? verticalMerge = tcPr?.Element(W + "vMerge") is { } merge ? ((string?)merge.Attribute(W + "val") ?? "continue") : null;
                if (verticalMerge == "continue") { column += span; continue; }

                int rowSpan = verticalMerge == "restart" ? CountVerticalMergeSpan(rows, rowIndex, column) : 1;
                html.Append("<td");
                if (span > 1) html.Append(" colspan=\"").Append(span).Append('"');
                if (rowSpan > 1) html.Append(" rowspan=\"").Append(rowSpan).Append('"');
                string cellCss = DocxTableFormatting.CellCss(tcPr, tableBorders, _styles.Theme)
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
