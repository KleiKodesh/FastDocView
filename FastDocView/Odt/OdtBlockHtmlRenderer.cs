using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// Renders one OpenDocument block (text:p, text:h or table:table) to HTML with resolved formatting,
/// and the bodies of the notes it references. Notes are inline in ODF (text:note inside the
/// paragraph): the reference becomes &lt;sup class="note-reference" data-note-kind data-note-number&gt;
/// with the number counted at index time, and the text:note-body is rendered separately as a
/// <see cref="WordDocumentNoteHtml"/> delivered with the block. No page layout; images are placeholders.
/// </summary>
public sealed class OdtBlockHtmlRenderer
{
    private static readonly XNamespace TextNs = OdtStyleCascade.Text;
    private static readonly XNamespace TableNs = OdtStyleCascade.Table;
    private static readonly XNamespace DrawNs = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";

    /// <summary>Inline elements the renderer handles; anything else renders its text content (fields) — the audit counts the rest.</summary>
    public static readonly HashSet<string> HandledInlineElements =
    [
        "text:span", "text:s", "text:tab", "text:line-break", "text:soft-page-break", "text:a", "text:note", "text:note-citation", "text:note-body",
        "text:bookmark", "text:bookmark-start", "text:bookmark-end", "text:reference-mark", "text:reference-mark-start", "text:reference-mark-end",
        "text:number", "draw:frame", "draw:a", "office:annotation", "office:annotation-end", "text:change", "text:change-start", "text:change-end",
        "text:alphabetical-index-mark", "text:alphabetical-index-mark-start", "text:alphabetical-index-mark-end", "text:toc-mark",
        "text:toc-mark-start", "text:toc-mark-end", "text:user-index-mark", "text:user-index-mark-start", "text:user-index-mark-end",
        "text:bibliography-mark", "text:meta", "text:ruby", "text:ruby-base", "text:ruby-text", "text:hidden-text",
        // fields: their current text is rendered
        "text:page-number", "text:page-count", "text:date", "text:time", "text:sequence", "text:sequence-ref", "text:bookmark-ref",
        "text:reference-ref", "text:note-ref", "text:chapter", "text:title", "text:subject", "text:author-name", "text:file-name",
        "text:variable-get", "text:variable-set", "text:user-field-get", "text:user-defined", "text:initial-creator", "text:creation-date",
        "text:word-count", "text:character-count", "text:paragraph-count", "text:hidden-paragraph", "text:conditional-text", "text:placeholder",
        "text:drop-down", "text:text-input", "text:expression", "text:sender-firstname", "text:sender-lastname",
        "draw:custom-shape", "draw:rect", "draw:line", "draw:g", "draw:ellipse", "draw:polygon", "draw:polyline", "draw:path", "draw:connector",
        "draw:control", "draw:text-box", "draw:image", "draw:object", "draw:object-ole",
    ];

    private readonly OdtStyleCascade _styles;
    private string? _listLabel;
    private int _listLevel;
    private string _listLayout = "";   // the label's level geometry (OdtListStyle.Level.Layout)
    private string? _previousStyleKey;
    private List<OdtNoteReference> _blockNotes = new();
    private int _nextNote;
    private int _noteBodyDepth;
    private List<WordDocumentNoteHtml>? _renderedNotes;

    public OdtBlockHtmlRenderer(OdtStyleCascade styles) => _styles = styles;

    /// <param name="listLevel">List depth (0-based) of a list paragraph, -1 otherwise.</param>
    /// <param name="notes">The note references of this block, in order (from the index).</param>
    public (string html, IReadOnlyList<WordDocumentNoteHtml> notes) Render(XElement block, string? listLabel, int listLevel, string listLayout,
                                                                          ReadOnlySpan<OdtNoteReference> notes, string? previousStyleKey)
    {
        var html = new StringBuilder(1024);
        _listLayout = listLayout;
        _listLabel = listLabel;
        _listLevel = listLevel;
        _previousStyleKey = previousStyleKey;
        _blockNotes = notes.ToArray().ToList();
        _nextNote = 0;
        _renderedNotes = notes.Length == 0 ? null : new List<WordDocumentNoteHtml>(notes.Length);
        RenderBlock(block, html, isTopLevel: true);
        return (html.ToString(), (IReadOnlyList<WordDocumentNoteHtml>?)_renderedNotes ?? []);
    }

    /// <summary>The paragraph style key of a block (for the next block's contextual spacing) — text:style-name.</summary>
    public static string? StyleKeyOf(XElement block) => (string?)block.Attribute(TextNs + "style-name");

    private void RenderBlock(XElement block, StringBuilder html, bool isTopLevel)
    {
        if (block.Name == TextNs + "p" || block.Name == TextNs + "h") RenderParagraph(block, html, isTopLevel);
        else if (block.Name == TableNs + "table") RenderTable(block, html);
        else if (block.Name == TextNs + "list") RenderNestedList(block, html, 0);
        else if (block.Name == TextNs + "section" || block.Name == TextNs + "index-body")
            foreach (XElement child in block.Elements()) RenderBlock(child, html, isTopLevel: false);
        else if (block.Name.LocalName is "table-of-content" or "alphabetical-index" or "illustration-index" or "table-index" or "object-index" or "user-index" or "bibliography")
        {
            if (block.Element(TextNs + "index-body") is { } body) RenderBlock(body, html, isTopLevel: false);
        }
    }

    private void RenderParagraph(XElement paragraph, StringBuilder html, bool isTopLevel)
    {
        string? styleName = (string?)paragraph.Attribute(TextNs + "style-name");
        OdtStyleCascade.OdtParagraphStyle style = _styles.ResolveParagraph(styleName);
        WordParagraphFormatting formatting = style.Formatting;
        bool isHeading = paragraph.Name == TextNs + "h";
        int level = 1;
        if (isHeading && int.TryParse((string?)paragraph.Attribute(TextNs + "outline-level"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 1) level = parsed;
        string tag = isHeading && level <= 6 ? "h" + level : "p";

        // label-alignment mode: the level's fo:margin-left / fo:text-indent indent the paragraph (ODF 1.2 19.881), its tab
        // stop is where the text after a listtab starts. Older label-width-and-position lists keep the per-level step below.
        int[]? layout = isTopLevel && _listLevel >= 0 ? ParseLayout(_listLayout) : null;
        if (layout is not null)
        {
            formatting = formatting.Clone();
            formatting.LeftIndentTwips = layout[0];
            formatting.IsLeftIndentPhysical = false;
            formatting.FirstLineIndentTwips = layout[1];
            if (layout[2] >= 0) formatting.TabStops = [new WordTabStop(layout[2], "left", "none")];
        }
        string css = formatting.ToCss(isTopLevel ? _previousStyleKey : null);
        if (isTopLevel && _listLevel >= 0 && layout is null)
        {
            // list indentation: list-level positions are not read; one step of 0.25in per level plus the label
            css += "margin-inline-start:" + ((formatting.LeftIndentTwips + 360 * (_listLevel + 1)) / 20.0).ToString("0.#", CultureInfo.InvariantCulture) + "pt;";
        }
        if (style.IsPageBreakBefore && isTopLevel) html.Append("<span class=\"page-break\"></span>");
        WordHtmlText.AppendOpenTag(html, tag, css, null, formatting.IsRightToLeft);
        bool hasLabel = false;
        if (isTopLevel && _listLabel is { Length: > 0 } label)
        {
            if (layout is not null)
                WordListLabelHtml.Append(html, label, null, layout[3] switch { 1 => WordListLabelSuffix.Space, 2 => WordListLabelSuffix.Nothing, _ => WordListLabelSuffix.Tab },
                                         formatting, WordListLabelHtml.DefaultTabStopTwips);
            else html.Append("<span class=\"list-marker\">").Append(WordHtmlText.Encode(label)).Append("</span> ");
            hasLabel = true;
        }
        RenderInline(paragraph, html, styleName, formatting.RunFormatting, formatting.RunFormatting, hasLabel);
        html.Append("</").Append(tag).Append('>');
    }

    /// <summary>"margin,indent,tab,follow" (OdtListStyle.Level.Layout), null when empty or damaged.</summary>
    private static int[]? ParseLayout(string layout)
    {
        string[] parts = layout.Split(',');
        if (parts.Length != 4) return null;
        var values = new int[4];
        for (int i = 0; i < 4; i++)
            if (!int.TryParse(parts[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out values[i])) return null;
        return values;
    }

    private void RenderInline(XElement container, StringBuilder html, string? paragraphStyle, WordRunFormatting paragraphBase,
                              WordRunFormatting current, bool hasLabel)
    {
        foreach (XNode node in container.Nodes())
        {
            if (node is XText text) { html.Append(WordHtmlText.Encode(text.Value)); continue; }
            if (node is not XElement element) continue;
            XName name = element.Name;
            if (name.Namespace == TextNs)
            {
                switch (name.LocalName)
                {
                    case "span":
                    {
                        string? spanStyle = (string?)element.Attribute(TextNs + "style-name");
                        WordRunFormatting run = spanStyle is null ? current
                            : ReferenceEquals(current, paragraphBase) ? _styles.ResolveSpan(paragraphStyle, current, spanStyle)
                            : _styles.ApplySpan(current, spanStyle);                // nested span: over its parent span, not cached
                        string css = run.ToCss(current);
                        if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">");
                        RenderInline(element, html, paragraphStyle, paragraphBase, run, hasLabel);
                        if (css.Length > 0) html.Append("</span>");
                        continue;
                    }
                    case "s":
                    {
                        int spaces = int.TryParse((string?)element.Attribute(TextNs + "c"), out int c) && c > 0 ? Math.Min(c, 1000) : 1;
                        html.Append(' ', spaces);
                        continue;
                    }
                    case "tab": html.Append('\t'); continue;
                    case "line-break": html.Append("<br>"); continue;
                    case "note": AppendNote(element, html); continue;
                    case "number": if (!hasLabel) html.Append(WordHtmlText.Encode(element.Value)).Append(' '); continue;
                    case "soft-page-break": case "bookmark": case "bookmark-start": case "bookmark-end": case "reference-mark": case "reference-mark-start":
                    case "reference-mark-end": case "change": case "change-start": case "change-end": case "toc-mark": case "toc-mark-start": case "toc-mark-end":
                    case "alphabetical-index-mark": case "alphabetical-index-mark-start": case "alphabetical-index-mark-end": case "user-index-mark":
                    case "user-index-mark-start": case "user-index-mark-end": case "bibliography-mark": case "ruby-text": case "hidden-text": case "hidden-paragraph":
                        continue;
                    default:
                        RenderInline(element, html, paragraphStyle, paragraphBase, current, hasLabel);     // text:a, text:meta, ruby base, fields: their text
                        continue;
                }
            }
            if (name.Namespace == DrawNs)
            {
                if (name.LocalName == "a") { RenderInline(element, html, paragraphStyle, paragraphBase, current, hasLabel); continue; }
                if (name.LocalName == "frame") { AppendFrame(element, html, paragraphStyle, paragraphBase, current); continue; }
                continue;                                                   // shapes: not rendered
            }
            // office:annotation (comments) and unknown namespaces render nothing
        }
    }

    /// <summary>A text box's paragraphs inline; an image or object as a placeholder.</summary>
    private void AppendFrame(XElement frame, StringBuilder html, string? paragraphStyle, WordRunFormatting paragraphBase, WordRunFormatting current)
    {
        if (frame.Element(DrawNs + "text-box") is { } textBox)
        {
            html.Append("<span class=\"text-box\">");
            foreach (XElement child in textBox.Elements())
                if (child.Name == TextNs + "p" || child.Name == TextNs + "h")
                {
                    html.Append("<span style=\"display:block\">");
                    RenderInline(child, html, paragraphStyle, paragraphBase, current, hasLabel: true);
                    html.Append("</span>");
                }
            html.Append("</span>");
            return;
        }
        html.Append("<span class=\"unsupported\">[image]</span>");
    }

    private void AppendNote(XElement note, StringBuilder html)
    {
        WordDocumentNoteKind kind = (string?)note.Attribute(TextNs + "note-class") == "endnote" ? WordDocumentNoteKind.Endnote : WordDocumentNoteKind.Footnote;
        // the scanner records only notes with a body, outside other notes: consume references the same way
        bool isIndexed = _noteBodyDepth == 0 && note.Element(TextNs + "note-body") is not null;
        if (_noteBodyDepth > 0) return;                                     // a note inside a note body: not shown
        OdtNoteReference? reference = isIndexed && _nextNote < _blockNotes.Count ? _blockNotes[_nextNote++] : null;
        int number = reference?.Number ?? 0;
        string label = reference?.Label ?? ((string?)note.Element(TextNs + "note-citation"))?.Trim() ?? "*";
        string kindName = kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote";
        html.Append("<sup class=\"note-reference\" data-note-kind=\"").Append(kindName).Append("\" data-note-number=\"")
            .Append(number.ToString(CultureInfo.InvariantCulture)).Append("\">").Append(WordHtmlText.Encode(label)).Append("</sup>");
        if (reference is null || _renderedNotes is null) return;

        // the body, rendered on its own (draft view: notes pane), never inline
        string? savedLabel = _listLabel; int savedLevel = _listLevel; string? savedPrevious = _previousStyleKey;
        _listLabel = null; _listLevel = -1; _previousStyleKey = null;
        var body = new StringBuilder(256);
        bool isFirst = true;
        foreach (XElement child in note.Element(TextNs + "note-body")?.Elements() ?? [])
        {
            int before = body.Length;
            _noteBodyDepth++;
            RenderBlock(child, body, isTopLevel: false);
            _noteBodyDepth--;
            if (isFirst && body.Length > before)
            {
                // the note's mark at the start of its first paragraph, as Word's notes pane shows it
                int afterOpenTag = body.ToString().IndexOf('>', before) + 1;
                body.Insert(afterOpenTag, "<span class=\"note-mark\">" + WordHtmlText.Encode(label) + "</span>");
                isFirst = false;
            }
        }
        _listLabel = savedLabel; _listLevel = savedLevel; _previousStyleKey = savedPrevious;
        _renderedNotes.Add(new WordDocumentNoteHtml(kind, number, label, body.ToString()));
    }

    /// <summary>A list inside a table cell or a note: not flattened by the index; items get a bullet / their order number.</summary>
    private void RenderNestedList(XElement list, StringBuilder html, int depth)
    {
        int ordinal = 0;
        foreach (XElement item in list.Elements())
        {
            bool isItem = item.Name == TextNs + "list-item";
            if (isItem) ordinal++;
            bool isFirstParagraph = true;
            foreach (XElement child in item.Elements())
            {
                if (child.Name == TextNs + "list") { RenderNestedList(child, html, depth + 1); isFirstParagraph = false; continue; }
                string? savedLabel = _listLabel; int savedLevel = _listLevel;
                _listLabel = isItem && isFirstParagraph ? "•" : null;
                _listLevel = depth;
                RenderParagraphAsListItem(child, html);
                _listLabel = savedLabel; _listLevel = savedLevel;
                isFirstParagraph = false;
            }
        }
    }

    private void RenderParagraphAsListItem(XElement paragraph, StringBuilder html)
    {
        if (paragraph.Name != TextNs + "p" && paragraph.Name != TextNs + "h") { RenderBlock(paragraph, html, isTopLevel: false); return; }
        RenderParagraph(paragraph, html, isTopLevel: true);
    }

    // ── tables ────────────────────────────────────────────────────────────────────────

    private void RenderTable(XElement table, StringBuilder html)
    {
        (string tableCss, bool? isRightToLeft) = _styles.TableCss((string?)table.Attribute(TableNs + "style-name"));
        WordHtmlText.AppendOpenTag(html, "table", tableCss, null, isRightToLeft);
        var widths = new List<string?>();
        CollectColumns(table, widths);
        if (widths.Any(w => w is not null))
        {
            html.Append("<colgroup>");
            foreach (string? width in widths) html.Append(width is null ? "<col>" : "<col style=\"width:" + width + "\">");
            html.Append("</colgroup>");
        }
        foreach (XElement row in Rows(table))
        {
            html.Append("<tr>");
            foreach (XElement cell in row.Elements())
            {
                if (cell.Name != TableNs + "table-cell") continue;          // table:covered-table-cell: part of a span
                html.Append("<td");
                if ((int?)cell.Attribute(TableNs + "number-columns-spanned") is > 1 and var columns) html.Append(" colspan=\"").Append(columns).Append('"');
                if ((int?)cell.Attribute(TableNs + "number-rows-spanned") is > 1 and var rows) html.Append(" rowspan=\"").Append(rows).Append('"');
                string cellCss = _styles.CellCss((string?)cell.Attribute(TableNs + "style-name"));
                if (cellCss.Length > 0) html.Append(" style=\"").Append(cellCss).Append('"');
                html.Append('>');
                string? savedLabel = _listLabel; int savedLevel = _listLevel;
                _listLabel = null; _listLevel = -1;
                foreach (XElement child in cell.Elements()) RenderBlock(child, html, isTopLevel: false);
                _listLabel = savedLabel; _listLevel = savedLevel;
                html.Append("</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</table>");
    }

    /// <summary>table:table-column (number-columns-repeated), also inside column groups / header columns.</summary>
    private void CollectColumns(XElement container, List<string?> widths)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == TableNs + "table-column")
            {
                int repeat = Math.Min(1024, Math.Max(1, (int?)child.Attribute(TableNs + "number-columns-repeated") ?? 1));
                string? width = _styles.ColumnWidthCss((string?)child.Attribute(TableNs + "style-name"));
                for (int i = 0; i < repeat; i++) widths.Add(width);
            }
            else if (child.Name.LocalName is "table-columns" or "table-column-group" or "table-header-columns") CollectColumns(child, widths);
        }
    }

    /// <summary>table:table-row, also inside header rows / row groups / table:table-rows.</summary>
    private static IEnumerable<XElement> Rows(XElement container)
    {
        foreach (XElement child in container.Elements())
        {
            if (child.Name == TableNs + "table-row") yield return child;
            else if (child.Name.LocalName is "table-rows" or "table-row-group" or "table-header-rows")
                foreach (XElement row in Rows(child)) yield return row;
        }
    }
}
