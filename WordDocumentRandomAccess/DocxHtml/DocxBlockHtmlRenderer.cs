using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// Renders one body-level block element (w:p, w:tbl, w:sdt, …) to HTML with resolved formatting.
/// No page layout: paragraphs flow, page breaks become a thin rule. Images are placeholders and
/// footnote references are a marker only (footnote bodies are deferred — see NOTES.md).
/// </summary>
public sealed class DocxBlockHtmlRenderer
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    private readonly DocxStyleCascade _styles;
    private string? _listLabel;     // the index-time label of the block being rendered (its first paragraph)

    public DocxBlockHtmlRenderer(DocxStyleCascade styles)
    {
        _styles = styles;
    }

    /// <param name="listLabel">The label counted at index time ("3.", "ב."), or null when the block is not a list item.</param>
    public string Render(XElement block, string? listLabel = null)
    {
        var html = new StringBuilder(1024);
        _listLabel = listLabel;
        RenderBlock(block, html);
        _listLabel = null;
        return html.ToString();
    }

    /// <summary>Plain text of a block — for headings in the table of contents.</summary>
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

    private void RenderBlock(XElement block, StringBuilder html)
    {
        if (block.Name == W + "p") RenderParagraph(block, html);
        else if (block.Name == W + "tbl") RenderTable(block, html);
        else if (block.Name == W + "sdt")
        {
            XElement? content = block.Element(W + "sdtContent");
            if (content is not null) foreach (XElement child in content.Elements()) RenderBlock(child, html);
        }
        // w:sectPr, bookmarks, proofing marks and other body children render nothing.
    }

    private void RenderParagraph(XElement paragraph, StringBuilder html)
    {
        XElement? pPr = paragraph.Element(W + "pPr");
        string? styleId = (string?)pPr?.Element(W + "pStyle")?.Attribute(W + "val");
        WordParagraphFormatting formatting = _styles.ResolveParagraph(styleId, pPr);
        int headingLevel = formatting.OutlineLevel is >= 0 and <= 8 ? formatting.OutlineLevel + 1 : _styles.HeadingLevelOf(styleId);
        string tag = headingLevel is >= 1 and <= 6 ? "h" + headingLevel : "p";

        WordHtmlText.AppendOpenTag(html, tag, formatting.ToCss(), null, formatting.IsRightToLeft);
        if (_listLabel is { Length: > 0 } label)
        {
            html.Append("<span class=\"list-marker\">").Append(WordHtmlText.Encode(label)).Append("</span> ");
            _listLabel = null;                                      // only the block's first paragraph carries it
        }
        RenderInlineChildren(paragraph, html, formatting.RunFormatting);
        html.Append("</").Append(tag).Append('>');
    }

    private void RenderInlineChildren(XElement container, StringBuilder html, WordRunFormatting paragraphBase)
    {
        foreach (XElement child in container.Elements())
        {
            string name = child.Name.LocalName;
            switch (name)
            {
                case "r": RenderRun(child, html, paragraphBase); break;
                case "hyperlink": case "ins": case "smartTag": case "sdt": case "customXml": case "fldSimple": case "dir": case "bdo":
                    XElement inner = name == "sdt" ? child.Element(W + "sdtContent") ?? child : child;
                    RenderInlineChildren(inner, html, paragraphBase);
                    break;
                // w:del (deleted text), w:pPr, bookmarks, proofErr, commentRange… render nothing.
            }
        }
    }

    private void RenderRun(XElement run, StringBuilder html, WordRunFormatting paragraphBase)
    {
        XElement? rPr = run.Element(W + "rPr");
        WordRunFormatting formatting = _styles.ResolveRun(paragraphBase, rPr);
        string css = formatting.ToCss(paragraphBase);
        bool hasSpan = css.Length > 0;
        if (hasSpan) html.Append("<span style=\"").Append(css).Append("\">");

        foreach (XElement child in run.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "t": case "delText": html.Append(WordHtmlText.Encode(child.Value)); break;
                case "tab": html.Append("<span class=\"tab\"></span>"); break;
                case "br":
                    html.Append(((string?)child.Attribute(W + "type")) == "page" ? "<div class=\"page-break\"></div>" : "<br>");
                    break;
                case "cr": html.Append("<br>"); break;
                case "noBreakHyphen": html.Append('‑'); break;
                case "softHyphen": html.Append('­'); break;
                case "sym": html.Append(RenderSymbol(child)); break;
                case "footnoteReference": case "endnoteReference":
                    html.Append("<sup class=\"footnote-reference\" data-note-id=\"").Append(WordHtmlText.Encode((string?)child.Attribute(W + "id") ?? ""))
                        .Append("\">[").Append(child.Name.LocalName == "footnoteReference" ? "*" : "†").Append("]</sup>");
                    break;
                case "drawing": case "pict": case "object":
                    html.Append("<span class=\"unsupported\">[image]</span>");
                    break;
                case "fldChar": case "instrText": case "rPr": case "lastRenderedPageBreak": break;
            }
        }
        if (hasSpan) html.Append("</span>");
    }

    private static string RenderSymbol(XElement symbol)
    {
        string? code = (string?)symbol.Attribute(W + "char");
        if (code is { Length: 4 } && int.TryParse(code, System.Globalization.NumberStyles.HexNumber, null, out int value))
        {
            if (value >= 0xF000) value -= 0xF000;          // symbol fonts live in the private-use area
            return WordHtmlText.Encode(((char)value).ToString());
        }
        return "";
    }

    // ── tables ───────────────────────────────────────────────────────────────────────

    private void RenderTable(XElement table, StringBuilder html)
    {
        bool isRightToLeft = table.Element(W + "tblPr")?.Element(W + "bidiVisual") is { } bidi && DocxStyleCascade.IsOn(bidi.Attribute(W + "val"));
        WordHtmlText.AppendOpenTag(html, "table", null, null, isRightToLeft);
        List<XElement> rows = table.Elements(W + "tr").ToList();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            html.Append("<tr>");
            int column = 0;
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
                string? fill = (string?)tcPr?.Element(W + "shd")?.Attribute(W + "fill");
                if (fill is { Length: 6 } && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase)) html.Append(" style=\"background:#").Append(fill).Append('"');
                html.Append('>');
                foreach (XElement child in cell.Elements()) RenderBlock(child, html);
                html.Append("</td>");
                column += span;
            }
            html.Append("</tr>");
        }
        html.Append("</table>");
    }

    private static int CountVerticalMergeSpan(List<XElement> rows, int startRow, int targetColumn)
    {
        int span = 1;
        for (int rowIndex = startRow + 1; rowIndex < rows.Count; rowIndex++)
        {
            int column = 0;
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
