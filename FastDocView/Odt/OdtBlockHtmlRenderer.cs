using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;
using FastDocView.DocxHtml;

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
    private static readonly XNamespace XLinkNs = "http://www.w3.org/1999/xlink";

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
    private bool _isInsideLink;      // HTML forbids a link inside a link: an inner one renders its text only
    private int _textBoxDepth;       // inside a text box / shape: paragraphs are block spans (no <p> inside a <p>)
    private bool _isParagraphRightToLeft;

    private static readonly XNamespace XlinkNs = "http://www.w3.org/1999/xlink";
    private static readonly XNamespace OfficeNs = OdtStyleCascade.Office;
    private static readonly XNamespace SvgNs = OdtStyleCascade.Svg;
    private XElement? _currentBlock;

    public OdtBlockHtmlRenderer(OdtStyleCascade styles) => _styles = styles;

    /// <summary>The block being rendered (its index names the pictures stored inside it); -1 when unknown.</summary>
    public int CurrentBlockIndex { get; set; } = -1;

    /// <summary>
    /// Picture ids (<see cref="IWordWindowedDocument.ReadPicture"/>): "p:" + the package entry a draw:image's xlink:href names,
    /// or "b:" + block index + ":" + ordinal for office:binary-data inside the block (flat files) - the ordinal counts the
    /// frames' first draw:image elements that carry binary data, in document order (<see cref="BinaryPictureImages"/>).
    /// </summary>
    public const string PackagePicturePrefix = "p:", BinaryPicturePrefix = "b:";

    /// <summary>The draw:image elements of a block whose office:binary-data a "b:" picture id counts, in document order.</summary>
    public static IEnumerable<XElement> BinaryPictureImages(XElement block)
        => block.Descendants(DrawNs + "frame").Select(frame => frame.Element(DrawNs + "image")).Where(image => image?.Element(OfficeNs + "binary-data") is not null)!;

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
        _isInsideLink = false;
        _renderedNotes = notes.Length == 0 ? null : new List<WordDocumentNoteHtml>(notes.Length);
        _currentBlock = block;
        RenderBlock(block, html, isTopLevel: true);
        _currentBlock = null;
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
        string tag = _textBoxDepth > 0 ? "span" : isHeading && level <= 6 ? "h" + level : "p";

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
        string css = formatting.ToCss(isTopLevel ? _previousStyleKey : null) + (_textBoxDepth > 0 ? "display:block;" : "");
        if (isTopLevel && _listLevel >= 0 && layout is null)
        {
            // list indentation: list-level positions are not read; one step of 0.25in per level plus the label
            css += "margin-inline-start:" + ((formatting.LeftIndentTwips + 360 * (_listLevel + 1)) / 20.0).ToString("0.#", CultureInfo.InvariantCulture) + "pt;";
        }
        if (style.IsPageBreakBefore && isTopLevel && _textBoxDepth == 0) html.Append("<span class=\"page-break\"></span>");
        WordHtmlText.AppendOpenTag(html, tag, css, _textBoxDepth > 0 ? "text-box-paragraph" : null, formatting.IsRightToLeft);
        bool outerIsRightToLeft = _isParagraphRightToLeft;
        _isParagraphRightToLeft = formatting.IsRightToLeft;
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
        _isParagraphRightToLeft = outerIsRightToLeft;
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
                    case "bookmark": case "bookmark-start": case "reference-mark": case "reference-mark-start":
                        // the jump target of in-document links (ODF 1.3 part 3 6.2.1 / 6.2.2, text:name)
                        if ((string?)element.Attribute(TextNs + "name") is { Length: > 0 } bookmark)
                            html.Append("<span class=\"bookmark\" id=\"").Append(WordHtmlText.Encode(DocxBlockHtmlRenderer.BookmarkElementId(bookmark))).Append("\"></span>");
                        continue;
                    case "a":
                    {
                        // text:a xlink:href (ODF 1.3 part 3 6.1.8): "#name" is a bookmark in this document; only safe targets become links
                        string? link = DocxBlockHtmlRenderer.LinkTag((string?)element.Attribute(XLinkNs + "href"), (string?)element.Attribute(OfficeNs + "title"));
                        if (link is not null && !_isInsideLink) html.Append(link);
                        bool wasInsideLink = _isInsideLink;
                        _isInsideLink |= link is not null;
                        RenderInline(element, html, paragraphStyle, paragraphBase, current, hasLabel);
                        _isInsideLink = wasInsideLink;
                        if (link is not null && !wasInsideLink) html.Append("</a>");
                        continue;
                    }
                    case "bookmark-ref": case "reference-ref":
                    {
                        // cross-references (ODF 1.3 part 3 7.7.5 / 7.7.6, text:ref-name): their current text, linked to the target
                        string? link = _isInsideLink ? null : DocxBlockHtmlRenderer.LinkTag("#" + (string?)element.Attribute(TextNs + "ref-name"), null);
                        if ((string?)element.Attribute(TextNs + "ref-name") is not { Length: > 0 }) link = null;
                        if (link is not null) { html.Append(link); _isInsideLink = true; }
                        RenderInline(element, html, paragraphStyle, paragraphBase, current, hasLabel);
                        if (link is not null) { html.Append("</a>"); _isInsideLink = false; }
                        continue;
                    }
                    case "soft-page-break": case "bookmark-end":
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
                if (name.LocalName == "frame") { AppendFrame(element, html); continue; }
                if (IsDrawnShape(element)) { AppendShape(element, html); continue; }
                continue;                                                   // other shapes (polygon, path, connector...), controls, objects: not rendered
            }
            // office:annotation (comments) and unknown namespaces render nothing
        }
    }

    /// <summary>A text box as a placed box (<see cref="AppendTextBox"/>); an image as a picture; an object as a placeholder.</summary>
    private void AppendFrame(XElement frame, StringBuilder html)
    {
        if (frame.Element(DrawNs + "text-box") is { } textBox) { AppendTextBox(frame, textBox, html); return; }
        if (frame.Element(DrawNs + "image") is { } image) { AppendPicture(frame, image, html); return; }
        html.Append("<span class=\"unsupported\">[image]</span>");
    }

    /// <summary>The bytes of an office:binary-data picture (base64), or null when they do not decode.</summary>
    public static byte[]? BinaryPictureBytes(XElement binary)
    {
        try { return Convert.FromBase64String(binary.Value); }
        catch (FormatException) { return null; }
    }

    /// <summary>True when an office:binary-data picture is a format a browser shows (its first bytes' signature).</summary>
    private static bool BinaryPictureFormatIsShown(XElement binary)
    {
        // the first 16 base64 characters are 12 bytes: enough for every signature WordPictureHtml.ExtensionOf knows
        string head = new string(binary.Value.Where(c => !char.IsWhiteSpace(c)).Take(16).ToArray());
        try { return WordPictureHtml.ExtensionOf(Convert.FromBase64String(head)).Length > 0; }
        catch (FormatException) { return false; }
    }

    /// <summary>
    /// A draw:frame's picture at svg:width / svg:height, placed as <see cref="WordPictureHtml"/> describes. text:anchor-type
    /// (ODF 1.3 part 3 19.759; the element attribute wins over the graphic property, default "paragraph"): as-char is in the
    /// line. Otherwise the graphic style (20.400 style:wrap: parallel / dynamic / biggest / left / right = text beside,
    /// none = above and below, run-through = through the text; 20.298 style:horizontal-pos: left / center / right, from-left
    /// with svg:x) and its fo:margin-* (the gap to the text). The bytes: office:binary-data inside draw:image (flat files),
    /// else the package entry xlink:href names (read on demand). The first draw:image is used (LibreOffice writes an SVG
    /// with its PNG fallback as two images of one frame; the first is the preferred one).
    /// </summary>
    private void AppendPicture(XElement frame, XElement image, StringBuilder html)
    {
        string? pictureId = null;
        if (image.Element(OfficeNs + "binary-data") is { } binary)
        {
            if (CurrentBlockIndex >= 0 && _currentBlock is not null && BinaryPictureFormatIsShown(binary))
            {
                int ordinal = 0;
                foreach (XElement candidate in BinaryPictureImages(_currentBlock)) { if (candidate == image) break; ordinal++; }
                pictureId = BinaryPicturePrefix + CurrentBlockIndex.ToString(CultureInfo.InvariantCulture) + ":" + ordinal.ToString(CultureInfo.InvariantCulture);
            }
        }
        else if ((string?)image.Attribute(XlinkNs + "href") is { Length: > 0 } href && !href.Contains(':'))
        {
            string part = href.StartsWith("./", StringComparison.Ordinal) ? href.Substring(2) : href;
            if (WordPictureHtml.ContentTypeOf(part) is not null) pictureId = PackagePicturePrefix + part;
        }
        WordPictureHtml.Append(html, Placement(frame, Points((string?)frame.Attribute(SvgNs + "width")), Points((string?)frame.Attribute(SvgNs + "height"))), pictureId);
    }

    private static double Points(string? length) => length is not null && OdtStyleCascade.Twips(length) is { } twips ? twips / 20.0 : 0;

    /// <summary>A graphic property of a frame / shape through its draw:style-name (parent chain, then the default graphic style).</summary>
    private string? GraphicProperty(XElement element, XName attribute) => _styles.GraphicProperty((string?)element.Attribute(DrawNs + "style-name"), attribute);

    /// <summary>
    /// Where a frame or shape sits, for <see cref="WordPictureHtml.AppendPlacement"/> (pictures, text boxes, shapes share it).
    /// text:anchor-type (ODF 1.3 part 3 19.759; the element attribute wins over the graphic property, default "paragraph"):
    /// as-char is in the line. Otherwise the graphic style: 20.400 style:wrap (parallel / dynamic / biggest / left / right = text
    /// beside, none = above and below, run-through = through the text; "left: text wraps around the left side of the shape",
    /// so the object goes right); 20.298 style:horizontal-pos (left / center / right, from-left with svg:x - physical sides;
    /// inside / outside / from-inside follow page parity, taken as an odd page) and fo:margin-* (the gap to the text).
    /// Text on both sides (parallel / dynamic of a centred object) cannot be laid out in CSS (see NOTES.md): it falls back to
    /// text above and below (<see cref="WordPictureHtml.AppendPlacement"/>). Without a horizontal position or a one-sided wrap
    /// the object goes to the paragraph's start side (right in a right-to-left paragraph) - our choice, the spec gives no
    /// default. style:vertical-pos from-top with svg:y below a paragraph / character anchor (20.397 / 20.398) is added to the
    /// top gap (a float cannot move up; page / frame relations are ignored - the flow view has no page).
    /// </summary>
    private WordPictureLayout Placement(XElement element, double width, double height)
    {
        string? Property(XNamespace ns, string name) => GraphicProperty(element, ns + name);
        string anchor = (string?)element.Attribute(TextNs + "anchor-type") ?? Property(TextNs, "anchor-type") ?? "paragraph";
        string? wrapValue = Property(OdtStyleCascade.Style, "wrap");
        WordPictureWrap wrap = anchor == "as-char" ? WordPictureWrap.Inline : wrapValue switch
        {
            "parallel" or "dynamic" or "biggest" or "left" or "right" => WordPictureWrap.Square,
            "run-through" => WordPictureWrap.None,
            _ => WordPictureWrap.TopAndBottom,
        };
        string? position = Property(OdtStyleCascade.Style, "horizontal-pos");
        WordPictureSide start = _isParagraphRightToLeft ? WordPictureSide.Right : WordPictureSide.Left;
        WordPictureSide side = position switch
        {
            "right" or "outside" => WordPictureSide.Right,
            "center" => WordPictureSide.Center,
            "left" or "inside" or "from-left" or "from-inside" => WordPictureSide.Left,
            _ => wrapValue switch { "left" => WordPictureSide.Right, "right" => WordPictureSide.Left, _ => start },
        };
        double offset = position is "from-left" or "from-inside" ? Math.Max(0, Points((string?)element.Attribute(SvgNs + "x"))) : 0;
        double all = Points(Property(OdtStyleCascade.Fo, "margin"));
        double Gap(string name) => Property(OdtStyleCascade.Fo, "margin-" + name) is { } value ? Points(value) : all;
        double top = Gap("top");
        if (wrap != WordPictureWrap.Inline && Property(OdtStyleCascade.Style, "vertical-pos") == "from-top"
            && Property(OdtStyleCascade.Style, "vertical-rel") is null or "paragraph" or "paragraph-content" or "char" or "line")
            top += Math.Max(0, Points((string?)element.Attribute(SvgNs + "y")));
        return new WordPictureLayout(width, height, wrap, side, top, Gap("right"), Gap("bottom"), Gap("left"), offset);
    }

    // -- text boxes and shapes -------------------------------------------------------------

    /// <summary>
    /// A draw:frame with a draw:text-box as a box where it is anchored (<see cref="Placement"/>), through the shared
    /// <see cref="WordPictureHtml.OpenTextBox"/>. Size: the frame's svg:width / svg:height; fo:min-height of the text box
    /// "overrides the height of a text box specified by the svg:height" (ODF 1.3 part 3 19.240) and the box grows with its
    /// text - LibreOffice xmloff XMLTextFrameContext.cxx:950-1010 reads fo:min-width / fo:min-height as the size with
    /// bMinWidth / bMinHeight, :637-661 sets SizeType::MIN only for text boxes. Percentages are not resolved (no page here).
    /// Look from the graphic style: fo:padding(-*), fo:border (first side found), fill (<see cref="FillCss"/>). The box's own
    /// paragraphs go through <see cref="RenderParagraph"/>; lists and tables in a text box are not drawn (search skips them too).
    /// </summary>
    private void AppendTextBox(XElement frame, XElement textBox, StringBuilder html)
    {
        double width = Points((string?)frame.Attribute(SvgNs + "width")), height = Points((string?)frame.Attribute(SvgNs + "height"));
        bool grows = false;
        // on the text box (19.240 / 19.241), else in the frame's graphic style (20.212 / 20.213 - LibreOffice 25.x writes it there)
        string? minWidth = (string?)textBox.Attribute(OdtStyleCascade.Fo + "min-width") ?? GraphicProperty(frame, OdtStyleCascade.Fo + "min-width");
        string? minHeight = (string?)textBox.Attribute(OdtStyleCascade.Fo + "min-height") ?? GraphicProperty(frame, OdtStyleCascade.Fo + "min-height");
        if (minWidth is not null && !minWidth.Contains('%') && Points(minWidth) > width) width = Points(minWidth);
        if (minHeight is not null && !minHeight.Contains('%')) { height = Points(minHeight); grows = true; }
        (string? borderColor, double borderWidth, bool borderDash) = Border(frame);
        WordPictureHtml.OpenTextBox(html, Placement(frame, width, height),
                                    new WordTextBoxStyle(Padding(frame, "left"), Padding(frame, "top"), Padding(frame, "right"), Padding(frame, "bottom"),
                                                         borderColor, borderWidth, borderDash, FillCss(frame, isTextBox: true), 0, null, grows));
        AppendOwnParagraphs(textBox, html);
        WordPictureHtml.CloseTextBox(html);
    }

    /// <summary>The text:p / text:h children of a text box or shape, rendered as block spans by the normal paragraph renderer.</summary>
    private void AppendOwnParagraphs(XElement container, StringBuilder html)
    {
        // _isInsideLink is kept: a box inside a text:a is inside that <a> in the HTML too (no link inside a link)
        string? savedLabel = _listLabel; int savedLevel = _listLevel; string? savedPrevious = _previousStyleKey; bool savedLink = _isInsideLink;
        _listLabel = null; _listLevel = -1; _previousStyleKey = null;
        _textBoxDepth++;
        foreach (XElement child in container.Elements())
            if (child.Name == TextNs + "p" || child.Name == TextNs + "h") RenderParagraph(child, html, isTopLevel: false);
        _textBoxDepth--;
        _listLabel = savedLabel; _listLevel = savedLevel; _previousStyleKey = savedPrevious; _isInsideLink = savedLink;
    }

    private double Padding(XElement element, string side)
        => GraphicProperty(element, OdtStyleCascade.Fo + ("padding-" + side)) is { } value ? Points(value) : Points(GraphicProperty(element, OdtStyleCascade.Fo + "padding"));

    /// <summary>fo:border (else the first of its sides) as colour / width / dashed; "none" or unset = no border.</summary>
    private (string? color, double width, bool dash) Border(XElement element)
    {
        string? value = GraphicProperty(element, OdtStyleCascade.Fo + "border");
        foreach (string side in new[] { "top", "left", "bottom", "right" })
            value ??= GraphicProperty(element, OdtStyleCascade.Fo + ("border-" + side));
        if (value is null) return (null, 0, false);
        string? color = null; double width = 0; bool dash = false;
        foreach (string part in value.Split(' '))
        {
            if (part.Length == 0) continue;
            if (part[0] == '#') color = Color(part);
            else if (part is "dashed" or "dotted") dash = true;
            else if (part is "none" or "hidden") return (null, 0, false);
            else if (OdtStyleCascade.Twips(part) is { } twips) width = twips / 20.0;
        }
        return (color ?? "#000000", width, dash);
    }

    private static string? Color(string? value) => value is { Length: 7 } && value[0] == '#' && value.Skip(1).All(Uri.IsHexDigit) ? value : null;

    /// <summary>
    /// The fill as a CSS background: draw:fill (ODF 1.3 part 3 20.118; LibreOffice sdpropls.cxx:115-119 maps draw:fill,
    /// draw:fill-color, draw:fill-gradient-name) solid = draw:fill-color (20.119), gradient = the named draw:gradient's
    /// draw:start-color to draw:end-color (GradientStyle.cxx:101-104) as a linear gradient at draw:angle (axial / radial /
    /// square styles, border and steps are approximated as linear); none / hatch / bitmap = no fill. A text box without
    /// draw:fill keeps fo:background-color (Writer frames). Null = transparent.
    /// </summary>
    private string? FillCss(XElement element, bool isTextBox)
    {
        string? fill = GraphicProperty(element, OdtStyleCascade.Draw + "fill");
        if (fill is null) return isTextBox ? Color(GraphicProperty(element, OdtStyleCascade.Fo + "background-color")) : null;
        if (fill == "solid") return Color(GraphicProperty(element, OdtStyleCascade.Draw + "fill-color"));
        if (fill != "gradient" || _styles.Gradient(GraphicProperty(element, OdtStyleCascade.Draw + "fill-gradient-name")) is not { } gradient) return null;
        string from = Color((string?)gradient.Attribute(OdtStyleCascade.Draw + "start-color")) ?? "#000000";
        string to = Color((string?)gradient.Attribute(OdtStyleCascade.Draw + "end-color")) ?? "#ffffff";
        return "linear-gradient(" + Css(GradientAngle((string?)gradient.Attribute(OdtStyleCascade.Draw + "angle"))) + "deg," + from + "," + to + ")";
    }

    /// <summary>
    /// draw:angle as a CSS angle. ODF 0 = the start colour on top, angles counter-clockwise; CSS 180deg runs top to bottom,
    /// clockwise. NOT verified in source: the unit of a plain number (read as degrees; "deg" accepted) and the exact
    /// orientation - a gradient is an approximation anyway.
    /// </summary>
    private static double GradientAngle(string? angle)
    {
        double degrees = 0;
        if (angle is { Length: > 0 })
            double.TryParse(angle.EndsWith("deg", StringComparison.Ordinal) ? angle.Substring(0, angle.Length - 3) : angle, NumberStyles.Float, CultureInfo.InvariantCulture, out degrees);
        return ((180 - degrees) % 360 + 360) % 360;
    }

    private static bool IsDrawnShape(XElement element)
        => element.Name.Namespace == DrawNs && element.Name.LocalName is "custom-shape" or "rect" or "ellipse" or "line" or "g";

    /// <summary>
    /// A drawing shape (draw:custom-shape / draw:rect / draw:ellipse / draw:line / draw:g) as inline SVG in a box placed like a
    /// picture (<see cref="Placement"/>, through <see cref="WordPictureHtml.OpenTextBox"/>): the box is the shape's bounding
    /// box (svg:width / svg:height; a line's svg:x1..y2; a group around its drawn children, each child at its svg:x / svg:y
    /// in it), the SVG fills it, the shape's own paragraphs sit on top from the box's top with its fo:padding
    /// (draw:textarea-vertical-align is not read). The box grows with text that does not fit. Only the shape's XML in the
    /// block being rendered is read (windowed).
    /// </summary>
    private void AppendShape(XElement shape, StringBuilder html)
    {
        (double x, double y, double width, double height) = Bounds(shape);
        WordPictureHtml.OpenTextBox(html, Placement(shape, width, height), new WordTextBoxStyle(0, 0, 0, 0, null, 0, false, null, 0, null, GrowsWithText: true), "text-box word-shape");
        AppendShapeContent(shape, html, x, y, width, height);
        WordPictureHtml.CloseTextBox(html);
    }

    private void AppendShapeContent(XElement shape, StringBuilder html, double x, double y, double width, double height)
    {
        if (shape.Name.LocalName == "g")
        {
            foreach (XElement child in shape.Elements())
            {
                if (!IsDrawnShape(child)) continue;
                (double childX, double childY, double childWidth, double childHeight) = Bounds(child);
                html.Append("<span class=\"word-shape-part\" style=\"left:").Append(Css(childX - x)).Append("pt;top:").Append(Css(childY - y))
                    .Append("pt;width:").Append(Css(childWidth)).Append("pt;height:").Append(Css(childHeight)).Append("pt\">");
                AppendShapeContent(child, html, childX, childY, childWidth, childHeight);
                html.Append("</span>");
            }
            return;
        }
        AppendShapeSvg(shape, html, x, y, width, height);
        if (shape.Elements().Any(child => child.Name == TextNs + "p" || child.Name == TextNs + "h"))
        {
            html.Append("<span class=\"word-shape-text\" style=\"padding:").Append(Css(Padding(shape, "top"))).Append("pt ").Append(Css(Padding(shape, "right")))
                .Append("pt ").Append(Css(Padding(shape, "bottom"))).Append("pt ").Append(Css(Padding(shape, "left"))).Append("pt\">");
            AppendOwnParagraphs(shape, html);
            html.Append("</span>");
        }
    }

    /// <summary>
    /// The shape's geometry in points from its box's corner. Stroke (LibreOffice sdpropls.cxx:99-102): draw:stroke (ODF 1.3
    /// part 3 20.167: none / solid / dash - a dash is one fixed pattern, the draw:stroke-dash style is not read),
    /// svg:stroke-color (20.408), svg:stroke-width (20.410; 0 = the thinnest line, drawn 0.5 pt). Unset stroke = solid
    /// black, unset fill = none (our choice; the document's default graphic style normally sets both). Geometry: draw:rect
    /// (svg:rx, else draw:corner-radius - 19.127), draw:ellipse, draw:line (svg:x1..y2), draw:custom-shape by its
    /// draw:enhanced-geometry draw:type. The spec says draw:type "shall not influence the geometry" (19.229.3) - the shape is
    /// draw:enhanced-path - but enhanced paths (formulas, modifiers) are not evaluated; the type names LibreOffice writes for
    /// its presets (svx EnhancedCustomShapeTypeNames.cxx:36-38, 48, 101-104) are drawn from its preset tables at their
    /// default adjustments (svx EnhancedCustomShapeGeometry.cxx: round-rectangle 174-216, corner mso_sptDefault3600 =
    /// 3600 / 21600; right-arrow 684-724 at 16200 / 5400; left-arrow 726-753 and up-arrow 792-819 at 5400 / 5400; down-arrow
    /// 761-784 at 16200 / 5400; left-right-arrow 827-868 at 4300 / 5400 - the calculation slots are read in mso_sptArrowCalc
    /// 693-703 and assumed alike for the other arrows, NOT verified). draw:modifiers are not applied. Any other type gets its
    /// box with a dashed outline.
    /// </summary>
    private void AppendShapeSvg(XElement shape, StringBuilder html, double x, double y, double width, double height)
    {
        string? strokeKind = GraphicProperty(shape, OdtStyleCascade.Draw + "stroke");
        string stroke = strokeKind == "none" ? "none" : Color(GraphicProperty(shape, OdtStyleCascade.Svg + "stroke-color")) ?? "#000000";
        double strokeWidth = Math.Max(0.5, Points(GraphicProperty(shape, OdtStyleCascade.Svg + "stroke-width")));
        string? fillCss = shape.Name.LocalName == "line" ? null : FillCss(shape, isTextBox: false);
        string fill = "none", defs = "";
        if (fillCss is not null && fillCss.StartsWith("linear-gradient(", StringComparison.Ordinal))
        {
            // the CSS gradient as an SVG linearGradient along the same angle (x1..y2 on the unit box)
            string[] parts = fillCss.Substring(16, fillCss.Length - 17).Split(',');
            double radians = double.Parse(parts[0].Substring(0, parts[0].Length - 3), CultureInfo.InvariantCulture) * Math.PI / 180;
            double dx = Math.Sin(radians) / 2, dy = -Math.Cos(radians) / 2;
            string id = "fdv-gradient-" + CurrentBlockIndex.ToString(CultureInfo.InvariantCulture) + "-" + html.Length.ToString(CultureInfo.InvariantCulture);
            defs = "<defs><linearGradient id=\"" + id + "\" x1=\"" + Css(0.5 - dx) + "\" y1=\"" + Css(0.5 - dy) + "\" x2=\"" + Css(0.5 + dx) + "\" y2=\"" + Css(0.5 + dy)
                 + "\"><stop offset=\"0\" stop-color=\"" + parts[1] + "\"/><stop offset=\"1\" stop-color=\"" + parts[2] + "\"/></linearGradient></defs>";
            fill = "url(#" + id + ")";
        }
        else if (fillCss is not null) fill = fillCss;
        string paint = " fill=\"" + fill + "\" stroke=\"" + stroke + "\" stroke-width=\"" + Css(strokeWidth) + "\""
                     + (strokeKind == "dash" ? " stroke-dasharray=\"" + Css(strokeWidth * 4) + " " + Css(strokeWidth * 2) + "\"" : "");
        html.Append("<svg class=\"word-shape-geometry\" viewBox=\"0 0 ").Append(Css(Math.Max(width, 0.01))).Append(' ').Append(Css(Math.Max(height, 0.01)))
            .Append("\" preserveAspectRatio=\"none\" aria-hidden=\"true\">").Append(defs);
        string local = shape.Name.LocalName;
        string? type = local == "custom-shape" ? (string?)shape.Element(DrawNs + "enhanced-geometry")?.Attribute(DrawNs + "type") : null;
        if (local == "rect" || type == "rectangle")
        {
            double radius = Points((string?)shape.Attribute(SvgNs + "rx")) is > 0 and var rx ? rx : Points((string?)shape.Attribute(DrawNs + "corner-radius"));
            html.Append("<rect width=\"").Append(Css(width)).Append("\" height=\"").Append(Css(height)).Append('"');
            if (radius > 0) html.Append(" rx=\"").Append(Css(radius)).Append('"');
            html.Append(paint).Append("/>");
        }
        else if (type == "round-rectangle")
            html.Append("<rect width=\"").Append(Css(width)).Append("\" height=\"").Append(Css(height)).Append("\" rx=\"").Append(Css(Math.Min(width, height) * 3600 / 21600))
                .Append('"').Append(paint).Append("/>");
        else if (local == "ellipse" || type == "ellipse")
            html.Append("<ellipse cx=\"").Append(Css(width / 2)).Append("\" cy=\"").Append(Css(height / 2)).Append("\" rx=\"").Append(Css(width / 2)).Append("\" ry=\"")
                .Append(Css(height / 2)).Append('"').Append(paint).Append("/>");
        else if (local == "line")
            html.Append("<line x1=\"").Append(Css(Points((string?)shape.Attribute(SvgNs + "x1")) - x)).Append("\" y1=\"").Append(Css(Points((string?)shape.Attribute(SvgNs + "y1")) - y))
                .Append("\" x2=\"").Append(Css(Points((string?)shape.Attribute(SvgNs + "x2")) - x)).Append("\" y2=\"").Append(Css(Points((string?)shape.Attribute(SvgNs + "y2")) - y))
                .Append('"').Append(paint).Append("/>");
        else if (ArrowPoints(type) is { } points)
        {
            html.Append("<polygon points=\"");
            for (int i = 0; i < points.Length; i += 2)
                html.Append(i == 0 ? "" : " ").Append(Css(points[i] * width / 21600)).Append(',').Append(Css(points[i + 1] * height / 21600));
            html.Append('"').Append(paint).Append("/>");
        }
        else
            html.Append("<rect class=\"word-shape-unknown\" width=\"").Append(Css(width)).Append("\" height=\"").Append(Css(height))
                .Append("\" fill=\"none\" stroke=\"#bbbbbb\" stroke-dasharray=\"3 2\"/>");
        html.Append("</svg>");
    }

    /// <summary>LibreOffice's arrow presets at their default adjustments, in the 21600 x 21600 preset frame (see <see cref="AppendShapeSvg"/>).</summary>
    private static int[]? ArrowPoints(string? type) => type switch
    {
        "right-arrow" => [0, 5400, 16200, 5400, 16200, 0, 21600, 10800, 16200, 21600, 16200, 16200, 0, 16200],
        "left-arrow" => [21600, 5400, 5400, 5400, 5400, 0, 0, 10800, 5400, 21600, 5400, 16200, 21600, 16200],
        "up-arrow" => [5400, 21600, 5400, 5400, 0, 5400, 10800, 0, 21600, 5400, 16200, 5400, 16200, 21600],
        "down-arrow" => [5400, 0, 5400, 16200, 0, 16200, 10800, 21600, 21600, 16200, 16200, 16200, 16200, 0],
        "left-right-arrow" => [0, 10800, 4300, 0, 4300, 5400, 17300, 5400, 17300, 0, 21600, 10800, 17300, 21600, 17300, 16200, 4300, 16200, 4300, 21600],
        _ => null,
    };

    /// <summary>A shape's bounding box in points: svg:x / svg:y / svg:width / svg:height; a line from svg:x1..y2; a group around its drawn children.</summary>
    private static (double x, double y, double width, double height) Bounds(XElement shape)
    {
        if (shape.Name.LocalName == "line")
        {
            double x1 = Points((string?)shape.Attribute(SvgNs + "x1")), y1 = Points((string?)shape.Attribute(SvgNs + "y1"));
            double x2 = Points((string?)shape.Attribute(SvgNs + "x2")), y2 = Points((string?)shape.Attribute(SvgNs + "y2"));
            return (Math.Min(x1, x2), Math.Min(y1, y2), Math.Abs(x2 - x1), Math.Abs(y2 - y1));
        }
        if (shape.Name.LocalName == "g")
        {
            double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
            foreach (XElement child in shape.Elements())
            {
                if (!IsDrawnShape(child)) continue;
                (double x, double y, double width, double height) = Bounds(child);
                left = Math.Min(left, x); top = Math.Min(top, y); right = Math.Max(right, x + width); bottom = Math.Max(bottom, y + height);
            }
            return left == double.MaxValue ? (0, 0, 0, 0) : (left, top, right - left, bottom - top);
        }
        return (Points((string?)shape.Attribute(SvgNs + "x")), Points((string?)shape.Attribute(SvgNs + "y")),
                Points((string?)shape.Attribute(SvgNs + "width")), Points((string?)shape.Attribute(SvgNs + "height")));
    }

    private static string Css(double value) => value.ToString("0.##", CultureInfo.InvariantCulture);


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
