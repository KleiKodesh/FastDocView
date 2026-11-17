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
public sealed partial class OdtBlockHtmlRenderer
{
    private static readonly XNamespace TextNs = OdtStyleCascade.Text;
    private static readonly XNamespace TableNs = OdtStyleCascade.Table;
    internal static readonly XNamespace DrawNs = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static readonly XNamespace XLinkNs = "http://www.w3.org/1999/xlink";
    private static readonly XNamespace LibreOfficeNs = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";

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
        "text:description", "text:keywords", "text:creation-time", "text:modification-date", "text:modification-time",
        "text:table-count", "text:image-count", "text:object-count",
        "draw:custom-shape", "draw:rect", "draw:line", "draw:g", "draw:ellipse", "draw:circle", "draw:polygon", "draw:polyline", "draw:path", "draw:connector",
        "draw:measure", "draw:caption", "draw:regular-polygon", "draw:control", "draw:text-box", "draw:image", "draw:object", "draw:object-ole",
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

    internal static readonly XNamespace XlinkNs = "http://www.w3.org/1999/xlink";
    private static readonly XNamespace OfficeNs = OdtStyleCascade.Office;
    private static readonly XNamespace SvgNs = OdtStyleCascade.Svg;
    private XElement? _currentBlock;

    public OdtBlockHtmlRenderer(OdtStyleCascade styles) => _styles = styles;

    /// <summary>settings.xml TabsRelativeToIndent (<see cref="OdtPackage.ReadTabsRelativeToIndent"/>): stops count from the start indent.</summary>
    public bool AreTabStopsRelativeToIndent { get; init; } = true;

    /// <summary>Computes the fields LibreOffice recomputes when it prints (<see cref="OdtFieldComputer"/>); null: every field shows its saved text.</summary>
    internal OdtFieldComputer? Fields { get; set; }

    /// <summary>The text: elements <see cref="OdtFieldComputer.Compute"/> may give a text for.</summary>
    private static readonly HashSet<string> ComputedFieldNames =
    [
        "sequence", "variable-set", "variable-get", "user-field-get", "expression", "conditional-text", "hidden-text", "bookmark-ref", "reference-ref",
        "sequence-ref", "note-ref", "chapter", "date", "time", "title", "subject", "description", "keywords", "initial-creator", "creation-date",
        "creation-time", "modification-date", "modification-time", "user-defined", "file-name",
        "word-count", "character-count", "paragraph-count", "table-count", "image-count", "object-count",
    ];

    /// <summary>The text: elements whose shown text may differ from the saved one: what search asks the computer for (same set).</summary>
    internal static IReadOnlyCollection<string> ComputedFieldElementNames => ComputedFieldNames;

    /// <summary>A computed field's text, run like a text node (drop cap, script pieces).</summary>
    private void AppendComputedField(StringBuilder html, string value, WordRunFormatting current)
    {
        if (_dropCapRemaining != 0 && value.Length > 0) value = TakeDropCap(value, current);
        AppendScriptPieces(html, value, current, -1);
    }

    /// <summary>The block being rendered (its index names the pictures stored inside it); -1 when unknown.</summary>
    /// <summary>The text width (points) of the section of the block being rendered: the column a floating object picks its nearer edge in (0 = unknown).</summary>
    public double ColumnWidthPt { get; set; }
    /// <summary>Where that column starts from the page's left edge (points; -1 = not known).</summary>
    public double ColumnLeftPt { get; set; } = -1;
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
        else if (block.Name.Namespace == DrawNs) RenderDrawingBlock(block, html);
        else if (block.Name == TextNs + "section" || block.Name == TextNs + "index-body")
            foreach (XElement child in block.Elements()) RenderBlock(child, html, isTopLevel: false);
        else if (block.Name.LocalName is "table-of-content" or "alphabetical-index" or "illustration-index" or "table-index" or "object-index" or "user-index" or "bibliography")
        {
            if (block.Element(TextNs + "index-body") is { } body) RenderBlock(body, html, isTopLevel: false);
        }
    }

    /// <summary>A page-anchored frame or shape standing between paragraphs: drawn in the flow where it stands (the flow view has no page), in a paragraph of its own.</summary>
    private void RenderDrawingBlock(XElement drawing, StringBuilder html)
    {
        string tag = HoldsBoxTable(drawing) ? "div" : "p";
        WordHtmlText.AppendOpenTag(html, tag, "", tag == "div" ? "word-paragraph-block" : null, null);
        RenderInline(new XNode[] { drawing }, html, null, _styles.DefaultParagraph.Formatting.RunFormatting, _styles.DefaultParagraph.Formatting.RunFormatting, hasLabel: false);
        html.Append("</").Append(tag).Append('>');
    }

    private void RenderParagraph(XElement paragraph, StringBuilder html, bool isTopLevel)
    {
        if (Fields?.IsParagraphHidden(paragraph) == true) return;          // text:hidden-paragraph with a true condition: not printed
        string? styleName = (string?)paragraph.Attribute(TextNs + "style-name");
        OdtStyleCascade.OdtParagraphStyle style = _styles.ResolveParagraph(styleName);
        WordParagraphFormatting formatting = style.Formatting;
        bool isHeading = paragraph.Name == TextNs + "h";
        int level = 1;
        if (isHeading && int.TryParse((string?)paragraph.Attribute(TextNs + "outline-level"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed >= 1) level = parsed;
        string tag = _textBoxDepth > 0 ? "span" : isHeading && level <= 6 ? "h" + level : HoldsBoxTable(paragraph) ? "div" : "p";

        // label-alignment mode: the level's fo:margin-left / fo:text-indent indent the paragraph (ODF 1.2 19.881), its tab
        // stop is where the text after a listtab starts. Older label-width-and-position lists keep the per-level step below.
        int[]? layout = isTopLevel && _listLevel >= 0 ? ParseLayout(_listLayout) : null;
        if (layout is not null)
        {
            formatting = formatting.Clone();
            formatting.LeftIndentTwips = layout[0];
            formatting.IsLeftIndentPhysical = false;
            formatting.FirstLineIndentTwips = layout[1];
            if (layout[2] >= 0) formatting.TabStops = [.. formatting.TabStops.Append(new WordTabStop(layout[2] - (AreTabStopsRelativeToIndent ? layout[0] : 0), "left", "none")).OrderBy(stop => stop.PositionTwips)];
        }
        // ODT stops count from the start indent (unless settings.xml says otherwise); the tab data and the list label take them
        // from the text area's edge
        if (layout is null) formatting = formatting.Clone();
        formatting.AreTabStopsRelativeToIndent = AreTabStopsRelativeToIndent;
        if (formatting.TabStops.Length > 0 && AreTabStopsRelativeToIndent)
        {
            int indent = formatting.StartIndentTwips;
            formatting.TabStops = [.. formatting.TabStops.Select(stop => stop with { PositionTwips = stop.PositionTwips + indent }).OrderBy(stop => stop.PositionTwips)];
        }
        if (formatting.RunFormatting.Hyphenate == true)
        {
            formatting = formatting.Clone();
            formatting.IsHyphenated = true;                     // ODT has no document switch: fo:hyphenate on the paragraph style
        }
        string css = formatting.ToCss(isTopLevel ? _previousStyleKey : null) + (_textBoxDepth > 0 ? "display:block;" : "");
        if (isTopLevel && _listLevel >= 0 && layout is null)
        {
            // list indentation: list-level positions are not read; one step of 0.25in per level plus the label
            css += "margin-inline-start:" + ((formatting.LeftIndentTwips + 360 * (_listLevel + 1)) / 20.0).ToString("0.#", CultureInfo.InvariantCulture) + "pt;";
        }
        if (style.IsPageBreakBefore && isTopLevel && _textBoxDepth == 0) html.Append("<span class=\"page-break\"></span>");
        WordHtmlText.AppendOpenTag(html, tag, css, _textBoxDepth > 0 ? "text-box-paragraph" : tag == "div" ? "word-paragraph-block" : null, formatting.IsRightToLeft,
                                   formatting.ParagraphAttributes);
        bool outerIsRightToLeft = _isParagraphRightToLeft;
        int outerDropCap = _dropCapRemaining;
        WordParagraphFormatting? outerDropCapFormat = _dropCapFormat;
        (StringBuilder? target, int insertAt, StringBuilder? captured) outerDropCapCapture = (_dropCapTarget, _dropCapInsertAt, _dropCapCaptured);
        _dropCapRemaining = formatting.DropCapLines > 0 ? formatting.DropCapCharacterCount : 0;
        _dropCapFormat = formatting;
        _dropCapCaptured = null;
        _isParagraphRightToLeft = formatting.IsRightToLeft;
        (Dictionary<XText, int>? offsets, bool[]? complex) outerScript = (_paragraphTextOffsets, _paragraphTextIsComplex);
        ClassifyParagraphScript(paragraph);
        bool hasLabel = false;
        if (isTopLevel && _listLabel is { Length: > 0 } label)
        {
            string? labelCss = LabelCss(formatting.RunFormatting, (string?)paragraph.Attribute(LibreOfficeNs + "marker-style-name"), label, formatting.IsRightToLeft);
            if (layout is not null)
                WordListLabelHtml.Append(html, label, labelCss, layout[3] switch { 1 => WordListLabelSuffix.Space, 2 => WordListLabelSuffix.Nothing, _ => WordListLabelSuffix.Tab },
                                         formatting, WordListLabelHtml.DefaultTabStopTwips);
            else WordListLabelHtml.Append(html, label, labelCss, WordListLabelSuffix.Space, formatting, WordListLabelHtml.DefaultTabStopTwips);
            hasLabel = true;
        }
        _dropCapTarget = html;
        _dropCapInsertAt = html.Length;                       // the initial goes here, before any span the paragraph opens
        RenderInline(paragraph, html, styleName, formatting.RunFormatting, formatting.RunFormatting, hasLabel);
        FinishDropCap();                                      // a paragraph shorter than the drop cap: what there is
        _dropCapRemaining = outerDropCap;
        _dropCapFormat = outerDropCapFormat;
        (_dropCapTarget, _dropCapInsertAt, _dropCapCaptured) = outerDropCapCapture;
        _isParagraphRightToLeft = outerIsRightToLeft;
        (_paragraphTextOffsets, _paragraphTextIsComplex) = outerScript;
        html.Append("</").Append(tag).Append('>');
    }

    private int _dropCapRemaining;                       // characters of the paragraph drop cap still to wrap (-1 = the first word)
    private WordParagraphFormatting? _dropCapFormat;
    private StringBuilder? _dropCapTarget, _dropCapCaptured;   // the paragraph's html and the initial's characters taken so far (each piece in its span's look)
    private int _dropCapInsertAt;

    /// <summary>
    /// Takes the paragraph drop cap's characters from the start of <paramref name="value"/> (style:length characters, or the first
    /// word), wherever they are: they may run across several text nodes and spans (LibreOffice SwDropPortion takes the first
    /// characters of the paragraph text whatever their attributes - sw/source/core/text/txtdrop.cxx SwTextFormatter::CalcDropLen /
    /// NewDropPortion count over the paragraph string). Returns the rest of the text.
    /// </summary>
    private string TakeDropCap(string value, WordRunFormatting current)
    {
        int take;
        bool isDone;
        if (_dropCapRemaining < 0)
        {
            take = 0;
            while (take < value.Length && !char.IsWhiteSpace(value[take])) take++;
            isDone = take < value.Length;
        }
        else
        {
            take = Math.Min(_dropCapRemaining, value.Length);
            if (take > 0 && take < value.Length && char.IsHighSurrogate(value[take - 1])) take++;
            _dropCapRemaining -= take;
            isDone = _dropCapRemaining <= 0;
        }
        if (take > 0)
        {
            _dropCapCaptured ??= new StringBuilder();
            // the piece in its span's look, at the paragraph's size: the drop-cap span sets the initial's size
            WordRunFormatting look = current;
            if (current.FontSizeHalfPoints != _dropCapFormat!.RunFormatting.FontSizeHalfPoints
                || current.ComplexScriptFontSizeHalfPoints != _dropCapFormat.RunFormatting.ComplexScriptFontSizeHalfPoints)
            {
                look = current.Clone();
                look.FontSizeHalfPoints = _dropCapFormat.RunFormatting.FontSizeHalfPoints;
                look.ComplexScriptFontSizeHalfPoints = _dropCapFormat.RunFormatting.ComplexScriptFontSizeHalfPoints;
            }
            string css = look.ToCss(_dropCapFormat.RunFormatting);
            if (css.Length > 0) _dropCapCaptured.Append("<span style=\"").Append(css).Append("\">");
            WordHtmlText.AppendScaled(_dropCapCaptured, WordHtmlText.Encode(value.Substring(0, take)), current.CharacterScalePercent);
            if (css.Length > 0) _dropCapCaptured.Append("</span>");
        }
        if (isDone) FinishDropCap();
        return value.Substring(take);
    }

    /// <summary>Puts the captured initial in its drop-cap span at the paragraph's start; nothing when none was captured.</summary>
    private void FinishDropCap()
    {
        _dropCapRemaining = 0;
        if (_dropCapCaptured is not { Length: > 0 } captured || _dropCapTarget is null) { _dropCapCaptured = null; return; }
        var span = new StringBuilder(captured.Length + 96);
        WordDropCapHtml.Open(span, _dropCapFormat!);
        span.Append(captured);
        WordDropCapHtml.Close(span);
        _dropCapTarget.Insert(Math.Min(_dropCapInsertAt, _dropCapTarget.Length), span.ToString());
        _dropCapCaptured = null;
    }

    /// <summary>
    /// The label's own look after the paragraph's: the level's text:style-name character style (LibreOffice draws the number
    /// portion with it: xmlnumi.cxx:276-278 CharStyleName) and a kept symbol bullet font (OdtListStyle.Level.Layout's tail).
    /// Under the character style: the paragraph mark's automatic text style, loext:marker-style-name (txtparai.cxx:1650-1652,
    /// :1765-1771 -> ListAutoFormat; LibreOffice draws it under the level's character style, which wins item by item:
    /// sw/source/core/text/txtfld.cxx:334-383 checkApplyParagraphMarkFormatToNumbering. The ApplyParagraphMarkFormatToNumbering
    /// setting that gates it there is not read: LibreOffice writes the attribute only for the Word imports that set it).
    /// The Latin or complex-script set follows the label's characters and the paragraph's direction (WordListLabelScript;
    /// a recoded symbol bullet counts as its symbol code, strong left-to-right).
    /// </summary>
    private string? LabelCss(WordRunFormatting paragraphRun, string? markerStyle, string label, bool isParagraphRightToLeft)
    {
        int look = _listLayout.IndexOf(OdtListStyle.LayoutLabelLookSeparator);
        string[] parts = look < 0 ? [] : _listLayout.Substring(look + 1).Split(OdtListStyle.LayoutLabelLookSeparator);
        // the script first: the text styles pick fo:font-weight or style:font-weight-complex (and size, font) by it
        bool isLatinBullet = parts.Length > 2 && parts[2] == "L";
        WordRunFormatting labelRun = WordListLabelScript.WithScript(paragraphRun, !isLatinBullet && WordListLabelScript.UsesComplexScript(label, isParagraphRightToLeft));
        if (markerStyle is { Length: > 0 }) labelRun = _styles.ApplySpan(labelRun, markerStyle, isComplexStrict: true);
        if (parts.Length > 0 && parts[0].Length > 0) labelRun = _styles.ApplySpan(labelRun, parts[0], isComplexStrict: true);
        string css = ReferenceEquals(labelRun, paragraphRun) ? "" : labelRun.ToCss(paragraphRun);
        if (parts.Length > 1 && parts[1].Length > 0) css += "font-family:'" + WordHtmlText.Encode(parts[1].Replace("'", "")) + "';";
        return css.Length > 0 ? css : null;
    }

    /// <summary>"margin,indent,tab,follow" (OdtListStyle.Level.Layout), null when empty or damaged.</summary>
    private static int[]? ParseLayout(string layout)
    {
        int look = layout.IndexOf(OdtListStyle.LayoutLabelLookSeparator);
        string[] parts = (look < 0 ? layout : layout.Substring(0, look)).Split(',');
        if (parts.Length != 4) return null;
        var values = new int[4];
        for (int i = 0; i < 4; i++)
            if (!int.TryParse(parts[i], NumberStyles.Integer, CultureInfo.InvariantCulture, out values[i])) return null;
        return values;
    }

    /// <summary>
    /// Text in the set LibreOffice draws each piece with (WordTextScript): in a right-to-left paragraph the run's formatting is
    /// the complex-script set, so an English piece switches to the western font / size / weight; in a left-to-right paragraph a
    /// Hebrew piece switches to the complex-script set.
    /// </summary>
    private void AppendScriptPieces(StringBuilder html, string value, WordRunFormatting current, int paragraphOffset)
    {
        if (value.Length == 0) return;
        if (paragraphOffset >= 0 && _paragraphTextIsComplex is { } complex && paragraphOffset + value.Length <= complex.Length)
        {
            // the paragraph-wide split (ClassifyParagraphScript): a piece may start or end inside this text node
            for (int start = 0; start < value.Length;)
            {
                bool isComplex = complex[paragraphOffset + start];
                int end = start + 1;
                while (end < value.Length && complex[paragraphOffset + end] == isComplex) end++;
                string piece = WordHtmlText.Encode(value.Substring(start, end - start));
                string css = isComplex == current.UsesComplexScriptProperties ? "" : WordListLabelScript.WithScript(current, isComplex).ToCss(current);
                if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">");
                WordHtmlText.AppendScaled(html, piece, current.CharacterScalePercent);
                if (css.Length > 0) html.Append("</span>");
                start = end;
            }
            return;
        }
        if (WordTextScript.PieceEnd(value, 0, _isParagraphRightToLeft, out bool firstIsComplex) == value.Length && firstIsComplex == current.UsesComplexScriptProperties)
        {
            WordHtmlText.AppendScaled(html, WordHtmlText.Encode(value), current.CharacterScalePercent);
            return;
        }
        for (int start = 0; start < value.Length;)
        {
            int end = WordTextScript.PieceEnd(value, start, _isParagraphRightToLeft, out bool isComplex);
            string piece = WordHtmlText.Encode(value.Substring(start, end - start));
            string css = isComplex == current.UsesComplexScriptProperties ? "" : WordListLabelScript.WithScript(current, isComplex).ToCss(current);
            if (css.Length > 0) html.Append("<span style=\"").Append(css).Append("\">");
            WordHtmlText.AppendScaled(html, piece, current.CharacterScalePercent);
            if (css.Length > 0) html.Append("</span>");
            start = end;
        }
    }

    private Dictionary<XText, int>? _paragraphTextOffsets;     // each text node of the current paragraph -> its offset in the paragraph text
    private bool[]? _paragraphTextIsComplex;                    // per character of the paragraph text: drawn in the complex-script set

    /// <summary>
    /// The script split over the whole paragraph text, as LibreOffice does it: SwScriptInfo::InitScriptInfo runs one
    /// direction / script scanner over the text node's whole string (sw/source/core/text/porlay.cxx:1368, 1467-1469), not per
    /// portion, so an English run broken by a span is still one left-to-right run. The text is built in the order RenderInline
    /// visits it; notes, frames and shapes stand as one neutral anchor character (CH_TXTATR_BREAKWORD), as in the node text.
    /// </summary>
    private void ClassifyParagraphScript(XElement paragraph)
    {
        var text = new StringBuilder();
        var offsets = new Dictionary<XText, int>();
        var pending = new Stack<IEnumerator<XNode>>();
        pending.Push(paragraph.Nodes().GetEnumerator());
        while (pending.Count > 0)
        {
            IEnumerator<XNode> nodes = pending.Peek();
            if (!nodes.MoveNext()) { pending.Pop(); continue; }
            if (nodes.Current is XText node) { offsets[node] = text.Length; text.Append(node.Value); continue; }
            if (nodes.Current is not XElement element) continue;
            XName name = element.Name;
            if (name.Namespace == TextNs)
            {
                switch (name.LocalName)
                {
                    case "s": text.Append(' ', int.TryParse((string?)element.Attribute(TextNs + "c"), out int c) && c > 0 ? Math.Min(c, 1000) : 1); continue;
                    case "tab": text.Append('\t'); continue;
                    case "line-break": text.Append('\n'); continue;
                    case "note": case "number": text.Append('\u0001'); continue;
                    case "bookmark": case "bookmark-start": case "reference-mark": case "reference-mark-start":
                    case "soft-page-break": case "bookmark-end":
                    case "reference-mark-end": case "change": case "change-start": case "change-end": case "toc-mark": case "toc-mark-start": case "toc-mark-end":
                    case "alphabetical-index-mark": case "alphabetical-index-mark-start": case "alphabetical-index-mark-end": case "user-index-mark":
                    case "user-index-mark-start": case "user-index-mark-end": case "bibliography-mark": case "ruby-text": case "hidden-text": case "hidden-paragraph":
                        continue;
                    default: pending.Push(element.Nodes().GetEnumerator()); continue;
                }
            }
            if (name.Namespace == DrawNs)
            {
                if (name.LocalName == "a") pending.Push(element.Nodes().GetEnumerator());
                else text.Append('\u0001');
            }
        }
        string value = text.ToString();
        var isComplex = new bool[value.Length];
        for (int start = 0; start < value.Length;)
        {
            int end = WordTextScript.PieceEnd(value, start, _isParagraphRightToLeft, out bool complex);
            if (complex) for (int k = start; k < end; k++) isComplex[k] = true;
            start = end;
        }
        _paragraphTextOffsets = offsets;
        _paragraphTextIsComplex = isComplex;
    }

    private void RenderInline(XElement container, StringBuilder html, string? paragraphStyle, WordRunFormatting paragraphBase,
                              WordRunFormatting current, bool hasLabel)
        => RenderInline(container.Nodes(), html, paragraphStyle, paragraphBase, current, hasLabel);

    private void RenderInline(IEnumerable<XNode> nodes, StringBuilder html, string? paragraphStyle, WordRunFormatting paragraphBase,
                              WordRunFormatting current, bool hasLabel)
    {
        foreach (XNode node in nodes)
        {
            if (node is XText text)
            {
                string value = text.Value;
                int offset = _paragraphTextOffsets is not null && _paragraphTextOffsets.TryGetValue(text, out int at) ? at : -1;
                // the paragraph drop cap: its first characters (or first word) in the drop-cap span, across text nodes and spans
                if (_dropCapRemaining != 0 && value.Length > 0) value = TakeDropCap(value, current);
                AppendScriptPieces(html, value, current, offset < 0 ? -1 : offset + text.Value.Length - value.Length);
                continue;
            }
            if (node is not XElement element) continue;
            XName name = element.Name;
            if (name.Namespace == TextNs)
            {
                string? computed = Fields is { } fields && ComputedFieldNames.Contains(name.LocalName) ? fields.Compute(element) : null;
                if (computed is not null && name.LocalName is not ("bookmark-ref" or "reference-ref")) { AppendComputedField(html, computed, current); continue; }
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
                        if (_dropCapRemaining != 0)                            // spaces are characters of the paragraph text: counted, or end the word
                        {
                            string rest = TakeDropCap(new string(' ', spaces), current);
                            spaces = rest.Length;
                        }
                        html.Append(' ', spaces);
                        continue;
                    }
                    // a tab, a line break or a note ends the drop cap (LibreOffice txtdrop.cxx SwTextNode::GetDropLen :157-165 stops at
                    // CH_TAB, CH_BREAK and text attributes such as a note's anchor)
                    case "tab": if (_dropCapRemaining != 0) FinishDropCap(); html.Append('\t'); continue;
                    case "line-break": if (_dropCapRemaining != 0) FinishDropCap(); html.Append("<br>"); continue;
                    case "note":
                    {
                        if (_dropCapRemaining != 0) FinishDropCap();
                        AppendNote(element, html);
                        continue;
                    }
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
                        if (computed is not null) AppendComputedField(html, computed, current);
                        else RenderInline(element, html, paragraphStyle, paragraphBase, current, hasLabel);
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
                continue;                                                   // controls, objects: not rendered
            }
            // office:annotation (comments) and unknown namespaces render nothing
        }
    }

    /// <summary>A text box as a placed box (<see cref="AppendTextBox"/>); an image as a picture; an object as a placeholder.</summary>
    private void AppendFrame(XElement frame, StringBuilder html)
    {
        if (frame.Element(DrawNs + "text-box") is { } textBox) { AppendTextBox(frame, textBox, html); return; }
        if (frame.Element(DrawNs + "object") is { } formulaObject && Formulas?.Invoke(formulaObject) is { } formula && WordMathMlHtml.Append(html, formula)) return;
        if (frame.Element(DrawNs + "object") is { } embedded && Charts?.Invoke(embedded) is { } chart) { AppendChart(frame, chart, html); return; }
        if (frame.Element(DrawNs + "image") is { } image) { AppendPicture(frame, image, html); return; }
        html.Append("<span class=\"unsupported\">[image]</span>");
    }

    /// <summary>The formula of a draw:object (its MathML math:math: inline in a flat file, else the sub-document's content.xml, read on demand); null when the object is not a formula.</summary>
    internal Func<XElement, XElement?>? Formulas { get; set; }

    /// <summary>The chart of a draw:object (its sub-document read on demand, <see cref="OdtChart"/>); null when the object is not a chart - its preview picture is shown.</summary>
    internal Func<XElement, DocxHtml.DocxChart?>? Charts { get; set; }

    /// <summary>A chart drawn as SVG in the frame's box, placed like a picture (the same renderer .docx charts use).</summary>
    private void AppendChart(XElement frame, DocxHtml.DocxChart chart, StringBuilder html)
    {
        double width = Points((string?)frame.Attribute(SvgNs + "width")), height = Points((string?)frame.Attribute(SvgNs + "height"));
        WordPictureLayout layout = WithContour(frame, Placement(frame, width, height), width, height, null);
        WordShapeHtml.OpenFrame(html, layout, growsWithText: false);
        html.Append(chart.Svg(width, height));
        WordShapeHtml.CloseFrame(html);
    }

    /// <summary>The bytes of an office:binary-data picture (base64), or null when they do not decode.</summary>
    public static byte[]? BinaryPictureBytes(XElement binary) => DecodeBase64LikeLibreOffice(binary.Value);

    /// <summary>
    /// office:binary-data decoded as LibreOffice decodes it (comphelper/source/misc/base64.cxx Base64::decodeSomeChars):
    /// characters outside the base64 alphabet (white space) are skipped, only complete groups of four are decoded, an
    /// unfinished last group (missing '=' padding, as in LibreOffice's own tdf145542 test files) is dropped, not refused.
    /// Null when nothing decodes.
    /// </summary>
    private static bool IsBase64Character(char c) => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z' or >= '0' and <= '9' or '+' or '/' or '=';

    internal static byte[]? DecodeBase64LikeLibreOffice(string text)
    {
        var kept = new System.Text.StringBuilder(text.Length);
        foreach (char c in text)
            if (IsBase64Character(c)) kept.Append(c);
        kept.Length -= kept.Length % 4;
        if (kept.Length == 0) return null;
        try { return Convert.FromBase64String(kept.ToString()); }
        catch (FormatException) { return null; }
    }

    /// <summary>True when an office:binary-data picture is a format a browser shows (its first bytes' signature).</summary>
    private static bool BinaryPictureFormatIsShown(XElement binary)
    {
        // the first 16 base64 characters are 12 bytes: enough for every signature WordPictureHtml.ExtensionOf knows; decoded the
        // way ReadPicture decodes the whole (white space skipped, an unpadded end tolerated)
        string head = new string(binary.Value.Where(IsBase64Character).Take(16).ToArray());
        return DecodeBase64LikeLibreOffice(head) is { } bytes && WordPictureHtml.ExtensionOf(bytes).Length > 0;
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
        double width = Points((string?)frame.Attribute(SvgNs + "width")), height = Points((string?)frame.Attribute(SvgNs + "height"));
        WordPictureLayout layout = WithContour(frame, Placement(frame, width, height), width, height, null);
        if (pictureId is not null && GraphicProperty(frame, OdtStyleCascade.Fo + "clip") is { } clip && ClipLengths(clip) is { } lengths
            && PictureSizeOf(pictureId) is { } size)
            layout = layout with { Crop = WordPictureCrop.FromLengths(lengths.Left, lengths.Top, lengths.Right, lengths.Bottom, size.W, size.H) };
        WordPictureHtml.Append(html, layout, pictureId);
    }

    /// <summary>
    /// fo:clip "rect(top, right, bottom, left)" (ODF 1.3 part 3 20.179; ODF 1.0 wrote it without commas) in points: lengths cut
    /// from the picture's own size, negative = padding, "auto" = 0 - as LibreOffice reads it (xmloff
    /// XMLClipPropertyHandler.cxx:60-107: comma or space separated, exactly four values, and a value over 400000 hundredths of
    /// a millimetre ignored as nonsense). Null when it is not such a rect.
    /// </summary>
    internal static (double Top, double Right, double Bottom, double Left)? ClipLengths(string clip)
    {
        clip = clip.Trim();
        if (clip.Length <= 6 || !clip.StartsWith("rect(", StringComparison.Ordinal) || clip[clip.Length - 1] != ')') return null;
        string inner = clip.Substring(5, clip.Length - 6);
        string[] parts = inner.Split(new[] { inner.IndexOf(',') >= 0 ? ',' : ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 4) return null;
        var points = new double[4];
        for (int i = 0; i < 4; i++)
        {
            string part = parts[i].Trim();
            if (part == "auto") continue;
            if (OdtStyleCascade.Twips(part) is not { } twips) return null;
            points[i] = Math.Abs(twips) > 400000 * 72 * 20 / 2540.0 ? 0 : twips / 20.0;
        }
        return (points[0], points[1], points[2], points[3]);
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
    /// Text on both sides (parallel / dynamic / biggest) cannot be laid out in CSS (see NOTES.md): the object floats flush to its
    /// side and the text runs along the other (<see cref="WordPictureHtml.AppendPlacement"/>). Without a horizontal position or a one-sided wrap
    /// the object goes to the paragraph's start side (right in a right-to-left paragraph) - our choice, the spec gives no
    /// default. style:vertical-pos from-top with svg:y below a paragraph / character anchor (20.397 / 20.398) is added to the
    /// top gap (a float cannot move up; page / frame relations are ignored - the flow view has no page).
    /// </summary>
    private WordPictureLayout Placement(XElement element, double width, double height)
        => Placement(element, width, height, Points((string?)element.Attribute(SvgNs + "x")), Points((string?)element.Attribute(SvgNs + "y")));

    /// <param name="x">The object's left edge from its anchor (a frame's svg:x; a shape's snap rectangle).</param>
    /// <param name="y">The object's top edge from its anchor.</param>
    private WordPictureLayout Placement(XElement element, double width, double height, double x, double y)
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
        double offset = position is "from-left" or "from-inside" ? Math.Max(0, x) : 0;
        double all = Points(Property(OdtStyleCascade.Fo, "margin"));
        double Gap(string name) => Property(OdtStyleCascade.Fo, "margin-" + name) is { } value ? Points(value) : all;
        double top = Gap("top");
        if (wrap != WordPictureWrap.Inline && Property(OdtStyleCascade.Style, "vertical-pos") == "from-top"
            && Property(OdtStyleCascade.Style, "vertical-rel") is null or "paragraph" or "paragraph-content" or "char" or "line")
            top += Math.Max(0, y);
        WordPictureTextSides sides = wrapValue switch { "left" => WordPictureTextSides.LeftOnly, "right" => WordPictureTextSides.RightOnly, _ => WordPictureTextSides.Both };
        return new WordPictureLayout(width, height, wrap, side, top, Gap("right"), Gap("bottom"), Gap("left"), offset, sides,
                                     ColumnWidthPt: ColumnWidthPt, ColumnLeftPt: ColumnLeftPt,
                                     // ODF 1.3 part 3 20.290 style:horizontal-rel: paragraph / page text area = the column; page / its start margin = the page edge; its end margin = the right margin area
                                     OffsetOrigin: Property(OdtStyleCascade.Style, "horizontal-rel") switch
                                     {
                                         null or "paragraph" or "paragraph-content" or "page-content" => WordPictureOffsetOrigin.Column,
                                         "page" or "page-start-margin" => WordPictureOffsetOrigin.Page,
                                         "page-end-margin" => WordPictureOffsetOrigin.EndMargin,
                                         _ => WordPictureOffsetOrigin.Other,
                                     });
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
        WordPictureLayout layout = WithContour(frame, Placement(frame, width, height), width, height, null);
        var box = new StringBuilder();
        WordPictureHtml.OpenTextBox(box, layout,
                                    new WordTextBoxStyle(Padding(frame, "left"), Padding(frame, "top"), Padding(frame, "right"), Padding(frame, "bottom"),
                                                         borderColor, borderWidth, borderDash, FillCss(frame, isTextBox: true), 0, null, grows));
        // draw:textarea-vertical-align (Writer frames: TextVerticalAdjust, sdpropls.cxx:149): the box a flex column
        string? vertical = GraphicProperty(frame, OdtStyleCascade.Draw + "textarea-vertical-align");
        bool painted = GraphicProperty(frame, OdtStyleCascade.Draw + "fill") is "gradient" or "hatch" or "bitmap";
        string extra = (painted ? "position:relative;isolation:isolate;" : "")
                     + (vertical is "middle" or "bottom" or "justify"
                        ? (layout.Wrap == WordPictureWrap.Inline ? "display:inline-flex;" : "display:flex;") + "flex-direction:column;justify-content:"
                          + (vertical == "middle" ? "center" : vertical == "bottom" ? "flex-end" : "space-between") + ";"
                        : "");
        int styleAt = box.ToString().LastIndexOf("\">", StringComparison.Ordinal);
        html.Append(box.ToString(0, styleAt)).Append(extra).Append(box.ToString(styleAt, box.Length - styleAt));
        AppendFrameFill(frame, html, width, height);
        AppendOwnParagraphs(textBox, html);
        WordPictureHtml.CloseTextBox(html);
    }

    /// <summary>The blocks of a text box or shape (paragraphs, lists, tables), rendered as block spans by the normal renderers.</summary>
    private void AppendOwnParagraphs(XElement container, StringBuilder html)
    {
        // _isInsideLink is kept: a box inside a text:a is inside that <a> in the HTML too (no link inside a link)
        string? savedLabel = _listLabel; int savedLevel = _listLevel; string? savedPrevious = _previousStyleKey; bool savedLink = _isInsideLink;
        _listLabel = null; _listLevel = -1; _previousStyleKey = null;
        _textBoxDepth++;
        foreach (XElement child in container.Elements()) RenderBoxBlock(child, html);
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
        return null;                                    // gradient / hatch / bitmap: an SVG layer under the text (AppendFrameFill)
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
            int rowHeight = _styles.RowHeightTwips((string?)row.Attribute(TableNs + "style-name"));
            WordTableRowHeight.AppendRowOpen(html, rowHeight);
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
                var content = new StringBuilder();
                foreach (XElement child in cell.Elements()) RenderBlock(child, content, isTopLevel: false);
                WordTableRowHeight.AppendCellContent(html, content.ToString(), rowHeight, (int?)cell.Attribute(TableNs + "number-rows-spanned") ?? 1);
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
