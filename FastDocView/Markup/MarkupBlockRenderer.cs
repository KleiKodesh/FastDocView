using System.Text;
using FastDocView.Html;

namespace FastDocView.Markup;

/// <summary>
/// One block of HTML-ish text → standalone, safe HTML. The block is parsed after the start tags that
/// were open at its start (from the index), with the same tag rules as <see cref="MarkupBlockScanner"/>:
/// a '<' that does not start a tag is text (escaped), comments and &lt;!...&gt; / &lt;?...?&gt; are dropped,
/// script / style / embedded objects are dropped, event-handler attributes and javascript: URLs are
/// removed, images become a placeholder, and every element left open is closed at the end.
/// </summary>
public static class MarkupBlockRenderer
{
    private static readonly HashSet<string> DroppedTags = new(StringComparer.Ordinal)
        { "html", "head", "body", "meta", "link", "base", "iframe", "frame", "frameset", "object", "embed", "applet", "form", "input", "button", "select", "option" };
    private static readonly HashSet<string> DroppedWithContent = new(StringComparer.Ordinal) { "script", "style", "title" };
    private static readonly HashSet<string> VoidElements = new(StringComparer.Ordinal)
        { "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr" };

    /// <summary>Plain text only: headings are bold in the Torat Emet size steps (the shared stylesheet resets them to inherit); no extra line after them (the page's zero-width-space ::after after a block-level heading).</summary>
    public const string PlainTextHeadingStylesheet =
        "\n.word-document .text-line > h1, .word-document .text-line > h2, .word-document .text-line > h3, .word-document .text-line > h4, .word-document .text-line > h5, .word-document .text-line > h6 { margin: .4em 0 .1em; border: 0; padding: 0; font-weight: 700; font-size: 1em; }"
        + "\n.word-document .text-line:has(> h1, > h2, > h3, > h4, > h5, > h6)::after { content: none !important; }"
        + "\n.word-document .text-line > h1 { font-size: 1.6em; }"
        + "\n.word-document .text-line > h2 { font-size: 1.35em; }"
        + "\n.word-document .text-line > h3 { font-size: 1.15em; }\n";

    public const string ImagePlaceholder = "<span class=\"image-placeholder\">[image]</span>";

    public static string RenderLine(string openTags, string line)
    {
        int end = line.Length;
        while (end > 0 && (line[end - 1] is '\n' or '\r' or ' ' or '\t' or '\u00A0' or '\u200B')) end--;   // blank lines folded into this block, trailing white space
        string body = Render(openTags, line.Substring(0, end));
        return "<div class=\"text-line\">" + (IsVisiblyEmpty(body) ? "<br>" : body) + "</div>";
    }

    private static bool IsVisiblyEmpty(string html)
    {
        foreach (char c in html) if (c != '<') { if (!char.IsWhiteSpace(c)) return false; } else return false;
        return true;
    }

    /// <param name="drawsPictures">A web archive: its &lt;img&gt; parts are read on request. Plain text has no parts, so an &lt;img&gt; typed in it stays the placeholder.</param>
    /// <param name="stylesheet">A web archive's head sheet (<see cref="WebArchiveStylesheet"/>): its rules are resolved per element into the element's inline style.</param>
    public static string Render(string openTags, string content, bool drawsPictures = false, WebArchiveStylesheet? stylesheet = null)
        => Render(openTags, content, drawsPictures, stylesheet, NextBlock.Unknown, shown: null, ends: null);

    /// <summary>What the next block's open-tag stack says about the elements open at this block's end.</summary>
    public readonly record struct NextBlock(int Depth)
    {
        /// <summary>Not known yet (the index is still being built): nothing ends here.</summary>
        public static readonly NextBlock Unknown = new(int.MinValue);
        /// <summary>This is the last block: everything open ends here.</summary>
        public static readonly NextBlock None = new(-1);
        public bool IsKnown => Depth != int.MinValue;
    }

    /// <summary>
    /// A block's open tags with the index pass's selector state (<see cref="MarkupBlockScanner.CurrentStructure"/>): body's as a
    /// leading &lt;fdv-root&gt; entry, each open element's as an attribute right after its name. The renderer reads them back
    /// (<see cref="Render(string, string, bool, WebArchiveStylesheet?, NextBlock, StringBuilder?)"/>) and never writes them out.
    /// </summary>
    public static string WithStructure(string openTags, string structure)
    {
        if (structure.Length == 0) return openTags;
        string[] states = structure.Split('|');
        var text = new StringBuilder(openTags.Length + structure.Length + 64);
        text.Append('<').Append(WebArchiveMatch.RootTagName).Append(' ').Append(WebArchiveMatch.StateAttribute).Append("=\"").Append(states[0]).Append("\">");
        int i = 0, entry = 1;
        while (i < openTags.Length)
        {
            int end = openTags[i] == '<' ? FindTagEnd(openTags, i) : -1;
            if (end < 0) { text.Append(openTags, i, openTags.Length - i); break; }
            string tag = openTags.Substring(i, end + 1 - i);
            string name = MarkupBlockScanner.ParseTag(tag).name;
            if (entry < states.Length && states[entry].Length > 0)
            {
                int at = Math.Min(tag.Length - 1, 1 + name.Length);
                tag = tag.Substring(0, at) + " " + WebArchiveMatch.StateAttribute + "=\"" + states[entry] + "\"" + tag.Substring(at);
            }
            text.Append(tag);
            entry++;
            i = end + 1;
        }
        return text.ToString();
    }

    /// <summary>The next block as <see cref="Render(string, string, bool, WebArchiveStylesheet?, NextBlock, StringBuilder?)"/> needs it: how many drawn elements its stack reopens.</summary>
    public static NextBlock NextBlockFrom(string nextOpenTags)
    {
        int depth = 0, i = 0;
        while (i < nextOpenTags.Length)
        {
            int start = nextOpenTags.IndexOf('<', i);
            if (start < 0) break;
            int end = FindTagEnd(nextOpenTags, start);
            if (end < 0) break;
            (string name, bool isEnd, _) = MarkupBlockScanner.ParseTag(nextOpenTags.Substring(start, end + 1 - start));
            if (!isEnd && name != WebArchiveMatch.RootTagName && !DroppedTags.Contains(name) && !VoidElements.Contains(name)
                && !DroppedWithContent.Contains(name) && !MarkupBlockScanner.RawTextElements.Contains(name)) depth++;
            i = end + 1;
        }
        return new NextBlock(depth);
    }

    /// <summary>One element the renderer has open (or the synthesized html / body): its state for the sheet and where its tag was written.</summary>
    private sealed class Node(string name, WebArchiveElement element, WebArchiveMatch? state, Node? parent, bool isOutput, bool isReopened, string? path = null, bool isPathKnown = true)
    {
        /// <summary>The element's path as the index pass keys its end (<see cref="WebArchiveElementEnds"/>): "" for body, null where unknown
        /// (a reopened element without its stored state, or one the index pass's implied end tags place elsewhere): unknown, never another element's.</summary>
        public readonly string? Path = path ?? (isPathKnown && state is not null && parent?.Path is { } parentPath ? parentPath + "/" + state.Index.ToString(System.Globalization.CultureInfo.InvariantCulture) : null);
        public readonly string Name = name;
        public readonly WebArchiveElement Element = element;
        public readonly WebArchiveMatch? State = state;
        public readonly Node? Parent = parent;
        public readonly bool IsOutput = isOutput, IsReopened = isReopened;
        public bool IsHidden;
        public bool IsClosed;                     // its end is in this block (explicit, or at the block's end when it does not go on)
        public int TagStart = -1, TagLength;      // its start tag in the output (deferred elements are redrawn there)
        public string Tag = "";                   // the checked start tag, without the sheet's style
        public WebArchiveStylesheet.ElementStyle Style;
    }

    /// <param name="next">The next block's stack depth (<see cref="NextBlock"/>): elements open at this block's end that the next block does not reopen end here.</param>
    /// <param name="shown">Receives what the block shows, as HTML text for the search (<see cref="WordDocumentSearchMatcher.FeedHtml(string)"/>):
    /// elements the sheet hides (display:none) left out, ::before / ::after strings put in; one path with the drawn HTML.</param>
    /// <param name="ends">The ends of the elements that span blocks (the index pass's): later-sibling rules of an element whose parent, or
    /// which itself, goes on past this block; null: only what this block shows is known.</param>
    public static string Render(string openTags, string content, bool drawsPictures, WebArchiveStylesheet? stylesheet, NextBlock next, StringBuilder? shown)
        => Render(openTags, content, drawsPictures, stylesheet, next, shown, ends: null);

    internal static string Render(string openTags, string content, bool drawsPictures, WebArchiveStylesheet? stylesheet, NextBlock next, StringBuilder? shown,
                                  WebArchiveElementEnds? ends)
        => Render(openTags, content, drawsPictures, stylesheet, next, shown, ends, out _);

    /// <param name="usedUnknownEnds">Some start tag was drawn without an end its sheet asks for (a later-sibling pseudo-class whose parent's
    /// or own end is not recorded, an ::after on an element whose end is not known yet): on a partial index the block is provisional
    /// (<see cref="WordDocumentBlockHtml.IsProvisional"/>) and is drawn again once the index is final.</param>
    internal static string Render(string openTags, string content, bool drawsPictures, WebArchiveStylesheet? stylesheet, NextBlock next, StringBuilder? shown,
                                  WebArchiveElementEnds? ends, out bool usedUnknownEnds)
    {
        usedUnknownEnds = false;
        if (stylesheet is { IsEmpty: true }) stylesheet = null;
        WebArchiveMatch? rootState = null;
        if (openTags.StartsWith("<" + WebArchiveMatch.RootTagName, StringComparison.Ordinal))
        {
            int rootEnd = openTags.IndexOf('>');
            if (stylesheet is not null && AttributeValue(openTags.Substring(0, rootEnd + 1), WebArchiveMatch.StateAttribute) is { } rootText) rootState = WebArchiveMatch.Deserialize(rootText, null);
            openTags = openTags.Substring(rootEnd + 1);
        }
        string text = openTags + content;
        var html = new StringBuilder(text.Length + 32);
        // the open elements as the sheet sees them, under the document's html / body (dropped from the output, still ancestors)
        WebArchiveElement htmlElement = new("html", null, [], false, "<html>"), bodyElement = new("body", null, [], false, "<body>");
        WebArchiveMatch? htmlState = stylesheet?.Open(null, htmlElement);
        var nodes = new List<Node> { new("html", htmlElement, htmlState, null, false, true) };
        nodes.Add(new Node("body", bodyElement, stylesheet is null ? null : rootState ?? stylesheet.Open(htmlState, bodyElement), nodes[0], false, true, path: ""));
        List<Node>? deferred = null;
        int hiddenDepth = 0;                       // open elements the sheet hides (their text is not shown)

        void AppendShown(string value) { if (shown is not null && hiddenDepth == 0) shown.Append(value); }
        void Close(Node node, bool endsHere)
        {
            node.IsClosed = endsHere;
            if (node.IsHidden) hiddenDepth--;
            if (endsHere && node.Style.AfterText.Length > 0) AppendShown(System.Net.WebUtility.HtmlEncode(node.Style.AfterText));
        }

        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (c != '<')
            {
                int next0 = text.IndexOf('<', i); if (next0 < 0) next0 = text.Length;
                html.Append(text, i, next0 - i);
                if (shown is not null) AppendShown(text.Substring(i, next0 - i));
                if (nodes[nodes.Count - 1].State is { } textParent && i >= openTags.Length) textParent.HasContent = true;
                i = next0;
                continue;
            }
            char d = i + 1 < text.Length ? text[i + 1] : '\0';
            if (d == '!' && string.CompareOrdinal(text, i, "<!--", 0, 4) == 0)
            {
                int close = text.IndexOf("-->", i + 4, Math.Min(text.Length - i - 4, MarkupBlockScanner.MaximumCommentBytes + 3), StringComparison.Ordinal);
                if (close < 0 && text.Length - i - 4 > MarkupBlockScanner.MaximumCommentBytes) { html.Append("&lt;"); AppendShown("&lt;"); i++; continue; }   // the scanner gave up on it: text
                i = close < 0 ? text.Length : close + 3;
                continue;
            }
            if (d == '!' || d == '?')
            {
                int close = text.IndexOf('>', i + 2, Math.Min(text.Length - i - 2, MarkupBlockScanner.MaximumCommentBytes + 1));
                if (close < 0 && text.Length - i - 2 > MarkupBlockScanner.MaximumCommentBytes) { html.Append("&lt;"); AppendShown("&lt;"); i++; continue; }
                i = close < 0 ? text.Length : close + 1;
                continue;
            }
            bool isEnd = d == '/';
            char first = isEnd ? (i + 2 < text.Length ? text[i + 2] : '\0') : d;
            int tagEnd = IsAsciiLetter(first) ? FindTagEnd(text, i) : -1;
            if (tagEnd < 0) { html.Append("&lt;"); AppendShown("&lt;"); i++; continue; }
            bool isReopened = i < openTags.Length;
            string tag = text.Substring(i, tagEnd + 1 - i);
            i = tagEnd + 1;
            (string name, _, bool isSelfClosing) = MarkupBlockScanner.ParseTag(tag);
            if (isEnd)
            {
                int at = -1;
                for (int k = nodes.Count - 1; k >= 2; k--) if (nodes[k].Name == name) { at = k; break; }
                if (at < 0) continue;
                for (int k = nodes.Count - 1; k >= at; k--)
                {
                    if (nodes[k].IsOutput) { html.Append("</").Append(nodes[k].Name).Append('>'); AppendShown("</" + nodes[k].Name + ">"); }
                    Close(nodes[k], endsHere: true);
                }
                nodes.RemoveRange(at, nodes.Count - at);
                continue;
            }
            // every start tag in the body is an element child (drawn or not), as the index pass counts them
            Node parent = nodes[nodes.Count - 1];
            if (!isReopened && parent.State is not null) parent.State.HasContent = true;
            WebArchiveElement? element = null;
            WebArchiveMatch? state = null;
            bool isPathKnown = true;
            if (ends is not null) for (int k = nodes.Count - 1; k >= 2 && isPathKnown; k--) isPathKnown = !ImpliesEnd(nodes[k].Name, name);
            if (stylesheet is not null && name is not ("html" or "head" or "body"))
            {
                element = WebArchiveElement.FromTag(name, tag);
                if (isReopened && AttributeValue(tag, WebArchiveMatch.StateAttribute) is { } stateText && WebArchiveMatch.Deserialize(stateText, parent.State) is { } stored)
                    state = stored;                                       // the index pass's state: the element's siblings before this block
                else { state = stylesheet.Open(parent.State, element); if (isReopened) isPathKnown = false; }
            }
            if (DroppedWithContent.Contains(name) || MarkupBlockScanner.RawTextElements.Contains(name))
            {
                int close = IndexOfEndTag(text, i, name);
                string raw = text.Substring(i, (close < 0 ? text.Length : close) - i);
                if (!DroppedWithContent.Contains(name))
                {
                    html.Append("<pre>").Append(Escape(raw)).Append("</pre>");    // textarea / xmp: show the text
                    AppendShown("<pre>" + Escape(raw) + "</pre>");
                }
                i = close < 0 ? text.Length : Math.Max(close, text.IndexOf('>', close) + 1);
                continue;
            }
            // Word's web archive writes each picture twice: VML (v:imagedata) and an <img> fallback inside <![if !vml]>, which is
            // what a browser shows (it has no VML). The <img> is drawn through the picture pipeline; the VML copy is skipped.
            if (name == "v:imagedata") continue;
            if (name == "img") { if (drawsPictures) AppendPicture(html, tag); else html.Append(ImagePlaceholder); continue; }
            bool isOpen = !VoidElements.Contains(name) && !isSelfClosing;
            bool isDrawn = !DroppedTags.Contains(name);
            var node = new Node(name, element ?? new WebArchiveElement(name, null, [], false, tag), state, parent, isDrawn, isReopened, isPathKnown: isPathKnown);
            if (isDrawn)
            {
                node.Tag = Sanitize(tag, name);
                node.Style = state is null ? WebArchiveStylesheet.ElementStyle.None : stylesheet!.StyleFor(state, null);
                string? inline = AttributeValue(node.Tag, "style");
                string? inlineDisplay = inline is null ? null : WebArchiveStylesheet.InlineDisplay(System.Net.WebUtility.HtmlDecode(inline));
                node.IsHidden = node.Style.IsDisplayImportant ? node.Style.IsHidden : inlineDisplay is not null ? inlineDisplay.Equals("none", StringComparison.OrdinalIgnoreCase) : node.Style.IsHidden;
                node.TagStart = html.Length;
                html.Append(WithStyle(node.Tag, node.Style.Css));
                node.TagLength = html.Length - node.TagStart;
                if (node.Style.IsDeferred) (deferred ??= []).Add(node);
                if (node.IsHidden) hiddenDepth++;
                AppendShown(node.Tag);
                if (!isReopened && node.Style.BeforeText.Length > 0) AppendShown(System.Net.WebUtility.HtmlEncode(node.Style.BeforeText));
                if (node.IsHidden && !isOpen) hiddenDepth--;
            }
            if (isOpen) nodes.Add(node);
            else node.IsClosed = true;
        }
        // the elements still open end here unless the next block reopens them (its stack holds the first ones)
        int drawnDepth = 0;
        for (int k = 2; k < nodes.Count; k++)
        {
            bool goesOn = !next.IsKnown || (next.Depth >= 0 && drawnDepth < next.Depth);
            if (nodes[k].IsOutput) drawnDepth++;
            nodes[k].IsClosed = !goesOn;
        }
        for (int k = nodes.Count - 1; k >= 2; k--)
        {
            if (nodes[k].IsOutput) html.Append("</").Append(nodes[k].Name).Append('>');
            if (nodes[k].IsClosed && nodes[k].Style.AfterText.Length > 0) AppendShown(System.Net.WebUtility.HtmlEncode(nodes[k].Style.AfterText));
        }
        if (nodes.Count > 1) nodes[1].IsClosed = next.IsKnown && next.Depth < 0;    // body ends with the last block
        if (deferred is not null) usedUnknownEnds = Redraw(html, deferred, stylesheet!, ends, next);
        return html.ToString();
    }

    /// <summary>
    /// The start tags drawn before their element's end was known, drawn again now: the later-sibling rules that hold
    /// (:last-child...), and the pseudo-elements on the part of the element this block shows (::before, ::first-letter and
    /// ::first-line where it starts, ::after where it ends). Returns whether an end the sheet asks for was not known (see usedUnknownEnds).
    /// </summary>
    private static bool Redraw(StringBuilder html, List<Node> deferred, WebArchiveStylesheet stylesheet, WebArchiveElementEnds? ends, NextBlock next)
    {
        bool usedUnknownEnds = false;
        for (int k = deferred.Count - 1; k >= 0; k--)
        {
            Node node = deferred[k];
            WebArchiveStylesheet.ElementStyle style = node.Style;
            if (style.HasLaterSiblingRules)
            {
                if (LaterFacts(node, ends) is { } facts) style = stylesheet.StyleFor(node.State!, facts);
                else usedUnknownEnds = true;
            }
            if (!node.IsClosed && !next.IsKnown && style.AfterPseudoIds.Length > 0) usedUnknownEnds = true;
            string ids = string.Join(" ", new[] { node.IsReopened ? "" : style.StartPseudoIds, node.IsClosed ? style.AfterPseudoIds : "" }.Where(part => part.Length > 0));
            string tag = WithStyle(node.Tag, style.Css);
            if (ids.Length > 0)
            {
                int end = tag.EndsWith("/>", StringComparison.Ordinal) ? tag.Length - 2 : tag.Length - 1;
                tag = tag.Substring(0, end) + " data-wa-pe=\"" + ids + "\"" + tag.Substring(end);
            }
            html.Remove(node.TagStart, node.TagLength).Insert(node.TagStart, tag);
        }
        return usedUnknownEnds;
    }

    /// <summary>
    /// An element's later siblings and emptiness: from this block when both it and its parent's children end here, else from the
    /// index pass's ends of the ones that go on past it (<see cref="WebArchiveElementEnds"/>); null when neither knows.
    /// </summary>
    private static WebArchiveLaterFacts? LaterFacts(Node node, WebArchiveElementEnds? ends)
    {
        if (node.State is not { } state || node.Parent is not { State: { } parent } parentNode) return null;
        int siblings, ofType;
        if (parentNode.IsClosed) { siblings = parent.Children; ofType = parent.TypeCount(node.Name); }
        else if (parentNode.Path is { } parentPath && ends?.Find(parentPath, node.Name) is { } parentEnd) { siblings = parentEnd.Children; ofType = parentEnd.OfType; }
        else return null;
        bool isEmpty;
        // an element that spans blocks: its own end says whether it has content anywhere (this block sees only its part)
        if ((!node.IsClosed || node.IsReopened) && node.Path is { } path && ends?.Find(path, node.Name) is { } own) isEmpty = !own.HasContent && own.Children == 0;
        else if (node.IsClosed) isEmpty = !state.HasContent;
        else return null;
        return new WebArchiveLaterFacts(siblings - state.Index, ofType - state.TypeIndex, isEmpty);
    }

    /// <summary>
    /// A start tag the index pass reads as ending the open element first (MarkupBlockScanner.PopForStart: a p ends a p, li an li,
    /// a cell a cell...), while the renderer nests it: its path would name another element.
    /// </summary>
    private static bool ImpliesEnd(string open, string name)
        => (open == "p" && (name is "li" or "dt" or "dd" || MarkupBlockScanner.ClosesOpenParagraph.Contains(name)))
           || (open == name && name is "li" or "option" or "tr")
           || (open is "dt" or "dd" && name is "dt" or "dd")
           || (open is "td" or "th" && name is "td" or "th" or "tr");

    /// <summary>The prefix of a web-archive picture id; the rest is the &lt;img&gt;'s src (see MarkupWindowedDocument.ReadPicture).</summary>
    public const string WebArchivePicturePrefix = "mh:";

    /// <summary>An &lt;img&gt; as a picture read on request: its src names the archive part, width / height (CSS px) keep the box.</summary>
    private static void AppendPicture(StringBuilder html, string tag)
    {
        string? source = AttributeValue(tag, "src");
        if (string.IsNullOrWhiteSpace(source) || !IsSafeValue(source!)) { html.Append(ImagePlaceholder); return; }
        double Points(string attribute) =>
            double.TryParse(AttributeValue(tag, attribute), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double px) && px > 0 ? px * 72 / 96 : 0;
        var layout = new WordPictureLayout(Points("width"), Points("height"), WordPictureWrap.Inline, WordPictureSide.Left, 0, 0, 0, 0);
        WordPictureHtml.Append(html, layout, WebArchivePicturePrefix + System.Net.WebUtility.HtmlDecode(source!.Trim()));
    }


    /// <summary>
    /// The tag with one style="": the sheet's declarations for the element first, its own style after them (it wins, as in the
    /// cascade), both through the sheet's checks (<see cref="WebArchiveStylesheet.SanitizeInline"/>) - nothing passes through raw.
    /// </summary>
    private static string WithStyle(string tag, string sheetCss)
    {
        string? inline = AttributeValue(tag, "style", out int attributeStart, out int attributeEnd);
        if (inline is null && sheetCss.Length == 0) return tag;
        string css = sheetCss + (inline is null ? "" : WebArchiveStylesheet.SanitizeInline(System.Net.WebUtility.HtmlDecode(inline)));
        while (attributeStart >= 0)                 // every style="" goes (a browser reads the first; a second must not take its place)
        {
            tag = tag.Remove(attributeStart, attributeEnd - attributeStart).Insert(attributeStart, " ");   // keep the name apart from the next attribute
            AttributeValue(tag, "style", out attributeStart, out attributeEnd);
        }
        if (css.Length == 0) return tag;
        int end = tag.EndsWith("/>", StringComparison.Ordinal) ? tag.Length - 2 : tag.Length - 1;
        while (end > 0 && char.IsWhiteSpace(tag[end - 1])) end--;
        return tag.Substring(0, end) + " style=\"" + css.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;") + "\"" + tag.Substring(end).TrimStart();
    }

    /// <summary>One attribute's value in a start tag (quoted or bare, names case-insensitive), or null - the same rules as <see cref="Sanitize"/>.</summary>
    internal static string? AttributeValue(string tag, string wanted) => AttributeValue(tag, wanted, out _, out _);

    /// <param name="attributeStart">Where the attribute starts (the white space before its name), <paramref name="attributeEnd"/> just after its value.</param>
    private static string? AttributeValue(string tag, string wanted, out int attributeStart, out int attributeEnd)
    {
        attributeStart = attributeEnd = -1;
        int i = 1;
        while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '>' && tag[i] != '/') i++;   // past the element name
        while (i < tag.Length)
        {
            int gapStart = i;
            while (i < tag.Length && (char.IsWhiteSpace(tag[i]) || tag[i] == '/')) i++;
            if (i >= tag.Length || tag[i] == '>') return null;
            int nameStart = i;
            while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '=' && tag[i] != '>' && tag[i] != '/') i++;
            int nameLength = i - nameStart;
            // the name compared in place (no copy per attribute); the value is copied only for the wanted one
            bool isWanted = nameLength == wanted.Length && string.Compare(tag, nameStart, wanted, 0, nameLength, StringComparison.OrdinalIgnoreCase) == 0;
            int valueStart = -1, valueLength = 0;
            int afterName = i;
            while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
            if (i < tag.Length && tag[i] == '=')
            {
                i++;
                while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
                if (i < tag.Length && (tag[i] == '"' || tag[i] == '\''))
                {
                    char quote = tag[i];
                    int close = tag.IndexOf(quote, i + 1);
                    if (close < 0) close = tag.Length - 1;
                    valueStart = i + 1; valueLength = close - i - 1;
                    i = close + 1;
                }
                else
                {
                    int start = i;
                    while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '>') i++;
                    valueStart = start; valueLength = i - start;
                }
            }
            else i = afterName;
            if (nameLength == 0) { i++; continue; }
            if (isWanted) { attributeStart = gapStart; attributeEnd = Math.Min(i, tag.Length - 1); return valueStart < 0 ? "" : tag.Substring(valueStart, valueLength); }
        }
        return null;
    }

    private static bool IsAsciiLetter(char c) => (c | 0x20) >= 'a' && (c | 0x20) <= 'z';

    /// <summary>The '>' closing the tag at <paramref name="start"/> (quotes after '=' respected), or -1 — the scanner's rule.</summary>
    private static int FindTagEnd(string text, int start)
    {
        bool quoted = false; char quote = '\0', lastNonSpace = 'a';
        int limit = Math.Min(text.Length, start + MarkupBlockScanner.MaximumTagUnits);
        for (int i = start + 1; i < limit; i++)
        {
            char c = text[i];
            if (quoted) { if (c == quote) quoted = false; }
            else if (c == '>') return i;
            else if ((c == '"' || c == '\'') && lastNonSpace == '=') { quoted = true; quote = c; }
            if (c > ' ') lastNonSpace = c;
        }
        return -1;
    }

    private static int IndexOfEndTag(string text, int from, string name)
    {
        string pattern = "</" + name;
        return text.IndexOf(pattern, from, StringComparison.OrdinalIgnoreCase);
    }

    private static string Escape(string text) => text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");

    /// <summary>Whitespace, quote or '/' followed by "on" — a possible event-handler attribute (cheap pre-check).</summary>
    private static bool HasEventAttribute(string tag)
    {
        for (int i = tag.IndexOf("on", StringComparison.OrdinalIgnoreCase); i > 0; i = tag.IndexOf("on", i + 2, StringComparison.OrdinalIgnoreCase))
        {
            char before = tag[i - 1];
            if (char.IsWhiteSpace(before) || before == '"' || before == '\'' || before == '/') return true;
        }
        return false;
    }

    private static readonly HashSet<string> UrlAttributes = new(StringComparer.Ordinal)
        { "href", "src", "action", "formaction", "xlink:href", "data", "background", "lowsrc", "dynsrc", "poster", "cite", "longdesc", "to", "from", "values", "by", "srcset", "codebase" };

    /// <summary>
    /// False for a value that resolves to a script-capable URL: entities decoded (browsers decode them in
    /// attributes), whitespace and controls removed, then the scheme checked: http(s), mailto, file, ftp,
    /// cid, mid and data:image (not svg) pass, as does any value without a URL scheme; anything else is dropped.
    /// </summary>
    private static readonly char[] PathStarts = ['/', '?', '#'];

    private static bool IsSafeValue(string value)
    {
        string decoded = System.Net.WebUtility.HtmlDecode(value);
        string compact = new string(decoded.Where(ch => !char.IsWhiteSpace(ch) && !char.IsControl(ch)).ToArray()).ToLowerInvariant();
        // A character reference the decoder left alone before the path / query / fragment may still be one to a browser:
        // HTML5 named references (&colon; &Tab; &NewLine;) and numeric ones without ';' (&#58) hide the scheme's ':' or split
        // its letters (WHATWG HTML 13.2.5.72-80: attribute values decode the full named table and unterminated numerics).
        int pathStart = compact.IndexOfAny(PathStarts);
        if (compact.IndexOf('&') is int ampersand && ampersand >= 0 && (pathStart < 0 || ampersand < pathStart)) return false;
        int colon = compact.IndexOf(':');
        if (colon <= 0) return true;
        string scheme = compact.Substring(0, colon);
        if (!char.IsLetter(scheme[0]) || scheme.Any(ch => !(ch is >= 'a' and <= 'z' || char.IsDigit(ch) || ch is '+' or '-' or '.'))) return true;   // not a URL scheme
        if (scheme is "http" or "https" or "mailto" or "file" or "ftp" or "cid" or "mid") return true;
        if (scheme == "data") return compact.StartsWith("data:image/", StringComparison.Ordinal) && !compact.StartsWith("data:image/svg", StringComparison.Ordinal);
        return false;
    }

    /// <summary>A link into a library app (Otzaria, Zayit, KitveiHakodesh's own) - the schemes the Word readers mark as app links.</summary>
    private static bool IsAppLink(string value) => WordLibraryAppLink.IsLibraryAppLink(System.Net.WebUtility.HtmlDecode(value));

    /// <summary>True when the tag's href stays in the document (#anchor) or has no scheme.</summary>
    private static bool IsInternalHref(string tag)
    {
        int at = tag.IndexOf("href", StringComparison.OrdinalIgnoreCase);
        if (at < 0) return true;
        int eq = tag.IndexOf('=', at);
        if (eq < 0) return true;
        int i = eq + 1;
        while (i < tag.Length && (char.IsWhiteSpace(tag[i]) || tag[i] == '"' || tag[i] == '\'')) i++;
        return i < tag.Length && tag[i] == '#';
    }

    /// <summary>The tag as written, minus event handlers (on*) and script URLs; a link gets the reader's link markup.</summary>
    private static string Sanitize(string tag, string name)
    {
        if (name != "a" && !HasEventAttribute(tag) && tag.IndexOf(':') < 0 && tag.IndexOf('&') < 0 && tag.IndexOf("data-fdv-s", StringComparison.OrdinalIgnoreCase) < 0 && tag.IndexOf("data-wa-pe", StringComparison.OrdinalIgnoreCase) < 0) return tag;      // no handler, no scheme, no entity: nothing to remove
        var result = new StringBuilder(tag.Length);
        bool isAppLink = false, hasHref = false;
        result.Append('<').Append(name);
        int i = 1 + name.Length;
        while (i < tag.Length)
        {
            while (i < tag.Length && (char.IsWhiteSpace(tag[i]) || tag[i] == '/')) i++;
            if (i >= tag.Length || tag[i] == '>') break;
            int nameStart = i;
            while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '=' && tag[i] != '>' && tag[i] != '/') i++;
            string attribute = tag.Substring(nameStart, i - nameStart);
            string value = "";
            int valueStart = i;
            while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
            if (i < tag.Length && tag[i] == '=')
            {
                i++;
                while (i < tag.Length && char.IsWhiteSpace(tag[i])) i++;
                if (i < tag.Length && (tag[i] == '"' || tag[i] == '\''))
                {
                    char quote = tag[i];
                    int close = tag.IndexOf(quote, i + 1);
                    if (close < 0) close = tag.Length - 1;
                    value = tag.Substring(i + 1, close - i - 1);
                    i = close + 1;
                }
                else
                {
                    int start = i;
                    while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '>') i++;
                    value = tag.Substring(start, i - start);
                }
            }
            else i = valueStart;
            if (attribute.Length == 0) { i++; continue; }
            if (attribute.StartsWith("on", StringComparison.OrdinalIgnoreCase)) continue;
            if (attribute.Equals(WebArchiveMatch.StateAttribute, StringComparison.OrdinalIgnoreCase) || attribute.Equals("data-wa-pe", StringComparison.OrdinalIgnoreCase)) continue;   // the reader's own markup, never the document's
            if (name == "a" && attribute.Equals("href", StringComparison.OrdinalIgnoreCase) && IsAppLink(value)) isAppLink = true;
            else if (UrlAttributes.Contains(attribute.ToLowerInvariant()) && !IsSafeValue(value)) continue;
            if (name == "a" && attribute.ToLowerInvariant() is "class" or "target" or "rel" or "data-app-link") continue;   // the reader's own link markup replaces them
            if (name == "a" && attribute.Equals("href", StringComparison.OrdinalIgnoreCase)) hasHref = true;
            result.Append(' ').Append(attribute);
            if (valueStart != i) result.Append("=\"").Append(value.Replace("\"", "&quot;")).Append('"');
        }
        if (name == "a" && hasHref)
        {
            // Links like the Word readers': app links (otzaria:// zayit:// ...) are routed by the host in place; web links open outside.
            result.Append(isAppLink ? " class=\"link app-link\" data-app-link=\"1\"" : " class=\"link\"");
            if (!isAppLink && !IsInternalHref(tag)) result.Append(" rel=\"noopener noreferrer\" target=\"_blank\"");
        }
        if (tag.EndsWith("/>", StringComparison.Ordinal)) result.Append(" /");
        return result.Append('>').ToString();
    }
}
