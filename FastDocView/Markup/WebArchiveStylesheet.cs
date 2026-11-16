using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// Word's HTML carries its styles as CSS in the head (p.MsoNormal, h1, span.MsoFootnoteReference, a:link, @font-face, @page,
/// @list...). The sheet never reaches the page as written: it is parsed here, each selector checked against a small grammar
/// (type, .class, #id, descendant and child combinators, the pseudo-classes that print: :link / :any-link / :first-child;
/// a printed link counts as unvisited, so :visited rules are dropped) and each declaration against a whitelist of print properties with safe values (no url(), expression,
/// behavior, escapes, @import or position / fixed layout). <see cref="MarkupBlockRenderer"/> then resolves the rules per element
/// into an inline style (sheet declarations by specificity and order, the element's own style="" after them), as a browser
/// cascades them. @-rules (fonts, pages, Word's @list definitions) and selectors outside the grammar are dropped.
/// </summary>
public sealed class WebArchiveStylesheet
{
    private const int MaximumCssCharacters = 256 * 1024;

    private readonly List<Rule> _rules = [];
    private readonly Dictionary<string, List<Rule>> _byId = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<Rule>> _byClass = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, List<Rule>> _byType = new(StringComparer.Ordinal);
    private readonly List<Rule> _universal = [];

    public static readonly WebArchiveStylesheet Empty = new();

    /// <summary>Declarations of html / body rules that inherit: drawn on the reader's root (.word-document), the document's own body.</summary>
    public string RootCss { get; private set; } = "";

    public bool IsEmpty => _rules.Count == 0;

    /// <summary>The sheet in its checked form (what the index stores): one rule per line, only selectors and declarations that passed.</summary>
    public static string Normalize(string css) => Parse(css).ToCssText();

    public static WebArchiveStylesheet Parse(string css)
    {
        var sheet = new WebArchiveStylesheet();
        if (css.Length == 0) return sheet;
        if (css.Length > MaximumCssCharacters) css = css.Substring(0, MaximumCssCharacters);
        string source = StripComments(css.Replace("<!--", " ").Replace("-->", " "));
        var root = new List<Declaration>();
        int at = 0, order = 0;
        while (at < source.Length)
        {
            while (at < source.Length && char.IsWhiteSpace(source[at])) at++;
            if (at >= source.Length) break;
            int open = source.IndexOf('{', at);
            if (source[at] == '@')
            {
                int semicolon = source.IndexOf(';', at);
                if (open < 0 || (semicolon >= 0 && semicolon < open)) { at = semicolon < 0 ? source.Length : semicolon + 1; continue; }
                at = MatchingBrace(source, open) + 1;
                continue;
            }
            if (open < 0) break;
            int close = MatchingBrace(source, open);
            string selectors = source.Substring(at, open - at);
            string body = source.Substring(open + 1, Math.Max(0, close - open - 1));
            at = close + 1;
            if (body.IndexOf('{') >= 0) continue;                      // a nested block: not a plain rule
            List<Declaration> declarations = ParseDeclarations(body);
            if (declarations.Count == 0) continue;
            foreach (string text in selectors.Split(','))
            {
                Selector? selector = Selector.Parse(text);
                if (selector is null) continue;
                if (selector.IsRoot) { root.AddRange(declarations.Where(d => InheritedProperties.Contains(d.Property))); continue; }
                sheet.Add(new Rule(selector, declarations, order++));
            }
        }
        sheet.RootCss = string.Concat(root.Select(d => d.ToCss()));
        return sheet;
    }

    private void Add(Rule rule)
    {
        _rules.Add(rule);
        Compound key = rule.Selector.Compounds[rule.Selector.Compounds.Count - 1];
        if (key.Id is not null) Bucket(_byId, key.Id).Add(rule);
        else if (key.Classes.Count > 0) Bucket(_byClass, key.Classes[0]).Add(rule);
        else if (key.Type is not null) Bucket(_byType, key.Type).Add(rule);
        else _universal.Add(rule);
    }

    private static List<Rule> Bucket(Dictionary<string, List<Rule>> buckets, string key)
    {
        if (!buckets.TryGetValue(key, out List<Rule>? list)) buckets[key] = list = [];
        return list;
    }

    private string ToCssText()
    {
        var text = new StringBuilder();
        if (RootCss.Length > 0) text.Append("body { ").Append(RootCss).Append(" }\n");
        foreach (Rule rule in _rules)
            text.Append(rule.Selector.Text).Append(" { ").Append(string.Concat(rule.Declarations.Select(d => d.ToCss()))).Append(" }\n");
        return text.ToString();
    }

    /// <summary>
    /// The sheet's declarations for the innermost element of <paramref name="ancestors"/> (outermost first), cascaded by
    /// specificity then source order, as one inline style text ("" when no rule matches).
    /// </summary>
    internal string StyleFor(IReadOnlyList<WebArchiveElement> ancestors)
    {
        if (_rules.Count == 0 || ancestors.Count == 0) return "";
        WebArchiveElement element = ancestors[ancestors.Count - 1];
        List<Rule>? matched = null;
        void Try(List<Rule>? rules)
        {
            if (rules is null) return;
            foreach (Rule rule in rules)
                if ((matched is null || !matched.Contains(rule)) && rule.Selector.Matches(ancestors)) (matched ??= []).Add(rule);
        }
        if (element.Id is not null && _byId.TryGetValue(element.Id, out List<Rule>? byId)) Try(byId);
        foreach (string name in element.Classes)
            if (_byClass.TryGetValue(name, out List<Rule>? byClass)) Try(byClass);
        if (_byType.TryGetValue(element.Name, out List<Rule>? byType)) Try(byType);
        Try(_universal);
        if (matched is null) return "";
        matched.Sort((a, b) => a.Selector.Specificity != b.Selector.Specificity ? a.Selector.Specificity.CompareTo(b.Selector.Specificity) : a.Order.CompareTo(b.Order));
        var css = new StringBuilder();
        foreach (Rule rule in matched)
            foreach (Declaration declaration in rule.Declarations) css.Append(declaration.ToCss());
        return css.ToString();
    }

    /// <summary>An inline style="" value with only the declarations that pass the same checks as the sheet's.</summary>
    public static string SanitizeInline(string style) => string.Concat(ParseDeclarations(style).Select(d => d.ToCss()));

    // -- declarations -----------------------------------------------------------------------

    private readonly record struct Declaration(string Property, string Value, bool IsImportant)
    {
        public string ToCss() => Property + ":" + Value + (IsImportant ? " !important" : "") + ";";
    }

    /// <summary>Print properties a sheet may set. No position / top / left / z-index, no content, no behavior / binding, no table-layout.</summary>
    private static readonly HashSet<string> AllowedProperties = new(StringComparer.Ordinal)
    {
        "color", "background", "background-color",
        "font", "font-family", "font-size", "font-weight", "font-style", "font-variant", "font-stretch", "font-kerning", "font-feature-settings",
        "text-decoration", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-underline-position", "text-underline-offset",
        "text-align", "text-align-last", "text-indent", "text-transform", "text-shadow", "line-height", "letter-spacing", "word-spacing",
        "white-space", "vertical-align", "direction", "unicode-bidi", "writing-mode", "text-orientation", "hyphens", "tab-size",
        "margin", "margin-top", "margin-right", "margin-bottom", "margin-left",
        "padding", "padding-top", "padding-right", "padding-bottom", "padding-left",
        "border", "border-top", "border-right", "border-bottom", "border-left", "border-width", "border-style", "border-color",
        "border-top-width", "border-right-width", "border-bottom-width", "border-left-width",
        "border-top-style", "border-right-style", "border-bottom-style", "border-left-style",
        "border-top-color", "border-right-color", "border-bottom-color", "border-left-color", "border-collapse", "border-spacing",
        "width", "height", "min-width", "min-height", "max-width", "max-height", "float", "clear",
        "display", "visibility", "list-style-type", "list-style-position", "caption-side", "empty-cells",
        "page-break-before", "page-break-after", "page-break-inside", "break-before", "break-after", "break-inside", "orphans", "widows",
        "column-count", "column-gap", "column-rule", "columns",
    };

    private static readonly HashSet<string> InheritedProperties = new(StringComparer.Ordinal)
    {
        "color", "font", "font-family", "font-size", "font-weight", "font-style", "font-variant", "font-stretch", "line-height",
        "letter-spacing", "word-spacing", "text-align", "text-indent", "text-transform", "direction", "white-space", "hyphens",
    };

    private static readonly HashSet<string> DisplayValues = new(StringComparer.Ordinal)
    {
        "none", "block", "inline", "inline-block", "list-item", "table", "inline-table", "table-row", "table-cell", "table-row-group",
        "table-header-group", "table-footer-group", "table-column", "table-column-group", "table-caption",
    };

    private static List<Declaration> ParseDeclarations(string body)
    {
        var result = new List<Declaration>();
        foreach (string part in SplitDeclarations(body))
        {
            int colon = part.IndexOf(':');
            if (colon <= 0) continue;
            string property = part.Substring(0, colon).Trim().ToLowerInvariant();
            string value = part.Substring(colon + 1).Trim();
            bool isImportant = false;
            int bang = value.LastIndexOf('!');
            if (bang >= 0)
            {
                if (!value.Substring(bang + 1).Trim().Equals("important", StringComparison.OrdinalIgnoreCase)) continue;
                isImportant = true;
                value = value.Substring(0, bang).Trim();
            }
            if (!AllowedProperties.Contains(property) || !IsSafeValue(value)) continue;
            if (property == "display" && !DisplayValues.Contains(value.ToLowerInvariant())) continue;
            if (property == "float" && value.ToLowerInvariant() is not ("left" or "right" or "none")) continue;
            if (WebArchiveCssLength.Clamp(property, value) is not { } bounded) continue;
            result.Add(new Declaration(property, bounded, isImportant));
        }
        return result;
    }

    /// <summary>Splits on ';' outside quotes and parentheses.</summary>
    private static IEnumerable<string> SplitDeclarations(string body)
    {
        int start = 0, depth = 0;
        char quote = '\0';
        for (int i = 0; i < body.Length; i++)
        {
            char c = body[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c == '(') depth++;
            else if (c == ')') depth = Math.Max(0, depth - 1);
            else if (c == ';' && depth == 0) { yield return body.Substring(start, i - start); start = i + 1; }
        }
        if (start < body.Length) yield return body.Substring(start);
    }

    /// <summary>
    /// A value is kept only when nothing in it can load, run or escape: no backslash escapes, no '&lt;' '&gt;' '{' '}' '@' ';'
    /// or controls, balanced quotes, no url( / image( / expression / behavior / javascript / binding / var( / attr(, and no "fixed".
    /// </summary>
    private static bool IsSafeValue(string value)
    {
        if (value.Length == 0 || value.Length > 1024) return false;
        int doubleQuotes = 0, singleQuotes = 0;
        foreach (char c in value)
        {
            if (c is '\\' or '<' or '>' or '{' or '}' or '@' or ';' || char.IsControl(c)) return false;
            if (c == '"') doubleQuotes++;
            else if (c == '\'') singleQuotes++;
        }
        if (doubleQuotes % 2 != 0 || singleQuotes % 2 != 0) return false;
        string lower = value.ToLowerInvariant();
        foreach (string word in UnsafeWords)
            if (lower.Contains(word)) return false;
        return true;
    }

    private static readonly string[] UnsafeWords =
        ["url(", "image(", "image-set(", "element(", "expression", "behavior", "javascript", "vbscript", "binding", "var(", "attr(", "env(", "fixed", "/*"];

    // -- selectors --------------------------------------------------------------------------

    private sealed record Rule(Selector Selector, List<Declaration> Declarations, int Order);

    private sealed class Compound
    {
        public string? Type;                 // lower case; null = any
        public string? Id;
        public readonly List<string> Classes = [];
        public bool IsLink, IsFirstChild;
        public bool IsChildOfPrevious;       // '>' before this compound

        public bool Matches(WebArchiveElement element)
        {
            if (Type is not null && Type != element.Name) return false;
            if (Id is not null && !string.Equals(Id, element.Id, StringComparison.OrdinalIgnoreCase)) return false;
            foreach (string name in Classes)
                if (!element.Classes.Contains(name, StringComparer.OrdinalIgnoreCase)) return false;
            if (IsLink && !element.IsLink) return false;
            if (IsFirstChild && !element.IsFirstChild) return false;
            return true;
        }
    }

    private sealed class Selector
    {
        public readonly List<Compound> Compounds = [];
        public int Specificity;
        public string Text = "";
        public bool IsRoot => Compounds.Count == 1 && Compounds[0] is { Type: "html" or "body", Id: null, Classes.Count: 0, IsLink: false, IsFirstChild: false };

        public static Selector? Parse(string text)
        {
            text = text.Trim();
            if (text.Length == 0 || text.Length > 512) return null;
            var selector = new Selector();
            int i = 0;
            bool childNext = false;
            while (i < text.Length)
            {
                while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
                if (i >= text.Length) break;
                if (text[i] == '>') { if (childNext || selector.Compounds.Count == 0) return null; childNext = true; i++; continue; }
                var compound = new Compound { IsChildOfPrevious = childNext };
                childNext = false;
                int start = i;
                if (text[i] == '*') i++;
                else if (IsNameCharacter(text[i])) { int s = i; while (i < text.Length && IsNameCharacter(text[i])) i++; compound.Type = text.Substring(s, i - s).ToLowerInvariant(); }
                while (i < text.Length && !char.IsWhiteSpace(text[i]) && text[i] != '>')
                {
                    char kind = text[i++];
                    int s = i;
                    while (i < text.Length && IsNameCharacter(text[i])) i++;
                    string name = text.Substring(s, i - s);
                    if (name.Length == 0) return null;
                    if (kind == '.') { compound.Classes.Add(name); selector.Specificity += 1 << 10; }
                    else if (kind == '#') { if (compound.Id is not null) return null; compound.Id = name; selector.Specificity += 1 << 20; }
                    else if (kind == ':')
                    {
                        switch (name.ToLowerInvariant())
                        {
                            // a printed link is unvisited (:visited never matches); hover / focus / active never print
                            case "link": case "any-link": compound.IsLink = true; break;
                            case "first-child": compound.IsFirstChild = true; break;
                            default: return null;
                        }
                        selector.Specificity += 1 << 10;
                    }
                    else return null;                                   // attribute selectors, pseudo-elements, sibling combinators
                }
                if (i == start) return null;
                if (compound.Type is not null) selector.Specificity += 1;
                selector.Compounds.Add(compound);
            }
            if (childNext || selector.Compounds.Count == 0 || selector.Compounds.Count > 16) return null;
            selector.Text = string.Join(" ", selector.Compounds.Select(ToText));
            return selector;
        }

        private static string ToText(Compound compound)
            => (compound.IsChildOfPrevious ? "> " : "") + (compound.Type ?? (compound.Id is null && compound.Classes.Count == 0 && !compound.IsLink && !compound.IsFirstChild ? "*" : ""))
               + (compound.Id is null ? "" : "#" + compound.Id) + string.Concat(compound.Classes.Select(c => "." + c))
               + (compound.IsLink ? ":link" : "") + (compound.IsFirstChild ? ":first-child" : "");

        /// <summary>Right to left over the open elements: the last compound is the element, the rest its ancestors (backtracking for descendants).</summary>
        public bool Matches(IReadOnlyList<WebArchiveElement> ancestors)
            => Compounds[Compounds.Count - 1].Matches(ancestors[ancestors.Count - 1])
               && MatchFrom(Compounds.Count - 2, ancestors, ancestors.Count - 2, Compounds[Compounds.Count - 1].IsChildOfPrevious);

        private bool MatchFrom(int compound, IReadOnlyList<WebArchiveElement> ancestors, int element, bool isChild)
        {
            if (compound < 0) return true;
            for (int e = element; e >= 0; e--)
            {
                if (Compounds[compound].Matches(ancestors[e]) && MatchFrom(compound - 1, ancestors, e - 1, Compounds[compound].IsChildOfPrevious)) return true;
                if (isChild) return false;
            }
            return false;
        }
    }

    private static bool IsNameCharacter(char c) => c is >= 'a' and <= 'z' or >= 'A' and <= 'Z' or >= '0' and <= '9' or '-' or '_' || c > 0x7F;

    private static string StripComments(string css)
    {
        var text = new StringBuilder(css.Length);
        for (int i = 0; i < css.Length; i++)
        {
            if (css[i] == '/' && i + 1 < css.Length && css[i + 1] == '*')
            {
                int close = css.IndexOf("*/", i + 2, StringComparison.Ordinal);
                i = close < 0 ? css.Length : close + 1;
                continue;
            }
            text.Append(css[i]);
        }
        return text.ToString();
    }

    private static int MatchingBrace(string text, int open)
    {
        int depth = 0;
        for (int i = open; i < text.Length; i++)
        {
            if (text[i] == '{') depth++;
            else if (text[i] == '}' && --depth == 0) return i;
        }
        return text.Length - 1;
    }
}

/// <summary>One open element as the stylesheet sees it: name (lower case), id, classes, and the facts the printing pseudo-classes need.</summary>
internal sealed class WebArchiveElement(string name, string? id, string[] classes, bool isLink, bool isFirstChild, bool areChildrenKnown)
{
    public string Name { get; } = name;
    public string? Id { get; } = id;
    public string[] Classes { get; } = classes;
    public bool IsLink { get; } = isLink;
    public bool IsFirstChild { get; } = isFirstChild;
    public bool AreChildrenKnown { get; } = areChildrenKnown;    // opened in the block's own text, not reopened from the open-tag stack
    public int ElementChildren;                                  // element children seen so far
}
