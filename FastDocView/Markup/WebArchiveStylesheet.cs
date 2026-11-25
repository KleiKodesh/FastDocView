using System.Globalization;
using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// A web archive's head CSS (Word: p.MsoNormal, h1, span.MsoFootnoteReference, a:link, @font-face, @page, @list...; a page a
/// browser saved: any selector). The sheet never reaches the page as written: it is parsed here, each selector checked against a
/// grammar - type, .class, #id, attribute selectors ([a], [a=v], ~=, |=, ^=, $=, *=, with the i flag), the descendant, child
/// ('&gt;'), adjacent ('+') and general sibling ('~') combinators, the pseudo-classes that print (:link / :any-link - a printed
/// link is unvisited, so :visited never matches -, :root, :first-child, :last-child, :only-child, :nth-child(), :nth-last-child(),
/// :first-of-type, :last-of-type, :only-of-type, :nth-of-type(), :nth-last-of-type(), :empty, :not() of simple selectors) and the
/// pseudo-elements ::before / ::after (content: strings and quote keywords), ::first-letter, ::first-line - and each declaration
/// against a whitelist of print properties with safe values (no url(), expression, behavior, escapes, @import or position / fixed
/// layout). @media rules apply as in print (Edge's print of the archive: print / all, not screen, orientation portrait, color, the
/// page's width and height - see <see cref="WebArchiveMedia"/>). Word's own print of an archive ignores all of these (only plain
/// type / class / id rules; test-data/word-mhtml/css-selectors-v1.pdf), and Word's own archives never use them; a browser's
/// archive prints as the browser prints it (css-selectors-v1-edge.pdf).
///
/// Resolution stays per element and windowed (<see cref="WebArchiveMatch"/>): each element's matched selector prefixes are worked
/// out from its parent's state when it opens - descendant and child combinators from the ancestors' matches, sibling combinators
/// and child-index pseudo-classes from the parent's state of its children so far. The index pass keeps that state for the
/// elements open at each block start (in the block's open-tag stack, <see cref="MarkupBlockScanner"/>), so a block resolves on
/// its own. :last-child and the other pseudo-classes that depend on later siblings resolve when the parent's children end inside
/// the block or at its end (the next block's stack shows what stays open); <see cref="MarkupBlockRenderer"/> writes the
/// declarations into each element's inline style (by specificity and order, the element's own style="" after them, as a browser
/// cascades them) and marks elements with pseudo-element rules (data-wa-pe), whose declarations the document stylesheet carries
/// (<see cref="PseudoElementCss"/>). Other @-rules (fonts, pages, Word's @list) and selectors outside the grammar are dropped.
/// </summary>
public sealed class WebArchiveStylesheet
{
    private const int MaximumCssCharacters = 256 * 1024;

    private readonly List<Rule> _rules = [];
    private readonly List<Prefix> _prefixes = [];
    private readonly Dictionary<int, Rule> _ruleByFullPrefix = new();
    private readonly HashSet<int> _descendantSources = [];   // prefixes a descendant combinator reads from an ancestor
    private readonly HashSet<int> _siblingSources = [];      // prefixes a sibling combinator reads from an earlier sibling
    // the same two sets by prefix id, and the full prefixes' rules by id: the matcher's compact form, built once after parsing
    private bool[] _isDescendantSource = [], _isSiblingSource = [];
    private Rule?[] _ruleById = [];
    [ThreadStatic] private static int[]? _matchScratch;

    public static readonly WebArchiveStylesheet Empty = new();

    /// <summary>Declarations of html / body rules that inherit: drawn on the reader's root (.word-document), the document's own body.</summary>
    public string RootCss { get; private set; } = "";

    public bool IsEmpty => _rules.Count == 0;

    /// <summary>The sheet has sibling combinators or child-index pseudo-classes: the index pass keeps the open elements' child state.</summary>
    public bool NeedsStructure { get; private set; }

    /// <summary>The sheet has rules that depend on where an element ends (pseudo-elements, later-sibling pseudo-classes): the renderer asks what the next block reopens.</summary>
    public bool NeedsEnds { get; private set; }

    /// <summary>Some rule's subject has a later-sibling pseudo-class (:last-child, :nth-last-*, :only-*, :empty): the index pass records the ends of elements that span blocks.</summary>
    public bool HasLaterSiblingRules { get; private set; }

    /// <summary>The pseudo-element rules for the document stylesheet: <c>.word-document [data-wa-pe~="N"]::before { ... }</c>, in cascade order.</summary>
    public string PseudoElementCss { get; private set; } = "";

    /// <summary>The sheet in its checked form (what the index stores): one rule per line, only selectors and declarations that passed; matching @media rules flattened.</summary>
    public static string Normalize(string css)
    {
        if (css.Length == 0) return "";
        lock (CacheLock)
            if (ReferenceEquals(_lastNormalizedSource, css) || string.Equals(_lastNormalizedSource, css, StringComparison.Ordinal)) return _lastNormalized!;
        string normalized = Parse(css).ToCssText();
        lock (CacheLock) { _lastNormalizedSource = css; _lastNormalized = normalized; }
        return normalized;
    }

    // the last sheet parsed / normalized: one document's index pass, rescans and renderer ask for the same sheet many times
    // (a sheet is never changed after parsing, so one instance serves them all)
    private static readonly object CacheLock = new();
    private static string? _lastNormalizedSource, _lastNormalized, _lastParsedSource;
    private static WebArchiveStylesheet? _lastParsed;

    /// <summary><see cref="Parse"/>, sharing the instance with the last caller that parsed the same text.</summary>
    public static WebArchiveStylesheet ParseShared(string css)
    {
        if (css.Length == 0) return Empty;
        lock (CacheLock)
            if (ReferenceEquals(_lastParsedSource, css) || string.Equals(_lastParsedSource, css, StringComparison.Ordinal)) return _lastParsed!;
        WebArchiveStylesheet sheet = Parse(css);
        lock (CacheLock) { _lastParsedSource = css; _lastParsed = sheet; }
        return sheet;
    }

    public static WebArchiveStylesheet Parse(string css)
    {
        var sheet = new WebArchiveStylesheet();
        if (css.Length == 0) return sheet;
        if (css.Length > MaximumCssCharacters) css = css.Substring(0, MaximumCssCharacters);
        string source = StripComments(css.Replace("<!--", " ").Replace("-->", " "));
        var root = new List<Declaration>();
        int order = 0;
        sheet.ParseRules(source, 0, source.Length, root, ref order, depth: 0);
        sheet.Compile();
        sheet.RootCss = string.Concat(root.Select(d => d.ToCss()));
        sheet.PseudoElementCss = string.Concat(sheet._rules.Where(r => r.Selector.PseudoElement != PseudoElement.None)
            .OrderBy(r => r.Selector.Specificity).ThenBy(r => r.Order)
            .Select(r => ".word-document [data-wa-pe~=\"" + r.Order.ToString(CultureInfo.InvariantCulture) + "\"]::" + PseudoElementName(r.Selector.PseudoElement)
                         + " { " + string.Concat(r.Declarations.Select(d => d.ToCss())) + " }\n"));
        return sheet;
    }

    private void ParseRules(string source, int at, int end, List<Declaration> root, ref int order, int depth)
    {
        while (at < end)
        {
            while (at < end && char.IsWhiteSpace(source[at])) at++;
            if (at >= end) break;
            int open = source.IndexOf('{', at, end - at);
            if (source[at] == '@')
            {
                int semicolon = source.IndexOf(';', at, end - at);
                if (open < 0 || (semicolon >= 0 && semicolon < open)) { at = semicolon < 0 ? end : semicolon + 1; continue; }
                int closeAt = Math.Min(end - 1, MatchingBrace(source, open));
                string prelude = source.Substring(at, open - at).Trim();
                // @media: its rules apply when the query matches print (Edge's print of the archive); nested @media too
                if (depth < 8 && prelude.StartsWith("@media", StringComparison.OrdinalIgnoreCase) && WebArchiveMedia.MatchesPrint(prelude.Substring(6)))
                    ParseRules(source, open + 1, closeAt, root, ref order, depth + 1);
                at = closeAt + 1;
                continue;
            }
            if (open < 0) break;
            int close = Math.Min(end - 1, MatchingBrace(source, open));
            string selectors = source.Substring(at, open - at);
            string body = source.Substring(open + 1, Math.Max(0, close - open - 1));
            at = close + 1;
            if (body.IndexOf('{') >= 0) continue;                      // a nested block: not a plain rule
            List<Declaration>? plain = null, generated = null;
            foreach (string text in SplitSelectorList(selectors))
            {
                Selector? selector = Selector.Parse(text);
                if (selector is null) continue;
                List<Declaration> declarations = selector.AllowsContent ? (generated ??= ParseDeclarations(body, allowsContent: true)) : (plain ??= ParseDeclarations(body));
                if (declarations.Count == 0) continue;
                if (selector.IsRoot) { root.AddRange(declarations.Where(d => InheritedProperties.Contains(d.Property))); continue; }
                Add(new Rule(selector, declarations, order++));
            }
        }
    }

    /// <summary>Splits a selector list on commas outside brackets and parentheses.</summary>
    private static IEnumerable<string> SplitSelectorList(string text)
    {
        int start = 0, depth = 0;
        char quote = '\0';
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c is '(' or '[') depth++;
            else if (c is ')' or ']') depth = Math.Max(0, depth - 1);
            else if (c == ',' && depth == 0) { yield return text.Substring(start, i - start); start = i + 1; }
        }
        yield return text.Substring(start);
    }

    private void Add(Rule rule)
    {
        _rules.Add(rule);
        Selector selector = rule.Selector;
        if (selector.UsesStructure) NeedsStructure = true;
        if (selector.PseudoElement != PseudoElement.None || selector.Subject.NeedsLaterSiblings) NeedsEnds = true;
        int first = _prefixes.Count;
        for (int i = 0; i < selector.Compounds.Count; i++)
        {
            var prefix = new Prefix(first + i, selector.Compounds[i], i == 0 ? Combinator.None : selector.Combinators[i], i == 0 ? -1 : first + i - 1);
            _prefixes.Add(prefix);
            if (prefix.Combinator == Combinator.Descendant) _descendantSources.Add(prefix.Previous);
            else if (prefix.Combinator is Combinator.Adjacent or Combinator.General) _siblingSources.Add(prefix.Previous);
        }
        _ruleByFullPrefix[first + selector.Compounds.Count - 1] = rule;
    }

    /// <summary>The per-id tables the matcher reads (no hashing per element).</summary>
    private void Compile()
    {
        int count = _prefixes.Count;
        _isDescendantSource = new bool[count];
        _isSiblingSource = new bool[count];
        _ruleById = new Rule?[count];
        foreach (int id in _descendantSources) _isDescendantSource[id] = true;
        foreach (int id in _siblingSources) _isSiblingSource[id] = true;
        foreach (KeyValuePair<int, Rule> pair in _ruleByFullPrefix) _ruleById[pair.Key] = pair.Value;
        HasLaterSiblingRules = _rules.Exists(rule => rule.Selector.Subject.NeedsLaterSiblings);
        // a source: a prefix some combinator reads from another element (its parent, an ancestor, an earlier sibling)
        var isSource = new bool[count];
        foreach (Prefix prefix in _prefixes) if (prefix.Previous >= 0) isSource[prefix.Previous] = true;
        _allBuckets = new Buckets();
        _sourceBuckets = new Buckets();
        foreach (Prefix prefix in _prefixes)
        {
            Place(_allBuckets, prefix);
            if (isSource[prefix.Id]) Place(_sourceBuckets, prefix);
            else _hasNonSourcePrefixes = true;
        }

        static void Place(Buckets buckets, Prefix prefix)
        {
            Compound key = prefix.Compound;
            if (key.Id is not null) Bucket(buckets.ById, key.Id).Add(prefix);
            else if (key.Classes.Count > 0) Bucket(buckets.ByClass, key.Classes[0]).Add(prefix);
            else if (key.Type is not null) Bucket(buckets.ByType, key.Type).Add(prefix);
            else buckets.Universal.Add(prefix);
        }
    }

    private static List<Prefix> Bucket(Dictionary<string, List<Prefix>> buckets, string key)
    {
        if (!buckets.TryGetValue(key, out List<Prefix>? list)) buckets[key] = list = [];
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

    private static string PseudoElementName(PseudoElement pseudo) => pseudo switch
    {
        PseudoElement.Before => "before", PseudoElement.After => "after", PseudoElement.FirstLetter => "first-letter", _ => "first-line",
    };

    // -- matching ---------------------------------------------------------------------------

    /// <summary>
    /// The state of <paramref name="element"/> opening under <paramref name="parent"/> (null: the document root): the selector
    /// prefixes it matches. Counts it as the parent's next element child (child index, type index, sibling matches).
    /// <paramref name="sourcesOnly"/> (the index pass): only the prefixes a combinator reads from another element are matched now
    /// - all that the state of later elements needs; the element's full match is worked out from the same inputs only when its
    /// state is written out (<see cref="WebArchiveMatch.FullMatched"/>), which is rare (elements open at a block start).
    /// </summary>
    internal WebArchiveMatch Open(WebArchiveMatch? parent, WebArchiveElement element, bool sourcesOnly = false)
    {
        int index = parent is null ? 1 : parent.Children + 1;
        int typeIndex = parent is null ? 1 : parent.TypeCount(element.Name) + 1;
        var context = new WebArchiveMatchContext(element, index, typeIndex, parent is not null, parent?.Matched ?? [], parent?.AncestorOrSelf ?? [],
                                                 parent?.PreviousChild ?? [], parent?.EarlierChildren ?? []);
        int[] matched = Match(context, sourcesOnly ? _sourceBuckets : _allBuckets);
        int[] ancestorOrSelf = Union(parent?.AncestorOrSelf, matched, _isDescendantSource);
        var state = sourcesOnly && _hasNonSourcePrefixes
            ? new WebArchiveMatch(matched, ancestorOrSelf, parent, index, typeIndex, this, context)
            : new WebArchiveMatch(matched, ancestorOrSelf, parent, index, typeIndex);
        parent?.CountChild(element.Name, Filter(matched, _isSiblingSource));
        return state;
    }

    /// <summary>Every prefix the element matches (a deferred full match of an index-pass state).</summary>
    internal int[] MatchAll(in WebArchiveMatchContext context) => Match(context, _allBuckets);

    /// <summary>The prefix candidates by the element's id, classes, type, then the universal ones (one set of them per use).</summary>
    private sealed class Buckets
    {
        public readonly Dictionary<string, List<Prefix>> ById = new(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, List<Prefix>> ByClass = new(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, List<Prefix>> ByType = new(StringComparer.Ordinal);
        public readonly List<Prefix> Universal = [];
        public bool IsEmpty => ById.Count == 0 && ByClass.Count == 0 && ByType.Count == 0 && Universal.Count == 0;
    }

    private Buckets _allBuckets = new(), _sourceBuckets = new();
    private bool _hasNonSourcePrefixes;

    private int[] Match(in WebArchiveMatchContext context, Buckets buckets)
    {
        if (buckets.IsEmpty) return [];
        int[] scratch = _matchScratch is { } buffer && buffer.Length >= _prefixes.Count ? buffer : _matchScratch = new int[Math.Max(64, _prefixes.Count)];
        int count = 0;
        WebArchiveElement element = context.Element;
        if (buckets.ById.Count > 0 && element.Id is { } elementId && buckets.ById.TryGetValue(elementId, out List<Prefix>? byId)) Try(byId, scratch, ref count, context);
        if (buckets.ByClass.Count > 0)
            foreach (string name in element.Classes)
                if (buckets.ByClass.TryGetValue(name, out List<Prefix>? byClass)) Try(byClass, scratch, ref count, context);
        if (buckets.ByType.Count > 0 && buckets.ByType.TryGetValue(element.Name, out List<Prefix>? byType)) Try(byType, scratch, ref count, context);
        if (buckets.Universal.Count > 0) Try(buckets.Universal, scratch, ref count, context);
        if (count == 0) return [];
        var matched = new int[count];
        Array.Copy(scratch, matched, count);
        if (count > 1) Array.Sort(matched);
        return matched;

        static void Try(List<Prefix> candidates, int[] into, ref int found, in WebArchiveMatchContext context)
        {
            foreach (Prefix prefix in candidates)
                if (prefix.Compound.Matches(context.Element, context.Index, context.TypeIndex, context.HasParent) && Related(prefix, context)
                    && Array.IndexOf(into, prefix.Id, 0, found) < 0) into[found++] = prefix.Id;
        }
    }

    private static bool Related(Prefix prefix, in WebArchiveMatchContext context) => prefix.Combinator switch
    {
        Combinator.None => true,
        Combinator.Child => context.HasParent && Contains(context.ParentMatched, prefix.Previous),
        Combinator.Descendant => context.HasParent && Contains(context.ParentAncestors, prefix.Previous),
        Combinator.Adjacent => context.HasParent && Contains(context.ParentPreviousChild, prefix.Previous),
        _ => context.HasParent && Contains(context.ParentEarlierChildren, prefix.Previous),
    };

    private static bool Contains(int[] sorted, int id) => sorted.Length != 0 && Array.BinarySearch(sorted, id) >= 0;

    /// <summary>The ids of <paramref name="matched"/> (sorted) that <paramref name="keep"/> marks, sorted; the empty array when none.</summary>
    private static int[] Filter(int[] matched, bool[] keep)
    {
        int count = 0;
        foreach (int id in matched) if (keep[id]) count++;
        if (count == 0) return [];
        if (count == matched.Length) return matched;
        var kept = new int[count];
        count = 0;
        foreach (int id in matched) if (keep[id]) kept[count++] = id;
        return kept;
    }

    /// <summary>The sorted union of <paramref name="inherited"/> and the kept ids of <paramref name="matched"/> (both sorted); shares an input when nothing is added.</summary>
    private static int[] Union(int[]? inherited, int[] matched, bool[] keep) => WebArchiveMatch.MergeSorted(inherited ?? [], Filter(matched, keep));

    /// <summary>What the sheet gives one element: its inline declarations, its pseudo-element rule ids, and whether it is hidden.</summary>
    internal readonly record struct ElementStyle(string Css, string StartPseudoIds, string AfterPseudoIds, bool IsHidden, bool IsDisplayImportant,
                                                  string BeforeText, string AfterText, bool HasLaterSiblingRules)
    {
        public static readonly ElementStyle None = new("", "", "", false, false, "", "", false);
        /// <summary>Whether the element is drawn by the renderer only once its end is known (pseudo-elements, later-sibling rules).</summary>
        public bool IsDeferred => HasLaterSiblingRules || StartPseudoIds.Length > 0 || AfterPseudoIds.Length > 0;
    }

    /// <summary>
    /// The declarations for an element by its state, cascaded by specificity then source order, as one inline style text, plus
    /// the pseudo-element rules it takes. A rule whose subject needs later siblings (:last-child...) applies only when
    /// <paramref name="later"/> answers for the element (null: not known yet; the rule is left out).
    /// </summary>
    internal ElementStyle StyleFor(WebArchiveMatch state, WebArchiveLaterFacts? later)
    {
        if (_rules.Count == 0 || state.FullMatched.Length == 0) return ElementStyle.None;
        List<Rule>? matched = null;
        bool hasLaterRules = false;
        foreach (int id in state.FullMatched)
        {
            if (_ruleById[id] is not { } rule) continue;
            if (rule.Selector.Subject.NeedsLaterSiblings)
            {
                hasLaterRules = true;
                if (later is not { } facts || !rule.Selector.Subject.MatchesLater(state, facts)) continue;
            }
            (matched ??= []).Add(rule);
        }
        if (matched is null) return ElementStyle.None with { HasLaterSiblingRules = hasLaterRules };
        matched.Sort((a, b) => a.Selector.Specificity != b.Selector.Specificity ? a.Selector.Specificity.CompareTo(b.Selector.Specificity) : a.Order.CompareTo(b.Order));
        var css = new StringBuilder();
        var pseudo = new StringBuilder();
        var afterPseudo = new StringBuilder();
        string? display = null, importantDisplay = null, before = null, after = null;
        foreach (Rule rule in matched)
        {
            if (rule.Selector.PseudoElement != PseudoElement.None)
            {
                StringBuilder ids = rule.Selector.PseudoElement == PseudoElement.After ? afterPseudo : pseudo;
                if (ids.Length > 0) ids.Append(' ');
                ids.Append(rule.Order.ToString(CultureInfo.InvariantCulture));
                foreach (Declaration declaration in rule.Declarations)
                    if (declaration.Property == "content" && WebArchiveGeneratedContent.TextOf(declaration.Value) is { } text)
                    {
                        if (rule.Selector.PseudoElement == PseudoElement.Before) before = text;
                        else if (rule.Selector.PseudoElement == PseudoElement.After) after = text;
                    }
                continue;
            }
            foreach (Declaration declaration in rule.Declarations)
            {
                css.Append(declaration.ToCss());
                if (declaration.Property == "display") { if (declaration.IsImportant) importantDisplay = declaration.Value; else display = declaration.Value; }
            }
        }
        bool isHidden = string.Equals(importantDisplay ?? display, "none", StringComparison.OrdinalIgnoreCase);
        return new ElementStyle(css.ToString(), pseudo.ToString(), afterPseudo.ToString(), isHidden, importantDisplay is not null, before ?? "", after ?? "", hasLaterRules);
    }

    /// <summary>Whether an inline style="" (checked) sets display:none (an element's own style wins over the sheet's unless !important there).</summary>
    internal static string? InlineDisplay(string checkedInline)
    {
        string? display = null;
        foreach (Declaration declaration in ParseDeclarations(checkedInline))
            if (declaration.Property == "display") display = declaration.Value;
        return display;
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
        "font", "font-family", "font-size", "font-weight", "font-style", "font-variant", "font-stretch", "font-kerning", "font-feature-settings", "font-variant-ligatures", "font-variant-numeric",
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

    private static List<Declaration> ParseDeclarations(string body, bool allowsContent = false)
    {
        var result = new List<Declaration>();
        foreach (string part in SplitDeclarations(body))
        {
            int colon = part.IndexOf(':');
            if (colon <= 0) continue;
            string property = part.Substring(0, colon).Trim().ToLowerInvariant();
            string value = part.Substring(colon + 1).Trim();
            if (property == "content")
            {
                // generated content (::before / ::after only): strings and quote keywords, re-written safely
                if (allowsContent && WebArchiveGeneratedContent.Parse(value) is { } generated) result.Add(new Declaration(property, generated.Css, generated.IsImportant));
                continue;
            }
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

    private sealed record Prefix(int Id, Compound Compound, Combinator Combinator, int Previous);

    private enum Combinator : byte { None, Descendant, Child, Adjacent, General }

    internal enum PseudoElement : byte { None, Before, After, FirstLetter, FirstLine }

    private enum AttributeOperator : byte { Exists, Equals, Word, DashPrefix, Prefix, Suffix, Substring }

    private sealed record AttributeTest(string Name, AttributeOperator Operator, string Value, bool IgnoresCase);

    /// <summary>An+B of :nth-child() and its kin.</summary>
    private readonly record struct Nth(int A, int B)
    {
        public bool Matches(int position) => A == 0 ? position == B : (position - B) / A >= 0 && (position - B) % A == 0;
        public override string ToString() => A.ToString(CultureInfo.InvariantCulture) + "n" + (B >= 0 ? "+" : "") + B.ToString(CultureInfo.InvariantCulture);

        public static Nth? Parse(string text)
        {
            string t = new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray()).ToLowerInvariant();
            if (t == "odd") return new Nth(2, 1);
            if (t == "even") return new Nth(2, 0);
            int n = t.IndexOf('n');
            if (n < 0) return int.TryParse(t, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int b0) ? new Nth(0, b0) : null;
            string a = t.Substring(0, n), rest = t.Substring(n + 1);
            int av = a switch { "" or "+" => 1, "-" => -1, _ => int.TryParse(a, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int parsed) ? parsed : int.MinValue };
            if (av == int.MinValue) return null;
            int bv = 0;
            if (rest.Length > 0 && !int.TryParse(rest, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out bv)) return null;
            return new Nth(av, bv);
        }
    }

    private sealed class Compound
    {
        public string? Type;                 // lower case; null = any
        public string? Id;
        public readonly List<string> Classes = [];
        public readonly List<AttributeTest> Attributes = [];
        public bool IsLink, IsRoot;
        public Nth? NthChild, NthOfType;                      // :first-child = nth-child(1), :first-of-type = nth-of-type(1)
        public Nth? NthLastChild, NthLastOfType;              // later siblings: resolved by the renderer (WebArchiveLaterFacts)
        public bool IsEmpty;
        public readonly List<Compound> Not = [];
        public readonly List<string> PseudoTexts = [];        // as written back (normalized)

        public bool UsesStructure => NthChild is not null || NthOfType is not null || NeedsLaterSiblings || Not.Any(n => n.UsesStructure);
        public bool NeedsLaterSiblings => NthLastChild is not null || NthLastOfType is not null || IsEmpty;
        public bool IsSimpleArgument => !NeedsLaterSiblings && Not.Count == 0;

        public bool Matches(WebArchiveElement element, int index, int typeIndex, bool hasParent)
        {
            if (Type is not null && Type != element.Name) return false;
            if (IsRoot && element.Name != "html") return false;
            if (Id is not null && !string.Equals(Id, element.Id, StringComparison.OrdinalIgnoreCase)) return false;
            foreach (string name in Classes)
                if (!HasClass(element.Classes, name)) return false;
            if (IsLink && !element.IsLink) return false;
            foreach (AttributeTest test in Attributes)
                if (!MatchesAttribute(test, element.Attribute(test.Name))) return false;
            if (NthChild is { } nthChild && (!hasParent || !nthChild.Matches(index))) return false;
            if (NthOfType is { } nthOfType && (!hasParent || !nthOfType.Matches(typeIndex))) return false;
            foreach (Compound not in Not)
                if (not.Matches(element, index, typeIndex, hasParent)) return false;
            return true;
        }

        private static bool HasClass(string[] classes, string name)
        {
            foreach (string candidate in classes) if (string.Equals(candidate, name, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        /// <summary>The later-sibling pseudo-classes of the subject, once the element's later siblings and content are known.</summary>
        public bool MatchesLater(WebArchiveMatch state, WebArchiveLaterFacts facts)
        {
            if (NthLastChild is { } nthLastChild && !nthLastChild.Matches(facts.LaterSiblings + 1)) return false;
            if (NthLastOfType is { } nthLastOfType && !nthLastOfType.Matches(facts.LaterSiblingsOfType + 1)) return false;
            if (IsEmpty && !facts.IsEmpty) return false;
            return true;
        }

        private static bool MatchesAttribute(AttributeTest test, string? value)
        {
            if (value is null) return false;
            if (test.Operator == AttributeOperator.Exists) return true;
            StringComparison comparison = test.IgnoresCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            string wanted = test.Value;
            return test.Operator switch
            {
                AttributeOperator.Equals => string.Equals(value, wanted, comparison),
                AttributeOperator.Word => wanted.Length > 0 && wanted.IndexOfAny([' ', '\t', '\n', '\r', '\f']) < 0
                                          && value.Split([' ', '\t', '\n', '\r', '\f'], StringSplitOptions.RemoveEmptyEntries).Any(word => string.Equals(word, wanted, comparison)),
                AttributeOperator.DashPrefix => string.Equals(value, wanted, comparison) || value.StartsWith(wanted + "-", comparison),
                AttributeOperator.Prefix => wanted.Length > 0 && value.StartsWith(wanted, comparison),
                AttributeOperator.Suffix => wanted.Length > 0 && value.EndsWith(wanted, comparison),
                _ => wanted.Length > 0 && value.IndexOf(wanted, comparison) >= 0,
            };
        }

        public string ToText(bool isFirst)
        {
            var text = new StringBuilder();
            text.Append(Type ?? "");
            if (Id is not null) text.Append('#').Append(Id);
            foreach (string name in Classes) text.Append('.').Append(name);
            foreach (AttributeTest test in Attributes)
            {
                text.Append('[').Append(test.Name);
                if (test.Operator != AttributeOperator.Exists)
                    text.Append(test.Operator switch
                    {
                        AttributeOperator.Equals => "=", AttributeOperator.Word => "~=", AttributeOperator.DashPrefix => "|=",
                        AttributeOperator.Prefix => "^=", AttributeOperator.Suffix => "$=", _ => "*=",
                    }).Append('"').Append(test.Value).Append('"').Append(test.IgnoresCase ? " i" : "");
                text.Append(']');
            }
            foreach (string pseudo in PseudoTexts) text.Append(':').Append(pseudo);
            if (text.Length == 0) text.Append('*');
            return text.ToString();
        }
    }

    private sealed class Selector
    {
        public readonly List<Compound> Compounds = [];
        public readonly List<Combinator> Combinators = [];     // [i]: between compound i - 1 and i ([0] unused)
        public int Specificity;
        public string Text = "";
        public PseudoElement PseudoElement;
        public Compound Subject => Compounds[Compounds.Count - 1];
        public bool AllowsContent => PseudoElement is PseudoElement.Before or PseudoElement.After;
        public bool UsesStructure => Combinators.Any(c => c is Combinator.Adjacent or Combinator.General) || Compounds.Any(c => c.UsesStructure);
        public bool IsRoot => PseudoElement == PseudoElement.None && Compounds.Count == 1
                              && Compounds[0] is { Id: null, Classes.Count: 0, Attributes.Count: 0, IsLink: false, PseudoTexts.Count: 0 } only
                              && (only.Type is "html" or "body" || only.IsRoot);

        public static Selector? Parse(string text)
        {
            text = text.Trim();
            if (text.Length == 0 || text.Length > 512) return null;
            var selector = new Selector();
            int i = 0;
            Combinator pending = Combinator.None;
            bool sawSpace = false;
            while (i < text.Length)
            {
                char c = text[i];
                if (char.IsWhiteSpace(c)) { sawSpace = true; i++; continue; }
                if (c is '>' or '+' or '~')
                {
                    if (selector.Compounds.Count == 0 || pending != Combinator.None || selector.PseudoElement != PseudoElement.None) return null;
                    pending = c == '>' ? Combinator.Child : c == '+' ? Combinator.Adjacent : Combinator.General;
                    sawSpace = false;
                    i++;
                    continue;
                }
                if (selector.Compounds.Count > 0)
                {
                    if (pending == Combinator.None && !sawSpace) return null;
                    if (selector.PseudoElement != PseudoElement.None) return null;     // a pseudo-element ends the selector
                }
                Compound? compound = ParseCompound(text, ref i, selector, allowsPseudoElement: true, out int specificity);
                if (compound is null) return null;
                selector.Specificity += specificity;
                selector.Combinators.Add(selector.Compounds.Count == 0 ? Combinator.None : pending == Combinator.None ? Combinator.Descendant : pending);
                selector.Compounds.Add(compound);
                pending = Combinator.None;
                sawSpace = false;
            }
            if (pending != Combinator.None || selector.Compounds.Count == 0 || selector.Compounds.Count > 16) return null;
            // a later-sibling pseudo-class is resolved for the element itself only (the renderer knows its later siblings)
            for (int k = 0; k + 1 < selector.Compounds.Count; k++) if (selector.Compounds[k].NeedsLaterSiblings) return null;
            if (selector.PseudoElement != PseudoElement.None) selector.Specificity += 1;
            var written = new StringBuilder();
            for (int k = 0; k < selector.Compounds.Count; k++)
            {
                if (k > 0) written.Append(selector.Combinators[k] switch { Combinator.Child => " > ", Combinator.Adjacent => " + ", Combinator.General => " ~ ", _ => " " });
                written.Append(selector.Compounds[k].ToText(k == 0));
            }
            if (selector.PseudoElement != PseudoElement.None) written.Append("::").Append(PseudoElementName(selector.PseudoElement));
            selector.Text = written.ToString();
            return selector;
        }

        /// <summary>One compound from <paramref name="i"/> (type or *, then #id .class [attr] :pseudo ::pseudo-element), or null.</summary>
        private static Compound? ParseCompound(string text, ref int i, Selector? selector, bool allowsPseudoElement, out int specificity)
        {
            specificity = 0;
            var compound = new Compound();
            int start = i;
            if (text[i] == '*') i++;
            else if (IsNameCharacter(text[i]))
            {
                int s = i;
                while (i < text.Length && IsNameCharacter(text[i])) i++;
                compound.Type = text.Substring(s, i - s).ToLowerInvariant();
                specificity += 1;
            }
            while (i < text.Length && !char.IsWhiteSpace(text[i]) && text[i] is not ('>' or '+' or '~' or ',' or ')'))
            {
                char kind = text[i++];
                if (kind == '[')
                {
                    int close = text.IndexOf(']', i);
                    if (close < 0 || ParseAttribute(text.Substring(i, close - i)) is not { } test) return null;
                    compound.Attributes.Add(test);
                    specificity += 1 << 10;
                    i = close + 1;
                    continue;
                }
                if (kind == ':' && i < text.Length && text[i] == ':')
                {
                    i++;
                    string elementName = ReadName(text, ref i).ToLowerInvariant();
                    if (!allowsPseudoElement || selector is null || selector.PseudoElement != PseudoElement.None || !SetPseudoElement(selector, elementName)) return null;
                    continue;
                }
                string name = ReadName(text, ref i);
                if (name.Length == 0) return null;
                if (selector?.PseudoElement is not (null or PseudoElement.None)) return null;   // nothing after a pseudo-element
                if (kind == '.') { compound.Classes.Add(name); specificity += 1 << 10; }
                else if (kind == '#') { if (compound.Id is not null) return null; compound.Id = name; specificity += 1 << 20; }
                else if (kind == ':')
                {
                    string lower = name.ToLowerInvariant();
                    string? argument = null;
                    if (i < text.Length && text[i] == '(')
                    {
                        int close = MatchingParenthesis(text, i);
                        if (close < 0) return null;
                        argument = text.Substring(i + 1, close - i - 1);
                        i = close + 1;
                    }
                    // CSS 2's single-colon pseudo-elements
                    if (argument is null && lower is "before" or "after" or "first-letter" or "first-line")
                    {
                        if (!allowsPseudoElement || selector is null || !SetPseudoElement(selector, lower)) return null;
                        continue;
                    }
                    int pseudoSpecificity = 1 << 10;
                    switch (lower)
                    {
                        // a printed link is unvisited (:visited never matches); hover / focus / active never print
                        case "link" when argument is null: case "any-link" when argument is null: compound.IsLink = true; compound.PseudoTexts.Add("link"); break;
                        case "root" when argument is null: compound.IsRoot = true; compound.PseudoTexts.Add("root"); break;
                        case "first-child" when argument is null: compound.NthChild = new Nth(0, 1); compound.PseudoTexts.Add("first-child"); break;
                        case "last-child" when argument is null: compound.NthLastChild = new Nth(0, 1); compound.PseudoTexts.Add("last-child"); break;
                        case "only-child" when argument is null:
                            compound.NthChild = new Nth(0, 1); compound.NthLastChild = new Nth(0, 1); compound.PseudoTexts.Add("only-child"); break;
                        case "first-of-type" when argument is null: compound.NthOfType = new Nth(0, 1); compound.PseudoTexts.Add("first-of-type"); break;
                        case "last-of-type" when argument is null: compound.NthLastOfType = new Nth(0, 1); compound.PseudoTexts.Add("last-of-type"); break;
                        case "only-of-type" when argument is null:
                            compound.NthOfType = new Nth(0, 1); compound.NthLastOfType = new Nth(0, 1); compound.PseudoTexts.Add("only-of-type"); break;
                        case "empty" when argument is null: compound.IsEmpty = true; compound.PseudoTexts.Add("empty"); break;
                        case "nth-child" or "nth-last-child" or "nth-of-type" or "nth-last-of-type" when argument is not null:
                        {
                            if (Nth.Parse(argument) is not { } nth) return null;
                            if (lower == "nth-child") compound.NthChild = nth;
                            else if (lower == "nth-last-child") compound.NthLastChild = nth;
                            else if (lower == "nth-of-type") compound.NthOfType = nth;
                            else compound.NthLastOfType = nth;
                            compound.PseudoTexts.Add(lower + "(" + nth + ")");
                            break;
                        }
                        case "not" when argument is not null:
                        {
                            int most = 0;
                            var written = new List<string>();
                            foreach (string part in SplitSelectorList(argument))
                            {
                                string trimmed = part.Trim();
                                int at = 0;
                                if (trimmed.Length == 0) return null;
                                Compound? inner = ParseCompound(trimmed, ref at, null, allowsPseudoElement: false, out int innerSpecificity);
                                if (inner is null || at != trimmed.Length || !inner.IsSimpleArgument) return null;
                                compound.Not.Add(inner);
                                written.Add(inner.ToText(true));
                                most = Math.Max(most, innerSpecificity);
                            }
                            compound.PseudoTexts.Add("not(" + string.Join(", ", written) + ")");
                            pseudoSpecificity = most;
                            break;
                        }
                        default: return null;                           // :visited, :hover, :checked... never print as written
                    }
                    specificity += pseudoSpecificity;
                }
                else return null;
            }
            if (i == start) return null;
            return compound;
        }

        private static bool SetPseudoElement(Selector selector, string name)
        {
            selector.PseudoElement = name switch
            {
                "before" => PseudoElement.Before, "after" => PseudoElement.After, "first-letter" => PseudoElement.FirstLetter, "first-line" => PseudoElement.FirstLine,
                _ => PseudoElement.None,
            };
            return selector.PseudoElement != PseudoElement.None;
        }

        private static string ReadName(string text, ref int i)
        {
            int s = i;
            while (i < text.Length && IsNameCharacter(text[i])) i++;
            return text.Substring(s, i - s);
        }

        private static int MatchingParenthesis(string text, int open)
        {
            int depth = 0;
            for (int k = open; k < text.Length; k++)
            {
                if (text[k] == '(') depth++;
                else if (text[k] == ')' && --depth == 0) return k;
            }
            return -1;
        }

        /// <summary>The inside of [...]: name, optional operator and value (ident or quoted, no escapes), optional i / s flag.</summary>
        private static AttributeTest? ParseAttribute(string inside)
        {
            string t = inside.Trim();
            int i = 0;
            while (i < t.Length && (IsNameCharacter(t[i]) || t[i] == ':')) i++;
            if (i == 0) return null;
            string name = t.Substring(0, i).ToLowerInvariant();
            while (i < t.Length && char.IsWhiteSpace(t[i])) i++;
            if (i == t.Length) return new AttributeTest(name, AttributeOperator.Exists, "", false);
            AttributeOperator op;
            if (t[i] == '=') { op = AttributeOperator.Equals; i++; }
            else if (i + 1 < t.Length && t[i + 1] == '=')
            {
                op = t[i] switch { '~' => AttributeOperator.Word, '|' => AttributeOperator.DashPrefix, '^' => AttributeOperator.Prefix, '$' => AttributeOperator.Suffix, '*' => AttributeOperator.Substring, _ => (AttributeOperator)255 };
                if (op == (AttributeOperator)255) return null;
                i += 2;
            }
            else return null;
            while (i < t.Length && char.IsWhiteSpace(t[i])) i++;
            if (i >= t.Length) return null;
            string value;
            if (t[i] is '"' or '\'')
            {
                char quote = t[i];
                int close = t.IndexOf(quote, i + 1);
                if (close < 0) return null;
                value = t.Substring(i + 1, close - i - 1);
                i = close + 1;
            }
            else
            {
                int s = i;
                while (i < t.Length && IsNameCharacter(t[i])) i++;
                value = t.Substring(s, i - s);
                if (value.Length == 0) return null;
            }
            if (value.IndexOfAny(['\\', '"', '\'', '<', '>', '{', '}', '\n', '\r']) >= 0) return null;   // escapes are not read; nothing unsafe is written back
            while (i < t.Length && char.IsWhiteSpace(t[i])) i++;
            bool ignoresCase = CaseInsensitiveAttributes.Contains(name);
            if (i < t.Length)
            {
                char flag = char.ToLowerInvariant(t[i]);
                if (flag == 'i') ignoresCase = true;
                else if (flag == 's') ignoresCase = false;
                else return null;
                i++;
                while (i < t.Length && char.IsWhiteSpace(t[i])) i++;
                if (i != t.Length) return null;
            }
            return new AttributeTest(name, op, value, ignoresCase);
        }
    }

    /// <summary>HTML attributes whose values selectors compare without case (HTML Living Standard 4.16.2, "case-sensitivity of selectors").</summary>
    private static readonly HashSet<string> CaseInsensitiveAttributes = new(StringComparer.Ordinal)
    {
        "accept", "accept-charset", "align", "alink", "axis", "bgcolor", "charset", "checked", "clear", "codetype", "color", "compact",
        "declare", "defer", "dir", "direction", "disabled", "enctype", "face", "frame", "hreflang", "http-equiv", "lang", "language",
        "link", "media", "method", "multiple", "nohref", "noresize", "noshade", "nowrap", "readonly", "rel", "rev", "rules", "scope",
        "scrolling", "selected", "shape", "target", "text", "type", "valign", "valuetype", "vlink",
    };

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

/// <summary>
/// One element's resolution state (<see cref="WebArchiveStylesheet.Open"/>): the selector prefixes it matches, those of it and its
/// ancestors that descendant combinators read, its position among its siblings, and - as a parent - the state of its element
/// children so far (count, count per type, the previous child's and all earlier children's matches that sibling combinators
/// read). The index pass writes it for the elements open at a block start into the block's open-tag stack
/// (<see cref="Serialize"/>), so a block resolves without the blocks before it.
/// </summary>
internal sealed class WebArchiveMatch(int[] matched, int[] ancestorOrSelf, WebArchiveMatch? parent, int index, int typeIndex)
{
    // an index-pass state matched for sources only (WebArchiveStylesheet.Open sourcesOnly): the sheet and inputs of its full match
    private WebArchiveStylesheet? _deferredSheet;
    private WebArchiveMatchContext _deferredContext;
    private int[] _matched = matched;

    public WebArchiveMatch(int[] sourceMatched, int[] ancestorOrSelf, WebArchiveMatch? parent, int index, int typeIndex, WebArchiveStylesheet sheet, WebArchiveMatchContext context)
        : this(sourceMatched, ancestorOrSelf, parent, index, typeIndex)
    {
        _deferredSheet = sheet;
        _deferredContext = context;
    }

    /// <summary>The attribute that carries the state on an open-tag stack entry (written by the index pass, never by a document).</summary>
    public const string StateAttribute = "data-fdv-s";
    /// <summary>The stack entry that carries the body's state (the children of body are blocks: their sibling state crosses blocks).</summary>
    public const string RootTagName = "fdv-root";

    /// <summary>
    /// The prefixes the element matches; for an index-pass state only the ones combinators read (<see cref="FullMatched"/> has
    /// all). Either way it holds every source prefix, which is all that the combinators of later elements look up.
    /// </summary>
    public int[] Matched => _matched;

    /// <summary>Every prefix the element matches (an index-pass state works it out now, once).</summary>
    public int[] FullMatched
    {
        get
        {
            if (_deferredSheet is { } sheet)
            {
                _matched = sheet.MatchAll(_deferredContext);
                _deferredSheet = null;
                _deferredContext = default;
            }
            return _matched;
        }
    }
    public int[] AncestorOrSelf { get; private set; } = ancestorOrSelf;
    public WebArchiveMatch? Parent { get; } = parent;
    public int Index { get; } = index;
    public int TypeIndex { get; } = typeIndex;

    public int Children { get; private set; }
    // children per element type: parallel arrays, copied before a change once a snapshot (Freeze) shares them
    private string[] _typeNames = [];
    private int[] _typeCounts = [];
    private int _typeCount;
    private bool _isTypeShared;
    public int[] PreviousChild { get; private set; } = [];
    public int[] EarlierChildren { get; private set; } = [];
    /// <summary>Renderer: text or an element child seen inside (for :empty).</summary>
    public bool HasContent;
    /// <summary>Index pass: a block started while the element was open (its end facts are recorded, <see cref="WebArchiveElementEnds"/>).</summary>
    public bool SpansBlocks;

    public int TypeCount(string name)
    {
        for (int i = 0; i < _typeCount; i++) if (string.Equals(_typeNames[i], name, StringComparison.Ordinal)) return _typeCounts[i];
        return 0;
    }

    public void CountChild(string name, int[] siblingMatches)
    {
        Children++;
        HasContent = true;
        if (_isTypeShared)
        {
            _typeNames = (string[])_typeNames.Clone();
            _typeCounts = (int[])_typeCounts.Clone();
            _isTypeShared = false;
        }
        int at = 0;
        while (at < _typeCount && !string.Equals(_typeNames[at], name, StringComparison.Ordinal)) at++;
        if (at == _typeCount)
        {
            if (_typeCount == _typeNames.Length)
            {
                Array.Resize(ref _typeNames, Math.Max(4, _typeCount * 2));
                Array.Resize(ref _typeCounts, Math.Max(4, _typeCount * 2));
            }
            _typeNames[at] = name;
            _typeCounts[at] = 0;
            _typeCount++;
        }
        _typeCounts[at]++;
        PreviousChild = siblingMatches;
        if (siblingMatches.Length > 0) EarlierChildren = MergeSorted(EarlierChildren, siblingMatches);
    }

    /// <summary>The sorted, distinct union of two sorted id arrays; returns an input itself when the other adds nothing.</summary>
    internal static int[] MergeSorted(int[] a, int[] b)
    {
        if (b.Length == 0) return a;
        if (a.Length == 0) return b;
        int i = 0, j = 0, count = 0;
        while (i < a.Length || j < b.Length)
        {
            if (j >= b.Length || (i < a.Length && a[i] < b[j])) i++;
            else if (i >= a.Length || b[j] < a[i]) j++;
            else { i++; j++; }
            count++;
        }
        if (count == a.Length) return a;
        if (count == b.Length) return b;
        var merged = new int[count];
        i = j = count = 0;
        while (i < a.Length || j < b.Length)
        {
            if (j >= b.Length || (i < a.Length && a[i] < b[j])) merged[count++] = a[i++];
            else if (i >= a.Length || b[j] < a[i]) merged[count++] = b[j++];
            else { merged[count++] = a[i++]; j++; }
        }
        return merged;
    }

    /// <summary>
    /// A copy of the state as it is now, which later children do not change (a block start's snapshot): it shares the arrays,
    /// which this state copies before its next change.
    /// </summary>
    public WebArchiveMatch Freeze()
    {
        _isTypeShared = true;
        return new WebArchiveMatch(_matched, AncestorOrSelf, Parent, Index, TypeIndex)
        {
            _deferredSheet = _deferredSheet, _deferredContext = _deferredContext,
            Children = Children, _typeNames = _typeNames, _typeCounts = _typeCounts, _typeCount = _typeCount, _isTypeShared = true,
            PreviousChild = PreviousChild, EarlierChildren = EarlierChildren, HasContent = HasContent, SpansBlocks = SpansBlocks,
        };
    }

    /// <summary>The state as an attribute value: "m,ids;a,ids;c,count;t,name=count,...;v,ids;e,ids;i,index,typeIndex" (digits, names, punctuation only).</summary>
    public string Serialize()
    {
        var text = new StringBuilder(64);
        AppendSerialized(text);
        return text.ToString();
    }

    /// <summary>Appends <see cref="Serialize"/>'s text.</summary>
    public void AppendSerialized(StringBuilder text)
    {
        text.Append("m,"); AppendIds(text, FullMatched);
        text.Append(";a,"); AppendIds(text, AncestorOrSelf);
        text.Append(";c,").Append(Children).Append(";t,");
        if (_typeCount > 0)
        {
            var order = new int[_typeCount];
            for (int i = 0; i < order.Length; i++) order[i] = i;
            string[] names = _typeNames;
            Array.Sort(order, (x, y) => string.CompareOrdinal(names[x], names[y]));
            for (int i = 0; i < order.Length; i++)
            {
                if (i > 0) text.Append('.');
                text.Append(_typeNames[order[i]]).Append('=').Append(_typeCounts[order[i]]);
            }
        }
        text.Append(";v,"); AppendIds(text, PreviousChild);
        text.Append(";e,"); AppendIds(text, EarlierChildren);
        text.Append(";i,").Append(Index).Append('.').Append(TypeIndex);
    }

    /// <summary>What a later-sibling pseudo-class needs of the element once it has ended: "children;content;name=count.name=count" (<see cref="WebArchiveElementEnds"/>).</summary>
    public string EndFacts()
    {
        var text = new StringBuilder(32);
        text.Append(Children).Append(';').Append(HasContent ? '1' : '0').Append(';');
        for (int i = 0; i < _typeCount; i++)
        {
            if (i > 0) text.Append('.');
            text.Append(_typeNames[i]).Append('=').Append(_typeCounts[i]);
        }
        return text.ToString();
    }

    private static void AppendIds(StringBuilder text, int[] ids)
    {
        for (int i = 0; i < ids.Length; i++)
        {
            if (i > 0) text.Append('.');
            text.Append(ids[i]);
        }
    }

    /// <summary>A state read back from <see cref="Serialize"/>; null when damaged.</summary>
    public static WebArchiveMatch? Deserialize(string text, WebArchiveMatch? parent)
    {
        int[] matched = [], ancestorOrSelf = [], previous = [], earlier = [];
        int children = 0, index = 1, typeIndex = 1;
        var typeNames = new List<string>();
        var typeCounts = new List<int>();
        foreach (string part in text.Split(';'))
        {
            if (part.Length < 2 || part[1] != ',') return null;
            string value = part.Substring(2);
            switch (part[0])
            {
                case 'm': if (Ids(value) is not { } m) return null; matched = m; break;
                case 'a': if (Ids(value) is not { } a) return null; ancestorOrSelf = a; break;
                case 'c': if (!int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out children)) return null; break;
                case 'v': if (Ids(value) is not { } v) return null; previous = v; break;
                case 'e': if (Ids(value) is not { } e) return null; earlier = e; break;
                case 'i':
                {
                    string[] pair = value.Split('.');
                    if (pair.Length != 2 || !int.TryParse(pair[0], NumberStyles.None, CultureInfo.InvariantCulture, out index) || !int.TryParse(pair[1], NumberStyles.None, CultureInfo.InvariantCulture, out typeIndex)) return null;
                    break;
                }
                case 't':
                    if (value.Length == 0) break;
                    foreach (string entry in value.Split('.'))
                    {
                        int equals = entry.LastIndexOf('=');
                        if (equals <= 0 || !int.TryParse(entry.Substring(equals + 1), NumberStyles.None, CultureInfo.InvariantCulture, out int count)) return null;
                        string typeName = entry.Substring(0, equals);
                        if (typeNames.Contains(typeName)) return null;
                        typeNames.Add(typeName);
                        typeCounts.Add(count);
                    }
                    break;
                default: return null;
            }
        }
        return new WebArchiveMatch(matched, ancestorOrSelf, parent, index, typeIndex)
            { Children = children, _typeNames = typeNames.ToArray(), _typeCounts = typeCounts.ToArray(), _typeCount = typeNames.Count,
              PreviousChild = previous, EarlierChildren = earlier, HasContent = children > 0 };
    }

    private static int[]? Ids(string value)
    {
        if (value.Length == 0) return [];
        string[] parts = value.Split('.');
        var ids = new int[parts.Length];
        for (int i = 0; i < parts.Length; i++)
            if (!int.TryParse(parts[i], NumberStyles.None, CultureInfo.InvariantCulture, out ids[i])) return null;
        return ids;
    }
}

/// <summary>
/// The selector state of body and of the elements open at a block start (<see cref="MarkupBlockScanner.CurrentStructure"/>):
/// frozen states, written out as text ("|"-joined <see cref="WebArchiveMatch.Serialize"/>, body first, "" for an entry without
/// a state) only when asked.
/// </summary>
internal sealed class WebArchiveStructure
{
    public static readonly WebArchiveStructure Empty = new("");
    private readonly WebArchiveMatch?[]? _entries;
    private string? _text;

    public WebArchiveStructure(string text) => _text = text;
    public WebArchiveStructure(WebArchiveMatch?[] entries) => _entries = entries;

    public string Text => _text ??= Write();

    private string Write()
    {
        var text = new StringBuilder(64 * _entries!.Length);
        for (int i = 0; i < _entries.Length; i++)
        {
            if (i > 0) text.Append('|');
            _entries[i]?.AppendSerialized(text);
        }
        return text.ToString();
    }
}

/// <summary>The inputs of an element's match: the element, its position, and its parent's state as it was when the element opened.</summary>
internal readonly record struct WebArchiveMatchContext(WebArchiveElement Element, int Index, int TypeIndex, bool HasParent, int[] ParentMatched,
                                                      int[] ParentAncestors, int[] ParentPreviousChild, int[] ParentEarlierChildren);

/// <summary>
/// The ends of the web archive elements that span blocks, recorded by the index pass when each one ends (its element children in
/// all, per type, and whether it has any content): what a block needs for :last-child, :nth-last-*, :only-of-type, :empty...
/// when an element's parent - or the element itself - goes on into blocks it does not read. Keyed by the element's path (its
/// child index under each open ancestor from body: "" is body, "/2/5" the fifth element child of body's second). Grows while
/// the index is built (a partial index shares the builder's), read under a lock.
/// </summary>
internal sealed class WebArchiveElementEnds
{
    private readonly Dictionary<string, string> _ends = new(StringComparer.Ordinal);

    public int Count { get { lock (_ends) return _ends.Count; } }

    public void Add(string path, string facts) { lock (_ends) _ends[path] = facts; }

    /// <summary>The element's end: element children in all, whether it has content, and the count of its children of one type.</summary>
    public (int Children, bool HasContent, int OfType)? Find(string path, string typeName)
    {
        string? facts;
        lock (_ends) if (!_ends.TryGetValue(path, out facts)) return null;
        string[] parts = facts.Split(';');
        if (parts.Length != 3 || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out int children)) return null;
        int ofType = 0;
        if (parts[2].Length > 0)
            foreach (string entry in parts[2].Split('.'))
            {
                int equals = entry.LastIndexOf('=');
                if (equals > 0 && string.CompareOrdinal(entry, 0, typeName, 0, Math.Max(equals, typeName.Length)) == 0 && equals == typeName.Length)
                    int.TryParse(entry.Substring(equals + 1), NumberStyles.None, CultureInfo.InvariantCulture, out ofType);
            }
        return (children, parts[1] == "1", ofType);
    }

    /// <summary>Path, facts, path, facts... (what the index stores).</summary>
    public string[] ToPairs()
    {
        lock (_ends)
        {
            var pairs = new List<string>(_ends.Count * 2);
            foreach (KeyValuePair<string, string> pair in _ends.OrderBy(p => p.Key, StringComparer.Ordinal)) { pairs.Add(pair.Key); pairs.Add(pair.Value); }
            return pairs.ToArray();
        }
    }

    public static WebArchiveElementEnds FromPairs(IReadOnlyList<string> pairs, int start)
    {
        var ends = new WebArchiveElementEnds();
        for (int i = start; i + 1 < pairs.Count; i += 2) ends._ends[pairs[i]] = pairs[i + 1];
        return ends;
    }
}

/// <summary>What an element's later-sibling pseudo-classes need, once known: its later element siblings (all, of its type) and whether it is empty.</summary>
internal readonly record struct WebArchiveLaterFacts(int LaterSiblings, int LaterSiblingsOfType, bool IsEmpty);

/// <summary>One element as the stylesheet sees it: name (lower case), id, classes, whether it is a link, and its start tag (attribute selectors).</summary>
internal sealed class WebArchiveElement
{
    private readonly string _tag;
    private string? _id;
    private string[]? _classes;
    private bool? _isLink;
    private bool _hasId;
    private string?[]? _attributes;             // name, value, name, value... (decoded)

    public WebArchiveElement(string name, string? id, string[] classes, bool isLink, string tag)
    {
        Name = name; _id = id; _hasId = true; _classes = classes; _isLink = isLink; _tag = tag;
    }

    /// <summary>An element read from its start tag: id, classes and link are read when first asked.</summary>
    private WebArchiveElement(string name, string tag) { Name = name; _tag = tag; }

    public string Name { get; }

    public string? Id
    {
        get
        {
            if (!_hasId) { _id = Decoded(MarkupBlockRenderer.AttributeValue(_tag, "id")); _hasId = true; }
            return _id;
        }
    }

    public string[] Classes => _classes ??= Decoded(MarkupBlockRenderer.AttributeValue(_tag, "class"))?.Split([' ', (char)9, (char)10, (char)12, (char)13], StringSplitOptions.RemoveEmptyEntries) ?? [];

    public bool IsLink => _isLink ??= (Name == "a" || Name == "area") && MarkupBlockRenderer.AttributeValue(_tag, "href") is not null;

    private static string? Decoded(string? value) => value is null ? null : (value.IndexOf('&') < 0 ? value : System.Net.WebUtility.HtmlDecode(value)).Trim() is { Length: > 0 } decoded ? decoded : null;

    /// <summary>An attribute's value as a browser reads it (entities decoded), or null when the tag has none.</summary>
    public string? Attribute(string attributeName)
    {
        // the tag's attributes read once, in one pass (names lower case, first occurrence wins), then looked up by name
        _attributes ??= ReadAttributes(_tag);
        for (int i = 0; i + 1 < _attributes.Length; i += 2)
            if (string.Equals(_attributes[i], attributeName, StringComparison.OrdinalIgnoreCase)) return _attributes[i + 1];
        return null;
    }

    private static string?[] ReadAttributes(string tag)
    {
        var pairs = new List<string?>(8);
        int i = 1;
        while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '>' && tag[i] != '/') i++;   // past the element name
        while (i < tag.Length)
        {
            while (i < tag.Length && (char.IsWhiteSpace(tag[i]) || tag[i] == '/')) i++;
            if (i >= tag.Length || tag[i] == '>') break;
            int nameStart = i;
            while (i < tag.Length && !char.IsWhiteSpace(tag[i]) && tag[i] != '=' && tag[i] != '>' && tag[i] != '/') i++;
            int nameLength = i - nameStart;
            string? value = null;
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
            else i = afterName;
            if (nameLength == 0) { i++; continue; }
            string name = tag.Substring(nameStart, nameLength);
            bool isKnown = false;
            for (int k = 0; k < pairs.Count; k += 2) if (string.Equals(pairs[k], name, StringComparison.OrdinalIgnoreCase)) { isKnown = true; break; }
            if (isKnown) continue;
            pairs.Add(name);
            pairs.Add(value is null ? "" : value.IndexOf('&') < 0 ? value : System.Net.WebUtility.HtmlDecode(value));
        }
        return pairs.ToArray();
    }

    /// <summary>The element of a start tag (its name already parsed).</summary>
    public static WebArchiveElement FromTag(string name, string tag) => new(name, tag);
}

/// <summary>
/// @media queries evaluated for print, as Edge prints a web archive (Media Queries 4): a query matches when its type is print or all
/// (not screen / speech / tty...; "not" inverts, "only" is ignored) and each feature holds on the printed page: orientation
/// portrait, color (8 bits), no monochrome, grid 0, the page box of Letter paper with Chromium's default margins of 1 cm
/// (width 8.5 in - 2 cm, height 11 in - 2 cm) for width / height and their min- / max- forms, aspect-ratio from those,
/// resolution 300 dpi. A feature this does not know makes its query false (as a browser treats an unknown feature).
/// </summary>
internal static class WebArchiveMedia
{
    private const double PageWidthPx = 8.5 * 96 - 2 * 96 / 2.54, PageHeightPx = 11 * 96 - 2 * 96 / 2.54;

    public static bool MatchesPrint(string queryList)
    {
        string list = queryList.Trim();
        if (list.Length == 0) return true;                               // @media { } : all
        foreach (string query in list.Split(','))
            if (MatchesQuery(query.Trim().ToLowerInvariant())) return true;
        return false;
    }

    private static bool MatchesQuery(string query)
    {
        if (query.Length == 0) return false;
        bool isNegated = false;
        if (query.StartsWith("not ", StringComparison.Ordinal)) { isNegated = true; query = query.Substring(4).Trim(); }
        else if (query.StartsWith("only ", StringComparison.Ordinal)) query = query.Substring(5).Trim();
        string type = "all";
        int at = 0;
        if (query.Length > 0 && query[0] != '(')
        {
            while (at < query.Length && !char.IsWhiteSpace(query[at]) && query[at] != '(') at++;
            type = query.Substring(0, at);
            query = query.Substring(at).Trim();
            if (query.StartsWith("and", StringComparison.Ordinal)) query = query.Substring(3).Trim();
            else if (query.Length > 0) return false;
        }
        bool result = type is "all" or "print";
        if (type is not ("all" or "print" or "screen" or "speech" or "tty" or "tv" or "projection" or "handheld" or "braille" or "embossed" or "aural")) return false;
        foreach (string feature in SplitFeatures(query))
        {
            if (feature.Length == 0) return false;
            if (Feature(feature) is not { } holds) return false;          // unknown or malformed: the query is false (not inverted)
            result &= holds;
        }
        return isNegated ? !result : result;
    }

    private static IEnumerable<string> SplitFeatures(string query)
    {
        int i = 0;
        while (i < query.Length)
        {
            while (i < query.Length && char.IsWhiteSpace(query[i])) i++;
            if (i >= query.Length) yield break;
            if (query[i] != '(') { yield return ""; yield break; }
            int close = query.IndexOf(')', i);
            if (close < 0) { yield return ""; yield break; }
            yield return query.Substring(i + 1, close - i - 1).Trim();
            i = close + 1;
            while (i < query.Length && char.IsWhiteSpace(query[i])) i++;
            if (i < query.Length)
            {
                if (!query.Substring(i).StartsWith("and", StringComparison.Ordinal)) { yield return ""; yield break; }
                i += 3;
            }
        }
    }

    private static bool? Feature(string feature)
    {
        int colon = feature.IndexOf(':');
        string name = (colon < 0 ? feature : feature.Substring(0, colon)).Trim();
        string? value = colon < 0 ? null : feature.Substring(colon + 1).Trim();
        switch (name)
        {
            case "orientation": return value is null ? true : value == "portrait" ? true : value == "landscape" ? false : null;
            case "color": return value is null ? true : Number(value) is { } bits ? bits == 8 : null;
            case "min-color": return Number(value) is { } minBits ? minBits <= 8 : null;
            case "max-color": return Number(value) is { } maxBits ? maxBits >= 8 : null;
            case "monochrome": return value is null ? false : Number(value) is { } mono ? mono == 0 : null;
            case "min-monochrome": return Number(value) is { } minMono ? minMono <= 0 : null;
            case "max-monochrome": return Number(value) is not null ? true : null;
            case "grid": return value is null ? false : Number(value) is { } grid ? grid == 0 : null;
            case "width": return value is null ? true : Length(value) is { } w ? Math.Abs(w - PageWidthPx) < 0.5 : null;
            case "min-width": return Length(value) is { } minWidth ? PageWidthPx >= minWidth : null;
            case "max-width": return Length(value) is { } maxWidth ? PageWidthPx <= maxWidth : null;
            case "height": return value is null ? true : Length(value) is { } h ? Math.Abs(h - PageHeightPx) < 0.5 : null;
            case "min-height": return Length(value) is { } minHeight ? PageHeightPx >= minHeight : null;
            case "max-height": return Length(value) is { } maxHeight ? PageHeightPx <= maxHeight : null;
            case "min-aspect-ratio": return Ratio(value) is { } minRatio ? PageWidthPx / PageHeightPx >= minRatio : null;
            case "max-aspect-ratio": return Ratio(value) is { } maxRatio ? PageWidthPx / PageHeightPx <= maxRatio : null;
            case "resolution": return value is null ? true : Resolution(value) is { } dpi ? Math.Abs(dpi - 300) < 0.5 : null;
            case "min-resolution": return Resolution(value) is { } minDpi ? 300 >= minDpi : null;
            case "max-resolution": return Resolution(value) is { } maxDpi ? 300 <= maxDpi : null;
            default: return null;
        }
    }

    private static double? Number(string? value)
        => value is not null && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : null;

    private static double? Length(string? value)
    {
        if (value is null) return null;
        int unit = 0;
        while (unit < value.Length && (char.IsDigit(value[unit]) || value[unit] is '.' or '-' or '+')) unit++;
        if (!double.TryParse(value.Substring(0, unit), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return null;
        return value.Substring(unit).Trim() switch
        {
            "px" => number, "pt" => number * 96 / 72, "pc" => number * 16, "in" => number * 96, "cm" => number * 96 / 2.54, "mm" => number * 96 / 25.4,
            "q" => number * 96 / 101.6, "em" or "rem" => number * 16, "" when number == 0 => 0, _ => null,
        };
    }

    private static double? Ratio(string? value)
    {
        if (value is null) return null;
        string[] parts = value.Split('/');
        if (parts.Length == 1) return Number(parts[0].Trim());
        return parts.Length == 2 && Number(parts[0].Trim()) is { } a && Number(parts[1].Trim()) is { } b && b > 0 ? a / b : null;
    }

    private static double? Resolution(string? value)
    {
        if (value is null) return null;
        if (value.EndsWith("dpi", StringComparison.Ordinal)) return Number(value.Substring(0, value.Length - 3));
        if (value.EndsWith("dpcm", StringComparison.Ordinal)) return Number(value.Substring(0, value.Length - 4)) * 2.54;
        if (value.EndsWith("dppx", StringComparison.Ordinal)) return Number(value.Substring(0, value.Length - 4)) * 96;
        if (value.EndsWith("x", StringComparison.Ordinal)) return Number(value.Substring(0, value.Length - 1)) * 96;
        return null;
    }
}

/// <summary>
/// The content property of ::before / ::after, kept only as strings and quote keywords (open-quote, close-quote, no-open-quote,
/// no-close-quote, none, normal): each string's CSS escapes are read and the text written back with every character outside
/// printable ASCII (and quotes, backslashes, angle brackets) as a CSS hex escape, so nothing in it can end the value or the rule.
/// counter(), attr(), url() and anything else drop the declaration.
/// </summary>
internal static class WebArchiveGeneratedContent
{
    public static (string Css, bool IsImportant)? Parse(string value)
    {
        string text = value.Trim();
        bool isImportant = false;
        if (text.EndsWith("!important", StringComparison.OrdinalIgnoreCase))
        {
            string before = text.Substring(0, text.Length - "!important".Length).TrimEnd();
            if (!InsideString(before)) { isImportant = true; text = before; }
        }
        var css = new StringBuilder();
        int i = 0;
        int tokens = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (char.IsWhiteSpace(c)) { i++; continue; }
            if (++tokens > 32) return null;
            if (css.Length > 0) css.Append(' ');
            if (c is '"' or '\'')
            {
                var decoded = new StringBuilder();
                i++;
                bool closed = false;
                while (i < text.Length)
                {
                    char d = text[i];
                    if (d == c) { closed = true; i++; break; }
                    if (d is '\n' or '\r') return null;
                    if (d == '\\')
                    {
                        i++;
                        if (i >= text.Length) return null;
                        int hex = 0, digits = 0;
                        while (i < text.Length && digits < 6 && Uri.IsHexDigit(text[i])) { hex = hex * 16 + Convert.ToInt32(text[i].ToString(), 16); i++; digits++; }
                        if (digits > 0)
                        {
                            if (i < text.Length && text[i] == ' ') i++;
                            if (hex is 0 or > 0x10FFFF || hex is >= 0xD800 and <= 0xDFFF) hex = 0xFFFD;
                            decoded.Append(char.ConvertFromUtf32(hex));
                        }
                        else if (text[i] is '\n' or '\r') i++;
                        else { decoded.Append(text[i]); i++; }
                        continue;
                    }
                    decoded.Append(d);
                    i++;
                }
                if (!closed || decoded.Length > 1024) return null;
                css.Append('"').Append(Escape(decoded.ToString())).Append('"');
                continue;
            }
            int start = i;
            while (i < text.Length && (char.IsLetter(text[i]) || text[i] == '-')) i++;
            string keyword = text.Substring(start, i - start).ToLowerInvariant();
            if (keyword is not ("open-quote" or "close-quote" or "no-open-quote" or "no-close-quote" or "none" or "normal")) return null;
            css.Append(keyword);
        }
        return css.Length == 0 ? null : (css.ToString(), isImportant);
    }

    private static bool InsideString(string text)
    {
        char quote = '\0';
        for (int i = 0; i < text.Length; i++)
        {
            char c = text[i];
            if (quote != '\0') { if (c == '\\') i++; else if (c == quote) quote = '\0'; }
            else if (c is '"' or '\'') quote = c;
        }
        return quote != '\0';
    }

    private static string Escape(string text)
    {
        var escaped = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c is >= ' ' and <= '~' && c is not ('"' or '\'' or '\\' or '<' or '>' or '{' or '}' or ';' or '@' or '/' or '*')) escaped.Append(c);
            else escaped.Append('\\').Append(((int)c).ToString("X", CultureInfo.InvariantCulture)).Append(' ');
        }
        return escaped.ToString();
    }

    /// <summary>The text a checked content value shows (its strings; quote keywords show no searchable text), or null for none / normal.</summary>
    public static string? TextOf(string checkedValue)
    {
        var text = new StringBuilder();
        int i = 0;
        bool any = false;
        while (i < checkedValue.Length)
        {
            if (checkedValue[i] != '"') { i++; continue; }
            any = true;
            i++;
            while (i < checkedValue.Length && checkedValue[i] != '"')
            {
                if (checkedValue[i] == '\\')
                {
                    int start = ++i;
                    while (i < checkedValue.Length && Uri.IsHexDigit(checkedValue[i])) i++;
                    text.Append(char.ConvertFromUtf32(Convert.ToInt32(checkedValue.Substring(start, i - start), 16)));
                    if (i < checkedValue.Length && checkedValue[i] == ' ') i++;
                    continue;
                }
                text.Append(checkedValue[i++]);
            }
            i++;
        }
        return any ? text.ToString() : null;
    }
}
