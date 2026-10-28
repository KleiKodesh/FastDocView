using System.IO.Compression;
using System.Xml;
using WordDocumentRandomAccess.Odt;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// The OpenDocument counterpart of <see cref="WordFormattingAudit"/>: streams content.xml and
/// styles.xml (or a .fodt) with XmlReader and counts what the ODT reader does not handle —
/// body-level elements without a block rule, inline elements the renderer does not know, and
/// attributes of style:text-properties / paragraph-properties / table properties the cascade does
/// not map. Anything ignored on purpose is listed below with its reason. Expected result: 0 unhandled.
/// </summary>
public static class OdtFormattingAudit
{
    // Body-level elements: rendered as blocks, flattened, transparent, or deliberately skipped.
    private static readonly HashSet<string> HandledBodyElements =
    [
        "text:p", "text:h", "table:table", "text:list", "text:list-item", "text:list-header", "text:section", "text:soft-page-break",
        "text:table-of-content", "text:alphabetical-index", "text:illustration-index", "text:table-index", "text:object-index", "text:user-index",
        "text:bibliography", "text:index-body", "text:index-title", "text:numbered-paragraph",
        // ignored by design: declarations (no visible text), revision history (the final text is shown), section links, forms
        "text:sequence-decls", "text:variable-decls", "text:user-field-decls", "text:dde-connection-decls", "text:tracked-changes",
        "text:section-source", "office:forms", "table:named-expressions",
        // ignored by design: index source templates are not content (only text:index-body is)
        "text:table-of-content-source", "text:alphabetical-index-source", "text:illustration-index-source", "text:table-index-source",
        "text:object-index-source", "text:user-index-source", "text:bibliography-source",
    ];

    // Ignored by design, with reasons: pagination/hyphenation (no pages), language tags (no spell checking),
    // East-Asian variants (no CJK script selection), kerning/scaling details, tab stops (CSS tab-size default), auto-direction flags.
    private static readonly HashSet<string> IgnoredProperties =
    [
        "fo:language", "fo:country", "style:language-asian", "style:country-asian", "style:language-complex", "style:country-complex",
        "fo:script", "style:script-asian", "style:script-complex", "style:rfc-language-tag", "style:rfc-language-tag-asian", "style:rfc-language-tag-complex",
        "style:font-name-asian", "fo:font-family-asian", "style:font-family-asian", "style:font-size-asian", "style:font-weight-asian", "style:font-style-asian",
        "style:font-family-generic", "style:font-pitch", "style:font-charset", "style:font-style-name", "style:font-family-generic-complex",
        "style:font-pitch-complex", "style:font-charset-complex", "style:font-style-name-complex", "style:font-family-generic-asian",
        "style:font-pitch-asian", "style:font-charset-asian", "style:font-style-name-asian",
        "fo:hyphenate", "fo:hyphenation-remain-char-count", "fo:hyphenation-push-char-count", "fo:hyphenation-ladder-count", "loext:hyphenation-no-caps",
        "loext:hyphenation-no-last-word", "loext:hyphenation-word-char-count", "loext:hyphenation-zone", "loext:hyphenation-keep",
        "loext:hyphenation-keep-type", "loext:hyphenation-compound-remain-char-count", "loext:hyphenation-zone-always",
        "loext:hyphenation-zone-column", "loext:hyphenation-zone-page", "loext:hyphenation-zone-spread", "loext:hyphenation-keep-line",
        "fo:widows", "fo:orphans", "fo:keep-with-next", "fo:keep-together", "fo:break-after", "style:page-number", "style:snap-to-layout-grid",
        "style:register-true", "style:line-break", "style:punctuation-wrap", "style:text-autospace", "style:writing-mode-automatic",
        "style:justify-single-word", "fo:text-align-last", "style:auto-text-indent", "style:join-border", "style:shadow", "style:tab-stop-distance",
        "style:vertical-align", "style:line-spacing", "style:contextual-spacing", "loext:contextual-spacing", "fo:padding-top", "fo:padding-bottom",
        "fo:padding-left", "fo:padding-right", "style:border-line-width", "style:border-line-width-top", "style:border-line-width-bottom",
        "style:border-line-width-left", "style:border-line-width-right", "draw:fill", "draw:fill-color", "loext:margin-left", "loext:margin-right",
        "loext:text-indent", "fo:hyphenation-keep", "style:letter-kerning", "style:text-scale", "style:use-window-font-color",
        "style:text-underline-mode", "style:text-line-through-mode", "style:text-line-through-width", "style:text-line-through-color",
        "style:text-line-through-text", "style:text-overline-style", "style:text-overline-width", "style:text-overline-color", "style:text-overline-mode",
        "style:text-emphasize", "style:text-combine", "style:text-rotation-angle", "style:text-rotation-scale", "style:text-blinking",
        "officeooo:rsid", "officeooo:paragraph-rsid", "loext:char-shading-value", "loext:opacity", "style:font-relief-color",
        "fo:hyphenation-ladder-count", "loext:graphic", "loext:padding", "loext:border", "loext:padding-left", "loext:padding-right",
        "loext:padding-top", "loext:padding-bottom", "loext:border-left", "loext:border-right", "loext:border-top", "loext:border-bottom",
        "fo:min-height", "style:shrink-to-fit", "style:may-break-between-rows", "table:border-model", "style:page-number", "fo:break-before",
        "style:min-row-height", "style:row-height", "fo:keep-together", "style:repeat-content", "style:cell-protect", "style:print-content",
        "style:decimal-places", "style:rotation-align", "style:direction", "style:glyph-orientation-vertical", "style:text-align-source",
        "fo:wrap-option", "style:use-optimal-column-width", "style:use-optimal-row-height", "fo:margin-top", "fo:margin-bottom", "fo:margin",
        "style:diagonal-bl-tr", "style:diagonal-tl-br", "loext:writing-mode",
        // justification tuning (Writer's paragraph composer, word-spacing limits) and line metrics: no CSS equivalent
        "loext:paragraph-composer", "loext:word-spacing-maximum", "loext:word-spacing-minimum", "style:font-independent-line-spacing",
        // tab stops are not mapped (the stylesheet's default tab size applies); line numbering belongs to a page margin
        "loext:tab-stop-distance", "text:line-number", "text:number-lines",
        // OpenType optical sizing, hyphenation detail, the script-type hint (western/complex is chosen by paragraph direction)
        "loext:font-optical-sizing", "loext:hyphenation-compound-push-char-count", "style:script-type",
    ];

    private sealed class Tally
    {
        public readonly SortedDictionary<string, long> Counts = new(StringComparer.Ordinal);
        public void Add(string key) => Counts[key] = Counts.TryGetValue(key, out long n) ? n + 1 : 1;
    }

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"OpenDocument audit: {Path.GetFileName(path)}");
            var tally = new Tally();
            if (WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.FlatOpenDocument)
            {
                using var file = File.OpenRead(path);
                AuditPart(file, tally);
            }
            else
            {
                using ZipArchive zip = ZipFile.OpenRead(path);
                foreach (string part in new[] { "styles.xml", "content.xml" })
                    if (zip.GetEntry(part) is { } entry) { using Stream stream = entry.Open(); AuditPart(stream, tally); }
            }
            foreach (KeyValuePair<string, long> entry in tally.Counts) Console.WriteLine($"      {entry.Value,8:N0}  {entry.Key}");
            TestReport.Check(tally.Counts.Count == 0, $"{tally.Counts.Count} unhandled OpenDocument constructs");
        }
    }

    private static void AuditPart(Stream stream, Tally tally)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore });
        var stack = new List<string>();
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement) { stack.RemoveAt(stack.Count - 1); continue; }
            if (reader.NodeType != XmlNodeType.Element) continue;
            string name = reader.Name;
            string parent = stack.Count > 0 ? stack[stack.Count - 1] : "";
            bool isInsideText = stack.Contains("office:text");
            bool isInsideParagraph = stack.Contains("text:p") || stack.Contains("text:h");
            bool isInsideSkipped = stack.Contains("text:tracked-changes") || stack.Contains("office:annotation") || stack.Contains("draw:frame")
                                   || stack.Any(s => s.EndsWith("-source", StringComparison.Ordinal));

            if (isInsideText && !isInsideSkipped)
            {
                if (isInsideParagraph)
                {
                    if (parent is "text:p" or "text:h" or "text:span" or "text:a" or "text:meta" or "text:ruby" or "text:ruby-base" or "text:note-citation"
                        && !OdtBlockHtmlRenderer.HandledInlineElements.Contains(name) && !HandledBodyElements.Contains(name))
                        tally.Add("inline <" + name + ">");
                }
                else if (parent is "office:text" or "text:section" or "text:list-item" or "text:list-header" or "text:index-body" or "text:list"
                         && !HandledBodyElements.Contains(name) && !name.StartsWith("table:", StringComparison.Ordinal))
                    tally.Add("block <" + name + "> in <" + parent + ">");
            }

            string? propertiesKind = name switch
            {
                "style:text-properties" => "text", "style:paragraph-properties" => "paragraph",
                "style:table-properties" or "style:table-cell-properties" or "style:table-column-properties" or "style:table-row-properties" => "table",
                _ => null,
            };
            if (propertiesKind is not null && reader.MoveToFirstAttribute())
            {
                do
                {
                    string attribute = reader.Name;
                    if (attribute.StartsWith("xmlns", StringComparison.Ordinal) || IgnoredProperties.Contains(attribute)) continue;
                    HashSet<string> handled = propertiesKind switch
                    {
                        "text" => OdtStyleCascade.HandledTextProperties,
                        "paragraph" => OdtStyleCascade.HandledParagraphProperties,
                        _ => OdtStyleCascade.HandledTableProperties,
                    };
                    if (!handled.Contains(attribute)) tally.Add(propertiesKind + "-properties @" + attribute);
                } while (reader.MoveToNextAttribute());
                reader.MoveToElement();
            }
            if (!reader.IsEmptyElement) stack.Add(name);
        }
    }
}
