using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// The base CSS every rendered block relies on. Paragraph-specific values come inline from
/// <see cref="WordParagraphFormatting.ToCss"/>; this covers the shared skeleton: the document's
/// default font, zero default paragraph margins (Word's defaults are explicit per style), table
/// borders, and the markers for list items and footnote references.
/// Its own colours are <c>var(--fdv-*, fallback)</c>: a host maps them to its theme, a bare page gets Word's look.
/// </summary>
public static class WordDocumentHtmlStylesheet
{
    /// <summary>A copy of a .doc / RTF document's default run that asks for Word's shaping (no ligatures, no contextual
    /// alternates): those formats cannot set them, and Word draws neither.</summary>
    public static WordRunFormatting WithWordShaping(WordRunFormatting run)
    {
        WordRunFormatting copy = run.Clone();
        copy.UsesWordOpenTypeDefaults = true;
        return copy;
    }

    public static string Build(WordRunFormatting documentDefaultRun, bool isRightToLeft)
    {
        string font = documentDefaultRun.EffectiveFontName ?? (isRightToLeft ? "David" : "Calibri");
        double sizePt = documentDefaultRun.EffectiveFontSizeHalfPoints > 0 ? documentDefaultRun.EffectiveFontSizeHalfPoints / 2.0 : 11;
        // Word's shaping (.docx, .doc, RTF): no ligatures, no contextual alternates unless a .docx run sets them (WordRunFormatting.OpenTypeCss)
        string shaping = documentDefaultRun.UsesWordOpenTypeDefaults ? " " + WordRunFormatting.WordComplexScriptLigaturesCss : "";
        return $$"""
            .word-document { font-family: '{{font.Replace("'", "")}}', 'Times New Roman', serif; font-size: {{sizePt.ToString("0.#", CultureInfo.InvariantCulture)}}pt; line-height: 1.15; color: var(--fdv-text, #000);{{shaping}} }
            .word-document p, .word-document .word-paragraph-block { margin: 0; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; min-height: 1em; }
            .word-document h1, .word-document h2, .word-document h3, .word-document h4, .word-document h5, .word-document h6 { margin: 0; font-size: inherit; font-weight: inherit; white-space: pre-wrap; }
            .word-document table { border-collapse: collapse; margin: 0; }
            .word-document td { border: 1px solid var(--fdv-table-border, #999); padding: 1pt 4pt; vertical-align: top; }
            .word-document .list-marker { display: inline-block; text-indent: 0; }
            .word-document p, .word-document .word-paragraph-block, .word-document h1, .word-document h2, .word-document h3, .word-document h4, .word-document h5, .word-document h6 { tab-size: 36pt; }
            .word-document .tab-space { position: relative; tab-size: 0; }
            .word-document .tab-space[data-leader]::before { position: absolute; inset-inline: 0; top: 0; line-height: normal; overflow: hidden; white-space: nowrap; text-align: end; pointer-events: none; }
            .word-document .tab-space[data-leader="d"]::before { content: "{{TabLeaders.Dots}}"; }
            .word-document .tab-space[data-leader="h"]::before { content: "{{TabLeaders.Hyphens}}"; }
            .word-document .tab-space[data-leader="u"]::before, .word-document .tab-space[data-leader="H"]::before { content: "{{TabLeaders.Underscores}}"; }
            .word-document .tab-space[data-leader="H"]::before { font-weight: bold; }
            .word-document .tab-space[data-leader="m"]::before { content: "{{TabLeaders.MiddleDots}}"; }
            .word-document .tab-break::before { content: "\A"; white-space: pre; }
            .word-document .tab-bar { position: absolute; top: 0; bottom: 0; border-inline-start: 0.75pt solid currentColor; pointer-events: none; }
            .word-document .note-reference { font-size: 0.7em; line-height: 0; color: var(--fdv-link, #06c); cursor: pointer; padding: 0 1px; }
            .word-document .note-mark { vertical-align: super; font-size: 0.7em; line-height: 0; color: var(--fdv-link, #06c); margin-inline-end: 0.25em; }
            .word-document .page-break { display: block; border-top: 1px dashed var(--fdv-rule, #bbb); margin: 6pt 0; }
            .word-document .field-page-placeholder { color: var(--fdv-muted-strong, #888); border: 1px dotted var(--fdv-rule, #bbb); border-radius: 2px; padding: 0 0.15em; font-size: 0.9em; user-select: none; }
            .word-document .form-checkbox { font-family: 'Segoe UI Symbol', sans-serif; }
            .word-document math.field-equation { font-family: inherit; }
            .word-document .text-box { display: inline-block; vertical-align: top; max-width: 100%; border: 1px solid var(--fdv-frame-border, #c8c8c8); padding: 2pt 6pt; margin: 2pt; }
            .word-document .section-break { border: 0; border-top: 3px double var(--fdv-rule, #bbb); margin: 8pt 0; }
            .word-document table.table-row { width: 100%; margin: 0; }
            .word-document table.table-row + table.table-row td { border-top: 0; }
            .word-document .word-picture { max-width: 100%; height: auto; box-sizing: border-box; }
            .word-document .text-box.word-shape { position: relative; }
            .word-document .text-box.word-shape .word-shape-geometry { position: absolute; left: 0; top: 0; width: 100%; height: 100%; overflow: visible; }
            .word-document .text-box.word-shape .word-shape-part { position: absolute; display: block; }
            .word-document .text-box.word-shape .word-shape-text { position: relative; display: block; }
            .word-document .word-picture-missing { display: inline-block; border: 1px dashed var(--fdv-rule, #bbb); color: var(--fdv-muted, #999); font-size: 9pt; text-align: center; overflow: hidden; }
            .word-document .unsupported { color: var(--fdv-muted, #999); font-style: italic; }
            .word-document .{{WordHtmlColor.DarkTextOnFillClass}} { color: light-dark(currentColor, #000); }
            .word-document .drop-cap { float: inline-start; font-size: calc(var(--drop-cap-lines) * 1.15em); line-height: 1; margin-block-start: 0.05em; }
            .word-document .drop-cap.drop-cap-in-margin { margin-inline-start: calc(-1 * var(--drop-cap-width, 0.8em) - var(--drop-cap-gap, 0pt)); }
            .word-document a.link { color: inherit; text-decoration: none; cursor: pointer; }
            """;
    }
}

/// <summary>The leader fills of a laid-out tab (the page's tab pass sets data-leader): long enough for any line, clipped to the tab's width.</summary>
internal static class TabLeaders
{
    public static readonly string Dots = new('.', 400);
    public static readonly string Hyphens = new('-', 400);
    public static readonly string Underscores = new('_', 400);
    public static readonly string MiddleDots = string.Concat(Enumerable.Repeat("\\0000B7", 400));    // CSS escapes (U+00B7)
}
