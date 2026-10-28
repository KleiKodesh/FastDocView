using System.Globalization;

namespace WordDocumentRandomAccess.Html;

/// <summary>
/// The base CSS every rendered block relies on. Paragraph-specific values come inline from
/// <see cref="WordParagraphFormatting.ToCss"/>; this covers the shared skeleton: the document's
/// default font, zero default paragraph margins (Word's defaults are explicit per style), table
/// borders, and the markers for list items and footnote references.
/// </summary>
public static class WordDocumentHtmlStylesheet
{
    public static string Build(WordRunFormatting documentDefaultRun, bool isRightToLeft)
    {
        string font = documentDefaultRun.EffectiveFontName ?? (isRightToLeft ? "David" : "Calibri");
        double sizePt = documentDefaultRun.EffectiveFontSizeHalfPoints > 0 ? documentDefaultRun.EffectiveFontSizeHalfPoints / 2.0 : 11;
        return $$"""
            .word-document { font-family: '{{font.Replace("'", "")}}', 'Times New Roman', serif; font-size: {{sizePt.ToString("0.#", CultureInfo.InvariantCulture)}}pt; line-height: 1.15; color: #000; }
            .word-document p { margin: 0; padding: 0; white-space: pre-wrap; overflow-wrap: anywhere; min-height: 1em; }
            .word-document h1, .word-document h2, .word-document h3, .word-document h4, .word-document h5, .word-document h6 { margin: 0; font-size: inherit; font-weight: inherit; white-space: pre-wrap; }
            .word-document table { border-collapse: collapse; margin: 2pt 0; }
            .word-document td { border: 1px solid #999; padding: 1pt 4pt; vertical-align: top; }
            .word-document .list-marker { display: inline-block; min-width: 1.2em; }
            .word-document .footnote-reference { vertical-align: super; font-size: 0.7em; color: #06c; cursor: help; }
            .word-document .tab { display: inline-block; width: 2em; }
            .word-document .page-break { border-top: 1px dashed #bbb; margin: 6pt 0; }
            .word-document .table-cell-end { color: #999; }
            .word-document .unsupported { color: #999; font-style: italic; }
            """;
    }
}
