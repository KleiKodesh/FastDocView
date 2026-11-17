using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .doc / RTF floating text boxes, groups and shapes placed as Word places them: test-data/word-shapes/floating-v1.doc / .rtf
/// (built by Word 16 through make-floating-fixtures.ps1) against Word's own print of each file (floating-v1-doc-word.pdf /
/// floating-v1-rtf-word.pdf, read when the fixture was made): a text box at an offset from the column with an 18 pt gap, a group
/// aligned right, a rotated shape centred above and below the text, a text box in a paragraph with a footnote, one in a
/// right-to-left paragraph and one placed from the character.
/// </summary>
public static class FloatingShapesFidelityTest
{
    private static readonly Regex FrameOpen = new("<span class=\"(word-shape|text-box)[^\"]*\" style=\"([^\"]*)\"", RegexOptions.CultureInvariant);

    /// <summary>
    /// Each marker's outer frame placement, as the file says (Word 16's COM read-back when the fixture was made: Left,
    /// RelativeHorizontalPosition, WrapFormat). Word's PDF of the same files (2026-11-17): tbcolumn (wrap none) prints at x 290 =
    /// the 90 pt margin + 200, rotbox centred; this Word prints every wrapped object (square, top and bottom, also in a one-shape
    /// .docx Word wrote itself) centred on the page whatever its position, so those are checked against the file's own values.
    /// </summary>
    private static readonly (string File, string Marker, string Css)[] Expected =
    [
        ("floating-v1", "tbcolumn", "display:block;margin:0pt auto 0pt 200pt;"),   // wrap none, 200 pt from the column
        ("floating-v1", "tbnote", "display:block;margin:0pt auto 0pt 0pt;"),       // top and bottom, 40 pt from the page edge: left of the column
        ("floating-v1", "rotbox", "display:block;margin:0pt auto 0pt auto;"),      // centred, rotated 30 degrees
        ("floating-v1", "tbrtl", "display:block;margin:0pt 0 0pt auto;"),          // aligned right of the margin, right-to-left paragraph
        ("floating-v1", "tbsquare", "float:right;margin:0pt 0pt 0pt 18pt;"),       // square, 300 pt from the column (nearer right), 18 pt gap
        ("floating-v1", "tbchar", "float:left;margin:0pt 9pt 0pt 0pt;"),           // square, from the character: its own side, Word's 9 pt gap
        ("floating-group-v1", "grpone", "display:block;margin:0pt 0 0pt auto;"),   // a group aligned right, top and bottom
        ("floating-group-v1", "grptwo", "display:block;margin:0pt 0 0pt auto;"),
    ];

    public static void Run()
    {
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-shapes");
        foreach (string extension in new[] { ".doc", ".rtf" })
            foreach (var file in Expected.GroupBy(e => e.File))
            {
                string path = Path.Combine(folder, file.Key + extension);
                TestReport.Step(".doc / RTF floating objects placed: " + file.Key + extension);
                if (!File.Exists(path)) { TestReport.Check(false, path + " is missing"); continue; }
                string html = string.Concat(Render(path));
                string text = Tags.Replace(html, "");
                foreach (var (_, marker, css) in file)
                {
                    string? style = FrameStyle(html, marker);
                    TestReport.Check(style is not null && style.Contains(css), $"{file.Key}{extension} '{marker}': frame {style ?? "(none)"}, expected {css}");
                    int shown = Occurrences(text, marker), found = SearchCount(path, marker);
                    TestReport.Check(shown == 1 && found == 1, $"{file.Key}{extension} '{marker}': shown {shown}, found {found}, expected 1 / 1");
                }
                if (file.Key == "floating-v1")
                {
                    TestReport.Check(Occurrences(text, "notebody") == 0 && html.Contains("note"), $"{file.Key}{extension}: the footnote in the text box's anchor paragraph stays a mark");
                    TestReport.Check(html.Contains("rotate(30"), $"{file.Key}{extension}: the rotated shape is drawn turned 30 degrees");
                }
            }
    }

    private static readonly Regex Tags = new("<[^>]+>", RegexOptions.CultureInvariant);

    /// <summary>The style of the outermost frame (shape or text box) holding <paramref name="marker"/>'s first occurrence.</summary>
    private static string? FrameStyle(string html, string marker)
    {
        int at = html.IndexOf(marker, StringComparison.Ordinal);
        if (at < 0) return null;
        string? style = null;
        foreach (Match match in FrameOpen.Matches(html.Substring(0, at)))
            if (!match.Groups[2].Value.Contains("position:absolute")) style = match.Groups[1].Value + " | " + match.Groups[2].Value;   // not a shape's own text over it
        return style;
    }

    private static int SearchCount(string path, string query)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        return document.Search(query, new WordDocumentSearchOptions(), 0, document.BlockCount).HitCounts.Sum();
    }

    private static int Occurrences(string text, string part)
    {
        int count = 0;
        for (int at = text.IndexOf(part, StringComparison.Ordinal); at >= 0; at = text.IndexOf(part, at + part.Length, StringComparison.Ordinal)) count++;
        return count;
    }

    private static List<string> Render(string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        var html = new List<string>();
        for (int first = 0; first < document.BlockCount; first += 100)
            html.AddRange(document.RenderBlocks(first, 100).Select(block => block.Html));
        return html;
    }
}
