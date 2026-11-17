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
    /// Each marker's outer frame placement against Word 16's PDF of the same file (print layout; measured 2026-11-17, x from the
    /// 90 pt left margin, y from the anchor line's top): tbcolumn x 200; tbrtl and the group flush right (right edge 414-415 of
    /// 415); rotbox centred (bounding box centre 206.9); tbsquare x 300 (text both sides: floats to its nearer edge by design);
    /// tbdown 36 pt below its paragraph's top; line-aligned boxes: .docx top / centre / bottom on the line's box, .doc / RTF top
    /// below, centre across, bottom above the line's top edge. tbnote (page, x -50) and tbchar (character) stay at the column's edge.
    /// </summary>
    private static readonly (string File, string Marker, string Css)[] Expected =
    [
        ("floating-v1", "tbcolumn", "display:block;margin:0pt auto 0pt 200pt;"),   // wrap none, 200 pt from the column
        ("floating-v1", "tbnote", "display:block;margin:0pt auto 0pt 0pt;"),       // top and bottom, 40 pt from the page edge: left of the column
        ("floating-v1", "rotbox", "display:block;margin:0pt auto 0pt 146.9pt;"),   // centred (Word saved it as 146.9 pt = the centre), rotated 30 degrees
        ("floating-v1", "tbrtl", "display:block;margin:0pt 0 0pt auto;"),          // aligned right of the margin, right-to-left paragraph
        ("floating-v1", "tbsquare", "float:right;margin:0pt 0pt 0pt 18pt;"),       // square, 300 pt from the column (nearer right), 18 pt gap
        ("floating-v1", "tbchar", "float:left;margin:0pt 9pt 0pt 0pt;"),           // square, from the character: its own side, Word's 9 pt gap
        ("floating-v1", "tbdown", "float:left;margin:0pt 9pt 0pt 0pt;margin-top:36pt;"),   // square, 36 pt down from its paragraph
        ("floating-v1", "tblinetop", "display:block;margin:0pt auto 0pt 315pt;"),          // in front, top on the line: no shift
        ("floating-v1", "tblinecenter", "margin-top:calc(0pt + (LINE - 40pt) / 2);"),      // centred on the line
        ("floating-v1", "tblinebottom", "margin-top:calc(0pt + LINE - 40pt);"),            // bottom on the line
        ("floating-group-v1", "grpone", "display:block;margin:0pt 0 0pt auto;"),   // a group aligned right, top and bottom
        ("floating-group-v1", "grptwo", "display:block;margin:0pt 0 0pt auto;"),
    ];

    public static void Run()
    {
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-shapes");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf" })
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
                    // the line a line-aligned object is aligned to: a .docx line's box (1lh), a .doc / RTF line's top edge (0pt)
                    string wanted = css.Replace("LINE", extension == ".docx" ? "1lh" : "0pt");
                    TestReport.Check(style is not null && style.Contains(wanted) && (wanted.Contains("margin-top") || !style.Contains("margin-top")), $"{file.Key}{extension} '{marker}': frame {style ?? "(none)"}, expected {wanted}");
                    int shown = Occurrences(text, marker), found = SearchCount(path, marker);
                    TestReport.Check(shown == 1 && found == 1, $"{file.Key}{extension} '{marker}': shown {shown}, found {found}, expected 1 / 1");
                }
                if (file.Key == "floating-v1")
                {
                    // anchored late in its paragraph: from the paragraph's top (Word's PDF: 24 pt below the top line), so lifted to
                    // the paragraph's start; from its line: left where it is
                    TestReport.Check(StartsParagraph(html, "tblate"), $"{file.Key}{extension}: the paragraph-relative float anchored late is lifted to its paragraph's start");
                    TestReport.Check(!StartsParagraph(html, "tblateline"), $"{file.Key}{extension}: the line-relative float anchored late stays on its line");
                    TestReport.Check(Occurrences(text, "notebody") == 0 && html.Contains("note"), $"{file.Key}{extension}: the footnote in the text box's anchor paragraph stays a mark");
                    TestReport.Check(html.Contains("rotate(30"), $"{file.Key}{extension}: the rotated shape is drawn turned 30 degrees");
                }
            }
    }

    private static readonly Regex Tags = new("<[^>]+>", RegexOptions.CultureInvariant);

    /// <summary>Whether <paramref name="marker"/>'s outer frame is the first thing in its paragraph.</summary>
    private static bool StartsParagraph(string html, string marker)
    {
        int at = html.IndexOf(marker, StringComparison.Ordinal);
        if (at < 0) return false;
        int frame = -1;
        foreach (Match match in FrameOpen.Matches(html.Substring(0, at)))
            if (!match.Groups[2].Value.Contains("position:absolute")) frame = match.Index;
        int paragraph = frame < 0 ? -1 : html.LastIndexOf("<p", frame, StringComparison.Ordinal);
        return paragraph >= 0 && html.IndexOf('>', paragraph) + 1 == frame;
    }

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
