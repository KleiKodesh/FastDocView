using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// RTF tables as Word draws them: test-data/word-tables/tables-v1.rtf (saved by Word 16, make-tables-fixtures.ps1: a 3x3 table
/// with A1:B1 merged across and C2:C3 merged down, a table nested 3 deep, a right-to-left table, per-cell borders and fill,
/// vertical alignment in a 40 pt row), plus hand-written RTF for what Word never saves (\clmgf / \clmrg, \clshdng, a nested
/// table merged down). Each outer row is one block; a cell merged down is drawn without the line between its rows, the
/// rows it spans found at index time.
/// </summary>
public static class WordTablesFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Tables: RTF merges, nesting, right-to-left rows, cell borders, vertical alignment");
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-tables", "tables-v1.rtf"));
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-tables/tables-v1.rtf present"); return; }
        List<string> blocks = Render(path, "tables-rtf");
        string Row(string text) => blocks.FirstOrDefault(block => block.Contains(">" + text + "<", StringComparison.Ordinal)) ?? "";
        string Cell(string row, string text)
        {
            int at = row.IndexOf(">" + text + "<", StringComparison.Ordinal);
            int open = at < 0 ? -1 : row.LastIndexOf("<td", at, StringComparison.Ordinal);
            return open < 0 ? "" : row.Substring(open, row.IndexOf('>', open) - open + 1);
        }

        // merges: Word saves A1:B1 as one wide cell holding both paragraphs; C2:C3 as \clvmgf + \clvmrg
        string first = Row("m11");
        TestReport.Check(first.Contains(">m12<") && Regex.Matches(first, "<td").Count == 2, "RTF: A1:B1 merged across is one cell with both paragraphs (2 cells in the row)");
        string second = Row("m21"), third = Row("m31");
        TestReport.Check(Cell(second, "m23").Contains("border-bottom:none") && second.Contains(">m33<"), "RTF: the cell merged down (\\clvmgf) has no bottom line and holds the text of the cells below (Word moves it there)");
        TestReport.Check(third.Contains("class=\"merged-up\"") && Regex.IsMatch(third, "class=\"merged-up\" style=\"[^\"]*border-top:none"), "RTF: the continued cell (\\clvmrg) is drawn empty with no top line");
        TestReport.Check(!Cell(Row("m22"), "m22").Contains("border-bottom:none"), "RTF: cells not merged keep their bottom line");

        // nested 3 deep: drawn inside their cells, the 2-row inner table as one table
        string nested = Row("outer2");
        TestReport.Check(Regex.Matches(nested, "<table class=\"nested-table\"").Count == 2, "RTF: tables nested 3 deep are drawn nested (2 nested tables in the outer row)");
        int inner = nested.IndexOf("nested-table", StringComparison.Ordinal);
        TestReport.Check(inner >= 0 && nested.IndexOf("deep1", StringComparison.Ordinal) > inner && nested.IndexOf("inner21", StringComparison.Ordinal) > nested.IndexOf("inner12", StringComparison.Ordinal)
                         && Regex.Matches(nested, "<tr>").Count == 4, "RTF: the inner table keeps its 2 rows and the deepest table sits in its first cell");

        // right-to-left row
        TestReport.Check(Row("first").Contains("dir=\"rtl\""), "RTF: \\rtlrow draws the row right to left");

        // per-cell borders and fill
        TestReport.Check(Cell(Row("red-top"), "red-top").Contains("border-top:2.25pt solid #FF0000") && Cell(Row("red-top"), "red-top").Contains("border-left:none"),
                         "RTF: a cell's own top border (2.25 pt red) and no line where none is set");
        TestReport.Check(Cell(Row("yellow"), "yellow").Contains("#FFFF00"), "RTF: cell fill (\\clcbpat)");
        string box = Cell(Row("blue-box"), "blue-box");
        TestReport.Check(new[] { "top", "bottom", "left", "right" }.All(side => box.Contains("border-" + side + ":0.75pt double #0000FF")), "RTF: a cell boxed by double blue lines");

        // vertical alignment
        TestReport.Check(!Cell(Row("top"), "top").Contains("vertical-align") && Cell(Row("center"), "center").Contains("vertical-align:middle")
                         && Cell(Row("bottom"), "bottom").Contains("vertical-align:bottom"), "RTF: \\clvertalt / c / b");

        HandWritten();
    }

    /// <summary>What Word does not save itself: \clmgf + \clmrg, \clshdng over a fill, \trbrdr* without cell borders, a nested table merged down.</summary>
    private static void HandWritten()
    {
        const string rtf = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}{\colortbl;\red255\green0\blue0;\red255\green255\blue255;}
\trowd\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\clmgf\cellx2000\clmrg\cellx4000\clcbpat2\clshdng5000\cellx6000
\pard\intbl wide\cell gone\cell grey\cell\row
\pard\intbl\itap1 outer\cell
\pard\intbl\itap2 a1\nestcell b1\nestcell{\*\nesttableprops\trowd\clvmgf\cellx1000\cellx2000\nestrow}
\pard\intbl\itap2 \nestcell b2\nestcell{\*\nesttableprops\trowd\clvmrg\cellx1000\cellx2000\nestrow}
\pard\intbl\itap1 \cell\trowd\cellx3000\cellx6000\row
\pard after\par}";
        string path = Path.Combine(TestReport.OutputDirectory, "tables-handwritten-v1.rtf");
        File.WriteAllText(path, rtf, Encoding.ASCII);
        List<string> blocks = Render(path, "tables-handwritten");
        string row = blocks.FirstOrDefault(block => block.Contains(">wide<")) ?? "";
        TestReport.Check(row.Contains("colspan=\"2\"") && !row.Contains("gone") && Regex.IsMatch(row, "colspan=\"2\" style=\"width:200pt"), "RTF: \\clmgf + \\clmrg draw one cell spanning both widths; the continued cell's text is dropped");
        TestReport.Check(row.Contains("background:#808080"), "RTF: \\clshdng5000 over a white fill is mid grey");
        TestReport.Check(Regex.IsMatch(row, ">wide<.*?</td>", RegexOptions.Singleline) && row.Contains("border-top:0.5pt solid #000") && row.Contains("border-right:0.5pt solid #000"), "RTF: row borders (\\trbrdr*) where the cells name none");
        string nested = blocks.FirstOrDefault(block => block.Contains(">outer<")) ?? "";
        TestReport.Check(nested.Contains("rowspan=\"2\"") && nested.Contains(">a1<") && Regex.Matches(nested, "<td").Count == 5, "RTF: a nested cell merged down is one rowspan cell (3 nested cells + 2 outer)");
    }

    private static List<string> Render(string path, string cache)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-" + cache) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }
}
