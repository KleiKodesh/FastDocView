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
        RowAwaitingNextIsProvisional();
        Doc();
        Extra();
    }

    /// <summary>
    /// .doc tables from the same Word 16 fixture (tables-v1.doc): TC merges (fVertMerge / fVertRestart) with the rows merged down
    /// found at index time, nesting by sprmPItap drawn inside the cell, per-cell borders (TC brc*) over the table's
    /// (sprmTTableBorders), TC vertAlign. Word 16 never saves fFirstMerged / fMerged, so a copy with those bits set on row 2 checks
    /// the horizontal merge.
    /// </summary>
    private static void Doc()
    {
        TestReport.Step("Tables: .doc merges, nesting, cell borders, vertical alignment");
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-tables", "tables-v1.doc"));
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-tables/tables-v1.doc present"); return; }
        List<string> blocks = Render(path, "tables-doc");
        string Row(string text) => blocks.FirstOrDefault(block => block.Contains(">" + text + "<", StringComparison.Ordinal)) ?? "";
        string Cell(string row, string text)
        {
            int at = row.IndexOf(">" + text + "<", StringComparison.Ordinal);
            int open = at < 0 ? -1 : row.LastIndexOf("<td", at, StringComparison.Ordinal);
            return open < 0 ? "" : row.Substring(open, row.IndexOf('>', open) - open + 1);
        }

        TestReport.Check(Regex.Matches(Row("m11"), "<td").Count == 2 && Row("m11").Contains(">m12<"), ".doc: A1:B1 merged across is one cell (Word saves one wide cell)");
        TestReport.Check(Cell(Row("m21"), "m23").Contains("border-bottom:none") && Row("m21").Contains(">m33<"), ".doc: the cell merged down (fVertRestart) has no bottom line and holds the text Word moved there");
        TestReport.Check(Regex.IsMatch(Row("m31"), "class=\"merged-up\" style=\"[^\"]*border-top:none"), ".doc: the continued cell (fVertMerge) is drawn empty with no top line");
        TestReport.Check(!Cell(Row("m22"), "m22").Contains("border-bottom:none") && Cell(Row("m22"), "m22").Contains("border-top:0.5pt solid #000"), ".doc: other cells keep the table's lines (sprmTTableBorders)");

        string nested = Row("outer2");
        TestReport.Check(Regex.Matches(nested, "<table class=\"nested-table\"").Count == 2 && Regex.Matches(nested, "<tr>").Count == 4,
                         ".doc: tables nested 3 deep are drawn nested (2 nested tables, the inner one with its 2 rows)");
        TestReport.Check(nested.IndexOf("deep1", StringComparison.Ordinal) > nested.IndexOf("nested-table", StringComparison.Ordinal)
                         && nested.IndexOf("inner21", StringComparison.Ordinal) > nested.IndexOf("inner12", StringComparison.Ordinal)
                         && blocks.Count(block => block.Contains("inner21") || block.Contains("deep1")) == 1, ".doc: inner rows are not blocks of their own");

        TestReport.Check(Cell(Row("red-top"), "red-top").Contains("border-top:2.25pt solid #FF0000") && Cell(Row("red-top"), "red-top").Contains("border-left:none"),
                         ".doc: a cell's own top border (2.25 pt red) and no line where none is set");
        string box = Cell(Row("blue-box"), "blue-box");
        TestReport.Check(new[] { "top", "bottom", "left", "right" }.All(side => box.Contains("border-" + side + ":0.75pt double #0000FF")), ".doc: a cell boxed by double blue lines");
        TestReport.Check(Cell(Row("yellow"), "yellow").Contains("#FFFF00"), ".doc: cell fill (sprmTDefTableShd)");
        TestReport.Check(Row("first").Contains("dir=\"rtl\"") && Row("m11").Contains("dir=\"ltr\"") && Row("red-top").Contains("dir=\"ltr\""),
                         ".doc: column order from sprmTFBiDi (Word 16: table 3 right to left, the others left to right), not from the row-end paragraph");
        TestReport.Check(!Cell(Row("top"), "top").Contains("vertical-align") && Cell(Row("center"), "center").Contains("vertical-align:middle")
                         && Cell(Row("bottom"), "bottom").Contains("vertical-align:bottom"), ".doc: TC vertAlign top / centre / bottom");

        // fFirstMerged / fMerged: a copy of the fixture with row 2's first two cells merged across
        byte[] bytes = File.ReadAllBytes(path);
        byte[] rowTwo = [0x03, 0x05, 0x00, 0xD5, 0x0A, 0xA6, 0x15, 0x77, 0x20];
        int at = -1;
        for (int i = 0; i + 60 < bytes.Length && at < 0; i++)
            if (bytes.AsSpan(i, rowTwo.Length).SequenceEqual(rowTwo) && bytes[i + 49] == 0x60) at = i;
        TestReport.Check(at >= 0, ".doc: row 2's sprmTDefTable found in the fixture");
        if (at < 0) return;
        bytes[at + 9] |= 0x01;                           // TC 0: fFirstMerged
        bytes[at + 29] |= 0x02;                          // TC 1: fMerged
        string merged = Path.Combine(TestReport.OutputDirectory, "tables-merged-across-v1.doc");
        File.WriteAllBytes(merged, bytes);
        string row = Render(merged, "tables-doc-merged").FirstOrDefault(block => block.Contains(">m21<")) ?? "";
        TestReport.Check(row.Contains("colspan=\"2\"") && !row.Contains(">m22<") && Regex.Matches(row, "<td").Count == 2 && Regex.IsMatch(row, "colspan=\"2\" style=\"width:276.9pt"),
                         ".doc: fFirstMerged + fMerged draw one cell spanning both widths; the merged cell's text is dropped");
    }

    /// <summary>
    /// test-data/word-tables/tables-extra-v1.doc / .rtf (make-tables-extra-fixtures.ps1, Word 16; what Word reported is in
    /// tables-extra-word.txt): row heights exact 30 pt / at least 50 pt, an explicit "no border" cell, a border colour outside the
    /// ico palette (.doc sprmTBrc*Cv), two tables Word joined into one when the paragraph between them was deleted (4 rows). The
    /// .odt is LibreOffice's conversion of the .doc (style:row-height / min-row-height). sprmTSetBrc / sprmTSetBrc80, which Word
    /// 16 does not write, are checked on a copy of the .doc whose sprmTBrcTopCv / LeftCv are rewritten into them.
    /// </summary>
    private static void Extra()
    {
        TestReport.Step("Tables: row heights, explicit no border, border colours, joined tables, sprmTSetBrc");
        string doc = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-tables", "tables-extra-v1.doc"));
        string rtf = Path.ChangeExtension(doc, ".rtf");
        if (!File.Exists(doc) || !File.Exists(rtf)) { TestReport.Check(false, "test-data/word-tables/tables-extra-v1.doc / .rtf present"); return; }
        foreach (string path in new[] { doc, rtf })
        {
            string name = Path.GetExtension(path);
            List<string> blocks = Render(path, "tables-extra" + name);
            string Row(string text) => blocks.FirstOrDefault(block => block.Contains(">" + text + "<", StringComparison.Ordinal)) ?? "";
            string Cell(string row, string text)
            {
                int at = row.IndexOf(">" + text + "<", StringComparison.Ordinal);
                int open = at < 0 ? -1 : row.LastIndexOf("<td", at, StringComparison.Ordinal);
                return open < 0 ? "" : row.Substring(open, row.IndexOf('>', open) - open + 1);
            }
            TestReport.Check(Row("exact-top").Contains("<tr style=\"height:30pt\">") && Row("exact-top").Contains("max-height:30pt;overflow:hidden"), $"{name}: an exact 30 pt row is 30 pt and clips its cells");
            TestReport.Check(Row("least-top").Contains("<tr style=\"height:50pt\">") && !Row("least-top").Contains("overflow:hidden"), $"{name}: an at-least 50 pt row is at least 50 pt, not clipped");
            TestReport.Check(Cell(Row("no-top-line"), "no-top-line").Contains("border-top:none") && Cell(Row("least-top"), "least-top").Contains("border-top:0.5pt solid #000"),
                             $"{name}: a cell's explicit no border hides the table's line there (Word 16 reports LineStyle 0)");
            TestReport.Check(new[] { "top", "bottom", "left", "right" }.All(side => Cell(Row("custom-colour"), "custom-colour").Contains("border-" + side + ":1.5pt solid #123456")),
                             $"{name}: a border colour outside the 16-colour palette (#123456)");
            TestReport.Check(Cell(Row("b11"), "b11").Contains("border-top:0.5pt solid #000") && Cell(Row("a21"), "a21").Contains("border-bottom:0.5pt solid #000")
                             && Cell(Row("b21"), "b21").Contains("border-bottom:2.25pt solid #FF0000"),
                             $"{name}: two tables joined by Word are one table: inside lines where they meet, the second's outer line at its end");
        }
        if (LibreOfficeConversion.Convert(doc, ".odt", "odt") is { } odt)
        {
            List<string> blocks = Render(odt, "tables-extra-odt");
            string table = blocks.FirstOrDefault(block => block.Contains("exact-top")) ?? "";
            TestReport.Check(table.Contains("<tr style=\"height:30pt\">") && table.Contains("max-height:30pt;overflow:hidden") && table.Contains("<tr style=\"height:50pt\">"),
                             ".odt: style:row-height (exact, clipped) and style:min-row-height");
        }
        else TestReport.Check(false, ".odt: LibreOffice converted tables-extra-v1.doc");

        byte[] bytes = File.ReadAllBytes(doc);
        byte[] colours = [0x00, 0x00, 0x00, 0xFF, 0x00, 0x00, 0x00, 0xFF, 0x12, 0x34, 0x56, 0x00];
        int Find(byte opcode)
        {
            for (int i = 0; i + 16 < bytes.Length; i++)
                if (bytes[i] == opcode && bytes[i + 1] == 0xD6 && bytes[i + 2] == 12 && bytes.AsSpan(i + 3, 12).SequenceEqual(colours)) return i;
            return -1;
        }
        int top = Find(0x1A), left = Find(0x1B);
        TestReport.Check(top >= 0 && left >= 0, ".doc: row 1's sprmTBrcTopCv / LeftCv found");
        if (top < 0 || left < 0) return;
        // sprmTSetBrc (0xD62F): cell 0, top, Brc = green (00 80 00), 2.25 pt single; sprmTSetBrc80 (0xD620): cell 0, left, Brc80 = 2.25 pt single, ico 6 (red)
        byte[] setBrc = [0x2F, 0xD6, 12, 0, 1, 0x01, 0x00, 0x80, 0x00, 0x00, 18, 1, 0, 0, 0];
        byte[] setBrc80 = [0x20, 0xD6, 12, 0, 1, 0x02, 18, 1, 6, 0, 0, 0, 0, 0, 0];
        setBrc.CopyTo(bytes, top);
        setBrc80.CopyTo(bytes, left);
        string patched = Path.Combine(TestReport.OutputDirectory, "tables-extra-setbrc-v1.doc");
        File.WriteAllBytes(patched, bytes);
        string cell = Render(patched, "tables-extra-setbrc").FirstOrDefault(block => block.Contains(">exact-top<")) ?? "";
        cell = cell.Substring(0, Math.Max(0, cell.IndexOf(">exact-top<", StringComparison.Ordinal)));
        TestReport.Check(cell.Contains("border-top:2.25pt solid #008000") && cell.Contains("border-left:2.25pt solid #FF0000"),
                         ".doc: sprmTSetBrc (Brc, green top) and sprmTSetBrc80 (Brc80, red left) put borders on the cells they name");
    }

    /// <summary>While the index builds, the last row known is provisional: whether the next row continues its cells is not known yet.</summary>
    private static void RowAwaitingNextIsProvisional()
    {
        var rtf = new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}" + "\r\n");
        rtf.Append(@"\trowd\trbrdrb\brdrs\brdrw10\clvmgf\cellx2000\cellx4000\pard\intbl top\cell b\cell\row" + "\r\n");
        for (int i = 0; i < 20_000; i++) rtf.Append(@"\trowd\trbrdrb\brdrs\brdrw10\clvmrg\cellx2000\cellx4000\pard\intbl \cell c").Append(i).Append(@"\cell\row").Append("\r\n");
        rtf.Append('}');
        string path = Path.Combine(TestReport.OutputDirectory, "tables-provisional-v3.rtf");
        File.WriteAllText(path, rtf.ToString(), Encoding.ASCII);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-tables-provisional"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        int last = document.KnownBlockCount - 1;
        WordDocumentBlockHtml early = document.RenderBlocks(last, 1)[0];
        TestReport.Check(!document.IsIndexComplete && early.IsProvisional && !early.Html.Contains("border-bottom:none") && !document.RenderBlocks(0, 1)[0].IsProvisional,
                         $"RTF: the last row known while the index builds is provisional, rows before it are not (block {last}, provisional={early.IsProvisional})");
        while (!document.IsIndexComplete) document.ScanAhead(5_000);
        WordDocumentBlockHtml final = document.RenderBlocks(last, 1)[0];
        TestReport.Check(!final.IsProvisional && final.Html.Contains("border-bottom:none"),
                         "RTF: once the index is complete the row is final and its cell merged down has no bottom line");
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

        // two nested tables with no paragraph between: Word 16 opens them as ONE table of 2 rows (Cell.Tables.Count = 1, checked by COM)
        // and a text box inside a nested row definition (with its own table) leaves the definition alone
        const string twoTables = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}
\pard\intbl\itap1
\pard\intbl\itap2 a1\nestcell a2\nestcell{\*\nesttableprops\trowd\cellx1000\cellx2000{\shp{\*\shpinst{\shptxt \trowd\cellx500\cellx900\cellx1300\pard\intbl x\cell y\cell z\cell\row}}}\nestrow}
\pard\intbl\itap2 b1\nestcell b2\nestcell b3\nestcell{\*\nesttableprops\trowd\cellx700\cellx1400\cellx2100\nestrow}
\pard\intbl\itap1 \cell\trowd\cellx3000\row
\pard after\par}";
        string twoPath = Path.Combine(TestReport.OutputDirectory, "tables-two-nested-v1.rtf");
        File.WriteAllText(twoPath, twoTables, Encoding.ASCII);
        string two = Render(twoPath, "tables-two-nested").FirstOrDefault(block => block.Contains(">a1<")) ?? "";
        TestReport.Check(Regex.Matches(two, "<table class=\"nested-table\"").Count == 1 && two.Contains(">b3<"), "RTF: two nested tables with no paragraph between are one table, as Word 16 opens them");
        TestReport.Check(Regex.IsMatch(two, "<td style=\"width:50pt[^>]*>(<[^>]+>)*a1<"), "RTF: a text box's table inside a nested row definition does not replace it (a1 keeps its 1000-twip width)");
    }

    private static List<string> Render(string path, string cache)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-" + cache) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }
}
