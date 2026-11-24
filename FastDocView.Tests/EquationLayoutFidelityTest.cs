using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// EQ field layout as Word 16 draws it: test-data/word-fields/eq-layout-v1.* (make-eq-layout-fixture.ps1; one "case NAME: {EQ} end"
/// paragraph per construct). Word's PDF (eq-layout-v1.pdf) shows: \o centres its elements on the widest one by default, \al / \ar
/// align them left / right; \i puts its limits above and below the symbol (integral, sum, \fc, \vc) and to its right only with \in;
/// \b with only \lc draws only the left bracket, with only \rc only the right one, with no option a pair of parentheses.
/// </summary>
public static class EquationLayoutFidelityTest
{
    public static void Run()
    {
        TestReport.Step("EQ field layout as Word 16 draws it: overstrike alignment, integral limits, one-sided brackets (.docx / .doc / .rtf)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-fields");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "eq-layout-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-fields/eq-layout-v1{extension} present"); continue; }
            Check(path, extension);
        }
    }

    private static void Check(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-eq-layout-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        var cases = new Dictionary<string, string>();
        foreach (var block in document.RenderBlocks(0, document.BlockCount))
        {
            Match name = Regex.Match(block.Html, @"case ([a-z0-9]+):");
            if (name.Success) cases[name.Groups[1].Value] = block.Html;
        }
        string Case(string name) => cases.TryGetValue(name, out string? html) ? html : "";
        void Expect(string name, bool condition, string what) =>
            TestReport.Check(condition, $"{format} EQ {name}: {what} (got [{BodyScriptFidelityTest.Ascii(Math(Case(name)))}])");

        Expect("o", Case("o").Contains("eq-overstrike\" style=\"justify-items:center\""), "\\o centres its elements");
        Expect("oal", Case("oal").Contains("justify-items:start"), "\\o\\al aligns them left");
        Expect("oac", Case("oac").Contains("justify-items:center"), "\\o\\ac centres them");
        Expect("oar", Case("oar").Contains("justify-items:end"), "\\o\\ar aligns them right");
        Expect("oi", Case("oi").Contains("justify-items:center") && !Case("oi").Contains("mpadded width=\"0\""), "every element takes part in the width (the widest decides)");
        foreach (string name in new[] { "int", "sum", "fc", "vc" })
            Expect(name, Case(name).Contains("<munderover>") && !Case(name).Contains("<msubsup>"), "limits above and below the symbol");
        foreach (string name in new[] { "intin", "sumin" })
            Expect(name, Case(name).Contains("<msubsup>") && !Case(name).Contains("<munderover>"), "\\in puts the limits to the right");
        Expect("blc", Brackets(Case("blc")) == "[", "\\lc\\[ alone: only the left bracket");
        Expect("brc", Brackets(Case("brc")) == "]", "\\rc\\] alone: only the right bracket");
        Expect("blcbrace", Brackets(Case("blcbrace")) == "{", "\\lc\\{ alone: only the left brace");
        Expect("blcbar", Brackets(Case("blcbar")) == "|", "\\lc\\| alone: only the bar");
        Expect("bfrac", Brackets(Case("bfrac")) == "()", "no option: a pair of parentheses");
    }

    /// <summary>The stretchy brackets of a case, in order.</summary>
    private static string Brackets(string html) => string.Concat(Regex.Matches(html, "<mo stretchy=\"true\">([^<]*)</mo>").Select(m => m.Groups[1].Value));

    private static string Math(string html)
    {
        int start = html.IndexOf("<math", StringComparison.Ordinal), end = html.IndexOf("</math>", StringComparison.Ordinal);
        return start >= 0 && end > start ? html.Substring(start, end - start) : html;
    }
}
