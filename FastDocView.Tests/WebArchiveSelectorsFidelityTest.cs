using System.Text;
using System.Text.RegularExpressions;
using FastDocView;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// A web archive's sheet beyond Word's own rules (test-data/word-mhtml/css-selectors-v1.mht, hand-written): attribute selectors,
/// sibling combinators, structural pseudo-classes, pseudo-elements and @media print. Word's print ignores all of them
/// (css-selectors-v1.pdf: only the plain class rules colour the controls); Edge's print applies them (css-selectors-v1-edge.pdf,
/// msedge --headless --print-to-pdf) - a browser's archive is printed as the browser prints it, so the expectations are Edge's.
/// Then the same body repeated many times is read windowed (sidecar spans rescanned, shuffled windows) and must equal the
/// in-memory read: the selector state the index pass keeps per block reproduces the continuous resolution.
/// </summary>
public static class WebArchiveSelectorsFidelityTest
{
    private static string Fixture => Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-mhtml", "css-selectors-v1.mht");

    public static void Run()
    {
        TestReport.Step("Web archive selectors: attribute, sibling, structural, pseudo-elements, @media print (Edge's print)");
        if (!File.Exists(Fixture)) { TestReport.Check(false, "test-data/word-mhtml/css-selectors-v1.mht present"); return; }
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(Fixture, policy))
        {
            string all = string.Join("\n", document.RenderBlocks(0, document.BlockCount).Select(b => b.Html));
            string Tag(string text)
            {
                int at = all.IndexOf(">" + text, StringComparison.Ordinal);
                if (at < 0) return "";
                int start = all.LastIndexOf('<', at);
                return all.Substring(start, at + 1 - start);
            }
            void Expect(string text, string what, params string[] wanted)
                => TestReport.Check(Tag(text) is { Length: > 0 } tag && wanted.All(w => w.StartsWith("!", StringComparison.Ordinal) ? !tag.Contains(w.Substring(1)) : tag.Contains(w)), $"css-selectors: {what} [{Tag(text)}]");
            Expect("Case attr-present", "[title] underlines", "text-decoration:underline;", "!color:");
            Expect("Case attr-exact", "[title=\"exact\"] red, [title] still underlines", "text-decoration:underline;", "color:#C00000;");
            Expect("Case attr-dash", "[lang|=en] bold", "font-weight:bold;");
            Expect("Case attr-word", "[class~=word] italic", "font-style:italic;");
            Expect("Case attr-prefix", "[data-x^=pre] blue", "color:#0070C0;", "!#00B050");
            Expect("Case attr-suffix", "[data-x$=suf] green", "color:#00B050;", "!#0070C0");
            Expect("Case attr-substring", "[data-x*=mid] spaced", "letter-spacing:3pt;");
            Expect("Case adjacent", "h2 + p across blocks purple", "color:#7030A0;");
            Expect("Case general-sibling", "h2 ~ p.later across blocks small caps, not adjacent", "font-variant:small-caps;", "!#7030A0");
            Expect("Item one", "li:nth-child(odd) underlines the first item", "text-decoration:underline;", "!font-weight", "!color");
            Expect("Item two", "li:nth-child(2) bold", "font-weight:bold;", "!underline");
            Expect("Item three", "li:last-child red, odd underlined", "text-decoration:underline;", "color:#C00000;");
            Expect("Only item", "li:only-child green, odd underlined", "text-decoration:underline;", "color:#00B050;");
            Expect("Cell A", "td:first-of-type yellow", "background:#FFFF00;", "!font-weight");
            Expect("Cell B", "a middle cell neither", "!background", "!font-weight");
            Expect("Cell C", "td:last-of-type bold", "font-weight:bold;", "!background");
            Expect("Case media", "@media print applies, @media screen does not", "color:#00B050;", "!#C00000");
            Expect("Case screen-only", "@media screen display:none does not print", "!display:none");
            Expect("Case print-hide", "@media print display:none hides", "display:none;");
            Expect("Control one", "a plain class rule before them", "color:#C00000;");
            Expect("Control two", "a plain class rule after them", "color:#0070C0;");
            Match before = Regex.Match(document.Stylesheet, @"\[data-wa-pe~=""(\d+)""\]::before \{ content:""\[B\] "";color:#C00000; \}");
            Match after = Regex.Match(document.Stylesheet, @"\[data-wa-pe~=""(\d+)""\]::after \{ content:"" \[A\]"";color:#0070C0; \}");
            Match letter = Regex.Match(document.Stylesheet, @"\[data-wa-pe~=""(\d+)""\]::first-letter \{ font-size:200%;color:#C00000; \}");
            Match line = Regex.Match(document.Stylesheet, @"\[data-wa-pe~=""(\d+)""\]::first-line \{ text-transform:uppercase; \}");
            TestReport.Check(before.Success && after.Success && letter.Success && line.Success, "css-selectors: the pseudo-element rules are in the page stylesheet with their content " + string.Join(" | ", document.Stylesheet.Split((char)10).Where(l => l.Contains("data-wa-pe"))));
            if (before.Success && after.Success && letter.Success && line.Success)
            {
                Expect("Case before", "::before on its element", $"data-wa-pe=\"{before.Groups[1].Value}\"");
                Expect("Case after", "::after on its element", $"data-wa-pe=\"{after.Groups[1].Value}\"");
                Expect("Case first-letter", "::first-letter on its element", $"data-wa-pe=\"{letter.Groups[1].Value}\"");
                Expect("Case first-line", "::first-line on its element", $"data-wa-pe=\"{line.Groups[1].Value}\"");
            }
            TestReport.Check(!all.Contains("data-fdv-s") && !all.Contains("fdv-root"), "css-selectors: the index pass's state never reaches the page");
            // search and text follow what is printed: generated content found, a hidden paragraph not
            TestReport.Check(Hits(document, "[B] Case before") == 1 && Hits(document, "suffixed. [A]") == 1, "css-selectors: search finds ::before / ::after text with the element's own");
            TestReport.Check(Hits(document, "hidden in print") == 0 && Hits(document, "hidden on screen only") == 1, "css-selectors: search skips what @media print hides, finds what only the screen hides");
            string documentText = WordDocumentText.GetText(document);
            TestReport.Check(documentText.Contains("[B] Case before") && !documentText.Contains("hidden in print"), "css-selectors: GetText shows what the page shows");
        }

        string css = File.ReadAllText(Fixture);
        int open = css.IndexOf("<style>", StringComparison.Ordinal) + 7, close = css.IndexOf("</style>", StringComparison.Ordinal);
        string normalized = WebArchiveStylesheet.Normalize(css.Substring(open, close - open));
        TestReport.Check(WebArchiveStylesheet.Normalize(normalized) == normalized && normalized.Contains("li:last-child") && !normalized.Contains("@media"),
                         "css-selectors: the stored sheet reads back the same (matching @media flattened)");
        TestReport.Check(MatchesPrint("print") && MatchesPrint("all") && !MatchesPrint("screen") && MatchesPrint("not screen") && !MatchesPrint("not print")
                         && MatchesPrint("screen, print") && MatchesPrint("print and (orientation: portrait)") && !MatchesPrint("print and (orientation: landscape)")
                         && MatchesPrint("(min-width: 600px)") && !MatchesPrint("(min-width: 1000px)") && MatchesPrint("(max-width: 8in)") && !MatchesPrint("print and (hover: hover)")
                         && MatchesPrint("(color)") && !MatchesPrint("(monochrome)"),
                         "css-selectors: @media queries evaluated for Edge's print (Letter, 1 cm margins)");

        Windowed();
    }

    private static bool MatchesPrint(string query) => WebArchiveStylesheet.Normalize("@media " + query + " { p { color:red; } }").Contains("color:red");

    /// <summary>The fixture's body repeated: every window read from rescanned spans equals the same window of the in-memory read.</summary>
    private static void Windowed()
    {
        const int Copies = 1500;
        string source = File.ReadAllText(Fixture);
        int bodyStart = source.IndexOf("<div class=WordSection1>", StringComparison.Ordinal), bodyEnd = source.IndexOf("</div>\r\n</body>", StringComparison.Ordinal) + 6;
        var big = new StringBuilder(source.Length * Copies / 2);
        big.Append(source, 0, bodyStart);
        for (int copy = 0; copy < Copies; copy++) big.Append(source, bodyStart, bodyEnd - bodyStart).Append("\r\n");
        big.Append(source, bodyEnd, source.Length - bodyEnd);
        string path = Path.Combine(TestReport.OutputDirectory, "css-selectors-windowed-v1-" + Guid.NewGuid().ToString("N") + ".mht");
        string cache = Path.Combine(TestReport.OutputDirectory, "css-selectors-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            File.WriteAllText(path, big.ToString(), new UTF8Encoding(false));
            List<string> expected;
            int expectedHits;
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = cache }))
            {
                expected = document.RenderBlocks(0, document.BlockCount).Select(b => b.Html).ToList();
                expectedHits = Hits(document, "[B] Case before");
            }
            TestReport.Check(expectedHits == Copies, $"css-selectors windowed: every copy's generated text is found ({expectedHits})");
            TestReport.Check(expected.Count(h => h.Contains("color:#7030A0;")) == Copies && expected.Count(h => h.Contains("Item three") && h.Contains("color:#C00000;")) == Copies,
                             "css-selectors windowed: the sibling and last-child rules hold in every copy");
            foreach ((string name, WordIndexPolicy policy) in new[]
                     {
                         ("on demand", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache, IndexesOnDemand = true }),
                         ("sidecar", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache }),
                     })
            {
                if (name == "on demand" && Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                if (!document.IsIndexComplete) document.ScanAhead(int.MaxValue / 2);
                var random = new Random(17);
                int mismatches = 0, checkedBlocks = 0;
                string firstMismatch = "";
                for (int round = 0; round < 60; round++)
                {
                    int first = random.Next(expected.Count);
                    IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(first, 25);
                    for (int i = 0; i < window.Count; i++, checkedBlocks++)
                        if (window[i].Html != expected[first + i]) { mismatches++; if (firstMismatch.Length == 0) firstMismatch = $"block {first + i}: {window[i].Html} | expected {expected[first + i]}"; }
                }
                TestReport.Check(mismatches == 0 && checkedBlocks > 0, $"css-selectors windowed ({name}): {checkedBlocks} blocks in shuffled windows equal the continuous read ({mismatches} differ) {firstMismatch}");
            }
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
        }
    }

    private static int Hits(IWordWindowedDocument document, string query)
    {
        int hits = 0, from = 0;
        while (true)
        {
            WordDocumentSearchBatch batch = document.Search(query, WordDocumentSearchOptions.Default, from, 500);
            hits += batch.HitCounts.Sum();
            if (batch.ReachedEnd) return hits;
            from = batch.NextBlock;
        }
    }
}
