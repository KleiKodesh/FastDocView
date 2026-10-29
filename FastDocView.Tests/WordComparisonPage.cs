using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Our side of the Word fidelity comparison (tools in NOTES.md, "Word comparison"): reads the paragraph texts Word
/// reported for the sampled pages (word.json from wordref.ps1), finds the blocks that hold them by reading the
/// document window by window in order, and writes a static page — the document stylesheet plus those blocks — at
/// the column width of Word's text area, for the browser side to measure. Writes a match map beside it.
/// </summary>
public static partial class WordComparisonPage
{
    private const int KeyLength = 20, MinimumKeyLength = 12, WindowBlocks = 64, SearchAheadBlocks = 4000, ContextBlocks = 1, MaximumPageSpanBlocks = 150;

    public static void Run(string documentPath, string wordJsonPath, string outputHtmlPath)
    {
        TestReport.Step($"Word comparison page: {Path.GetFileName(documentPath)}");
        using JsonDocument word = JsonDocument.Parse(File.ReadAllText(wordJsonPath));
        var targets = new List<(int page, int paragraph, string key, string? alternate)>();
        double textWidthPoints = word.RootElement.TryGetProperty("textWidth", out JsonElement width) ? width.GetDouble() : 0;
        foreach (JsonElement page in Elements(word.RootElement.GetProperty("pages")))
        {
            if (textWidthPoints == 0)
                textWidthPoints = page.GetProperty("pageWidth").GetDouble() - page.GetProperty("leftMargin").GetDouble() - page.GetProperty("rightMargin").GetDouble();
            int paragraphIndex = 0;
            foreach (JsonElement paragraph in Elements(page.GetProperty("paragraphs")))
            {
                string key = Middle(Normalize(paragraph.GetProperty("text").GetString() ?? ""));
                string? alternate = paragraph.TryGetProperty("alt", out JsonElement alt) ? Middle(Normalize(alt.GetString() ?? "")) : null;
                if (key.Length >= MinimumKeyLength) targets.Add((page.GetProperty("page").GetInt32(), paragraphIndex, key, alternate));
                paragraphIndex++;
            }
        }

        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(documentPath);
        while (!document.IsIndexComplete) document.ScanAhead(10_000);
        var blockTexts = new Dictionary<int, string>();
        var blockHtml = new SortedDictionary<int, string>();
        string TextOf(int blockIndex)
        {
            if (blockTexts.TryGetValue(blockIndex, out string? text)) return text;
            int first = blockIndex - blockIndex % WindowBlocks;
            if (blockTexts.Count > 4 * WindowBlocks) blockTexts.Clear();   // bounded: only the window being searched
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, WindowBlocks))
                blockTexts[block.BlockIndex] = Normalize(WebUtility.HtmlDecode(Tags().Replace(block.Html, "")));
            return blockTexts[blockIndex];
        }

        var matches = new List<object>();
        int cursor = 0, matched = 0;
        var pageBlocks = new Dictionary<int, List<int>>();
        foreach ((int page, int paragraph, string key, string? alternate) in targets)
        {
            int found = -1;
            string usedKey = key;
            for (int blockIndex = cursor; blockIndex < document.BlockCount && blockIndex < cursor + SearchAheadBlocks; blockIndex++)
            {
                string text = TextOf(blockIndex);
                if (text.Contains(key, StringComparison.Ordinal)) { found = blockIndex; break; }
                if (alternate is { Length: > 0 } && text.Contains(alternate, StringComparison.Ordinal)) { found = blockIndex; usedKey = alternate; break; }
            }
            if (found < 0) continue;
            matched++;
            cursor = found;
            matches.Add(new { page, paragraph, block = found, key = usedKey });
            (pageBlocks.TryGetValue(page, out var list) ? list : pageBlocks[page] = []).Add(found);
        }
        // a page's blocks: around the median match (a stray match elsewhere must not pull in thousands of blocks)
        var pageRanges = new Dictionary<int, (int first, int last)>();
        foreach ((int page, List<int> found) in pageBlocks)
        {
            found.Sort();
            int median = found[found.Count / 2];
            List<int> near = found.Where(block => Math.Abs(block - median) <= MaximumPageSpanBlocks).ToList();
            pageRanges[page] = (near[0], near[^1]);
        }
        foreach ((int first, int last) in pageRanges.Values)
            for (int blockIndex = Math.Max(0, first - ContextBlocks); blockIndex <= Math.Min(document.BlockCount - 1, last + ContextBlocks); blockIndex++)
                if (!blockHtml.ContainsKey(blockIndex)) blockHtml[blockIndex] = document.RenderBlocks(blockIndex, 1)[0].Html;

        double widthPixels = textWidthPoints * 96 / 72;
        var html = new StringBuilder();
        html.Append("<!doctype html><html><head><meta charset=\"utf-8\"><style>body{margin:0;background:#fff}")
            .Append(".word-document{margin:0 8px}")
            .Append(document.Stylesheet).Append("</style></head><body><div class=\"word-document\" dir=\"")
            .Append(document.IsRightToLeft ? "rtl" : "ltr").Append("\" style=\"width:").Append(widthPixels.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)).Append("px\">");
        foreach ((int blockIndex, string blockMarkup) in blockHtml)
            html.Append("<div class=\"block\" data-block=\"").Append(blockIndex).Append("\">").Append(blockMarkup).Append("</div>");
        html.Append("</div></body></html>");
        File.WriteAllText(outputHtmlPath, html.ToString(), new UTF8Encoding(false));
        File.WriteAllText(Path.ChangeExtension(outputHtmlPath, ".matches.json"), JsonSerializer.Serialize(matches));
        Console.WriteLine($"  {matched}/{targets.Count} Word paragraphs matched to blocks, {blockHtml.Count} blocks written");
    }

    /// <summary>Letters and digits only — the text both sides agree on (no list labels, spaces, marks or field codes).</summary>
    public static string Normalize(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (char c in text) if (char.IsLetterOrDigit(c)) builder.Append(c);
        return builder.ToString();
    }

    /// <summary>The middle of a line's letters: list labels and hyphenation sit at the edges.</summary>
    private static string Middle(string key) => key.Length <= KeyLength ? key : key.Substring((key.Length - KeyLength) / 2, KeyLength);

    private static IEnumerable<JsonElement> Elements(JsonElement element) =>
        element.ValueKind == JsonValueKind.Array ? element.EnumerateArray() : [element];

    [GeneratedRegex("<[^>]*>")]
    private static partial Regex Tags();
}
