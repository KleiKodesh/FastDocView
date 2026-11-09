using System.Globalization;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// List number formats against Word 16 (--fidelity, seconds): test-data/word-lists/list-formats-v1-a/-b .docx / .doc / .rtf,
/// one single-level "%1." list per MSONFC code (= WdListNumberStyle) and start value, built by make-list-formats-fixture.ps1,
/// which also saved Word's own ListFormat.ListString for every paragraph (list-formats-v1-*.expected.txt). The label each
/// reader draws must equal Word's, paragraph by paragraph. RTF draws Word's \listtext, so it checks the parser, the .docx and
/// .doc sides check the counted labels. Word starts a hebrew2 ListString with an invisible right-to-left mark; it is not compared.
/// </summary>
public static class ListNumberFormatsFidelityTest
{
    private static readonly Regex Marker = new("<span class=\"list-marker\"[^>]*>(.*?)</span>(?!</span>)", RegexOptions.Compiled);

    /// <summary>MSONFC codes whose Word labels are not reproduced yet (they show decimal): counting words in CJK / Korean / Thai /
    /// Vietnamese / Hindi, kana orders (aiueo, iroha), Korean letters (ganada, chosung), koreanDigital and Word's custom formats (60+).</summary>
    private static readonly HashSet<string> NotImplemented =
        ["11", "12", "13", "16", "20", "21", "24", "25", "33", "34", "35", "37", "38", "39", "41", "42", "43", "52", "55", "56",
         "60", "61", "62", "63", "64", "65", "66", "67", "68"];

    public static void Run()
    {
        TestReport.Step("List number formats: every Word number style against Word 16's own labels (.docx, .doc, .rtf)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists");
        foreach (string part in new[] { "a", "b" })
        {
            string expectedPath = Path.Combine(folder, $"list-formats-v1-{part}.expected.txt");
            if (!File.Exists(expectedPath)) { TestReport.Check(false, "missing fixture " + expectedPath); continue; }
            var expected = new List<(string code, string value, string label)>();
            foreach (string line in File.ReadAllLines(expectedPath))
            {
                string[] fields = line.Split('\t');
                if (fields.Length == 3) expected.Add((fields[0], fields[1], Unescape(fields[2])));
            }
            foreach (string extension in new[] { "docx", "doc", "rtf" })
            {
                string path = Path.Combine(folder, $"list-formats-v1-{part}.{extension}");
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
                while (!document.IsIndexComplete) document.ScanAhead(10000);
                var labels = new List<string>();
                foreach (WordDocumentBlockHtml block in document.RenderBlocks(1, document.BlockCount - 1))
                    labels.Add(LabelOf(block.Html));
                int wrong = 0, notImplemented = 0;
                var mismatches = new StringBuilder();
                for (int i = 0; i < expected.Count; i++)
                {
                    string actual = i < labels.Count ? labels[i] : "(missing)";
                    if (actual == expected[i].label) continue;
                    if (NotImplemented.Contains(expected[i].code)) { notImplemented++; continue; }
                    wrong++;
                    if (mismatches.Length < 3000)
                        mismatches.Append(" [code ").Append(expected[i].code).Append(" v").Append(expected[i].value).Append(": word ")
                                  .Append(Escape(expected[i].label)).Append(" got ").Append(Escape(actual)).Append(']');
                }
                TestReport.Check(wrong == 0, $"{extension} part {part}: {expected.Count - wrong - notImplemented}/{expected.Count - notImplemented} labels equal Word's"
                                             + $" ({notImplemented} in formats drawn as decimal: CJK / Korean / Thai / Vietnamese / Hindi counting words, kana orders, Word's custom formats)" + mismatches);
            }
        }
    }

    /// <summary>The text of a block's list label (tags stripped, the tab suffix and trailing spaces dropped).</summary>
    private static string LabelOf(string html)
    {
        int start = html.IndexOf("<span class=\"list-marker\"", StringComparison.Ordinal);
        if (start < 0) return "";
        int depth = 0, i = start, end = -1;
        while (i < html.Length)
        {
            if (string.CompareOrdinal(html, i, "<span", 0, 5) == 0) depth++;
            else if (string.CompareOrdinal(html, i, "</span>", 0, 7) == 0 && --depth == 0) { end = i; break; }
            i++;
        }
        if (end < 0) return "";
        string inner = html.Substring(start, end - start);
        inner = Regex.Replace(inner, "<[^>]*>", "");
        return WebUtility.HtmlDecode(inner).Replace(((char)0x200F).ToString(), "").TrimEnd('\t', ' ', ' ');
    }

    private static string Unescape(string text)
        => Regex.Replace(text.Replace(@"\u200F", ""), @"\\u([0-9A-F]{4})", m => ((char)int.Parse(m.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString())
                .TrimEnd('\t', ' ');

    private static string Escape(string text)
    {
        var escaped = new StringBuilder();
        foreach (char c in text) escaped.Append(c is >= ' ' and < (char)127 ? c.ToString() : "\\u" + ((int)c).ToString("X4", CultureInfo.InvariantCulture));
        return escaped.ToString();
    }
}
