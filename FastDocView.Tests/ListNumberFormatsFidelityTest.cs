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
/// .doc sides check the counted labels; list-formats-words-v1 checks the script formats over 1..129 and large values. Word starts a hebrew2 ListString with an invisible right-to-left mark; it is not compared.
/// </summary>
public static class ListNumberFormatsFidelityTest
{
    private static readonly Regex Marker = new("<span class=\"list-marker\"[^>]*>(.*?)</span>(?!</span>)", RegexOptions.Compiled);

    public static void Run()
    {
        TestReport.Step("List number formats: every Word number style against Word 16's own labels (.docx, .doc, .rtf)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists");
        RunWords(folder);
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
                int wrong = 0;
                var mismatches = new StringBuilder();
                for (int i = 0; i < expected.Count; i++)
                {
                    string actual = i < labels.Count ? labels[i] : "(missing)";
                    // Word's custom formats (60+) live only in .docx: Word 16 saves them in .doc / RTF as decimal and, reopening
                    // the .doc, shows 1., 2. itself (the RTF carries \levelnfc0 and \pntext 1.).
                    string expectedLabel = extension != "docx" && int.Parse(expected[i].code, CultureInfo.InvariantCulture) >= 60
                        ? expected[i].value + "." : expected[i].label;
                    if (actual == expectedLabel) continue;
                    wrong++;
                    if (mismatches.Length < 3000)
                        mismatches.Append(" [code ").Append(expected[i].code).Append(" v").Append(expected[i].value).Append(": word ")
                                  .Append(Escape(expectedLabel)).Append(" got ").Append(Escape(actual)).Append(']');
                }
                TestReport.Check(wrong == 0, $"{extension} part {part}: {expected.Count - wrong}/{expected.Count} labels equal Word's" + mismatches);
            }
        }
    }

    /// <summary>
    /// The counting / letter formats over a wider range: list-formats-words-v1.expected.txt (make-list-formats-words.ps1) holds
    /// Word 16's ListString for values 1..129 and large start values of each format WordListNumberScripts draws.
    /// </summary>
    private static void RunWords(string folder)
    {
        string path = Path.Combine(folder, "list-formats-words-v1.expected.txt");
        if (!File.Exists(path)) { TestReport.Check(false, "missing fixture " + path); return; }
        int total = 0, wrong = 0;
        var mismatches = new StringBuilder();
        foreach (string line in File.ReadAllLines(path))
        {
            string[] fields = line.Split('\t');
            if (fields.Length != 3) continue;
            byte code = byte.Parse(fields[0], CultureInfo.InvariantCulture);
            int value = int.Parse(fields[1], CultureInfo.InvariantCulture);
            string expected = Unescape(fields[2]);
            string format = code >= 60 ? FastDocView.Html.WordListNumberScripts.CustomFormatsByCode[code - 60] : FastDocView.DocBinary.DocBinaryListDefinitions.NumberFormatName(code);
            string actual = FastDocView.DocxHtml.DocxNumberingDefinitions.FormatNumber(value, format);
            total++;
            if (actual == expected) continue;
            wrong++;
            if (mismatches.Length < 3000)
                mismatches.Append(" [code ").Append(code).Append(" v").Append(value).Append(": word ").Append(Escape(expected)).Append(" got ").Append(Escape(actual)).Append(']');
        }
        TestReport.Check(total > 3000 && wrong == 0, $"counting words, kana, Korean letters, custom formats over 1..129 and large values: {total - wrong}/{total} equal Word's" + mismatches);

        // .odt: LibreOffice's num-format names for the same kana / Hangul / zero-padded orders (defaultnumberingprovider.cxx:1160
        // makeNumberingIdentifier: the first three labels; :1154-1155 fixed names) draw what Word draws
        (string odf, string word)[] odtNames =
        [
            ("001, 002, 003, ...", FastDocView.Html.WordListNumberScripts.CustomZero3), ("0001, 0002, 0003, ...", FastDocView.Html.WordListNumberScripts.CustomZero4),
            ("00001, 00002, 00003, ...", FastDocView.Html.WordListNumberScripts.CustomZero5), ("\uFF71, \uFF72, \uFF73, ...", "aiueo"),
            ("\u30A2, \u30A4, \u30A6, ...", "aiueoFullWidth"), ("\uFF72, \uFF9B, \uFF8A, ...", "iroha"), ("\u30A4, \u30ED, \u30CF, ...", "irohaFullWidth"),
            ("\uAC00, \uB098, \uB2E4, ...", "ganada"), ("\u3131, \u3134, \u3137, ...", "chosung"),
        ];
        int odtWrong = 0;
        foreach ((string odf, string word) in odtNames)
        {
            if (FastDocView.Odt.OdtListStyle.FormatNumber(1, odf) + ", " + FastDocView.Odt.OdtListStyle.FormatNumber(2, odf) + ", "
                + FastDocView.Odt.OdtListStyle.FormatNumber(3, odf) + ", ..." != odf) odtWrong++;
            for (int value = 1; value <= 129; value++)
                if (FastDocView.Odt.OdtListStyle.FormatNumber(value, odf) != FastDocView.DocxHtml.DocxNumberingDefinitions.FormatNumber(value, word)) odtWrong++;
        }
        TestReport.Check(odtWrong == 0, $"odt: LibreOffice's kana / Hangul / 001 num-format names draw Word's labels ({odtWrong} wrong)");
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
