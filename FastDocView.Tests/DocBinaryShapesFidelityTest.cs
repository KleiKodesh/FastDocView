using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// .doc text boxes and drawn shapes (OfficeArt via FSPA, txbx story via lTxid): the .odt shapes fixture saved as .doc by
/// LibreOffice (headless, into TestOutput, never committed), and Apache POI / LibreOffice test files with text boxes - each
/// text-box text rendered in a box once and found by the search exactly once.
/// </summary>
public static class DocBinaryShapesFidelityTest
{
    private static readonly Regex Tags = new("<[^>]+>", RegexOptions.CultureInvariant);

    public static void Run()
    {
        TestReport.Step(".doc shapes: the .odt shapes fixture saved as .doc by LibreOffice");
        string fodt = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-shapes.fodt");
        if (!File.Exists(fodt)) OdtShapesFidelityTest.Run();
        if (File.Exists(fodt) && LibreOfficeConversion.Convert(fodt, ".doc", "doc:MS Word 97") is { } doc)
        {
            string html = string.Concat(Render(doc));
            foreach (string word in new[] { "boxtext one", "inlinebox", "centrebox", "rtlbox", "shapetext", "gradtext", "groupone" })
            {
                int shown = Occurrences(Tags.Replace(html, ""), word), found = SearchCount(doc, word);
                TestReport.Check(shown == 1 && found == 1, $"LibreOffice .doc: '{word}' shown {shown}, found {found}, expected 1 / 1");
            }
            TestReport.Check(html.Contains("<svg") && html.Contains("<ellipse") && html.Contains("text-box"), "LibreOffice .doc: shapes drawn, text boxes placed");
            int placeholder = html.IndexOf("[image]", StringComparison.Ordinal);
            TestReport.Check(placeholder < 0, "LibreOffice .doc: no shape left as a placeholder" + (placeholder < 0 ? "" : " (" + html.Substring(Math.Max(0, placeholder - 300), Math.Min(400, html.Length - Math.Max(0, placeholder - 300))) + ")"));
        }

        TestReport.Step(".doc shapes: Apache POI / LibreOffice files with text boxes");
        string data = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data");
        (string file, string word)[] cases =
        [
            (Path.Combine("apache-poi", "Bug41898.doc"), ""),
            (Path.Combine("libreoffice", "sw", "qa", "extras", "ww8export", "data", "tdf112618_textbox_no_bg.doc"), ""),
            (Path.Combine("libreoffice", "sw", "qa", "extras", "ww8export", "data", "tdf76349_textboxMargins.doc"), ""),
            (Path.Combine("libreoffice", "sw", "qa", "extras", "ww8export", "data", "shapes-line-ellipse.doc"), ""),
        ];
        foreach ((string file, _) in cases)
        {
            string path = Path.Combine(data, file);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/{file} present"); continue; }
            List<string> html = Render(path);
            string all = string.Concat(html);
            int boxes = Occurrences(all, "class=\"text-box") + Occurrences(all, "class=\"word-shape\"");
            string boxText = string.Concat(Regex.Matches(all, "<span class=\"text-box-paragraph\"[^>]*>(.*?)</span>", RegexOptions.Singleline).Select(m => Tags.Replace(m.Groups[1].Value, "")));
            string word = boxText.Split([' ', '\t', ' '], StringSplitOptions.RemoveEmptyEntries).Where(w => w.Length >= 4 && w.All(char.IsLetter)).FirstOrDefault() ?? "";
            Console.WriteLine($"      {Path.GetFileName(file)}: {boxes} drawing(s), text-box text '{(boxText.Length > 60 ? boxText[..60] : boxText)}'");
            TestReport.Check(boxes > 0, $"{Path.GetFileName(file)}: drawings rendered ({boxes})");
            if (word.Length > 0)
            {
                int shown = Occurrences(Tags.Replace(all, " "), word), found = SearchCount(path, word);
                TestReport.Check(found == shown, $"{Path.GetFileName(file)}: '{word}' found {found} times, shown {shown}");
            }
        }
    }

    private static List<string> Render(string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        var html = new List<string>();
        for (int first = 0; first < document.BlockCount; first += 100)
            html.AddRange(document.RenderBlocks(first, 100).Select(block => block.Html));
        return html;
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
}
