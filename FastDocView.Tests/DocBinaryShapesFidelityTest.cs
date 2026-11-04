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
            TestReport.Check(html.Contains("<svg") && html.Contains("<path d=") && html.Contains("text-box"), "LibreOffice .doc: shapes drawn, text boxes placed");
            TestReport.Check(html.Contains("word-officeart-gradient-"), "LibreOffice .doc: the gradient rectangle is a gradient");
            int placeholder = html.IndexOf("[image]", StringComparison.Ordinal);
            TestReport.Check(placeholder < 0, "LibreOffice .doc: no shape left as a placeholder" + (placeholder < 0 ? "" : " (" + html.Substring(Math.Max(0, placeholder - 300), Math.Min(400, html.Length - Math.Max(0, placeholder - 300))) + ")"));
        }

        TestReport.Step(".doc shapes: the hand-written RTF shapes saved as .doc by LibreOffice (WordArt, own geometry, fills, system colours)");
        string rtf = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-drawn-shapes.rtf");
        if (!File.Exists(rtf)) RtfShapesFidelityTest.Run();
        if (File.Exists(rtf) && LibreOfficeConversion.Convert(rtf, ".doc", "doc:MS Word 97") is { } fromRtf)
        {
            string html = string.Concat(Render(fromRtf));
            int shown = Occurrences(Tags.Replace(html, ""), "WordArtText"), found = SearchCount(fromRtf, "WordArtText");
            TestReport.Check(shown == 1 && found == 1, $"LibreOffice .doc: the WordArt text shown {shown}, found {found}, expected 1 / 1");
            TestReport.Check(html.Contains("word-officeart-gradient-"), $"LibreOffice .doc: shades drawn as gradients (gradients {Occurrences(html, "word-officeart-gradient-")}, rectangular {Occurrences(html, "word-officeart-shade-")}, patterns {Occurrences(html, "word-officeart-pattern-")}, picture fills {Occurrences(html, "word-shape-fill-picture")})");
            foreach (string word in new[] { "boxtext one", "arrowtext", "groupone", "grouptwo" })
            {
                int wordShown = Occurrences(Tags.Replace(html, ""), word), wordFound = SearchCount(fromRtf, word);
                TestReport.Check(wordShown == 1 && wordFound == 1, $"LibreOffice .doc from RTF: '{word}' shown {wordShown}, found {wordFound}, expected 1 / 1");
            }
        }

        TestReport.Step(".doc shapes: Fontwork saved as .doc by LibreOffice (WordArt: gtextUNICODE)");
        string fontwork = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-fontwork.fodt");
        const string FontworkDocument = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
            + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" "
            + "xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:body><office:text><text:p><draw:custom-shape text:anchor-type=\"paragraph\" svg:width=\"2in\" svg:height=\"0.5in\" svg:x=\"0in\" svg:y=\"0in\">"
            + "<text:p>FontworkText</text:p><draw:enhanced-geometry svg:viewBox=\"0 0 21600 21600\" draw:type=\"fontwork-plain-text\" draw:text-path=\"true\" "
            + "draw:text-path-mode=\"shape\" draw:text-path-scale=\"shape\"/></draw:custom-shape>after</text:p></office:text></office:body></office:document>";
        if (!File.Exists(fontwork) || File.ReadAllText(fontwork) != FontworkDocument) File.WriteAllText(fontwork, FontworkDocument);
        if (LibreOfficeConversion.Convert(fontwork, ".doc", "doc:MS Word 97") is { } wordArtDoc)
        {
            string html = string.Concat(Render(wordArtDoc));
            int shown = Occurrences(Tags.Replace(html, ""), "FontworkText"), found = SearchCount(wordArtDoc, "FontworkText");
            TestReport.Check(shown == 1 && found == 1, $"LibreOffice .doc Fontwork: shown {shown}, found {found}, expected 1 / 1");
            TestReport.Check(html.Contains("class=\"text-box-paragraph\" style=\"display:block;text-align:center;white-space:pre-wrap;") && !html.Contains("<path"), "LibreOffice .doc Fontwork: the WordArt string drawn centred in its font, no outline (a text path) (" + html + ")");
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
