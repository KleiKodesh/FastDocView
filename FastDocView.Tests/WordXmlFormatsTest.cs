using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Word's two XML formats, against Word's own save of the same document (test-data/word-xml/make-word-xml-fixtures.ps1:
/// word-xml-v1.docx, .xml = Word XML Document / Flat OPC, -2003.xml = Word 2003 XML Document, .pdf = Word's print).
/// Word's PDF text (pdftotext): heading, "Plain text, bold and italic with a footnote1 and an endnotei.", "1. First numbered",
/// "2. Second numbered", a bullet, "A link", the 2 x 2 table, the Hebrew line, the picture, "Last paragraph.", then the footnote
/// and (on the last page) the endnote. Checked: detection by content (not the extension), the Flat OPC file renders exactly as
/// the .docx (the same reader over the same parts), both in memory and windowed from a sidecar, and the 2003 file shows the same
/// text, formatting, list labels, notes, table, link and picture.
/// </summary>
public static class WordXmlFormatsTest
{
    private static string Output => Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestOutput")).FullName;

    public static void Run()
    {
        TestReport.Step("Word XML formats: Flat OPC and Word 2003 XML against Word's .docx and PDF of the same document");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-xml");
        string docx = Path.Combine(folder, "word-xml-v1.docx"), flat = Path.Combine(folder, "word-xml-v1.xml"), word2003 = Path.Combine(folder, "word-xml-v1-2003.xml");
        if (!File.Exists(docx) || !File.Exists(flat) || !File.Exists(word2003)) { TestReport.Check(false, "test-data/word-xml fixtures present"); return; }

        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(flat) == WordWindowedDocumentOpener.WordFileFormat.FlatXmlPackage, "Flat OPC detected by its pkg:package root");
        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(word2003) == WordWindowedDocumentOpener.WordFileFormat.Word2003Xml, "Word 2003 XML detected by its w:wordDocument root");
        string renamed = Path.Combine(Output, "word-xml-v1-flat-named.docx");
        File.Copy(flat, renamed, overwrite: true);
        try { TestReport.Check(WordWindowedDocumentOpener.DetectFormat(renamed) == WordWindowedDocumentOpener.WordFileFormat.FlatXmlPackage, "Flat OPC named .docx: still detected by content"); }
        finally { File.Delete(renamed); }

        string cache = Path.Combine(Output, "word-xml-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var inMemory = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = cache };
            var sidecar = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache };
            var onDemand = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = cache, IndexesOnDemand = true };
            (string html, string notes, int blocks) expected = Render(docx, inMemory);
            foreach ((string name, WordIndexPolicy policy) in new[] { ("in memory", inMemory), ("sidecar", sidecar), ("sidecar reopen", sidecar), ("on demand", onDemand) })
            {
                if (name == "on demand") Directory.Delete(cache, recursive: true);
                (string html, string notes, int blocks) actual = Render(flat, policy);
                TestReport.Check(actual.blocks == expected.blocks && Normalise(actual.html) == Normalise(expected.html) && actual.notes == expected.notes,
                                 $"Flat OPC ({name}): the same {expected.blocks} blocks and notes as the .docx ({actual.blocks} blocks)");
            }
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(flat, inMemory))
            {
                string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
                string? pictureId = Regex.Match(html, "data-word-picture=\"([^\"]*)\"").Groups[1].Value;
                TestReport.Check(document.ReadPicture(pictureId) is { } picture && FastDocView.Html.WordPictureSize.Read(picture.Bytes) is { } size && Math.Abs(size.WidthPt - 150) < 0.5,
                                 "Flat OPC: the picture decoded from its base64 part (150 pt wide)");
            }

            foreach ((string name, WordIndexPolicy policy) in new[] { ("in memory", inMemory), ("sidecar", sidecar), ("sidecar reopen", sidecar), ("on demand", onDemand) })
            {
                if (name == "on demand") Directory.Delete(cache, recursive: true);
                (string html, string notes, int blocks) actual = Render(word2003, policy);
                string text = Text(actual.html);
                TestReport.Check(actual.blocks == expected.blocks, $"Word 2003 XML ({name}): {expected.blocks} blocks like the .docx ({actual.blocks})");
                TestReport.Check(text == Text(expected.html), $"Word 2003 XML ({name}): the same text as the .docx  -- {Shorten(text)}");
                TestReport.Check(Regex.IsMatch(actual.html, "<h1[^>]*>.*Flat XML heading", RegexOptions.Singleline) || actual.html.Contains("Flat XML heading"), $"Word 2003 XML ({name}): heading");
                TestReport.Check(Regex.IsMatch(actual.html, "font-weight:\\s*bold[^>]*>bold") && Regex.IsMatch(actual.html, "font-style:\\s*italic[^>]*>italic"), $"Word 2003 XML ({name}): bold and italic runs");
                TestReport.Check(actual.html.Contains(">1.") && actual.html.Contains(">2."), $"Word 2003 XML ({name}): list labels 1. and 2.");
                TestReport.Check(actual.html.Contains("https://example.com/"), $"Word 2003 XML ({name}): the hyperlink's target");
                TestReport.Check(actual.html.Contains("<table") && actual.html.Contains("B2"), $"Word 2003 XML ({name}): the table");
                TestReport.Check(actual.html.Contains("data-word-picture="), $"Word 2003 XML ({name}): the picture");
                TestReport.Check(actual.notes.Contains("The footnote text.") && actual.notes.Contains("The endnote text."), $"Word 2003 XML ({name}): the footnote and the endnote  -- {Shorten(actual.notes)}");
            }
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(word2003, inMemory))
            {
                string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
                string? pictureId = Regex.Match(html, "data-word-picture=\"([^\"]*)\"").Groups[1].Value;
                TestReport.Check(document.ReadPicture(pictureId) is { } picture && FastDocView.Html.WordPictureSize.Read(picture.Bytes) is { } size && Math.Abs(size.WidthPt - 150) < 0.5,
                                 "Word 2003 XML: the picture decoded from its w:binData (150 pt wide)");
                TestReport.Check(Regex.Matches(html, "list-style|word-list-label|•||·").Count > 0, "Word 2003 XML: the bullet item has a label");
            }
            // Apache Tika's Word 2003 XML saved under a .doc name (a table inside a paragraph): detected by content, opens with its table
            string tika = Path.Combine(folder, "..", "apache-tika", "testWORDML_tableInsideParagraph.doc");
            if (File.Exists(tika))
            {
                TestReport.Check(WordWindowedDocumentOpener.DetectFormat(tika) == WordWindowedDocumentOpener.WordFileFormat.Word2003Xml, "Tika testWORDML_tableInsideParagraph.doc: Word 2003 XML by content");
                (string html, _, int blocks) = Render(tika, inMemory);
                TestReport.Check(blocks > 0 && html.Contains("<table"), $"Tika testWORDML_tableInsideParagraph.doc: opens with its table ({blocks} blocks) -- {Shorten(Text(html))}");
            }
            File.WriteAllText(Path.Combine(Output, "word-xml-2003.html"), Render(word2003, inMemory).html);
            File.WriteAllText(Path.Combine(Output, "word-xml-docx.html"), expected.html);
        }
        finally
        {
            try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { }
        }
    }

    private static (string html, string notes, int blocks) Render(string path, WordIndexPolicy policy)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        if (!document.IsIndexComplete) document.ScanAhead(int.MaxValue / 2);
        var blocks = new List<WordDocumentBlockHtml>();
        for (int first = 0; first < document.BlockCount; first += 3) blocks.AddRange(document.RenderBlocks(first, Math.Min(3, document.BlockCount - first)));   // small windows
        string notes = string.Join("|", blocks.SelectMany(block => block.Notes).Select(note => note.Kind + ":" + note.Label + ":" + Text(note.Html)));
        return (string.Concat(blocks.Select(block => block.Html)), notes, document.BlockCount);
    }

    /// <summary>Picture ids carry the part name and may differ in form; everything else must be the same.</summary>
    private static string Normalise(string html) => Regex.Replace(html, "data-word-picture=\"[^\"]*\"", "data-word-picture");

    private static string Text(string html) => Regex.Replace(System.Net.WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]*>", " ")), "\\s+", " ").Trim();

    private static string Shorten(string text)
    {
        var ascii = new System.Text.StringBuilder();
        foreach (char c in text.Length > 240 ? text.Substring(0, 240) : text) ascii.Append(c < 128 ? c.ToString() : $"U+{(int)c:X4}");
        return ascii.ToString();
    }
}
