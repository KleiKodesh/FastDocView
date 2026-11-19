using System.IO.Compression;
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

            // a host's Decide then Open of a Flat OPC file: one pass over the file, not two
            string flatCopy = Path.Combine(Output, "word-xml-v1-scan-once-" + Guid.NewGuid().ToString("N") + ".xml");
            File.Copy(flat, flatCopy);
            try
            {
                int scansBefore = FastDocView.Docx.DocxFlatPackageStream.ScanCount;
                WordWindowedDocumentOpener.Decide(flatCopy, onDemand);
                using (IWordWindowedDocument opened = WordWindowedDocumentOpener.Open(flatCopy, onDemand)) opened.RenderBlocks(0, 1);
                int scans = FastDocView.Docx.DocxFlatPackageStream.ScanCount - scansBefore;
                TestReport.Check(scans == 1, $"Flat OPC: Decide + Open scan the file once ({scans})");
                File.AppendAllText(flatCopy, "\n");                       // changed: scanned again
                File.SetLastWriteTimeUtc(flatCopy, DateTime.UtcNow.AddMinutes(1));
                scansBefore = FastDocView.Docx.DocxFlatPackageStream.ScanCount;
                WordWindowedDocumentOpener.Decide(flatCopy, onDemand);
                int rescans = FastDocView.Docx.DocxFlatPackageStream.ScanCount - scansBefore;
                TestReport.Check(rescans == 1, $"Flat OPC: a changed file is scanned again ({rescans})");
            }
            finally { File.Delete(flatCopy); }
            CheckStreamingBase64(docx);

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
            // a picture bullet (Word: w:listPicBullet in 2003 XML, w:numPicBullet in .docx): the same picture label as Word's .docx
            string bulletDocx = Path.Combine(folder, "word-xml-picture-bullet-v1.docx"), bullet2003 = Path.Combine(folder, "word-xml-picture-bullet-v1-2003.xml");
            if (File.Exists(bulletDocx) && File.Exists(bullet2003))
            {
                var docxBullet = Render(bulletDocx, inMemory); var xmlBullet = Render(bullet2003, inMemory);
                int Pictures(string html) => Regex.Matches(html, "<img").Count;
                TestReport.Check(Pictures(docxBullet.html) == 2 && Pictures(xmlBullet.html) == 2 && Normalise(Text(xmlBullet.html)) == Normalise(Text(docxBullet.html)),
                                 $"Word 2003 XML picture bullet: two picture labels like the .docx ({Pictures(xmlBullet.html)} vs {Pictures(docxBullet.html)})");
                using IWordWindowedDocument bulletDocument = WordWindowedDocumentOpener.Open(bullet2003, inMemory);
                string bulletHtml = bulletDocument.RenderBlocks(0, 1)[0].Html;
                string bulletId = Regex.Match(bulletHtml, "word-picture-bullet[^>]*data-word-picture=\"([^\"]*)\"").Groups[1].Value;
                TestReport.Check(bulletId.Length > 0 && bulletDocument.ReadPicture(bulletId) is { } bulletPicture && FastDocView.Html.WordPictureSize.Read(bulletPicture.Bytes) is not null, $"Word 2003 XML picture bullet: the label is a picture -- {Shorten(bulletHtml.Substring(Math.Max(0, bulletHtml.IndexOf("list-marker"))))}");
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

    /// <summary>
    /// A Flat OPC file with a 12 MB binary part (base64 in 76-character lines, as Word writes it): the part reads back
    /// exactly, and decoding it allocates about the part itself, not its base64 text as well (once ~5x the part).
    /// </summary>
    private static void CheckStreamingBase64(string docx)
    {
        string path = Path.Combine(Output, "flat-big-binary-v1-" + Guid.NewGuid().ToString("N") + ".xml");
        var big = new byte[12 << 20];
        new Random(7).NextBytes(big);
        try
        {
            using (var writer = new StreamWriter(path, false, new System.Text.UTF8Encoding(false)))
            {
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">");
                void Part(string name, string type, byte[] bytes)
                {
                    writer.Write($"<pkg:part pkg:name=\"/{name}\" pkg:contentType=\"{type}\" pkg:compression=\"store\"><pkg:binaryData>");
                    string text = Convert.ToBase64String(bytes);
                    for (int at = 0; at < text.Length; at += 76) { writer.Write(text.AsSpan(at, Math.Min(76, text.Length - at))); writer.Write("\r\n"); }
                    writer.Write("</pkg:binaryData></pkg:part>");
                }
                using (ZipArchive zip = ZipFile.OpenRead(docx))
                    foreach (ZipArchiveEntry entry in zip.Entries)
                    {
                        using var data = new MemoryStream();
                        using (Stream stream = entry.Open()) stream.CopyTo(data);
                        Part(entry.FullName, "application/octet-stream", data.ToArray());
                    }
                Part("word/media/big.bin", "application/octet-stream", big);
                writer.Write("</pkg:package>");
            }
            using var package = new FastDocView.Docx.DocxFlatPackageStream(WordDocumentFileStream.OpenShared(path));
            long before = GC.GetAllocatedBytesForCurrentThread();
            var entryLocation = SeekableDeflate.ZipEntryLocator.Locate(package, "word/media/big.bin");
            TestReport.Check(entryLocation is not null, "Flat OPC big binary: the part is in the package");
            if (entryLocation is null) return;
            package.Position = entryLocation.DataOffset;
            var first = new byte[1];
            package.Read(first, 0, 1);                                   // the part is decoded by the first read that reaches it (here or in Locate)
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
            var back = new byte[big.Length];
            back[0] = first[0];
            int filled = 1;
            while (filled < back.Length) { int read = package.Read(back, filled, back.Length - filled); if (read <= 0) break; filled += read; }
            TestReport.Check(back.AsSpan().SequenceEqual(big), "Flat OPC big binary: decoded exactly");
            TestReport.Check(allocated >= big.Length && allocated < big.Length * 1.2, $"Flat OPC big binary: decoding allocated {allocated >> 20} MB for a {big.Length >> 20} MB part (streamed, < 1.2x)");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            TestReport.Check(document.BlockCount > 0, $"Flat OPC with every part in base64: opens ({document.BlockCount} blocks)");
        }
        finally { File.Delete(path); }
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
