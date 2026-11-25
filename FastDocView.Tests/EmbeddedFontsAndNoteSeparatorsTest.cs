using System.Text;
using System.Text.RegularExpressions;
using FastDocView;
using FastDocView.Docx;

namespace FastDocView.Tests;

/// <summary>
/// Embedded fonts and note separators (test-data/word2003-fonts/embedded-font-v1.*, Word 16: Caladea - a free font - regular +
/// bold, "Embed fonts in the file", subsetted; saved as .docx and as Word 2003 XML). The .docx faces reach the stylesheet as
/// @font-face rules (installed font first) and read de-obfuscated through ReadPicture; Word writes no font data to Word 2003
/// XML, so that file has no faces. Note separators (Word 2003 XML w:docPr/w:footnotePr, .docx separator notes) are not text:
/// a custom separator's text is neither shown nor found, the notes are numbered as without it.
/// </summary>
public static class EmbeddedFontsAndNoteSeparatorsTest
{
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word2003-fonts", name);
    private static string Output => Directory.CreateDirectory(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "TestOutput")).FullName;

    public static void Run()
    {
        TestReport.Step("Embedded fonts (.docx, Word 2003 XML) and note separators");
        // the key of ECMA-376 17.8.1's example GUID: its hex pairs from the last to the first (LibreOffice FontTable.cxx:280-293)
        byte[] key = DocxEmbeddedFonts.KeyOf("{62E79491-959F-41E9-B76B-6B32631DEA5C}");
        TestReport.Check(BitConverter.ToString(key) == "5C-EA-1D-63-32-6B-6B-B7-E9-41-9F-95-91-94-E7-62", $"font key from the GUID, last pair first ({BitConverter.ToString(key)})");
        TestReport.Check(DocxEmbeddedFonts.KeyOf("{not a guid}").Length == 0 && DocxEmbeddedFonts.KeyOf(null).Length == 0, "a missing or broken font key gives no key");

        string docx = Fixture("embedded-font-v1.docx"), xml = Fixture("embedded-font-v1.xml");
        if (!File.Exists(docx) || !File.Exists(xml)) { TestReport.Check(false, "test-data/word2003-fonts fixtures present"); return; }
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(docx, policy))
        {
            MatchCollection faces = Regex.Matches(document.Stylesheet, @"@font-face \{ font-family: '([^']+)'; font-weight: (\d+); font-style: (\w+); src: local\('([^']+)'\), url\(""fdv-font:([^""]+)""\); \}");
            string listed = string.Join(", ", faces.Cast<Match>().Select(m => $"{m.Groups[1].Value} {m.Groups[2].Value}"));
            TestReport.Check(faces.Count == 2 && listed.Contains("Caladea 400") && listed.Contains("Caladea 700"),
                $".docx: one @font-face per embedded face, installed font first ({listed})");
            TestReport.Check(faces.Cast<Match>().Any(m => m.Groups[4].Value == "Caladea Bold"), ".docx: a bold face's local() names its full name");
            int fonts = 0;
            foreach (Match face in faces)
            {
                string id = Uri.UnescapeDataString(face.Groups[5].Value);
                WordDocumentPicture? font = document.ReadPicture(id);
                bool isSfnt = font is { Bytes.Length: > 12 } && font.Bytes[0] == 0 && font.Bytes[1] == 1 && font.Bytes[2] == 0 && font.Bytes[3] == 0;
                // the table directory is past the obfuscated 32 bytes: its first tag must read as ASCII after de-obfuscation too
                bool tablesReadable = font is { Bytes.Length: > 16 } && font.Bytes.Skip(12).Take(4).All(b => b is >= 0x20 and < 0x7F);
                if (isSfnt && tablesReadable && font!.ContentType == "font/ttf") fonts++;
            }
            TestReport.Check(fonts == 2, $".docx: every embedded face reads back as a TrueType font, de-obfuscated ({fonts}/2)");
            TestReport.Check(document.ReadPicture("o:word/document.xml") is null && document.ReadPicture("o:word/fonts/none.odttf") is null,
                ".docx: only a part the font table names is served as a font");
        }

        // Word 2003 XML: Word saved the same document with embedTrueTypeFonts on and no font data (the format has no element for it)
        string xmlText = File.ReadAllText(xml, Encoding.UTF8);
        TestReport.Check(xmlText.Contains("<w:embedTrueTypeFonts/>") && !Regex.IsMatch(xmlText, @"<w:font\b[^>]*>(?:(?!</w:font>).)*<w:binData", RegexOptions.Singleline),
            "Word 2003 XML: Word writes the embed setting but no font data");
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(xml, policy))
        {
            TestReport.Check(!document.Stylesheet.Contains("@font-face"), "Word 2003 XML: no embedded faces");
            TestReport.Check(document.NoteCounts.footnotes == 1, $"Word 2003 XML: the separators are not notes ({document.NoteCounts.footnotes} footnote)");
        }

        // a custom separator with text (Word: References > Show Notes > Footnote Separator), in Word 2003 XML and in .docx
        string custom = Path.Combine(Output, "note-separator-custom-v1-" + Guid.NewGuid().ToString("N"));
        try
        {
            File.WriteAllText(custom + ".xml", xmlText.Replace("<w:r><w:separator/></w:r>", "<w:r><w:t>SEPARATORTEXT</w:t></w:r>")
                                                     .Replace("<w:r><w:continuationSeparator/></w:r>", "<w:r><w:t>CONTINUEDTEXT</w:t></w:r>"), new UTF8Encoding(false));
            using (var zip = System.IO.Compression.ZipFile.Open(custom + ".docx", System.IO.Compression.ZipArchiveMode.Create))
            using (var source = System.IO.Compression.ZipFile.OpenRead(docx))
                foreach (var entry in source.Entries)
                {
                    using Stream from = entry.Open();
                    using Stream to = zip.CreateEntry(entry.FullName).Open();
                    if (entry.FullName is "word/footnotes.xml" or "word/endnotes.xml")
                    {
                        string part = new StreamReader(from, Encoding.UTF8).ReadToEnd()
                            .Replace("<w:separator/>", "<w:t>SEPARATORTEXT</w:t>").Replace("<w:continuationSeparator/>", "<w:t>CONTINUEDTEXT</w:t>");
                        byte[] bytes = new UTF8Encoding(false).GetBytes(part);
                        to.Write(bytes, 0, bytes.Length);
                    }
                    else from.CopyTo(to);
                }
            foreach (string path in new[] { custom + ".xml", custom + ".docx" })
            {
                string kind = path.EndsWith(".xml") ? "Word 2003 XML" : ".docx";
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                string shown = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(b => b.Html))
                               + string.Concat(Enumerable.Range(0, document.NoteReferences.Count).SelectMany(n => document.RenderNotes(n, 1)).Select(n => n.Html));
                bool customInFile = File.ReadAllText(path).Contains("SEPARATORTEXT") || path.EndsWith(".docx");
                TestReport.Check(customInFile && !shown.Contains("SEPARATORTEXT") && !shown.Contains("CONTINUEDTEXT") && shown.Contains("Note body"),
                    $"{kind}: a custom separator is not shown, the note is");
                TestReport.Check(Hits(document, "SEPARATORTEXT") == 0 && Hits(document, "Note body") == 1, $"{kind}: search finds the note, not the separator");
                TestReport.Check(document.NoteCounts.footnotes == 1, $"{kind}: one footnote with a custom separator");
            }
        }
        finally
        {
            foreach (string path in new[] { custom + ".xml", custom + ".docx" }) if (File.Exists(path)) File.Delete(path);
        }
    }

    private static int Hits(IWordWindowedDocument document, string query)
    {
        int hits = 0, from = 0;
        while (true)
        {
            WordDocumentSearchBatch batch = document.Search(query, WordDocumentSearchOptions.Default, from, 500);
            hits += batch.HitCounts.Sum();
            if (batch.ReachedEnd) break;
            from = batch.NextBlock;
        }
        int note = 0;
        while (note < document.NoteReferences.Count)
        {
            WordDocumentSearchBatch batch = document.SearchNotes(query, WordDocumentSearchOptions.Default, note, 500);
            hits += batch.HitCounts.Sum();
            if (batch.ReachedEnd) break;
            note = batch.NextBlock;
        }
        return hits;
    }
}
