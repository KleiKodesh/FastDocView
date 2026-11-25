using System.IO.Compression;
using System.Text;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Search during a first open (the index still building) over blocks whose shown text depends on what only the final index
/// knows - the facts that make a rendered block provisional (<see cref="WordDocumentBlockHtml.IsProvisional"/>): a web-archive
/// later-sibling rule (display:none, ::before content) whose parent's end is not indexed yet, a REF to a bookmark not reached yet
/// (.docx, .rtf, .odt). The batch names those blocks (<see cref="WordDocumentSearchBatch.ProvisionalBlocks"/>); the host searches
/// them again once the index is complete. Checked: the partial hits, the provisional list, and that the partial hits with the
/// provisional blocks searched again equal - exactly, block by block and count by count - a search over the complete index.
/// </summary>
internal static class ProvisionalSearchTest
{
    public static void Run()
    {
        TestReport.Step("Search during a first open: provisional blocks searched again once the index is complete");
        string directory = TestReport.OutputDirectory;
        string cache = Path.Combine(directory, "index-cache-provisional-search-" + Guid.NewGuid().ToString("N"));
        var onDemand = new WordIndexPolicy { CacheDirectory = cache, ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        var files = new List<string>();
        try
        {
            string archive = WriteArchive(Path.Combine(directory, "provisional-search-v1-" + Guid.NewGuid().ToString("N") + ".mht"), 40_000);
            files.Add(archive);
            RunOne("web archive ::before", archive, onDemand, "Zebra", partialShouldHit: false, finalShouldHit: true);
            RunOne("web archive display:none", archive, onDemand, "Quokka", partialShouldHit: true, finalShouldHit: false);
            RunOne("web archive unaffected text", archive, onDemand, "Paragraph 3", partialShouldHit: true, finalShouldHit: true);

            string rtf = Path.Combine(directory, "provisional-search-v1-" + Guid.NewGuid().ToString("N") + ".rtf");
            files.Add(rtf);
            var body = new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard Early {\field{\*\fldinst REF late }} end.\par" + "\r\n");
            for (int i = 0; i < 20_000; i++) body.Append(@"\pard Filler paragraph ").Append(i).Append(@".\par").Append("\r\n");
            body.Append(@"\pard {\*\bkmkstart late}Target words{\*\bkmkend late} here.\par}");
            File.WriteAllText(rtf, body.ToString(), Encoding.ASCII);
            RunOne(".rtf REF to a late bookmark", rtf, onDemand, "Target", partialShouldHit: false, finalShouldHit: true);

            string odt = Path.Combine(directory, "provisional-search-v1-" + Guid.NewGuid().ToString("N") + ".fodt");
            files.Add(odt);
            var odf = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>\n");
            odf.Append("<text:p>Early <text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"late\">old</text:bookmark-ref> end.</text:p>\n");
            for (int i = 0; i < 20_000; i++) odf.Append("<text:p>Filler paragraph ").Append(i).Append(".</text:p>\n");
            odf.Append("<text:p><text:bookmark-start text:name=\"late\"/>Target words<text:bookmark-end text:name=\"late\"/></text:p>\n</office:text></office:body></office:document>");
            File.WriteAllText(odt, odf.ToString(), new UTF8Encoding(false));
            RunOne(".odt bookmark-ref to a late bookmark", odt, onDemand, "Target", partialShouldHit: false, finalShouldHit: true);

            string docx = Path.Combine(directory, "provisional-search-v1-" + Guid.NewGuid().ToString("N") + ".docx");
            files.Add(docx);
            WriteDocx(docx);
            RunOne(".docx REF to a late bookmark", docx, onDemand, "Target", partialShouldHit: false, finalShouldHit: true);

            // notes: a footnote whose REF reads a bookmark 20,000 paragraphs on (the notes search and the notes pane, per note)
            string rtfNotes = Path.Combine(directory, "provisional-notes-v1-" + Guid.NewGuid().ToString("N") + ".rtf");
            files.Add(rtfNotes);
            var rtfNoteBody = new StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}\pard Plain{\super\chftn}{\footnote\pard{\super\chftn} Plain note.}\par" + "\r\n");
            rtfNoteBody.Append(@"\pard Early{\super\chftn}{\footnote\pard{\super\chftn} See {\field{\*\fldinst REF late }} end.}\par" + "\r\n");
            for (int i = 0; i < 20_000; i++) rtfNoteBody.Append(@"\pard Filler paragraph ").Append(i).Append(@".\par").Append("\r\n");
            rtfNoteBody.Append(@"\pard {\*\bkmkstart late}Target words{\*\bkmkend late} here.\par}");
            File.WriteAllText(rtfNotes, rtfNoteBody.ToString(), Encoding.ASCII);
            RunNotes(".rtf note REF to a late bookmark", rtfNotes, onDemand, "Target");

            string odtNotes = Path.Combine(directory, "provisional-notes-v1-" + Guid.NewGuid().ToString("N") + ".fodt");
            files.Add(odtNotes);
            var odfNotes = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>\n");
            odfNotes.Append("<text:p>Plain<text:note text:id=\"n1\" text:note-class=\"footnote\"><text:note-citation>1</text:note-citation><text:note-body><text:p>Plain note.</text:p></text:note-body></text:note></text:p>\n");
            odfNotes.Append("<text:p>Early<text:note text:id=\"n2\" text:note-class=\"footnote\"><text:note-citation>2</text:note-citation><text:note-body><text:p>See <text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"late\">old</text:bookmark-ref> end.</text:p></text:note-body></text:note></text:p>\n");
            for (int i = 0; i < 20_000; i++) odfNotes.Append("<text:p>Filler paragraph ").Append(i).Append(".</text:p>\n");
            odfNotes.Append("<text:p><text:bookmark-start text:name=\"late\"/>Target words<text:bookmark-end text:name=\"late\"/></text:p>\n</office:text></office:body></office:document>");
            File.WriteAllText(odtNotes, odfNotes.ToString(), new UTF8Encoding(false));
            RunNotes(".odt note bookmark-ref to a late bookmark", odtNotes, onDemand, "Target");

            string docxNotes = Path.Combine(directory, "provisional-notes-v1-" + Guid.NewGuid().ToString("N") + ".docx");
            files.Add(docxNotes);
            WriteDocx(docxNotes, withNotes: true);
            RunNotes(".docx note REF to a late bookmark", docxNotes, onDemand, "Target");
        }
        finally
        {
            foreach (string file in files) if (File.Exists(file)) File.Delete(file);
            if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true);
        }
    }

    private const int PartialBlocks = 50;

    private static void RunOne(string name, string path, WordIndexPolicy policy, string query, bool partialShouldHit, bool finalShouldHit)
    {
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        var options = new WordDocumentSearchOptions();
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        TestReport.Check(!document.IsIndexComplete, $"{name}: the first open indexes on demand");
        WordDocumentSearchBatch partial = document.Search(query, options, 0, PartialBlocks);
        TestReport.Check(!document.IsIndexComplete, $"{name}: the index is still building after the partial search");
        var counts = new SortedDictionary<int, int>();
        for (int i = 0; i < partial.HitBlocks.Length; i++) counts[partial.HitBlocks[i]] = partial.HitCounts[i];
        TestReport.Check(partial.HitBlocks.Length > 0 == partialShouldHit,
                         $"{name}: over the partial index '{query}' {(partialShouldHit ? "is" : "is not")} found (hits [{string.Join(",", partial.HitBlocks)}])");
        bool expectsProvisional = partialShouldHit != finalShouldHit;
        TestReport.Check(!expectsProvisional || partial.ProvisionalBlocks.Length > 0,
                         $"{name}: the partial batch names its provisional blocks ([{Head(partial.ProvisionalBlocks)}])");
        TestReport.Check(partial.ProvisionalBlocks.All(b => b >= 0 && b < PartialBlocks) && partial.ProvisionalBlocks.SequenceEqual(partial.ProvisionalBlocks.Distinct().OrderBy(b => b)),
                         $"{name}: provisional blocks lie in the scanned range, ascending, each once");

        while (!document.IsIndexComplete) document.ScanAhead(5_000);
        // the host's refresh: each provisional block searched again, its hits replaced (a count of 0 drops it)
        bool anyStillProvisional = false;
        foreach (int block in partial.ProvisionalBlocks)
        {
            WordDocumentSearchBatch again = document.Search(query, options, block, 1);
            anyStillProvisional |= again.ProvisionalBlocks.Length > 0;
            counts.Remove(block);
            if (again.HitBlocks.Length == 1 && again.HitBlocks[0] == block) counts[block] = again.HitCounts[0];
        }
        TestReport.Check(!anyStillProvisional, $"{name}: nothing is provisional when searched again on the complete index");
        WordDocumentSearchBatch final = document.Search(query, options, 0, PartialBlocks);
        TestReport.Check(final.ProvisionalBlocks.Length == 0, $"{name}: a batch over the complete index has no provisional blocks");
        TestReport.Check(final.HitBlocks.Length > 0 == finalShouldHit,
                         $"{name}: over the complete index '{query}' {(finalShouldHit ? "is" : "is not")} found (hits [{string.Join(",", final.HitBlocks)}])");
        string patched = string.Join(" ", counts.Select(pair => $"{pair.Key}:{pair.Value}"));
        string expected = string.Join(" ", final.HitBlocks.Select((block, i) => $"{block}:{final.HitCounts[i]}"));
        TestReport.Check(patched == expected, $"{name}: partial hits with the provisional blocks searched again = the complete index's hits exactly ([{patched}] vs [{expected}])");
    }

    /// <summary>
    /// The notes search and the notes pane over a partial index, per note: the note whose REF reads a bookmark not reached yet
    /// is drawn provisionally (<see cref="WordDocumentNoteHtml.IsProvisional"/>) and named by the notes batch; the plain note is
    /// not. Once the index is complete the note renders final, and the partial hits with the provisional notes searched again
    /// equal the complete index's notes hits exactly.
    /// </summary>
    private static void RunNotes(string name, string path, WordIndexPolicy policy, string query)
    {
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        var options = new WordDocumentSearchOptions();
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        document.RenderBlocks(0, PartialBlocks);   // the reader's first window: its notes become known
        TestReport.Check(!document.IsIndexComplete, $"{name}: the first open indexes on demand");
        int known = document.NoteReferences.Count;
        TestReport.Check(known == 2, $"{name}: both notes are known from the first window ({known})");
        IReadOnlyList<WordDocumentNoteHtml> drawn = document.RenderNotesForDisplay(0, known);
        TestReport.Check(drawn.Count == 2 && !drawn[0].IsProvisional && drawn[1].IsProvisional,
                         $"{name}: the notes pane draws the REF note provisionally, the plain note final ([{string.Join(",", drawn.Select(n => n.IsProvisional))}])");
        WordDocumentSearchBatch partial = document.SearchNotes(query, options, 0, known);
        TestReport.Check(!document.IsIndexComplete, $"{name}: the index is still building after the partial notes search");
        TestReport.Check(partial.HitBlocks.Length == 0, $"{name}: over the partial index '{query}' is not found in the notes (hits [{string.Join(",", partial.HitBlocks)}])");
        TestReport.Check(partial.ProvisionalBlocks.SequenceEqual(new[] { 1 }), $"{name}: the notes batch names exactly the provisional note ([{Head(partial.ProvisionalBlocks)}])");
        var counts = new SortedDictionary<int, int>();
        for (int i = 0; i < partial.HitBlocks.Length; i++) counts[partial.HitBlocks[i]] = partial.HitCounts[i];

        while (!document.IsIndexComplete) document.ScanAhead(5_000);
        IReadOnlyList<WordDocumentNoteHtml> refreshed = document.RenderNotesForDisplay(1, 1);
        TestReport.Check(refreshed.Count == 1 && !refreshed[0].IsProvisional && refreshed[0].Html.Contains("Target"),
                         $"{name}: fetched again on the complete index the note is final and shows the bookmark's text ({(refreshed.Count == 1 ? System.Text.RegularExpressions.Regex.Replace(refreshed[0].Html, "<[^>]+>", "") : "-")})");
        bool anyStillProvisional = false;
        foreach (int note in partial.ProvisionalBlocks)
        {
            WordDocumentSearchBatch again = document.SearchNotes(query, options, note, 1);
            anyStillProvisional |= again.ProvisionalBlocks.Length > 0;
            counts.Remove(note);
            if (again.HitBlocks.Length == 1 && again.HitBlocks[0] == note) counts[note] = again.HitCounts[0];
        }
        TestReport.Check(!anyStillProvisional, $"{name}: no note is provisional when searched again on the complete index");
        WordDocumentSearchBatch final = document.SearchNotes(query, options, 0, int.MaxValue);
        TestReport.Check(final.ProvisionalBlocks.Length == 0 && final.ReachedEnd, $"{name}: a notes batch over the complete index has no provisional notes");
        string patched = string.Join(" ", counts.Select(pair => $"{pair.Key}:{pair.Value}"));
        string expected = string.Join(" ", final.HitBlocks.Select((note, i) => $"{note}:{final.HitCounts[i]}"));
        TestReport.Check(expected.Length > 0 && patched == expected,
                         $"{name}: partial notes hits with the provisional notes searched again = the complete index's notes hits exactly ([{patched}] vs [{expected}])");
    }

    private static string Head(int[] blocks) => string.Join(",", blocks.Take(8)) + (blocks.Length > 8 ? $",... ({blocks.Length})" : "");

    /// <summary>One long div: its first paragraph gets generated text and its second is hidden by rules that need the div's end.</summary>
    private static string WriteArchive(string path, int paragraphs)
    {
        const string Boundary = "----=_NextPart_01DB0000.00000003";
        var body = new StringBuilder(paragraphs * 40);
        body.Append("MIME-Version: 1.0\r\nContent-Type: multipart/related; boundary=\"").Append(Boundary).Append("\"\r\n\r\n--").Append(Boundary).Append("\r\n")
            .Append("Content-Location: file:///C:/provisional-search.htm\r\nContent-Transfer-Encoding: 8bit\r\nContent-Type: text/html; charset=\"us-ascii\"\r\n\r\n")
            .Append("<html>\r\n<head>\r\n<style>\r\n<!--\r\n")
            .Append("div.long > p:nth-last-of-type(").Append(paragraphs).Append(")::before { content:\"Zebra \"; }\r\n")
            .Append("div.long > p:nth-last-of-type(").Append(paragraphs - 1).Append(") { display:none; }\r\n-->\r\n</style>\r\n</head>\r\n<body>\r\n<div class=long>\r\n");
        for (int i = 1; i <= paragraphs; i++) body.Append("<p>").Append(i == 2 ? "Quokka" : "Paragraph").Append(' ').Append(i).Append("</p>\r\n");
        body.Append("</div>\r\n</body>\r\n</html>\r\n\r\n--").Append(Boundary).Append("--\r\n");
        File.WriteAllText(path, body.ToString(), new UTF8Encoding(false));
        return path;
    }

    private static void WriteDocx(string path, bool withNotes = false)
    {
        const string W = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";
        var body = new StringBuilder();
        if (withNotes)
            body.Append("<w:p><w:r><w:t>Plain</w:t></w:r><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"1\"/></w:r></w:p>")
                .Append("<w:p><w:r><w:t>Early</w:t></w:r><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"2\"/></w:r></w:p>")
                .Append("<w:p><w:pPr><w:sectPr/></w:pPr></w:p>");   // the notes' section ends early: their numbering is known (the REF alone is provisional)
        else
            body.Append("<w:p><w:r><w:t xml:space=\"preserve\">Early </w:t></w:r><w:fldSimple w:instr=\" REF late \"/><w:r><w:t xml:space=\"preserve\"> end.</w:t></w:r></w:p>");
        for (int i = 0; i < 20_000; i++) body.Append("<w:p><w:r><w:t>Filler paragraph ").Append(i).Append(".</w:t></w:r></w:p>");
        body.Append("<w:p><w:bookmarkStart w:id=\"1\" w:name=\"late\"/><w:r><w:t>Target words</w:t></w:r><w:bookmarkEnd w:id=\"1\"/></w:p>");
        body.Append("<w:sectPr/>");
        const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
        var parts = new Dictionary<string, string>
        {
            ["[Content_Types].xml"] = "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>",
            ["_rels/.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"{Rel}officeDocument\" Target=\"word/document.xml\"/></Relationships>",
            ["word/document.xml"] = $"<w:document {W}><w:body>{body}</w:body></w:document>",
        };
        if (withNotes)
        {
            parts["[Content_Types].xml"] = parts["[Content_Types].xml"].Replace("</Types>",
                "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/></Types>");
            parts["word/_rels/document.xml.rels"] = $"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" Type=\"{Rel}footnotes\" Target=\"footnotes.xml\"/></Relationships>";
            const string Mark = "<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteRef/></w:r>";
            parts["word/footnotes.xml"] = $"<w:footnotes {W}>"
                + "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                + $"<w:footnote w:id=\"1\"><w:p>{Mark}<w:r><w:t xml:space=\"preserve\"> Plain note.</w:t></w:r></w:p></w:footnote>"
                + $"<w:footnote w:id=\"2\"><w:p>{Mark}<w:r><w:t xml:space=\"preserve\"> See </w:t></w:r><w:fldSimple w:instr=\" REF late \"/><w:r><w:t xml:space=\"preserve\"> end.</w:t></w:r></w:p></w:footnote>"
                + "</w:footnotes>";
        }
        using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var part in parts)
            using (var writer = new StreamWriter(zip.CreateEntry(part.Key).Open())) writer.Write(part.Value);
    }
}
