using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// In-document search (<see cref="IWordWindowedDocument.Search"/>): the app's normalisation (niqqud / teamim ignored,
/// maqaf a space, quote glyphs folded), case, matches across runs, never across paragraphs / cells, deleted text and
/// field codes not matched, notes optional - on every format; and, with paths, the batch / full-scan timings.
/// Hebrew is written as \u escapes (the tooling's traffic must stay ASCII).
/// </summary>
public static class DocumentSearchTest
{
    private const string Shalom = "\u05E9\u05DC\u05D5\u05DD";                                  // the query, plain letters
    private const string ShalomPointed = "\u05E9\u05B8\u05C1\u05DC\u05D5\u05B9\u0591\u05DD";    // qamats, shin dot, holam, etnahta

    public static void Run(IReadOnlyList<string> timingPaths)
    {
        TestReport.Step("Search: docx");
        Docx();
        TestReport.Step("Search: odt / fodt");
        Odt();
        TestReport.Step("Search: rtf");
        Rtf();
        TestReport.Step("Search: txt");
        Text();
        foreach (string path in timingPaths) Timing(path);
    }

    private static int[] All(IWordWindowedDocument document, string query, WordDocumentSearchOptions? options = null, int batch = 2)
    {
        var hits = new List<int>();
        int from = 0;
        while (true)
        {
            WordDocumentSearchBatch result = document.Search(query, options ?? WordDocumentSearchOptions.Default, from, batch);
            hits.AddRange(result.HitBlocks);
            if (result.ReachedEnd) { if (result.NextBlock != -1) throw new InvalidOperationException("ReachedEnd with a next block"); break; }
            if (result.NextBlock <= from) throw new InvalidOperationException("Search did not advance");
            from = result.NextBlock;
        }
        return hits.ToArray();
    }

    private static string Show(int[] hits) => "[" + string.Join(",", hits) + "]";

    private static void Docx()
    {
        static string R(string text, string properties = "") => $"<w:r>{(properties.Length > 0 ? "<w:rPr>" + properties + "</w:rPr>" : "")}<w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        string body =
            "<w:p>" + R("\u05E9\u05B8\u05C1") + R("\u05DC\u05D5\u05B9\u0591\u05DD", "<w:b/>") + R(" world") + "</w:p>"                       // 0: split across runs + marks
          + "<w:p>" + R("Hello World") + "</w:p>"                                                                                             // 1: case
          + "<w:p>" + R("abc") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>"                                                         // 2: note
          + "<w:tbl><w:tr><w:tc><w:p>" + R("split") + "</w:p><w:p>" + R("here") + "</w:p></w:tc><w:tc><w:p>" + R("cell") + "</w:p></w:tc></w:tr></w:tbl>" // 3: table
          + "<w:p>" + R("\u05D1\u05BE\u05D0 \u05E8\u05D1\u05F4\u05D9") + "</w:p>"                                                             // 4: maqaf, gershayim
          + "<w:p>" + R("kept ") + "<w:del w:id=\"9\" w:author=\"a\"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>"                   // 5: deleted text
          + "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + R("7") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"; // 6: field
        string footnotes = "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                         + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
                         + "<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r>" + R(" needle in a note") + "</w:p></w:footnote>";
        DocxFidelityTest.RenderAll("search", body, footnotes: footnotes);
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-search.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        Common(document, "docx", shalomBlock: 0, helloBlock: 1, noteBlock: 2, tableBlock: 3, maqafBlock: 4);
        TestReport.Check(All(document, "gone").Length == 0, "docx: deleted text (w:delText) is not matched");
        TestReport.Check(All(document, "kept").SequenceEqual([5]), "docx: the paragraph around a deletion still matches");
        TestReport.Check(All(document, "PAGE").Length == 0, "docx: a field code (w:instrText) is not matched");
        TestReport.Check(All(document, "7").SequenceEqual([6]), "docx: a field result is matched");
    }

    /// <summary>Checks every format shares: marks, runs, case, separators, maqaf / gershayim, notes, batches, empty query.</summary>
    private static void Common(IWordWindowedDocument document, string format, int shalomBlock, int helloBlock, int noteBlock, int tableBlock, int maqafBlock)
    {
        int[] hits = All(document, Shalom);
        TestReport.Check(hits.SequenceEqual([shalomBlock]), $"{format}: plain query matches pointed text (niqqud / teamim ignored){(format == "docx" ? ", across runs" : "")} {Show(hits)}");
        TestReport.Check(All(document, ShalomPointed).SequenceEqual([shalomBlock]), $"{format}: a pointed query matches too");
        TestReport.Check(All(document, "hello").Length == 0, $"{format}: case is matched exactly by default (as the app does)");
        TestReport.Check(All(document, "hello", new WordDocumentSearchOptions(MatchCase: false)).SequenceEqual([helloBlock]), $"{format}: MatchCase false folds case");
        TestReport.Check(All(document, "needle").SequenceEqual([noteBlock]), $"{format}: a note body makes its referencing block a hit {Show(All(document, "needle"))}");
        TestReport.Check(All(document, "needle", new WordDocumentSearchOptions(IncludeNotes: false)).Length == 0, $"{format}: IncludeNotes false skips note bodies");
        TestReport.Check(All(document, "here").SequenceEqual([tableBlock]), $"{format}: text in a table matches");
        TestReport.Check(All(document, "splithere").Length == 0 && All(document, "herecell").Length == 0, $"{format}: no match across paragraphs / cells");
        TestReport.Check(All(document, "\u05D1 \u05D0").SequenceEqual([maqafBlock]), $"{format}: the maqaf matches a space");
        TestReport.Check(All(document, "\u05E8\u05D1\"\u05D9").SequenceEqual([maqafBlock]), $"{format}: gershayim match an ASCII double quote");
        TestReport.Check(All(document, "world", batch: 1000).SequenceEqual(All(document, "world", batch: 1)), $"{format}: batch size does not change the hits");
        WordDocumentSearchBatch empty = document.Search("  \u05B8 ", WordDocumentSearchOptions.Default, 0, 10);
        TestReport.Check(empty.ReachedEnd && empty.HitBlocks.Length == 0, $"{format}: an empty (marks-only) query ends at once");
        WordDocumentSearchBatch past = document.Search("x", WordDocumentSearchOptions.Default, document.BlockCount + 5, 10);
        TestReport.Check(past.ReachedEnd && past.NextBlock == -1, $"{format}: a start past the end reports the end");
    }

    private const string OdfNamespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" office:version=\"1.3\"";

    private static string OdfBody()
        => "<text:p>\u05E9\u05B8\u05C1<text:span>\u05DC\u05D5\u05B9\u0591\u05DD</text:span> world</text:p>"   // 0
         + "<text:p>Hello World</text:p>"                                                                       // 1
         + "<text:p>abc<text:note text:id=\"ftn1\" text:note-class=\"footnote\"><text:note-citation>1</text:note-citation><text:note-body><text:p>needle in a note</text:p></text:note-body></text:note></text:p>" // 2
         + "<table:table table:name=\"T\"><table:table-column/><table:table-column/><table:table-row><table:table-cell><text:p>split</text:p><text:p>here</text:p></table:table-cell><table:table-cell><text:p>cell</text:p></table:table-cell></table:table-row></table:table>" // 3
         + "<text:p>\u05D1\u05BE\u05D0 \u05E8\u05D1\u05F4\u05D9</text:p>"                                       // 4
         + "<text:p>a<text:s text:c=\"2\"/>b<office:annotation><text:p>comment zzz</text:p></office:annotation></text:p>"; // 5

    private static void Odt()
    {
        string package = Path.Combine(TestReport.OutputDirectory, "search-v1.odt");
        string flat = Path.Combine(TestReport.OutputDirectory, "search-v1.fodt");
        if (File.Exists(package)) File.Delete(package);
        using (var file = File.Create(package))
        using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
        {
            void Add(string entry, string text, CompressionLevel level = CompressionLevel.Optimal)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entry, level).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
            Add("META-INF/manifest.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>");
            Add("content.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {OdfNamespaces}><office:body><office:text>{OdfBody()}</office:text></office:body></office:document-content>");
        }
        File.WriteAllText(flat, $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>{OdfBody()}</office:text></office:body></office:document>", new UTF8Encoding(false));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        foreach (string path in new[] { package, flat })
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            string format = Path.GetExtension(path).TrimStart('.');
            Common(document, format, 0, 1, 2, 3, 4);
            TestReport.Check(All(document, "a  b").SequenceEqual([5]), $"{format}: text:s counts as its spaces");
            TestReport.Check(All(document, "zzz").Length == 0, $"{format}: comments (office:annotation) are not matched");
        }
    }

    private static string RtfUnicode(string text)
    {
        var rtf = new StringBuilder();
        foreach (char c in text) rtf.Append(c < 128 ? c.ToString() : $"\\u{(short)c}?");
        return rtf.ToString();
    }

    private static void Rtf()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "search-v1.rtf");
        string rtf = "{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\n"
            + "\\pard " + RtfUnicode("\u05E9\u05B8\u05C1") + "{\\b " + RtfUnicode("\u05DC\u05D5\u05B9\u0591\u05DD") + "} world\\par\n"   // 0
            + "\\pard Hello World\\par\n"                                                                                                 // 1
            + "\\pard abc{\\super\\chftn}{\\footnote\\pard{\\super\\chftn} needle in a note}\\par\n"                                     // 2
            + "\\trowd\\cellx3000\\cellx6000\\pard\\intbl split\\line here\\cell cell\\cell\\row\n"                                    // 3
            + "\\pard " + RtfUnicode("\u05D1\u05BE\u05D0 \u05E8\u05D1\u05F4\u05D9") + "\\par\n"                                         // 4
            + "\\pard {\\listtext 1.\\tab}listed\\par\n"                                                                               // 5
            + "}";
        File.WriteAllText(path, rtf, Encoding.ASCII);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        int Block(string word) => All(document, word).FirstOrDefault(-1);
        Common(document, "rtf", Block("world"), Block("Hello"), Block("abc"), Block("split"), Block("\u05D1"));
        TestReport.Check(All(document, "1.").Length == 0 && Block("listed") >= 0, "rtf: a \\listtext label is not matched, its paragraph is");
    }

    private static void Text()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "search-v1.txt");
        File.WriteAllText(path, string.Join("\r\n", ShalomPointed + " world", "Hello World", "abc", "here", "\u05D1\u05BE\u05D0 \u05E8\u05D1\u05F4\u05D9", "a & b"), new UTF8Encoding(true));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        TestReport.Check(All(document, Shalom).SequenceEqual([0]), "txt: niqqud ignored");
        TestReport.Check(All(document, "hello", new WordDocumentSearchOptions(MatchCase: false)).SequenceEqual([1]), "txt: MatchCase false");
        TestReport.Check(All(document, "\u05E8\u05D1\"\u05D9").SequenceEqual([4]), "txt: gershayim fold");
        TestReport.Check(All(document, "a & b").SequenceEqual([5]), "txt: a plain-text '&' matches literally (the line is text, not HTML)");
        TestReport.Check(All(document, "abchere").Length == 0, "txt: no match across lines");
    }

    /// <summary>Batch and full-scan timings on a real file: a frequent word (time to the first hit) and a word that is absent (a full scan).</summary>
    private static void Timing(string path)
    {
        TestReport.Step($"Search timing: {Path.GetFileName(path)}");
        CrossCheckWithRendering(path, "\u05E9\u05D1\u05EA");
        string coldCache = Path.Combine(TestReport.OutputDirectory, "search-cold-cache");
        foreach ((string label, string query, int batch, bool cold) in new[] { ("frequent", "\u05E9\u05D1\u05EA", 500, false), ("absent", "qqzzqq", 500, false), ("absent, cold first open", "qqzzqq", 500, true) })
        {
            if (cold && Directory.Exists(coldCache)) Directory.Delete(coldCache, recursive: true);
            var open = Stopwatch.StartNew();
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, cold ? new WordIndexPolicy { CacheDirectory = coldCache, IndexesOnDemand = true } : null);
            open.Stop();
            bool completeAtOpen = document.IsIndexComplete;
            var watch = Stopwatch.StartNew();
            long firstHitMs = -1, worstBatchMs = 0;
            int batches = 0, hits = 0, from = 0;
            while (true)
            {
                long before = watch.ElapsedMilliseconds;
                WordDocumentSearchBatch result = document.Search(query, WordDocumentSearchOptions.Default, from, batch);
                worstBatchMs = Math.Max(worstBatchMs, watch.ElapsedMilliseconds - before);
                batches++;
                hits += result.HitBlocks.Length;
                if (firstHitMs < 0 && result.HitBlocks.Length > 0) firstHitMs = watch.ElapsedMilliseconds;
                if (result.ReachedEnd) break;
                from = result.NextBlock;
            }
            Console.WriteLine($"   {label}: open {open.ElapsedMilliseconds} ms (index complete at open: {completeAtOpen}), {document.BlockCount} blocks, "
                            + $"{hits} hit blocks, first hit {firstHitMs} ms, full scan {watch.ElapsedMilliseconds} ms, {batches} batches of {batch}, "
                            + $"avg {watch.ElapsedMilliseconds / Math.Max(1, batches)} ms, worst {worstBatchMs} ms, heap {GC.GetTotalMemory(false) >> 20} MB");
        }
    }

    /// <summary>The search's hits over the first blocks must equal the blocks whose RENDERED html (+ notes) holds the query, walked as the app walks html.</summary>
    private static void CrossCheckWithRendering(string path, string query)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        const int count = 3000;
        int[] searched = document.Search(query, WordDocumentSearchOptions.Default, 0, count).HitBlocks;
        var expected = new List<int>();
        var matcher = new WordDocumentSearchMatcher(query);
        for (int first = 0; first < Math.Min(count, document.BlockCount); first += 200)
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, Math.Min(200, count - first)))
            {
                matcher.Reset();
                matcher.FeedHtml(block.Html);
                foreach (WordDocumentNoteHtml note in block.Notes) { matcher.Separator(); matcher.FeedHtml(note.Html); }
                if (matcher.Found) expected.Add(block.BlockIndex);
            }
        int[] onlySearch = searched.Except(expected).ToArray(), onlyRendered = expected.Except(searched).ToArray();
        TestReport.Check(onlySearch.Length == 0 && onlyRendered.Length == 0,
            $"{Path.GetFileName(path)}: first {count} blocks: search hits {searched.Length} = blocks whose rendered html holds the query {expected.Count}"
            + $" (only search: {Show(onlySearch.Take(10).ToArray())}, only rendered: {Show(onlyRendered.Take(10).ToArray())})");
    }
}
