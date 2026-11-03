using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using FastDocView;
using FastDocView.Docx;
using FastDocView.Odt;
using SearchMode = FastDocView.Odt.OdtWindowedDocument.OdtSearchTextMode;

namespace FastDocView.Tests;

/// <summary>
/// The .odt / .fodt byte-level search (<see cref="OdtSearchTextScanner"/>): a word split by markup still matches, only
/// paragraph text is text, block boundaries found from the bytes - checked on a synthetic package and flat file, every
/// case against the expected hits AND against the XmlReader reference; the piece verdict (a scan that disagrees with the
/// index is never kept); and, with .odt / .fodt paths, the golden check (every block and count equal to the reference over
/// the whole document, several queries, match case on / off) with the profile (best of 3, alternated).
/// Hebrew is written as \u escapes.
/// </summary>
public static class OdtSearchBytesTest
{
    private static (int[] blocks, int[] counts) All(OdtWindowedDocument document, string query, SearchMode mode, WordDocumentSearchOptions? options = null, int batch = 500, int from = 0)
    {
        var hits = new List<int>();
        var counts = new List<int>();
        while (true)
        {
            WordDocumentSearchBatch result = document.Search(query, options ?? WordDocumentSearchOptions.Default, from, batch, mode);
            hits.AddRange(result.HitBlocks);
            counts.AddRange(result.HitCounts);
            if (result.ReachedEnd) break;
            from = result.NextBlock;
        }
        return (hits.ToArray(), counts.ToArray());
    }

    private static string Show(int[] values) => "[" + string.Join(",", values.Take(12)) + (values.Length > 12 ? ",..." : "") + "]";

    public static void Run(IReadOnlyList<string> goldenPaths)
    {
        TestReport.Step("Search bytes: piece verdicts");
        Verdicts();
        TestReport.Step("Search bytes odt / fodt: words split by markup");
        Synthetic();
        foreach (string path in goldenPaths) Golden(path);
    }

    /// <summary>A piece whose scan does not come out where the index says is redone block by block, never kept.</summary>
    private static void Verdicts()
    {
        TestReport.Check(WordSearchPiece.Judge(WordSearchBlocksResult.Stopped, isEndExact: true, stoppedAt: 40, length: 100) == WordSearchPieceVerdict.RedoPiece,
                         "verdict: Stopped inside bytes that end exactly at the piece's end is a failure (redo the piece)");
        TestReport.Check(WordSearchPiece.Judge(WordSearchBlocksResult.Stopped, isEndExact: false, stoppedAt: 40, length: 100) == WordSearchPieceVerdict.Accept,
                         "verdict: Stopped inside bytes that run on to the next anchor is the normal case");
        TestReport.Check(WordSearchPiece.Judge(WordSearchBlocksResult.Done, isEndExact: false, stoppedAt: 100, length: 100) == WordSearchPieceVerdict.RedoPiece,
                         "verdict: Done on bytes that run on to the next anchor is a failure");
        TestReport.Check(WordSearchPiece.Judge(WordSearchBlocksResult.Done, isEndExact: true, stoppedAt: 100, length: 100) == WordSearchPieceVerdict.Accept,
                         "verdict: Done at an exact end is the normal case");
        TestReport.Check(WordSearchPiece.Judge(WordSearchBlocksResult.Failed, isEndExact: true, stoppedAt: 7, length: 100) == WordSearchPieceVerdict.FallBackFromFailedBlock,
                         "verdict: Failed falls back from the failing block");

        // the .docx scanner over exact bytes holding one block more than the index says: Stopped before the end -> redo
        const string root = "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">";
        byte[] xml = Encoding.UTF8.GetBytes("<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p><w:p><w:r><w:t>three</w:t></w:r></w:p>");
        var docxScanner = new DocxSearchTextScanner(root);
        var matcher = new WordDocumentSearchMatcher("t");
        WordSearchBlocksResult docxResult = docxScanner.FeedBlocks(xml, 10, 12, matcher, _ => matcher.Reset(), out long docxStoppedAt);
        TestReport.Check(docxResult == WordSearchBlocksResult.Stopped && docxStoppedAt < xml.Length
                         && WordSearchPiece.Judge(docxResult, isEndExact: true, docxStoppedAt, xml.Length) == WordSearchPieceVerdict.RedoPiece,
                         $"docx: an extra block in exact bytes is caught ({docxResult} at {docxStoppedAt} of {xml.Length})");
        WordSearchBlocksResult fewer = docxScanner.FeedBlocks(xml, 10, 14, matcher, _ => matcher.Reset(), out long fewerAt);
        TestReport.Check(fewer == WordSearchBlocksResult.Failed && fewerAt == 10, $"docx: one block fewer than the index fails from the piece's first block ({fewer} at {fewerAt})");

        // the same for the .odt scanner
        const string odtRoot = "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">";
        byte[] odt = Encoding.UTF8.GetBytes("<text:p>one</text:p><text:p>two</text:p><text:p>three</text:p>");
        var odtScanner = new OdtSearchTextScanner(odtRoot);
        OdtSearchTextScanner.Frame[] frames = [OdtSearchTextScanner.Frame.Outer, OdtSearchTextScanner.Frame.Outer, OdtSearchTextScanner.Frame.Container];
        WordSearchBlocksResult odtResult = odtScanner.FeedBlocks(odt, 10, 12, frames, matcher, true, _ => matcher.Reset(), out long odtStoppedAt);
        TestReport.Check(odtResult == WordSearchBlocksResult.Stopped && WordSearchPiece.Judge(odtResult, isEndExact: true, odtStoppedAt, odt.Length) == WordSearchPieceVerdict.RedoPiece,
                         $"odt: an extra block in exact bytes is caught ({odtResult} at {odtStoppedAt} of {odt.Length})");
        WordSearchBlocksResult odtFewer = odtScanner.FeedBlocks(odt, 10, 14, frames, matcher, true, _ => matcher.Reset(), out long odtFewerAt);
        TestReport.Check(odtFewer == WordSearchBlocksResult.Failed && odtFewerAt == 10, $"odt: one block fewer than the index fails from the piece's first block ({odtFewer} at {odtFewerAt})");
    }

    private const string Namespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" office:version=\"1.3\"";

    private static string Body()
        => "<text:tracked-changes><text:changed-region text:id=\"c1\"><text:deletion><text:p>removed</text:p></text:deletion></text:changed-region></text:tracked-changes>"
         + "<text:sequence-decls><text:sequence-decl text:display-outline-level=\"0\" text:name=\"Table\"/></text:sequence-decls>"
         + "<text:p>wo<text:span text:style-name=\"T1\">rd</text:span></text:p>"                                                          // 0: partly formatted word
         + "<text:p>al<text:soft-page-break/>pha</text:p>"                                                                                // 1: soft page break mid-word
         + "<text:p>be<text:bookmark text:name=\"zeta\"/>t<text:bookmark-start text:name=\"m\"/>a<text:bookmark-end text:name=\"m\"/>"
             + "<text:reference-mark-start text:name=\"r\"/>x<text:reference-mark-end text:name=\"r\"/></text:p>"                         // 2: bookmarks, reference marks mid-word
         + "<text:p>gam<office:annotation office:name=\"n1\"><dc:creator>someone</dc:creator><text:p>inner</text:p></office:annotation>ma<office:annotation-end office:name=\"n1\"/></text:p>" // 3: comment inside a word
         + "<text:p>a<text:s text:c=\"3\"/>b c<text:s/>d e<text:s text:c=\"x\"/>f</text:p>"                                         // 4: text:s with and without c
         + "<text:p>x<text:tab/>y c<text:line-break/>d</text:p>"                                                                         // 5: tab = space, line break = separator
         + "<text:p>R&amp;D &lt;x&gt; &quot;q&apos; &#x5D0;&#1489;ג</text:p>"                                                         // 6: entities, character references
         + "<text:p>שׁ<text:span text:style-name=\"T2\">ָ</text:span>לוֹם</text:p>"                      // 7: niqqud in its own span
         + "<text:p>two  \n   terms</text:p>"                                                                                             // 8: white space collapses
         + "<text:list text:style-name=\"L1\"><text:list-item><text:p>listed one</text:p></text:list-item><text:list-header><text:p>listed two</text:p></text:list-header></text:list>" // 9, 10
         + "<text:section text:name=\"S\"><text:section-source xlink:href=\"x\" xmlns:xlink=\"http://www.w3.org/1999/xlink\"/><text:p>in section</text:p></text:section>"   // 11
         + "<text:table-of-content text:name=\"TOC\"><text:table-of-content-source><text:index-title-template>template</text:index-title-template></text:table-of-content-source>"
             + "<text:index-body><text:p>toc entry</text:p></text:index-body></text:table-of-content>"                                    // 12
         + "<text:p>ep<!-- note -->si<![CDATA[lon]]></text:p>"                                                                            // 13: a comment and CDATA inside a word
         + "<text:p>abc<text:note text:id=\"ftn1\" text:note-class=\"footnote\"><text:note-citation>1</text:note-citation><text:note-body><text:p>needle</text:p></text:note-body></text:note>def</text:p>" // 14
         + "<text:p/>"                                                                                                                   // 15: empty
         + "<table:table table:name=\"T\"><table:table-column/><table:table-row><table:table-cell><text:p>split</text:p><text:p>here</text:p></table:table-cell></table:table-row></table:table>" // 16
         + "<text:p xmlns:t2=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">other<t2:span>prefix</t2:span><t2:s t2:c=\"2\"/>z</text:p>" // 17: a namespace bound to another prefix
         + "<text:h text:outline-level=\"1\">\U0001D400plane<text:span>s</text:span></text:h>";                                          // 18: supplementary character, heading

    private static (string package, string flat) WriteSamples()
    {
        string package = Path.Combine(TestReport.OutputDirectory, "search-bytes-v1.odt");
        string flat = Path.Combine(TestReport.OutputDirectory, "search-bytes-v1.fodt");
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
            Add("content.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {Namespaces}><office:body><office:text>{Body()}</office:text></office:body></office:document-content>");
        }
        File.WriteAllText(flat, $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document {Namespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>{Body()}</office:text></office:body></office:document>", new UTF8Encoding(false));
        return (package, flat);
    }

    private static void Synthetic()
    {
        (string package, string flat) = WriteSamples();
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        foreach (string path in new[] { package, flat })
        {
            string format = Path.GetExtension(path).TrimStart('.');
            using var document = (OdtWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
            TestReport.Check(document.BlockCount == 19, $"{format} bytes: the sample has the expected 19 blocks ({document.BlockCount})");
            void Expect(string query, int[] blocks, string what, WordDocumentSearchOptions? options = null)
            {
                foreach (int batch in new[] { 3, 500 })
                {
                    (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options, batch);
                    (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options, batch);
                    TestReport.Check(hits.SequenceEqual(blocks), $"{format} bytes (batch {batch}): {what} {Show(hits)}");
                    TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts), $"{format} bytes = reference (batch {batch}): {what} ({Show(referenceHits)} x {Show(referenceCounts)})");
                }
            }
            Expect("word", [0], "text:span formatting inside a word");
            Expect("alpha", [1], "text:soft-page-break inside a word");
            Expect("betax", [2], "bookmarks and reference marks inside a word");
            Expect("zeta", [], "a bookmark name is not text");
            Expect("gamma", [3], "office:annotation inside a word is skipped");
            Expect("inner", [], "comment text is not matched");
            Expect("someone", [], "the comment's author is not matched");
            Expect("a   b", [4], "text:s text:c=3 is three spaces");
            Expect("c d", [4], "text:s with no c is one space");
            Expect("e f", [4], "text:s with a bad c is one space");
            Expect("x y", [5], "text:tab is a space");
            Expect("cd", [], "text:line-break is a separator");
            Expect("R&D <x> \"q'", [6], "predefined entities are decoded");
            Expect("אבג", [6], "hex + decimal character references are decoded");
            Expect("amp", [], "entity names are not text");
            Expect("שלום", [7], "niqqud in a separate span is ignored");
            Expect("שָׁלוֹם", [7], "a pointed query matches the split word");
            Expect("two terms", [8], "white space collapses to one space");
            Expect("listed", [9, 10], "list items and list headers are blocks");
            Expect("in section", [11], "a paragraph in a section");
            Expect("toc entry", [12], "an index body holds blocks");
            Expect("template", [], "an index source template is not content");
            Expect("removed", [], "tracked changes are not blocks");
            Expect("epsilon", [13], "a comment and CDATA inside a word");
            Expect("needle", [14], "a note body counts with its block");
            Expect("needle", [], "notes excluded", new WordDocumentSearchOptions(IncludeNotes: false));
            Expect("1", [], "the note citation is not text");
            Expect("abcdef", [14], "the note adds no character around it (notes excluded)", new WordDocumentSearchOptions(IncludeNotes: false));
            Expect("here", [16], "text in a table");
            Expect("splithere", [], "no match across paragraphs");
            Expect("otherprefix  z", [17], "elements are resolved by namespace, not prefix");
            Expect("\U0001D400planes", [18], "a supplementary character; a heading");
            Expect("WORD", [0], "MatchCase false", new WordDocumentSearchOptions(MatchCase: false));
            TestReport.Check(document.BlocksSearchedFromRawBytes > 0, $"{format} bytes: the raw-run path ran ({document.BlocksSearchedFromRawBytes} blocks)");
            // a batch begun out of order (not at an anchor, not where a batch stopped)
            (int[] late, _) = All(document, "listed", SearchMode.Bytes, batch: 2, from: 10);
            TestReport.Check(late.SequenceEqual([10]), $"{format} bytes: a batch begun out of order {Show(late)}");
        }
    }

    private static readonly string[] GoldenQueries =
    [
        "שבת",                                   // frequent Hebrew word, plain
        "שַׁבָּת",           // the same, pointed
        "אמר רבי",                // two words
        "ר״י",                                   // gershayim
        "ו־",                                         // maqaf
        "the", "a", "Shabbat", "qqzzqq",                        // English, a one-letter query, absent
        "ס׳ 1",                                       // mixed
    ];

    private static void Golden(string path)
    {
        if (!(path.EndsWith(".odt", StringComparison.OrdinalIgnoreCase) || path.EndsWith(".fodt", StringComparison.OrdinalIgnoreCase))) return;
        TestReport.Step($"Search bytes golden: {Path.GetFileName(path)}");
        var open = Stopwatch.StartNew();
        using var document = (OdtWindowedDocument)WordWindowedDocumentOpener.Open(path);
        Console.WriteLine($"   open {open.ElapsedMilliseconds} ms, {document.BlockCount} blocks");
        All(document, "warmup", SearchMode.Bytes, batch: 5000);
        int queryIndex = 0;
        foreach (string query in GoldenQueries)
        {
            queryIndex++;
            foreach (WordDocumentSearchOptions options in new[] { WordDocumentSearchOptions.Default, new WordDocumentSearchOptions(MatchCase: false) })
            {
                if (!options.MatchCase && !query.Any(char.IsLetter)) continue;
                var watch = Stopwatch.StartNew();
                (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options);
                long bytesMs = watch.ElapsedMilliseconds;
                watch.Restart();
                (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options);
                TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts),
                    $"golden query {queryIndex} (match case {options.MatchCase}): {hits.Length} hit blocks, {counts.Sum()} matches; reference {referenceHits.Length}, {referenceCounts.Sum()}"
                    + $" (bytes {bytesMs} ms, reference {watch.ElapsedMilliseconds} ms; only bytes: {Show(hits.Except(referenceHits).ToArray())}, only reference: {Show(referenceHits.Except(hits).ToArray())})");
            }
        }
        TestReport.Check(document.BlocksSearchedFromRawBytes > (long)document.BlockCount * 10, $"golden: the raw-run path read the blocks ({document.BlocksSearchedFromRawBytes})");
        // batch size and odd starts do not change the hits
        (int[] small, int[] smallCounts) = All(document, GoldenQueries[0], SearchMode.Bytes, batch: 37);
        (int[] large, int[] largeCounts) = All(document, GoldenQueries[0], SearchMode.Bytes, batch: 5000);
        TestReport.Check(small.SequenceEqual(large) && smallCounts.SequenceEqual(largeCounts), $"golden: batches of 37 = batches of 5000 ({small.Length} hit blocks)");
        int middle = document.BlockCount / 2 + 7;
        (int[] fromMiddle, _) = All(document, GoldenQueries[0], SearchMode.Bytes, from: middle);
        (int[] referenceFromMiddle, _) = All(document, GoldenQueries[0], SearchMode.ReferenceXmlReader, from: middle);
        TestReport.Check(fromMiddle.SequenceEqual(referenceFromMiddle), $"golden: a search begun at block {middle} ({fromMiddle.Length} hit blocks)");

        TestReport.Step("Search bytes profile (best of 3, alternated, full scan in batches of 500)");
        string profileQuery = GoldenQueries[0];
        var best = new Dictionary<string, long>();
        for (int round = 0; round < 3; round++)
            foreach ((string label, SearchMode mode) in new[] { ("reference XmlReader", SearchMode.ReferenceXmlReader), ("bytes", SearchMode.Bytes) })
            {
                var watch = Stopwatch.StartNew();
                All(document, profileQuery, mode);
                long ms = watch.ElapsedMilliseconds;
                best[label] = best.TryGetValue(label, out long previous) ? Math.Min(previous, ms) : ms;
                Console.WriteLine($"   round {round + 1}: {label}: {ms} ms");
            }
        foreach (var pair in best) Console.WriteLine($"   best: {pair.Key}: {pair.Value} ms");
        foreach (SearchMode mode in new[] { SearchMode.ReferenceXmlReader, SearchMode.Bytes })
        {
            using var fresh = (OdtWindowedDocument)WordWindowedDocumentOpener.Open(path);
            var first = Stopwatch.StartNew();
            WordDocumentSearchBatch firstBatch = fresh.Search(profileQuery, WordDocumentSearchOptions.Default, 0, 500, mode);
            Console.WriteLine($"   first batch of 500 ({mode}, fresh open): {first.Elapsed.TotalMilliseconds:F1} ms, {firstBatch.HitBlocks.Length} hits");
        }
    }
}
