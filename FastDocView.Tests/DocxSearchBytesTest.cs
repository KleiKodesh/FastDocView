using System.Diagnostics;
using FastDocView;
using FastDocView.Docx;
using SearchMode = FastDocView.Docx.DocxWindowedDocument.DocxSearchTextMode;

namespace FastDocView.Tests;

/// <summary>
/// The .docx byte-level search (<see cref="DocxSearchTextScanner"/>): a word split by markup still matches, only w:t is
/// text - checked on a synthetic file, every case against the expected hits AND against the XmlReader reference; and, with
/// paths, the golden check (every block and count equal to the reference over the whole document, several queries) with
/// the profile (read only / reference / bytes, best of 3, alternated). Hebrew is written as \u escapes.
/// </summary>
public static class DocxSearchBytesTest
{
    private static (int[] blocks, int[] counts) All(DocxWindowedDocument document, string query, SearchMode mode, WordDocumentSearchOptions? options = null, int batch = 500)
    {
        var hits = new List<int>();
        var counts = new List<int>();
        int from = 0;
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
        TestReport.Step("Search bytes: words split by markup");
        Synthetic();
        foreach (string path in goldenPaths) Golden(path);
    }

    private static void Synthetic()
    {
        static string R(string text, string properties = "") => $"<w:r>{(properties.Length > 0 ? "<w:rPr>" + properties + "</w:rPr>" : "")}<w:t>{text}</w:t></w:r>";
        string body =
            "<w:p>" + R("wo") + R("rd", "<w:b/>") + "</w:p>"                                                                                        // 0: partly formatted word
          + "<w:p>" + R("al") + "<w:proofErr w:type=\"spellStart\"/>" + R("ph") + "<w:bookmarkStart w:id=\"0\" w:name=\"zeta\"/><w:commentRangeStart w:id=\"1\"/>"
              + R("a") + "<w:bookmarkEnd w:id=\"0\"/><w:proofErr w:type=\"spellEnd\"/><w:commentRangeEnd w:id=\"1\"/></w:p>"                        // 1: proofErr / bookmark / comment range inside a word
          + "<w:p>" + R("שׁ") + R("ָ", "<w:color w:val=\"FF0000\"/>") + R("לוֹם") + "</w:p>"                 // 2: niqqud in its own run
          + "<w:p>" + R("R&amp;D &lt;x&gt; &quot;q&apos;") + R("&#x5D0;&#1489;ג") + "</w:p>"                                               // 3: entities, character references
          + "<w:p><w:r><w:t xml:space=\"preserve\">two </w:t></w:r>" + R("terms") + "<w:r><w:t>a</w:t></w:r><w:r><w:t> </w:t></w:r>" + R("b") + "</w:p>" // 4: xml:space, a space-only run
          + "<w:p><w:r><w:instrText xml:space=\"preserve\"> beta </w:instrText></w:r><w:del w:id=\"2\" w:author=\"a\"><w:r><w:delText>gamma</w:delText></w:r></w:del>"
              + "<w:bookmarkStart w:id=\"3\" w:name=\"delta\"/>" + R("kept") + "</w:p>"                                                             // 5: field code, deleted text, attribute values
          + "<w:p>" + R("x") + "<w:r><w:tab/></w:r>" + R("y") + "<w:r><w:t>c</w:t><w:br/><w:t>d</w:t></w:r></w:p>"                               // 6: tab = space, break = separator
          + "<w:p><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"wps\">" + R("choice")
              + "</mc:Choice><mc:Fallback>" + R("fallback") + "</mc:Fallback></mc:AlternateContent>" + R("after") + "</w:p>"                         // 7: mc:Fallback skipped
          + "<w:p>" + "<w:r><w:t>ep<!-- note --></w:t></w:r>" + R("si") + "<w:r><w:t><![CDATA[lon]]></w:t></w:r>" + "</w:p>"                       // 8: a comment and CDATA inside a word
          + "<w:p xmlns:v=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><v:r><v:t>other</v:t></v:r>" + R("prefix") + "</w:p>"  // 9: namespace bound to another prefix
          + "<w:p>" + R("split") + "</w:p><w:p>" + R("here") + "</w:p>"                                                                           // 10, 11
          + "<w:p>" + R("𝐀plane") + "<w:r><w:rPr><w:rFonts w:ascii=\"x&gt;y\"/></w:rPr><w:t>s</w:t></w:r></w:p>"                      // 12: supplementary character; '>' entity in an attribute
          // 13: a picture: its title / description are attributes (wp:docPr, pic:cNvPr), never text; the renderer shows [image]
          + "<w:p>" + R("pic") + "<w:r><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:docPr id=\"1\" name=\"Picture 1\" title=\"pictitle\" descr=\"picdesc\"/>"
              + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"x\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"0\" name=\"picname\" descr=\"picdesc2\"/></pic:nvPicPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>" + R("ture") + "</w:p>"
          // 14: a text box (mc:Choice, rendered) with its VML copy (mc:Fallback, not rendered): the text once
          + "<w:p><w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"wps\"><w:drawing><wp:anchor xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><wp:docPr id=\"2\" name=\"Text Box 2\" descr=\"boxdesc\"/>"
              + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"y\"><wps:wsp xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:txbx><w:txbxContent><w:p>" + R("boxtext") + "</w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>"
              + "<mc:Fallback><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\" alt=\"vmlalt\"><v:textbox><w:txbxContent><w:p>" + R("boxtext") + "</w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>";
        DocxFidelityTest.RenderAll("search-bytes", body);
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-search-bytes.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);

        void Expect(string query, int[] blocks, string what, WordDocumentSearchOptions? options = null)
        {
            (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options, batch: 3);
            (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options, batch: 3);
            TestReport.Check(hits.SequenceEqual(blocks), $"docx bytes: {what} {Show(hits)}");
            TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts), $"docx bytes = reference: {what} ({Show(referenceHits)} x {Show(referenceCounts)})");
        }
        Expect("word", [0], "a partly formatted word (two runs) matches");
        Expect("alpha", [1], "proofErr / bookmarkStart / commentRange inside a word");
        Expect("שלום", [2], "niqqud in a separate run is ignored");
        Expect("שָׁלוֹם", [2], "a pointed query matches the split word");
        Expect("R&D <x> \"q'", [3], "predefined entities are decoded");
        Expect("אבג", [3], "hex + decimal character references are decoded");
        Expect("amp", [], "entity names are not text");
        Expect("two terms", [4], "xml:space preserve keeps the space");
        Expect("a b", [4], "a space-only run is a space (as the reference reads it)");
        Expect("beta", [], "w:instrText is not text");
        Expect("gamma", [], "w:delText is not text");
        Expect("delta", [], "attribute values are not text");
        Expect("zeta", [], "a bookmark name is not text");
        Expect("kept", [5], "text after a field code and a deletion");
        Expect("x y", [6], "w:tab is a space");
        Expect("cd", [], "w:br is a separator");
        Expect("choice", [7], "mc:Choice is read");
        Expect("fallback", [], "mc:Fallback is skipped");
        Expect("after", [7], "text after mc:AlternateContent");
        Expect("epsilon", [8], "a comment and CDATA inside a word");
        Expect("otherprefix", [9], "elements are resolved by namespace, not prefix");
        Expect("splithere", [], "no match across paragraphs");
        Expect("𝐀planes", [12], "a supplementary character; a '>' entity in an attribute");
        Expect("WORD", [0], "MatchCase false", new WordDocumentSearchOptions(MatchCase: false));
        Expect("pictitle", [], "a picture's title (wp:docPr) is not text");
        Expect("picdesc", [], "a picture's description (wp:docPr / pic:cNvPr) is not text");
        Expect("picname", [], "a picture's name is not text");
        Expect("boxtext", [14], "a text box's text counts once (mc:Fallback copy skipped)");
        Expect("boxdesc", [], "a text box's description is not text");
        Expect("vmlalt", [], "a VML shape's alt text is not text");
        foreach ((int block, string shown, string hidden) in new[] { (13, "pic", "pictitle"), (14, "boxtext", "boxdesc") })
        {
            string html = document.RenderBlocks(block, 1)[0].Html;
            TestReport.Check(html.Contains(shown) && !html.Contains(hidden), $"docx: block {block} renders \"{shown}\", not \"{hidden}\" (search = rendered text)");
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
        if (!path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) return;
        TestReport.Step($"Search bytes golden: {Path.GetFileName(path)}");
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path);
        All(document, "warmup", SearchMode.Bytes, batch: 5000);
        int queryIndex = 0;
        foreach (string query in GoldenQueries)
        {
            queryIndex++;
            foreach (WordDocumentSearchOptions options in new[] { WordDocumentSearchOptions.Default, new WordDocumentSearchOptions(MatchCase: false) })
            {
                if (!options.MatchCase && !query.Any(char.IsLetter)) continue;
                (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options);
                (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options);
                TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts),
                    $"golden query {queryIndex} (match case {options.MatchCase}): {hits.Length} hit blocks, {counts.Sum()} matches; reference {referenceHits.Length}, {referenceCounts.Sum()}"
                    + $" (only bytes: {Show(hits.Except(referenceHits).ToArray())}, only reference: {Show(referenceHits.Except(hits).ToArray())})");
            }
        }
        TestReport.Step("Search bytes profile (best of 3, alternated, full scan in batches of 500)");
        string profileQuery = GoldenQueries[0];
        var best = new Dictionary<string, long>();
        for (int round = 0; round < 3; round++)
            foreach ((string label, SearchMode mode) in new[]
            {
                ("read bytes only, index ranges (span scan)", SearchMode.ReadBytesOnly),
                ("read bytes only, raw search read", SearchMode.ReadSearchBytesOnly),
                ("reference XmlReader", SearchMode.ReferenceXmlReader),
                ("bytes", SearchMode.Bytes),
            })
            {
                var watch = Stopwatch.StartNew();
                All(document, profileQuery, mode, WordDocumentSearchOptions.Default);
                long ms = watch.ElapsedMilliseconds;
                best[label] = best.TryGetValue(label, out long previous) ? Math.Min(previous, ms) : ms;
                Console.WriteLine($"   round {round + 1}: {label}: {ms} ms");
            }
        foreach (var pair in best) Console.WriteLine($"   best: {pair.Key}: {pair.Value} ms");
        var first = Stopwatch.StartNew();
        WordDocumentSearchBatch firstBatch = document.Search(profileQuery, WordDocumentSearchOptions.Default, 0, 500);
        Console.WriteLine($"   first batch of 500: {first.Elapsed.TotalMilliseconds:F1} ms, {firstBatch.HitBlocks.Length} hits");
    }
}
