using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using FastDocView;
using FastDocView.Docx;
using SearchMode = FastDocView.Docx.DocxWindowedDocument.DocxSearchTextMode;

namespace FastDocView.Tests;

/// <summary>
/// .docx note bodies searched from their raw bytes (<see cref="DocxNoteSearchTextScanner"/>): every split-word case on a
/// synthetic file against the expected hits AND the parsed-note reference; and a notes-heavy synthetic document
/// (100,000 footnotes) - golden check (blocks and counts equal the reference) and the timing, best of 3, alternated.
/// (.odt notes are inline in the body and go through the body's byte scan.) Hebrew is written as \u escapes.
/// </summary>
public static class DocxNoteSearchBytesTest
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

    public static void Run(bool withNotesHeavy)
    {
        TestReport.Step("Search bytes: note bodies, words split by markup");
        Synthetic();
        if (withNotesHeavy) NotesHeavy(100_000);
    }

    private static void Synthetic()
    {
        static string R(string text, string properties = "") => $"<w:r>{(properties.Length > 0 ? "<w:rPr>" + properties + "</w:rPr>" : "")}<w:t>{text}</w:t></w:r>";
        string[] notes =
        [
            "<w:p>" + R("wo") + R("rd", "<w:b/>") + "</w:p>",                                                                              // 0: two runs
            "<w:p>" + R("al") + "<w:proofErr w:type=\"spellStart\"/>" + R("ph") + "<w:bookmarkStart w:id=\"0\" w:name=\"zeta\"/>" + R("a") + "<w:bookmarkEnd w:id=\"0\"/></w:p>", // 1
            "<w:p>" + R("שׁ") + R("ָ", "<w:color w:val=\"FF0000\"/>") + R("לוֹם") + "</w:p>",          // 2: niqqud in its own run
            "<w:p>" + R("R&amp;D &lt;x&gt; &#x5D0;&#1489;ג") + "</w:p>",                                                              // 3: entities
            "<w:p>" + R("x") + "<w:r><w:tab/></w:r>" + R("y") + "<w:r><w:t>c</w:t><w:br/><w:t>d</w:t></w:r></w:p>",                       // 4: tab, break
            "<w:p>" + R("split") + "</w:p><w:p>" + R("here") + "</w:p>",                                                                  // 5: two paragraphs
            "<w:p>" + R("box") + "<w:r><mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"wps\"><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"/></w:drawing></mc:Choice>"
                + "<mc:Fallback><w:pict><v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\"><v:textbox><w:txbxContent><w:p>" + R("inside") + "</w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>" + R("after") + "</w:p>", // 6: drawing + text box
            "<w:p>" + R("a") + "<w:r><w:t> </w:t></w:r>" + R("b") + "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>" + R("e") + "</w:p>",  // 7: whitespace-only runs
            "<w:p><w:r><w:instrText xml:space=\"preserve\"> beta </w:instrText></w:r><w:del w:id=\"2\" w:author=\"a\"><w:r><w:delText>gamma</w:delText></w:r></w:del>" + R("kept") + "</w:p>", // 8
            "<w:p>" + "<w:r><w:t>ep<!-- c --></w:t></w:r>" + R("si") + "<w:r><w:t><![CDATA[lon]]></w:t></w:r>" + "</w:p>",                // 9: comment, CDATA
            "<w:p xmlns:v2=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><v2:r><v2:t>other</v2:t></v2:r>" + R("prefix") + "</w:p>", // 10
            "<w:p>" + R("w&#x20;s") + "<w:r><w:t>&#x20;</w:t></w:r>" + R("q") + "</w:p>",                                                  // 11: a space written as a reference
        ];
        var footnotes = new StringBuilder("<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"
                                        + "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>");
        var body = new StringBuilder();
        for (int i = 0; i < notes.Length; i++)
        {
            footnotes.Append($"<w:footnote w:id=\"{i + 1}\">{notes[i]}</w:footnote>");
            body.Append($"<w:p>{R("block" + i)}<w:r><w:footnoteReference w:id=\"{i + 1}\"/></w:r></w:p>");
        }
        DocxFidelityTest.RenderAll("search-note-bytes", body.ToString(), footnotes: footnotes.ToString());
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-search-note-bytes.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);

        void Expect(string query, int[] blocks, string what, WordDocumentSearchOptions? options = null)
        {
            (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options, batch: 3);
            (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options, batch: 3);
            TestReport.Check(hits.SequenceEqual(blocks), $"docx note bytes: {what} {Show(hits)}");
            TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts), $"docx note bytes = reference: {what} ({Show(referenceHits)} x {Show(referenceCounts)})");
        }
        Expect("word", [0], "a word split by runs");
        Expect("alpha", [1], "proofErr / bookmarks inside a word");
        Expect("zeta", [], "a bookmark name is not text");
        Expect("שלום", [2], "niqqud in a separate run");
        Expect("שָׁלוֹם", [2], "a pointed query");
        Expect("R&D <x>", [3], "entities decoded");
        Expect("אבג", [3], "character references decoded");
        Expect("x y", [4], "w:tab in a run is a space");
        Expect("cd", [], "w:br is a separator");
        Expect("splithere", [], "no match across note paragraphs");
        Expect("here", [5], "the note's second paragraph");
        Expect("inside", [6], "a text box in a note (left to the parser)");
        Expect("boxafter", [], "a drawing keeps its text box paragraph between the words (left to the parser)");
        Expect("ab", [7], "a whitespace-only run without xml:space is dropped (as the parser drops it)");
        Expect("b e", [7], "a whitespace-only run with xml:space preserve is a space");
        Expect("beta", [], "w:instrText is not text");
        Expect("gamma", [], "w:delText is not text");
        Expect("kept", [8], "text after a field code and a deletion");
        Expect("epsilon", [9], "a comment and CDATA inside a word");
        Expect("otherprefix", [10], "resolved by namespace, not prefix");
        Expect("w s", [11], "spaces written as references (left to the parser)");
        Expect("WORD", [0], "MatchCase false", new WordDocumentSearchOptions(MatchCase: false));
        Expect("word", [], "notes excluded", new WordDocumentSearchOptions(IncludeNotes: false));
        TestReport.Check(document.NotesSearchedFromBytes > 0, $"docx note bytes: the byte path ran ({document.NotesSearchedFromBytes} notes)");
    }

    /// <summary>A body of paragraphs that each reference one footnote; every note has a word split by runs, niqqud in a run, a bookmark, a tab.</summary>
    private static string WriteNotesHeavy(int noteCount)
    {
        string path = Path.Combine(TestReport.OutputDirectory, $"notes-heavy-{noteCount}-v1.docx");
        if (File.Exists(path)) return path;
        TestReport.Step($"Writing {path}");
        const string w = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";
        using var file = File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);
        void Text(string name, string text)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name, CompressionLevel.Optimal).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
        Text("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/></Types>");
        Text("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        Text("word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/></Relationships>");
        string[] words = ["שבת", "אמר", "רבי", "הלכה", "Shabbat", "the", "כתב", "ס׳", "1", "ו־א"];
        var random = new Random(4242);
        using (var notes = new StreamWriter(zip.CreateEntry("word/footnotes.xml", CompressionLevel.Optimal).Open(), new UTF8Encoding(false), 1 << 16))
        {
            notes.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:footnotes {w}>");
            notes.Write("<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>");
            for (int i = 1; i <= noteCount; i++)
            {
                notes.Write($"<w:footnote w:id=\"{i}\"><w:p><w:pPr><w:pStyle w:val=\"FootnoteText\"/></w:pPr><w:r><w:rPr><w:rStyle w:val=\"FootnoteReference\"/></w:rPr><w:footnoteRef/></w:r>");
                notes.Write("<w:r><w:t xml:space=\"preserve\"> שׁ</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ָבָּת</w:t></w:r>");   // a pointed word split by runs
                notes.Write($"<w:bookmarkStart w:id=\"{i}\" w:name=\"n{i}\"/><w:proofErr w:type=\"spellStart\"/>");
                for (int k = 0, n = 4 + random.Next(20); k < n; k++) notes.Write($"<w:r><w:t xml:space=\"preserve\"> {words[random.Next(words.Length)]}</w:t></w:r>");
                notes.Write($"<w:proofErr w:type=\"spellEnd\"/><w:bookmarkEnd w:id=\"{i}\"/><w:r><w:tab/></w:r><w:r><w:t>{i}</w:t></w:r></w:p>");
                if (i % 3 == 0) notes.Write($"<w:p><w:r><w:t xml:space=\"preserve\">{words[random.Next(words.Length)]} &amp; {words[random.Next(words.Length)]}</w:t></w:r></w:p>");
                notes.Write("</w:footnote>");
            }
            notes.Write("</w:footnotes>");
        }
        using (var document = new StreamWriter(zip.CreateEntry("word/document.xml", CompressionLevel.Optimal).Open(), new UTF8Encoding(false), 1 << 16))
        {
            document.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document {w}><w:body>");
            for (int i = 1; i <= noteCount; i++)
                document.Write($"<w:p><w:r><w:t xml:space=\"preserve\">p{i} {words[random.Next(words.Length)]}</w:t></w:r><w:r><w:footnoteReference w:id=\"{i}\"/></w:r></w:p>");
            document.Write("<w:sectPr/></w:body></w:document>");
        }
        return path;
    }

    private static void NotesHeavy(int noteCount)
    {
        string path = WriteNotesHeavy(noteCount);
        TestReport.Step($"Search note bytes golden: {Path.GetFileName(path)} ({new FileInfo(path).Length >> 20} MB)");
        var open = Stopwatch.StartNew();
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path);
        Console.WriteLine($"   open {open.ElapsedMilliseconds} ms, {document.BlockCount} blocks");
        All(document, "warmup", SearchMode.Bytes, batch: 5000);
        string[] queries = ["שבת", "שַׁבָּת", "אמר רבי", "Shabbat", "the & ", "ס׳ 1", "ו א", "qqzzqq"];
        int queryIndex = 0;
        foreach (string query in queries)
        {
            queryIndex++;
            foreach (WordDocumentSearchOptions options in new[] { WordDocumentSearchOptions.Default, new WordDocumentSearchOptions(MatchCase: false) })
            {
                var watch = Stopwatch.StartNew();
                (int[] hits, int[] counts) = All(document, query, SearchMode.Bytes, options);
                long bytesMs = watch.ElapsedMilliseconds;
                watch.Restart();
                (int[] referenceHits, int[] referenceCounts) = All(document, query, SearchMode.ReferenceXmlReader, options);
                TestReport.Check(hits.SequenceEqual(referenceHits) && counts.SequenceEqual(referenceCounts),
                    $"notes golden query {queryIndex} (match case {options.MatchCase}): {hits.Length} hit blocks, {counts.Sum()} matches; reference {referenceHits.Length}, {referenceCounts.Sum()}"
                    + $" (bytes {bytesMs} ms, reference {watch.ElapsedMilliseconds} ms)");
            }
        }
        TestReport.Check(document.NotesSearchedFromBytes >= noteCount, $"notes golden: the byte path read the notes ({document.NotesSearchedFromBytes})");
        TestReport.Step("Search note bytes profile (best of 3, alternated, full scan in batches of 500)");
        var best = new Dictionary<string, long>();
        for (int round = 0; round < 3; round++)
            foreach ((string label, SearchMode mode, bool notes) in new[]
            {
                ("bytes + notes parsed (before)", SearchMode.BytesWithParsedNotes, true),
                ("bytes, no notes", SearchMode.Bytes, false),
                ("bytes + notes from bytes", SearchMode.Bytes, true),
            })
            {
                var watch = Stopwatch.StartNew();
                All(document, queries[0], mode, new WordDocumentSearchOptions(IncludeNotes: notes));
                long ms = watch.ElapsedMilliseconds;
                best[label] = best.TryGetValue(label, out long previous) ? Math.Min(previous, ms) : ms;
                Console.WriteLine($"   round {round + 1}: {label}: {ms} ms");
            }
        foreach (var pair in best) Console.WriteLine($"   best: {pair.Key}: {pair.Value} ms");
    }
}
