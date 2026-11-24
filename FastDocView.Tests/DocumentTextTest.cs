using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// <see cref="WordDocumentText"/> (GetText / GetNotes / GetTextWithNotes), one real file per format, notes included: the
/// text equals what the search reads EXACTLY, per block and per note - the search's scan read one block / one note at a
/// time gives the same text as GetText's batches, the golden reference readers (.docx / .odt XmlReader, .doc block ranges)
/// agree, and for every word of a block / note the search's count in that block / note equals the count in its text.
/// GetTextWithNotes is the body, a blank line and the notes; a cancelled token stops it. (The app's full-text index
/// does not index document contents - its local-file index holds names and folders only - so this is the one text path.)
/// </summary>
public static class DocumentTextTest
{
    private static readonly string[] SampleFiles =
    [
        "test-data/apache-poi/footnote.doc", "test-data/apache-poi/footnotes.docx", "test-data/apache-poi/endnotes.docx",
        "test-data/apache-tika/testRTF_annotation_spacing.rtf", "test-data/libreoffice/sw/qa/extras/rtfexport/data/tdf87034.rtf", "test-data/libreoffice/sw/qa/core/layout/data/inline-endnote-and-section.odt",
        "test-data/libreoffice/sw/qa/core/text/data/footnote-tag.fodt", "test-data/apache-tika/english.txt",
        "test-data/apache-tika/testMHTMLFirefox.mhtml", "test-data/word-mhtml/pictures-v1.mht",
    ];

    public static void Run()
    {
        TestReport.Step("Text API: GetText / GetNotes / GetTextWithNotes equal what the search indexes (one file per format)");
        string root = FindRepositoryRoot();
        foreach (string file in SampleFiles)
        {
            string path = Path.Combine(root, file);
            string name = Path.GetFileName(path);
            if (!File.Exists(path)) { TestReport.Check(false, $"{name}: sample missing"); continue; }
            string body = WordDocumentText.GetText(path);
            string notes = WordDocumentText.GetNotes(path);
            string both = WordDocumentText.GetTextWithNotes(path);
            TestReport.Check(body.Length > 0, $"{name}: body text read ({body.Length} chars, notes {notes.Length})");
            TestReport.Check(both == (notes.Length == 0 ? body : body + "\n\n" + notes), $"{name}: with notes = body, blank line, notes");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            TestReport.Check(WordDocumentText.GetText(document) == body && WordDocumentText.GetNotes(document) == notes, $"{name}: open document gives the same text");
            if (document.NoteReferences.Count > 0) TestReport.Check(notes.Length > 0, $"{name}: {document.NoteReferences.Count} notes have text");
            CheckBlocks(name, document, body);
            CheckNotes(name, document, notes);
            CheckWords(name, "body", body, word => SumCounts(from => document.Search(word, WordDocumentSearchOptions.Default, from, 1000)));
            CheckWords(name, "notes", notes, word => SumCounts(from => document.SearchNotes(word, WordDocumentSearchOptions.Default, from, 1000)));
        }
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        bool stopped = false;
        try { WordDocumentText.GetText(Path.Combine(root, SampleFiles[1]), cancelled.Token); }
        catch (OperationCanceledException) { stopped = true; }
        TestReport.Check(stopped, "a cancelled token stops GetText");
        CheckWhileIndexing();
        CheckNoteInFieldKeyword(root);
        CheckFieldLeadingSwitch(root);
        CheckEquationSearchText(root);
    }

    /// <summary>
    /// test-data/word-fields/eq-search-text-v1 (.rtf / .docx / .doc, Word 16): the text Word itself gives an EQ field
    /// (eq-search-text-v1.word-find.txt): Range.Text and Find skip every EQ drawing (fraction, root, overstrike, array, bracket,
    /// script, box, integral, list), while a phonetic guide (Word saves it to .docx as w:ruby) is found as guide + base, one
    /// run with the text around it ("gdbase", ": gd", "base tail"; never "basegd"). GetText and the search read exactly that.
    /// </summary>
    private static void CheckEquationSearchText(string root)
    {
        string[] cases = { "f", "r", "o", "a", "b", "s", "x", "i", "l" };
        string expected = string.Join("\n", cases.Select(name => $"case {name}:  tail{name}").Append("case ruby: gdbase tailruby"));
        foreach (string extension in new[] { "rtf", "docx", "doc" })
        {
            string name = "eq-search-text-v1." + extension;
            string path = Path.Combine(root, "test-data", "word-fields", name);
            string body = WordDocumentText.GetText(path);
            TestReport.Check(body == expected, $"{name}: GetText = Word's text of every EQ exactly (got \"{Ascii(body.Replace('\n', '|'))}\")");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            CheckBlocks(name, document, body);
            foreach ((string word, int count) in new[] { ("ab", 0), ("cd", 0), ("xy", 0), ("ef", 0), ("pq", 0), ("mn", 0), ("up", 0), ("bx", 0), ("ig", 0), ("l1", 0),
                                                         ("gd", 1), ("base", 1), ("gdbase", 1), ("basegd", 0), (": gd", 1), ("base tail", 1) })
            {
                int found = SumCounts(from => document.Search(word, WordDocumentSearchOptions.Default, from, 1000));
                TestReport.Check(found == count, $"{name}: search finds \"{word}\" {count} time(s), as Word's Find ({found})");
            }
        }
    }

    /// <summary>
    /// test-data/word-ruby/field-leading-switch-v1 (.rtf and Word's own .docx / .doc): a code that starts with a switch or with
    /// a backslash before a non-keyword has no type in Word (Field.Type -1) and shows nothing; a backslash before a keyword is
    /// skipped. GetText, the search and the rendered text all read Word's printed text (field-leading-switch-v1.*.pdf). Case I
    /// nests a MERGEFIELD with a note in its broken keyword inside an EQ with a note of its own: the EQ draws ITS note (ii),
    /// the inner field's note keeps its number and draws nothing (each field's code notes are its own, .doc as .docx).
    /// </summary>
    private static void CheckFieldLeadingSwitch(string root)
    {
        const string order = "ABCDEFGHJKLI";
        var shown = new Dictionary<char, string> { ['E'] = "resE", ['H'] = "resH" };
        string[] lines = order.Select(letter => $"case {letter}: {(shown.TryGetValue(letter, out string? result) ? result : "")} tail{letter}").ToArray();
        foreach (string extension in new[] { "rtf", "docx", "doc" })
        {
            string name = "field-leading-switch-v1." + extension;
            string path = Path.Combine(root, "test-data", "word-ruby", name);
            string body = WordDocumentText.GetText(path);
            string firstDiff = string.Join(" | ", body.Split('\n').Zip(lines, (got, want) => (got, want)).Where(pair => pair.got != pair.want).Take(3).Select(pair => $"got \"{Ascii(pair.got)}\" want \"{Ascii(pair.want)}\""));
            TestReport.Check(body == string.Join("\n", lines), $"{name}: GetText = Word's printed text exactly {firstDiff}");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            CheckBlocks(name, document, body);
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, lines.Length);
            List<string> rendered = blocks.Select(block => System.Net.WebUtility.HtmlDecode(Regex.Replace(block.Html, "<[^>]+>", ""))).ToList();
            int eq = order.IndexOf('I');
            var renderWrong = Enumerable.Range(0, lines.Length).Where(i => i != eq && (i >= rendered.Count || rendered[i] != lines[i])).ToList();
            TestReport.Check(renderWrong.Count == 0, $"{name}: the rendered text of every case but the EQ is GetText's line{(renderWrong.Count > 0 && renderWrong[0] < rendered.Count ? $" (first: \"{Ascii(rendered[renderWrong[0]])}\")" : "")}");
            string eqHtml = eq < blocks.Count ? blocks[eq].Html : "";
            TestReport.Check(Regex.Matches(eqHtml, "data-note-number=\"(\\d+)\"").Cast<Match>().Select(match => match.Groups[1].Value).SequenceEqual(new[] { "2" }),
                $"{name}: the EQ draws its own note's mark (2) and not the nested field's (1), as Word's PDF");
            foreach (string hidden in new[] { "resA", "resB", "resC", "resD", "resF", "resG", "resJ", "resK", "resL", "MERGEFORMAT" })
                TestReport.Check(SumCounts(from => document.Search(hidden, WordDocumentSearchOptions.Default, from, 1000)) == 0, $"{name}: search finds no \"{hidden}\"");
        }
    }

    /// <summary>
    /// test-data/word-ruby/note-in-field-keyword-v1 (.rtf and Word's own .docx / .doc): a field whose keyword a note reference
    /// breaks has no type in Word and shows nothing; GetText (and so the search) reads exactly the text Word prints in its PDFs
    /// (note-in-field-keyword-v1.*.pdf) - no hidden result, no code.
    /// </summary>
    private static void CheckNoteInFieldKeyword(string root)
    {
        // the cases whose result Word prints (the PDFs of all three formats); every other case shows nothing between its label and
        // tail (case P, an EQ overstrike, is drawn but an EQ's drawing is not search text in any format - WordFieldEvaluator EQ)
        var shown = new Dictionary<char, string> { ['B'] = "resB", ['E'] = "resE", ['J'] = "resJ", ['M'] = "resM", ['R'] = "resR", ['S'] = "resS" };
        string expected = string.Join("\n", "ABCDEFGHIJKLMNOPQRS".Select(letter => $"case {letter}: {(shown.TryGetValue(letter, out string? result) ? result : "")} tail{letter}"));
        string[] expectedNotes = "ABCDFGHILNOPRS".Select(letter => $"note {letter}").ToArray();
        foreach (string extension in new[] { "rtf", "docx", "doc" })
        {
            string path = Path.Combine(root, "test-data", "word-ruby", "note-in-field-keyword-v1." + extension);
            string body = WordDocumentText.GetText(path);
            string firstDiff = string.Join(" | ", body.Split('\n').Zip(expected.Split('\n'), (got, want) => (got, want)).Where(pair => pair.got != pair.want).Take(3).Select(pair => $"got \"{Ascii(pair.got)}\" want \"{Ascii(pair.want)}\""));
            TestReport.Check(body == expected, $"note-in-field-keyword-v1.{extension}: GetText = Word's printed text exactly {firstDiff}");
            string[] notes = WordDocumentText.GetNotes(path).Split('\n').Select(line => line.Trim()).ToArray();
            TestReport.Check(notes.Length == expectedNotes.Length && notes.Zip(expectedNotes, (got, want) => got.EndsWith(want, StringComparison.Ordinal)).All(ok => ok),
                $"note-in-field-keyword-v1.{extension}: GetNotes reads the {expectedNotes.Length} notes Word numbers ({notes.Length})");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            CheckBlocks("note-in-field-keyword-v1." + extension, document, body);
            // what renders (tags dropped) is the same text, case by case - the EQ case P draws its overstrike and note mark
            string[] lines = expected.Split('\n');
            List<string> rendered = document.RenderBlocks(0, lines.Length).Select(block => System.Net.WebUtility.HtmlDecode(Regex.Replace(block.Html, "<[^>]+>", ""))).ToList();
            var renderWrong = Enumerable.Range(0, lines.Length).Where(i => i != 'P' - 'A' && (i >= rendered.Count || rendered[i] != lines[i])).ToList();
            TestReport.Check(renderWrong.Count == 0, $"note-in-field-keyword-v1.{extension}: the rendered text of every case but the EQ is GetText's line{(renderWrong.Count > 0 && renderWrong[0] < rendered.Count ? $" (first: \"{Ascii(rendered[renderWrong[0]])}\")" : "")}");
            foreach (string hidden in new[] { "resA", "resD", "resF", "resH", "resL", "MERGEFIELD", "name" })
                TestReport.Check(SumCounts(from => document.Search(hidden, WordDocumentSearchOptions.Default, from, 1000)) == 0, $"note-in-field-keyword-v1.{extension}: search finds no \"{hidden}\"");
        }
    }

    /// <summary>
    /// GetText on a document opened on demand (index still building, no sidecar yet): the scan must extend the index as far
    /// as each batch asks, or GetText never ends (RTF and plain text / web archives once returned no progress past the known blocks).
    /// </summary>
    private static void CheckWhileIndexing()
    {
        const int paragraphs = 5000;
        string folder = Path.Combine(TestReport.OutputDirectory, "text-while-indexing-v1");
        if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
        Directory.CreateDirectory(folder);
        var rtf = new System.Text.StringBuilder(@"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}").Append('\n');
        var txt = new System.Text.StringBuilder();
        for (int i = 0; i < paragraphs; i++) { rtf.Append(@"\pard line").Append(i).Append(@"\par").Append('\n'); txt.Append("line").Append(i).Append('\n'); }
        rtf.Append('}');
        string rtfPath = Path.Combine(folder, "many-paragraphs.rtf"), txtPath = Path.Combine(folder, "many-lines.txt");
        File.WriteAllText(rtfPath, rtf.ToString());
        File.WriteAllText(txtPath, txt.ToString());
        foreach (string path in new[] { rtfPath, txtPath })
        {
            var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true, CacheDirectory = Path.Combine(folder, "index-" + Path.GetExtension(path).TrimStart('.')) };
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            bool wasIndexing = !document.IsIndexComplete;
            using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
            string text = "";
            try { text = WordDocumentText.GetText(document, timeout.Token); }
            catch (OperationCanceledException) { }
            int lines = text.Split('\n').Count(line => line.StartsWith("line", StringComparison.Ordinal));
            TestReport.Check(wasIndexing && lines == paragraphs, $"{Path.GetFileName(path)}: GetText while the index builds reads all {paragraphs} paragraphs (opened indexing: {wasIndexing}, read {lines})");
        }
    }

    /// <summary>Up to 6 words of the text (longest first): the search's count over the document equals the count in the text.</summary>
    private static void CheckWords(string name, string what, string text, Func<string, int> searchCount)
    {
        string[] words = Regex.Matches(text, "\\p{L}{3,}").Cast<Match>().Select(match => match.Value).Distinct()
                              .OrderByDescending(word => word.Length).ThenBy(word => word, StringComparer.Ordinal).Take(6).ToArray();
        foreach (string word in words)
        {
            var matcher = new WordDocumentSearchMatcher(word);
            matcher.Feed(text.AsSpan());
            int found = searchCount(word);
            TestReport.Check(found == matcher.Count && found > 0, $"{name}: {what} search \"{Ascii(word)}\" {found} = text {matcher.Count}");
        }
    }

    /// <summary>Every block's text read by the search scan one block at a time (and by the golden reader) equals GetText's entry for it, and its words count the same in the search.</summary>
    private static void CheckBlocks(string name, IWordWindowedDocument document, string body)
    {
        var scan = (IWordDocumentSearchScan)document;
        List<(int Index, string Text)> blocks = CaptureEach(scan.ScanBlocks);
        TestReport.Check(string.Join("\n", blocks.Select(block => block.Text)) == body, $"{name}: GetText = the search's text of each block, read one at a time ({blocks.Count} blocks)");
        Func<WordDocumentSearchMatcher, int, int, WordDocumentSearchBatch>? golden = document switch
        {
            FastDocView.Docx.DocxWindowedDocument docx => (matcher, from, count) => docx.SearchWith(matcher, from, count, FastDocView.Docx.DocxWindowedDocument.DocxSearchTextMode.ReferenceXmlReader),
            FastDocView.Odt.OdtWindowedDocument odt => (matcher, from, count) => odt.SearchWith(matcher, from, count, FastDocView.Odt.OdtWindowedDocument.OdtSearchTextMode.ReferenceXmlReader),
            FastDocView.DocBinary.DocBinaryWindowedDocument doc => (matcher, from, count) => doc.SearchBlocksWith(matcher, from, count, onePass: false),
            _ => null,
        };
        if (golden is not null)
            TestReport.Check(CaptureEach(golden).SequenceEqual(blocks), $"{name}: the golden reference reader reads the same text per block");
        CheckEach(name, "block", blocks, (word, index) => document.Search(word, WordDocumentSearchOptions.Default, index, 1));
    }

    /// <summary>Every note's text read by the notes search one note at a time equals GetNotes' entry for it, and its words count the same in the notes search.</summary>
    private static void CheckNotes(string name, IWordWindowedDocument document, string notes)
    {
        List<(int Index, string Text)> each = CaptureEach((matcher, from, count) => WordDocumentNoteSearch.Search(document, matcher, from, count));
        TestReport.Check(string.Join("\n", each.Select(note => note.Text)) == notes, $"{name}: GetNotes = the notes search's text of each note, read one at a time ({each.Count} notes)");
        CheckEach(name, "note", each, (word, index) => document.SearchNotes(word, WordDocumentSearchOptions.Default, index, 1));
    }

    /// <summary>The scan's capture, one block / note per batch: (index, text) of each one with text.</summary>
    private static List<(int Index, string Text)> CaptureEach(Func<WordDocumentSearchMatcher, int, int, WordDocumentSearchBatch> scan)
    {
        var captured = new List<(int, string)>();
        int from = 0;
        while (true)
        {
            WordDocumentSearchMatcher matcher = WordDocumentSearchMatcher.ForTextCapture();
            WordDocumentSearchBatch batch = scan(matcher, from, 1);
            for (int i = 0; i < batch.HitBlocks.Length; i++)
            {
                string text = matcher.CapturedText(batch.HitCounts[i]);
                if (text.Length > 0) captured.Add((batch.HitBlocks[i], text));
            }
            if (batch.ReachedEnd || batch.NextBlock < 0) return captured;
            from = batch.NextBlock;
        }
    }

    /// <summary>For each block / note, every distinct word of its text: the search's count in that one block / note equals the count in the text.</summary>
    private static void CheckEach(string name, string what, List<(int Index, string Text)> entries, Func<string, int, WordDocumentSearchBatch> search)
    {
        int words = 0, wrong = 0;
        string firstWrong = "";
        foreach ((int index, string text) in entries)
            foreach (string word in Regex.Matches(text, "\\p{L}{2,}").Cast<Match>().Select(match => match.Value).Distinct())
            {
                var matcher = new WordDocumentSearchMatcher(word);
                matcher.Feed(text.AsSpan());
                WordDocumentSearchBatch batch = search(word, index);
                int found = batch.HitBlocks.Length == 1 && batch.HitBlocks[0] == index ? batch.HitCounts[0] : 0;
                words++;
                if (found != matcher.Count) { if (wrong++ == 0) firstWrong = $" (first: {what} {index} \"{Ascii(word)}\" search {found}, text {matcher.Count})"; }
            }
        TestReport.Check(wrong == 0, $"{name}: every word of every {what} counts the same in the search and the text ({words} checked, {wrong} differ){firstWrong}");
    }

    private static int SumCounts(Func<int, WordDocumentSearchBatch> step)
    {
        int total = 0, from = 0;
        while (true)
        {
            WordDocumentSearchBatch batch = step(from);
            total += batch.HitCounts.Sum();
            if (batch.ReachedEnd || batch.NextBlock < 0) return total;
            from = batch.NextBlock;
        }
    }

    private static string Ascii(string word) => word.All(c => c < 128) ? word : $"<{word.Length} letters>";

    private static string FindRepositoryRoot()
    {
        string? directory = AppContext.BaseDirectory;
        while (directory is not null && !Directory.Exists(Path.Combine(directory, "test-data"))) directory = Path.GetDirectoryName(directory);
        return directory ?? Directory.GetCurrentDirectory();
    }
}
