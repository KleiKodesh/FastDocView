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
