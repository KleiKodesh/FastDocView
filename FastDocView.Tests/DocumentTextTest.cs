using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// <see cref="WordDocumentText"/> (GetText / GetNotes / GetTextWithNotes), one real file per format: the text equals what
/// the search indexes - for words taken from the text, the whole-document search count (<see cref="IWordWindowedDocument.Search"/>,
/// <see cref="IWordWindowedDocument.SearchNotes"/>) equals the count in the returned text; GetTextWithNotes is the body,
/// a blank line and the notes; a cancelled token stops it.
/// </summary>
public static class DocumentTextTest
{
    private static readonly string[] SampleFiles =
    [
        "test-data/apache-poi/footnote.doc", "test-data/apache-poi/footnotes.docx", "test-data/apache-poi/endnotes.docx",
        "test-data/apache-tika/testRTF_annotation_spacing.rtf", "test-data/libreoffice/sw/qa/core/layout/data/inline-endnote-and-section.odt",
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
