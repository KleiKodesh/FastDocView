using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Notes are searched apart from the main text, on real files with footnotes / endnotes (.doc from test-data, and any
/// given paths): (1) a block's search count equals what its own rendered HTML holds - never its notes'; (2)
/// <see cref="IWordWindowedDocument.SearchNotes"/> counts per note equal what each note's displayed HTML holds (walked as
/// the app's highlighter walks it), for a word taken from a note; (3) that word's note is found.
/// </summary>
public static class NoteSearchTest
{
    private static readonly string[] SampleFiles =
    [
        "test-data/apache-poi/53379.doc", "test-data/apache-poi/58804.doc", "test-data/apache-poi/Bug33519.doc",
        "test-data/apache-poi/Bug47958.doc", "test-data/apache-poi/Bug53380_1.doc", "test-data/apache-poi/GaiaTest.doc",
        "test-data/apache-poi/Bug61268.doc", "test-data/apache-poi/ob_is.doc", "test-data/apache-poi/Bug46220.doc",
        "test-data/apache-poi/endingnote.doc", "test-data/apache-poi/footnote.doc", "test-data/apache-poi/endnotes.docx",
        "test-data/apache-poi/footnotes.docx", "test-data/apache-poi/table_footnotes.docx",
    ];

    private const int MaxBlocks = 3000, MaxNotes = 500;

    public static void Run(IEnumerable<string> extraPaths)
    {
        TestReport.Step("Search: notes apart from the main text (real files with notes)");
        string root = FindRepositoryRoot();
        int checkedFiles = 0;
        foreach (string path in SampleFiles.Select(file => Path.Combine(root, file)).Concat(extraPaths).Where(File.Exists))
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
            if (document.NoteReferences.Count == 0) continue;
            checkedFiles++;
            Check(document, Path.GetFileName(path));
        }
        TestReport.Check(checkedFiles > 0, $"notes apart: {checkedFiles} files with notes checked");
    }

    private static void Check(IWordWindowedDocument document, string name)
    {
        int noteCount = Math.Min(MaxNotes, document.NoteReferences.Count);
        IReadOnlyList<WordDocumentNoteHtml> notes = document.RenderNotesForDisplay(0, noteCount);
        string? word = notes.Select(note => LongestWord(note.Html)).FirstOrDefault(found => found is not null);
        if (word is null) { Console.WriteLine($"   {name}: no word in its notes, skipped"); return; }

        // notes: per-note counts equal the displayed html's
        var expected = new List<(int note, int count)>();
        var matcher = new WordDocumentSearchMatcher(word);
        for (int i = 0; i < notes.Count; i++)
        {
            matcher.Reset();
            matcher.FeedHtml(notes[i].Html, separatesParagraphs: true);
            if (matcher.Count > 0) expected.Add((i, matcher.Count));
        }
        WordDocumentSearchBatch notesBatch = document.SearchNotes(word, WordDocumentSearchOptions.Default, 0, noteCount);
        var found = notesBatch.HitBlocks.Select((note, i) => (note, notesBatch.HitCounts[i])).ToList();
        TestReport.Check(found.SequenceEqual(expected) && found.Count > 0,
            $"{name}: notes search \"{word}\" = the notes' html ({found.Count} notes, {found.Sum(hit => hit.Item2)} matches; expected {expected.Count})");
        (int[] stepped, _) = DocumentSearchTest.AllNotes(document, word, batch: 7);
        TestReport.Check(stepped.Where(note => note < noteCount).SequenceEqual(expected.Select(hit => hit.note)), $"{name}: notes search in batches of 7 gives the same notes");

        // blocks: counts equal the block's own html, never its notes
        int blockCount = Math.Min(MaxBlocks, document.BlockCount);
        WordDocumentSearchBatch blocks = document.Search(word, WordDocumentSearchOptions.Default, 0, blockCount);
        var expectedBlocks = new List<(int block, int count)>();
        for (int first = 0; first < blockCount; first += 200)
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, Math.Min(200, blockCount - first)))
            {
                matcher.Reset();
                matcher.FeedHtml(block.Html);
                if (matcher.Count > 0) expectedBlocks.Add((block.BlockIndex, matcher.Count));
            }
        var foundBlocks = blocks.HitBlocks.Select((block, i) => (block, blocks.HitCounts[i])).ToList();
        TestReport.Check(foundBlocks.SequenceEqual(expectedBlocks),
            $"{name}: block search \"{word}\" counts only the blocks' own html ({foundBlocks.Count} blocks vs {expectedBlocks.Count} rendered)");
    }

    /// <summary>The longest run of letters (3+) in the html's text, or null.</summary>
    private static string? LongestWord(string html)
    {
        string text = Regex.Replace(Regex.Replace(html, "<[^>]*>", " "), "&[^;\\s<]{1,11};", " ");
        return Regex.Matches(text, "\\p{L}{3,}").Cast<Match>().Select(match => match.Value).OrderByDescending(value => value.Length).FirstOrDefault();
    }

    private static string FindRepositoryRoot()
    {
        string? directory = AppContext.BaseDirectory;
        while (directory is not null && !Directory.Exists(Path.Combine(directory, "test-data"))) directory = Path.GetDirectoryName(directory);
        return directory ?? Directory.GetCurrentDirectory();
    }
}
