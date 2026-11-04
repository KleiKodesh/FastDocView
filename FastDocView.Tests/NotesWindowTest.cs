using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// RenderNotes (the notes pane's own windows over every note) against the notes the blocks deliver: same notes, same order,
/// same HTML, for every reader - synthetic .docx / .odt / .rtf plus any real files given.
/// </summary>
public static class NotesWindowTest
{
    public static void Run(IEnumerable<string> realPaths)
    {
        string directory = Path.Combine(TestReport.OutputDirectory, "notes-window");
        Directory.CreateDirectory(directory);
        string docx = Path.Combine(directory, "notes.docx"), odt = Path.Combine(directory, "notes.odt"), rtf = Path.Combine(directory, "notes.rtf");
        string fodt = Path.Combine(directory, "notes.fodt");
        if (!File.Exists(docx)) SyntheticDocxGenerator.Write(docx, 600);
        if (!File.Exists(odt)) SyntheticOdtGenerator.WritePackage(odt, 600);
        if (!File.Exists(rtf)) SyntheticRtfGenerator.Write(rtf, 600);
        if (!File.Exists(fodt)) SyntheticOdtGenerator.WriteFlat(fodt, 600);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache") };
        foreach (string path in new[] { docx, odt, rtf, fodt }.Concat(realPaths))
        {
            TestReport.Step($"Notes windows: {Path.GetFileName(path)}");
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            while (!document.IsIndexComplete) document.ScanAhead(100_000);
            var expected = new List<WordDocumentNoteHtml>();
            for (int block = 0; block < document.BlockCount; block += 200) foreach (WordDocumentBlockHtml rendered in document.RenderBlocks(block, 200)) expected.AddRange(rendered.Notes);
            int noteCount = document.NoteReferences.Count;
            bool isSynthetic = path == docx || path == odt || path == rtf || path == fodt;              // generated with notes; a real file may have none
            TestReport.Check((noteCount > 0 || !isSynthetic) && expected.Count == noteCount, $"{noteCount} references, the blocks deliver {expected.Count} notes");
            var windowed = new List<WordDocumentNoteHtml>();
            for (int first = 0; first < noteCount; first += 7) windowed.AddRange(document.RenderNotes(first, 7));
            TestReport.Check(windowed.Count == noteCount, $"windows of 7 deliver {windowed.Count} notes");
            int mismatches = Enumerable.Range(0, Math.Min(windowed.Count, expected.Count))
                .Count(i => windowed[i].Kind != expected[i].Kind || windowed[i].Number != expected[i].Number || windowed[i].Label != expected[i].Label || windowed[i].Html != expected[i].Html);
            TestReport.Check(mismatches == 0, $"every windowed note equals the block's note ({mismatches} differ)");
            TestReport.Check(Enumerable.Range(0, noteCount).All(i => document.NoteReferences[i].Number == windowed[i].Number && document.NoteReferences[i].Kind == windowed[i].Kind),
                             "note i of the windows is reference i");
            TestReport.Check(document.RenderNotes(noteCount, 5).Count == 0 && document.RenderNotes(-3, 2).Count == Math.Min(2, noteCount), "windows clamp to the notes");
        }
    }
}
