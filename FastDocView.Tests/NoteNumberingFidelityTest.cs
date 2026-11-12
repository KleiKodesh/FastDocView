using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Footnote and endnote numbers as Word prints them: test-data/word-notes/note-numbering-v1.docx (make-note-numbering-docx.py:
/// five sections with their own number format, start and restart) and the .doc / .rtf Word 16 saved from it; the reference is
/// Word's own PDF (note-numbering-v1-word.pdf). Section 4 restarts each page: Word prints 1, 1 (two pages); a flow view has no
/// pages and counts it like each section (1, 2) - see WordDocumentNoteNumberRule. Also an .fodt with
/// text:start-numbering-at="chapter" and no citation text (the fallback label).
/// </summary>
internal static class NoteNumberingFidelityTest
{
    // body order; Word's PDF: 1 2 b c 14 15 1 1 IX / i i iii D v
    private static readonly string[] Footnotes = ["1", "2", "b", "c", "14", "15", "1", "2", "IX"];
    private static readonly string[] Endnotes = ["i", "i", "iii", "D", "v"];

    public static void Run()
    {
        foreach (string extension in new[] { ".docx", ".doc", ".rtf" })
        {
            string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-notes", "note-numbering-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-notes/note-numbering-v1{extension} present"); continue; }
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-note-numbering-" + extension.TrimStart('.')) };
            for (int pass = 0; pass < 2; pass++)          // first open, then the reopen from the saved index
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                while (!document.IsIndexComplete) document.ScanAhead(1000);
                IReadOnlyList<WordDocumentNoteHtml> notes = document.RenderNotes(0, document.NoteReferences.Count);
                string footnotes = string.Join(" ", notes.Where(note => note.Kind == WordDocumentNoteKind.Footnote).Select(note => note.Label));
                string endnotes = string.Join(" ", notes.Where(note => note.Kind == WordDocumentNoteKind.Endnote).Select(note => note.Label));
                string label = pass == 0 ? "first open" : "reopen";
                TestReport.Check(footnotes == string.Join(" ", Footnotes), $"{extension} ({label}): footnote numbers as Word prints them (got [{footnotes}])");
                TestReport.Check(endnotes == string.Join(" ", Endnotes), $"{extension} ({label}): endnote numbers as Word prints them (got [{endnotes}])");
                // the mark in the body text shows the same label
                string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
                TestReport.Check(html.Contains(">14<") && html.Contains(">IX<") && html.Contains(">D<"), $"{extension} ({label}): body marks carry the section's numbers");
            }
        }
        RunOdtChapterRestart();
    }

    private static void RunOdtChapterRestart()
    {
        const string Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0", Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0", Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        string note(string id, string body) => $"<text:note text:id=\"{id}\" text:note-class=\"footnote\"><text:note-citation/><text:note-body><text:p>{body}</text:p></text:note-body></text:note>";
        string xml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <office:document xmlns:office="{Office}" xmlns:text="{Text}" xmlns:style="{Style}" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
            <office:styles><text:notes-configuration text:note-class="footnote" style:num-format="a" text:start-value="1" text:start-numbering-at="chapter"/></office:styles>
            <office:body><office:text>
            <text:h text:outline-level="1">One</text:h><text:p>a{note("n1", "x")}</text:p><text:p>b{note("n2", "y")}</text:p>
            <text:h text:outline-level="2">One.one</text:h><text:p>c{note("n3", "z")}</text:p>
            <text:h text:outline-level="1">Two</text:h><text:p>d{note("n4", "w")}</text:p>
            </office:text></office:body></office:document>
            """;
        string path = Path.Combine(TestReport.OutputDirectory, "note-numbering-chapter-v1.fodt");
        File.WriteAllText(path, xml);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-note-numbering-fodt") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        while (!document.IsIndexComplete) document.ScanAhead(1000);
        string labels = string.Join(" ", document.RenderNotes(0, document.NoteReferences.Count).Select(note => note.Label));
        // start-value 1 is LibreOffice's 0-based offset (XMLFootnoteConfigurationImportContext.cxx:175-181): the first note is "b"
        TestReport.Check(labels == "b c d b", $".fodt: footnotes restart at each outline-level-1 heading (start-numbering-at=chapter) (got [{labels}])");
    }
}
