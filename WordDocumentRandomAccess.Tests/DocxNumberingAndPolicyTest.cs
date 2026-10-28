using WordDocumentRandomAccess.Docx;
using WordDocumentRandomAccess.DocxHtml;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Index-time list numbering (labels must be the running count, not restart per window) and the
/// index policy (small file → in-memory with no sidecar and no checkpoints; forced sidecar → file).
/// </summary>
public static class DocxNumberingAndPolicyTest
{
    public static void Run()
    {
        TestReport.Step("Number formats");
        // Word 16 ListFormat.ListString (2026-10-28): no geresh / gershayim; the sequence cycles after 392.
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(1) == "א", "hebrew1 1 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(11) == "יא", "hebrew1 11 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(15) == "טו", "hebrew1 15 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(16) == "טז", "hebrew1 16 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(115) == "קטו", "hebrew1 115 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(500) == "קח", "hebrew1 500 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(1000) == "רטז", "hebrew1 1000 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewNumeral(5784) == "רצו", "hebrew1 5784 as Word shows it");
        TestReport.Check(DocxNumberingDefinitions.HebrewLetters(23) == "אא", "hebrew2 23 → אא");
        TestReport.Check(DocxNumberingDefinitions.FormatNumber(4, "lowerRoman") == "iv", "lowerRoman 4 → iv");
        TestReport.Check(DocxNumberingDefinitions.FormatNumber(28, "upperLetter") == "BB", "upperLetter 28 → BB");

        TestReport.Step("Small synthetic docx: policy, in-memory mode, running list numbers");
        string path = Path.Combine(TestReport.OutputDirectory, "synthetic-300.docx");
        string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache");
        SyntheticDocxGenerator.Write(path, 300);
        var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory };
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);

        WordIndexDecision decision = WordWindowedDocumentOpener.Decide(path, policy);
        TestReport.Check(decision.Mode == WordIndexMode.InMemory, $"policy: {decision.Mode} — {decision.Reason}");

        using (var document = (DocxWindowedDocument)TestReport.Timed("open in-memory", () => WordWindowedDocumentOpener.Open(path, policy)))
        {
            TestReport.Check(document.Decision.Mode == WordIndexMode.InMemory, "document opened in-memory");
            TestReport.Check(!File.Exists(indexPath), "no sidecar written for a small file");
            TestReport.Check(document.Index.Checkpoints.Count == 0, "no checkpoints captured in-memory");
            TestReport.Check(document.BlockCount == SyntheticDocxGenerator.ExpectedBlockCount(300), $"block count {document.BlockCount}");

            // paragraphs 3, 13, 23 … are decimal list items 1., 2., 3. …; 4, 14 … are a., b. (level 1 restarts after each level-0 item → always "1.a" style)
            string block3 = document.RenderBlocks(3, 1)[0].Html, block13 = document.RenderBlocks(13, 1)[0].Html, block123 = document.RenderBlocks(123, 1)[0].Html;
            TestReport.Check(block3.Contains("<span class=\"list-marker\">1.</span>"), "block 3 is list item 1.");
            TestReport.Check(block13.Contains("<span class=\"list-marker\">2.</span>"), "block 13 is list item 2.");
            TestReport.Check(block123.Contains("<span class=\"list-marker\">13.</span>"), $"block 123 is list item 13. (running count, not per window)");
            string block14 = document.RenderBlocks(14, 1)[0].Html;
            TestReport.Check(block14.Contains("<span class=\"list-marker\">2.a</span>"), "block 14 is sub-item 2.a");
            string block7 = document.RenderBlocks(7, 1)[0].Html, block97 = document.RenderBlocks(97, 1)[0].Html;
            TestReport.Check(block7.Contains("<span class=\"list-marker\">א.</span>"), "block 7 is Hebrew item א.");
            TestReport.Check(block97.Contains("<span class=\"list-marker\">י.</span>"), "block 97 is Hebrew item י.");
            TestReport.Check(!document.RenderBlocks(5, 1)[0].Html.Contains("list-marker"), "plain paragraph has no marker");

            // notes: paragraph 21 references footnote 2, paragraph 102 endnote ii (Word's endnote default is lowerRoman)
            TestReport.Check(document.NoteCounts == (15, 3), $"note counts {document.NoteCounts} == (15, 3)");
            WordDocumentBlockHtml withFootnote = document.RenderBlocks(21, 1)[0];
            TestReport.Check(withFootnote.Html.Contains("data-note-kind=\"footnote\" data-note-number=\"2\""), "block 21 references footnote 2");
            TestReport.Check(withFootnote.Notes.Count == 1 && withFootnote.Notes[0].Label == "2" && withFootnote.Notes[0].Html.Contains("של פסקה 21")
                             && withFootnote.Notes[0].Html.Contains("note-mark\">2<"), "footnote 2 body comes with block 21, numbered 2");
            WordDocumentBlockHtml withEndnote = document.RenderBlocks(102, 1)[0];
            TestReport.Check(withEndnote.Notes.Count == 1 && withEndnote.Notes[0].Kind == WordDocumentNoteKind.Endnote && withEndnote.Notes[0].Label == "ii"
                             && withEndnote.Notes[0].Html.Contains("של פסקה 102"), "endnote ii body comes with block 102");
            TestReport.Check(document.RenderBlocks(0, 300).Sum(b => b.Notes.Count) == 18, "every reference in the document delivers its note");
        }

        var forced = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
        using (var document = (DocxWindowedDocument)TestReport.Timed("open forced sidecar", () => WordWindowedDocumentOpener.Open(path, forced)))
        {
            TestReport.Check(document.Decision.Mode == WordIndexMode.Sidecar, "forced sidecar mode");
            TestReport.Check(File.Exists(indexPath), "sidecar written");
            TestReport.Check(document.Index.Checkpoints.Count >= 1, $"{document.Index.Checkpoints.Count} checkpoint(s)");
            TestReport.Check(document.RenderBlocks(123, 1)[0].Html.Contains("list-marker\">13.<"), "labels persist in the sidecar");
        }
        WordIndexDecision cachedDecision = WordWindowedDocumentOpener.Decide(path, policy);
        TestReport.Check(cachedDecision is { Mode: WordIndexMode.Sidecar, HasCachedIndex: true }, "a cached sidecar is used even below the threshold");

        TestReport.Step("Folder decisions");
        List<WordIndexDecision> folder = WordWindowedDocumentOpener.DecideFolder(TestReport.OutputDirectory, policy, recursive: false).ToList();
        TestReport.Check(folder.Any(d => d.FilePath == path && d.HasCachedIndex), $"{folder.Count} Word file(s) found in TestOutput, the synthetic one shows its cached index");
        foreach (WordIndexDecision d in folder) Console.WriteLine($"      {Path.GetFileName(d.FilePath),-28} {d.Mode,-9} cached={d.HasCachedIndex,-5} body={d.BodyBytes >> 10} KB  {d.Reason}");
    }
}
