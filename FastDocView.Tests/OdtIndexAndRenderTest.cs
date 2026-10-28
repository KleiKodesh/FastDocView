using System.Diagnostics;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// Synthetic OpenDocument text, package (.odt) and flat (.fodt) from the same generator: format
/// detection, block count, list labels (nested, display-levels, start-value, list-header,
/// continue-numbering, continue-list, bullet), headings, notes numbered and delivered with their
/// block (and excluded from the text length), the table, the index saved and reopened, random windows.
/// </summary>
public static class OdtIndexAndRenderTest
{
    public static void Run(int count)
    {
        string directory = TestReport.OutputDirectory;
        string packagePath = Path.Combine(directory, $"synthetic-{count}-v2.odt");
        string flatPath = Path.Combine(directory, $"synthetic-{count}-v2.fodt");
        TestReport.Step($"Synthetic OpenDocument with {SyntheticOdtGenerator.ExpectedBlockCount(count):N0} blocks");
        if (!File.Exists(packagePath)) TestReport.Timed("generate .odt", () => SyntheticOdtGenerator.WritePackage(packagePath, count));
        if (!File.Exists(flatPath)) TestReport.Timed("generate .fodt", () => SyntheticOdtGenerator.WriteFlat(flatPath, count));
        Console.WriteLine($"   .odt {new FileInfo(packagePath).Length >> 10} KB, .fodt {new FileInfo(flatPath).Length >> 10} KB");

        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(packagePath) == WordWindowedDocumentOpener.WordFileFormat.OpenDocument, ".odt detected by its mimetype entry");
        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(flatPath) == WordWindowedDocumentOpener.WordFileFormat.FlatOpenDocument, ".fodt detected by its office:document root");

        // chunk edges everywhere: the flat file fed to the scanner 1-7 bytes per read
        using (var trickle = new TrickleStream(File.ReadAllBytes(flatPath)))
        using (OdtWindowedDocument trickled = OdtWindowedDocument.OpenFromStream(trickle, flatPath, isFlat: true,
                   new WordIndexDecision(flatPath, WordIndexMode.InMemory, false, trickle.Length, 0, "test"), null, 0, 0))
            Verify(trickled, count, ".fodt fed in 1-7 byte reads");

        foreach ((string path, string what) in new[] { (packagePath, ".odt"), (flatPath, ".fodt") })
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            Console.WriteLine($"   — {what}, {mode}");
            string cacheDirectory = Path.Combine(directory, "index-cache");
            var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = mode };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (var document = (OdtWindowedDocument)TestReport.Timed("open (cold)", () => WordWindowedDocumentOpener.Open(path, policy)))
                Verify(document, count, what);
            if (mode != WordIndexMode.Sidecar) continue;
            TestReport.Check(File.Exists(indexPath), "sidecar index written");
            using var reopened = (OdtWindowedDocument)TestReport.Timed("re-open from cached index", () => WordWindowedDocumentOpener.Open(path, policy));
            TestReport.Check(reopened.Decision.HasCachedIndex, "cached index used");
            Verify(reopened, count, what + " reopened");
        }
    }

    private static void Verify(OdtWindowedDocument document, int count, string what)
    {
        OdtRandomAccessIndex index = document.Index;
        TestReport.Check(document.BlockCount == SyntheticOdtGenerator.ExpectedBlockCount(count), $"{what}: block count {document.BlockCount:N0} == {SyntheticOdtGenerator.ExpectedBlockCount(count):N0}");
        bool labelsMatch = true;
        foreach ((int block, string? label) in SyntheticOdtGenerator.ExpectedPrefixLabels)
            if (index.ListLabelOf(block) != label) { labelsMatch = false; Console.WriteLine($"      block {block}: label '{index.ListLabelOf(block)}', expected '{label}'"); }
        TestReport.Check(labelsMatch, $"{what}: list labels (nested 2.a), start-value 7, header unnumbered, continue-numbering 8, new list 1, continue-list 9, bullet)");
        TestReport.Check(index.ListLevelOf(4) == 1 && index.ListLevelOf(2) == 0 && index.ListLevelOf(7) == -1, $"{what}: list depths");
        TestReport.Check(index.BlockKinds[11] == WordDocumentBlockKind.Table, $"{what}: block 11 is the table");
        TestReport.Check(document.IsRightToLeft, $"{what}: right-to-left document");
        TestReport.Check(document.Headings.Count == SyntheticOdtGenerator.ExpectedHeadings(count) && document.Headings[0].Level == 1 && document.Headings[1].Level == 2
                         && document.Headings[0].Text.EndsWith(" 1", StringComparison.Ordinal),
                         $"{what}: {document.Headings.Count} headings (expected {SyntheticOdtGenerator.ExpectedHeadings(count)}), levels and text");
        (int footnotes, int endnotes) = document.NoteCounts;
        TestReport.Check(footnotes == SyntheticOdtGenerator.ExpectedFootnotes(count) && endnotes == SyntheticOdtGenerator.ExpectedEndnotes,
                         $"{what}: {footnotes} footnotes / {endnotes} endnotes counted");
        TestReport.Check(document.NoteReferences.Where(r => r.Kind == WordDocumentNoteKind.Footnote).Select(r => r.Number).SequenceEqual(Enumerable.Range(1, footnotes)),
                         $"{what}: footnotes numbered 1..{footnotes} in body order");

        IReadOnlyList<WordDocumentBlockHtml> prefix = document.RenderBlocks(0, SyntheticOdtGenerator.PrefixBlockCount);
        WordDocumentBlockHtml withNotes = prefix[1];
        TestReport.Check(withNotes.Notes.Count == 2 && withNotes.Notes[0].Kind == WordDocumentNoteKind.Footnote && withNotes.Notes[0].Number == 1
                         && withNotes.Notes[0].Html.Contains("footnote body 1") && withNotes.Notes[1].Kind == WordDocumentNoteKind.Endnote && withNotes.Notes[1].Label == "i",
                         $"{what}: block 1 delivers its footnote and endnote bodies, numbered");
        TestReport.Check(withNotes.Html.Contains("data-note-kind=\"footnote\" data-note-number=\"1\"") && !withNotes.Html.Contains("footnote body"),
                         $"{what}: note reference inline, body not inline");
        TestReport.Check(document.EstimatedTextLength(1) < 60, $"{what}: note bodies excluded from text length ({document.EstimatedTextLength(1)})");
        TestReport.Check(withNotes.Html.Contains("<span style=\"") && withNotes.Html.Contains("font-style:italic") && withNotes.Html.Contains("  "), $"{what}: span style (italic via parent text style) and text:s spaces");
        TestReport.Check(prefix[0].Html.StartsWith("<h1", StringComparison.Ordinal) && prefix[0].Html.Contains("dir=\"rtl\"") && prefix[0].Html.Contains("font-weight:bold"), $"{what}: heading renders <h1 dir=rtl> bold");
        TestReport.Check(prefix[4].Html.Contains("list-marker\">2.a)"), $"{what}: nested list item renders its label");
        string table = prefix[11].Html;
        TestReport.Check(table.Contains("colspan=\"2\"") && table.Contains("width:144pt") && table.Contains("background-color:#EEEEEE") && table.Contains("vertical-align:middle")
                         && System.Text.RegularExpressions.Regex.Matches(table, "<td").Count == 5 && prefix[11].Notes.Count == 1,
                         $"{what}: table spans, covered cell skipped, column width, cell style, footnote in a cell");

        var random = new Random(7);
        double worst = 0;
        bool ordinals = true;
        for (int i = 0; i < 30; i++)
        {
            int first = SyntheticOdtGenerator.PrefixBlockCount + random.Next(Math.Max(1, count - 40));
            var watch = Stopwatch.StartNew();
            IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(first, 40);
            worst = Math.Max(worst, watch.Elapsed.TotalMilliseconds);
            foreach (WordDocumentBlockHtml block in window)
            {
                int ordinal = block.BlockIndex - SyntheticOdtGenerator.PrefixBlockCount;
                if (ordinal % 50 == 0 && ordinal % 200 != 199) continue;          // headings carry "section N"
                if (!block.Html.Contains(" " + block.BlockIndex)) { ordinals = false; Console.WriteLine($"      block {block.BlockIndex} lacks its ordinal"); }
                if (ordinal % 10 == 5 && ordinal % 200 != 199 && block.Notes.Count != 1) { ordinals = false; Console.WriteLine($"      block {block.BlockIndex} lacks its footnote"); }
            }
        }
        Console.WriteLine($"   30 random windows of 40 blocks: worst {worst:0.#} ms");
        TestReport.Check(ordinals, $"{what}: every rendered block holds its own ordinal and its notes");
        TestReport.Check(worst < 500, $"{what}: worst random window under 500 ms");
        File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "synthetic-odt-sample.html"),
            $"<!doctype html><meta charset=utf-8><style>{document.Stylesheet}</style><div class=word-document dir=rtl>{string.Concat(prefix.Select(b => b.Html + string.Concat(b.Notes.Select(n => "<div style='color:#666'>" + n.Html + "</div>"))))}</div>");
    }
}

/// <summary>A memory stream whose reads return 1-7 bytes, so every tag, comment and capture crosses chunk edges.</summary>
internal sealed class TrickleStream(byte[] bytes) : MemoryStream(bytes, writable: false)
{
    private int _next;
    public override int Read(byte[] buffer, int offset, int count) => base.Read(buffer, offset, Math.Min(count, _next++ % 7 + 1));
}
