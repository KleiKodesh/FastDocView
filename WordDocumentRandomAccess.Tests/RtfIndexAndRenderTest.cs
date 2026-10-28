using System.Diagnostics;
using System.Text.RegularExpressions;
using WordDocumentRandomAccess.Rtf;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Synthetic RTF: detection, block count, headings, list labels (counted and \listtext), notes numbered and
/// delivered with their block, table rows (cell widths, shading, footnote in a cell, nested table), Hebrew from
/// \'xx through the font's \fcharset177, \uN with \ucN skipping, \bin data with braces, pictures, \upr/\ud,
/// fields, character formatting; the index fed in 1-7 byte reads, in memory and as a sidecar reopened; random windows.
/// </summary>
public static class RtfIndexAndRenderTest
{
    private static readonly string Hebrew = new([(char)0x05D0, (char)0x05D1, (char)0x05D2]);
    private static readonly string HebrewFour = new([(char)0x05D0, (char)0x05D1, (char)0x05D2, (char)0x05D3]);

    public static void Run(int count)
    {
        string directory = TestReport.OutputDirectory;
        string path = Path.Combine(directory, $"synthetic-{count}-v1.rtf");
        TestReport.Step($"Synthetic RTF with {SyntheticRtfGenerator.ExpectedBlockCount(count):N0} blocks");
        if (!File.Exists(path)) TestReport.Timed("generate .rtf", () => SyntheticRtfGenerator.Write(path, count));
        Console.WriteLine($"   .rtf {new FileInfo(path).Length >> 10} KB");
        TestReport.Check(WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.RichText, ".rtf detected by its {\\rtf signature");

        string[] reference;
        using (var trickle = new TrickleStream(File.ReadAllBytes(path)))
        using (RtfWindowedDocument trickled = RtfWindowedDocument.OpenFromStream(trickle, path,
                   new WordIndexDecision(path, WordIndexMode.InMemory, false, trickle.Length, 0, "test"), null, 0, 0))
            reference = Verify(trickled, count, "fed in 1-7 byte reads", null);

        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            Console.WriteLine($"   - {mode}");
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (var document = (RtfWindowedDocument)TestReport.Timed("open (cold)", () => WordWindowedDocumentOpener.Open(path, policy)))
                Verify(document, count, mode.ToString(), reference);
            if (mode != WordIndexMode.Sidecar) continue;
            TestReport.Check(File.Exists(indexPath), "sidecar index written");
            using var reopened = (RtfWindowedDocument)TestReport.Timed("re-open from cached index", () => WordWindowedDocumentOpener.Open(path, policy));
            TestReport.Check(reopened.Decision.HasCachedIndex, "cached index used");
            Verify(reopened, count, "sidecar reopened", reference);
        }
    }

    /// <returns>The HTML of the prefix blocks, compared across modes.</returns>
    private static string[] Verify(RtfWindowedDocument document, int count, string what, string[]? reference)
    {
        RtfRandomAccessIndex index = document.Index;
        TestReport.Check(document.BlockCount == SyntheticRtfGenerator.ExpectedBlockCount(count),
                         $"{what}: block count {document.BlockCount:N0} == {SyntheticRtfGenerator.ExpectedBlockCount(count):N0}");
        TestReport.Check(document.IsIndexComplete && document.KnownBlockCount == document.BlockCount, $"{what}: index complete");
        bool labelsMatch = true;
        foreach ((int block, string? label) in SyntheticRtfGenerator.ExpectedPrefixLabels)
            if (index.ListLabelOf(block) != label) { labelsMatch = false; Console.WriteLine($"      block {block}: label '{index.ListLabelOf(block)}', expected '{label}'"); }
        TestReport.Check(labelsMatch, $"{what}: list labels counted from the list table (1., 1.1, bullet), none where \\listtext gives it");
        TestReport.Check(index.BlockKinds[6] == WordDocumentBlockKind.TableRow && index.BlockKinds[7] == WordDocumentBlockKind.TableRow
                         && index.BlockKinds[14] == WordDocumentBlockKind.TableRow && index.BlockKinds[5] == WordDocumentBlockKind.Paragraph,
                         $"{what}: table rows are blocks of their own");
        TestReport.Check(document.Headings.Count == SyntheticRtfGenerator.ExpectedHeadings(count) && document.Headings[0].Level == 1
                         && document.Headings[0].Text.Contains(Hebrew) && document.Headings[1].Level == 2 && document.Headings[1].Text == "Heading two"
                         && document.Headings[2].Text == $"Section {SyntheticRtfGenerator.PrefixBlockCount}",
                         $"{what}: {document.Headings.Count} headings (expected {SyntheticRtfGenerator.ExpectedHeadings(count)}): \\outlinelevel and the \"heading 2\" style, Hebrew text");
        (int footnotes, int endnotes) = document.NoteCounts;
        TestReport.Check(footnotes == SyntheticRtfGenerator.ExpectedFootnotes(count) && endnotes == SyntheticRtfGenerator.ExpectedEndnotes,
                         $"{what}: {footnotes} footnotes / {endnotes} endnotes counted (expected {SyntheticRtfGenerator.ExpectedFootnotes(count)} / 1)");
        TestReport.Check(document.NoteReferences.Where(r => r.Kind == WordDocumentNoteKind.Footnote).Select(r => r.Number).SequenceEqual(Enumerable.Range(1, footnotes)),
                         $"{what}: footnotes numbered 1..{footnotes} in body order");

        IReadOnlyList<WordDocumentBlockHtml> prefix = document.RenderBlocks(0, SyntheticRtfGenerator.PrefixBlockCount);
        string Html(int i) => prefix[i].Html;
        TestReport.Check(Html(0).StartsWith("<h1", StringComparison.Ordinal) && Html(0).Contains("dir=\"rtl\"") && Html(0).Contains(Hebrew)
                         && Html(0).Contains("font-weight:bold") && Html(0).Contains("David"),
                         $"{what}: heading <h1 dir=rtl>, bold, Hebrew decoded from \\'xx via \\fcharset177, font David");
        WordDocumentBlockHtml withNotes = prefix[1];
        TestReport.Check(withNotes.Notes.Count == 2 && withNotes.Notes[0].Kind == WordDocumentNoteKind.Footnote && withNotes.Notes[0].Number == 1
                         && withNotes.Notes[0].Html.Contains("footnote body 1") && withNotes.Notes[0].Html.Contains("note-mark")
                         && withNotes.Notes[1].Kind == WordDocumentNoteKind.Endnote && withNotes.Notes[1].Label == "i",
                         $"{what}: block 1 delivers its footnote and endnote (\\ftnalt, \\aftnnrlc -> i) bodies");
        TestReport.Check(withNotes.Html.Contains("data-note-kind=\"footnote\" data-note-number=\"1\"") && !withNotes.Html.Contains("body 1"),
                         $"{what}: note reference inline, body not inline");
        TestReport.Check(document.EstimatedTextLength(1) < 20, $"{what}: note bodies excluded from text length ({document.EstimatedTextLength(1)})");
        TestReport.Check(Html(2).Contains("list-marker\">1.<") && Html(3).Contains("list-marker\">1.1<") && Html(4).Contains("list-marker\">2.<")
                         && Html(5).Contains("list-marker\">•<"), $"{what}: list markers (counted, \\listtext, bullet)");
        TestReport.Check(Html(6).Contains("<table") && Regex.Matches(Html(6), "<td").Count == 2 && Html(6).Contains("width:144pt") && Html(6).Contains("background:#EEEEEE")
                         && prefix[6].Notes.Count == 1 && prefix[6].Notes[0].Html.Contains("footnote in cell"),
                         $"{what}: table row: two cells, \\cellx widths, \\clcbpat shading, footnote in a cell");
        TestReport.Check(Html(8).Contains(HebrewFour) && Html(8).Contains(" end") && !Html(8).Contains("?") && !Html(8).Contains("x"),
                         $"{what}: \\uN with \\uc1 / \\uc2 / \\uc0 skipping, \\bin data with braces skipped");
        TestReport.Check(Regex.Matches(Html(9), "\\[image\\]").Count == 2 && Html(9).Contains("done") && !Html(9).Contains("8950"),
                         $"{what}: \\pict and \\shppict as one placeholder each, \\nonshppict dropped");
        string formats = Html(10);
        TestReport.Check(formats.Contains("font-weight:bold") && formats.Contains("font-style:italic") && formats.Contains("text-decoration:underline")
                         && formats.Contains("line-through") && formats.Contains("font-size:20pt") && formats.Contains("color:#FF0000")
                         && formats.Contains("background-color:yellow") && formats.Contains("vertical-align:super") && formats.Contains("vertical-align:sub"),
                         $"{what}: \\b \\i \\ul \\strike \\fs \\cf \\highlight \\super \\sub");
        TestReport.Check(Html(11).StartsWith("<h2", StringComparison.Ordinal), $"{what}: \\s2 \"heading 2\" renders <h2>");
        TestReport.Check(Html(12).Contains(((char)0x05D0).ToString()) && !Html(12).Contains("ansi fallback"), $"{what}: \\upr shows the \\ud text only");
        TestReport.Check(Html(13).Contains("link text") && !Html(13).Contains("HYPERLINK"), $"{what}: field result shown, instruction hidden");
        TestReport.Check(Regex.Matches(Html(14), "<table").Count == 2 && Html(14).Contains("inner1") && Html(14).Contains("inner2") && Html(14).Contains("outer"),
                         $"{what}: nested table (\\itap2, \\nestcell, \\nestrow in \\nesttableprops) inside the outer cell");
        TestReport.Check(Html(15).Contains("text-align:center") && Html(15).Contains("margin-top:6pt") && Html(15).Contains("margin-bottom:12pt")
                         && Html(15).Contains("line-height:18pt") && Html(15).Contains("\u201Cquoted\u201D\u2014&#160;x<br>y\tz")
                         && Html(15).Contains("margin-inline-start:5pt"),
                         $"{what}: alignment, spacing, exact line spacing, indents, cp1252 quotes, \\emdash \\~ \\line \\tab");
        string[] prefixHtml = prefix.Select(b => b.Html + string.Concat(b.Notes.Select(n => n.Html))).ToArray();
        if (reference is not null) TestReport.Check(prefixHtml.SequenceEqual(reference), $"{what}: same HTML as the trickle-fed in-memory index");

        var random = new Random(7);
        double worst = 0;
        bool ordinals = true;
        for (int i = 0; i < 30; i++)
        {
            int first = SyntheticRtfGenerator.PrefixBlockCount + random.Next(Math.Max(1, count - 40));
            var watch = Stopwatch.StartNew();
            IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(first, 40);
            worst = Math.Max(worst, watch.Elapsed.TotalMilliseconds);
            foreach (WordDocumentBlockHtml block in window)
            {
                int ordinal = block.BlockIndex - SyntheticRtfGenerator.PrefixBlockCount;
                if (!block.Html.Contains(" " + block.BlockIndex)) { ordinals = false; Console.WriteLine($"      block {block.BlockIndex} lacks its ordinal: {block.Html}"); }
                if (ordinal % 10 == 5 && ordinal % 50 != 25 && (block.Notes.Count != 1 || !block.Notes[0].Html.Contains("note body " + block.BlockIndex)))
                { ordinals = false; Console.WriteLine($"      block {block.BlockIndex} lacks its footnote"); }
                if (ordinal % 50 == 25 && !block.Html.StartsWith("<table", StringComparison.Ordinal)) { ordinals = false; Console.WriteLine($"      block {block.BlockIndex} is not a row"); }
                if (ordinal % 7 == 3 && ordinal % 10 != 5 && ordinal % 50 is not (0 or 25) && !block.Html.Contains("font-style:italic"))
                { ordinals = false; Console.WriteLine($"      block {block.BlockIndex}: group state at its start lost"); }
            }
        }
        Console.WriteLine($"   30 random windows of 40 blocks: worst {worst:0.#} ms");
        TestReport.Check(ordinals, $"{what}: every rendered block holds its own ordinal, notes, row and group formatting");
        TestReport.Check(worst < 500, $"{what}: worst random window under 500 ms");
        File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "synthetic-rtf-sample.html"),
            $"<!doctype html><meta charset=utf-8><style>{document.Stylesheet}</style><div class=word-document dir=rtl>{string.Concat(prefixHtml)}</div>");
        return prefixHtml;
    }
}
