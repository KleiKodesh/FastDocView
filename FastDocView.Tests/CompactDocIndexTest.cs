using FastDocView.DocBinary;

namespace FastDocView.Tests;

/// <summary>
/// The compact .doc index: for random + span-edge blocks of a real .doc the compact table's CP range, kind and list
/// label (rescanned from the span anchor) equal a reference pass that scans the whole text once with the same
/// scanner; plus the numbers (sidecar size, reopen, first window, random block median / p90).
/// </summary>
public static class CompactDocIndexTest
{
    public static void Run(IEnumerable<string> docPaths)
    {
        foreach (string path in docPaths)
        {
            TestReport.Step($"Compact .doc index: {Path.GetFileName(path)}");
            string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-compact");
            var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (DocBinaryWindowedDocument first = TestReport.Timed("cold open (index pass)", () => DocBinaryWindowedDocument.Open(path, policy)))
            {
                Console.WriteLine($"   {first.BlockCount:N0} blocks, {first.BlockIndex.Anchors.Length} anchors, sidecar {new FileInfo(indexPath).Length:N0} bytes");
                foreach ((string section, long bytes) in first.BlockIndex.SizeBreakdown()) Console.WriteLine($"      {bytes,9:N0}  {section}");
            }
            using DocBinaryWindowedDocument document = TestReport.Timed("reopen from sidecar", () => DocBinaryWindowedDocument.Open(path, policy));
            TestReport.Timed("first window (40 blocks)", () => document.RenderBlocks(0, 40));
            TestReport.Check(document.IsIndexComplete && document.KnownBlockCount == document.BlockCount, "index complete, known = all");

            // reference: one pass with the same scanner over the whole text
            var reference = new DocBinaryBlockScanner(RendererOf(document), document.StyleSheet, document.ListDefinitions);
            int mainText = document.FileInformationBlock.MainTextCharacterCount;
            TestReport.Timed("reference full scan", () =>
            {
                document.PieceTable.ScanCharacters(0, mainText, (characters, count, cp) => reference.Feed(characters, count, cp));
                reference.Finish(mainText);
            });
            TestReport.Check(reference.BlockCount == document.BlockCount, "block count matches the reference scan");

            var random = new Random(17);
            var probes = new List<int>();
            for (int i = 0; i < 200; i++) probes.Add(random.Next(document.BlockCount));
            var times = new List<double>();
            foreach (int block in probes)
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                document.RenderBlocks(block, 1);
                times.Add(watch.Elapsed.TotalMilliseconds);
            }
            times.Sort();
            Console.WriteLine($"   {probes.Count} random single blocks (render): median {times[times.Count / 2]:F1} ms, p90 {times[times.Count * 9 / 10]:F1} ms");

            foreach (WordIndexSpanAnchor anchor in document.BlockIndex.Anchors) { probes.Add(anchor.FirstBlock); if (anchor.FirstBlock > 0) probes.Add(anchor.FirstBlock - 1); }
            probes.Add(document.BlockCount - 1);
            int mismatches = 0;
            foreach (int block in probes)
            {
                (int start, int end) = document.BlockRange(block);
                int referenceEnd = block + 1 < reference.BlockCount ? reference.BlockStartCps[block + 1] : mainText;
                int labelIndex = reference.BlockListLabelIds[block];
                string? referenceLabel = labelIndex >= 0 ? reference.ListLabels[labelIndex] : null;
                bool ok = start == reference.BlockStartCps[block] && end == referenceEnd
                       && document.BlockKindOf(block) == reference.BlockKinds[block]
                       && document.EstimatedTextLength(block) == end - start
                       && document.ListLabelOf(block) == referenceLabel;
                if (!ok && mismatches++ < 5) Console.WriteLine($"   mismatch at block {block}: {start}-{end} vs {reference.BlockStartCps[block]}-{referenceEnd}");
            }
            TestReport.Check(mismatches == 0, $"{probes.Count} random + span-edge blocks: CP range, kind, text length and list label equal the reference scan");
            int cpMismatches = 0;
            for (int i = 0; i < 50; i++)
            {
                int cp = random.Next(Math.Max(1, mainText));
                int expected = DocBinaryBlockIndex.BlockAt(reference.BlockStartCps, cp);
                if (document.BlockAtCp(cp) != expected && cpMismatches++ < 5) Console.WriteLine($"   BlockAt({cp}) mismatch");
            }
            TestReport.Check(cpMismatches == 0, "block-at-CP lookups equal the reference");
            int labelled = Enumerable.Range(0, reference.BlockCount).Count(b => reference.BlockListLabelIds[b] >= 0);
            Console.WriteLine($"   reference: {labelled:N0} list paragraphs, {reference.ListLabels.Count} distinct labels, {reference.Headings.Count} headings");
        }
    }

    private static DocBinaryParagraphHtmlRenderer RendererOf(DocBinaryWindowedDocument document) => document.Renderer;
}
