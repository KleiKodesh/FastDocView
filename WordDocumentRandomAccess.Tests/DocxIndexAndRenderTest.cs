using System.Text.RegularExpressions;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Synthetic big .docx: index build time and size, exact block count, cached re-open time,
/// random windows render in milliseconds, and each rendered block contains its own ordinal.
/// </summary>
public static class DocxIndexAndRenderTest
{
    public static void Run(int blockCount)
    {
        // the generator's structure is part of the checks — bump the suffix when it changes so a stale file is never reused
        string path = Path.Combine(TestReport.OutputDirectory, $"synthetic-{blockCount}-v2.docx");
        string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache");
        TestReport.Step($"Synthetic docx with {blockCount:N0} blocks");
        if (!File.Exists(path)) TestReport.Timed("generate", () => SyntheticDocxGenerator.Write(path, blockCount));
        Console.WriteLine($"   file size {new FileInfo(path).Length / 1024 / 1024} MB");

        string indexPath = WordDocumentIndexCacheLocation.IndexPathFor(path, cacheDirectory);
        if (File.Exists(indexPath)) File.Delete(indexPath);

        long lastReported = -1;
        var progress = new Progress<WordDocumentIndexProgress>(p =>
        {
            long megabyte = p.BytesDone >> 24;
            if (megabyte != lastReported) { lastReported = megabyte; Console.WriteLine($"      indexing {p.BytesDone >> 20} / {p.BytesTotal >> 20} MB, {p.BlocksFound:N0} blocks"); }
        });
        var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
        DocxRandomAccessIndex index;
        using (var cold = (DocxWindowedDocument)TestReport.Timed("index build (cold)", () => WordWindowedDocumentOpener.Open(path, policy, progress))) index = cold.Index;
        Console.WriteLine($"   document.xml {index.DocumentEntry.UncompressedSize >> 20} MB uncompressed, index file {new FileInfo(indexPath).Length >> 10} KB, {index.Checkpoints.Count} checkpoints");
        TestReport.Check(index.BlockCount == SyntheticDocxGenerator.ExpectedBlockCount(blockCount), $"block count {index.BlockCount:N0} == expected {SyntheticDocxGenerator.ExpectedBlockCount(blockCount):N0}");
        TestReport.Check(index.BlockKinds[199] == WordDocumentBlockKind.Table, "block 199 is a table");
        TestReport.Check(index.BlockKinds[index.BlockCount - 1] == WordDocumentBlockKind.Other, "last block (sectPr) is Other");
        TestReport.Check(index.StyleIdOf(50) == "Heading1", "block 50 has style Heading1");

        using DocxWindowedDocument document = TestReport.Timed("open from cached index", () => DocxWindowedDocument.Open(path, policy));
        TestReport.Check(document.Decision.HasCachedIndex, "cached index used");
        TestReport.Check(document.IsRightToLeft, "document detected as right-to-left");

        var random = new Random(3);
        bool allContainOrdinal = true;
        double worst = 0;
        for (int i = 0; i < 40; i++)
        {
            int first = random.Next(Math.Max(1, document.BlockCount - 40));
            var watch = System.Diagnostics.Stopwatch.StartNew();
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(first, 40);
            worst = Math.Max(worst, watch.Elapsed.TotalMilliseconds);
            foreach (WordDocumentBlockHtml block in blocks)
            {
                if (block.Kind == WordDocumentBlockKind.Other) continue;
                string expected = block.Kind == WordDocumentBlockKind.Table ? $"טבלה {block.BlockIndex} " : $"פסקה {block.BlockIndex}";
                if (block.BlockIndex % 10 == 3 && block.BlockIndex % 50 != 0 && !block.Html.Contains("list-marker")) { allContainOrdinal = false; Console.WriteLine($"   block {block.BlockIndex} lacks its list marker"); }
                if (!block.Html.Contains(expected)) { allContainOrdinal = false; Console.WriteLine($"   block {block.BlockIndex} lacks '{expected}': {Truncate(block.Html)}"); }
            }
        }
        Console.WriteLine($"   40 random windows of 40 blocks: worst {worst:0.#} ms");
        TestReport.Check(allContainOrdinal, "every rendered block contains its own ordinal");
        TestReport.Check(worst < 500, "worst random window under 500 ms");

        // sequential windows (the common scroll case) should be faster than the cold seek
        var sequential = System.Diagnostics.Stopwatch.StartNew();
        for (int first = 1000; first < 1400; first += 40) document.RenderBlocks(first, 40);
        Console.WriteLine($"   10 consecutive windows: {sequential.Elapsed.TotalMilliseconds:0.#} ms total");

        IReadOnlyList<WordDocumentHeading> headings = TestReport.Timed("collect headings", () => document.Headings);
        TestReport.Check(headings.Count == (blockCount + 49) / 50 - CountTablesOnHeadingSlots(blockCount), $"{headings.Count} headings found");
        TestReport.Check(Regex.IsMatch(document.RenderBlocks(0, 1)[0].Html, "^<h1[^>]*dir=\"rtl\""), "heading block renders as <h1 dir=rtl>");

        File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "synthetic-sample.html"),
            $"<!doctype html><meta charset=utf-8><style>{document.Stylesheet}</style><div class=word-document dir=rtl>{string.Concat(document.RenderBlocks(0, 60).Select(b => b.Html))}</div>");
    }

    private static int CountTablesOnHeadingSlots(int blockCount)
    {
        int count = 0;
        for (int i = 0; i < blockCount; i += 50) if (i % 200 == 199) count++;
        return count;
    }

    private static string Truncate(string html) => html.Length > 160 ? html.Substring(0, 160) + "…" : html;
}
