using System.Diagnostics;
using FastDocView.DocBinary;
using FastDocView.Docx;

namespace FastDocView.Tests;

/// <summary>
/// Where open time goes, step by step, for a cached (sidecar) reopen and for the first window —
/// the numbers the speed work is planned from. Each step is timed in isolation, in the order the
/// documents perform them; three runs, the last (warm) one printed alongside the first (cold).
/// </summary>
public static class OpenTimingProfile
{
    public static void Run(IEnumerable<string> paths)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        foreach (string path in paths)
        {
            TestReport.Step($"Open timing: {Path.GetFileName(path)}");
            using (WordWindowedDocumentOpener.Open(path, policy)) { }          // make sure the sidecar exists
            for (int run = 1; run <= 3; run++)
            {
                var steps = new List<(string name, double milliseconds)>();
                if (WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.Docx) ProfileDocx(path, policy, steps);
                else ProfileDocBinary(path, policy, steps);
                if (run is 1 or 3)
                    Console.WriteLine($"   run {run}: " + string.Join(" | ", steps.Select(s => $"{s.name} {s.milliseconds:0.#}")) + $" | total {steps.Sum(s => s.milliseconds):0.#} ms");
            }
        }
    }

    private static T Time<T>(List<(string, double)> steps, string name, Func<T> action)
    {
        var watch = Stopwatch.StartNew();
        T result = action();
        steps.Add((name, watch.Elapsed.TotalMilliseconds));
        return result;
    }

    private static void ProfileDocx(string path, WordIndexPolicy policy, List<(string, double)> steps)
    {
        DocxRandomAccessIndex index = Time(steps, "load index", () => DocxRandomAccessIndex.TryLoad(policy.IndexPathFor(path), path)!);
        using var package = Time(steps, "open file", () => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16));
        Time(steps, "decide (zip dir)", () => DocxRandomAccessIndexBuilder.TextPartsBytes(package));
        using (var zip = Time(steps, "zip dir", () => new System.IO.Compression.ZipArchive(package, System.IO.Compression.ZipArchiveMode.Read, leaveOpen: true)))
        {
            Time(steps, "styles.xml full DOM (old)", () => LoadPart(zip, "word/styles.xml"));
        }
        package.Position = 0;
        DocxDocumentParts parts = Time(steps, "parts (styles streamed + settings)", () => new DocxDocumentParts(package));
        var info = new FileInfo(path);
        var decision = new WordIndexDecision(path, WordIndexMode.Sidecar, true, 0, 0, "");
        using DocxWindowedDocument document = Time(steps, "construct", () => DocxWindowedDocument.OpenFromStream(
            new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16), path, decision, index, info.Length, info.LastWriteTimeUtc.Ticks));
        Time(steps, "first 40 blocks", () => document.RenderBlocks(0, 40));
        Time(steps, "headings", () => document.Headings.Count);
        Time(steps, "weights", () => { long sum = 0; for (int i = 0; i < document.BlockCount; i++) sum += document.EstimatedTextLength(i); return sum; });
        Time(steps, "window at middle", () => document.RenderBlocks(document.BlockCount / 2, 40));
    }

    private static System.Xml.Linq.XDocument? LoadPart(System.IO.Compression.ZipArchive zip, string name)
    {
        System.IO.Compression.ZipArchiveEntry? entry = zip.GetEntry(name);
        if (entry is null) return null;
        using Stream stream = entry.Open();
        return System.Xml.Linq.XDocument.Load(stream);
    }

    private static void ProfileDocBinary(string path, WordIndexPolicy policy, List<(string, double)> steps)
    {
        Time(steps, "load index", () => DocBinaryBlockIndex.TryLoad(policy.IndexPathFor(path), path)!);
        using DocBinaryWindowedDocument document = Time(steps, "open (all)", () => DocBinaryWindowedDocument.Open(path, policy));
        Time(steps, "first 40 blocks", () => document.RenderBlocks(0, 40));
        Time(steps, "headings", () => document.Headings.Count);
        Time(steps, "window at middle", () => document.RenderBlocks(document.BlockCount / 2, 40));
    }
}
