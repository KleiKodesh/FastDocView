using System.Diagnostics;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// Real files: (1) .txt corpus — a sample (the largest N + M random others) opened and rendered,
/// encoding mix reported, 0 crashes; (2) a Word file list scanned for files that start with "MIME"
/// (first 4 bytes only, shared read) — counted, and each one opened and rendered.
/// </summary>
public static class MarkupCorpusTest
{
    public static void RunText(string listPath, int largest, int random)
    {
        try { Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
        var files = File.ReadAllLines(listPath).Where(line => line.Length > 0 && File.Exists(line))
                        .Select(path => (path, length: new FileInfo(path).Length)).OrderByDescending(f => f.length).ToList();
        var sample = files.Take(largest).ToList();
        var rest = files.Skip(largest).ToList();
        var shuffle = new Random(2026);
        sample.AddRange(rest.OrderBy(_ => shuffle.Next()).Take(random));
        TestReport.Step($".txt corpus: {sample.Count} of {files.Count:N0} files (largest {largest} + random {random})");
        RunSample(sample.Select(f => f.path).ToList(), MarkupSourceKind.PlainText);
    }

    public static void RunMimeScan(string wordListPath)
    {
        try { Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
        string[] paths = File.ReadAllLines(wordListPath).Where(line => line.Length > 0).ToArray();
        TestReport.Step($"Scanning {paths.Length:N0} Word files for a 'MIME' signature (first 4 bytes)");
        var found = new List<string>();
        var header = new byte[4];
        int done = 0;
        foreach (string path in paths)
        {
            if (++done % 2000 == 0) Console.WriteLine($"   {done:N0} / {paths.Length:N0}, {found.Count} found");
            try
            {
                using FileStream file = WordDocumentFileStream.OpenShared(path);
                if (file.Read(header, 0, 4) == 4 && header[0] == 'M' && header[1] == 'I' && header[2] == 'M' && header[3] == 'E') found.Add(path);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) { }
        }
        Console.WriteLine($"   {found.Count} files start with 'MIME'");
        File.WriteAllLines(Path.Combine(TestReport.OutputDirectory, "mime-word-files.txt"), found);
        if (found.Count > 0) RunSample(found, MarkupSourceKind.WebArchive);
    }

    private static void RunSample(List<string> paths, MarkupSourceKind kind)
    {
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        var encodings = new Dictionary<string, int>(StringComparer.Ordinal);
        var failures = new List<string>();
        var watch = Stopwatch.StartNew();
        int done = 0, blocks = 0, headings = 0;
        foreach (string path in paths)
        {
            if (++done % 100 == 0) Console.WriteLine($"   [{watch.Elapsed:mm\\:ss}] {done} / {paths.Count}, {failures.Count} failures");
            try
            {
                using var document = MarkupWindowedDocument.Open(path, kind, policy);
                string key = $"{document.Index.CodePage} ({document.Index.EncodingReason})";
                encodings[key] = encodings.TryGetValue(key, out int n) ? n + 1 : 1;
                blocks += document.BlockCount;
                headings += document.Headings.Count;
                document.RenderBlocks(0, 50);
                if (document.BlockCount > 50) document.RenderBlocks(document.BlockCount / 2, 50);
                if (document.BlockCount > 100) document.RenderBlocks(document.BlockCount - 50, 50);
            }
            catch (Exception exception)
            {
                failures.Add($"{exception.GetType().Name}: {exception.Message} [{done}]");
            }
        }
        Console.WriteLine($"   {paths.Count} files in {watch.Elapsed.TotalSeconds:0.#} s, {blocks:N0} blocks, {headings:N0} headings");
        foreach (KeyValuePair<string, int> pair in encodings.OrderByDescending(p => p.Value)) Console.WriteLine($"   encoding {pair.Key}: {pair.Value}");
        foreach (string failure in failures.Take(20)) Console.WriteLine($"   failure {failure}");
        TestReport.Check(failures.Count == 0, $"{kind} corpus: {failures.Count} crashes of {paths.Count}");
    }
}
