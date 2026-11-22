using System.Diagnostics;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Cold first open, per format: every run is a fresh process (nothing JIT-warmed, no static state) with an empty
/// sidecar cache, timed to its first paint - the open (as the dev server / app service does it, on-demand indexing)
/// plus the summary the page asks for and the first window of 40 blocks rendered for display.
///   --cold-bench [--runs N] file...   runs each file N times (default 3) in child processes, prints a table
///   --cold-open file cacheDirectory   one child run (prints one "COLD ..." line)
/// Children run at Idle priority; the parent waits for each in turn (one at a time).
/// </summary>
public static class ColdOpenBenchmark
{
    private const int FirstWindowBlocks = 40;

    public static int RunChild(string path, string cacheDirectory)
    {
        double processMilliseconds = (DateTime.Now - Process.GetCurrentProcess().StartTime).TotalMilliseconds;
        var watch = Stopwatch.StartNew();
        using var documents = new WordWindowedDocumentSet(cacheDirectory);
        string id = documents.Open(path);
        double open = watch.Elapsed.TotalMilliseconds;
        documents.Use(id, document =>
        {
            _ = document.BlockCount; _ = document.Stylesheet; _ = document.Headings.Count; _ = document.NoteCounts;
            _ = document.KnownBlockCount; _ = document.IsIndexComplete;
            return 0;
        });
        double summary = watch.Elapsed.TotalMilliseconds;
        int rendered = documents.Use(id, document => document.RenderBlocksForDisplay(0, FirstWindowBlocks).Count);
        double paint = watch.Elapsed.TotalMilliseconds;
        Console.WriteLine(string.Create(System.Globalization.CultureInfo.InvariantCulture,
            $"COLD start={processMilliseconds:0} open={open:0} summary={summary - open:0} window={paint - summary:0} paint={paint:0} blocks={rendered}"));
        if (Environment.GetEnvironmentVariable("FASTDOCVIEW_COLD_WARM") == "1")
        {
            // the same open again in this (now warm) process, with a fresh cache: what is left is the work itself, not JIT
            string warmCache = cacheDirectory + "-warm";
            watch.Restart();
            using (var warm = new WordWindowedDocumentSet(warmCache))
            {
                string warmId = warm.Open(path);
                double warmOpen = watch.Elapsed.TotalMilliseconds;
                warm.Use(warmId, document => document.RenderBlocksForDisplay(0, FirstWindowBlocks).Count);
                Console.WriteLine(string.Create(System.Globalization.CultureInfo.InvariantCulture, $"WARM open={warmOpen:0} paint={watch.Elapsed.TotalMilliseconds:0}"));
            }
            try { Directory.Delete(warmCache, recursive: true); } catch (IOException) { }
        }
        return 0;
    }

    public static void RunParent(IReadOnlyList<string> paths, int runs)
    {
        string self = Environment.ProcessPath!;
        string selfDll = typeof(ColdOpenBenchmark).Assembly.Location;
        bool isHostDotnet = Path.GetFileNameWithoutExtension(self).Equals("dotnet", StringComparison.OrdinalIgnoreCase);
        var rows = new List<string>();
        foreach (string path in paths)
        {
            var paints = new List<double>();
            string last = "";
            for (int run = 0; run < runs; run++)
            {
                string cache = Path.Combine(TestReport.OutputDirectory, "cold-cache-" + Guid.NewGuid().ToString("N"));
                try
                {
                    var start = new ProcessStartInfo(self)
                    {
                        RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false,
                    };
                    if (isHostDotnet) start.ArgumentList.Add(selfDll);
                    start.ArgumentList.Add("--cold-open"); start.ArgumentList.Add(path); start.ArgumentList.Add(cache);
                    start.Environment["DOTNET_TieredPGO"] = Environment.GetEnvironmentVariable("DOTNET_TieredPGO") ?? "1";
                    var wall = Stopwatch.StartNew();
                    using Process child = Process.Start(start)!;
                    try { child.PriorityClass = ProcessPriorityClass.Idle; } catch (InvalidOperationException) { }
                    Task<string> errorRead = child.StandardError.ReadToEndAsync();     // both pipes at once: a big stderr must not block the child
                    string output = child.StandardOutput.ReadToEnd();
                    string error = errorRead.Result;
                    child.WaitForExit();
                    string line = output.Split('\n').Select(l => l.Trim()).FirstOrDefault(l => l.StartsWith("COLD ")) ?? "";
                    if (line.Length == 0) { last = "FAILED " + (error.Length > 200 ? error[..200] : error).Replace('\n', ' '); break; }
                    last = line[5..] + $" wall={wall.Elapsed.TotalMilliseconds:0}";
                    paints.Add(double.Parse(Value(line, "paint"), System.Globalization.CultureInfo.InvariantCulture));
                }
                finally { try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { } }
            }
            string median = paints.Count == 0 ? "-" : paints.OrderBy(p => p).ElementAt(paints.Count / 2).ToString("0");
            string row = $"{Path.GetFileName(path),-34} {new FileInfo(path).Length / (1 << 20),5} MB  paint median {median,6} ms   last: {last}";
            Console.WriteLine(row);
            rows.Add(row);
        }
        Console.WriteLine();
        Console.WriteLine("Cold first paint (fresh process, empty cache, " + runs + " runs):");
        foreach (string row in rows) Console.WriteLine("  " + row);
    }

    private static string Value(string line, string key)
    {
        int at = line.IndexOf(key + "=", StringComparison.Ordinal) + key.Length + 1;
        int end = line.IndexOf(' ', at);
        return end < 0 ? line[at..] : line[at..end];
    }
}
