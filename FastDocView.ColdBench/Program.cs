using System.Diagnostics;
using System.Globalization;
using FastDocView;

// Cold first paint as the app service does it: WordWindowedDocumentSet.Open (on demand), the summary the page
// asks for, then the first 40 blocks rendered for display. Parent mode runs each file N times, each in a fresh
// child process (Idle priority, one at a time) with an empty cache, and prints the median.
const int FirstWindowBlocks = 40;
if (args.Length == 3 && args[0] == "--child") return Child(args[1], args[2]);
int runs = 3;
var files = new List<string>();
for (int i = 0; i < args.Length; i++)
{
    if (args[i] == "--runs") runs = int.Parse(args[++i], CultureInfo.InvariantCulture);
    else files.Add(args[i]);
}
var rows = new List<string>();
foreach (string file in files)
{
    var paints = new List<double>();
    string last = "";
    for (int run = 0; run < runs; run++)
    {
        string cache = Path.Combine(Path.GetDirectoryName(Environment.ProcessPath)!, "cold-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var start = new ProcessStartInfo(Environment.ProcessPath!) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            start.ArgumentList.Add("--child"); start.ArgumentList.Add(file); start.ArgumentList.Add(cache);
            using Process child = Process.Start(start)!;
            try { child.PriorityClass = ProcessPriorityClass.Idle; } catch (InvalidOperationException) { }
            Task<string> error = child.StandardError.ReadToEndAsync();
            string output = child.StandardOutput.ReadToEnd();
            child.WaitForExit();
            string line = output.Split('\n').Select(l => l.Trim()).FirstOrDefault(l => l.StartsWith("COLD ", StringComparison.Ordinal)) ?? "";
            if (line.Length == 0) { last = "FAILED " + error.Result.Replace('\n', ' '); break; }
            last = line[5..];
            int at = line.IndexOf("paint=", StringComparison.Ordinal) + 6;
            paints.Add(double.Parse(line[at..line.IndexOf(' ', at)], CultureInfo.InvariantCulture));
        }
        finally { try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { } }
    }
    string median = paints.Count == 0 ? "-" : paints.Order().ElementAt(paints.Count / 2).ToString("0", CultureInfo.InvariantCulture);
    string row = $"{Path.GetFileName(file),-40} paint median {median,6} ms   last: {last}";
    Console.WriteLine(row);
    rows.Add(row);
}
return 0;

static int Child(string path, string cacheDirectory)
{
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
    Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
        $"COLD open={open:0} summary={summary - open:0} window={paint - summary:0} paint={paint:0} blocks={rendered}"));
    return 0;
}
