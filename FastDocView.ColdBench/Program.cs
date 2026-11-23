using System.Diagnostics;
using System.Globalization;
using FastDocView;

// Cold first paint as the app service does it: WordWindowedDocumentSet.Open (on demand), the summary the page
// asks for, then the first 40 blocks rendered for display. Parent mode runs each file N times, each in a fresh
// child process (Idle priority, one at a time) with an empty cache, and prints the median.
// --notes: the note paths instead, each beside its body equivalent (see NoteProbe): a single note by index vs its own
// block, a notes page of 20 vs a body window of 20 at the page's first note block, a tooltip lookup (kind + number) vs the
// note's block. Cold = the first call in a fresh process (after the index completed), warm = median of 20 more at random
// places (the same seed for the note and the body side). Prints extensions only.
const int FirstWindowBlocks = 40;
if (args.Length == 3 && args[0] == "--child") return Child(args[1], args[2]);
if (args.Length == 4 && args[0] == "--child-probe") return NoteProbe(args[1], args[2], args[3]);
int runs = 3;
bool notes = false;
var files = new List<string>();
for (int i = 0; i < args.Length; i++)
{
    if (args[i] == "--runs") runs = int.Parse(args[++i], CultureInfo.InvariantCulture);
    else if (args[i] == "--notes") notes = true;
    else files.Add(args[i]);
}
if (notes) return RunNoteProbes(files, runs);
var rows = new List<string>();
foreach (string file in files)
{
    var paints = new List<double>();
    string last = "";
    for (int run = 0; run < runs; run++)
    {
        string cache = Path.Combine(AppContext.BaseDirectory, "cold-cache-" + Guid.NewGuid().ToString("N"));
        try
        {
            var start = new ProcessStartInfo(Environment.ProcessPath!) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
            // under "dotnet FastDocView.ColdBench.dll" (not the published exe) the process is the dotnet host: the child needs the dll
            if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
                start.ArgumentList.Add(Path.Combine(AppContext.BaseDirectory, "FastDocView.ColdBench.dll"));
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
        finally { try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
    }
    string median = paints.Count == 0 ? "-" : paints.Order().ElementAt(paints.Count / 2).ToString("0", CultureInfo.InvariantCulture);
    string row = $"{Path.GetFileName(file),-40} paint median {median,6} ms   last: {last}";
    Console.WriteLine(row);
    rows.Add(row);
}
return 0;

static string RunChild(params string[] arguments)
{
    var start = new ProcessStartInfo(Environment.ProcessPath!) { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false };
    if (string.Equals(Path.GetFileNameWithoutExtension(Environment.ProcessPath), "dotnet", StringComparison.OrdinalIgnoreCase))
        start.ArgumentList.Add(Path.Combine(AppContext.BaseDirectory, "FastDocView.ColdBench.dll"));
    foreach (string argument in arguments) start.ArgumentList.Add(argument);
    using Process child = Process.Start(start)!;
    try { child.PriorityClass = ProcessPriorityClass.Idle; } catch (InvalidOperationException) { }
    Task<string> error = child.StandardError.ReadToEndAsync();
    string output = child.StandardOutput.ReadToEnd();
    child.WaitForExit();
    return output.Split('\n').Select(l => l.Trim()).FirstOrDefault(l => l.StartsWith("PROBE ", StringComparison.Ordinal)) ?? "FAILED " + error.Result.Replace('\n', ' ');
}

static int RunNoteProbes(List<string> files, int runs)
{
    // pairs: the note path, then its body equivalent
    // (noteblocks20: the page's own note blocks through the body path, one call per run of adjacent blocks - the exact
    // body equivalent of a notes page; body20 is the same count of blocks where the page starts)
    string[] probes = ["note1", "block1", "tooltip", "notes20", "noteblocks20", "body20"];
    Console.WriteLine($"{"file",-8} {"probe",-12} {"cold ms",8} {"warm ms",8}");
    foreach (string file in files)
    {
        File.ReadAllBytes(file);     // the OS file cache warm for every probe alike (else the first probe pays the disk)
        foreach (string probe in probes)
        {
            var cold = new List<double>(); var warm = new List<double>(); string failure = "";
            for (int run = 0; run < runs; run++)
            {
                string cache = Path.Combine(AppContext.BaseDirectory, "cold-cache-" + Guid.NewGuid().ToString("N"));
                try
                {
                    string line = RunChild("--child-probe", file, cache, probe);
                    if (!line.StartsWith("PROBE ", StringComparison.Ordinal)) { failure = line; break; }
                    string[] parts = line.Split(' ');
                    cold.Add(double.Parse(parts[1], CultureInfo.InvariantCulture)); warm.Add(double.Parse(parts[2], CultureInfo.InvariantCulture));
                }
                finally { try { if (Directory.Exists(cache)) Directory.Delete(cache, recursive: true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
            }
            static string Median(List<double> values) => values.Count == 0 ? "-" : values.Order().ElementAt(values.Count / 2).ToString("0.00", CultureInfo.InvariantCulture);
            Console.WriteLine($"{Path.GetExtension(file),-8} {probe,-12} {Median(cold),8} {Median(warm),8} {failure}");
        }
    }
    return 0;
}

static int NoteProbe(string path, string cacheDirectory, string probe)
{
    using var documents = new WordWindowedDocumentSet(cacheDirectory);
    string id = documents.Open(path);
    return documents.Use(id, document =>
    {
        while (!document.IsIndexComplete) document.ScanAhead(100_000);
        IReadOnlyList<WordDocumentNoteReference> references = document.NoteReferences;
        if (references.Count == 0) { Console.WriteLine("PROBE 0 0"); return 0; }
        var random = new Random(7);
        int[] targets = [references.Count / 2, .. Enumerable.Range(0, 20).Select(_ => random.Next(references.Count))];
        var times = new List<double>();
        foreach (int note in targets)
        {
            WordDocumentNoteReference reference = references[note];
            var watch = Stopwatch.StartNew();
            int produced = probe switch
            {
                "note1" => document.RenderNotes(note, 1).Count,
                "block1" => document.RenderBlocksForDisplay(reference.BlockIndex, 1).Count,
                "tooltip" => TooltipLookup(document, reference.Kind, reference.Number),
                "notes20" => document.RenderNotesForDisplay(note, 20).Count,
                "body20" => document.RenderBlocksForDisplay(reference.BlockIndex, 20).Count,
                "noteblocks20" => NoteBlocksAsBody(document, references, note, 20),
                _ => throw new ArgumentException(probe),
            };
            times.Add(watch.Elapsed.TotalMilliseconds);
            if (produced == 0) throw new InvalidOperationException("nothing rendered");
        }
        double warmMedian = times.Skip(1).Order().ElementAt(10);
        Console.WriteLine(string.Create(CultureInfo.InvariantCulture, $"PROBE {times[0]:0.000} {warmMedian:0.000}"));
        return 0;
    });
}

static int NoteBlocksAsBody(IWordWindowedDocument document, IReadOnlyList<WordDocumentNoteReference> references, int first, int count)
{
    int end = Math.Min(references.Count, first + count), rendered = 0;
    for (int i = first; i < end;)
    {
        int runStart = references[i].BlockIndex, runEnd = runStart;
        while (i < end && references[i].BlockIndex <= runEnd + 1) runEnd = references[i++].BlockIndex;
        rendered += document.RenderBlocksForDisplay(runStart, runEnd - runStart + 1).Count;
    }
    return rendered;
}

// the tooltip of a note reference in the page: the note's index found by kind + number, then that note for display
static int TooltipLookup(IWordWindowedDocument document, WordDocumentNoteKind kind, int number)
{
    IReadOnlyList<WordDocumentNoteReference> references = document.NoteReferences;
    for (int i = 0; i < references.Count; i++)
        if (references[i].Kind == kind && references[i].Number == number) return document.RenderNotesForDisplay(i, 1).Count;
    return 0;
}

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

