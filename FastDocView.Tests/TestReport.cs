using System.Diagnostics;

namespace FastDocView.Tests;

/// <summary>Per-step console reporting with timings; every check prints PASS/FAIL and the run exits non-zero on any FAIL.</summary>
public static class TestReport
{
    private static int _failureCount;

    /// <summary>Checks that failed so far (counted atomically: some checks run on worker threads).</summary>
    public static int FailureCount => Volatile.Read(ref _failureCount);

    public static void Step(string message) => Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {message}");

    public static void Check(bool condition, string what)
    {
        if (!condition) Interlocked.Increment(ref _failureCount);
        Console.WriteLine($"   {(condition ? "PASS" : "FAIL")}  {what}");
    }

    public static T Timed<T>(string what, Func<T> action)
    {
        var watch = Stopwatch.StartNew();
        T result = action();
        Console.WriteLine($"   {what}: {watch.Elapsed.TotalMilliseconds:0.#} ms");
        return result;
    }

    public static void Timed(string what, Action action) => Timed(what, () => { action(); return 0; });

    public static string OutputDirectory
    {
        get
        {
            // The FastDocView repo root is the folder holding FastDocView.Tests; found by walking up so
            // any output depth (rid, publish, worktree) lands in its gitignored TestOutput.
            DirectoryInfo? folder = new DirectoryInfo(AppContext.BaseDirectory);
            while (folder != null && !Directory.Exists(Path.Combine(folder.FullName, "FastDocView.Tests")))
                folder = folder.Parent;
            if (folder == null)
                throw new InvalidOperationException($"FastDocView repo root not found above {AppContext.BaseDirectory}");
            string directory = Path.Combine(folder.FullName, "TestOutput");
            Directory.CreateDirectory(directory);
            return directory;
        }
    }
}
