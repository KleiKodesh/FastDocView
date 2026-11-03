namespace FastDocView.Tests;

/// <summary>
/// The render times of a run of random windows, judged by their median and 95th percentile instead of the single worst
/// one. The harness runs at idle priority on a shared machine: one window that lost the CPU (or met a GC or a first-call
/// JIT) can take a second while the code is fine, and a worst-of-N check then fails on load, not on code. A real slowdown
/// moves every window, so the median catches it; the p95 catches a slow path hit by a few windows (more than one window
/// in twenty). The worst is still printed.
/// </summary>
public sealed class WindowTimings
{
    private readonly List<double> _milliseconds = [];

    public void Add(double milliseconds) => _milliseconds.Add(milliseconds);

    public double Worst => _milliseconds.Count == 0 ? 0 : _milliseconds.Max();
    public double Median => Percentile(0.5);
    public double Percentile95 => Percentile(0.95);

    /// <summary>Nearest-rank percentile.</summary>
    private double Percentile(double fraction)
    {
        if (_milliseconds.Count == 0) return 0;
        double[] sorted = _milliseconds.OrderBy(value => value).ToArray();
        int rank = (int)Math.Ceiling(fraction * sorted.Length);
        return sorted[Math.Clamp(rank, 1, sorted.Length) - 1];
    }

    /// <summary>Prints the spread and checks the median and the p95 against their budgets.</summary>
    public void Check(string what, double medianBudgetMilliseconds, double percentile95BudgetMilliseconds)
    {
        Console.WriteLine($"   {what}: {_milliseconds.Count} windows, median {Median:0.#} ms, p95 {Percentile95:0.#} ms, worst {Worst:0.#} ms");
        TestReport.Check(Median < medianBudgetMilliseconds, $"{what}: median window under {medianBudgetMilliseconds:0} ms ({Median:0.#})");
        TestReport.Check(Percentile95 < percentile95BudgetMilliseconds, $"{what}: p95 window under {percentile95BudgetMilliseconds:0} ms ({Percentile95:0.#})");
    }
}
