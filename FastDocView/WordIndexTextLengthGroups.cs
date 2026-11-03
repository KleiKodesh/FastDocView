namespace FastDocView;

/// <summary>
/// Text-length totals per group of 64 blocks - the scroll weights a compact index keeps instead of per-block lengths:
/// a block's estimate is its group's average until the span holding it has been scanned (exact afterwards).
/// </summary>
public static class WordIndexTextLengthGroups
{
    public const int GroupSize = 64;

    public static int[] Totals(IReadOnlyList<int> lengths)
    {
        var totals = new int[(lengths.Count + GroupSize - 1) / GroupSize];
        for (int i = 0; i < lengths.Count; i++) totals[i / GroupSize] += lengths[i];
        return totals;
    }

    public static int Average(int[] totals, int blockCount, int blockIndex)
    {
        if ((uint)blockIndex >= (uint)blockCount) return 0;
        int group = blockIndex / GroupSize;
        int members = Math.Min(GroupSize, blockCount - group * GroupSize);
        return totals[group] / members;
    }

    public static void Validate(int[] totals, int blockCount)
    {
        if (totals.Length != (blockCount + GroupSize - 1) / GroupSize) throw new InvalidDataException("Index: text-length groups do not match the blocks.");
    }
}
