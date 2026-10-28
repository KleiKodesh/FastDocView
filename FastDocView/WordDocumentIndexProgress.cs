namespace FastDocView;

/// <summary>
/// Progress of the one-time index pass over a document. Every long run reports per-step progress
/// (project rule), so the viewer can show "indexing 340 / 1200 MB" instead of a frozen window.
/// </summary>
public readonly record struct WordDocumentIndexProgress(long BytesDone, long BytesTotal, int BlocksFound)
{
    public double Fraction => BytesTotal <= 0 ? 0 : Math.Min(1.0, (double)BytesDone / BytesTotal);
}
