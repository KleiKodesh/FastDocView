namespace FastDocView.Odt;

/// <summary>
/// Records, during the index pass, what the compact index keeps per span anchor: at the first block that starts past
/// every <see cref="OdtRandomAccessIndex.AnchorSpanBytes"/> of body (and at block 0), the block's index + offset and the
/// scanner's frames + list-numbering state at that moment (captured live, before the block's label is counted - it
/// cannot be recovered later). <see cref="Build"/> fills in each anchor's previous block from the scanner's lists and
/// stores the state as a delta from the previous anchor. The one-go and the on-demand builders attach the same recorder,
/// so both produce the same anchors.
/// </summary>
internal sealed class OdtIndexAnchorRecorder
{
    private readonly OdtBodyBlockScanner _scanner;
    private readonly List<(int block, long offset, byte[] state)> _anchors = new();
    private long _nextAnchorAt;

    public OdtIndexAnchorRecorder(OdtBodyBlockScanner scanner)
    {
        _scanner = scanner;
        scanner.BlockStarting += OnBlockStarting;
    }

    private void OnBlockStarting(int blockIndex, long offset)
    {
        if (blockIndex != 0 && offset < _nextAnchorAt) return;
        if (_anchors.Count > 0 && _anchors[_anchors.Count - 1].block == blockIndex) return;     // the same block starting twice cannot happen; guards a damaged body
        _anchors.Add((blockIndex, offset, _scanner.ExportScanState()));
        _nextAnchorAt = offset + OdtRandomAccessIndex.AnchorSpanBytes;
    }

    /// <summary>The anchors of the finished index (block order, only those within <paramref name="blockCount"/>), scan states as deltas.</summary>
    public WordIndexSpanAnchor[] Build(int blockCount)
    {
        var kept = _anchors.Where(anchor => anchor.block < blockCount).ToList();
        var anchors = new WordIndexSpanAnchor[kept.Count];
        byte[] previousState = [];
        for (int i = 0; i < anchors.Length; i++)
        {
            (int block, long offset, byte[] state) = kept[i];
            anchors[i] = new WordIndexSpanAnchor(block, offset, block > 0 ? _scanner.BlockStartOffsets[block - 1] : -1,
                                                 block > 0 ? _scanner.BlockKinds[block - 1] : WordDocumentBlockKind.Other,
                                                 OdtBodyBlockScanner.DeltaScanState(previousState, state));
            previousState = state;
        }
        return anchors;
    }
}
