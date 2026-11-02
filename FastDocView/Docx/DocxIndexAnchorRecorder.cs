using SeekableDeflate;

namespace FastDocView.Docx;

/// <summary>
/// Records, during the index pass, what the compact index keeps per span anchor: at the first body block that
/// starts past every <see cref="DocxRandomAccessIndex.AnchorSpanBytes"/> of output, the block's index and the
/// list-numbering state at that moment (captured live — it cannot be recovered later). Anchors are independent
/// of the deflate checkpoints (a rescan seeks to the nearest checkpoint anyway); they are dense (~256 KB) so a
/// random block costs one short scan on top of the inflate. <see cref="Build"/> fills in each anchor's previous
/// block from the scanner's lists and stores the list state as a delta from the previous anchor.
/// </summary>
internal sealed class DocxIndexAnchorRecorder
{
    private readonly DocxBodyBlockScanner _scanner;
    private readonly DocxDocumentParts.DocxListLabeller _labeller;
    private readonly List<(int block, byte[] state)> _anchors = new();
    private long _nextAnchorAt;

    public DocxIndexAnchorRecorder(DocxBodyBlockScanner scanner, DocxDocumentParts.DocxListLabeller labeller)
    {
        _scanner = scanner;
        _labeller = labeller;
        scanner.BlockStarted += OnBlockStarted;
    }

    private void OnBlockStarted(int blockIndex, long offset)
    {
        if (blockIndex != 0 && offset < _nextAnchorAt) return;
        _anchors.Add((blockIndex, _labeller.ExportState()));
        _nextAnchorAt = offset + DocxRandomAccessIndex.AnchorSpanBytes;
    }

    /// <summary>The anchors of the finished index (block order), list states as deltas.</summary>
    public DocxIndexAnchor[] Build()
    {
        var anchors = new DocxIndexAnchor[_anchors.Count];
        byte[] previousState = [];
        for (int i = 0; i < anchors.Length; i++)
        {
            (int block, byte[] state) = _anchors[i];
            long previousOffset = block > 0 ? _scanner.BlockStartOffsets[block - 1] : -1;
            WordDocumentBlockKind previousKind = block > 0 ? _scanner.BlockKinds[block - 1] : WordDocumentBlockKind.Other;
            string? previousStyle = null;
            if (block > 0 && _scanner.BlockStyleIds[block - 1] is var styleIndex and >= 0) previousStyle = _scanner.StyleIdTable[styleIndex];
            anchors[i] = new DocxIndexAnchor(block, _scanner.BlockStartOffsets[block], previousOffset, previousKind, previousStyle,
                                             DocxHtml.DocxListCounter.DeltaState(previousState, state));
            previousState = state;
        }
        return anchors;
    }
}
