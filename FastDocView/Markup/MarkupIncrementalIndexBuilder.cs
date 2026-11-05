using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// The one streaming index pass of <see cref="MarkupRandomAccessIndexBuilder"/> run in steps: <see cref="AdvanceUntilKnown"/>
/// feeds on (.txt: one 64 KB piece at a time; MHTML: one encoded line at a time) until that many blocks are known;
/// <see cref="CreatePartialIndex"/> is the compact index of the blocks known so far (a window reads them exactly as it will from
/// the final index); <see cref="Complete"/> the final index. One pass either way: the same blocks, anchors and values.
/// A block is known once the next one has started (its bounds and text length are final then).
/// </summary>
internal sealed class MarkupIncrementalIndexBuilder
{
    private const int RightToLeftSampleBytes = 64 * 1024;
    private const int HandOffBlocks = 4096;                 // finished blocks kept before they are folded into the groups

    private readonly MarkupSourceKind _kind;
    private readonly Stream _file;
    private readonly long _sourceLength, _sourceTicks;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private CancellationToken _cancellation;
    private readonly MarkupBlockScanner _scanner;
    private readonly List<int> _groupTotals = [];
    private readonly List<MarkupSpanAnchor> _anchors;
    private readonly int _codePage, _unitSize;
    private readonly string _reason;
    private readonly Encoding _encoding;
    private MarkupRandomAccessIndex? _final;

    // .txt
    private readonly byte[]? _chunk;
    private readonly int _bomLength, _sampleLength;
    private readonly bool _isPlainTextRightToLeft;
    private readonly ToratEmetHead? _toratEmet;             // a Torat Emet book: its metadata head
    private long _done;

    // MHTML
    private readonly WebArchiveHtmlPart? _part;
    private readonly MimeTransferDecoder? _decoder;
    private readonly long _startResume;
    private (MarkupScanState state, long position)? _pending;
    private readonly long[]? _resumeRing;
    private readonly List<long>? _resumes;                  // aligned with the scanner's (not yet handed off) blocks
    private readonly byte[]? _decoded;
    private readonly long[]? _resumePiece;
    private readonly List<byte>? _bodySample;
    private long _decodedTotal, _encodedEnd;
    private int _pieces;
    private readonly MarkupLineReader? _lines;

    public bool IsEveryBlockKnown => _final is not null;

    /// <summary>Blocks whose values are final (the next block has started), or all of them once the pass ended.</summary>
    public int KnownBlockCount => _final?.BlockCount ?? Math.Max(0, CurrentBlock);

    /// <summary>MHTML: the body sample the document's direction is read from is full (or the pass ended); .txt: always.</summary>
    public bool HasDirectionSample => _kind != MarkupSourceKind.WebArchive || IsEveryBlockKnown || _bodySample!.Count >= RightToLeftSampleBytes;

    private int CurrentBlock => _scanner.BlockIndexBase + _scanner.BlockStartOffsets.Count - 1;

    public MarkupIncrementalIndexBuilder(MarkupSourceKind kind, Stream file, long sourceLength, long sourceTicks,
                                         IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation, bool isToratEmetByPath = false)
    {
        _kind = kind; _file = file; _sourceLength = sourceLength; _sourceTicks = sourceTicks; _progress = progress; _cancellation = cancellation;
        if (kind == MarkupSourceKind.PlainText)
        {
            file.Position = 0;
            _chunk = new byte[Math.Max(MarkupRandomAccessIndexBuilder.ChunkSize, MarkupTextEncoding.DetectionSampleBytes)];
            _sampleLength = MarkupRandomAccessIndexBuilder.ReadFully(file, _chunk, MarkupTextEncoding.DetectionSampleBytes);
            (_codePage, _bomLength, _reason) = MarkupTextEncoding.DetectPlainText(_chunk.AsSpan(0, _sampleLength),
                _sampleLength < MarkupTextEncoding.DetectionSampleBytes ? _sampleLength : file.Length); // a short read (file shrank) = the file ends there
            (_, _unitSize, _encoding) = MarkupRandomAccessIndexBuilder.PlainTextEncoding(_codePage);
            int sampleText = Math.Min(_sampleLength - _bomLength, RightToLeftSampleBytes) & ~(_unitSize - 1);
            _isPlainTextRightToLeft = MarkupTextEncoding.ContainsHebrew(_encoding.GetString(_chunk, _bomLength, Math.Max(0, sampleText)));
            _toratEmet = isToratEmetByPath ? ToratEmetPathRule.ReadHead(_chunk.AsSpan(0, _sampleLength), _bomLength, _encoding)   // the path rule: no detection
                                           : ToratEmetFormat.TryReadHead(_chunk.AsSpan(0, _sampleLength), _bomLength, _encoding);
            _scanner = new MarkupBlockScanner(MarkupBlockMode.Lines, _unitSize, _codePage == MarkupTextEncoding.Utf16BigEndian, _encoding, _toratEmet?.HeadEndOffset ?? -1)
                { BookName = _toratEmet?.BookName };
            _anchors = [new(null, _bomLength, 0, 0)];
            _scanner.Begin(_bomLength);
            _scanner.Feed(_chunk, _bomLength, _sampleLength - _bomLength, _bomLength);
            _done = _sampleLength;
            AfterPlainTextPiece();
            return;
        }

        _part = WebArchiveContainer.LocateHtmlPart(file);
        (_codePage, _reason, _unitSize, _encoding) = MarkupRandomAccessIndexBuilder.WebArchiveEncoding(_part.Charset);
        _scanner = new MarkupBlockScanner(MarkupBlockMode.Elements, _unitSize, _codePage == MarkupTextEncoding.Utf16BigEndian, _encoding);
        _decoder = new MimeTransferDecoder(_part.TransferEncoding);
        _startResume = MimeTransferDecoder.ResumePoint(_part.BodyStart, 0);
        _anchors = [new(null, 0, _startResume, _startResume)];
        _resumeRing = new long[MarkupRandomAccessIndexBuilder.ResumeRingSize];
        _resumes = [];
        _decoded = new byte[MarkupRandomAccessIndexBuilder.FeedPieceBytes + 4];
        _resumePiece = new long[MarkupRandomAccessIndexBuilder.FeedPieceBytes + 4];
        _bodySample = [];
        _encodedEnd = file.Length;
        file.Position = _part.BodyStart;
        _lines = new MarkupLineReader(file);
    }

    /// <summary>An on-demand open hands the pass to the document: later steps are not tied to the open's cancellation.</summary>
    public void ClearCancellation() => _cancellation = default;

    /// <summary>Feeds on until at least <paramref name="blocks"/> blocks are known or the text ends.</summary>
    public void AdvanceUntilKnown(int blocks)
    {
        _cancellation.ThrowIfCancellationRequested();       // also when the blocks asked for are known already
        while (!IsEveryBlockKnown && KnownBlockCount < blocks)
        {
            _cancellation.ThrowIfCancellationRequested();
            bool more = _kind == MarkupSourceKind.PlainText ? FeedPlainTextPiece() : FeedWebArchiveLine();
            if (!more) _final = _kind == MarkupSourceKind.PlainText ? FinishPlainText() : FinishWebArchive();
        }
    }

    /// <summary>The finished index (feeds on to the end first when needed).</summary>
    public MarkupRandomAccessIndex Complete()
    {
        AdvanceUntilKnown(int.MaxValue);
        return _final!;
    }

    // -- groups and anchors -----------------------------------------------------------------

    /// <summary>Folds the scanner's first <paramref name="count"/> blocks into the text-length groups and drops them.</summary>
    private void HandOff(int count)
    {
        if (count <= 0) return;
        AddToGroups(_groupTotals, count);
        _scanner.RemoveFirstBlocks(count);
        _resumes?.RemoveRange(0, count);
    }

    private void AddToGroups(List<int> totals, int count)
    {
        for (int i = 0; i < count; i++)
        {
            int group = (_scanner.BlockIndexBase + i) / WordIndexTextLengthGroups.GroupSize;
            while (totals.Count <= group) totals.Add(0);
            totals[group] += _scanner.BlockTextLengths[i];
        }
    }

    private static MarkupSpanAnchor[] FinishAnchors(List<MarkupSpanAnchor> anchors, int blockCount)
    {
        if (blockCount == 0) return [];
        int count = anchors.Count;
        while (count > 1 && anchors[count - 1].FirstBlock >= blockCount) count--;
        return anchors.GetRange(0, count).ToArray();
    }

    private static int[] FinishGroups(List<int> totals, int blockCount)
    {
        int groups = (blockCount + WordIndexTextLengthGroups.GroupSize - 1) / WordIndexTextLengthGroups.GroupSize;
        var result = new int[groups];
        for (int i = 0; i < groups && i < totals.Count; i++) result[i] = totals[i];
        return result;
    }

    /// <summary>
    /// The compact index of the blocks known so far. Its last span ends where the pass has read to (<see cref="MarkupRandomAccessIndex.TextEnd"/>),
    /// past the start of the block still being read, so the last known block's end is found by the rescan. Never saved.
    /// </summary>
    public MarkupRandomAccessIndex CreatePartialIndex()
    {
        if (_final is not null) return _final;
        int known = KnownBlockCount;
        var groups = new List<int>(_groupTotals);
        AddToGroups(groups, Math.Max(0, known - _scanner.BlockIndexBase));
        WordDocumentHeading[] headings = _scanner.Headings.Where(heading => heading.BlockIndex < known).ToArray();
        var stacks = new List<string>(_scanner.Stacks);     // a rescan interns into the index's list: never the pass's own
        if (_kind == MarkupSourceKind.PlainText)
            return new MarkupRandomAccessIndex(MarkupSourceKind.PlainText, _sourceLength, _sourceTicks, _codePage, _reason, MimeTransferEncoding.Identity,
                                               _done, _done, _isPlainTextRightToLeft, "", _bomLength, _sampleLength, known, FinishAnchors(_anchors, known),
                                               FinishGroups(groups, known), stacks, headings, _toratEmet) { IsPartial = true };
        return new MarkupRandomAccessIndex(MarkupSourceKind.WebArchive, _sourceLength, _sourceTicks, _codePage, _reason, _part!.TransferEncoding,
                                           _decodedTotal, _file.Length, WebArchiveRightToLeft(), WebArchiveStylesheet.Scope(_scanner.StyleText), _startResume, 0,
                                           known, FinishAnchors(_anchors, known), FinishGroups(groups, known), stacks, headings) { IsPartial = true };
    }

    // -- plain text -----------------------------------------------------------------------

    private void AfterPlainTextPiece()
    {
        MarkupSpanAnchor last = _anchors[_anchors.Count - 1];
        if (_done - last.ScanPosition >= MarkupRandomAccessIndex.AnchorSpanBytes && _scanner.IsAtSpanPoint && CurrentBlock > last.FirstBlock)
            _anchors.Add(new MarkupSpanAnchor(_scanner.ExportSpanState(), _done, 0, 0));
        if (_scanner.BlockStartOffsets.Count > HandOffBlocks) HandOff(_scanner.BlockStartOffsets.Count - 1);
        _progress?.Report(new WordDocumentIndexProgress(_done, _file.Length, CurrentBlock + 1));
    }

    private bool FeedPlainTextPiece()
    {
        _file.Position = _done;                                     // the stream may be shared (windows read between two steps)
        int read = MarkupRandomAccessIndexBuilder.ReadFully(_file, _chunk!, MarkupRandomAccessIndexBuilder.ChunkSize);   // full pieces: a rescan feeds the same boundaries
        if (read <= 0) return false;
        _scanner.Feed(_chunk!, 0, read, _done);
        _done += read;
        AfterPlainTextPiece();
        return true;
    }

    private MarkupRandomAccessIndex FinishPlainText()
    {
        _scanner.Finish(_done);
        int blockCount = _scanner.BlockIndexBase + _scanner.BlockStartOffsets.Count;
        AddToGroups(_groupTotals, _scanner.BlockTextLengths.Count);
        return new MarkupRandomAccessIndex(MarkupSourceKind.PlainText, _sourceLength, _sourceTicks, _codePage, _reason, MimeTransferEncoding.Identity,
                                           _done, _done, _isPlainTextRightToLeft, "", _bomLength, _sampleLength, blockCount, FinishAnchors(_anchors, blockCount),
                                           FinishGroups(_groupTotals, blockCount), _scanner.Stacks, _scanner.Headings.ToArray(), _toratEmet);
    }

    // -- web archive ----------------------------------------------------------------------

    private void FinishPending()
    {
        if (_pending is not { } p || _decodedTotal <= p.position) return;
        int local = p.state.Block - _scanner.BlockIndexBase;
        _anchors.Add(new MarkupSpanAnchor(p.state, p.position, _resumeRing![p.position & (MarkupRandomAccessIndexBuilder.ResumeRingSize - 1)], _resumes![local]));
        HandOff(local);
        _pending = null;
    }

    /// <summary>Decodes and feeds one encoded line; false at the end of the HTML part (boundary, &lt;/body&gt; or end of file).</summary>
    private bool FeedWebArchiveLine()
    {
        if (_scanner.IsBodyEnded) return false;
        if (!_lines!.ReadLine(out ReadOnlySpan<byte> line, out long lineOffset)) return false;
        if (_part!.Boundary is not null && WebArchiveContainer.IsBoundaryLine(line, _part.Boundary, out _)) { _encodedEnd = lineOffset; return false; }
        byte[] lineBytes = line.ToArray();
        for (int at = 0; at < lineBytes.Length; at += MarkupRandomAccessIndexBuilder.FeedPieceBytes)
        {
            int count = Math.Min(MarkupRandomAccessIndexBuilder.FeedPieceBytes, lineBytes.Length - at), produced = 0;
            _decoder!.Feed(lineBytes, at, count, lineOffset + at, _decoded!, _resumePiece, ref produced);
            if (produced == 0) continue;
            for (int k = 0; k < produced; k++) _resumeRing![(_decodedTotal + k) & (MarkupRandomAccessIndexBuilder.ResumeRingSize - 1)] = _resumePiece![k];
            if (_scanner.IsBodyStarted && _bodySample!.Count < RightToLeftSampleBytes)
                for (int k = 0; k < produced && _bodySample.Count < RightToLeftSampleBytes; k++) _bodySample.Add(_decoded![k]);
            _scanner.Feed(_decoded!, 0, produced, _decodedTotal);
            _decodedTotal += produced;
            MarkupRandomAccessIndexBuilder.MapResumePoints(_scanner, _resumes!, _resumeRing!, _decodedTotal);
            FinishPending();
            MarkupSpanAnchor last = _anchors[_anchors.Count - 1];
            if (_pending is null && _decodedTotal - last.ScanPosition >= MarkupRandomAccessIndex.WebArchiveAnchorSpanBytes && _scanner.IsAtSpanPoint
                && CurrentBlock > last.FirstBlock)
                _pending = (_scanner.ExportSpanState(), _decodedTotal);
            if (_pending is null && _resumes!.Count > HandOffBlocks) HandOff(Math.Min(_resumes.Count, _scanner.BlockStartOffsets.Count - 1));
        }
        if (_progress is not null && (++_pieces & 255) == 0)
            _progress.Report(new WordDocumentIndexProgress(lineOffset - _part.BodyStart, _file.Length - _part.BodyStart, CurrentBlock + 1));
        return true;
    }

    private bool WebArchiveRightToLeft()
    {
        int sampleLength = _bodySample!.Count & ~(_unitSize - 1);
        return MarkupTextEncoding.ContainsHebrew(_encoding.GetString(_bodySample.ToArray(), 0, sampleLength));
    }

    private MarkupRandomAccessIndex FinishWebArchive()
    {
        if (!_scanner.IsBodyStarted) throw new InvalidDataException("Web archive: the HTML part has no <body>.");
        long textEnd = _scanner.IsBodyEnded ? _scanner.BodyEndTagOffset : _decodedTotal;
        _scanner.Finish(textEnd);
        List<long> resumes = _resumes!;
        while (resumes.Count > _scanner.BlockStartOffsets.Count) resumes.RemoveAt(resumes.Count - 1);
        if (resumes.Count < _scanner.BlockStartOffsets.Count) throw new InvalidDataException("Web archive: block start past the decoded text.");
        int blockCount = _scanner.BlockIndexBase + _scanner.BlockStartOffsets.Count;
        if (_pending is { } open && open.state.Block < blockCount && _decodedTotal > open.position)
            _anchors.Add(new MarkupSpanAnchor(open.state, open.position, _resumeRing![open.position & (MarkupRandomAccessIndexBuilder.ResumeRingSize - 1)],
                                              resumes[open.state.Block - _scanner.BlockIndexBase]));
        AddToGroups(_groupTotals, _scanner.BlockTextLengths.Count);
        _progress?.Report(new WordDocumentIndexProgress(_file.Length - _part!.BodyStart, _file.Length - _part.BodyStart, blockCount));
        return new MarkupRandomAccessIndex(MarkupSourceKind.WebArchive, _sourceLength, _sourceTicks, _codePage, _reason, _part!.TransferEncoding,
                                           textEnd, _encodedEnd, WebArchiveRightToLeft(), WebArchiveStylesheet.Scope(_scanner.StyleText), _startResume, 0,
                                           blockCount, FinishAnchors(_anchors, blockCount), FinishGroups(_groupTotals, blockCount), _scanner.Stacks,
                                           _scanner.Headings.ToArray());
    }
}
