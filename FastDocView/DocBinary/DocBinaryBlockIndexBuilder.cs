namespace FastDocView.DocBinary;

/// <summary>
/// The .doc index pass, resumable: the block scanner is fed the main text chunk by chunk (the very chunks a pass in one go
/// reads, <see cref="DocBinaryPieceTable.CharacterChunks"/>) and can pause between chunks. An on-demand first open advances
/// it only until the first blocks are known and reads windows from <see cref="CreatePartialIndex"/>; later calls scan on, and
/// <see cref="Complete"/> gives the same index <see cref="DocBinaryBlockIndex.Build"/> does (Build runs this to the end).
/// </summary>
internal sealed class DocBinaryBlockIndexBuilder
{
    private readonly DocBinaryBlockScanner _scanner;
    private readonly IEnumerator<(char[] characters, int count, int chunkCp)> _chunks;
    private readonly List<int> _anchorBlocks = new();
    private readonly List<byte[]> _anchorStates = new();
    private readonly IReadOnlyList<int> _noteReferenceCps, _sectionStartCps;
    private readonly int _mainTextCharacterCount;
    private readonly long _sourceLength, _sourceLastWriteTicksUtc;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private CancellationToken _cancellation;

    public DocBinaryBlockIndexBuilder(DocBinaryPieceTable pieces, DocBinaryParagraphHtmlRenderer renderer, DocBinaryStyleSheet styles,
                                      DocBinaryListDefinitions lists, IReadOnlyList<int> noteReferenceCps, IReadOnlyList<int> sectionStartCps,
                                      int mainTextCharacterCount, long sourceLength, long sourceLastWriteTicksUtc,
                                      IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        _noteReferenceCps = noteReferenceCps;
        _sectionStartCps = sectionStartCps;
        _mainTextCharacterCount = mainTextCharacterCount;
        _sourceLength = sourceLength;
        _sourceLastWriteTicksUtc = sourceLastWriteTicksUtc;
        _progress = progress;
        _cancellation = cancellation;
        _scanner = new DocBinaryBlockScanner(renderer, styles, lists) { CaptureStateAtOrAfterCp = 0 };
        _scanner.BlockAdded += (blockIndex, startCp) =>
        {
            if (_scanner.StateOfLastBlock is null && blockIndex != 0) return;
            _anchorBlocks.Add(blockIndex);
            _anchorStates.Add(_scanner.StateOfLastBlock ?? []);
            _scanner.CaptureStateAtOrAfterCp = startCp + DocBinaryBlockIndex.AnchorSpanCharacters;
        };
        _chunks = pieces.CharacterChunks(0, mainTextCharacterCount).GetEnumerator();
    }

    public bool IsEveryBlockKnown { get; private set; }
    public int KnownBlockCount => _scanner.BlockCount;

    /// <summary>The open has returned: later steps run under the host's own calls, not the open's cancellation.</summary>
    public void ClearCancellation() => _cancellation = default;

    /// <summary>Scans on until at least <paramref name="blockCount"/> blocks are known or the text ends.</summary>
    public void AdvanceUntilKnown(int blockCount)
    {
        while (!IsEveryBlockKnown && _scanner.BlockCount < blockCount)
        {
            _cancellation.ThrowIfCancellationRequested();
            if (!_chunks.MoveNext())
            {
                _scanner.Finish(_mainTextCharacterCount);
                _chunks.Dispose();
                IsEveryBlockKnown = true;
                break;
            }
            (char[] characters, int count, int chunkCp) = _chunks.Current;
            _scanner.Feed(characters, count, chunkCp);
            _progress?.Report(new WordDocumentIndexProgress(chunkCp + count, _mainTextCharacterCount, _scanner.BlockCount));
        }
    }

    /// <summary>The index of the blocks known so far (at least one: call after <see cref="AdvanceUntilKnown"/>).</summary>
    public DocBinaryBlockIndex CreatePartialIndex()
        => IsEveryBlockKnown ? Complete()
         : DocBinaryBlockIndex.FromScan(_scanner, _anchorBlocks, _anchorStates, _scanner.AddedBlocksEndCp, _noteReferenceCps, _sectionStartCps,
                                        _sourceLength, _sourceLastWriteTicksUtc, isComplete: false);

    /// <summary>Scans to the end of the text and returns the finished index.</summary>
    public DocBinaryBlockIndex Complete()
    {
        AdvanceUntilKnown(int.MaxValue);
        return DocBinaryBlockIndex.FromScan(_scanner, _anchorBlocks, _anchorStates, _mainTextCharacterCount, _noteReferenceCps, _sectionStartCps,
                                            _sourceLength, _sourceLastWriteTicksUtc, isComplete: true);
    }
}
