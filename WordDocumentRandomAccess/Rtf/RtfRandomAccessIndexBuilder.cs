using System.Text;
using WordDocumentRandomAccess.DocxHtml;

namespace WordDocumentRandomAccess.Rtf;

/// <summary>
/// The one streaming pass over an RTF file. Blocks: a paragraph ends at \par (or \sect) outside a table; a table
/// row (\trowd … \cell … \row) is one block, like the .doc reader's rows. A block's range starts right after the
/// previous block's terminator (or after the last header group), and the parser's group-state stack at that point
/// is stored with it (interned), so any block renders from its own bytes. List labels, note numbers and headings
/// are settled here; nothing of the text is kept.
/// </summary>
internal sealed class RtfRandomAccessIndexBuilder : RtfSink
{
    private const int HeadingTextLimit = 200;

    private readonly RtfParser _parser;
    private readonly RtfDocumentTables _tables = new();
    private RtfListNumbering? _lists;
    private readonly IProgress<WordDocumentIndexProgress>? _progress;
    private readonly CancellationToken _cancellation;
    private readonly long _totalBytes;

    private readonly List<long> _starts = new(), _ends = new();
    private readonly List<WordDocumentBlockKind> _kinds = new();
    private readonly List<int> _textLengths = new(), _stateIds = new(), _labelIds = new();
    private readonly List<byte> _listLevels = new();
    private readonly List<byte[]> _states = new();
    private readonly Dictionary<string, int> _stateIdByContent = new();
    private readonly List<string> _labels = new();
    private readonly Dictionary<string, int> _labelIdByText = new();
    private readonly List<WordDocumentHeading> _headings = new();
    private readonly List<RtfNoteReference> _notes = new();

    // the block being read
    private bool _hasStart, _hasContent, _isRow;
    private long _start;
    private int _stateId;
    private int _textLength;
    private readonly StringBuilder _paragraphText = new();
    private readonly StringBuilder _listText = new();
    private bool _hasListText;
    private int _rowCells;
    private int _rightToLeftParagraphs, _paragraphs;

    // notes
    private bool _hasAutoMark;
    private int _footnoteCount, _endnoteCount;

    private RtfRandomAccessIndexBuilder(Stream file, IProgress<WordDocumentIndexProgress>? progress, CancellationToken cancellation)
    {
        _parser = new RtfParser(this, _tables, readsHeaderTables: true);
        _progress = progress;
        _cancellation = cancellation;
        _totalBytes = file.Length;
    }

    public static RtfRandomAccessIndex Build(Stream file, long sourceLength, long sourceLastWriteTicksUtc,
                                             IProgress<WordDocumentIndexProgress>? progress = null, CancellationToken cancellation = default)
    {
        var builder = new RtfRandomAccessIndexBuilder(file, progress, cancellation);
        builder._parser.Run(file, 0, file.Length);
        builder.EndBlock(builder._parser.Position, final: true);
        progress?.Report(new WordDocumentIndexProgress(file.Length, file.Length, builder._starts.Count));
        return builder.CreateIndex(sourceLength, sourceLastWriteTicksUtc);
    }

    private RtfRandomAccessIndex CreateIndex(long sourceLength, long sourceTicks)
    {
        bool isRightToLeft = _tables.IsRightToLeftDocument || (_paragraphs > 0 && _rightToLeftParagraphs * 2 >= _paragraphs);
        return new RtfRandomAccessIndex(sourceLength, sourceTicks, _tables, isRightToLeft, _starts.ToArray(), _ends.ToArray(), _kinds.ToArray(),
                                        _textLengths.ToArray(), _stateIds.ToArray(), _states.ToArray(), _listLevels.ToArray(), _labelIds.ToArray(),
                                        _labels.ToArray(), _headings.ToArray(), _notes.ToArray());
    }

    // ── block boundaries ─────────────────────────────────────────────────────────────

    private void TakeStart()
    {
        _start = _parser.Position;
        byte[] snapshot = _parser.Snapshot();
        string key = Convert.ToBase64String(snapshot);
        if (!_stateIdByContent.TryGetValue(key, out _stateId))
        {
            _stateId = _states.Count;
            _states.Add(snapshot);
            _stateIdByContent[key] = _stateId;
        }
        _hasStart = true;
        _hasContent = false;
        _isRow = false;
        _textLength = 0;
        _rowCells = 0;
        _paragraphText.Clear();
        _hasListText = false;
        _listText.Clear();
    }

    private void EndBlock(long end, bool final = false, string? listLabel = null, int listLevel = -1)
    {
        if (!_hasStart || (final && !_hasContent)) return;
        int blockIndex = _starts.Count;
        _starts.Add(_start);
        _ends.Add(end);
        _kinds.Add(_isRow ? WordDocumentBlockKind.TableRow : WordDocumentBlockKind.Paragraph);
        _textLengths.Add(_textLength);
        _stateIds.Add(_stateId);
        _listLevels.Add(listLabel is null ? (byte)255 : (byte)Math.Max(0, Math.Min(254, listLevel)));
        _labelIds.Add(listLabel is null ? -1 : InternLabel(listLabel));
        if ((blockIndex & 4095) == 0)
        {
            _cancellation.ThrowIfCancellationRequested();
            _progress?.Report(new WordDocumentIndexProgress(end, _totalBytes, blockIndex));
        }
        if (!final) TakeStart();
        else _hasStart = false;
    }

    private int InternLabel(string label)
    {
        if (!_labelIdByText.TryGetValue(label, out int id))
        {
            id = _labels.Count;
            _labels.Add(label);
            _labelIdByText[label] = id;
        }
        return id;
    }

    // ── parser events ───────────────────────────────────────────────────────────────

    public override bool WantsText
    {
        get
        {
            if (_parser.IsInFootnote) return false;
            if (_parser.IsInListText) return true;
            if (_paragraphText.Length >= HeadingTextLimit) return false;
            ref readonly RtfGroupState state = ref _parser.Current;
            return state.OutlineLevel is >= 0 and <= 8 || _tables.OutlineLevelOfStyle(state.Style) >= 0;
        }
    }

    public override void DocumentStarted() => TakeStart();

    public override void GroupClosed(int depth)
    {
        if (depth == 0 && _hasStart && !_hasContent) TakeStart();          // header tables and other root-level groups before the text
    }

    public override void Text(ReadOnlySpan<char> text, in RtfGroupState state)
    {
        if (_parser.IsInListText) { _hasContent = true; _listText.Append(text.ToArray()); return; }
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _textLength += text.Length;
        int room = HeadingTextLimit - _paragraphText.Length;
        if (room > 0) _paragraphText.Append(text.Slice(0, Math.Min(room, text.Length)).ToArray());
    }

    public override void TextLength(int length, in RtfGroupState state)
    {
        if (_parser.IsInFootnote || _parser.IsInListText) return;
        _hasContent = true;
        _textLength += length;
    }

    public override void ListTextStart() { _hasListText = true; _listText.Clear(); }

    public override void Embedded(RtfEmbeddedKind kind, in RtfGroupState state) { if (!_parser.IsInFootnote) { _hasContent = true; _textLength += 8; } }

    public override void Break(RtfWord word, in RtfGroupState state) { if (!_parser.IsInFootnote) _hasContent = true; }

    public override void ParagraphEnd(in RtfGroupState state, bool isSectionEnd)
    {
        if (_parser.IsInFootnote || _parser.IsInListText) return;
        _hasContent = true;
        string? label = CountList(state, out bool hasListParagraph);
        if (state.IsInTable || _rowCells > 0)
        {
            _isRow = true;
            _paragraphText.Clear();
            _hasListText = false;
            return;
        }
        _paragraphs++;
        if (state.Has(RtfStateFlags.RightToLeftParagraph)) _rightToLeftParagraphs++;
        int outline = state.OutlineLevel is >= 0 and <= 8 ? state.OutlineLevel : state.OutlineLevel == 9 ? -1 : _tables.OutlineLevelOfStyle(state.Style);
        if (outline >= 0)
        {
            string text = _paragraphText.ToString().Trim();
            if (text.Length > 0) _headings.Add(new WordDocumentHeading(_starts.Count, outline + 1, text));
        }
        string? storedLabel = _hasListText ? null : label;
        EndBlock(_parser.Position, listLabel: storedLabel, listLevel: hasListParagraph ? state.ListLevel : -1);
    }

    /// <summary>Runs the list counter for a list paragraph; returns its label when the index should store one.</summary>
    private string? CountList(in RtfGroupState state, out bool hasListParagraph)
    {
        hasListParagraph = state.ListOverride > 0;
        if (!hasListParagraph) return null;
        _lists ??= new RtfListNumbering(_tables);
        return _lists.Next(state.ListOverride, state.ListLevel);
    }

    public override void CellEnd(in RtfGroupState state, bool isNested)
    {
        if (_parser.IsInFootnote) return;
        _hasContent = true;
        _isRow = true;
        if (!isNested) _rowCells++;
        CountList(state, out _);
        _paragraphText.Clear();
    }

    public override void RowEnd(in RtfGroupState state, bool isNested)
    {
        if (_parser.IsInFootnote || isNested) return;
        _hasContent = true;
        _isRow = true;
        EndBlock(_parser.Position);
    }

    public override void FootnoteMark(in RtfGroupState state)
    {
        if (!_parser.IsInFootnote) _hasAutoMark = true;
    }

    public override void FootnoteStart() => _hasContent = true;

    public override void FootnoteEnd(bool isEndnote)
    {
        WordDocumentNoteKind kind = isEndnote ? WordDocumentNoteKind.Endnote : WordDocumentNoteKind.Footnote;
        int number = 0;
        string label = "*";
        if (_hasAutoMark)
        {
            number = isEndnote ? ++_endnoteCount : ++_footnoteCount;
            int shown = number - 1 + (isEndnote ? _tables.EndnoteStart : _tables.FootnoteStart);
            label = FormatNoteNumber(shown, isEndnote ? _tables.EndnoteFormat : _tables.FootnoteFormat);
        }
        _hasAutoMark = false;
        _notes.Add(new RtfNoteReference(_starts.Count, kind, number, label));
    }

    /// <summary>\ftnn* formats; Chicago (*, †, ‡, § then doubled) is from memory, not source-checked.</summary>
    internal static string FormatNoteNumber(int number, string format)
    {
        if (format != "chicago") return DocxNumberingDefinitions.FormatNumber(number, format);
        string[] symbols = ["*", "\u2020", "\u2021", "\u00A7"];
        if (number <= 0) return "*";
        return string.Concat(Enumerable.Repeat(symbols[(number - 1) % 4], (number - 1) / 4 + 1));
    }
}
