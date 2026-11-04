using FastDocView.Fields;

namespace FastDocView.DocBinary;

/// <summary>
/// The field tables of a .doc ([MS-DOC] PlcFld): per story a PLC of the CPs of every field mark (0x13 begin, 0x14
/// separator, 0x15 end) with a 2-byte FLD each. Read at open like the bookmarks (small: 6 bytes per mark), from the
/// FibRgFcLcb97 pairs in LibreOffice's ww8scan.cxx WW8Fib read order: PlcffldMom 16, PlcffldHdr 17, PlcffldFootnote 18,
/// PlcffldEdn 48, PlcffldTxbx 57, PlcffldHdrTxbx 59 (ww8scan.cxx:6042-6048, 6134, 6152, 6156); a story's CPs count from
/// its start (WW8PLCFx_FLD, ww8scan.cxx:3974-4012, reads them per story). The FLD of a begin mark holds the field type, of
/// an end mark the flags: fResultDirty 0x04, fLocked 0x10 (Apache POI FLDAbstractType.java:34-41; LibreOffice reads
/// fLocked the same way, ww8par5.cxx:1628 / 1952).
/// What the reader needs from them: each field's dirty / locked flags (by its begin CP), the fields open where a block
/// starts (a code or a result running in from an earlier block), and the counters of SEQ / AUTONUM / LISTNUM before a CP.
/// </summary>
internal sealed class DocBinaryFields
{
    private readonly int[] _beginCps;            // every field's begin CP (absolute), ascending
    private readonly int[] _separatorCps;        // its separator CP, -1 when it has none
    private readonly int[] _endCps;              // its end CP
    private readonly byte[] _endFlags;           // the end mark's FLD flags
    private readonly int[] _parents;             // the field it is nested in, -1 at the top
    private readonly int _mainTextEnd;

    public static readonly DocBinaryFields Empty = new([], [], [], [], [], 0);

    private DocBinaryFields(int[] beginCps, int[] separatorCps, int[] endCps, byte[] endFlags, int[] parents, int mainTextEnd)
    {
        _beginCps = beginCps;
        _separatorCps = separatorCps;
        _endCps = endCps;
        _endFlags = endFlags;
        _parents = parents;
        _mainTextEnd = mainTextEnd;
    }

    public int Count => _beginCps.Length;

    public static DocBinaryFields Read(Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        var begins = new List<int>();
        var separators = new List<int>();
        var ends = new List<int>();
        var flags = new List<byte>();
        var parents = new List<int>();
        foreach ((int pair, int storyStart) in new[]
                 {
                     (16, 0), (18, fib.FootnoteStoryStart), (17, fib.MainTextCharacterCount + fib.FootnoteCharacterCount),
                     (48, fib.EndnoteStoryStart), (57, fib.TextBoxStoryStart), (59, fib.TextBoxStoryStart + fib.TextBoxCharacterCount),
                 })
        {
            (uint offset, uint length) = fib.FieldTable(pair);
            if (length < 10 || DocBinaryStructureReader.TryRead(tableStream, (offset, length)) is not { } bytes) continue;
            int count = (bytes.Length - 4) / 6;                    // n + 1 CPs of 4 bytes, n FLDs of 2 bytes
            var open = new Stack<int>();
            for (int i = 0; i < count; i++)
            {
                int cp = BitConverter.ToInt32(bytes, i * 4) + storyStart;
                int fld = (count + 1) * 4 + i * 2;
                byte ch = (byte)(bytes[fld] & 0x1F), data = bytes[fld + 1];
                switch (ch)
                {
                    case 0x13:
                        parents.Add(open.Count > 0 ? open.Peek() : -1);
                        open.Push(begins.Count);
                        begins.Add(cp); separators.Add(-1); ends.Add(int.MaxValue); flags.Add(0);
                        break;
                    case 0x14 when open.Count > 0: separators[open.Peek()] = cp; break;
                    case 0x15 when open.Count > 0:
                        int field = open.Pop();
                        ends[field] = cp;
                        flags[field] = data;
                        break;
                }
            }
            // a damaged table: a field never ended counts as ended where it began (else every later block would replay it)
            foreach (int unended in open) ends[unended] = begins[unended];
        }
        // the stories follow one another, so CP order is story order then PLC order; keep it sorted should a damaged file break it
        int[] order = Enumerable.Range(0, begins.Count).OrderBy(i => begins[i]).ToArray();
        if (!order.Select((value, index) => value == index).All(same => same))
        {
            var position = new int[order.Length];
            for (int i = 0; i < order.Length; i++) position[order[i]] = i;
            return new DocBinaryFields(order.Select(i => begins[i]).ToArray(), order.Select(i => separators[i]).ToArray(), order.Select(i => ends[i]).ToArray(),
                                       order.Select(i => flags[i]).ToArray(), order.Select(i => parents[i] < 0 ? -1 : position[parents[i]]).ToArray(), fib.MainTextCharacterCount);
        }
        return new DocBinaryFields(begins.ToArray(), separators.ToArray(), ends.ToArray(), flags.ToArray(), parents.ToArray(), fib.MainTextCharacterCount);
    }

    private int IndexOfBegin(int cp)
    {
        int index = Array.BinarySearch(_beginCps, cp);
        return index >= 0 ? index : -1;
    }

    /// <summary>The dirty / locked / private-result flags of the field that begins at <paramref name="beginCp"/> (all false when unknown; fPrivateResult 0x20, POI FLDAbstractType.java:39).</summary>
    public (bool isDirty, bool isLocked, bool isPrivate) FlagsAt(int beginCp)
    {
        int index = IndexOfBegin(beginCp);
        if (index < 0) return (false, false, false);
        byte flags = _endFlags[index];
        return ((flags & 0x04) != 0, (flags & 0x10) != 0, (flags & 0x20) != 0);
    }

    /// <summary>The fields open at <paramref name="cp"/> (begun before it, ending at or after it), outermost first: begin and separator CPs.</summary>
    public List<(int beginCp, int separatorCp)> OpenAt(int cp)
    {
        var open = new List<(int, int)>();
        int low = 0, high = _beginCps.Length;
        while (low < high) { int middle = (low + high) / 2; if (_beginCps[middle] < cp) low = middle + 1; else high = middle; }
        int field = low - 1;                                       // the last field begun before cp
        while (field >= 0 && _endCps[field] < cp) field = _parents[field];
        for (; field >= 0; field = _parents[field]) open.Insert(0, (_beginCps[field], _separatorCps[field]));
        return open;
    }

    // ---- counters (SEQ / AUTONUM / LISTNUM), computed once on first need: only a window holding a counted field Word
    //      saved without a result (or dirty) asks, and Word writes results into .doc files, so this is rare

    private List<(int cp, WordFieldInstruction instruction)>? _countedFields;
    private (int cp, WordFieldSequenceCounters counters)? _lastCounters;
    private (int cp, int level)[]? _headings;

    /// <summary>
    /// The counters as they stand at <paramref name="cp"/> in the main text: every counted field before it that is not inside
    /// another field's code advances them (the .docx index pass's rule); headings reset SEQ \s.
    /// </summary>
    public WordFieldSequenceCounters CountersBefore(int cp, Func<int, int, string> readText, Func<IEnumerable<(int cp, int level)>> headings)
    {
        if (cp > _mainTextEnd) return new WordFieldSequenceCounters();
        _countedFields ??= CountedFields(readText);
        WordFieldSequenceCounters counters;
        int fromCp;
        if (_lastCounters is { } last && last.cp <= cp) { counters = last.counters.Clone(); fromCp = last.cp; }
        else { counters = new WordFieldSequenceCounters(); fromCp = 0; }
        // the heading positions once, sorted; each call starts at the first one at or after its range (a binary search)
        // (kept once there are any: during the index pass the list is not there yet)
        var headingList = _headings ?? headings().OrderBy(h => h.cp).ToArray();
        if (headingList.Length > 0) _headings = headingList;
        int nextHeading = 0;
        { int low = 0, high = headingList.Length; while (low < high) { int middle = (low + high) / 2; if (headingList[middle].cp < fromCp) low = middle + 1; else high = middle; } nextHeading = low; }
        foreach ((int fieldCp, WordFieldInstruction instruction) in _countedFields)
        {
            if (fieldCp < fromCp) continue;
            if (fieldCp >= cp) break;
            while (nextHeading < headingList.Length && headingList[nextHeading].cp <= fieldCp) counters.Heading(headingList[nextHeading++].level);
            WordFieldEvaluator.Count(instruction, counters);
        }
        while (nextHeading < headingList.Length && headingList[nextHeading].cp < cp) counters.Heading(headingList[nextHeading++].level);
        _lastCounters = (cp, counters.Clone());
        return counters;
    }

    private List<(int, WordFieldInstruction)> CountedFields(Func<int, int, string> readText)
    {
        var counted = new List<(int, WordFieldInstruction)>();
        for (int i = 0; i < _beginCps.Length && _beginCps[i] < _mainTextEnd; i++)
        {
            if (InsideCode(i)) continue;
            int codeEnd = _separatorCps[i] >= 0 ? _separatorCps[i] : Math.Min(_endCps[i], _mainTextEnd);
            // the name first (a short read); the whole code only for a counted field
            string head = readText(_beginCps[i] + 1, Math.Min(codeEnd, _beginCps[i] + 16)).TrimStart();
            int nameEnd = 0;
            while (nameEnd < head.Length && char.IsLetter(head[nameEnd])) nameEnd++;
            if (!WordFieldEvaluator.IsCounted(head.Substring(0, nameEnd).ToUpperInvariant())) continue;
            string code = new string(readText(_beginCps[i] + 1, codeEnd).Select(c => c < ' ' ? ' ' : c).ToArray());
            counted.Add((_beginCps[i], WordFieldInstruction.Parse(code)));
        }
        return counted;
    }

    private List<(int cp, string code)>? _indexEntries;

    /// <summary>The XE / TA codes of the main text (read once, on first need: a result-less INDEX / TOA in a window).</summary>
    public List<(int cp, string code)> IndexEntries(Func<int, int, string> readText)
    {
        if (_indexEntries is not null) return _indexEntries;
        var entries = new List<(int, string)>();
        for (int i = 0; i < _beginCps.Length && _beginCps[i] < _mainTextEnd; i++)
        {
            if (InsideCode(i)) continue;
            int codeEnd = _separatorCps[i] >= 0 ? _separatorCps[i] : Math.Min(_endCps[i], _mainTextEnd);
            string head = readText(_beginCps[i] + 1, Math.Min(codeEnd, _beginCps[i] + 8)).TrimStart();
            if (!head.StartsWith("XE", StringComparison.OrdinalIgnoreCase) && !head.StartsWith("TA", StringComparison.OrdinalIgnoreCase)) continue;
            string code = new string(readText(_beginCps[i] + 1, codeEnd).Select(c => c < ' ' ? ' ' : c).ToArray());
            if (WordFieldInstruction.Parse(code).Name is "XE" or "TA") entries.Add((_beginCps[i], code));
        }
        return _indexEntries = entries;
    }

    /// <summary>Whether a field sits in the code of a field around it.</summary>
    private bool InsideCode(int field)
    {
        for (int parent = _parents[field]; parent >= 0; field = parent, parent = _parents[parent])
            if (_separatorCps[parent] < 0 || _beginCps[field] < _separatorCps[parent]) return true;
        return false;
    }
}
