namespace FastDocView.DocBinary;

/// <summary>
/// The main text's text-box stories of a .doc, as LibreOffice finds a shape's text (sw/source/filter/ww8/ww8graf.cxx:852-945
/// GetTxbxTextSttEndCp, read 2026-11-03):
/// - The shape's lTxid (FOPT property 128, include/svx/msdffdef.hxx:107) = story number (high 16 bits, 1-based) and the box's
///   place in a linked chain (low 16 bits; msdffimp.cxx:5560-5561 nTxBxS / nSequence).
/// - PlcftxbxTxt (FibRgFcLcb97 index 56, ww8scan.cxx:6150 read order): CPs and 22-byte FTXBXS (ww8struc.hxx:967-975
///   WW8_TXBXS; ww8scan.cxx:1866 nLenTxBxS); entry nTxBxS - 1, stepping on while fReusable is set; the story runs to the next
///   CP. In a chain, PlcftxbxBkd (index 75, ww8scan.cxx:6215; a PLC of 6-byte Tbkd, [MS-DOC] 2.8.30 PlcfTxbxBkd / 2.9.312 Tbkd)
///   splits it: from the break entry at the story start, nSequence
///   entries on, to the next one. The last character (the story's mark) is not shown (rEndCp -= 1).
/// - Story CPs count from the text-box story start: after the main text, footnotes, headers, comments and endnotes
///   (ww8scan.cxx:5730-5760 GetBaseCp MAN_TXBX).
/// The two tables are read on the first text box (CPs only: 4 bytes a story / break).
/// </summary>
internal sealed class DocBinaryTextBoxStories
{
    private const int StoryEntrySize = 22, BreakEntrySize = 6;
    private readonly Stream _tableStream;
    private readonly (uint offset, uint length) _stories, _breaks;
    private readonly int _storyStart, _storyLength;
    private (int[] cps, bool[] reusable)? _storyTable;
    private int[]? _breakCps;

    public DocBinaryTextBoxStories(Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        _tableStream = tableStream;
        _stories = fib.TextBoxTexts;
        _breaks = fib.TextBoxBreaks;
        _storyStart = fib.TextBoxStoryStart;
        _storyLength = fib.TextBoxCharacterCount;
    }

    /// <summary>The document CP range of the text an lTxid shows; null when it names no story.</summary>
    public (int start, int end)? RangeOf(uint textId)
    {
        int story = (int)(textId >> 16), sequence = (int)(textId & 0xFFFF);
        (int[] cps, bool[] reusable) = _storyTable ??= ReadStories();
        if (story < 1 || story >= cps.Length) return null;
        int index = story - 1;
        while (index < reusable.Length && reusable[index]) index++;
        if (index + 1 >= cps.Length) return null;
        int start = cps[index], end = cps[index + 1];
        if (sequence > 0)
        {
            int[] breaks = _breakCps ??= ReadCps(_breaks, BreakEntrySize);
            int first = Array.IndexOf(breaks, start);
            if (first < 0 || first + sequence >= breaks.Length) return null;
            int partStart = breaks[first + sequence];
            if (partStart < start) return null;
            if (partStart >= end) return (_storyStart + partStart, _storyStart + partStart);
            int partEnd = first + sequence + 1 < breaks.Length ? breaks[first + sequence + 1] : end;
            start = partStart; end = Math.Min(end, partEnd);
        }
        end -= 1;                                           // the story's last mark is not shown
        if (start < 0 || end < start || end > _storyLength) return null;
        return (_storyStart + start, _storyStart + end);
    }

    private (int[] cps, bool[] reusable) ReadStories()
    {
        int[] cps = ReadCps(_stories, StoryEntrySize);
        var reusable = new bool[Math.Max(0, cps.Length - 1)];
        if (reusable.Length > 0)
        {
            byte[] entries = DocBinaryPictures.Read(_tableStream, _stories.offset + (long)cps.Length * 4, reusable.Length * StoryEntrySize);
            for (int i = 0; i < reusable.Length && (i + 1) * StoryEntrySize <= entries.Length; i++)
                reusable[i] = BitConverter.ToUInt16(entries, i * StoryEntrySize + 8) != 0;     // after cTxbx_iNextReuse, cReusable
        }
        return (cps, reusable);
    }

    private int[] ReadCps((uint offset, uint length) plc, int entrySize)
    {
        if (plc.length < 8 || plc.offset + (long)plc.length > _tableStream.Length) return [];
        int count = (int)((plc.length - 4) / (4 + entrySize));
        byte[] bytes = DocBinaryPictures.Read(_tableStream, plc.offset, (count + 1) * 4);
        if (bytes.Length < (count + 1) * 4) return [];
        var cps = new int[count + 1];
        for (int i = 0; i <= count; i++) cps[i] = BitConverter.ToInt32(bytes, i * 4);
        return cps;
    }
}
