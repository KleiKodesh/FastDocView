using System.Text;

namespace FastDocView.DocBinary;

/// <summary>
/// The bookmarks of a .doc: names from SttbfBkmk and start CPs from PlcfBkf, both in the table stream (FibRgFcLcb97
/// pairs 21 and 22, the order LibreOffice's ww8scan.cxx WW8Fib reads them: ... PlcffldMcr, Sttbfbkmk, Plcfbkf,
/// Plcfbkl). As LibreOffice's WW8PLCFx_Book reads them: PlcfBkf is a PLC with 4-byte FBKF entries (n + 1 CPs, then n
/// FBKFs: n = (lcb - 4) / 8), the i-th name belongs to the i-th Bkf entry, and the count used is the smaller of the two;
/// the STTB is extended (first uint16 0xFFFF: UTF-16 strings, uint16 count, uint16 cbExtra, then per string a uint16
/// character count and the characters - ww8scan.cxx WW8ReadSTTBF), else 8-bit strings with a byte length. Only the
/// start CP matters here (PlcfBkl gives the end). The tables are small (one name and 8 bytes per bookmark): read at open.
/// </summary>
public sealed class DocBinaryBookmarks
{
    public static readonly DocBinaryBookmarks Empty = new([], []);

    /// <summary>Names and start CPs, sorted by CP (PLC order).</summary>
    public string[] Names { get; }
    public int[] StartCps { get; }
    private Dictionary<string, int>? _byName;

    private DocBinaryBookmarks(string[] names, int[] startCps)
    {
        Names = names;
        StartCps = startCps;
    }

    public static DocBinaryBookmarks Read(Stream tableStream, DocBinaryFileInformationBlock fib)
    {
        if (fib.BookmarkNames.length < 6 || fib.BookmarkStarts.length < 12) return Empty;
        if (DocBinaryStructureReader.TryRead(tableStream, fib.BookmarkNames) is not { } nameBytes
            || DocBinaryStructureReader.TryRead(tableStream, fib.BookmarkStarts) is not { } startBytes) return Empty;
        List<string> names = ReadStringTable(nameBytes);
        int count = Math.Min(names.Count, (startBytes.Length - 4) / 8);
        var resultNames = new List<string>(count);
        var resultCps = new List<int>(count);
        for (int i = 0; i < count; i++)
        {
            int cp = BitConverter.ToInt32(startBytes, i * 4);
            // _GoBack is Word's "last edit position", not a target (same rule as the .docx reader)
            if (cp < 0 || names[i].Length == 0 || names[i].StartsWith("_GoBack", StringComparison.Ordinal)) continue;
            resultNames.Add(names[i]);
            resultCps.Add(cp);
        }
        // PLC CPs are ascending; keep the order stable by CP should a damaged file break it
        int[] cps = resultCps.ToArray();
        string[] sortedNames = resultNames.ToArray();
        Array.Sort(cps, sortedNames);
        return new DocBinaryBookmarks(sortedNames, cps);
    }

    /// <summary>The start CP of the first bookmark called <paramref name="name"/> (PLC order), or null.</summary>
    public int? StartOf(string name)
    {
        if (_byName is null)
        {
            var byName = new Dictionary<string, int>(Names.Length, StringComparer.Ordinal);
            for (int i = 0; i < Names.Length; i++)
                if (!byName.ContainsKey(Names[i]) || StartCps[i] < byName[Names[i]]) byName[Names[i]] = StartCps[i];
            _byName = byName;
        }
        return _byName.TryGetValue(name, out int cp) ? cp : null;
    }

    /// <summary>The index of the first bookmark starting at or after <paramref name="cp"/>.</summary>
    public int FirstAtOrAfter(int cp)
    {
        int low = 0, high = StartCps.Length;
        while (low < high)
        {
            int middle = (low + high) / 2;
            if (StartCps[middle] < cp) low = middle + 1; else high = middle;
        }
        return low;
    }

    private static List<string> ReadStringTable(byte[] bytes)
    {
        var strings = new List<string>();
        if (bytes.Length < 4) return strings;
        int at = 0;
        ushort first = BitConverter.ToUInt16(bytes, at); at += 2;
        bool isUnicode = first == 0xFFFF;
        int count = isUnicode ? BitConverter.ToUInt16(bytes, at) : first;
        if (isUnicode) at += 2;
        if (at + 2 > bytes.Length) return strings;
        int extra = BitConverter.ToUInt16(bytes, at); at += 2;
        for (int i = 0; i < count; i++)
        {
            if (isUnicode)
            {
                if (at + 2 > bytes.Length) break;
                int characters = BitConverter.ToUInt16(bytes, at); at += 2;
                if (at + characters * 2 > bytes.Length) break;
                strings.Add(Encoding.Unicode.GetString(bytes, at, characters * 2));
                at += characters * 2;
            }
            else
            {
                if (at + 1 > bytes.Length) break;
                int length = bytes[at++];
                if (at + length > bytes.Length) break;
                var text = new StringBuilder(length);
                for (int c = 0; c < length; c++) text.Append(DocBinaryTextDecoder.DecodeCompressed(bytes[at + c]));   // 8-bit: cp1252 as the text is
                strings.Add(text.ToString());
                at += length;
            }
            at += extra;
        }
        return strings;
    }
}
