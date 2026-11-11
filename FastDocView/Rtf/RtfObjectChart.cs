using FastDocView.DocxHtml;

namespace FastDocView.Rtf;

/// <summary>
/// The chart of an RTF {\object} group: its {\*\objdata} hex digits ([MS-OLEDS] 2.2.4 ObjectHeader + the compound file, as
/// LibreOffice reads them: sw/source/writerfilter/rtftok/rtfdocumentimpl.cxx:3051-3058 handleEmbeddedObject, :4022
/// ExtractOLE2FromObjdata, filter/source/msfilter/rtfutil.cxx:267-330 - CR / LF skipped, any other non-hex character ends the
/// read as a failure) read by <see cref="ExcelBinaryChart"/>; its size from \objw / \objh (twips) scaled by \objscalex /
/// \objscaley (percent). Only the group at hand is read, when its block is rendered or searched.
/// </summary>
internal static class RtfObjectChart
{
    public readonly record struct Found(DocxChart Chart, double WidthPt, double HeightPt);

    public static Found? Read(byte[] group)
    {
        int width = 0, height = 0, scaleX = 100, scaleY = 100;
        byte[]? data = null;
        int depth = 0;
        for (int i = 0; i < group.Length; i++)
        {
            byte b = group[i];
            if (b == '{') { depth++; continue; }
            if (b == '}') { depth--; continue; }
            if (b != '\\' || i + 1 >= group.Length) continue;
            int start = i + 1, end = start;
            while (end < group.Length && IsLetter(group[end])) end++;
            if (end == start) { i++; continue; }                   // an escaped symbol
            string word = System.Text.Encoding.ASCII.GetString(group, start, end - start);
            int numberEnd = end;
            if (numberEnd < group.Length && group[numberEnd] == '-') numberEnd++;
            while (numberEnd < group.Length && group[numberEnd] is >= (byte)'0' and <= (byte)'9') numberEnd++;
            int.TryParse(System.Text.Encoding.ASCII.GetString(group, end, numberEnd - end), out int parameter);
            i = numberEnd - 1;
            if (depth != 0 && word != "objdata") continue;               // the object's own size words are at its top level
            switch (word)
            {
                case "objw": width = parameter; break;
                case "objh": height = parameter; break;
                case "objscalex": scaleX = parameter; break;
                case "objscaley": scaleY = parameter; break;
                case "objdata": data = HexDigits(group, numberEnd, out i); i--; break;
            }
        }
        if (data is null || width <= 0 || height <= 0) return null;
        if (ExcelBinaryChart.FromObjectData(data) is not { } chart) return null;
        if (scaleX <= 0) scaleX = 100;
        if (scaleY <= 0) scaleY = 100;
        return new Found(chart, width / 20.0 * scaleX / 100.0, height / 20.0 * scaleY / 100.0);
    }

    /// <summary>
    /// Where the first {\pict} body of the object's {\result} starts in the group (just after the "\pict" word): the preview picture
    /// Word saves for the object (LibreOffice rtfdocumentimpl.cxx:3069-3075 uses \result when the object itself is not read);
    /// null when there is none.
    /// </summary>
    public static int? ResultPictureOffset(byte[] group)
    {
        int result = IndexOf(group, "\\result", 0);
        if (result < 0) return null;
        for (int at = IndexOf(group, "\\pict", result); at >= 0; at = IndexOf(group, "\\pict", at + 5))
            if (at + 5 >= group.Length || !IsLetter(group[at + 5])) return at + 5;
        return null;
    }

    private static int IndexOf(byte[] group, string word, int from)
    {
        for (int i = Math.Max(0, from); i + word.Length <= group.Length; i++)
        {
            int k = 0;
            while (k < word.Length && group[i + k] == word[k]) k++;
            if (k == word.Length) return i;
        }
        return -1;
    }

    /// <summary>The bytes of the hex digits from <paramref name="at"/> to the group's closing brace (white space skipped); null on any other character.</summary>
    private static byte[]? HexDigits(byte[] group, int at, out int end)
    {
        var bytes = new List<byte>(Math.Max(0, (group.Length - at) / 2));
        int high = -1;
        for (end = at; end < group.Length; end++)
        {
            byte b = group[end];
            if (b == '}') break;
            if (b is (byte)'\r' or (byte)'\n' or (byte)' ' or (byte)'\t') continue;
            int nibble = b >= '0' && b <= '9' ? b - '0' : b >= 'a' && b <= 'f' ? b - 'a' + 10 : b >= 'A' && b <= 'F' ? b - 'A' + 10 : -1;
            if (nibble < 0) return null;
            if (high < 0) high = nibble;
            else { bytes.Add((byte)(high << 4 | nibble)); high = -1; }
        }
        return bytes.ToArray();
    }

    private static bool IsLetter(byte b) => b is >= (byte)'a' and <= (byte)'z' or >= (byte)'A' and <= (byte)'Z';
}
