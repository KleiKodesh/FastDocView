namespace WordDocumentRandomAccess.DocBinary;

/// <summary>
/// Reads a structure the FIB locates by (fc, lcb) — a PLC, the style sheet, the font table, the
/// piece table — after checking that the range lies inside its stream. A damaged FIB can name any
/// offset and length; allocating or reading such a range unchecked is what threw raw
/// OverflowException / IndexOutOfRangeException on fuzzed files. LibreOffice does the same check
/// (sw/source/filter/ww8/ww8scan.cxx ReadPLCF: "rSt.remainingSize() >= nPLCF", otherwise the PLCF
/// is made empty) — so an optional structure that does not fit is treated as absent.
/// </summary>
internal static class DocBinaryStructureReader
{
    /// <summary>The structure's bytes, or null when its range does not lie inside the stream.</summary>
    public static byte[]? TryRead(Stream stream, (uint offset, uint length) location)
    {
        if (!Fits(stream, location)) return null;
        var bytes = new byte[location.length];
        stream.Position = location.offset;
        int read = 0;
        while (read < bytes.Length)
        {
            int n = stream.Read(bytes, read, bytes.Length - read);
            if (n <= 0) return null;                    // the stream is shorter than its directory entry says
            read += n;
        }
        return bytes;
    }

    public static bool Fits(Stream stream, (uint offset, uint length) location)
        => location.offset <= stream.Length && location.length <= stream.Length - location.offset && location.length <= int.MaxValue;

    /// <summary>
    /// The number of leading entries of a PLC's CP/FC array that are in ascending order — LibreOffice's
    /// TruncToSortedRange (ww8scan.cxx: "if (pPLCF_PosArray[nI] > pPLCF_PosArray[nI+1]) nIMax = nI;"),
    /// applied to the piece table and every PLC it reads.
    /// Returns the entry count to keep (positions 0..count are then sorted).
    /// </summary>
    public static int SortedEntryCount(Func<int, long> positionAt, int entryCount)
    {
        for (int i = 0; i < entryCount; i++)
            if (positionAt(i) > positionAt(i + 1)) return i;
        return entryCount;
    }
}
