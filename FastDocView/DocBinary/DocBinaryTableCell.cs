namespace FastDocView.DocBinary;

/// <summary>
/// One cell of a .doc row definition: the TC of sprmTDefTable ([MS-DOC] 2.9.316 TC80: tcgrf, wWidth, four Brc80) with the
/// borders sprmTSetBrc80 / sprmTSetBrc put on it. Read as LibreOffice reads it (ww8par2.cxx WW8TabBandDesc::ReadDef, :1081-1187;
/// bits in ww8struc.hxx WW8_TCell, :526-555): 0x0001 fFirstMerged, 0x0002 fMerged, 0x0020 fVertMerge, 0x0040 fVertRestart,
/// 0x0180 vertAlign (0 top, 1 centre, 2 bottom).
/// </summary>
public sealed class DocBinaryTableCell
{
    public const int Top = 0, Left = 1, Bottom = 2, Right = 3;

    public ushort Flags { get; set; }

    /// <summary>The cell's own borders (top, left, bottom, right) as CSS; null where the cell sets none (brcType 0: the table's border is used).</summary>
    public string?[] Borders { get; } = new string?[4];

    public bool IsFirstMerged => (Flags & 0x0001) != 0;
    public bool IsMerged => (Flags & 0x0002) != 0;
    public bool IsVerticallyMerged => (Flags & 0x0020) != 0;
    public bool IsVerticalRestart => (Flags & 0x0040) != 0;
    /// <summary>A cell merged into the one above it (fVertMerge without fVertRestart): its content shows in the first cell of the merge.</summary>
    public bool ContinuesAbove => IsVerticallyMerged && !IsVerticalRestart;
    /// <summary>Merged into the cell before it (fMerged without fFirstMerged): the first cell spans it.</summary>
    public bool ContinuesBefore => IsMerged && !IsFirstMerged;
    public int VerticalAlignment => (Flags >> 7) & 0x3;
}
