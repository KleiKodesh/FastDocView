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

    /// <summary>
    /// The cell's TextFlow ([MS-DOC] 2.9.336: 0 lrtb, 1 tbrl, 3 btlr, 4 lrtbv, 5 tbrlv): sprmTTextFlow when the row has one for this
    /// cell, else tcgrf bits 2-4 (fVertical 0x4, fBackward 0x8, fRotateFont 0x10). LibreOffice ww8par2.cxx ReadDef :1190-1205 reads the
    /// TC bits only when no sprmTTextFlow set the cell (ProcessDirection :1408-1423).
    /// </summary>
    public int? TextFlowOverride { get; set; }
    public int TextFlow => TextFlowOverride ?? ((Flags >> 2) & 0x7);

    /// <summary>The TextFlow as the docx w:textDirection name, so the cell is drawn by the same CSS mapping (null = horizontal).</summary>
    public string? TextDirectionName => TextFlow switch { 1 => "tbRl", 3 => "btLr", 5 => "tbRlV", _ => null };
}
