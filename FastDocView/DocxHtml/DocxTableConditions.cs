using System.Xml.Linq;

namespace FastDocView.DocxHtml;

/// <summary>
/// The conditional parts of a table style (w:tblStylePr, ECMA-376 §17.7.6.6 / ST_TblStyleOverrideType) that apply to one
/// cell. The bit values are LibreOffice's CNF_* masks (DomainMapperTableHandler.cxx).
/// </summary>
[Flags]
public enum DocxTableConditions : short
{
    None = 0,
    SouthWestCell = 0x001, SouthEastCell = 0x002, NorthWestCell = 0x004, NorthEastCell = 0x008,
    Band2Horizontal = 0x010, Band1Horizontal = 0x020, Band2Vertical = 0x040, Band1Vertical = 0x080,
    LastColumn = 0x100, FirstColumn = 0x200, LastRow = 0x400, FirstRow = 0x800,
}

public static class DocxTableConditionOrder
{
    /// <summary>
    /// The order the conditional parts are layered in, later ones winning — LibreOffice
    /// TableStyleSheetEntry::GetLocalPropertiesFromMask: bands (horizontal, then vertical), last / first column,
    /// last / first row, then the corner cells.
    /// </summary>
    public static readonly (DocxTableConditions condition, string type)[] Layers =
    [
        (DocxTableConditions.Band2Horizontal, "band2Horz"), (DocxTableConditions.Band1Horizontal, "band1Horz"),
        (DocxTableConditions.Band2Vertical, "band2Vert"), (DocxTableConditions.Band1Vertical, "band1Vert"),
        (DocxTableConditions.LastColumn, "lastCol"), (DocxTableConditions.FirstColumn, "firstCol"),
        (DocxTableConditions.LastRow, "lastRow"), (DocxTableConditions.FirstRow, "firstRow"),
        (DocxTableConditions.SouthWestCell, "swCell"), (DocxTableConditions.SouthEastCell, "seCell"),
        (DocxTableConditions.NorthWestCell, "nwCell"), (DocxTableConditions.NorthEastCell, "neCell"),
    ];

    /// <summary>
    /// w:tblLook → LibreOffice's look bits (DomainMapperTableManager::finishTableLook): 0x020 first row, 0x040 last row,
    /// 0x080 first column, 0x100 last column, 0x200 no horizontal bands, 0x400 no vertical bands. Any explicit attribute
    /// wins over w:val and an absent attribute then counts as on; no tblLook at all = 0x4A0 ([MS-OI29500] errata default).
    /// </summary>
    public static int LookBits(XElement? tableLook)
    {
        if (tableLook is null) return 0x4A0;
        XNamespace w = DocxStyleCascade.W;
        string[] names = ["firstRow", "lastRow", "firstColumn", "lastColumn", "noHBand", "noVBand"];
        int[] bits = [0x020, 0x040, 0x080, 0x100, 0x200, 0x400];
        if (names.Any(name => tableLook.Attribute(w + name) is not null))
        {
            int look = 0;
            for (int i = 0; i < names.Length; i++)
                if (tableLook.Attribute(w + names[i]) is not { } attribute || DocxStyleCascade.IsOn(attribute)) look |= bits[i];
            return look;
        }
        string? value = (string?)tableLook.Attribute(w + "val");
        return value is not null && int.TryParse(value, System.Globalization.NumberStyles.HexNumber, null, out int legacy) ? legacy : 0x4A0;
    }
}
