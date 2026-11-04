namespace FastDocView.DocxHtml;

/// <summary>
/// DrawingML a:pattFill presets as LibreOffice draws them: a hatch (oox/inc/drawingml/hatchmap.hxx createHatch, its own
/// "best-effort mapping"): 1, 2 or 3 line sets (HatchStyle SINGLE / DOUBLE / TRIPLE: drawinglayer fillhatchprimitive2d.cxx:62-88, the second set at -90 degrees, the
/// third at -45), line spacing in hundredths of a millimetre, angle in tenths of a degree. Hairlines in the foreground colour.
/// </summary>
internal static class DocxPatternHatch
{
    public static (double DistancePt, double AngleDegrees, int Sets)? Of(string preset)
        => Hatches.TryGetValue(preset, out var hatch) ? (hatch.Distance * 72.0 / 2540, hatch.Angle / 10.0, hatch.Style) : null;

    private static readonly Dictionary<string, (int Style, int Distance, int Angle)> Hatches = new(StringComparer.Ordinal)
    {
        ["pct5"] = (1, 250, 450),
        ["pct10"] = (1, 200, 450),
        ["pct20"] = (1, 150, 450),
        ["dashUpDiag"] = (1, 150, 450),
        ["pct25"] = (2, 200, 450),
        ["pct30"] = (2, 175, 450),
        ["pct40"] = (2, 150, 450),
        ["weave"] = (2, 150, 450),
        ["pct50"] = (2, 125, 450),
        ["pct60"] = (3, 150, 450),
        ["pct70"] = (3, 125, 450),
        ["pct75"] = (3, 100, 450),
        ["pct80"] = (3, 75, 450),
        ["pct90"] = (3, 50, 450),
        ["horz"] = (1, 100, 0),
        ["wave"] = (1, 100, 0),
        ["vert"] = (1, 100, 900),
        ["ltHorz"] = (1, 50, 0),
        ["narHorz"] = (1, 50, 0),
        ["ltVert"] = (1, 50, 900),
        ["narVert"] = (1, 50, 900),
        ["dkHorz"] = (1, 25, 0),
        ["dkVert"] = (1, 25, 900),
        ["dashHorz"] = (1, 150, 0),
        ["dashVert"] = (1, 150, 900),
        ["cross"] = (2, 100, 0),
        ["lgGrid"] = (2, 100, 0),
        ["dnDiag"] = (1, 100, 1350),
        ["wdDnDiag"] = (1, 100, 1350),
        ["upDiag"] = (1, 100, 450),
        ["wdUpDiag"] = (1, 100, 450),
        ["ltDnDiag"] = (1, 50, 1350),
        ["dkDnDiag"] = (1, 50, 1350),
        ["ltUpDiag"] = (1, 50, 450),
        ["dkUpDiag"] = (1, 50, 450),
        ["dashDnDiag"] = (1, 150, 1350),
        ["diagCross"] = (2, 100, 450),
        ["smCheck"] = (2, 50, 450),
        ["lgCheck"] = (2, 100, 450),
        ["smGrid"] = (2, 50, 0),
        ["dotGrid"] = (2, 400, 0),
        ["smConfetti"] = (1, 200, 600),
        ["lgConfetti"] = (1, 100, 600),
        ["horzBrick"] = (2, 300, 0),
        ["diagBrick"] = (2, 300, 450),
        ["dotDmnd"] = (2, 300, 450),
        ["solidDmnd"] = (2, 100, 450),
        ["openDmnd"] = (2, 100, 450),
        ["plaid"] = (3, 200, 900),
        ["sphere"] = (3, 100, 0),
        ["divot"] = (3, 400, 450),
        ["shingle"] = (1, 200, 1350),
        ["trellis"] = (2, 75, 450),
        ["zigZag"] = (1, 75, 0),
    };
}
