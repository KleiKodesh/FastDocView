using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// The colour a patterned shading paints, format-neutral: the pattern's share of the pattern colour (auto = black) mixed over
/// the fill (auto = white), per channel (pattern * share + fill * (1000 - share)) / 1000 — the rule the .docx w:shd path uses
/// (DocxStyleCascade.FillHex, after LibreOffice's CellColorHandler). A result of white is no background.
/// </summary>
public static class WordShadingColor
{
    /// <param name="patternHex">"RRGGBB" of the pattern colour, null = auto (black).</param>
    /// <param name="fillHex">"RRGGBB" of the fill, null = auto (white).</param>
    /// <param name="shareThousandths">The pattern's share, 0 (fill only) to 1000 (pattern colour only).</param>
    public static string? Mix(string? patternHex, string? fillHex, int shareThousandths)
    {
        int share = Math.Max(0, Math.Min(1000, shareThousandths));
        string? result;
        if (share == 0) result = fillHex;
        else
        {
            int foreground = Rgb(patternHex ?? "000000");
            int background = Rgb(fillHex ?? "FFFFFF");
            int Channel(int shift) => (((foreground >> shift) & 0xFF) * share + ((background >> shift) & 0xFF) * (1000 - share)) / 1000;
            result = (Channel(16) << 16 | Channel(8) << 8 | Channel(0)).ToString("X6", CultureInfo.InvariantCulture);
        }
        return result is not null && !result.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase) ? result : null;
    }

    private static int Rgb(string hex) => int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int rgb) ? rgb : 0;
}
