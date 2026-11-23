using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// Unicode cursive joining (Unicode 16.0 ArabicShaping.txt, generated; Arabic, Syriac, N'Ko, Mandaic, ... in the BMP) and the
/// positional form each character takes - the OpenType 'isol' / 'fina' / 'medi' / 'init' feature a shaper applies to it
/// (Microsoft "Developing OpenType fonts for Arabic script": a character joins the previous one when both sides can join,
/// transparent characters - marks, general category Mn / Me / Cf - are skipped).
/// </summary>
internal static class WordArabicJoining
{
    /// <summary>First, last, joining type (D dual, R right, L left, C join-causing, T transparent); characters not listed are U (non-joining) unless Mn / Me / Cf.</summary>
    private static readonly int[] Ranges =
    [
        0x0620, 0x0620, 'D', 0x0622, 0x0625, 'R', 0x0626, 0x0626, 'D', 0x0627, 0x0627, 'R', 0x0628, 0x0628, 'D', 0x0629, 0x0629, 'R',
        0x062A, 0x062E, 'D', 0x062F, 0x0632, 'R', 0x0633, 0x063F, 'D', 0x0640, 0x0640, 'C', 0x0641, 0x0647, 'D', 0x0648, 0x0648, 'R',
        0x0649, 0x064A, 'D', 0x066E, 0x066F, 'D', 0x0671, 0x0673, 'R', 0x0675, 0x0677, 'R', 0x0678, 0x0687, 'D', 0x0688, 0x0699, 'R',
        0x069A, 0x06BF, 'D', 0x06C0, 0x06C0, 'R', 0x06C1, 0x06C2, 'D', 0x06C3, 0x06CB, 'R', 0x06CC, 0x06CC, 'D', 0x06CD, 0x06CD, 'R',
        0x06CE, 0x06CE, 'D', 0x06CF, 0x06CF, 'R', 0x06D0, 0x06D1, 'D', 0x06D2, 0x06D3, 'R', 0x06D5, 0x06D5, 'R', 0x06EE, 0x06EF, 'R',
        0x06FA, 0x06FC, 'D', 0x06FF, 0x06FF, 'D', 0x070F, 0x070F, 'T', 0x0710, 0x0710, 'R', 0x0712, 0x0714, 'D', 0x0715, 0x0719, 'R',
        0x071A, 0x071D, 'D', 0x071E, 0x071E, 'R', 0x071F, 0x0727, 'D', 0x0728, 0x0728, 'R', 0x0729, 0x0729, 'D', 0x072A, 0x072A, 'R',
        0x072B, 0x072B, 'D', 0x072C, 0x072C, 'R', 0x072D, 0x072E, 'D', 0x072F, 0x072F, 'R', 0x074D, 0x074D, 'R', 0x074E, 0x0758, 'D',
        0x0759, 0x075B, 'R', 0x075C, 0x076A, 'D', 0x076B, 0x076C, 'R', 0x076D, 0x0770, 'D', 0x0771, 0x0771, 'R', 0x0772, 0x0772, 'D',
        0x0773, 0x0774, 'R', 0x0775, 0x0777, 'D', 0x0778, 0x0779, 'R', 0x077A, 0x077F, 'D', 0x07CA, 0x07EA, 'D', 0x07FA, 0x07FA, 'C',
        0x0840, 0x0840, 'R', 0x0841, 0x0845, 'D', 0x0846, 0x0847, 'R', 0x0848, 0x0848, 'D', 0x0849, 0x0849, 'R', 0x084A, 0x0853, 'D',
        0x0854, 0x0854, 'R', 0x0855, 0x0855, 'D', 0x0856, 0x0858, 'R', 0x0860, 0x0860, 'D', 0x0862, 0x0865, 'D', 0x0867, 0x0867, 'R',
        0x0868, 0x0868, 'D', 0x0869, 0x086A, 'R', 0x0870, 0x0882, 'R', 0x0883, 0x0885, 'C', 0x0886, 0x0886, 'D', 0x0889, 0x088D, 'D',
        0x088E, 0x088E, 'R', 0x08A0, 0x08A9, 'D', 0x08AA, 0x08AC, 'R', 0x08AE, 0x08AE, 'R', 0x08AF, 0x08B0, 'D', 0x08B1, 0x08B2, 'R',
        0x08B3, 0x08B8, 'D', 0x08B9, 0x08B9, 'R', 0x08BA, 0x08C8, 'D', 0x1807, 0x1807, 'D', 0x180A, 0x180A, 'C', 0x1820, 0x1878, 'D',
        0x1885, 0x1886, 'T', 0x1887, 0x18A8, 'D', 0x18AA, 0x18AA, 'D', 0x200D, 0x200D, 'C', 0xA840, 0xA871, 'D', 0xA872, 0xA872, 'L',
    ];

    /// <summary>The joining type of a character: 'D', 'R', 'L', 'C', 'T' or 'U'.</summary>
    public static char TypeOf(int codePoint)
    {
        int low = 0, high = Ranges.Length / 3 - 1;
        while (low <= high)
        {
            int middle = (low + high) / 2;
            if (codePoint < Ranges[middle * 3]) high = middle - 1;
            else if (codePoint > Ranges[middle * 3 + 1]) low = middle + 1;
            else return (char)Ranges[middle * 3 + 2];
        }
        if (codePoint <= 0xFFFF)
        {
            UnicodeCategory category = CharUnicodeInfo.GetUnicodeCategory((char)codePoint);
            if (category is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format) return 'T';
        }
        return 'U';
    }

    /// <summary>The feature tag of each character's positional form ("isol", "fina", "medi", "init"; null for a non-joining or transparent one).</summary>
    public static string?[] Forms(IReadOnlyList<int> codePoints)
    {
        var forms = new string?[codePoints.Count];
        var types = new char[codePoints.Count];
        for (int i = 0; i < types.Length; i++) types[i] = TypeOf(codePoints[i]);
        for (int i = 0; i < types.Length; i++)
        {
            char type = types[i];
            if (type is 'U' or 'T') continue;
            int before = i - 1, after = i + 1;
            while (before >= 0 && types[before] == 'T') before--;
            while (after < types.Length && types[after] == 'T') after++;
            // joins the previous character: this one joins on its right (D, R, C) and the previous one on its left (D, L, C)
            bool joinsPrevious = type is 'D' or 'R' or 'C' && before >= 0 && types[before] is 'D' or 'L' or 'C';
            bool joinsNext = type is 'D' or 'L' or 'C' && after < types.Length && types[after] is 'D' or 'R' or 'C';
            if (type == 'C') continue;   // tatweel has one form
            forms[i] = joinsPrevious ? joinsNext ? "medi" : "fina" : joinsNext ? "init" : "isol";
        }
        return forms;
    }
}
