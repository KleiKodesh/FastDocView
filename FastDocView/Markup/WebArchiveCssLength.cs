using System.Globalization;
using System.Text;

namespace FastDocView.Markup;

/// <summary>
/// Keeps every length a web archive's CSS sets (sheet rules and inline style="", both through
/// <see cref="WebArchiveStylesheet"/>'s one declaration check) inside sane bounds, so no margin, padding, indent, size or
/// spacing can move text out of the document area or over the host page: each number-with-unit token is converted to points
/// (em / ex / ch at a 12 pt font) and clamped to its property's range; a token already in range is kept as written. Keywords
/// (auto, normal, bold...) pass; a function such as calc(), an exponent or an unknown unit drops the declaration, since it cannot
/// be bounded. Relative font sizes are held to 0.1-4 em / 10-400 % (they compound down nested elements).
/// </summary>
internal static class WebArchiveCssLength
{
    private readonly record struct Bounds(double MinimumPoints, double MaximumPoints, double MinimumPercent, double MaximumPercent,
                                          double MinimumNumber = 0, double MaximumNumber = 0);

    private static readonly Bounds Spacing = new(0, 720, 0, 50);                 // margins, padding: never negative, at most 10 in
    private static readonly Bounds Indent = new(-108, 720, -50, 50);             // a hanging indent, at most 1.5 in out
    private static readonly Bounds Box = new(0, 1440, 0, 100);                   // width / height: at most 20 in, a full line
    private static readonly Bounds FontSize = new(1, 400, 10, 1000);
    private static readonly Bounds LineHeight = new(0, 400, 0, 1000, 0, 10);     // a bare number is a multiplier
    private static readonly Bounds LetterSpacing = new(-20, 100, -50, 100);
    private static readonly Bounds Line = new(0, 72, 0, 0);                      // border widths, rules, gaps
    private static readonly Bounds Offset = new(-200, 200, -100, 100);           // vertical-align, underline offset, shadows

    private static Bounds? BoundsOf(string property) => property switch
    {
        _ when property.StartsWith("margin", StringComparison.Ordinal) || property.StartsWith("padding", StringComparison.Ordinal) => Spacing,
        "text-indent" => Indent,
        "width" or "height" or "min-width" or "min-height" or "max-width" or "max-height" => Box,
        "font-size" or "font" => FontSize,
        "line-height" => LineHeight,
        "letter-spacing" or "word-spacing" => LetterSpacing,
        "vertical-align" or "text-underline-offset" or "text-shadow" => Offset,
        "column-gap" or "column-rule" or "border-spacing" => Line,
        _ when property.StartsWith("border", StringComparison.Ordinal) && property != "border-collapse" => Line,
        _ => null,
    };

    /// <summary>The value with its lengths in bounds, or null when it holds something that cannot be bounded.</summary>
    public static string? Clamp(string property, string value)
    {
        if (BoundsOf(property) is not { } bounds) return value;
        if (value.IndexOf('(') >= 0 && property is not ("text-shadow" or "column-rule") && !property.StartsWith("border", StringComparison.Ordinal)) return null;
        var result = new StringBuilder(value.Length);
        int i = 0;
        while (i < value.Length)
        {
            char c = value[i];
            if (char.IsWhiteSpace(c) || c is '/' or ',' or '(' or ')' or '"' or '\'')
            {
                if (c is '"' or '\'') { int close = value.IndexOf(c, i + 1); if (close < 0) close = value.Length - 1; result.Append(value, i, close + 1 - i); i = close + 1; continue; }
                result.Append(c); i++; continue;
            }
            int start = i;
            while (i < value.Length && !char.IsWhiteSpace(value[i]) && value[i] is not ('/' or ',' or '(' or ')' or '"' or '\'')) i++;
            string token = value.Substring(start, i - start);
            // in "font" the line height follows a '/': it gets the line-height range
            Bounds tokenBounds = property == "font" && start > 0 && value.LastIndexOf('/', start - 1) >= 0 && !value.Substring(value.LastIndexOf('/', start - 1) + 1, start - value.LastIndexOf('/', start - 1) - 1).Trim().Any(char.IsLetterOrDigit)
                ? LineHeight : bounds;
            string? bounded = ClampToken(token, tokenBounds, property is "font" or "text-shadow" or "column-rule" || property.StartsWith("border", StringComparison.Ordinal));
            if (bounded is null) return null;
            result.Append(bounded);
        }
        return result.ToString();
    }

    private static string? ClampToken(string token, Bounds bounds, bool hasOtherTokens)
    {
        bool isFontSize = bounds.Equals(FontSize);
        int end = 0;
        if (end < token.Length && token[end] is '+' or '-') end++;
        int digits = end;
        while (end < token.Length && (char.IsDigit(token[end]) || token[end] == '.')) end++;
        if (end == digits) return token;                                                       // a keyword or colour name
        if (token[0] == '#') return token;
        if (!double.TryParse(token.Substring(0, end), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) return null;
        string unit = token.Substring(end).ToLowerInvariant();
        double low, high, scale;
        switch (unit)
        {
            case "%": low = bounds.MinimumPercent; high = isFontSize ? Math.Min(bounds.MaximumPercent, 400) : bounds.MaximumPercent; scale = 1; break;
            case "": if (number == 0) return token; low = bounds.MinimumNumber; high = bounds.MaximumNumber; scale = 1; if (high == 0) return hasOtherTokens ? token : null; break;
            default:
                scale = unit switch { "pt" => 1, "px" => 0.75, "in" => 72, "cm" => 72 / 2.54, "mm" => 72 / 25.4, "q" => 72 / 101.6, "pc" => 12, "em" or "rem" => 12, "ex" or "ch" => 6, _ => 0 };
                if (scale == 0) return null;                                                   // an unknown unit or an exponent (1e9px): cannot be bounded
                low = bounds.MinimumPoints; high = bounds.MaximumPoints;
                if (isFontSize && unit is "em" or "rem") { low = Math.Max(low, 0.1 * scale); high = Math.Min(high, 4 * scale); break; }   // relative: compounds down nested elements
                break;
        }
        double measured = number * scale;
        if (measured >= low && measured <= high) return token;
        double clamped = Math.Max(low, Math.Min(high, measured)) / scale;
        return clamped.ToString("0.###", CultureInfo.InvariantCulture) + token.Substring(end);
    }
}
