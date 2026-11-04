using System.Globalization;

namespace FastDocView.Html;

/// <summary>
/// A document's explicit colour as CSS that reads in both themes, without making the HTML theme-dependent:
/// <c>light-dark(original, darkVariant)</c>. The page's <c>color-scheme</c> picks one (the app sets <c>dark</c> on its
/// root in a dark theme), so a theme switch needs no re-render. "Auto" colours are never written: they inherit the
/// page's text colour (currentColor).
/// <para>
/// The dark variant keeps hue and chroma and inverts lightness in CIELAB: L' = min(110 - L, 100), a and b unchanged.
/// Source: Chromium's dark mode, third_party/blink/renderer/platform/graphics/dark_mode_color_filter.cc
/// (LABColorFilter::InvertColor) — the same remap for text and backgrounds. Word's own dark mode likewise "shifts"
/// the document's colours instead of inverting them (support.microsoft.com, "Dark Mode in Word"); its exact formula
/// is not published, so Chromium's is the cited one. Black text becomes near-white, white shading near-black.
/// </para>
/// </summary>
public static class WordHtmlColor
{
    /// <summary>
    /// <paramref name="color"/> as theme-aware CSS. Accepts "RRGGBB", "#RRGGBB", "#RGB" and the basic named colours
    /// Word's highlights use; any other value (a gradient, "transparent", an unknown name) is returned unchanged.
    /// </summary>
    public static string Css(string color)
    {
        if (!TryParse(color, out int red, out int green, out int blue)) return color;
        // the light value is the document's own, as written (light mode output stays as before)
        string light = color.StartsWith("#", StringComparison.Ordinal) ? color : IsHexDigits(color) ? "#" + color : Hex(red, green, blue);
        (int darkRed, int darkGreen, int darkBlue) = DarkVariant(red, green, blue);
        return "light-dark(" + light + "," + Hex(darkRed, darkGreen, darkBlue) + ")";
    }

    /// <summary>
    /// <paramref name="html"/> for display in either theme: every explicit colour in a colour-bearing place is written
    /// through <see cref="Css"/>. Only these places are touched (an allowlist): the attributes in
    /// <see cref="ColorAttributes"/>, and inside style="..." and &lt;style&gt; elements the properties in
    /// <see cref="IsColorProperty"/>. Text, every other attribute (data-*, title, alt, class, href...), url(...), quoted
    /// strings and values already in light-dark() are copied as they are. Pictures are &lt;img&gt; and never change.
    /// <para>
    /// A shape or text box (class word-shape / text-box) whose first solid fill (its own background, else the first
    /// fill="..." inside it) is better read with dark text once remapped gets the class <see cref="DarkTextOnFillClass"/>;
    /// the stylesheet gives it light-dark(currentColor, #000), so text with no colour of its own picks black over a
    /// light remapped fill in the dark theme (WCAG contrast against the remapped fill, light text taken as #d4d4d4),
    /// and light mode is unchanged. One linear pass.
    /// </para>
    /// </summary>
    public static string ThemeAwareMarkup(string html)
    {
        if (html.IndexOf('<') < 0) return html;
        var output = new System.Text.StringBuilder(html.Length + html.Length / 8);
        var shapes = new List<ShapeScope>();
        int depth = 0, position = 0;
        while (position < html.Length)
        {
            int open = html.IndexOf('<', position);
            if (open < 0) { output.Append(html, position, html.Length - position); break; }
            output.Append(html, position, open - position);
            if (string.CompareOrdinal(html, open, "<!--", 0, 4) == 0)
            {
                int commentEnd = html.IndexOf("-->", open + 4, StringComparison.Ordinal);
                position = commentEnd < 0 ? html.Length : commentEnd + 3;
                output.Append(html, open, position - open);
                continue;
            }
            int close = TagEnd(html, open);
            if (close < 0) { output.Append(html, open, html.Length - open); break; }
            position = close + 1;
            if (open + 1 < html.Length && html[open + 1] is '!' or '?') { output.Append(html, open, position - open); continue; }
            if (open + 1 < html.Length && html[open + 1] == '/')
            {
                output.Append(html, open, position - open);
                depth--;
                while (shapes.Count > 0 && shapes[shapes.Count - 1].Depth > depth) CloseShape(output, shapes);
                continue;
            }
            string name = AppendTag(html, open, close, output, shapes, depth, out bool isSelfClosing);
            if (isSelfClosing || VoidElements.Contains(name))
            {
                while (shapes.Count > 0 && shapes[shapes.Count - 1].Depth > depth) CloseShape(output, shapes);
                continue;
            }
            depth++;
            if (name is "style" or "script")
            {
                int end = html.IndexOf("</" + name, position, StringComparison.OrdinalIgnoreCase);
                if (end < 0) end = html.Length;
                string raw = html.Substring(position, end - position);
                output.Append(name == "style" ? ThemeAwareStyleSheet(raw) : raw);
                position = end;
            }
        }
        while (shapes.Count > 0) CloseShape(output, shapes);
        return output.ToString();
    }

    /// <summary>The class a shape gets when dark text reads better on its remapped fill (styled by the stylesheet).</summary>
    public const string DarkTextOnFillClass = "dark-text-on-fill";

    /// <summary>The attributes whose whole value is a colour (SVG paint and HTML 4 colour attributes).</summary>
    public static readonly IReadOnlyCollection<string> ColorAttributes =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "fill", "stroke", "stop-color", "flood-color", "lighting-color", "color", "bgcolor" };

    /// <summary>A CSS property whose value can hold a colour.</summary>
    public static bool IsColorProperty(string property)
    {
        property = property.Trim().ToLowerInvariant();
        return ColorProperties.Contains(property) || property.StartsWith("border", StringComparison.Ordinal)
               || property.StartsWith("outline", StringComparison.Ordinal) || property.StartsWith("column-rule", StringComparison.Ordinal);
    }

    private static readonly HashSet<string> ColorProperties = new(StringComparer.Ordinal)
    {
        "color", "background", "background-color", "background-image", "fill", "stroke", "stop-color", "flood-color", "lighting-color",
        "text-decoration", "text-decoration-color", "text-emphasis", "text-emphasis-color", "-webkit-text-stroke",
        "-webkit-text-stroke-color", "-webkit-text-fill-color", "box-shadow", "text-shadow", "caret-color",
    };

    private static readonly HashSet<string> VoidElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "param", "source", "track", "wbr",
    };

    private sealed class ShapeScope
    {
        public int Depth;                 // the depth of the shape's content (its element's depth + 1)
        public int ClassValueEnd;         // output index just before the class value's closing quote
        public (int R, int G, int B)? Fill;
    }

    private static void CloseShape(System.Text.StringBuilder output, List<ShapeScope> shapes)
    {
        ShapeScope shape = shapes[shapes.Count - 1];
        shapes.RemoveAt(shapes.Count - 1);
        if (shape.Fill is { } fill && PrefersDarkText(fill.R, fill.G, fill.B)) output.Insert(shape.ClassValueEnd, " " + DarkTextOnFillClass);
    }

    /// <summary>Whether black text beats the theme's light text (#d4d4d4) on the dark-theme variant of this fill (WCAG 2 contrast).</summary>
    public static bool PrefersDarkText(int red, int green, int blue)
    {
        (int r, int g, int b) = DarkVariant(red, green, blue);
        double fill = RelativeLuminance(r, g, b), light = RelativeLuminance(0xd4, 0xd4, 0xd4);
        double withBlack = (fill + 0.05) / 0.05;
        double withLight = (Math.Max(fill, light) + 0.05) / (Math.Min(fill, light) + 0.05);
        return withBlack > withLight;
    }

    private static double RelativeLuminance(int red, int green, int blue)
        => 0.2126 * ToLinear(red / 255.0) + 0.7152 * ToLinear(green / 255.0) + 0.0722 * ToLinear(blue / 255.0);

    /// <summary>The '&gt;' ending the tag that starts at <paramref name="open"/>, skipping quoted values; -1 when none.</summary>
    private static int TagEnd(string html, int open)
    {
        char quote = '\0';
        for (int i = open + 1; i < html.Length; i++)
        {
            char c = html[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; }
            else if (c is '"' or '\'') quote = c;
            else if (c == '>') return i;
        }
        return -1;
    }

    /// <summary>Copies the start tag [open, close] to <paramref name="output"/>, rewriting colour attributes and style; returns its name.</summary>
    private static string AppendTag(string html, int open, int close, System.Text.StringBuilder output, List<ShapeScope> shapes, int depth, out bool isSelfClosing)
    {
        int i = open + 1;
        while (i < close && !char.IsWhiteSpace(html[i]) && html[i] != '/') i++;
        string name = html.Substring(open + 1, i - open - 1).ToLowerInvariant();
        isSelfClosing = false;
        output.Append(html, open, i - open);
        ShapeScope? shape = null;
        (int, int, int)? ownFill = null, paintFill = null;
        while (i < close)
        {
            int start = i;
            while (i < close && (char.IsWhiteSpace(html[i]) || html[i] == '/')) { if (html[i] == '/') isSelfClosing = true; i++; }
            output.Append(html, start, i - start);
            if (i >= close) break;
            int nameStart = i;
            while (i < close && html[i] != '=' && !char.IsWhiteSpace(html[i]) && html[i] != '/') i++;
            string attribute = html.Substring(nameStart, i - nameStart);
            output.Append(attribute);
            isSelfClosing = false;   // a '/' before an attribute does not close the tag
            if (i >= close || html[i] != '=') continue;
            output.Append('=');
            i++;
            char quote = i < close && html[i] is '"' or '\'' ? html[i] : '\0';
            int valueStart = quote != '\0' ? i + 1 : i;
            int valueEnd = quote != '\0' ? html.IndexOf(quote, valueStart) : valueStart;
            if (quote == '\0') while (valueEnd < close && !char.IsWhiteSpace(html[valueEnd])) valueEnd++;
            if (valueEnd < 0 || valueEnd > close) valueEnd = close;
            string value = html.Substring(valueStart, valueEnd - valueStart);
            if (quote != '\0') output.Append(quote);
            if (attribute.Equals("style", StringComparison.OrdinalIgnoreCase))
                output.Append(ThemeAwareDeclarations(value, out ownFill));
            else if (ColorAttributes.Contains(attribute))
            {
                output.Append(ThemeAwareValue(value, out (int, int, int)? first));
                if (attribute.Equals("fill", StringComparison.OrdinalIgnoreCase)) paintFill = first;
            }
            else
            {
                output.Append(value);
                if (quote != ' ' && attribute.Equals("class", StringComparison.OrdinalIgnoreCase) && HasShapeClass(value))
                    shape = new ShapeScope { Depth = depth + 1, ClassValueEnd = output.Length };
            }
            if (quote != '\0') output.Append(quote);
            i = quote != '\0' ? Math.Min(close, valueEnd + 1) : valueEnd;
        }
        output.Append('>');
        // the fill this element paints: an SVG fill inside open shapes, or a shape's own background
        if (paintFill is not null)
            for (int s = shapes.Count - 1; s >= 0; s--) if (shapes[s].Fill is null) shapes[s].Fill = paintFill;
        if (shape is not null)
        {
            shape.Fill = ownFill ?? paintFill;
            shapes.Add(shape);
        }
        return name;
    }

    private static bool HasShapeClass(string value)
    {
        foreach (string token in value.Split(' '))
            if (token is "word-shape" or "text-box") return true;
        return false;
    }

    /// <summary>A style sheet's declaration blocks through <see cref="ThemeAwareDeclarations"/>; selectors and at-rules are copied.</summary>
    private static string ThemeAwareStyleSheet(string css)
    {
        var output = new System.Text.StringBuilder(css.Length + 32);
        int position = 0;
        while (position < css.Length)
        {
            int open = css.IndexOf('{', position);
            if (open < 0) { output.Append(css, position, css.Length - position); break; }
            int close = css.IndexOf('}', open + 1);
            int nested = css.IndexOf('{', open + 1);
            if (nested >= 0 && (close < 0 || nested < close)) { output.Append(css, position, open + 1 - position); position = open + 1; continue; }   // @media { ... }
            if (close < 0) close = css.Length;
            output.Append(css, position, open + 1 - position).Append(ThemeAwareDeclarations(css.Substring(open + 1, close - open - 1), out _));
            position = close;
        }
        return output.ToString();
    }

    /// <summary>"prop: value; ..." with the values of colour properties through <see cref="ThemeAwareValue"/>.
    /// <paramref name="background"/>: the first colour of a background / background-color declaration.</summary>
    private static string ThemeAwareDeclarations(string declarations, out (int, int, int)? background)
    {
        background = null;
        var output = new System.Text.StringBuilder(declarations.Length + 32);
        int position = 0;
        while (position <= declarations.Length)
        {
            int end = DeclarationEnd(declarations, position);
            string declaration = declarations.Substring(position, end - position);
            int colon = declaration.IndexOf(':');
            if (colon > 0 && IsColorProperty(declaration.Substring(0, colon)))
            {
                string property = declaration.Substring(0, colon).Trim().ToLowerInvariant();
                output.Append(declaration, 0, colon + 1).Append(ThemeAwareValue(declaration.Substring(colon + 1), out (int, int, int)? first));
                if (background is null && property is "background" or "background-color") background = first;
            }
            else output.Append(declaration);
            if (end >= declarations.Length) break;
            output.Append(';');
            position = end + 1;
        }
        return output.ToString();
    }

    private static int DeclarationEnd(string declarations, int start)
    {
        int parentheses = 0;
        char quote = '\0';
        for (int i = start; i < declarations.Length; i++)
        {
            char c = declarations[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c == '(') parentheses++;
            else if (c == ')') parentheses = Math.Max(0, parentheses - 1);
            else if (c == ';' && parentheses == 0) return i;
        }
        return declarations.Length;
    }

    /// <summary>
    /// A colour value with each colour token (#rgb, #rrggbb, a named colour, rgb() / rgba() / hsl() / hsla()) through the same
    /// light-dark rule as <see cref="Css"/>; url(...),
    /// light-dark(...), quoted strings and character references (&amp;#160;) are copied.
    /// <paramref name="first"/>: the first colour, for a shape's fill.
    /// </summary>
    private static string ThemeAwareValue(string value, out (int, int, int)? first)
    {
        first = null;
        System.Text.StringBuilder? output = null;
        int copied = 0;
        for (int i = 0; i < value.Length; i++)
        {
            char c = value[i];
            if (c is '"' or '\'') { int q = value.IndexOf(c, i + 1); i = q < 0 ? value.Length : q; continue; }
            if (c is 'u' or 'U' or 'l' or 'L' && (StartsWithAt(value, i, "url(") || StartsWithAt(value, i, "light-dark(")))
            {
                // to the matching ')': light-dark(rgb(...), ...) and url(a(b).png) hold nested parentheses
                int openParenthesis = value.IndexOf('(', i);
                // an unquoted url(...) may hold a quote character and ends at its first ')' (CSS Syntax 3, 4.3.6)
                bool isUnquotedUrl = c is 'u' or 'U' && value.Substring(openParenthesis + 1).TrimStart() is { Length: > 0 } rest && rest[0] is not ('"' or '\'');
                int closeParenthesis = isUnquotedUrl ? value.IndexOf(')', openParenthesis) : MatchingParenthesis(value, openParenthesis);
                i = closeParenthesis < 0 ? value.Length : closeParenthesis;
                continue;
            }
            if (c is 'r' or 'R' or 'h' or 'H' && (i == 0 || !(char.IsLetterOrDigit(value[i - 1]) || value[i - 1] is '-' or '_'))
                && (StartsWithAt(value, i, "rgb(") || StartsWithAt(value, i, "rgba(") || StartsWithAt(value, i, "hsl(") || StartsWithAt(value, i, "hsla(")))
            {
                // a functional colour (MHTML writes them): the same light-dark pair, the alpha kept on the dark side
                int open = value.IndexOf('(', i);
                int closeParenthesis = MatchingParenthesis(value, open);
                if (closeParenthesis < 0) { i = value.Length; continue; }
                string function = value.Substring(i, closeParenthesis + 1 - i);
                if (TryParseFunction(value.Substring(i, open - i), value.Substring(open + 1, closeParenthesis - open - 1), out int fRed, out int fGreen, out int fBlue, out string? alpha))
                {
                    first ??= (fRed, fGreen, fBlue);
                    (int darkRed, int darkGreen, int darkBlue) = DarkVariant(fRed, fGreen, fBlue);
                    string dark = alpha is null ? Hex(darkRed, darkGreen, darkBlue)
                        : "rgba(" + darkRed + "," + darkGreen + "," + darkBlue + "," + alpha + ")";
                    output ??= new System.Text.StringBuilder(value.Length + 32);
                    output.Append(value, copied, i - copied).Append("light-dark(").Append(function).Append(',').Append(dark).Append(')');
                    copied = closeParenthesis + 1;
                }
                i = closeParenthesis;
                continue;
            }
            int end;
            if (c == '#' && (i == 0 || value[i - 1] != '&'))
            {
                end = i + 1;
                while (end < value.Length && Uri.IsHexDigit(value[end])) end++;
                if (end - i - 1 is not (3 or 6) || (end < value.Length && (char.IsLetterOrDigit(value[end]) || value[end] is '_' or '-'))) { i = end - 1; continue; }
            }
            else if (char.IsLetter(c) && (i == 0 || !(char.IsLetterOrDigit(value[i - 1]) || value[i - 1] is '-' or '_' or '&' or '#')))
            {
                end = i;
                while (end < value.Length && (char.IsLetterOrDigit(value[end]) || value[end] is '-' or '_')) end++;
                string word = value.Substring(i, end - i);
                if (word.All(Uri.IsHexDigit) || !word.All(char.IsLetter) || !TryParse(word, out _, out _, out _)) { i = end - 1; continue; }
            }
            else continue;
            string token = value.Substring(i, end - i);
            if (TryParse(token, out int red, out int green, out int blue))
            {
                first ??= (red, green, blue);
                output ??= new System.Text.StringBuilder(value.Length + 32);
                output.Append(value, copied, i - copied).Append(Css(token));
                copied = end;
            }
            i = end - 1;
        }
        return output is null ? value : output.Append(value, copied, value.Length - copied).ToString();
    }

    /// <summary>The ')' matching the '(' at <paramref name="open"/> (quotes skipped); -1 when unbalanced.</summary>
    private static int MatchingParenthesis(string value, int open)
    {
        if (open < 0) return -1;
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < value.Length; i++)
        {
            char c = value[i];
            if (quote != '\0') { if (c == quote) quote = '\0'; continue; }
            if (c is '"' or '\'') quote = c;
            else if (c == '(') depth++;
            else if (c == ')' && --depth == 0) return i;
        }
        return -1;
    }

    /// <summary>
    /// rgb() / rgba() / hsl() / hsla() arguments (CSS Color 4: commas or spaces, "/ alpha", channels as numbers or percentages,
    /// hue in degrees); false for anything else (var(), calc(), none) - such a value is left as written.
    /// </summary>
    private static bool TryParseFunction(string name, string arguments, out int red, out int green, out int blue, out string? alpha)
    {
        red = green = blue = 0;
        alpha = null;
        string[] parts = arguments.Replace("/", " / ").Replace(',', ' ').Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var channels = new List<string>(4);
        foreach (string part in parts) if (part != "/") channels.Add(part);
        if (channels.Count is not (3 or 4)) return false;
        if (channels.Count == 4)
        {
            if (!TryNumber(channels[3], 1, out double a)) return false;
            alpha = Math.Max(0, Math.Min(1, a)).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture);
        }
        bool isHsl = name.StartsWith("hsl", StringComparison.OrdinalIgnoreCase);
        if (!isHsl)
        {
            if (!TryNumber(channels[0], 255, out double r) || !TryNumber(channels[1], 255, out double g) || !TryNumber(channels[2], 255, out double b)) return false;
            red = Channel(r); green = Channel(g); blue = Channel(b);
            return true;
        }
        string hueText = channels[0].EndsWith("deg", StringComparison.OrdinalIgnoreCase) ? channels[0].Substring(0, channels[0].Length - 3) : channels[0];
        if (!double.TryParse(hueText, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double hue)
            || double.IsNaN(hue) || double.IsInfinity(hue)
            || !channels[1].EndsWith("%", StringComparison.Ordinal) || !channels[2].EndsWith("%", StringComparison.Ordinal)
            || !TryNumber(channels[1], 1, out double saturation) || !TryNumber(channels[2], 1, out double lightness)) return false;
        saturation = Math.Max(0, Math.Min(1, saturation));
        lightness = Math.Max(0, Math.Min(1, lightness));
        // CSS Color 4, 7.1 "Converting HSL Colors to sRGB"
        hue = ((hue % 360) + 360) % 360;
        double F(int n)
        {
            double k = (n + hue / 30) % 12;
            return lightness - saturation * Math.Min(lightness, 1 - lightness) * Math.Max(-1, Math.Min(Math.Min(k - 3, 9 - k), 1));
        }
        red = Channel(F(0) * 255); green = Channel(F(8) * 255); blue = Channel(F(4) * 255);
        return true;
    }

    /// <summary>A number, or a percentage of <paramref name="full"/>.</summary>
    private static bool TryNumber(string text, double full, out double number)
    {
        bool percent = text.EndsWith("%", StringComparison.Ordinal);
        bool ok = double.TryParse(percent ? text.Substring(0, text.Length - 1) : text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number);
        if (!ok || double.IsNaN(number) || double.IsInfinity(number)) return false;
        if (percent) number = number / 100 * full;
        return true;
    }

    private static int Channel(double value) => (int)Math.Round(Math.Max(0, Math.Min(255, value)));

    private static bool StartsWithAt(string text, int index, string prefix)
        => index + prefix.Length <= text.Length && string.Compare(text, index, prefix, 0, prefix.Length, StringComparison.OrdinalIgnoreCase) == 0;

    /// <summary>The dark-theme variant: CIELAB lightness inverted (110 - L, at most 100), a and b kept.</summary>
    public static (int Red, int Green, int Blue) DarkVariant(int red, int green, int blue)
    {
        (double l, double a, double b) = ToLab(red, green, blue);
        return FromLab(Math.Min(110 - l, 100), a, b);
    }

    private static bool IsHexDigits(string value) => value.Length is 3 or 6 && value.All(Uri.IsHexDigit);

    private static string Hex(int red, int green, int blue) => "#" + red.ToString("x2") + green.ToString("x2") + blue.ToString("x2");

    private static bool TryParse(string color, out int red, out int green, out int blue)
    {
        red = green = blue = 0;
        string value = color.Trim().ToLowerInvariant();
        value = value switch
        {
            "black" => "000000", "white" => "ffffff", "red" => "ff0000", "lime" => "00ff00", "green" => "008000",
            "blue" => "0000ff", "yellow" => "ffff00", "cyan" or "aqua" => "00ffff", "magenta" or "fuchsia" => "ff00ff",
            "gray" or "grey" => "808080", "silver" => "c0c0c0", "maroon" => "800000", "navy" => "000080", "olive" => "808000",
            "purple" => "800080", "teal" => "008080",
            _ => value.StartsWith("#", StringComparison.Ordinal) ? value.Substring(1) : value,
        };
        if (value.Length == 3) value = new string([value[0], value[0], value[1], value[1], value[2], value[2]]);
        if (value.Length != 6 || !int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int packed)) return false;
        red = packed >> 16 & 0xff; green = packed >> 8 & 0xff; blue = packed & 0xff;
        return true;
    }

    // sRGB (D65) <-> CIELAB, the standard formulas (IEC 61966-2-1 transfer curve, CIE 1976 L*a*b*).
    private const double WhiteX = 0.95047, WhiteY = 1.0, WhiteZ = 1.08883;

    private static (double L, double A, double B) ToLab(int red, int green, int blue)
    {
        double r = ToLinear(red / 255.0), g = ToLinear(green / 255.0), b = ToLinear(blue / 255.0);
        double x = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / WhiteX;
        double y = (0.2126729 * r + 0.7151522 * g + 0.0721750 * b) / WhiteY;
        double z = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / WhiteZ;
        double fx = LabF(x), fy = LabF(y), fz = LabF(z);
        return (116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz));
    }

    private static (int Red, int Green, int Blue) FromLab(double l, double a, double b)
    {
        double fy = (l + 16) / 116, fx = fy + a / 500, fz = fy - b / 200;
        double x = LabFInverse(fx) * WhiteX, y = LabFInverse(fy) * WhiteY, z = LabFInverse(fz) * WhiteZ;
        double r = 3.2404542 * x - 1.5371385 * y - 0.4985314 * z;
        double g = -0.9692660 * x + 1.8760108 * y + 0.0415560 * z;
        double bl = 0.0556434 * x - 0.2040259 * y + 1.0572252 * z;
        return (ToByte(r), ToByte(g), ToByte(bl));
    }

    private static double ToLinear(double channel) => channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);

    private static int ToByte(double linear)
    {
        double clamped = Math.Max(0, Math.Min(1, linear));
        double encoded = clamped <= 0.0031308 ? clamped * 12.92 : 1.055 * Math.Pow(clamped, 1 / 2.4) - 0.055;
        return (int)Math.Round(encoded * 255);
    }

    private const double Epsilon = 216.0 / 24389, Kappa = 24389.0 / 27;
    private static double LabF(double t) => t > Epsilon ? Math.Pow(t, 1.0 / 3) : (Kappa * t + 16) / 116;
    private static double LabFInverse(double f) { double cube = f * f * f; return cube > Epsilon ? cube : (116 * f - 16) / Kappa; }
}
