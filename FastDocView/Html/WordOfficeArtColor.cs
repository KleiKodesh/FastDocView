namespace FastDocView.Html;

/// <summary>
/// An OfficeArt colour (MSO_CLR: .doc FOPT values and the RTF {\sp} colour properties, which carry the same numbers) as a CSS
/// colour, resolved as LibreOffice's SvxMSDffManager::MSO_CLR_ToColor does (filter/source/msfilter/msdffimp.cxx:3420-3640,
/// master read 2026-11-04):
/// - 0xFE in the top byte is dropped (:3425-3426); red is the low byte (:3640 Color(code, code >> 8, code >> 16)).
/// - Upper byte &amp; 0x19 with 0x08 or without 0x10: a scheme colour; Word documents have no palette, so the property's default
///   (fill / fill back: white, line: black, :3436-3454).
/// - Otherwise with 0x10: a system colour (:3456-3631): the index picks LibreOffice's StyleSettings colour - its standard
///   values (vcl/source/app/settings.cxx:497-559 SetStandardStyles: face / menu light grey C0C0C0, shadow / deactive grey
///   808080, highlight blue 000080, highlight text / window white, texts black; include/tools/color.hxx:444-462) - or another
///   property (0xF0 fill, 0xF1 line-or-fill, 0xF2 line, 0xF3 shadow, 0xF4 this = fill, 0xF5 fill back, 0xF6 line back, 0xF7
///   fill-then-line, 0xFF fill); then gray (0x8000, luminance color.hxx:276-279), the function (0x0100 darken .. 0x0600
///   black / white, by the parameter in bits 16-23) and the inverts (0x4000 top bit, 0x2000 all).
/// </summary>
public static class WordOfficeArtColor
{
    public const int FillColor = 385, FillBackColor = 387, LineColor = 448, LineBackColor = 450, ShadowColor = 513;

    /// <param name="property">The colour property's value by id (null = not set), for colours that name another property.</param>
    /// <param name="hasLine">The shape draws a line (for "line or fill").</param>
    public static string Css(uint code, int contentProperty, Func<int, uint?> property, bool hasLine, int depth = 0)
    {
        (byte r, byte g, byte b) = Resolve(code, contentProperty, property, hasLine, depth);
        return "#" + r.ToString("x2") + g.ToString("x2") + b.ToString("x2");
    }

    private static uint Default(int contentProperty) => contentProperty switch
    {
        LineColor => 0x000000, ShadowColor => 0x808080, _ => 0xFFFFFF,
    };

    private static (byte r, byte g, byte b) Rgb(uint value) => ((byte)value, (byte)(value >> 8), (byte)(value >> 16));

    private static (byte r, byte g, byte b) Resolve(uint code, int contentProperty, Func<int, uint?> property, bool hasLine, int depth)
    {
        if ((code & 0xFE000000) == 0xFE000000) code &= 0x00FFFFFF;
        uint upper = code >> 24;
        if ((upper & 0x19) == 0) return Rgb(code);
        if ((upper & 0x08) != 0 || (upper & 0x10) == 0) return Rgb(Default(contentProperty));   // scheme colour, no palette
        int parameter = (int)((code >> 16) & 0xFF), function = (int)((code & 0x0F00) >> 8), flags = (int)((code & 0xF000) >> 8);
        // the manager's default colour for Word is white (ww8par.cxx:500-501 COL_WHITE; msdffimp.cxx:3422): kept when the index is
        // unknown or names a property that is itself a system colour ("beware of looping recursive", :3619)
        (byte r, byte g, byte b) color = Rgb(0xFFFFFF);
        int? other = null;
        switch (code & 0xFF)
        {
            case 0: case 13: case 16: color = Rgb(0xC0C0C0); break;                 // button face, info background, scroll bar
            case 1: case 5: case 9: case 14: case 15: color = Rgb(0x000000); break; // window / caption / button / info / menu text
            case 2: color = Rgb(0xC0C0C0); break;                                   // menu
            case 3: case 6: case 7: case 19: color = Rgb(0x800000); break;          // highlight, active caption, button highlight, 3D light (blue 000080)
            case 4: color = Rgb(0xFFFFFF); break;                                   // highlight text
            case 8: case 10: case 11: case 12: color = Rgb(0x808080); break;        // button shadow, gray text, inactive caption (text)
            case 17: case 18: color = Rgb(0xFFFFFF); break;                         // window, window frame
            case 0xF0: case 0xF4: case 0xF7: case 0xFF: other = FillColor; break;
            case 0xF1: other = hasLine ? LineColor : FillColor; break;
            case 0xF2: other = LineColor; break;
            case 0xF3: other = ShadowColor; break;
            case 0xF5: other = FillBackColor; break;
            case 0xF6: other = LineBackColor; break;
        }
        if (other is { } id)
        {
            uint value = property(id) ?? Default(id);
            if ((value & 0x10000000) == 0) color = Resolve(value, id, property, hasLine, depth + 1);
        }
        if ((flags & 0x80) != 0)
        {
            byte luminance = (byte)((color.b * 29 + color.g * 151 + color.r * 76) >> 8);
            color = (luminance, luminance, luminance);
        }
        static byte Clamp(int v) => (byte)Math.Max(0, Math.Min(255, v));
        color = function switch
        {
            1 => ((byte)(parameter * color.r >> 8), (byte)(parameter * color.g >> 8), (byte)(parameter * color.b >> 8)),
            2 => ((byte)(((0xFF - parameter) * 0xFF + parameter * color.r) >> 8), (byte)(((0xFF - parameter) * 0xFF + parameter * color.g) >> 8),
                  (byte)(((0xFF - parameter) * 0xFF + parameter * color.b) >> 8)),
            3 => (Clamp(color.r + parameter), Clamp(color.g + parameter), Clamp(color.b + parameter)),
            4 => (Clamp(color.r - parameter), Clamp(color.g - parameter), Clamp(color.b - parameter)),
            5 => (Clamp(parameter - color.r), Clamp(parameter - color.g), Clamp(parameter - color.b)),
            6 => ((byte)(color.r < parameter ? 0 : 255), (byte)(color.g < parameter ? 0 : 255), (byte)(color.b < parameter ? 0 : 255)),
            _ => color,
        };
        if ((flags & 0x40) != 0) color = ((byte)(color.r ^ 0x80), (byte)(color.g ^ 0x80), (byte)(color.b ^ 0x80));
        if ((flags & 0x20) != 0) color = ((byte)(0xFF - color.r), (byte)(0xFF - color.g), (byte)(0xFF - color.b));
        return color;
    }
}
