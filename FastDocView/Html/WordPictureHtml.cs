using System.Globalization;
using System.Text;

namespace FastDocView.Html;

/// <summary>How the text flows around a picture (format-neutral).</summary>
public enum WordPictureWrap : byte
{
    /// <summary>In the line, like a character (.docx wp:inline, ODF anchor-type as-char, RTF \pict in the text).</summary>
    Inline,
    /// <summary>Text on its sides (.docx wrapSquare / wrapTight / wrapThrough, ODF style:wrap left / right / parallel / dynamic).</summary>
    Square,
    /// <summary>Text above and below only (.docx wrapTopAndBottom, ODF style:wrap none).</summary>
    TopAndBottom,
    /// <summary>In front of or behind the text (.docx wrapNone, ODF style:wrap run-through).</summary>
    None,
}

/// <summary>Where a floating picture sits across the column - physical sides, as DrawingML wp:align and ODF style:horizontal-pos.</summary>
public enum WordPictureSide : byte { Left, Center, Right }

/// <summary>A picture's box: size and the gap kept to the text around it (points), wrap and side.</summary>
public readonly record struct WordPictureLayout(double WidthPt, double HeightPt, WordPictureWrap Wrap, WordPictureSide Side,
                                                double TopPt, double RightPt, double BottomPt, double LeftPt, double OffsetPt = 0);

/// <summary>
/// Draws a picture at its size where the document puts it. The browser has no page or anchor, so a floating picture is a
/// CSS float in its own paragraph (the page keeps each block a formatting context, so a float never runs into blocks that
/// are not on screen):
/// - square / tight / through: float to its side with the gap to the text as margins; the margin on the side without text
///   is dropped, as LibreOffice emulates Word (GraphicImport.cxx: "when impossible to wrap text on a side, remove the margin
///   gap"). A centred picture cannot float in CSS: it becomes a centred block (text above and below).
/// - top and bottom: a block on its side with the top / bottom gap.
/// - none (in front of / behind the text): LibreOffice imports it as wrap THROUGH with every gap dropped (GraphicImport.cxx
///   LN_EG_WrapType_wrapNone); text cannot run under a picture here without hiding it, so it is drawn as a block on its side
///   without gaps - the closest layout that keeps both readable.
/// A picture the browser cannot show (EMF / WMF / TIFF, a missing part) keeps its box, marked as an image.
/// </summary>
public static class WordPictureHtml
{
    public static void Append(StringBuilder html, WordPictureLayout layout, string? dataUri)
    {
        var css = new StringBuilder();
        if (layout.WidthPt > 0) css.Append("width:").Append(Points(layout.WidthPt)).Append("pt;");
        if (layout.WidthPt > 0 && layout.HeightPt > 0)
            css.Append("aspect-ratio:").Append(Points(layout.WidthPt)).Append('/').Append(Points(layout.HeightPt)).Append(';');
        else if (layout.HeightPt > 0) css.Append("height:").Append(Points(layout.HeightPt)).Append("pt;");
        WordPictureSide side = layout.Side;
        WordPictureWrap wrap = layout.Wrap;
        if (wrap == WordPictureWrap.Square && side == WordPictureSide.Center) wrap = WordPictureWrap.TopAndBottom;
        switch (wrap)
        {
            case WordPictureWrap.Inline:
                if (layout.TopPt + layout.RightPt + layout.BottomPt + layout.LeftPt > 0) css.Append("margin:").Append(Margins(layout.TopPt, layout.RightPt, layout.BottomPt, layout.LeftPt)).Append(';');
                break;
            case WordPictureWrap.Square:
                css.Append("float:").Append(side == WordPictureSide.Right ? "right" : "left").Append(";margin:")
                   .Append(side == WordPictureSide.Right ? Margins(layout.TopPt, 0, layout.BottomPt, layout.LeftPt)
                                                         : Margins(layout.TopPt, layout.RightPt, layout.BottomPt, layout.OffsetPt))
                   .Append(';');
                break;
            default:
                double top = wrap == WordPictureWrap.None ? 0 : layout.TopPt, bottom = wrap == WordPictureWrap.None ? 0 : layout.BottomPt;
                css.Append("display:block;margin:").Append(Points(top)).Append("pt ")
                   .Append(side == WordPictureSide.Right ? "0" : "auto").Append(' ')
                   .Append(Points(bottom)).Append("pt ")
                   .Append(side == WordPictureSide.Left ? Points(layout.OffsetPt) + "pt" : "auto").Append(';');
                break;
        }
        if (dataUri is not null)
            html.Append("<img class=\"word-picture\" alt=\"\" src=\"").Append(dataUri).Append("\" style=\"").Append(css).Append("\">");
        else
            html.Append("<span class=\"word-picture word-picture-missing\" style=\"").Append(css).Append("\">[image]</span>");
    }

    /// <summary>A browser-renderable picture as a data: URI by its file extension, null for formats a browser cannot show.</summary>
    public static string? DataUri(string partName, byte[]? bytes)
    {
        if (bytes is not { Length: > 0 }) return null;
        string? type = Path.GetExtension(partName).ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" or ".jpe" or ".jfif" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" or ".dib" => "image/bmp",
            ".webp" => "image/webp",
            ".svg" => "image/svg+xml",   // as <img> an SVG runs no script and loads nothing
            _ => null,
        };
        return type is null ? null : "data:" + type + ";base64," + Convert.ToBase64String(bytes);
    }

    /// <summary>A file extension for picture bytes by their signature (an embedded picture without a name), "" when unknown.</summary>
    public static string ExtensionOf(byte[] bytes)
    {
        bool Starts(params byte[] signature) => bytes.Length >= signature.Length && bytes.AsSpan(0, signature.Length).SequenceEqual(signature);
        if (Starts(0x89, 0x50, 0x4E, 0x47)) return ".png";
        if (Starts(0xFF, 0xD8, 0xFF)) return ".jpg";
        if (Starts(0x47, 0x49, 0x46, 0x38)) return ".gif";
        if (Starts(0x42, 0x4D)) return ".bmp";
        return "";
    }

    private static string Margins(double top, double right, double bottom, double left)
        => Points(top) + "pt " + Points(right) + "pt " + Points(bottom) + "pt " + Points(left) + "pt";

    private static string Points(double value) => Math.Max(0, value).ToString("0.##", CultureInfo.InvariantCulture);
}
