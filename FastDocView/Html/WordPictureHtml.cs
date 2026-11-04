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

/// <summary>A text box's own look: inner margins (points), border / fill as CSS colours (null = none), rotation, writing-mode CSS (null = horizontal), auto-grow height.</summary>
public readonly record struct WordTextBoxStyle(double LeftInsetPt, double TopInsetPt, double RightInsetPt, double BottomInsetPt,
                                               string? BorderColor, double BorderWidthPt, bool BorderDash, string? FillColor,
                                               double RotationDegrees, string? WritingMode, bool GrowsWithText);

/// <summary>
/// Which sides of a <see cref="WordPictureWrap.Square"/> object the text may use: both (.docx wrapText bothSides / largest, VML
/// side both / largest, .doc / RTF wrk 0 / 3, ODF style:wrap parallel / dynamic / biggest) or one only (wrapText / side / wrk
/// left / right, ODF style:wrap left / right: "text wraps around the left side", so the object goes right).
/// </summary>
public enum WordPictureTextSides : byte { Both, LeftOnly, RightOnly }

/// <summary>
/// A picture's box: size and the gap kept to the text around it (points), wrap and side; for a floating object the sides its text
/// may use and its wrap contour (outlines in its own box points, <see cref="WordPictureContour"/>; null = its box).
/// </summary>
public readonly record struct WordPictureLayout(double WidthPt, double HeightPt, WordPictureWrap Wrap, WordPictureSide Side,
                                                double TopPt, double RightPt, double BottomPt, double LeftPt, double OffsetPt = 0,
                                                WordPictureTextSides TextSides = WordPictureTextSides.Both,
                                                IReadOnlyList<IReadOnlyList<(double X, double Y)>>? Contour = null, double ColumnWidthPt = 0);

/// <summary>
/// Draws a picture at its size where the document puts it. The browser has no page or anchor, so a floating picture is a
/// CSS float in its own paragraph (the page keeps each block a formatting context, so a float never runs into blocks that
/// are not on screen):
/// - square / tight / through: float with the gap to the text as margins; the margin on the side without text is dropped, as
///   LibreOffice emulates Word (GraphicImport.cxx: "when impossible to wrap text on a side, remove the margin gap"). CSS cannot
///   put text on both sides of a float (CSS Shapes 1, section 2), so an object with text on both sides floats flush to the
///   column edge nearest its position (left / right as placed; a centred one to the left - our choice, the tie has no rule),
///   without its horizontal offset, and the text wraps fully along its other side. A one-sided wrap floats to the side
///   opposite its text and keeps its offset. A contour (<see cref="WordPictureLayout.Contour"/>) becomes shape-outside.
/// - top and bottom: a block on its side with the top / bottom gap.
/// - none (in front of / behind the text): LibreOffice imports it as wrap THROUGH with every gap dropped (GraphicImport.cxx
///   LN_EG_WrapType_wrapNone); text cannot run under a picture here without hiding it, so it is drawn as a block on its side
///   without gaps - the closest layout that keeps both readable.
/// A picture the browser cannot show (EMF / WMF / TIFF, a missing part) keeps its box, marked as an image.
///
/// The bytes never travel in the block HTML: the &lt;img&gt; carries only the picture's id in
/// <see cref="PictureIdAttribute"/> (no src), its width / height (CSS px at 96 dpi, so the row keeps its height before the
/// picture loads) and loading="lazy" decoding="async". The host turns the id into a URL of its own (the service's picture
/// endpoint) and the bytes are read on request by <see cref="IWordWindowedDocument.ReadPicture"/> - one picture, never the
/// block or the document.
/// </summary>
public static class WordPictureHtml
{
    public static void Append(StringBuilder html, WordPictureLayout layout, string? pictureId)
    {
        var css = new StringBuilder();
        if (layout.WidthPt > 0) css.Append("width:").Append(Points(layout.WidthPt)).Append("pt;");
        if (layout.WidthPt > 0 && layout.HeightPt > 0)
            css.Append("aspect-ratio:").Append(Points(layout.WidthPt)).Append('/').Append(Points(layout.HeightPt)).Append(';');
        else if (layout.HeightPt > 0) css.Append("height:").Append(Points(layout.HeightPt)).Append("pt;");
        AppendPlacement(css, layout);
        if (pictureId is not null)
        {
            html.Append("<img class=\"word-picture\" alt=\"\" ").Append(PictureIdAttribute).Append("=\"").Append(WordHtmlText.Encode(pictureId)).Append('"');
            if (layout.WidthPt > 0 && layout.HeightPt > 0)
                html.Append(" width=\"").Append(Pixels(layout.WidthPt)).Append("\" height=\"").Append(Pixels(layout.HeightPt)).Append('"');
            html.Append(" loading=\"lazy\" decoding=\"async\" style=\"").Append(css).Append("\">");
        }
        else
            html.Append("<span class=\"word-picture word-picture-missing\" style=\"").Append(css).Append("\">[image]</span>");
    }

    /// <summary>The float / block / margin CSS of an object placed as <see cref="WordPictureHtml"/> describes (pictures, text boxes, shapes share it).</summary>
    public static void AppendPlacement(StringBuilder css, WordPictureLayout layout)
    {
        WordPictureSide side = layout.Side;
        switch (layout.Wrap)
        {
            case WordPictureWrap.Inline:
                if (layout.TopPt + layout.RightPt + layout.BottomPt + layout.LeftPt > 0) css.Append("margin:").Append(Margins(layout.TopPt, layout.RightPt, layout.BottomPt, layout.LeftPt)).Append(';');
                break;
            case WordPictureWrap.Square:
                bool floatsRight = FloatsRight(layout);
                double offset = layout.TextSides == WordPictureTextSides.Both ? 0 : layout.OffsetPt;
                css.Append("float:").Append(floatsRight ? "right" : "left").Append(";margin:")
                   .Append(floatsRight ? Margins(layout.TopPt, 0, layout.BottomPt, layout.LeftPt)
                                       : Margins(layout.TopPt, layout.RightPt, layout.BottomPt, offset))
                   .Append(';');
                if (layout.Contour is { } contour
                    && WordPictureContour.ShapeOutsideCss(contour, layout.WidthPt, layout.HeightPt, !floatsRight, floatsRight ? layout.LeftPt : layout.RightPt) is { } shape)
                    css.Append(shape);
                break;
            default:
                double top = layout.Wrap == WordPictureWrap.None ? 0 : layout.TopPt, bottom = layout.Wrap == WordPictureWrap.None ? 0 : layout.BottomPt;
                css.Append("display:block;margin:").Append(Points(top)).Append("pt ")
                   .Append(side == WordPictureSide.Right ? "0" : "auto").Append(' ')
                   .Append(Points(bottom)).Append("pt ")
                   .Append(side == WordPictureSide.Left ? Points(layout.OffsetPt) + "pt" : "auto").Append(';');
                break;
        }
    }

    /// <summary>Whether a <see cref="WordPictureWrap.Square"/> object floats right: a one-sided wrap opposite its text, else its own side (centre = left).</summary>
    public static bool FloatsRight(WordPictureLayout layout) => layout.TextSides switch
    {
        WordPictureTextSides.LeftOnly => true,
        WordPictureTextSides.RightOnly => false,
        _ => layout.Side == WordPictureSide.Right || (layout.Side == WordPictureSide.Left && NearerRight(layout)),
    };

    /// <summary>
    /// An object placed by its offset from the left (no alignment) whose centre lies right of the column's centre: the right edge
    /// is nearer. The column is the section's text width (<see cref="WordPictureLayout.ColumnWidthPt"/>, 0 = unknown: left). A
    /// centred object's centre is the column's own - the tie, left.
    /// </summary>
    private static bool NearerRight(WordPictureLayout layout)
        => layout.ColumnWidthPt > 0 && layout.OffsetPt > 0 && layout.OffsetPt + layout.WidthPt / 2 > layout.ColumnWidthPt / 2;

    /// <summary>
    /// Opens a text box: a box at its placement (<see cref="AppendPlacement"/>), its width, its height (a minimum when it grows
    /// with its text - DrawingML a:spAutoFit, LibreOffice textbodypropertiescontext.cxx:255-259 PROP_TextAutoGrowHeight), the inner
    /// margins, border, fill and rotation. The caller renders the box's own paragraphs and then <see cref="CloseTextBox"/>.
    /// </summary>
    /// <param name="className">The box's class ("text-box"; drawn shapes add "word-shape": a positioned box their SVG fills).</param>
    public static void OpenTextBox(StringBuilder html, WordPictureLayout layout, WordTextBoxStyle style, string className = "text-box")
    {
        var css = new StringBuilder("box-sizing:border-box;overflow-wrap:break-word;");
        if (layout.Wrap == WordPictureWrap.Inline) css.Append("display:inline-block;vertical-align:bottom;");
        if (layout.WidthPt > 0) css.Append("width:").Append(Points(layout.WidthPt)).Append("pt;");
        if (layout.HeightPt > 0) css.Append(style.GrowsWithText ? "min-height:" : "height:").Append(Points(layout.HeightPt)).Append("pt;");
        if (!style.GrowsWithText && layout.HeightPt > 0) css.Append("overflow:hidden;");
        AppendPlacement(css, layout);
        css.Append("padding:").Append(Margins(style.TopInsetPt, style.RightInsetPt, style.BottomInsetPt, style.LeftInsetPt)).Append(';');
        if (style.BorderColor is { } border) css.Append("border:").Append(Points(Math.Max(0.25, style.BorderWidthPt))).Append("pt ").Append(style.BorderDash ? "dashed" : "solid").Append(' ').Append(border).Append(';');
        else css.Append("border:0;");   // over the stylesheet's frame for boxes without a placement
        if (style.FillColor is { } fill) css.Append("background:").Append(fill).Append(';');
        if (style.RotationDegrees != 0) css.Append("transform:rotate(").Append(style.RotationDegrees.ToString("0.##", CultureInfo.InvariantCulture)).Append("deg);");
        if (style.WritingMode is { } writingMode) css.Append(writingMode);
        html.Append("<span class=\"").Append(className).Append("\" style=\"").Append(css).Append("\">");
    }

    public static void CloseTextBox(StringBuilder html) => html.Append("</span>");

    /// <summary>The &lt;img&gt; attribute naming the picture for <see cref="IWordWindowedDocument.ReadPicture"/>.</summary>
    public const string PictureIdAttribute = "data-word-picture";

    /// <summary>The media type of a picture a browser can show, by its file extension; null for formats it cannot show.</summary>
    public static string? ContentTypeOf(string partNameOrExtension) => Path.GetExtension(partNameOrExtension).ToLowerInvariant() switch
    {
        ".png" => "image/png",
        ".jpg" or ".jpeg" or ".jpe" or ".jfif" => "image/jpeg",
        ".gif" => "image/gif",
        ".bmp" or ".dib" => "image/bmp",
        ".webp" => "image/webp",
        ".svg" => "image/svg+xml",   // as <img> an SVG runs no script and loads nothing
        _ => null,
    };

    /// <summary>A picture's bytes as the host serves them, or null when the browser cannot show its format.</summary>
    public static WordDocumentPicture? Picture(string partNameOrExtension, byte[]? bytes)
        => bytes is { Length: > 0 } && ContentTypeOf(partNameOrExtension) is { } type ? new WordDocumentPicture(type, bytes) : null;

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

    private static string Pixels(double points) => Math.Max(1, Math.Round(points * 96 / 72)).ToString(CultureInfo.InvariantCulture);

    private static string Points(double value) => Math.Max(0, value).ToString("0.##", CultureInfo.InvariantCulture);
}
