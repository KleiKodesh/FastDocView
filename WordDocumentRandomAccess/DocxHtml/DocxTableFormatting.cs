using System.Globalization;
using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Html;

namespace WordDocumentRandomAccess.DocxHtml;

/// <summary>
/// Table and cell properties → CSS (ECMA-376 §17.4): widths (tblW / tcW in twips, fiftieths of a
/// percent, or auto), borders (tblBorders / tcBorders, w:sz in eighths of a point), cell margins
/// (tblCellMar), vertical alignment (vAlign) and cell shading. Cell borders override the table's.
/// </summary>
public static class DocxTableFormatting
{
    private static readonly XNamespace W = DocxStyleCascade.W;

    /// <summary>Per-side CSS border values ("1pt solid #000" / "none"); null side = not specified.</summary>
    public sealed record Borders(string? Top, string? Start, string? Bottom, string? End, string? InsideHorizontal, string? InsideVertical)
    {
        public static readonly Borders None = new(null, null, null, null, null, null);
    }

    public static string? WidthCss(XElement? width)
    {
        if (width is null) return null;
        string type = (string?)width.Attribute(W + "type") ?? "dxa";
        if (!int.TryParse((string?)width.Attribute(W + "w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int value) || value <= 0) return null;
        return type switch
        {
            "dxa" => WordParagraphFormatting.TwipsToPt(value) + "pt",
            "pct" => (value / 50.0).ToString("0.#", CultureInfo.InvariantCulture) + "%",
            _ => null,
        };
    }

    public static Borders ReadBorders(XElement? borders)
    {
        if (borders is null) return Borders.None;
        return new Borders(
            BorderCss(borders.Element(W + "top")),
            BorderCss(borders.Element(W + "start") ?? borders.Element(W + "left")),
            BorderCss(borders.Element(W + "bottom")),
            BorderCss(borders.Element(W + "end") ?? borders.Element(W + "right")),
            BorderCss(borders.Element(W + "insideH")),
            BorderCss(borders.Element(W + "insideV")));
    }

    /// <summary>One border element → CSS, or null when absent.</summary>
    public static string? BorderCss(XElement? border)
    {
        if (border is null) return null;
        string style = (string?)border.Attribute(W + "val") ?? "single";
        if (style is "nil" or "none") return "none";
        int eighths = DocxStyleCascade.TwipsAttribute(border, "sz") ?? 4;
        string width = Math.Max(0.25, eighths / 8.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        string cssStyle = style switch { "double" => "double", "dotted" => "dotted", "dashed" or "dashSmallGap" => "dashed", _ => "solid" };
        string? color = (string?)border.Attribute(W + "color");
        string cssColor = color is { Length: 6 } && !color.Equals("auto", StringComparison.OrdinalIgnoreCase) ? "#" + color : "#000";
        return $"{width} {cssStyle} {cssColor}";
    }

    public static string CellCss(XElement? tcPr, Borders tableBorders)
    {
        var css = new StringBuilder();
        if (WidthCss(tcPr?.Element(W + "tcW")) is { } width) css.Append("width:").Append(width).Append(';');
        if (tcPr?.Element(W + "shd") is { } shading && DocxStyleCascade.FillHex(shading) is { } fill) css.Append("background:#").Append(fill).Append(';');
        string? verticalAlign = (string?)tcPr?.Element(W + "vAlign")?.Attribute(W + "val");
        if (verticalAlign is "center" or "bottom") css.Append("vertical-align:").Append(verticalAlign == "center" ? "middle" : "bottom").Append(';');

        Borders cell = ReadBorders(tcPr?.Element(W + "tcBorders"));
        // A cell takes its own border, else the table's inside border, else the table's outer one.
        AppendBorder(css, "border-top", cell.Top ?? tableBorders.InsideHorizontal ?? tableBorders.Top);
        AppendBorder(css, "border-bottom", cell.Bottom ?? tableBorders.InsideHorizontal ?? tableBorders.Bottom);
        AppendBorder(css, "border-inline-start", cell.Start ?? tableBorders.InsideVertical ?? tableBorders.Start);
        AppendBorder(css, "border-inline-end", cell.End ?? tableBorders.InsideVertical ?? tableBorders.End);
        return css.ToString();
    }

    public static string CellMarginCss(XElement? cellMargins)
    {
        if (cellMargins is null) return "";
        var css = new StringBuilder();
        AppendMargin(css, "padding-top", cellMargins.Element(W + "top"));
        AppendMargin(css, "padding-bottom", cellMargins.Element(W + "bottom"));
        AppendMargin(css, "padding-inline-start", cellMargins.Element(W + "start") ?? cellMargins.Element(W + "left"));
        AppendMargin(css, "padding-inline-end", cellMargins.Element(W + "end") ?? cellMargins.Element(W + "right"));
        return css.ToString();
    }

    private static void AppendBorder(StringBuilder css, string property, string? value)
    {
        if (value is not null) css.Append(property).Append(':').Append(value).Append(';');
    }

    private static void AppendMargin(StringBuilder css, string property, XElement? margin)
    {
        if (DocxStyleCascade.TwipsAttribute(margin, "w") is { } twips) css.Append(property).Append(':').Append(WordParagraphFormatting.TwipsToPt(twips)).Append("pt;");
    }
}
