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

    public static Borders ReadBorders(XElement? borders) => ReadBorders(borders, null);

    public static Borders ReadBorders(XElement? borders, DocxTheme? theme)
    {
        if (borders is null) return Borders.None;
        return new Borders(
            BorderCss(borders.Element(W + "top"), theme),
            BorderCss(borders.Element(W + "start") ?? borders.Element(W + "left"), theme),
            BorderCss(borders.Element(W + "bottom"), theme),
            BorderCss(borders.Element(W + "end") ?? borders.Element(W + "right"), theme),
            BorderCss(borders.Element(W + "insideH"), theme),
            BorderCss(borders.Element(W + "insideV"), theme));
    }

    /// <summary>One border element → CSS, or null when absent.</summary>
    public static string? BorderCss(XElement? border) => BorderCss(border, null);

    /// <summary>One border element → CSS; the color goes through the theme (themeColor + tint / shade supersede @color, §17.3.2.3).</summary>
    public static string? BorderCss(XElement? border, DocxTheme? theme)
    {
        if (border is null) return null;
        string style = (string?)border.Attribute(W + "val") ?? "single";
        if (style is "nil" or "none") return "none";
        int eighths = DocxStyleCascade.TwipsAttribute(border, "sz") ?? 4;
        string width = Math.Max(0.25, eighths / 8.0).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        string cssStyle = style switch { "double" => "double", "dotted" => "dotted", "dashed" or "dashSmallGap" => "dashed", _ => "solid" };
        string? color = DocxTheme.ResolveColorHex(border, "color", theme);
        string cssColor = color is not null ? "#" + color : "#000";
        return $"{width} {cssStyle} {cssColor}";
    }

    public static string CellCss(XElement? tcPr, Borders tableBorders) => CellCss(tcPr, tableBorders, null);

    /// <summary>
    /// w:textDirection (ST_TextDirection, transitional and strict names) as CSS writing-mode. Mapping read in
    /// LibreOffice DomainMapper.cxx: tbRl → TB_RL, tbRlV → TB_RL90, btLr → BT_LR, tbLrV → TB_LR, lrTb / lrTbV → LR_TB.
    /// </summary>
    public static string? TextDirectionCss(string? value) => value switch
    {
        "tbRl" or "rl" or "tbRlV" or "rlV" => "writing-mode:vertical-rl;",
        "btLr" or "lr" => "writing-mode:sideways-lr;",
        "tbLrV" or "lrV" => "writing-mode:vertical-lr;",
        _ => null,                               // lrTb / tb / lrTbV / tbV: horizontal, the default
    };

    public static string CellCss(XElement? tcPr, Borders tableBorders, DocxTheme? theme)
    {
        var css = new StringBuilder();
        if (WidthCss(tcPr?.Element(W + "tcW")) is { } width) css.Append("width:").Append(width).Append(';');
        if (tcPr?.Element(W + "shd") is { } shading && DocxStyleCascade.FillHex(shading, theme) is { } fill) css.Append("background:#").Append(fill).Append(';');
        string? verticalAlign = (string?)tcPr?.Element(W + "vAlign")?.Attribute(W + "val");
        if (verticalAlign is "center" or "bottom") css.Append("vertical-align:").Append(verticalAlign == "center" ? "middle" : "bottom").Append(';');
        if (TextDirectionCss((string?)tcPr?.Element(W + "textDirection")?.Attribute(W + "val")) is { } direction) css.Append(direction);

        Borders cell = ReadBorders(tcPr?.Element(W + "tcBorders"), theme);
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
