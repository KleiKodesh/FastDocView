using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// DrawingML shape text (a:txBody / dsp:txBody, ECMA-376 Part 1 21.1.2: a:bodyPr, a:p with a:pPr and a:r / a:br / a:fld) as
/// HTML, for SmartArt drawings and charts, prefix-free elements. Defaults as LibreOffice gives a DrawingML shape (oox
/// shape.cxx:340-350): 18 pt text, insets 0.1in / 0.05in, top anchored, left aligned. Runs: a:rPr sz (hundredths of a
/// point) times the bodyPr normAutofit fontScale (textbodypropertiescontext.cxx:249), b, i, u, strike, baseline, the fill
/// colour, a:latin / a:cs / a:ea typefaces (+mn- / +mj- = the theme's minor / major face); the shape style's a:fontRef gives
/// the default face and colour. Paragraphs: algn, lnSpc spcPct, spcBef / spcAft spcPts, marL / indent, rtl.
/// The same walk yields the plain text search reads (<see cref="Text"/>): runs and fields in order, a separator per
/// paragraph and line break.
/// </summary>
internal static class DocxDrawingTextHtml
{
    private const double EmuPerPoint = 12700;

    public static string Render(XElement body, XElement? style, DocxTheme? theme)
    {
        XElement? bodyProperties = body.Element("bodyPr");
        double fontScale = bodyProperties?.Element("normAutofit") is { } autofit && autofit.Attribute("fontScale") is { } scale
                           && double.TryParse(scale.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double s) ? s / 100000 : 1;
        XElement? fontReference = style?.Element("fontRef");
        string? defaultColor = DocxDrawingColor.Css(fontReference, theme);
        string? defaultFace = (string?)fontReference?.Attribute("idx") switch
        {
            "major" => theme?.FontFor("majorHAnsi"), "minor" => theme?.FontFor("minorHAnsi"), _ => null,
        };
        var html = new StringBuilder();
        foreach (XElement paragraph in body.Elements("p"))
        {
            XElement? properties = paragraph.Element("pPr");
            var css = new StringBuilder("display:block;margin:0;");
            css.Append("text-align:").Append((string?)properties?.Attribute("algn") switch { "ctr" => "center", "r" => "right", "just" or "dist" => "justify", _ => "left" }).Append(';');
            if (Number(properties, "marL") is { } left && left != 0) css.Append("padding-left:").Append(Points(left / EmuPerPoint)).Append("pt;");
            if (Number(properties, "indent") is { } indent && indent != 0) css.Append("text-indent:").Append(Points(indent / EmuPerPoint)).Append("pt;");
            if (properties?.Element("lnSpc")?.Element("spcPct") is { } lineSpacing && Number(lineSpacing, "val") is { } percent) css.Append("line-height:").Append(Points(percent / 100000 * 1.15)).Append(';');   // the stylesheet's single line is 1.15
            if (properties?.Element("spcBef")?.Element("spcPts") is { } before && Number(before, "val") is { } beforePoints) css.Append("margin-top:").Append(Points(beforePoints / 100)).Append("pt;");
            if (properties?.Element("spcAft")?.Element("spcPts") is { } after && Number(after, "val") is { } afterPoints) css.Append("margin-bottom:").Append(Points(afterPoints / 100)).Append("pt;");
            bool rightToLeft = (string?)properties?.Attribute("rtl") is "1" or "true";
            html.Append("<span class=\"text-box-paragraph\" dir=\"").Append(rightToLeft ? "rtl" : "auto").Append("\" style=\"").Append(css).Append("\">");
            bool any = false;
            foreach (XElement part in paragraph.Elements())
            {
                switch (part.Name.LocalName)
                {
                    case "r" or "fld":
                    {
                        string text = (string?)part.Element("t") ?? "";
                        if (text.Length == 0) continue;
                        html.Append("<span style=\"").Append(RunCss(part.Element("rPr") ?? properties?.Element("defRPr"), theme, fontScale, defaultColor, defaultFace)).Append("\">")
                            .Append(WordHtmlText.Encode(text)).Append("</span>");
                        any = true;
                        break;
                    }
                    case "br":
                        html.Append("<br>");
                        break;
                }
            }
            if (!any)
            {
                // an empty paragraph keeps its line at its end mark's size
                XElement? end = paragraph.Element("endParaRPr");
                html.Append("<span style=\"").Append(RunCss(end, theme, fontScale, defaultColor, defaultFace)).Append("\">&#8203;</span>");
            }
            html.Append("</span>");
        }
        return html.ToString();
    }

    /// <summary>The text search reads: runs and fields in order, a separator (newline) after each paragraph and at each line break.</summary>
    public static string Text(XElement body)
    {
        var text = new StringBuilder();
        foreach (XElement paragraph in body.Elements("p"))
        {
            foreach (XElement part in paragraph.Elements())
                if (part.Name.LocalName is "r" or "fld") text.Append((string?)part.Element("t"));
                else if (part.Name.LocalName == "br") text.Append('\n');
            text.Append('\n');
        }
        return text.ToString();
    }

    /// <summary>Vertical anchor of the text (bodyPr anchor: t, ctr, b; LibreOffice default top).</summary>
    public static string Anchor(XElement body) => (string?)body.Element("bodyPr")?.Attribute("anchor") ?? "t";

    /// <summary>The text insets in points (bodyPr lIns / tIns / rIns / bIns; defaults 0.1in / 0.05in).</summary>
    public static (double Left, double Top, double Right, double Bottom) Insets(XElement body)
    {
        XElement? properties = body.Element("bodyPr");
        double Inset(string name, double fallback) => (Number(properties, name) ?? fallback) / EmuPerPoint;
        return (Inset("lIns", 91440), Inset("tIns", 45720), Inset("rIns", 91440), Inset("bIns", 45720));
    }

    private static string RunCss(XElement? properties, DocxTheme? theme, double fontScale, string? defaultColor, string? defaultFace)
    {
        var css = new StringBuilder();
        double size = (Number(properties, "sz") ?? 1800) / 100 * fontScale;
        css.Append("font-size:").Append(Points(size)).Append("pt;");
        if ((string?)properties?.Attribute("b") is "1" or "true") css.Append("font-weight:bold;");
        if ((string?)properties?.Attribute("i") is "1" or "true") css.Append("font-style:italic;");
        string decoration = ((string?)properties?.Attribute("u") is { } u && u != "none" ? "underline " : "") + ((string?)properties?.Attribute("strike") is { } strike && strike != "noStrike" ? "line-through" : "");
        if (decoration.Length > 0) css.Append("text-decoration:").Append(decoration.Trim()).Append(';');
        if (Number(properties, "baseline") is { } baseline && baseline != 0) css.Append("vertical-align:").Append(baseline > 0 ? "super" : "sub").Append(";font-size:").Append(Points(size * 2 / 3)).Append("pt;");
        string? color = properties?.Element("solidFill") is { } fill ? DocxDrawingColor.Css(fill, theme) : defaultColor;
        if (color is not null) css.Append("color:").Append(color).Append(';');
        var faces = new List<string>();
        foreach (string slot in new[] { "latin", "cs", "ea" })
            if ((string?)properties?.Element(slot)?.Attribute("typeface") is { Length: > 0 } face)
                faces.Add(face switch
                {
                    "+mn-lt" => theme?.FontFor("minorHAnsi") ?? face, "+mj-lt" => theme?.FontFor("majorHAnsi") ?? face,
                    "+mn-cs" => theme?.FontFor("minorBidi") ?? face, "+mj-cs" => theme?.FontFor("majorBidi") ?? face,
                    "+mn-ea" => theme?.FontFor("minorEastAsia") ?? face, "+mj-ea" => theme?.FontFor("majorEastAsia") ?? face,
                    _ => face,
                });
        if (faces.Count == 0 && defaultFace is not null) faces.Add(defaultFace);
        if (faces.Count > 0) css.Append("font-family:").Append(string.Join(",", faces.Distinct().Select(f => "'" + f.Replace("'", "") + "'"))).Append(';');
        return css.ToString();
    }

    private static double? Number(XElement? element, string name)
        => double.TryParse((string?)element?.Attribute(name), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : null;

    private static string Points(double value) => Math.Round(value, 2).ToString("0.##", CultureInfo.InvariantCulture);
}
