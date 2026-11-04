using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// An Office 2016 chart part (cx:chartSpace, [MS-ODRAWXML] 2.24, schema section 5.22: CT_ChartSpace, CT_ChartData, CT_Data,
/// CT_StringDimension / CT_NumericDimension with their CT_StringLevel / CT_NumericLevel caches, CT_Series layoutId
/// ST_SeriesLayout, CT_SeriesLayoutProperties) drawn as one SVG from the cached data, never the workbook:
/// - clusteredColumn = histogram: CT_Binning binSize / binCount / underflow / overflow / intervalClosed (default right-closed;
///   no size or count = "Automatic", Scott's normal reference rule h = 3.49 s n^-1/3, as Excel's "Create a histogram" help
///   states), or CT_Aggregation = one column per distinct category with the values summed; a paretoLine series (ownerIdx)
///   adds the cumulative percentage of the columns sorted largest first on a 0-100 % axis;
/// - waterfall: running total; CT_Subtotals points drawn from 0 to their value; connector lines unless visibility says not;
/// - funnel: one centred bar per category, width by value;
/// - boxWhisker: per category the quartiles by CT_Statistics quartileMethod (exclusive default: position (n+1)p, inclusive:
///   (n-1)p+1 - Excel's QUARTILE.EXC / QUARTILE.INC), whiskers to the furthest values within 1.5 IQR of the box, outliers
///   beyond as points, mean marker and line per CT_SeriesElementVisibilities (defaults: marker and outliers shown);
/// - treemap: squarified rectangles (Bruls, Huizing, van Wijk 2000) of the leaf values, grouped by the top category level;
///   parentLabelLayout banner / overlapping / none;
/// - sunburst: one ring per category level from the centre, each node's angle its leaves' share;
/// - regionMap: the polygons of the cached geography (CT_GeoCache with a "clear" CT_Clear: the data points' locations,
///   the point-to-entity mapping and each entity's CT_GeoPolygon pcaRings, decoded with the Bing Maps point compression
///   algorithm as the CT_GeoPolygon description names it), projected (ST_GeoProjectionType mercator, miller, robinson by
///   Robinson's table, albers) and coloured between CT_ValueColors min / mid / max by value. A geography cached only as
///   "binary" ([MS-ODRAWXML] 2.24.3.29: base64 with no format specified) cannot be drawn: <see cref="CanDraw"/> is false
///   and the renderer shows the drawing's mc:Fallback picture.
/// Multi-level categories: the first cx:lvl is the leaf level, the following ones its parents (as CT_StringDimension lists
/// them for hierarchical layouts); a parent label repeats until the next non-empty one.
/// Layout, axis scaling, the albers standard parallels (1/6 and 5/6 of the data's latitude span), the projection when none is
/// set (mercator) and the default colours (series and points from accent 1 on; waterfall increase / decrease / total =
/// accent 1 / 2 / 3; region maps from accent 1 at 20 % to accent 1) are this library's own: the format does not state them.
/// </summary>
internal sealed class DocxChartEx
{
    private sealed class Series
    {
        public string Layout = "";
        public string? Name;
        public bool Hidden;
        public int? OwnerIndex;
        public XElement? Element;
        public List<List<string>> Categories = new();     // levels, leaf first
        public List<double?> Values = new();
        public string? FormatCode;
        public Dictionary<int, XElement> Points = new();
        public XElement? Layouts, Labels;
    }

    private readonly List<Series> _series = new();
    private string? _title;
    private string? _legendPosition;
    private double? _valueMin, _valueMax, _valueUnit;
    private string? _categoryTitle, _valueTitle;
    private readonly DocxTheme? _theme;

    private DocxChartEx(DocxTheme? theme) => _theme = theme;

    /// <summary>The chart of a cx:chartSpace (namespace-free); null for a dashboard or a chart without series.</summary>
    public static DocxChartEx? Read(XElement chartSpace, DocxTheme? theme)
    {
        XElement? chart = chartSpace.Element("chart");
        if (chart is null) return null;
        var model = new DocxChartEx(theme);
        var data = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (XElement d in chartSpace.Element("chartData")?.Elements("data") ?? []) if ((string?)d.Attribute("id") is { } id) data[id] = d;
        if (chart.Element("title") is { } title) model._title = Text(title.Element("tx"));
        if (chart.Element("legend") is { } legend) model._legendPosition = (string?)legend.Attribute("pos") ?? "r";
        foreach (XElement axis in chart.Element("plotArea")?.Elements("axis") ?? [])
        {
            if ((string?)axis.Attribute("hidden") is "1" or "true") continue;
            string? axisTitle = Text(axis.Element("title")?.Element("tx"));
            if (axis.Element("valScaling") is { } scaling)
            {
                model._valueMin ??= Double((string?)scaling.Attribute("min")); model._valueMax ??= Double((string?)scaling.Attribute("max"));
                model._valueUnit ??= Double((string?)scaling.Attribute("majorUnit"));
                model._valueTitle ??= axisTitle;
            }
            else model._categoryTitle ??= axisTitle;
        }
        foreach (XElement element in chart.Element("plotArea")?.Element("plotAreaRegion")?.Elements("series") ?? [])
        {
            var series = new Series
            {
                Layout = (string?)element.Attribute("layoutId") ?? "", Name = Text(element.Element("tx")), Hidden = (string?)element.Attribute("hidden") is "1" or "true",
                OwnerIndex = int.TryParse((string?)element.Attribute("ownerIdx"), out int owner) ? owner : null, Element = element,
                Layouts = element.Element("layoutPr"), Labels = element.Element("dataLabels"),
            };
            foreach (XElement point in element.Elements("dataPt"))
                if (int.TryParse((string?)point.Attribute("idx"), out int idx) && point.Element("spPr") is { } properties) series.Points[idx] = properties;
            if ((string?)element.Element("dataId")?.Attribute("val") is { } dataId && data.TryGetValue(dataId, out XElement? set))
            {
                foreach (XElement dimension in set.Elements())
                {
                    string type = (string?)dimension.Attribute("type") ?? "";
                    if (dimension.Name.LocalName == "strDim" && type == "cat")
                        foreach (XElement level in dimension.Elements("lvl")) series.Categories.Add(StringLevel(level));
                    else if (dimension.Name.LocalName == "numDim" && type is "val" or "size" && series.Values.Count == 0 && dimension.Element("lvl") is { } numbers)
                    {
                        series.FormatCode = (string?)numbers.Attribute("formatCode");
                        series.Values = NumberLevel(numbers);
                    }
                }
            }
            model._series.Add(series);
        }
        return model._series.Count == 0 ? null : model;
    }

    private static List<string> StringLevel(XElement level)
    {
        int count = int.TryParse((string?)level.Attribute("ptCount"), out int c) ? c : 0;
        var values = new string[count];
        foreach (XElement point in level.Elements("pt"))
            if (int.TryParse((string?)point.Attribute("idx"), out int idx) && idx >= 0 && idx < count) values[idx] = point.Value;
        return values.Select(v => v ?? "").ToList();
    }

    private static List<double?> NumberLevel(XElement level)
    {
        int count = int.TryParse((string?)level.Attribute("ptCount"), out int c) ? c : 0;
        var values = new double?[count];
        foreach (XElement point in level.Elements("pt"))
            if (int.TryParse((string?)point.Attribute("idx"), out int idx) && idx >= 0 && idx < count && Double(point.Value) is { } value) values[idx] = value;
        return values.ToList();
    }

    private static string? Text(XElement? tx)
    {
        if (tx is null) return null;
        if (tx.Element("txData")?.Element("v") is { } v) return v.Value;
        if (tx.Element("rich") is { } rich) return string.Join("\n", rich.Elements("p").Select(p => string.Concat(p.Elements().Where(e => e.Name.LocalName is "r" or "fld").Select(r => (string?)r.Element("t")))));
        return null;
    }

    private static double? Double(string? text) => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : null;

    private Series? Main => _series.FirstOrDefault(s => !s.Hidden && s.Layout != "paretoLine");

    /// <summary>False for a region map whose geography is not cached in the documented ("clear") form.</summary>
    public bool CanDraw => Main is { } main && (main.Layout != "regionMap" || Geography(main) is not null);

    // ── colours ─────────────────────────────────────────────────────────────────────

    private string Accent(int index, string? modification = null, int modValue = 0)
    {
        var color = new XElement("schemeClr", new XAttribute("val", "accent" + ((index % 6) + 1).ToString(CultureInfo.InvariantCulture)));
        if (modification is not null) color.Add(new XElement(modification, new XAttribute("val", modValue.ToString(CultureInfo.InvariantCulture))));
        return DocxDrawingColor.Css(new XElement("c", color), _theme) ?? "#4472C4";
    }

    private string PointFill(Series series, int point, string fallback)
        => series.Points.TryGetValue(point, out XElement? properties) && properties.Element("solidFill") is { } solid && DocxDrawingColor.Css(solid, _theme) is { } own ? own
         : series.Element?.Element("spPr")?.Element("solidFill") is { } seriesFill && DocxDrawingColor.Css(seriesFill, _theme) is { } css ? css : fallback;

    // ── texts ───────────────────────────────────────────────────────────────────────

    /// <summary>The texts the drawing shows from the document: title, axis titles, legend, category labels (bins for a histogram), data labels.</summary>
    public IEnumerable<string> Texts()
    {
        var texts = new List<string>();
        var svg = new StringBuilder();
        Draw(svg, 400, 300, texts);
        return texts;
    }

    public string Svg(double width, double height)
    {
        var svg = new StringBuilder();
        Draw(svg, Math.Max(1, width), Math.Max(1, height), null);
        return svg.ToString();
    }

    // ── drawing ─────────────────────────────────────────────────────────────────────

    private string _textColor = "#595959";

    private void Label(StringBuilder svg, List<string>? texts, double x, double y, string text, double size, string anchor, string? color = null, bool searchable = true)
    {
        if (texts is not null && searchable && text.Length > 0) texts.Add(text);
        svg.Append("<text x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" font-size=\"").Append(N(size)).Append("\" fill=\"").Append(color ?? _textColor)
           .Append("\" text-anchor=\"").Append(anchor).Append("\">").Append(WordHtmlText.Encode(text)).Append("</text>");
    }

    private void Draw(StringBuilder svg, double width, double height, List<string>? texts)
    {
        _textColor = DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", "tx1"), new XElement("lumMod", new XAttribute("val", "65000")), new XElement("lumOff", new XAttribute("val", "35000")))), _theme) ?? "#595959";
        svg.Append("<svg class=\"word-chart\" aria-hidden=\"true\" viewBox=\"0 0 ").Append(N(width)).Append(' ').Append(N(height)).Append("\" width=\"").Append(N(width)).Append("pt\" height=\"").Append(N(height))
           .Append("pt\" style=\"position:absolute;left:0;top:0;width:100%;height:100%;overflow:hidden\" font-family=\"Calibri, sans-serif\"><rect width=\"").Append(N(width)).Append("\" height=\"").Append(N(height)).Append("\" fill=\"#FFFFFF\"/>");
        double top = 6, bottom = height - 6, left = 6, right = width - 6;
        if (_title is not null)
        {
            foreach (string line in _title.Split('\n')) { Label(svg, texts, width / 2, top + 14, line, 14, "middle"); top += 16.8; }
            top += 6;
        }
        Series? main = Main;
        if (main is null) { svg.Append("</svg>"); return; }
        // legend (CT_Legend pos, default r): the series names; for treemap / sunburst the top-level categories (their colours)
        if (_legendPosition is not null && main.Layout != "regionMap")
        {
            var entries = main.Layout is "treemap" or "sunburst"
                ? Hierarchy(main).Select(l => l.Path[0]).Distinct().Select((name, i) => (name, Accent(i))).ToList()
                : _series.Where(s => !s.Hidden).Select((s, i) => (s.Name ?? "", PointFill(s, -1, Accent(i)))).Where(e => e.Item1.Length > 0).ToList();
            if (entries.Count > 0)
            {
                double legendWidth = Math.Min(width / 3, entries.Max(e => e.Item1.Length) * 9 * 0.55 + 18);
                double y = (top + bottom) / 2 - entries.Count * 13 / 2.0, x = _legendPosition == "l" ? left : right - legendWidth;
                foreach ((string name, string color) in entries)
                {
                    svg.Append("<rect x=\"").Append(N(x + 2)).Append("\" y=\"").Append(N(y + 1)).Append("\" width=\"7\" height=\"7\" fill=\"").Append(color).Append("\"/>");
                    Label(svg, texts, x + 14, y + 8, name, 9, "start");
                    y += 13;
                }
                if (_legendPosition == "l") left += legendWidth + 6; else right -= legendWidth + 6;
            }
        }
        switch (main.Layout)
        {
            case "clusteredColumn": DrawHistogram(svg, texts, main, left, top, right, bottom); break;
            case "waterfall": DrawWaterfall(svg, texts, main, left, top, right, bottom); break;
            case "funnel": DrawFunnel(svg, texts, main, left, top, right, bottom); break;
            case "boxWhisker": DrawBoxes(svg, texts, left, top, right, bottom); break;
            case "treemap": DrawTreemap(svg, texts, main, left, top, right, bottom); break;
            case "sunburst": DrawSunburst(svg, texts, main, left, top, right, bottom); break;
            case "regionMap": DrawRegionMap(svg, texts, main, left, top, right, bottom); break;
        }
        svg.Append("</svg>");
    }

    private (double Min, double Max, double Step) ValueScale(double dataMin, double dataMax)
    {
        double min = _valueMin ?? Math.Min(0, dataMin), max = _valueMax ?? Math.Max(0, dataMax);
        if (max <= min) max = min + 1;
        double step = _valueUnit ?? NiceStep((max - min) / 5);
        if (_valueMin is null) min = Math.Floor(min / step) * step;
        if (_valueMax is null) max = Math.Ceiling(max / step) * step;
        if (max <= min) max = min + step;
        return (min, max, step);
    }

    private static double NiceStep(double raw)
    {
        if (raw <= 0) return 1;
        double magnitude = Math.Pow(10, Math.Floor(Math.Log10(raw))), fraction = raw / magnitude;
        return (fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10) * magnitude;
    }

    /// <summary>Value axis with gridlines and labels on the left of [left, right] x [top, bottom]; returns the plot's left edge and the value-to-y map.</summary>
    private (double Left, Func<double, double> Y) ValueAxis(StringBuilder svg, double min, double max, double step, double left, double top, double bottom, double right, string? format, List<string>? texts)
    {
        var labels = new List<(double, string)>();
        for (double v = min; v <= max + step / 1000; v += step) labels.Add((v, DocxChart.FormatNumber(Math.Round(v, 10), format)));
        double labelWidth = labels.Max(l => l.Item2.Length) * 9 * 0.55 + 4;
        if (_valueTitle is not null) { Label(svg, texts, left + 6, (top + bottom) / 2, _valueTitle, 10, "middle"); left += 14; }
        left += labelWidth;
        double Y(double v) => bottom - (v - min) / (max - min) * (bottom - top);
        foreach ((double v, string text) in labels)
        {
            svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(Y(v))).Append('H').Append(N(right)).Append("\" stroke=\"#D9D9D9\" stroke-width=\"0.75\"/>");
            Label(svg, null, left - 4, Y(v) + 3, text, 9, "end");
        }
        return (left, Y);
    }

    private static string Range(double from, double to, bool leftOpen, bool rightOpen)
        => (leftOpen ? "(" : "[") + DocxChart.FormatNumber(Math.Round(from, 10), "General") + ", " + DocxChart.FormatNumber(Math.Round(to, 10), "General") + (rightOpen ? ")" : "]");

    /// <summary>The columns of a histogram (bins or aggregated categories): label and height.</summary>
    private List<(string Label, double Value)> Columns(Series series)
    {
        var columns = new List<(string, double)>();
        XElement? binning = series.Layouts?.Element("binning");
        if (series.Layouts?.Element("aggregation") is not null)
        {
            var sums = new Dictionary<string, double>();
            var order = new List<string>();
            List<string> categories = series.Categories.Count > 0 ? series.Categories[0] : new();
            for (int i = 0; i < series.Values.Count; i++)
            {
                string key = i < categories.Count ? categories[i] : (i + 1).ToString(CultureInfo.InvariantCulture);
                if (!sums.ContainsKey(key)) { sums[key] = 0; order.Add(key); }
                sums[key] += series.Values[i] ?? 0;
            }
            return order.Select(k => (k, sums[k])).ToList();
        }
        var values = series.Values.Where(v => v is not null).Select(v => v!.Value).OrderBy(v => v).ToList();
        if (values.Count == 0) return columns;
        bool rightClosed = (string?)binning?.Attribute("intervalClosed") != "l";
        double? underflow = Double((string?)binning?.Attribute("underflow")), overflow = Double((string?)binning?.Attribute("overflow"));
        double low = underflow ?? values[0], high = overflow ?? values[values.Count - 1];
        double size;
        if (Double(binning?.Element("binSize")?.Value) is { } binSize && binSize > 0) size = binSize;
        else if (int.TryParse(binning?.Element("binCount")?.Value, out int binCount) && binCount > 0) size = (high - low) / binCount;
        else
        {
            double mean = values.Average(), deviation = Math.Sqrt(values.Sum(v => (v - mean) * (v - mean)) / Math.Max(1, values.Count - 1));
            size = 3.49 * deviation * Math.Pow(values.Count, -1.0 / 3);
        }
        if (!(size > 0)) size = Math.Max(1, high - low);
        int bins = Math.Max(1, (int)Math.Ceiling((high - low) / size - 1e-9));
        if (underflow is not null) columns.Add(("≤" + DocxChart.FormatNumber(low, "General"), values.Count(v => rightClosed ? v <= low : v < low)));
        for (int b = 0; b < bins; b++)
        {
            double from = low + b * size, to = b == bins - 1 ? Math.Max(high, low + (b + 1) * size) : low + (b + 1) * size;
            bool first = b == 0 && underflow is null, last = b == bins - 1 && overflow is null;
            int count = values.Count(v => rightClosed ? (first ? v >= from : v > from) && v <= to : v >= from && (last ? v <= to : v < to));
            columns.Add((Range(from, to, rightClosed && !first, !rightClosed && !last), count));
        }
        if (overflow is not null) columns.Add((">" + DocxChart.FormatNumber(high, "General"), values.Count(v => rightClosed ? v > high : v >= high)));
        return columns;
    }

    private void DrawHistogram(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        List<(string Label, double Value)> columns = Columns(series);
        Series? pareto = _series.FirstOrDefault(s => s.Layout == "paretoLine" && !s.Hidden);
        if (pareto is not null) columns = columns.OrderByDescending(c => c.Value).ToList();
        if (columns.Count == 0) return;
        if (pareto is not null) right -= 30;
        bottom -= 14;
        if (_categoryTitle is not null) { Label(svg, texts, (left + right) / 2, bottom + 12, _categoryTitle, 10, "middle"); bottom -= 14; }
        (double min, double max, double step) = ValueScale(0, columns.Max(c => c.Value));
        (double plotLeft, Func<double, double> y) = ValueAxis(svg, min, max, step, left, top, bottom, right, null, texts);
        double slot = (right - plotLeft) / columns.Count;
        string fill = PointFill(series, -1, Accent(0));
        for (int i = 0; i < columns.Count; i++)
        {
            double x = plotLeft + i * slot;
            svg.Append("<rect x=\"").Append(N(x + 0.5)).Append("\" y=\"").Append(N(y(columns[i].Value))).Append("\" width=\"").Append(N(slot - 1)).Append("\" height=\"").Append(N(y(min) - y(columns[i].Value))).Append("\" fill=\"").Append(PointFill(series, i, fill)).Append("\"/>");
            // bin ranges are computed, not document text: shown, not searched; aggregated categories are the document's
            Label(svg, texts, x + slot / 2, bottom + 11, columns[i].Label, 9, "middle", searchable: series.Layouts?.Element("aggregation") is not null);
        }
        if (pareto is null) return;
        double total = columns.Sum(c => c.Value), running = 0;
        var d = new StringBuilder();
        for (int i = 0; i < columns.Count; i++)
        {
            running += columns[i].Value;
            double share = total > 0 ? running / total : 0, py = bottom - share * (bottom - top);
            d.Append(i == 0 ? 'M' : 'L').Append(N(plotLeft + (i + 0.5) * slot)).Append(' ').Append(N(py));
        }
        svg.Append("<path d=\"").Append(d).Append("\" fill=\"none\" stroke=\"").Append(PointFill(pareto, -1, Accent(1))).Append("\" stroke-width=\"2.25\"/>");
        for (int p = 0; p <= 100; p += 20) Label(svg, null, right + 4, bottom - p / 100.0 * (bottom - top) + 3, p.ToString(CultureInfo.InvariantCulture) + "%", 9, "start");
    }

    private List<string> Leaves(Series series)
        => series.Categories.Count > 0 ? series.Categories[0] : Enumerable.Range(1, series.Values.Count).Select(i => i.ToString(CultureInfo.InvariantCulture)).ToList();

    private void DrawWaterfall(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        var totals = new HashSet<int>(series.Layouts?.Element("subtotals")?.Elements("idx").Select(e => int.TryParse((string?)e.Attribute("val"), out int v) ? v : -1) ?? []);
        List<string> categories = Leaves(series);
        int count = series.Values.Count;
        var bars = new List<(double From, double To, int Kind)>();   // kind 0 increase, 1 decrease, 2 total
        double running = 0;
        for (int i = 0; i < count; i++)
        {
            double value = series.Values[i] ?? 0;
            if (totals.Contains(i)) { bars.Add((0, value, 2)); running = value; }
            else { bars.Add((running, running + value, value >= 0 ? 0 : 1)); running += value; }
        }
        if (bars.Count == 0) return;
        bottom -= 14;
        (double min, double max, double step) = ValueScale(bars.Min(b => Math.Min(b.From, b.To)), bars.Max(b => Math.Max(b.From, b.To)));
        (double plotLeft, Func<double, double> y) = ValueAxis(svg, min, max, step, left, top, bottom, right, series.FormatCode, texts);
        double slot = (right - plotLeft) / count, barWidth = slot * 0.6;
        bool connectors = (string?)series.Layouts?.Element("visibility")?.Attribute("connectorLines") is not ("0" or "false");
        for (int i = 0; i < count; i++)
        {
            (double from, double to, int kind) = bars[i];
            double x = plotLeft + i * slot + (slot - barWidth) / 2;
            string fill = PointFill(series, i, Accent(kind));
            svg.Append("<rect x=\"").Append(N(x)).Append("\" y=\"").Append(N(y(Math.Max(from, to)))).Append("\" width=\"").Append(N(barWidth)).Append("\" height=\"").Append(N(Math.Abs(y(from) - y(to)))).Append("\" fill=\"").Append(fill).Append("\"/>");
            if (connectors && i + 1 < count) svg.Append("<path d=\"M").Append(N(x + barWidth)).Append(' ').Append(N(y(to))).Append('H').Append(N(x + slot)).Append("\" stroke=\"#A6A6A6\" stroke-width=\"0.75\"/>");
            Label(svg, texts, x + barWidth / 2, bottom + 11, i < categories.Count ? categories[i] : "", 9, "middle");
            Label(svg, texts, x + barWidth / 2, y(Math.Max(from, to)) - 2, DocxChart.FormatNumber(series.Values[i] ?? 0, series.FormatCode), 9, "middle");
        }
    }

    private void DrawFunnel(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        List<string> categories = Leaves(series);
        int count = series.Values.Count;
        if (count == 0) return;
        double labelWidth = categories.DefaultIfEmpty("").Max(c => c.Length) * 9 * 0.55 + 6, max = series.Values.Max(v => Math.Abs(v ?? 0));
        double plotLeft = left + labelWidth, plotWidth = right - plotLeft, slot = (bottom - top) / count, centre = plotLeft + plotWidth / 2;
        for (int i = 0; i < count; i++)
        {
            double value = Math.Abs(series.Values[i] ?? 0), w = max > 0 ? value / max * plotWidth : 0, y = top + i * slot;
            svg.Append("<rect x=\"").Append(N(centre - w / 2)).Append("\" y=\"").Append(N(y + slot * 0.1)).Append("\" width=\"").Append(N(w)).Append("\" height=\"").Append(N(slot * 0.8)).Append("\" fill=\"").Append(PointFill(series, i, Accent(0))).Append("\"/>");
            Label(svg, texts, plotLeft - 4, y + slot / 2 + 3, i < categories.Count ? categories[i] : "", 9, "end");
            Label(svg, texts, centre, y + slot / 2 + 3, DocxChart.FormatNumber(series.Values[i] ?? 0, series.FormatCode), 9, "middle", "#FFFFFF");
        }
    }

    /// <summary>Quartile at <paramref name="p"/> of sorted values: QUARTILE.EXC position (n+1)p, QUARTILE.INC (n-1)p+1 (1-based), interpolated.</summary>
    private static double Quartile(List<double> sorted, double p, bool exclusive)
    {
        int n = sorted.Count;
        double position = exclusive ? (n + 1) * p : (n - 1) * p + 1;
        if (position <= 1) return sorted[0];
        if (position >= n) return sorted[n - 1];
        int lower = (int)Math.Floor(position);
        return sorted[lower - 1] + (position - lower) * (sorted[lower] - sorted[lower - 1]);
    }

    private void DrawBoxes(StringBuilder svg, List<string>? texts, double left, double top, double right, double bottom)
    {
        List<Series> visible = _series.Where(s => !s.Hidden && s.Layout == "boxWhisker").ToList();
        var categories = new List<string>();
        foreach (Series s in visible) foreach (string c in s.Categories.Count > 0 ? s.Categories[0] : new List<string> { "" }) if (!categories.Contains(c)) categories.Add(c);
        if (categories.Count == 0) categories.Add("");
        var all = visible.SelectMany(s => s.Values).Where(v => v is not null).Select(v => v!.Value).ToList();
        if (all.Count == 0) return;
        bottom -= 14;
        (double min, double max, double step) = ValueScale(Math.Min(0, all.Min()), all.Max());
        (double plotLeft, Func<double, double> y) = ValueAxis(svg, min, max, step, left, top, bottom, right, visible[0].FormatCode, texts);
        double slot = (right - plotLeft) / categories.Count, boxWidth = slot * 0.6 / Math.Max(1, visible.Count);
        for (int c = 0; c < categories.Count; c++)
        {
            Label(svg, texts, plotLeft + (c + 0.5) * slot, bottom + 11, categories[c], 9, "middle");
            for (int s = 0; s < visible.Count; s++)
            {
                Series series = visible[s];
                List<string> names = series.Categories.Count > 0 ? series.Categories[0] : new();
                var values = series.Values.Select((v, i) => (v, i)).Where(p => p.v is not null && (names.Count == 0 || (p.i < names.Count && names[p.i] == categories[c]))).Select(p => p.v!.Value).OrderBy(v => v).ToList();
                if (values.Count == 0) continue;
                bool exclusive = (string?)series.Layouts?.Element("statistics")?.Attribute("quartileMethod") != "inclusive";
                double q1 = Quartile(values, 0.25, exclusive), median = Quartile(values, 0.5, exclusive), q3 = Quartile(values, 0.75, exclusive), iqr = q3 - q1;
                double low = values.Where(v => v >= q1 - 1.5 * iqr).DefaultIfEmpty(q1).Min(), high = values.Where(v => v <= q3 + 1.5 * iqr).DefaultIfEmpty(q3).Max();
                double x = plotLeft + c * slot + (slot - boxWidth * visible.Count) / 2 + s * boxWidth, mid = x + boxWidth / 2;
                string fill = PointFill(series, -1, Accent(s)), line = "#595959";
                XElement? visibility = series.Layouts?.Element("visibility");
                bool On(string name, bool fallback) => (string?)visibility?.Attribute(name) is { } v ? v is "1" or "true" : fallback;
                svg.Append("<path d=\"M").Append(N(mid)).Append(' ').Append(N(y(high))).Append('V').Append(N(y(q3))).Append('M').Append(N(mid)).Append(' ').Append(N(y(q1))).Append('V').Append(N(y(low)))
                   .Append('M').Append(N(mid - boxWidth / 4)).Append(' ').Append(N(y(high))).Append('h').Append(N(boxWidth / 2)).Append('M').Append(N(mid - boxWidth / 4)).Append(' ').Append(N(y(low))).Append('h').Append(N(boxWidth / 2))
                   .Append("\" stroke=\"").Append(line).Append("\" stroke-width=\"0.75\" fill=\"none\"/>");
                svg.Append("<rect x=\"").Append(N(x + boxWidth * 0.1)).Append("\" y=\"").Append(N(y(q3))).Append("\" width=\"").Append(N(boxWidth * 0.8)).Append("\" height=\"").Append(N(y(q1) - y(q3))).Append("\" fill=\"").Append(fill).Append("\" stroke=\"").Append(line).Append("\" stroke-width=\"0.75\"/>");
                svg.Append("<path d=\"M").Append(N(x + boxWidth * 0.1)).Append(' ').Append(N(y(median))).Append('h').Append(N(boxWidth * 0.8)).Append("\" stroke=\"").Append(line).Append("\" stroke-width=\"1\"/>");
                if (On("meanMarker", true))
                {
                    double mean = values.Average();
                    svg.Append("<path d=\"M").Append(N(mid - 2)).Append(' ').Append(N(y(mean) - 2)).Append("l4 4m0 -4l-4 4\" stroke=\"").Append(line).Append("\" stroke-width=\"0.75\"/>");
                }
                if (On("outliers", true))
                    foreach (double v in values.Where(v => v < low || v > high))
                        svg.Append("<circle cx=\"").Append(N(mid)).Append("\" cy=\"").Append(N(y(v))).Append("\" r=\"1.5\" fill=\"none\" stroke=\"").Append(fill).Append("\"/>");
            }
        }
    }

    /// <summary>The hierarchy of a treemap / sunburst: the path of each leaf (root first) and its value.</summary>
    private List<(List<string> Path, double Value, int Index)> Hierarchy(Series series)
    {
        var leaves = new List<(List<string>, double, int)>();
        int levels = series.Categories.Count;
        var carried = new string[levels];
        for (int i = 0; i < series.Values.Count; i++)
        {
            var path = new List<string>();
            for (int level = levels - 1; level >= 0; level--)
            {
                string name = i < series.Categories[level].Count ? series.Categories[level][i] : "";
                if (name.Length > 0) carried[level] = name;
                else if (level > 0) name = carried[level] ?? "";
                if (name.Length > 0) path.Add(name);
            }
            if (path.Count == 0) path.Add((i + 1).ToString(CultureInfo.InvariantCulture));
            leaves.Add((path, Math.Max(0, series.Values[i] ?? 0), i));
        }
        return leaves;
    }

    private void DrawTreemap(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        var leaves = Hierarchy(series);
        var groups = leaves.GroupBy(l => l.Path.Count > 1 ? l.Path[0] : "").ToList();
        string parentLayout = (string?)series.Layouts?.Element("parentLabelLayout")?.Attribute("val") ?? "overlapping";
        var groupRects = Squarify(groups.Select(g => g.Sum(l => l.Value)).ToList(), left, top, right - left, bottom - top);
        for (int g = 0; g < groups.Count; g++)
        {
            (double gx, double gy, double gw, double gh) = groupRects[g];
            string color = Accent(g);
            double innerTop = gy;
            if (groups[g].Key.Length > 0 && parentLayout == "banner")
            {
                svg.Append("<rect x=\"").Append(N(gx)).Append("\" y=\"").Append(N(gy)).Append("\" width=\"").Append(N(gw)).Append("\" height=\"14\" fill=\"").Append(color).Append("\" stroke=\"#FFFFFF\"/>");
                Label(svg, texts, gx + 3, gy + 10, groups[g].Key, 9, "start", "#FFFFFF");
                innerTop += 14;
            }
            var items = groups[g].ToList();
            var rects = Squarify(items.Select(l => l.Value).ToList(), gx, innerTop, gw, gy + gh - innerTop);
            for (int i = 0; i < items.Count; i++)
            {
                (double x, double y, double w, double h) = rects[i];
                svg.Append("<rect x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" width=\"").Append(N(w)).Append("\" height=\"").Append(N(h)).Append("\" fill=\"").Append(PointFill(series, items[i].Index, color)).Append("\" stroke=\"#FFFFFF\"/>");
                Label(svg, texts, x + 3, y + 11, items[i].Path[items[i].Path.Count - 1], 9, "start", "#FFFFFF");
            }
            if (groups[g].Key.Length > 0 && parentLayout == "overlapping") Label(svg, texts, gx + 3, gy + gh / 2, groups[g].Key, 12, "start", "#FFFFFF");
        }
    }

    /// <summary>Squarified treemap layout (Bruls, Huizing, van Wijk, "Squarified Treemaps", 2000) of values, largest first, into a rectangle; rectangles in input order.</summary>
    private static List<(double X, double Y, double W, double H)> Squarify(List<double> values, double x, double y, double w, double h)
    {
        var result = new (double, double, double, double)[values.Count];
        double total = values.Sum();
        if (total <= 0 || w <= 0 || h <= 0) return result.ToList();
        var order = Enumerable.Range(0, values.Count).OrderByDescending(i => values[i]).ToList();
        double scale = w * h / total;
        int start = 0;
        while (start < order.Count)
        {
            double side = Math.Min(w, h);
            var row = new List<int> { order[start] };
            double Worst(List<int> r)
            {
                double sum = r.Sum(i => values[i] * scale), maxV = r.Max(i => values[i] * scale), minV = r.Min(i => values[i] * scale);
                if (sum <= 0 || minV <= 0) return double.MaxValue;
                return Math.Max(side * side * maxV / (sum * sum), sum * sum / (side * side * minV));
            }
            int next = start + 1;
            while (next < order.Count)
            {
                var candidate = new List<int>(row) { order[next] };
                if (Worst(candidate) > Worst(row)) break;
                row = candidate; next++;
            }
            double rowArea = row.Sum(i => values[i] * scale), thickness = rowArea / side, offset = 0;
            foreach (int i in row)
            {
                double length = thickness > 0 ? values[i] * scale / thickness : 0;
                result[i] = w >= h ? (x, y + offset, thickness, length) : (x + offset, y, length, thickness);
                offset += length;
            }
            if (w >= h) { x += thickness; w -= thickness; } else { y += thickness; h -= thickness; }
            start = next;
        }
        return result.ToList();
    }

    private void DrawSunburst(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        var leaves = Hierarchy(series);
        double total = leaves.Sum(l => l.Value);
        if (total <= 0) return;
        int depth = leaves.Max(l => l.Path.Count);
        double cx = (left + right) / 2, cy = (top + bottom) / 2, radius = Math.Min(right - left, bottom - top) / 2, ring = radius / (depth + 0.5), hole = ring / 2;
        var topLevel = leaves.Select(l => l.Path[0]).Distinct().ToList();
        for (int level = 0; level < depth; level++)
        {
            double angle = -90;
            var nodes = new List<(string Name, double Value, string Root)>();
            foreach (var leaf in leaves)
            {
                string name = level < leaf.Path.Count ? leaf.Path[level] : "";
                string key = string.Join("\u0001", leaf.Path.Take(level + 1));
                if (nodes.Count > 0 && nodes[nodes.Count - 1].Name == key) nodes[nodes.Count - 1] = (key, nodes[nodes.Count - 1].Value + leaf.Value, leaf.Path[0]);
                else nodes.Add((key, level < leaf.Path.Count ? leaf.Value : -leaf.Value, leaf.Path[0]));
            }
            double inner = hole + level * ring, outer = inner + ring;
            foreach ((string key, double value, string root) in nodes)
            {
                double sweep = Math.Abs(value) / total * 360;
                if (value > 0)
                {
                    double a0 = angle * Math.PI / 180, a1 = (angle + sweep) * Math.PI / 180;
                    string large = sweep > 180 ? "1" : "0";
                    svg.Append("<path d=\"M").Append(N(cx + inner * Math.Cos(a0))).Append(' ').Append(N(cy + inner * Math.Sin(a0))).Append('L').Append(N(cx + outer * Math.Cos(a0))).Append(' ').Append(N(cy + outer * Math.Sin(a0)))
                       .Append('A').Append(N(outer)).Append(' ').Append(N(outer)).Append(" 0 ").Append(large).Append(" 1 ").Append(N(cx + outer * Math.Cos(a1))).Append(' ').Append(N(cy + outer * Math.Sin(a1)))
                       .Append('L').Append(N(cx + inner * Math.Cos(a1))).Append(' ').Append(N(cy + inner * Math.Sin(a1)))
                       .Append('A').Append(N(inner)).Append(' ').Append(N(inner)).Append(" 0 ").Append(large).Append(" 0 ").Append(N(cx + inner * Math.Cos(a0))).Append(' ').Append(N(cy + inner * Math.Sin(a0)))
                       .Append("Z\" fill=\"").Append(Accent(topLevel.IndexOf(root))).Append("\" stroke=\"#FFFFFF\"/>");
                    double middle = (angle + sweep / 2) * Math.PI / 180, at = (inner + outer) / 2;
                    Label(svg, texts, cx + at * Math.Cos(middle), cy + at * Math.Sin(middle) + 3, key.Split('\u0001').Last(), 8, "middle", "#FFFFFF");
                }
                angle += sweep;
            }
        }
    }

    // ── region map ───────────────────────────────────────────────────────────────────

    private sealed record GeoEntity(string Id, List<List<(double Lat, double Lon)>> Rings);

    /// <summary>The "clear" geography of a region-map series: per data point, its entity's polygons; null when there is none.</summary>
    private List<GeoEntity?>? Geography(Series series)
    {
        XElement? clear = series.Layouts?.Element("geography")?.Element("geoCache")?.Element("clear");
        if (clear is null) return null;
        var polygons = new Dictionary<string, List<List<(double, double)>>>(StringComparer.Ordinal);
        foreach (XElement data in clear.Element("geoDataEntityQueryResults")?.Elements("geoDataEntityQueryResult").Select(r => r.Element("geoData")).OfType<XElement>() ?? [])
        {
            var rings = new List<List<(double, double)>>();
            foreach (XElement polygon in data.Element("geoPolygons")?.Elements("geoPolygon") ?? [])
                foreach (string part in ((string?)polygon.Attribute("pcaRings") ?? "").Split(','))
                    if (!int.TryParse(part, out _) && DecodeRing(part) is { Count: > 2 } ring) rings.Add(ring);
            if ((string?)data.Attribute("entityId") is { } id) polygons[id] = rings;
        }
        // data point -> its location -> the entity containing that location
        var locations = clear.Element("geoLocationQueryResults")?.Elements("geoLocationQueryResult").ToList() ?? new();
        var toEntity = clear.Element("geoDataPointToEntityQueryResults")?.Elements("geoDataPointToEntityQueryResult").ToList() ?? new();
        List<string> names = Leaves(series);
        var result = new List<GeoEntity?>();
        for (int i = 0; i < series.Values.Count; i++)
        {
            string name = i < names.Count ? names[i] : "";
            XElement? location = locations.Select(l => l.Element("geoLocations")?.Element("geoLocation")).FirstOrDefault(g => g is not null
                && ((string?)g.Attribute("entityName") == name || g.Parent?.Parent?.Element("geoLocationQuery")?.Attributes().Any(a => a.Value == name) == true));
            string? entityId = null;
            if (location is not null)
                entityId = toEntity.Where(t => t.Element("geoDataPointQuery") is { } q && Double((string?)q.Attribute("latitude")) == Double((string?)location.Attribute("latitude"))
                                                && Double((string?)q.Attribute("longitude")) == Double((string?)location.Attribute("longitude")))
                                   .Select(t => (string?)t.Element("geoDataPointToEntityQuery")?.Attribute("entityId")).FirstOrDefault();
            result.Add(entityId is not null && polygons.TryGetValue(entityId, out var rings) ? new GeoEntity(entityId, rings) : null);
        }
        return result.Any(e => e is not null) ? result : null;
    }

    private const string SafeCharacters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-";

    /// <summary>The Bing Maps point compression algorithm's decoder (learn.microsoft.com bingmaps "Geodata API", Decompression Algorithm): 5-bit groups, diagonal pair, zig-zag sign, delta, 1e-5 degrees.</summary>
    internal static List<(double Lat, double Lon)>? DecodeRing(string value)
    {
        var list = new List<(double, double)>();
        int index = 0, xsum = 0, ysum = 0;
        while (index < value.Length)
        {
            long n = 0;
            int k = 0;
            while (true)
            {
                if (index >= value.Length) return null;
                int b = SafeCharacters.IndexOf(value[index++]);
                if (b == -1) return null;
                n |= ((long)b & 31) << k;
                k += 5;
                if (b < 32) break;
            }
            int diagonal = (int)((Math.Sqrt(8 * n + 5) - 1) / 2);
            n -= diagonal * (diagonal + 1L) / 2;
            int ny = (int)n, nx = diagonal - ny;
            nx = (nx >> 1) ^ -(nx & 1);
            ny = (ny >> 1) ^ -(ny & 1);
            xsum += nx; ysum += ny;
            list.Add((ysum * 0.00001, xsum * 0.00001));
        }
        return list;
    }

    private static readonly double[] RobinsonX = [1.0000, 0.9986, 0.9954, 0.9900, 0.9822, 0.9730, 0.9600, 0.9427, 0.9216, 0.8962, 0.8679, 0.8350, 0.7986, 0.7597, 0.7186, 0.6732, 0.6213, 0.5722, 0.5322];
    private static readonly double[] RobinsonY = [0.0000, 0.0620, 0.1240, 0.1860, 0.2480, 0.3100, 0.3720, 0.4340, 0.4958, 0.5571, 0.6176, 0.6769, 0.7346, 0.7903, 0.8435, 0.8936, 0.9394, 0.9761, 1.0000];

    /// <summary>Map projections (x right, y up, unit sphere): Mercator ln tan(pi/4 + phi/2); Miller 1.25 ln tan(pi/4 + 0.4 phi); Robinson by its
    /// 5-degree table (x = 0.8487 X lambda, y = 1.3523 Y, linear between rows); Albers equal-area conic with the given standard parallels.</summary>
    private static (double X, double Y) Project(string projection, double lat, double lon, double parallel1, double parallel2, double centre)
    {
        double phi = Math.Max(-89.5, Math.Min(89.5, lat)) * Math.PI / 180, lambda = (lon - centre) * Math.PI / 180;
        switch (projection)
        {
            case "miller": return (lambda, 1.25 * Math.Log(Math.Tan(Math.PI / 4 + 0.4 * phi)));
            case "robinson":
            {
                double a = Math.Abs(lat) / 5;
                int row = Math.Min(17, (int)Math.Floor(a));
                double t = a - row, x = RobinsonX[row] + t * (RobinsonX[row + 1] - RobinsonX[row]), y = RobinsonY[row] + t * (RobinsonY[row + 1] - RobinsonY[row]);
                return (0.8487 * x * lambda, 1.3523 * y * Math.Sign(lat));
            }
            case "albers":
            {
                double p1 = parallel1 * Math.PI / 180, p2 = parallel2 * Math.PI / 180;
                double n = (Math.Sin(p1) + Math.Sin(p2)) / 2, c = Math.Cos(p1) * Math.Cos(p1) + 2 * n * Math.Sin(p1);
                if (Math.Abs(n) < 1e-9) return (lambda, phi);
                double rho = Math.Sqrt(Math.Max(0, c - 2 * n * Math.Sin(phi))) / n, theta = n * lambda;
                return (rho * Math.Sin(theta), -rho * Math.Cos(theta));
            }
            default: return (lambda, Math.Log(Math.Tan(Math.PI / 4 + phi / 2)));
        }
    }

    private void DrawRegionMap(StringBuilder svg, List<string>? texts, Series series, double left, double top, double right, double bottom)
    {
        List<GeoEntity?>? entities = Geography(series);
        if (entities is null) return;
        string projection = (string?)series.Layouts?.Element("geography")?.Attribute("projectionType") ?? "mercator";
        var points = entities.OfType<GeoEntity>().SelectMany(e => e.Rings).SelectMany(r => r).ToList();
        double minLat = points.Min(p => p.Lat), maxLat = points.Max(p => p.Lat), centre = (points.Min(p => p.Lon) + points.Max(p => p.Lon)) / 2;
        double p1 = minLat + (maxLat - minLat) / 6, p2 = minLat + (maxLat - minLat) * 5 / 6;
        var projected = entities.Select(e => e?.Rings.Select(r => r.Select(p => Project(projection, p.Lat, p.Lon, p1, p2, centre)).ToList()).ToList()).ToList();
        var all = projected.Where(e => e is not null).SelectMany(e => e!).SelectMany(r => r).ToList();
        double minX = all.Min(p => p.X), maxX = all.Max(p => p.X), minY = all.Min(p => p.Y), maxY = all.Max(p => p.Y);
        double legendWidth = _legendPosition is null ? 0 : 50;
        double scale = Math.Min((right - left - legendWidth) / Math.Max(1e-9, maxX - minX), (bottom - top) / Math.Max(1e-9, maxY - minY));
        double offsetX = left + (right - left - legendWidth - (maxX - minX) * scale) / 2, offsetY = top + (bottom - top - (maxY - minY) * scale) / 2;
        var values = series.Values.Where(v => v is not null).Select(v => v!.Value).ToList();
        double vMin = values.DefaultIfEmpty(0).Min(), vMax = values.DefaultIfEmpty(1).Max();
        XElement? colors = series.Element?.Element("valueColors");
        string low = DocxDrawingColor.Css(colors?.Element("minColor"), _theme) ?? LightAccent();
        string high = DocxDrawingColor.Css(colors?.Element("maxColor"), _theme) ?? Accent(0);
        string? middle = DocxDrawingColor.Css(colors?.Element("midColor"), _theme);
        List<string> names = Leaves(series);
        for (int i = 0; i < projected.Count; i++)
        {
            if (projected[i] is not { } rings || series.Values[i] is not { } value) continue;
            double t = vMax > vMin ? (value - vMin) / (vMax - vMin) : 1;
            string fill = middle is null ? Mix(low, high, t) : t < 0.5 ? Mix(low, middle, t * 2) : Mix(middle, high, (t - 0.5) * 2);
            var d = new StringBuilder();
            foreach (var ring in rings)
                for (int k = 0; k < ring.Count; k++)
                    d.Append(k == 0 ? 'M' : 'L').Append(N(offsetX + (ring[k].X - minX) * scale)).Append(' ').Append(N(offsetY + (maxY - ring[k].Y) * scale)).Append(k == ring.Count - 1 ? "Z" : "");
            svg.Append("<path d=\"").Append(d).Append("\" fill=\"").Append(fill).Append("\" fill-rule=\"evenodd\" stroke=\"#FFFFFF\" stroke-width=\"0.5\"/>");
            if (texts is not null && i < names.Count) texts.Add(names[i]);
        }
        if (_legendPosition is not null)
        {
            string id = "word-chart-scale-" + (low + high).GetHashCode().ToString("x", CultureInfo.InvariantCulture);
            svg.Append("<defs><linearGradient id=\"").Append(id).Append("\" x1=\"0\" y1=\"1\" x2=\"0\" y2=\"0\"><stop offset=\"0\" stop-color=\"").Append(low).Append("\"/>")
               .Append(middle is null ? "" : "<stop offset=\"0.5\" stop-color=\"" + middle + "\"/>").Append("<stop offset=\"1\" stop-color=\"").Append(high).Append("\"/></linearGradient></defs>")
               .Append("<rect x=\"").Append(N(right - 46)).Append("\" y=\"").Append(N(top + 10)).Append("\" width=\"8\" height=\"").Append(N(bottom - top - 20)).Append("\" fill=\"url(#").Append(id).Append(")\"/>");
            Label(svg, null, right - 34, top + 14, DocxChart.FormatNumber(vMax, series.FormatCode), 8, "start");
            Label(svg, null, right - 34, bottom - 10, DocxChart.FormatNumber(vMin, series.FormatCode), 8, "start");
        }
    }

    private string LightAccent()
        => DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", "accent1"), new XElement("lumMod", new XAttribute("val", "20000")), new XElement("lumOff", new XAttribute("val", "80000")))), _theme) ?? "#DAE3F3";

    private static string Mix(string from, string to, double t)
    {
        int Channel(string css, int shift) => (int.Parse(css.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture) >> shift) & 0xFF;
        t = Math.Max(0, Math.Min(1, t));
        int Blend(int shift) => (int)Math.Round(Channel(from, shift) + (Channel(to, shift) - Channel(from, shift)) * t);
        return "#" + ((Blend(16) << 16) | (Blend(8) << 8) | Blend(0)).ToString("X6", CultureInfo.InvariantCulture);
    }

    private static string N(double value) { double r = Math.Round(value, 2); return (r == 0 ? 0 : r).ToString("0.##", CultureInfo.InvariantCulture); }
}
