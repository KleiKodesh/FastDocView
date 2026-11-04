using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.DocxHtml;

/// <summary>
/// A DrawingML chart part (ECMA-376 Part 1 21.2, c:chartSpace) read from its cached values only (c:strCache / c:numCache: the
/// embedded workbook is never opened) and drawn as one SVG: bar / column (clustered, stacked, percent-stacked), line, area,
/// pie / doughnut / of-pie (the "of" pie drawn as a pie of every point), scatter, radar, bubble, stock (its series as lines)
/// and surface (as lines of its series); the 3-D kinds (bar3D, line3D, pie3D, area3D, surface3D) are drawn as their 2-D kinds.
/// Titles (rich text or, without text, the single series' name), axis titles, category and value axes with major gridlines,
/// legend (c:legendPos, c:legendEntry delete), data labels (showVal / showCatName / showSerName / showPercent). Colours: the
/// series' and points' own c:spPr, else LibreOffice's automatic chart style (oox/source/drawingml/chart/objectformatter.cxx:
/// c:style default 2 (chartspacemodel.cxx:27), the series pattern of spFilledSeries2dFills / spLinearSeriesLines :246-300
/// with patterns 1-4 :72-115 and the per-cycle shade / tint of getPhColor :766-812; varyColors = one colour per point).
/// Charts this does not read (chartex: cx:chart - waterfall, treemap, sunburst, histogram, box and whisker, funnel, region map)
/// are left to the drawing's mc:Fallback picture by the renderer.
/// Layout and automatic axis scaling are this library's own (LibreOffice's chart2 layout is not ported): the title on top,
/// the legend on its side, the plot area in the rest or at its c:manualLayout; value axes from 0 (or the data minimum when
/// it is negative) to a 1 / 2 / 5 x 10^n step above the data, or the axis' c:min / c:max / c:majorUnit.
/// </summary>
internal sealed class DocxChart
{
    private const double EmuPerPoint = 12700;

    private enum Kind { Bar, Line, Area, Pie, Doughnut, Scatter, Radar, Bubble }

    private sealed class Series
    {
        public int Index, Order;
        public string? Name;
        public List<string> Categories = new();
        public List<double?> Values = new(), XValues = new();
        public List<double?> BubbleSizes = new();
        public string? FormatCode;
        public XElement? ShapeProperties;
        public Dictionary<int, XElement> PointProperties = new();
        public XElement? Labels;
        public bool Smooth, NoMarker;
    }

    private sealed class Group
    {
        public Kind Kind;
        public bool Horizontal;
        public string Grouping = "clustered";
        public bool VaryColors;
        public double GapWidth = 150, Overlap, HoleSize = 50, FirstSliceAngle;
        public string? ScatterStyle;
        public List<Series> Series = new();
        public XElement? Labels;
    }

    private sealed class Axis
    {
        public string Id = "";
        public bool IsCategory, Deleted, Horizontal;
        public double? Min, Max, MajorUnit;
        public bool Gridlines;
        public string? Title;
        public string? FormatCode;
        public bool Reversed;
    }

    private readonly List<Group> _groups = new();
    private readonly Dictionary<string, Axis> _axes = new(StringComparer.Ordinal);
    private string? _title;
    private string? _legendPosition;
    private readonly HashSet<int> _deletedLegendEntries = new();
    private int _style = 2;
    private XElement? _chartSpaceProperties, _plotAreaProperties;
    private (double X, double Y, double W, double H)? _plotLayout;
    private readonly DocxTheme? _theme;

    private DocxChart(DocxTheme? theme) => _theme = theme;

    /// <summary>The chart of a chart part (namespace-free c:chartSpace); null when it has no c:chart.</summary>
    public static DocxChart? Read(XElement chartSpace, DocxTheme? theme)
    {
        XElement? chart = chartSpace.Element("chart");
        if (chart is null) return null;
        var model = new DocxChart(theme);
        if (Value(chartSpace.Element("style")) is { } style && int.TryParse(style, out int styleIndex)) model._style = styleIndex;
        else if (chartSpace.Element("AlternateContent")?.Element("Fallback")?.Element("style") is { } fallbackStyle && int.TryParse(Value(fallbackStyle), out int fallbackIndex)) model._style = fallbackIndex;
        model._chartSpaceProperties = chartSpace.Element("spPr");
        if (chart.Element("title") is { } title) model._title = TitleText(title);
        XElement? plotArea = chart.Element("plotArea");
        model._plotAreaProperties = plotArea?.Element("spPr");
        if (plotArea?.Element("layout")?.Element("manualLayout") is { } manual && Number(manual.Element("w")) is { } w && Number(manual.Element("h")) is { } h)
            model._plotLayout = (Number(manual.Element("x")) ?? 0, Number(manual.Element("y")) ?? 0, w, h);
        foreach (XElement element in plotArea?.Elements() ?? [])
        {
            string name = element.Name.LocalName;
            if (name is "catAx" or "dateAx" or "serAx" or "valAx")
            {
                var axis = new Axis
                {
                    Id = Value(element.Element("axId")) ?? "", IsCategory = name != "valAx", Deleted = Value(element.Element("delete")) is "1" or "true",
                    Min = Number(element.Element("scaling")?.Element("min")), Max = Number(element.Element("scaling")?.Element("max")),
                    MajorUnit = Number(element.Element("majorUnit")), Gridlines = element.Element("majorGridlines") is not null,
                    Title = element.Element("title") is { } axisTitle ? TitleText(axisTitle) : null,
                    FormatCode = (string?)element.Element("numFmt")?.Attribute("formatCode"),
                    Reversed = Value(element.Element("scaling")?.Element("orientation")) == "maxMin",
                };
                model._axes[axis.Id] = axis;
                continue;
            }
            Kind? kind = name switch
            {
                "barChart" or "bar3DChart" => Kind.Bar, "lineChart" or "line3DChart" or "stockChart" or "surfaceChart" or "surface3DChart" => Kind.Line,
                "areaChart" or "area3DChart" => Kind.Area, "pieChart" or "pie3DChart" or "ofPieChart" => Kind.Pie, "doughnutChart" => Kind.Doughnut,
                "scatterChart" => Kind.Scatter, "radarChart" => Kind.Radar, "bubbleChart" => Kind.Bubble, _ => null,
            };
            if (kind is null) continue;
            var group = new Group
            {
                Kind = kind.Value, Horizontal = Value(element.Element("barDir")) == "bar", Grouping = Value(element.Element("grouping")) ?? "clustered",
                VaryColors = Value(element.Element("varyColors")) is "1" or "true", GapWidth = Number(element.Element("gapWidth")) ?? 150,
                Overlap = Number(element.Element("overlap")) ?? (Value(element.Element("grouping")) is "stacked" or "percentStacked" ? 100 : 0),
                HoleSize = Number(element.Element("holeSize")) ?? 50, FirstSliceAngle = Number(element.Element("firstSliceAng")) ?? 0,
                ScatterStyle = Value(element.Element("scatterStyle")), Labels = element.Element("dLbls"),
            };
            if (name is "stockChart" or "surfaceChart" or "surface3DChart") group.Grouping = "standard";
            foreach (XElement series in element.Elements("ser")) group.Series.Add(ReadSeries(series));
            group.Series.Sort((a, b) => a.Order.CompareTo(b.Order));
            if (group.Series.Count > 0) model._groups.Add(group);
        }
        if (chart.Element("legend") is { } legend)
        {
            model._legendPosition = Value(legend.Element("legendPos")) ?? "r";
            foreach (XElement entry in legend.Elements("legendEntry"))
                if (Value(entry.Element("delete")) is "1" or "true" && int.TryParse(Value(entry.Element("idx")), out int deleted)) model._deletedLegendEntries.Add(deleted);
        }
        // c:title without text: the single series' name (as Office shows it), else no title
        if (chart.Element("title") is not null && model._title is null && model._groups.SelectMany(g => g.Series).ToList() is { Count: 1 } only) model._title = only[0].Name;
        if (Value(chart.Element("autoTitleDeleted")) is "1" or "true" && chart.Element("title") is null) model._title = null;
        return model;
    }

    private static Series ReadSeries(XElement element)
    {
        var series = new Series
        {
            Index = int.TryParse(Value(element.Element("idx")), out int index) ? index : 0,
            Order = int.TryParse(Value(element.Element("order")), out int order) ? order : 0,
            Name = SeriesName(element.Element("tx")), ShapeProperties = element.Element("spPr"), Labels = element.Element("dLbls"),
            Smooth = Value(element.Element("smooth")) is "1" or "true", NoMarker = Value(element.Element("marker")?.Element("symbol")) == "none",
        };
        series.Categories = Strings(element.Element("cat") ?? element.Element("xVal"));
        (series.Values, series.FormatCode) = Numbers(element.Element("val") ?? element.Element("yVal"));
        (series.XValues, _) = Numbers(element.Element("xVal"));
        (series.BubbleSizes, _) = Numbers(element.Element("bubbleSize"));
        foreach (XElement point in element.Elements("dPt"))
            if (int.TryParse(Value(point.Element("idx")), out int pointIndex) && point.Element("spPr") is { } properties) series.PointProperties[pointIndex] = properties;
        return series;
    }

    private static string? SeriesName(XElement? tx)
    {
        if (tx is null) return null;
        if (tx.Element("v") is { } literal) return literal.Value;
        XElement? cache = tx.Element("strRef")?.Element("strCache");
        return cache?.Elements("pt").FirstOrDefault()?.Element("v")?.Value;
    }

    /// <summary>Category texts from a str / num / multi-level cache (pt idx order, ptCount slots); numbers formatted by their format code.</summary>
    private static List<string> Strings(XElement? data)
    {
        var result = new List<string>();
        if (data is null) return result;
        XElement? cache = data.Element("strRef")?.Element("strCache") ?? data.Element("strLit") ?? data.Element("numRef")?.Element("numCache") ?? data.Element("numLit")
                          ?? data.Element("multiLvlStrRef")?.Element("multiLvlStrCache")?.Elements("lvl").FirstOrDefault();
        if (cache is null) return result;
        string? format = (string?)cache.Element("formatCode");
        int count = int.TryParse(Value(cache.Element("ptCount")), out int c) ? c : 0;
        var points = new SortedDictionary<int, string>();
        foreach (XElement point in cache.Elements("pt"))
            if (int.TryParse((string?)point.Attribute("idx"), out int idx))
            {
                string text = point.Element("v")?.Value ?? "";
                if (data.Element("numRef") is not null || data.Element("numLit") is not null)
                    text = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? FormatNumber(number, (string?)point.Attribute("formatCode") ?? format) : text;
                points[idx] = text;
            }
        int slots = Math.Max(count, points.Count == 0 ? 0 : points.Keys.Max() + 1);
        for (int i = 0; i < slots; i++) result.Add(points.TryGetValue(i, out string? text) ? text : "");
        return result;
    }

    private static (List<double?> Values, string? FormatCode) Numbers(XElement? data)
    {
        var result = new List<double?>();
        XElement? cache = data?.Element("numRef")?.Element("numCache") ?? data?.Element("numLit");
        if (cache is null) return (result, null);
        int count = int.TryParse(Value(cache.Element("ptCount")), out int c) ? c : 0;
        var points = new Dictionary<int, double>();
        foreach (XElement point in cache.Elements("pt"))
            if (int.TryParse((string?)point.Attribute("idx"), out int idx) && double.TryParse(point.Element("v")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
                points[idx] = value;
        int slots = Math.Max(count, points.Count == 0 ? 0 : points.Keys.Max() + 1);
        for (int i = 0; i < slots; i++) result.Add(points.TryGetValue(i, out double value) ? value : null);
        return (result, (string?)cache.Element("formatCode"));
    }

    private static string? TitleText(XElement title)
    {
        if (title.Element("tx")?.Element("rich") is { } rich) return string.Join("\n", rich.Elements("p").Select(p => string.Concat(p.Elements().Where(e => e.Name.LocalName is "r" or "fld").Select(r => (string?)r.Element("t")))));
        if (title.Element("tx")?.Element("strRef")?.Element("strCache") is { } cache) return cache.Elements("pt").FirstOrDefault()?.Element("v")?.Value;
        return null;
    }

    private static string? Value(XElement? element) => (string?)element?.Attribute("val");

    private static double? Number(XElement? element)
        => double.TryParse((string?)element?.Attribute("val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) ? value : null;

    /// <summary>
    /// A number by an Excel format code, the parts this needs for chart labels: the first section; "General" = shortest form;
    /// 0 / # digit placeholders with a decimal count; "," thousands; "%" (times 100); quoted and escaped literals kept.
    /// </summary>
    internal static string FormatNumber(double value, string? format)
    {
        if (string.IsNullOrEmpty(format) || format!.Equals("General", StringComparison.OrdinalIgnoreCase)) return value.ToString("0.##########", CultureInfo.InvariantCulture);
        string section = format.Split(';')[value < 0 && format.Contains(';') ? 1 : 0];
        bool sectionHasSign = value < 0 && format.Contains(';');
        var prefix = new StringBuilder();
        var suffix = new StringBuilder();
        int decimals = 0;
        bool percent = false, thousands = false, inNumber = false, afterNumber = false, inDecimals = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '"')
            {
                int close = section.IndexOf('"', i + 1);
                string literal = close < 0 ? section.Substring(i + 1) : section.Substring(i + 1, close - i - 1);
                (afterNumber || inNumber ? suffix : prefix).Append(literal);
                i = close < 0 ? section.Length : close;
                if (inNumber) { inNumber = false; afterNumber = true; }
                continue;
            }
            if (c == '\\' && i + 1 < section.Length) { (afterNumber || inNumber ? suffix : prefix).Append(section[++i]); continue; }
            if (c == '[') { int close = section.IndexOf(']', i); i = close < 0 ? section.Length : close; continue; }
            if (c is '_' or '*') { i++; continue; }
            if (c is '0' or '#' or '?') { if (afterNumber) continue; inNumber = true; if (inDecimals) decimals++; continue; }
            if (c == '.' && inNumber) { inDecimals = true; continue; }
            if (c == ',' && inNumber) { thousands = true; continue; }
            if (c == '%') { percent = true; (inNumber || afterNumber ? suffix : prefix).Append('%'); continue; }
            if (inNumber) { inNumber = false; afterNumber = true; }
            (afterNumber ? suffix : prefix).Append(c);
        }
        double shown = percent ? value * 100 : value;
        if (sectionHasSign) shown = Math.Abs(shown);
        string number = shown.ToString((thousands ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : ""), CultureInfo.InvariantCulture);
        return prefix + number + suffix;
    }

    // ── texts (search) ─────────────────────────────────────────────────────────────────────────

    /// <summary>Every text the drawing shows from the document (title, axis titles, legend entries, category labels, data labels), in drawing order.</summary>
    public IEnumerable<string> Texts()
    {
        if (_title is not null) yield return _title;
        foreach (Axis axis in _axes.Values) if (!axis.Deleted && axis.Title is not null) yield return axis.Title;
        foreach (string label in CategoryLabels()) yield return label;
        foreach ((string text, _, _) in LegendEntries()) yield return text;
        foreach (Group group in _groups)
            foreach (Series series in group.Series)
                for (int i = 0; i < series.Values.Count; i++)
                    if (LabelText(group, series, i) is { } label) yield return label;
    }

    private List<string> CategoryLabels()
    {
        Group? first = _groups.FirstOrDefault(g => g.Kind is not (Kind.Pie or Kind.Doughnut or Kind.Scatter or Kind.Bubble));
        if (first is null || CategoryAxis() is { Deleted: true }) return new();
        Series? withCategories = first.Series.FirstOrDefault(s => s.Categories.Count > 0);
        if (withCategories is not null) return withCategories.Categories;
        int count = first.Series.Max(s => s.Values.Count);
        return Enumerable.Range(1, count).Select(i => i.ToString(CultureInfo.InvariantCulture)).ToList();
    }

    private Axis? CategoryAxis() => _axes.Values.FirstOrDefault(a => a.IsCategory);
    private Axis? ValueAxis() => _axes.Values.FirstOrDefault(a => !a.IsCategory);

    /// <summary>Legend entries: (text, fill colour, is a line); varyColors pies list their points.</summary>
    private List<(string Text, string Color, bool Line)> LegendEntries()
    {
        var entries = new List<(string, string, bool)>();
        if (_legendPosition is null) return entries;
        int totalSeries = _groups.Sum(g => g.Series.Count);
        int position = 0;
        foreach (Group group in _groups)
        {
            if (group.VaryColors && group.Series.Count == 1 && group.Kind is Kind.Pie or Kind.Doughnut or Kind.Bar or Kind.Line or Kind.Area)
            {
                Series only = group.Series[0];
                List<string> names = only.Categories.Count > 0 ? only.Categories : Enumerable.Range(1, only.Values.Count).Select(i => i.ToString(CultureInfo.InvariantCulture)).ToList();
                for (int i = 0; i < names.Count; i++)
                    if (!_deletedLegendEntries.Contains(i)) entries.Add((names[i], FillOf(group, only, i, names.Count).Css, false));
                position += 1;
                continue;
            }
            foreach (Series series in group.Series)
            {
                if (!_deletedLegendEntries.Contains(position))
                    entries.Add((series.Name ?? "Series" + (series.Index + 1), group.Kind is Kind.Line or Kind.Scatter or Kind.Radar ? LineOf(series, totalSeries).Css : FillOf(group, series, -1, totalSeries).Css,
                                 group.Kind is Kind.Line or Kind.Scatter or Kind.Radar));
                position++;
            }
        }
        return entries;
    }

    private string? LabelText(Group group, Series series, int point)
    {
        XElement? labels = series.Labels ?? group.Labels;
        if (labels is null || Value(labels.Element("delete")) is "1" or "true") return null;
        bool On(string name) => Value(labels.Element(name)) is "1" or "true";
        var parts = new List<string>();
        if (On("showSerName") && series.Name is not null) parts.Add(series.Name);
        if (On("showCatName") && point < series.Categories.Count) parts.Add(series.Categories[point]);
        if (On("showVal") && point < series.Values.Count && series.Values[point] is { } value)
            parts.Add(FormatNumber(value, (string?)labels.Element("numFmt")?.Attribute("formatCode") ?? series.FormatCode));
        if (On("showPercent") && point < series.Values.Count && series.Values[point] is { } part)
        {
            double total = series.Values.Sum(v => Math.Abs(v ?? 0));
            if (total > 0) parts.Add(FormatNumber(Math.Abs(part) / total, "0%"));
        }
        return parts.Count == 0 ? null : string.Join(", ", parts);
    }

    // ── colours ─────────────────────────────────────────────────────────────────────────

    private static readonly (string Scheme, string? Mod, int Value)[] Pattern1 =
        [("dk1", "tint", 88500), ("dk1", "tint", 55000), ("dk1", "tint", 78000), ("dk1", "tint", 92500), ("dk1", "tint", 70000), ("dk1", "tint", 30000)];
    private static readonly (string Scheme, string? Mod, int Value)[] Pattern2 =
        [("accent1", null, 0), ("accent2", null, 0), ("accent3", null, 0), ("accent4", null, 0), ("accent5", null, 0), ("accent6", null, 0)];
    private static readonly (string Scheme, string? Mod, int Value)[] Pattern4 =
        [("dk1", "tint", 5000), ("dk1", "tint", 55000), ("dk1", "tint", 78000), ("dk1", "tint", 15000), ("dk1", "tint", 70000), ("dk1", "tint", 30000)];

    /// <summary>The series colour pattern of the chart style (spFilledSeries2dFills / spLinearSeriesLines: 1, 9, 17, 25, 33 = pattern 1;
    /// 2, 10, 18, 26, 34, 42 = pattern 2; 41 = pattern 4; 3-8, 11-16, 19-24, 27-32, 35-40, 43-48 = one accent, faded per series).</summary>
    private (string Scheme, string? Mod, int Value)[] StylePattern()
    {
        int style = _style is >= 1 and <= 48 ? _style : 2;
        if (style == 41) return Pattern4;
        int inRow = (style - 1) % 8;
        if (inRow == 0) return Pattern1;
        if (inRow == 1) return Pattern2;
        return [("accent" + (inRow - 1).ToString(CultureInfo.InvariantCulture), null, 0)];
    }

    /// <summary>getPhColor: the pattern colour of the index, shaded / tinted per cycle (objectformatter.cxx:766-812, Color::addChartTintTransformation color.cxx:488-495).</summary>
    private string AutomaticColor(int index, int count)
    {
        var pattern = StylePattern();
        var (scheme, mod, modValue) = pattern[((index % pattern.Length) + pattern.Length) % pattern.Length];
        int cycle = index / pattern.Length, maxCycle = Math.Max(0, count - 1) / pattern.Length;
        double shadeTint = (cycle + 1.0) / (maxCycle + 2) * 1.4 - 0.7;
        var color = new XElement("schemeClr", new XAttribute("val", scheme));
        if (mod is not null) color.Add(new XElement(mod, new XAttribute("val", modValue.ToString(CultureInfo.InvariantCulture))));
        int tint = (int)Math.Max(-100000, Math.Min(100000, shadeTint * 100000 + 0.5));
        if (tint < 0) color.Add(new XElement("shade", new XAttribute("val", (tint + 100000).ToString(CultureInfo.InvariantCulture))));
        else if (tint > 0) color.Add(new XElement("tint", new XAttribute("val", (100000 - tint).ToString(CultureInfo.InvariantCulture))));
        return DocxDrawingColor.Css(new XElement("c", color), _theme) ?? "#4472C4";
    }

    private DocxResolvedColor FillOf(Group group, Series series, int point, int count)
    {
        XElement? properties = point >= 0 && series.PointProperties.TryGetValue(point, out XElement? own) ? own : series.ShapeProperties;
        if (properties?.Element("solidFill") is { } solid && DocxDrawingColor.Resolve(solid, _theme) is { } color) return color;
        if (properties?.Element("gradFill")?.Element("gsLst")?.Element("gs") is { } stop && DocxDrawingColor.Resolve(stop, _theme) is { } first) return first;
        int index = group.VaryColors && point >= 0 ? point : series.Index;
        return new DocxResolvedColor(AutomaticColor(index, group.VaryColors && point >= 0 ? count : count), 1);
    }

    private (string Css, double WidthPt, bool None) LineOf(Series series, int count)
    {
        XElement? line = series.ShapeProperties?.Element("ln");
        if (line?.Element("noFill") is not null) return ("none", 0, true);
        double width = line?.Attribute("w") is { } w && double.TryParse(w.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double emu) ? emu / EmuPerPoint : 2.25;
        string css = line?.Element("solidFill") is { } solid ? DocxDrawingColor.Css(solid, _theme) ?? AutomaticColor(series.Index, count) : AutomaticColor(series.Index, count);
        return (css, width, false);
    }

    private string? PropertiesFill(XElement? properties, string? fallback)
    {
        if (properties?.Element("noFill") is not null) return null;
        if (properties?.Element("solidFill") is { } solid) return DocxDrawingColor.Css(solid, _theme);
        return fallback;
    }

    // ── drawing ─────────────────────────────────────────────────────────────────────────

    public string Svg(double width, double height)
    {
        width = Math.Max(width, 1); height = Math.Max(height, 1);
        var svg = new StringBuilder();
        svg.Append("<svg class=\"word-chart\" aria-hidden=\"true\" viewBox=\"0 0 ").Append(N(width)).Append(' ').Append(N(height)).Append("\" width=\"").Append(N(width))
           .Append("pt\" height=\"").Append(N(height)).Append("pt\" style=\"position:absolute;left:0;top:0;width:100%;height:100%;overflow:hidden\" font-family=\"Calibri, sans-serif\">");
        // chart area: spChartSpaceFill bg1 for styles 1-32, lt1 33-40, dk1 41-48
        string chartFill = PropertiesFill(_chartSpaceProperties, DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", _style >= 41 ? "dk1" : _style >= 33 ? "lt1" : "bg1"))), _theme) ?? "#FFFFFF") ?? "none";
        string? chartLine = _chartSpaceProperties?.Element("ln") is { } chartLn && chartLn.Element("noFill") is null && chartLn.Element("solidFill") is { } chartLineFill ? DocxDrawingColor.Css(chartLineFill, _theme) : null;
        svg.Append("<rect width=\"").Append(N(width)).Append("\" height=\"").Append(N(height)).Append("\" fill=\"").Append(chartFill).Append('"')
           .Append(chartLine is null ? " stroke=\"none\"" : " stroke=\"" + chartLine + "\" stroke-width=\"0.75\"").Append("/>");
        string textColor = DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", "tx1"), new XElement("lumMod", new XAttribute("val", "65000")), new XElement("lumOff", new XAttribute("val", "35000")))), _theme) ?? "#595959";
        double top = 6, bottom = height - 6, left = 6, right = width - 6;
        if (_title is not null)
        {
            string[] lines = _title.Split('\n');
            for (int i = 0; i < lines.Length; i++) Text(svg, width / 2, top + 14 + i * 16.8, lines[i], 14, textColor, "middle");
            top += 8 + lines.Length * 16.8;
        }
        List<(string Text, string Color, bool Line)> legend = LegendEntries();
        if (legend.Count > 0)
        {
            const double entryHeight = 13, size = 9;
            if (_legendPosition is "r" or "l" or "tr")
            {
                double legendWidth = Math.Min(width / 3, legend.Max(e => e.Text.Length) * size * 0.55 + 18);
                double y = (top + bottom) / 2 - legend.Count * entryHeight / 2;
                double x = _legendPosition == "l" ? left : right - legendWidth;
                foreach ((string text, string color, bool line) in legend) { LegendKey(svg, x, y, color, line); Text(svg, x + 14, y + 8, text, size, textColor, "start"); y += entryHeight; }
                if (_legendPosition == "l") left += legendWidth + 6; else right -= legendWidth + 6;
            }
            else
            {
                double total = legend.Sum(e => e.Text.Length * size * 0.55 + 22);
                double x = Math.Max(left, width / 2 - total / 2), y = _legendPosition == "t" ? top : bottom - entryHeight;
                foreach ((string text, string color, bool line) in legend) { LegendKey(svg, x, y, color, line); Text(svg, x + 14, y + 8, text, size, textColor, "start"); x += text.Length * size * 0.55 + 22; }
                if (_legendPosition == "t") top += entryHeight + 6; else bottom -= entryHeight + 6;
            }
        }
        if (_plotLayout is { } manual) (left, top, right, bottom) = (manual.X * width, manual.Y * height, (manual.X + manual.W) * width, (manual.Y + manual.H) * height);
        if (_groups.Count == 0) { svg.Append("</svg>"); return svg.ToString(); }
        Group main = _groups[0];
        if (main.Kind is Kind.Pie or Kind.Doughnut) DrawPie(svg, main, left, top, right, bottom, textColor);
        else if (main.Kind == Kind.Radar) DrawRadar(svg, main, left, top, right, bottom, textColor);
        else DrawAxes(svg, left, top, right, bottom, textColor);
        svg.Append("</svg>");
        return svg.ToString();
    }

    private static void LegendKey(StringBuilder svg, double x, double y, string color, bool line)
    {
        if (line) svg.Append("<path d=\"M").Append(N(x)).Append(' ').Append(N(y + 5)).Append('h').Append(N(10)).Append("\" stroke=\"").Append(color).Append("\" stroke-width=\"2\"/>");
        else svg.Append("<rect x=\"").Append(N(x + 2)).Append("\" y=\"").Append(N(y + 1)).Append("\" width=\"7\" height=\"7\" fill=\"").Append(color).Append("\"/>");
    }

    private static void Text(StringBuilder svg, double x, double y, string text, double size, string color, string anchor, double rotate = 0)
    {
        svg.Append("<text x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" font-size=\"").Append(N(size)).Append("\" fill=\"").Append(color)
           .Append("\" text-anchor=\"").Append(anchor).Append('"');
        if (rotate != 0) svg.Append(" transform=\"rotate(").Append(N(rotate)).Append(' ').Append(N(x)).Append(' ').Append(N(y)).Append(")\"");
        svg.Append('>').Append(WordHtmlText.Encode(text)).Append("</text>");
    }

    /// <summary>An axis scale: the axis' own min / max / major unit, else 0 (or below the data when it is negative) to a 1 / 2 / 5 x 10^n step above the data, about 5 steps.</summary>
    private static (double Min, double Max, double Step) Scale(double dataMin, double dataMax, Axis? axis, bool percent)
    {
        if (percent) return (axis?.Min ?? 0, axis?.Max ?? 1, axis?.MajorUnit ?? 0.1);
        double min = axis?.Min ?? Math.Min(0, dataMin), max = axis?.Max ?? Math.Max(0, dataMax);
        if (max <= min) max = min + 1;
        double step = axis?.MajorUnit ?? NiceStep((max - min) / 5);
        if (axis?.Min is null) min = Math.Floor(min / step) * step;
        if (axis?.Max is null) max = Math.Ceiling(max / step) * step;
        if (max <= min) max = min + step;
        return (min, max, step);
    }

    private static double NiceStep(double raw)
    {
        if (raw <= 0) return 1;
        double magnitude = Math.Pow(10, Math.Floor(Math.Log10(raw))), fraction = raw / magnitude;
        return (fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10) * magnitude;
    }

    private void DrawAxes(StringBuilder svg, double left, double top, double right, double bottom, string textColor)
    {
        Axis? valueAxis = ValueAxis(), categoryAxis = CategoryAxis();
        bool horizontal = _groups.Any(g => g.Kind == Kind.Bar && g.Horizontal);
        bool xy = _groups[0].Kind is Kind.Scatter or Kind.Bubble;
        bool percent = _groups.Any(g => g.Grouping == "percentStacked");
        List<string> categories = CategoryLabels();
        int categoryCount = Math.Max(1, _groups.Max(g => g.Series.Max(s => s.Values.Count)));
        // value range over every group (stacked groups by their sums)
        double dataMin = 0, dataMax = 0;
        foreach (Group group in _groups)
        {
            bool stacked = group.Grouping is "stacked";
            for (int i = 0; i < categoryCount; i++)
            {
                double positive = 0, negative = 0;
                foreach (Series series in group.Series)
                {
                    double value = i < series.Values.Count ? series.Values[i] ?? 0 : 0;
                    if (stacked) { if (value >= 0) positive += value; else negative += value; }
                    else { dataMax = Math.Max(dataMax, value); dataMin = Math.Min(dataMin, value); }
                }
                if (stacked) { dataMax = Math.Max(dataMax, positive); dataMin = Math.Min(dataMin, negative); }
            }
        }
        (double vMin, double vMax, double vStep) = Scale(dataMin, dataMax, valueAxis, percent);
        // x range for scatter / bubble
        double xMin = 0, xMax = categoryCount, xStep = 1;
        if (xy)
        {
            var xs = _groups.SelectMany(g => g.Series).SelectMany(s => s.XValues.Count > 0 ? s.XValues : Enumerable.Range(1, s.Values.Count).Select(i => (double?)i)).Where(v => v is not null).Select(v => v!.Value).ToList();
            Axis? xAxis = _axes.Values.Where(a => !a.IsCategory).Skip(1).FirstOrDefault() ?? categoryAxis;
            (xMin, xMax, xStep) = Scale(xs.Count > 0 ? xs.Min() : 0, xs.Count > 0 ? xs.Max() : 1, xAxis, false);
            if (xAxis is not null && !xAxis.IsCategory) categoryAxis = xAxis;
        }
        // room for tick labels and axis titles
        double labelSize = 9;
        string valueFormat = valueAxis?.FormatCode ?? (percent ? "0%" : _groups[0].Series[0].FormatCode ?? "General");
        var valueLabels = new List<(double Value, string Text)>();
        if (valueAxis is not { Deleted: true })
            for (double v = vMin; v <= vMax + vStep / 1000; v += vStep) valueLabels.Add((v, FormatNumber(Math.Round(v, 10), valueFormat)));
        double valueLabelWidth = valueLabels.Count == 0 ? 0 : valueLabels.Max(l => l.Text.Length) * labelSize * 0.55 + 4;
        if (valueAxis?.Title is { } valueTitle) { if (horizontal) bottom -= 14; else left += 14; Text(svg, horizontal ? (left + right) / 2 : left - 6, horizontal ? bottom + 12 : (top + bottom) / 2, valueTitle, 10, textColor, "middle", horizontal ? 0 : -90); }
        if (categoryAxis?.Title is { } categoryTitle && !categoryAxis.Deleted) { if (horizontal) left += 14; else bottom -= 14; Text(svg, horizontal ? left - 6 : (left + right) / 2, horizontal ? (top + bottom) / 2 : bottom + 12, categoryTitle, 10, textColor, "middle", horizontal ? -90 : 0); }
        if (horizontal) { bottom -= labelSize + 4; left += categoryAxis is { Deleted: true } ? 0 : Math.Min((right - left) / 3, categories.DefaultIfEmpty("").Max(c => c.Length) * labelSize * 0.55 + 6); }
        else { left += valueLabelWidth; bottom -= categoryAxis is { Deleted: true } ? 0 : labelSize + 4; }
        double plotWidth = Math.Max(1, right - left), plotHeight = Math.Max(1, bottom - top);
        string? plotFill = PropertiesFill(_plotAreaProperties, null);
        if (plotFill is not null) svg.Append("<rect x=\"").Append(N(left)).Append("\" y=\"").Append(N(top)).Append("\" width=\"").Append(N(plotWidth)).Append("\" height=\"").Append(N(plotHeight)).Append("\" fill=\"").Append(plotFill).Append("\"/>");
        string gridColor = DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", "tx1"), new XElement("lumMod", new XAttribute("val", "15000")), new XElement("lumOff", new XAttribute("val", "85000")))), _theme) ?? "#D9D9D9";
        double ValueToAxis(double value) => (value - vMin) / (vMax - vMin);
        // gridlines and value labels
        foreach ((double value, string text) in valueLabels)
        {
            double t = ValueToAxis(value);
            if (horizontal)
            {
                double x = left + t * plotWidth;
                if (valueAxis?.Gridlines != false) svg.Append("<path d=\"M").Append(N(x)).Append(' ').Append(N(top)).Append('V').Append(N(bottom)).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
                Text(svg, x, bottom + labelSize + 2, text, labelSize, textColor, "middle");
            }
            else
            {
                double y = bottom - t * plotHeight;
                if (valueAxis?.Gridlines != false) svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(y)).Append('H').Append(N(right)).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
                Text(svg, left - 4, y + labelSize / 3, text, labelSize, textColor, "end");
            }
        }
        // category labels
        if (!xy && categoryAxis is not { Deleted: true })
            for (int i = 0; i < categoryCount; i++)
            {
                string label = i < categories.Count ? categories[i] : (i + 1).ToString(CultureInfo.InvariantCulture);
                double centre = (i + 0.5) / categoryCount;
                if (horizontal) Text(svg, left - 4, bottom - centre * plotHeight + labelSize / 3, label, labelSize, textColor, "end");
                else Text(svg, left + centre * plotWidth, bottom + labelSize + 2, label, labelSize, textColor, "middle");
            }
        if (xy)
            for (double x = xMin; x <= xMax + xStep / 1000; x += xStep)
                Text(svg, left + (x - xMin) / (xMax - xMin) * plotWidth, bottom + labelSize + 2, FormatNumber(Math.Round(x, 10), "General"), labelSize, textColor, "middle");
        int totalSeries = _groups.Sum(g => g.Series.Count);
        foreach (Group group in _groups)
        {
            switch (group.Kind)
            {
                case Kind.Bar: DrawBars(svg, group, categoryCount, left, top, plotWidth, plotHeight, ValueToAxis, totalSeries, textColor); break;
                case Kind.Line or Kind.Area: DrawLines(svg, group, categoryCount, left, top, plotWidth, plotHeight, ValueToAxis, totalSeries, textColor, vMin, vMax); break;
                case Kind.Scatter or Kind.Bubble: DrawScatter(svg, group, left, top, plotWidth, plotHeight, ValueToAxis, x => (x - xMin) / (xMax - xMin), totalSeries, textColor); break;
            }
        }
        // axis lines
        string axisColor = gridColor;
        svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(bottom)).Append('H').Append(N(right)).Append("\" stroke=\"").Append(axisColor).Append("\" stroke-width=\"0.75\"/>");
    }

    private void DrawBars(StringBuilder svg, Group group, int categories, double left, double top, double width, double height, Func<double, double> axis, int totalSeries, string textColor)
    {
        bool stacked = group.Grouping is "stacked" or "percentStacked", percent = group.Grouping == "percentStacked";
        int count = group.Series.Count;
        double slot = 1.0 / categories, gap = group.GapWidth / 100, overlap = group.Overlap / 100;
        // bar width b: count bars overlapping by overlap * b, plus a gap of gap * b, fill the category slot
        double bars = stacked ? 1 : count - (count - 1) * overlap;
        double barWidth = slot / (bars + gap);
        for (int i = 0; i < categories; i++)
        {
            double positive = 0, negative = 0, total = group.Series.Sum(s => Math.Abs(i < s.Values.Count ? s.Values[i] ?? 0 : 0));
            for (int s = 0; s < count; s++)
            {
                Series series = group.Series[s];
                if (i >= series.Values.Count || series.Values[i] is not { } raw) continue;
                double value = percent ? (total == 0 ? 0 : raw / total) : raw;
                double from, to;
                if (stacked) { if (value >= 0) { from = positive; positive += value; to = positive; } else { from = negative; negative += value; to = negative; } }
                else { from = 0; to = value; }
                double offset = i * slot + gap * barWidth / 2 + (stacked ? 0 : s * barWidth * (1 - overlap));
                double a = axis(Math.Min(from, to)), b = axis(Math.Max(from, to));
                DocxResolvedColor fill = FillOf(group, series, group.VaryColors && count == 1 ? i : -1, group.VaryColors && count == 1 ? categories : totalSeries);
                double x, y, w, h;
                if (group.Horizontal) { x = left + a * width; w = (b - a) * width; y = top + height - (offset + barWidth) * height; h = barWidth * height; }
                else { x = left + offset * width; w = barWidth * width; y = top + height - b * height; h = (b - a) * height; }
                svg.Append("<rect x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" width=\"").Append(N(Math.Max(0, w))).Append("\" height=\"").Append(N(Math.Max(0, h)))
                   .Append("\" fill=\"").Append(fill.Css).Append('"').Append(fill.Opacity < 1 ? " fill-opacity=\"" + N(fill.Opacity) + "\"" : "").Append("/>");
                if (LabelText(group, series, i) is { } label)
                    Text(svg, group.Horizontal ? x + w + 2 : x + w / 2, group.Horizontal ? y + h / 2 + 3 : y - 2, label, 9, textColor, group.Horizontal ? "start" : "middle");
            }
        }
    }

    private void DrawLines(StringBuilder svg, Group group, int categories, double left, double top, double width, double height, Func<double, double> axis, int totalSeries, string textColor, double vMin, double vMax)
    {
        bool stacked = group.Grouping is "stacked" or "percentStacked", percent = group.Grouping == "percentStacked";
        var baseline = new double[categories];
        for (int s = 0; s < group.Series.Count; s++)
        {
            Series series = group.Series[s];
            var points = new List<(double X, double Y, int Index)>();
            var lower = new List<(double X, double Y)>();
            for (int i = 0; i < categories && i < series.Values.Count; i++)
            {
                if (series.Values[i] is not { } raw) continue;
                double total = group.Series.Sum(x => Math.Abs(i < x.Values.Count ? x.Values[i] ?? 0 : 0));
                double value = percent ? (total == 0 ? 0 : raw / total) : raw;
                double from = stacked ? baseline[i] : Math.Max(vMin, Math.Min(vMax, 0));
                double to = stacked ? baseline[i] + value : value;
                if (stacked) baseline[i] = to;
                double x = left + (i + 0.5) / categories * width;
                points.Add((x, top + height - axis(to) * height, i));
                lower.Add((x, top + height - axis(from) * height));
            }
            if (points.Count == 0) continue;
            if (group.Kind == Kind.Area)
            {
                DocxResolvedColor fill = FillOf(group, series, -1, totalSeries);
                var d = new StringBuilder("M").Append(N(points[0].X)).Append(' ').Append(N(points[0].Y));
                foreach (var p in points.Skip(1)) d.Append('L').Append(N(p.X)).Append(' ').Append(N(p.Y));
                for (int k = lower.Count - 1; k >= 0; k--) d.Append('L').Append(N(lower[k].X)).Append(' ').Append(N(lower[k].Y));
                svg.Append("<path d=\"").Append(d).Append("Z\" fill=\"").Append(fill.Css).Append('"').Append(fill.Opacity < 1 ? " fill-opacity=\"" + N(fill.Opacity) + "\"" : "").Append("/>");
            }
            else
            {
                (string css, double lineWidth, bool none) = LineOf(series, totalSeries);
                if (!none)
                {
                    var d = new StringBuilder("M").Append(N(points[0].X)).Append(' ').Append(N(points[0].Y));
                    foreach (var p in points.Skip(1)) d.Append('L').Append(N(p.X)).Append(' ').Append(N(p.Y));
                    svg.Append("<path d=\"").Append(d).Append("\" fill=\"none\" stroke=\"").Append(css).Append("\" stroke-width=\"").Append(N(lineWidth)).Append("\" stroke-linejoin=\"round\"/>");
                }
                if (!series.NoMarker)
                    foreach (var p in points) svg.Append("<circle cx=\"").Append(N(p.X)).Append("\" cy=\"").Append(N(p.Y)).Append("\" r=\"2.5\" fill=\"").Append(css == "none" ? AutomaticColor(series.Index, totalSeries) : css).Append("\"/>");
            }
            foreach (var p in points)
                if (LabelText(group, series, p.Index) is { } label) Text(svg, p.X, p.Y - 4, label, 9, textColor, "middle");
        }
    }

    private void DrawScatter(StringBuilder svg, Group group, double left, double top, double width, double height, Func<double, double> axisY, Func<double, double> axisX, int totalSeries, string textColor)
    {
        bool lines = group.Kind == Kind.Scatter && group.ScatterStyle is not ("marker" or null) && group.ScatterStyle != "none";
        double maxBubble = group.Series.SelectMany(s => s.BubbleSizes).Where(v => v is not null).Select(v => Math.Abs(v!.Value)).DefaultIfEmpty(1).Max();
        foreach (Series series in group.Series)
        {
            (string css, double lineWidth, bool none) = LineOf(series, totalSeries);
            var points = new List<(double X, double Y, int Index)>();
            for (int i = 0; i < series.Values.Count; i++)
            {
                if (series.Values[i] is not { } y) continue;
                double x = i < series.XValues.Count && series.XValues[i] is { } xv ? xv : i + 1;
                points.Add((left + axisX(x) * width, top + height - axisY(y) * height, i));
            }
            if (group.Kind == Kind.Bubble)
            {
                DocxResolvedColor fill = FillOf(group, series, -1, totalSeries);
                foreach (var p in points)
                {
                    double size = p.Index < series.BubbleSizes.Count && series.BubbleSizes[p.Index] is { } b ? Math.Abs(b) : 1;
                    double radius = Math.Sqrt(size / maxBubble) * Math.Min(width, height) / 8;
                    svg.Append("<circle cx=\"").Append(N(p.X)).Append("\" cy=\"").Append(N(p.Y)).Append("\" r=\"").Append(N(radius)).Append("\" fill=\"").Append(fill.Css).Append("\" fill-opacity=\"0.75\"/>");
                }
            }
            else
            {
                if (lines && !none && points.Count > 1)
                {
                    var d = new StringBuilder("M").Append(N(points[0].X)).Append(' ').Append(N(points[0].Y));
                    foreach (var p in points.Skip(1)) d.Append('L').Append(N(p.X)).Append(' ').Append(N(p.Y));
                    svg.Append("<path d=\"").Append(d).Append("\" fill=\"none\" stroke=\"").Append(css).Append("\" stroke-width=\"").Append(N(lineWidth)).Append("\"/>");
                }
                if (!series.NoMarker)
                    foreach (var p in points) svg.Append("<circle cx=\"").Append(N(p.X)).Append("\" cy=\"").Append(N(p.Y)).Append("\" r=\"2.5\" fill=\"").Append(css == "none" ? AutomaticColor(series.Index, totalSeries) : css).Append("\"/>");
            }
            foreach (var p in points)
                if (LabelText(group, series, p.Index) is { } label) Text(svg, p.X, p.Y - 4, label, 9, textColor, "middle");
        }
    }

    private void DrawPie(StringBuilder svg, Group group, double left, double top, double right, double bottom, string textColor)
    {
        Series series = group.Series[0];
        double cx = (left + right) / 2, cy = (top + bottom) / 2, radius = Math.Max(1, Math.Min(right - left, bottom - top) / 2 - 4);
        double inner = group.Kind == Kind.Doughnut ? radius * group.HoleSize / 100 : 0;
        double total = series.Values.Sum(v => Math.Abs(v ?? 0));
        if (total <= 0) return;
        // slices clockwise from 12 o'clock turned by firstSliceAng
        double angle = group.FirstSliceAngle - 90;
        int count = series.Values.Count;
        for (int i = 0; i < count; i++)
        {
            double value = Math.Abs(series.Values[i] ?? 0);
            if (value <= 0) continue;
            double sweep = value / total * 360, start = angle * Math.PI / 180, end = (angle + sweep) * Math.PI / 180;
            DocxResolvedColor fill = FillOf(group, series, group.VaryColors ? i : -1, count);
            var d = new StringBuilder();
            bool large = sweep > 180;
            if (sweep >= 359.999)
                d.Append("M").Append(N(cx + radius)).Append(' ').Append(N(cy)).Append("A").Append(N(radius)).Append(' ').Append(N(radius)).Append(" 0 1 1 ").Append(N(cx - radius)).Append(' ').Append(N(cy))
                 .Append("A").Append(N(radius)).Append(' ').Append(N(radius)).Append(" 0 1 1 ").Append(N(cx + radius)).Append(' ').Append(N(cy)).Append('Z');
            else
            {
                d.Append('M').Append(N(cx + inner * Math.Cos(start))).Append(' ').Append(N(cy + inner * Math.Sin(start)))
                 .Append('L').Append(N(cx + radius * Math.Cos(start))).Append(' ').Append(N(cy + radius * Math.Sin(start)))
                 .Append('A').Append(N(radius)).Append(' ').Append(N(radius)).Append(" 0 ").Append(large ? '1' : '0').Append(" 1 ").Append(N(cx + radius * Math.Cos(end))).Append(' ').Append(N(cy + radius * Math.Sin(end)));
                if (inner > 0)
                    d.Append('L').Append(N(cx + inner * Math.Cos(end))).Append(' ').Append(N(cy + inner * Math.Sin(end)))
                     .Append('A').Append(N(inner)).Append(' ').Append(N(inner)).Append(" 0 ").Append(large ? '1' : '0').Append(" 0 ").Append(N(cx + inner * Math.Cos(start))).Append(' ').Append(N(cy + inner * Math.Sin(start)));
                else d.Append('L').Append(N(cx)).Append(' ').Append(N(cy));
                d.Append('Z');
            }
            string? sliceLine = PropertiesFill(series.ShapeProperties?.Element("ln") is { } ln ? new XElement("x", ln.Elements()) : null, "#FFFFFF");
            svg.Append("<path d=\"").Append(d).Append("\" fill=\"").Append(fill.Css).Append('"').Append(fill.Opacity < 1 ? " fill-opacity=\"" + N(fill.Opacity) + "\"" : "")
               .Append(sliceLine is null ? " stroke=\"none\"" : " stroke=\"" + sliceLine + "\" stroke-width=\"0.75\"").Append("/>");
            if (LabelText(group, series, i) is { } label)
            {
                double middle = (angle + sweep / 2) * Math.PI / 180, at = inner > 0 ? (inner + radius) / 2 : radius * 0.65;
                Text(svg, cx + at * Math.Cos(middle), cy + at * Math.Sin(middle) + 3, label, 9, textColor, "middle");
            }
            angle += sweep;
        }
    }

    private void DrawRadar(StringBuilder svg, Group group, double left, double top, double right, double bottom, string textColor)
    {
        double cx = (left + right) / 2, cy = (top + bottom) / 2, radius = Math.Max(1, Math.Min(right - left, bottom - top) / 2 - 14);
        int count = Math.Max(1, group.Series.Max(s => s.Values.Count));
        double dataMax = group.Series.SelectMany(s => s.Values).Where(v => v is not null).Select(v => v!.Value).DefaultIfEmpty(1).Max();
        (double min, double max, double step) = Scale(0, dataMax, ValueAxis(), false);
        string gridColor = "#D9D9D9";
        for (double v = min + step; v <= max + step / 1000; v += step)
        {
            var ring = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                double a = -Math.PI / 2 + i * 2 * Math.PI / count, r = (v - min) / (max - min) * radius;
                ring.Append(i == 0 ? 'M' : 'L').Append(N(cx + r * Math.Cos(a))).Append(' ').Append(N(cy + r * Math.Sin(a)));
            }
            svg.Append("<path d=\"").Append(ring).Append("Z\" fill=\"none\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
        }
        List<string> categories = CategoryLabels();
        for (int i = 0; i < count; i++)
        {
            double a = -Math.PI / 2 + i * 2 * Math.PI / count;
            svg.Append("<path d=\"M").Append(N(cx)).Append(' ').Append(N(cy)).Append('L').Append(N(cx + radius * Math.Cos(a))).Append(' ').Append(N(cy + radius * Math.Sin(a))).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
            if (i < categories.Count) Text(svg, cx + (radius + 8) * Math.Cos(a), cy + (radius + 8) * Math.Sin(a) + 3, categories[i], 9, textColor, "middle");
        }
        int totalSeries = group.Series.Count;
        foreach (Series series in group.Series)
        {
            var d = new StringBuilder();
            for (int i = 0; i < series.Values.Count; i++)
            {
                double a = -Math.PI / 2 + i * 2 * Math.PI / count, r = ((series.Values[i] ?? 0) - min) / (max - min) * radius;
                d.Append(i == 0 ? 'M' : 'L').Append(N(cx + r * Math.Cos(a))).Append(' ').Append(N(cy + r * Math.Sin(a)));
            }
            (string css, double width, bool none) = LineOf(series, totalSeries);
            if (!none) svg.Append("<path d=\"").Append(d).Append("Z\" fill=\"none\" stroke=\"").Append(css).Append("\" stroke-width=\"").Append(N(width)).Append("\"/>");
        }
    }

    private static string N(double value) { double r = Math.Round(value, 2); return (r == 0 ? 0 : r).ToString("0.##", CultureInfo.InvariantCulture); }
}
