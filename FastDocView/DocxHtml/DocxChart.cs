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
internal sealed partial class DocxChart
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
        public List<XElement> ErrorBars = new(), Trendlines = new();
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
        public List<string> AxisIds = new();
        public XElement? HiLowLines, UpDownBars;
        public bool IsStock;
    }

    private sealed class Axis
    {
        public string Id = "";
        public bool IsCategory, Deleted, Horizontal;
        /// <summary>No tick labels: a deleted axis, or c:tickLblPos none (21.2.2.207; LibreOffice oox axisconverter.cxx: none = no labels, the axis line stays).</summary>
        public bool NoLabels;
        /// <summary>The tick labels' own colour (c:txPr a:defRPr a:solidFill); null = the chart's text colour.</summary>
        public string? LabelColor;
        public double? Min, Max, MajorUnit;
        public bool Gridlines;
        public string? Title;
        public double TitleSize = 10, LabelSize = 9;
        public bool MidCategory, LineBreak = true, TryStaggeringFirst, NoTickMarks;
        public string? FormatCode;
        public bool Reversed;
        public double? LogBase;
        public string? Position;
    }

    private readonly List<Group> _groups = new();
    private readonly Dictionary<string, Axis> _axes = new(StringComparer.Ordinal);
    private string? _title;
    private string? _titleColor;
    private string? _legendColor;
    private string? _legendPosition;
    private readonly HashSet<int> _deletedLegendEntries = new();
    private int _style = 2;
    private XElement? _chartSpaceProperties, _plotAreaProperties;
    private (double X, double Y, double W, double H)? _plotLayout;
    private bool _plotLayoutInner;
    private List<double> _titleSizes = new();
    private bool _libreOfficeLayout, _separateTitleLines;
    private int _chartFont;
    private double _legendSize = 9, _dataLabelSize = 9;
    private (double X, double Y)? _legendLayout;
    private bool _is3D;
    private XElement? _floorProperties, _backWallProperties;
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
        // text sizes (a:defRPr / a:rPr sz in hundredths of a point): the element's own, else the chart's c:txPr; Office's defaults otherwise
        double? chartSize = FontSize(chartSpace.Element("txPr"));
        // the chart's Latin font (a:latin typeface), for text widths
        string? typeface = chartSpace.Descendants().FirstOrDefault(e => e.Name.LocalName == "latin")?.Attribute("typeface")?.Value;
        model._chartFont = ChartFontIndex(typeface);
        // a chart read from OpenDocument is laid out by LibreOffice's rules (OdtChart marks it)
        model._libreOfficeLayout = (string?)chartSpace.Attribute("layoutRules") == "libreoffice";
        model._separateTitleLines = (string?)chart.Element("title")?.Attribute("separateLines") == "1";
        if (chart.Element("title") is { } title)
        {
            model._title = TitleText(title);
            double fallback = FontSize(title.Element("txPr")) ?? 14;
            model._titleSizes = title.Element("tx")?.Element("rich")?.Elements("p").Select(p => FontSize(p) ?? fallback).ToList() ?? new List<double> { fallback };
            // the title's own text colour (a:rPr / a:defRPr a:solidFill, 20.1.2.3.32 / 20.1.2.3.33), else the chart's text colour
            if (title.Descendants().FirstOrDefault(e => e.Name.LocalName is "rPr" or "defRPr" && e.Element(e.Name.Namespace + "solidFill") is not null) is { } coloured)
                model._titleColor = DocxDrawingColor.Css(coloured.Element(coloured.Name.Namespace + "solidFill")!, theme);
        }
        if (chart.Element("legend") is { } legendText) model._legendSize = FontSize(legendText.Element("txPr")) ?? chartSize ?? 9;
        model._dataLabelSize = FontSize(chart.Element("plotArea")?.Descendants("dLbls").Select(d => d.Element("txPr")).FirstOrDefault(t => t is not null)) ?? chartSize ?? 9;
        XElement? plotArea = chart.Element("plotArea");
        model._plotAreaProperties = plotArea?.Element("spPr");
        if (plotArea?.Element("layout")?.Element("manualLayout") is { } manual && Number(manual.Element("w")) is { } w && Number(manual.Element("h")) is { } h)
        {
            model._plotLayout = (Number(manual.Element("x")) ?? 0, Number(manual.Element("y")) ?? 0, w, h);
            // c:layoutTarget (21.2.2.101): inner = the plot rectangle without tick labels; outer (the default) = with them
            model._plotLayoutInner = Value(manual.Element("layoutTarget")) == "inner";
        }
        // 3-D: c:view3D, c:floor, c:backWall (21.2.2.228, 21.2.2.69, 21.2.2.11)
        model._is3D = plotArea?.Elements().Any(e => e.Name.LocalName.EndsWith("3DChart", StringComparison.Ordinal)) == true;   // c:view3D alone is written for 2-D charts too
        model._floorProperties = chart.Element("floor")?.Element("spPr");
        if (model._is3D) model._view3D = ReadView3D(chart.Element("view3D"), plotArea?.Element("pie3DChart") is not null);
        model._backWallProperties = chart.Element("backWall")?.Element("spPr");
        foreach (XElement element in plotArea?.Elements() ?? [])
        {
            string name = element.Name.LocalName;
            if (name is "catAx" or "dateAx" or "serAx" or "valAx")
            {
                var axis = new Axis
                {
                    Id = Value(element.Element("axId")) ?? "", IsCategory = name != "valAx", Deleted = Value(element.Element("delete")) is "1" or "true",
                    NoLabels = Value(element.Element("delete")) is "1" or "true" || Value(element.Element("tickLblPos")) == "none",
                    LabelColor = TextColorOf(element.Element("txPr"), theme),
                    Min = Number(element.Element("scaling")?.Element("min")), Max = Number(element.Element("scaling")?.Element("max")),
                    MajorUnit = Number(element.Element("majorUnit")), Gridlines = element.Element("majorGridlines") is not null,
                    Title = element.Element("title") is { } axisTitle ? TitleText(axisTitle) : null,
                    MidCategory = Value(element.Element("crossBetween")) == "midCat",
                    NoTickMarks = Value(element.Element("majorTickMark")) is "none" or "in",
                    LineBreak = (string?)element.Attribute("lineBreak") != "0", TryStaggeringFirst = (string?)element.Attribute("tryStaggeringFirst") == "1",
                    TitleSize = FontSize(element.Element("title")) ?? (chartSize is { } c10 ? c10 : 10), LabelSize = FontSize(element.Element("txPr")) ?? chartSize ?? 9,
                    FormatCode = (string?)element.Element("numFmt")?.Attribute("formatCode"),
                    Reversed = Value(element.Element("scaling")?.Element("orientation")) == "maxMin",
                    LogBase = Number(element.Element("scaling")?.Element("logBase")), Position = Value(element.Element("axPos")),
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
            group.AxisIds = element.Elements("axId").Select(a => Value(a) ?? "").ToList();
            // stock: c:hiLowLines and c:upDownBars (21.2.2.80, 21.2.2.218)
            group.IsStock = name == "stockChart";
            group.HiLowLines = element.Element("hiLowLines");
            group.UpDownBars = element.Element("upDownBars");
            foreach (XElement series in element.Elements("ser")) group.Series.Add(ReadSeries(series));
            group.Series.Sort((a, b) => a.Order.CompareTo(b.Order));
            if (group.Series.Count > 0) model._groups.Add(group);
        }
        if (chart.Element("legend") is { } legend)
        {
            model._legendPosition = Value(legend.Element("legendPos")) ?? "r";
            model._legendColor = TextColorOf(legend.Element("txPr"), theme);
            // only an edge-mode position (c:xMode / c:yMode edge, 21.2.2.233 / 21.2.2.235); the default factor mode is an offset from the automatic place
            if (legend.Element("layout")?.Element("manualLayout") is { } legendManual && Value(legendManual.Element("xMode")) == "edge" && Value(legendManual.Element("yMode")) == "edge"
                && Number(legendManual.Element("x")) is { } lx && Number(legendManual.Element("y")) is { } ly)
                model._legendLayout = (lx, ly);
            foreach (XElement entry in legend.Elements("legendEntry"))
                if (Value(entry.Element("delete")) is "1" or "true" && int.TryParse(Value(entry.Element("idx")), out int deleted)) model._deletedLegendEntries.Add(deleted);
        }
        // c:title without text: the single series' name (as Office shows it), else no title
        if (chart.Element("title") is not null && model._title is null && model._groups.SelectMany(g => g.Series).ToList() is { Count: 1 } only) model._title = only[0].Name;
        if (Value(chart.Element("autoTitleDeleted")) is "1" or "true" && chart.Element("title") is null) model._title = null;
        return model;
    }

    /// <summary>The colour of a text-properties element's first run properties with a fill (a:defRPr / a:rPr a:solidFill); null when none.</summary>
    private static string? TextColorOf(XElement? textProperties, DocxTheme? theme)
        => textProperties?.Descendants().FirstOrDefault(e => e.Name.LocalName is "defRPr" or "rPr" && e.Element(e.Name.Namespace + "solidFill") is not null) is { } run
            ? DocxDrawingColor.Css(run.Element(run.Name.Namespace + "solidFill")!, theme) : null;

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
        series.ErrorBars = element.Elements("errBars").ToList();
        series.Trendlines = element.Elements("trendline").ToList();
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
        if (first is null || CategoryAxis() is { NoLabels: true }) return new();
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
            // stacked columns list their series top down, as they are stacked (both Word's and LibreOffice's PDFs of the fixture: the last series first)
            int groupStart = entries.Count;
            foreach (Series series in group.Series)
            {
                if (!_deletedLegendEntries.Contains(position))
                    entries.Add((series.Name ?? "Series" + (series.Index + 1), group.Kind is Kind.Line or Kind.Scatter or Kind.Radar ? LineOf(series, totalSeries).Css : FillOf(group, series, -1, totalSeries).Css,
                                 group.Kind is Kind.Line or Kind.Scatter or Kind.Radar));
                position++;
            }
            if (group.Kind == Kind.Bar && !group.Horizontal && group.Grouping is "stacked" or "percentStacked") entries.Reverse(groupStart, entries.Count - groupStart);
            // horizontal bars list their series in the order they stand from the top, the last series first (Word's PDF of
            // word-charts-binary-v3: West, South, North for a clustered bar chart of North, South, West)
            else if (group.Kind == Kind.Bar && group.Horizontal && group.Grouping == "clustered") entries.Reverse(groupStart, entries.Count - groupStart);   // stacked bars keep their order (Word's PDF of odt-charts-v5)
        }
        // Word lists a trend line after the series (its c:name, else a name in Office's own interface language: Word's PDFs of
        // word-charts-binary-v3 and odt-charts-v5, made by a Hebrew Office, both show "<Hebrew 'linear'> (North)" whatever the
        // chart's c:lang) - so the name is not in the file and follows the reader's interface language here
        if (!_libreOfficeLayout)
            foreach (Group group in _groups)
                foreach (Series series in group.Series)
                    foreach (XElement trend in series.Trendlines)
                    {
                        string name = (string?)trend.Element("name") is { Length: > 0 } own ? own : TrendLineName(Value(trend.Element("trendlineType")) ?? "linear", trend, series.Name ?? "Series" + (series.Index + 1));
                        string color = trend.Element("spPr")?.Element("ln")?.Element("solidFill") is { } solid ? DocxDrawingColor.Css(solid, _theme) ?? LineOf(series, totalSeries).Css : LineOf(series, totalSeries).Css;
                        entries.Add((name, color, true));
                    }
        return entries;
    }

    /// <summary>
    /// Office's automatic trend line name, "type (series)", in the reader's interface language: Hebrew as a Hebrew Office writes it
    /// (the linear name read from Word's PDFs; the others from memory, unchecked), else English (Excel's names, from memory).
    /// </summary>
    internal static string TrendLineName(string type, XElement trend, string seriesName)
    {
        bool hebrew = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "he";
        string period = ((int?)Number(trend.Element("period")) ?? 2).ToString(CultureInfo.InvariantCulture);
        string kind = type switch
        {
            "exp" => hebrew ? "\u05DE\u05E2\u05E8\u05D9\u05DB\u05D9" : "Expon.",
            "log" => hebrew ? "\u05DC\u05D5\u05D2\u05E8\u05D9\u05EA\u05DE\u05D9" : "Log.",
            "poly" => hebrew ? "\u05E4\u05D5\u05DC\u05D9\u05E0\u05D5\u05DE\u05D9" : "Poly.",
            "power" => hebrew ? "\u05D7\u05D6\u05E7\u05D4" : "Power",
            "movingAvg" => hebrew ? period + " \u05DE\u05DE\u05D5\u05E6\u05E2 \u05E0\u05E2" : period + " per. Mov. Avg.",
            _ => hebrew ? "\u05DC\u05D9\u05E0\u05D9\u05D0\u05E8\u05D9" : "Linear",
        };
        return kind + " (" + seriesName + ")";
    }

    private string? LabelText(Group group, Series series, int point)
    {
        XElement? labels = series.Labels ?? group.Labels;
        // c:dLbl (21.2.2.47): one point's own label settings replace the series' (Word: a label on a single point)
        if (labels?.Elements("dLbl").FirstOrDefault(d => Value(d.Element("idx")) == point.ToString(CultureInfo.InvariantCulture)) is { } own) labels = own;
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
        double top = 6, bottom = height - (_libreOfficeLayout ? 6 : 10.8), left = 6, right = width - 6;   // Word: 10.8 pt to the bottom edge (measured)
        if (_title is not null)
        {
            // each line at its own size, 1.2 x the size apart, the first baseline one size below the top
            string[] lines = _title.Split('\n');
            double y = top;
            for (int i = 0; i < lines.Length; i++)
            {
                double size = TitleSize(i);
                // LibreOffice's title and subtitle are two objects: the next baseline is the two sizes apart (measured 13 + 11 = 24 pt in its PDF)
                if (_separateTitleLines && i > 0) y += size - TitleSize(i - 1) * 0.2;
                Text(svg, width / 2, y + size, lines[i], size, _titleColor ?? textColor, "middle");
                y += size * 1.2;
            }
            top = y + 8 + (_libreOfficeLayout ? 0 : TitleSize(0) * 0.35);   // Word leaves about a third of the title size more below it (measured on its PDFs: the top tick 18-19 pt below the title baseline)
        }
        List<(string Text, string Color, bool Line)> legend = LegendEntries();
        if (legend.Count > 0)
        {
            double size = _legendSize, entryHeight = size * 13 / 9;
            if (_legendLayout is { } at)
            {
                // c:legend c:manualLayout (21.2.2.95): the legend's top left corner where it was saved; the plot area keeps its room
                double x = at.X * width, y = at.Y * height;
                bool row = _legendPosition is "t" or "b";
                foreach ((string text, string color, bool line) in legend)
                {
                    LegendKey(svg, x, y, color, line); Text(svg, x + 14, y + 8, text, size, _legendColor ?? textColor, "start");
                    if (row) x += TextWidth(text, size) + 22; else y += entryHeight;
                }
                double legendWidth = row ? 0 : legend.Max(e => TextWidth(e.Text, size)) + (_libreOfficeLayout ? 18 : 19.4);
                if (_plotLayout is null && !row) { if (at.X < 0.5) left = Math.Max(left, at.X * width + legendWidth + 6); else right = Math.Min(right, at.X * width - 6); }
            }
            else if (_legendPosition is "r" or "l" or "tr")
            {
                double legendWidth = Math.Min(width / 3, legend.Max(e => TextWidth(e.Text, size)) + (_libreOfficeLayout ? 18 : 19.4));
                double y = (top + bottom) / 2 - legend.Count * entryHeight / 2;
                double x = _legendPosition == "l" ? left : right - legendWidth;
                foreach ((string text, string color, bool line) in legend) { LegendKey(svg, x, y, color, line); Text(svg, x + 14 + (_libreOfficeLayout ? 0 : 1.4), y + 8 + (_libreOfficeLayout ? 0 : 1.3), text, size, _legendColor ?? textColor, "start"); y += entryHeight; }
                if (_legendPosition == "l") left += legendWidth + 6; else right -= legendWidth + 6;
            }
            else
            {
                // a legend row. Word (measured on its PDF of word-charts-binary-v3, three bottom legends): the name 10.7 pt after
                // the entry's start, the next entry 14.4 pt after the name, the row 1.8 pt lower; LibreOffice: 14 pt and 22 pt
                double nameAt = _libreOfficeLayout ? 14 : 10.7, advance = _libreOfficeLayout ? 22 : 14.4;
                double total = legend.Sum(e => TextWidth(e.Text, size) + advance) - (_libreOfficeLayout ? 0 : advance - nameAt);
                double rowShift = _libreOfficeLayout || _legendPosition == "t" ? 0 : 1.8;
                double x = Math.Max(left, width / 2 - total / 2), y = (_legendPosition == "t" ? top : bottom - entryHeight) + rowShift;
                foreach ((string text, string color, bool line) in legend)
                {
                    LegendKey(svg, _libreOfficeLayout ? x : x + nameAt - 10.7, y, color, line);
                    Text(svg, x + nameAt, y + 8 + (_libreOfficeLayout ? 0 : 1.3), text, size, _legendColor ?? textColor, "start");
                    x += TextWidth(text, size) + advance;
                }
                if (_legendPosition == "t") top += entryHeight + 6; else bottom -= entryHeight + 6;
            }
        }
        if (_plotLayout is { } manual) (left, top, right, bottom) = (manual.X * width, manual.Y * height, (manual.X + manual.W) * width, (manual.Y + manual.H) * height);
        if (_groups.Count == 0) { svg.Append("</svg>"); return svg.ToString(); }
        Group main = _groups[0];
        if (main.Kind is Kind.Pie && _is3D) Draw3DPie(svg, main, left, top, right, bottom, textColor);
        else if (main.Kind is Kind.Pie or Kind.Doughnut) DrawPie(svg, main, left, top, right, bottom, textColor);
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

    /// <summary>
    /// A chart text. One whose first strong character is right-to-left (Hebrew, Arabic) is laid out right to left (direction rtl), as
    /// Word draws it (its PDF puts a Hebrew name's trailing number at the left end); the anchor keeps the same physical edge, since
    /// text-anchor start / end follow the text's direction (SVG 1.1 10.9.1).
    /// </summary>
    private static void Text(StringBuilder svg, double x, double y, string text, double size, string color, string anchor, double rotate = 0)
    {
        if (text.Length == 0) return;                // a tick whose label is turned off
        bool rightToLeft = IsRightToLeftText(text);
        if (rightToLeft) anchor = anchor switch { "start" => "end", "end" => "start", _ => anchor };
        svg.Append("<text x=\"").Append(N(x)).Append("\" y=\"").Append(N(y)).Append("\" font-size=\"").Append(N(size)).Append("\" fill=\"").Append(color)
           .Append("\" text-anchor=\"").Append(anchor).Append('"');
        if (rightToLeft) svg.Append(" direction=\"rtl\" unicode-bidi=\"embed\"");
        if (rotate != 0) svg.Append(" transform=\"rotate(").Append(N(rotate)).Append(' ').Append(N(x)).Append(' ').Append(N(y)).Append(")\"");
        svg.Append('>').Append(WordHtmlText.Encode(text)).Append("</text>");
    }

    /// <summary>True when the first strong character is right-to-left (Hebrew U+0590-08FF and presentation forms U+FB1D-FDFF, U+FE70-FEFF).</summary>
    private static bool IsRightToLeftText(string text)
    {
        foreach (char c in text)
        {
            if (c is >= '\u0590' and <= '\u08FF' or >= '\uFB1D' and <= '\uFDFF' or >= '\uFE70' and <= '\uFEFF') return true;
            if (char.IsLetter(c)) return false;
        }
        return false;
    }

    /// <summary>An axis scale: the axis' own min / max / major unit, else 0 (or below the data when it is negative) to a 1 / 2 / 5 x 10^n step above the data, about 5 steps.</summary>
    private static (double Min, double Max, double Step) Scale(double dataMin, double dataMax, Axis? axis, bool percent)
    {
        if (percent) return (axis?.Min ?? 0, axis?.Max ?? 1, axis?.MajorUnit ?? 0.1);
        double min = axis?.Min ?? Math.Min(0, dataMin), max = axis?.Max ?? Math.Max(0, dataMax);
        if (max <= min) max = min + 1;
        double step = axis?.MajorUnit is { } unit && unit > (max - min) / 1000 ? unit : NiceStep((max - min) / 5);   // a c:majorUnit of 0 or less, or tiny, would never end
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

    /// <summary>The value range of a set of groups over the categories (stacked groups by their sums), and the smallest positive value (a logarithmic axis starts below it).</summary>
    private static (double Min, double Max, double PositiveMin) DataRange(IEnumerable<Group> groups, int categoryCount, bool includeErrorBars = false)
    {
        double dataMin = 0, dataMax = 0, positiveMin = double.MaxValue;
        foreach (Group group in groups)
        {
            bool stacked = group.Grouping is "stacked";
            for (int i = 0; i < categoryCount; i++)
            {
                double positive = 0, negative = 0;
                foreach (Series series in group.Series)
                {
                    double value = i < series.Values.Count ? series.Values[i] ?? 0 : 0;
                    if (value > 0) positiveMin = Math.Min(positiveMin, value);
                    if (stacked) { if (value >= 0) positive += value; else negative += value; }
                    else { dataMax = Math.Max(dataMax, value); dataMin = Math.Min(dataMin, value); }
                }
                if (stacked) { dataMax = Math.Max(dataMax, positive); dataMin = Math.Min(dataMin, negative); }
            }
            // the ends of the y error bars count too (Word's PDFs: odt-charts-v5 "Error Bars" / "Range Errors", word-charts-binary-v3)
            if (!stacked && includeErrorBars)
                foreach (Series series in group.Series)
                    foreach (double end in ErrorBarEnds(series, "y")) { dataMax = Math.Max(dataMax, end); dataMin = Math.Min(dataMin, end); }
        }
        return (dataMin, dataMax, positiveMin == double.MaxValue ? 1 : positiveMin);
    }

    /// <summary>
    /// A value axis' scale and its tick values. Logarithmic (c:logBase; ODF chart:logarithmic): from the base's power at or below the
    /// smallest positive value to the one at or above the largest, a tick per power - LibreOffice's automatic logarithmic scale works on
    /// the logarithms the same way (chart2/source/view/axes/ScaleAutomatism.cxx:294-330, calculateExplicitIncrementAndScaleForLogarithmic).
    /// </summary>
    /// <summary>
    /// The automatic linear scale as LibreOffice or Word makes it. Both fit as many ticks as the axis has room for, one label height
    /// apart (LibreOffice: VCartesianAxis.cxx:1559-1612, axis length / text height, at most 10, ScaleAutomatism.cxx:43-49), with an interval
    /// of 1, 2 or 5 x 10^n at or above range / count (ScaleAutomatism.cxx:851-913). LibreOffice rounds the ends out to the interval and adds
    /// one more interval when the data reach past 20/21 of the range (:918-940); Word first adds 5 % of the range above the largest value
    /// (measured on Word's PDFs of the fixture: 40 -> 45, 3 -> 3.5, 7 -> 8, 10 -> 12, 8 -> 9) and rounds out.
    /// </summary>
    private (double Min, double Max, double Step)? AutoScale(double dataMin, double dataMax, Axis? axis, double axisLength, double labelSize)
    {
        if (axis?.MajorUnit is > 0 && axis.Min is not null && axis.Max is not null) return null;
        double min = Math.Min(0, dataMin), max = Math.Max(0, dataMax);
        if (!_libreOfficeLayout)
        {
            double span = max - min;
            if (max > 0) max += span * 0.05;
            if (min < 0) min -= span * 0.05;
        }
        min = axis?.Min ?? min; max = axis?.Max ?? max;
        if (max <= min) max = min + 1;
        int count = Math.Max(2, Math.Min(10, (int)(axisLength / Math.Max(1, labelSize * (_libreOfficeLayout ? 1.2 : 1.0)))));   // LibreOffice: the text height; Word: one size (measured: 45 / 5 over 107 pt at 10 pt)
        double step = axis?.MajorUnit is { } unit && unit > (max - min) / 1000 ? unit : NiceStep((max - min) / count);
        double low = axis?.Min ?? Math.Floor(min / step + 1e-9) * step, high = axis?.Max ?? Math.Ceiling(max / step - 1e-9) * step;
        if (_libreOfficeLayout && axis?.Max is null && high != 0 && (dataMax - low) / (high - low) > 20.0 / 21.0) high += step;
        if (_libreOfficeLayout && axis?.Min is null && low != 0 && (high - dataMin) / (high - low) > 20.0 / 21.0) low -= step;
        if (high <= low) high = low + step;
        return (low, high, step);
    }

    private (double Min, double Max, List<double> Ticks) AxisScale(double dataMin, double dataMax, double positiveMin, Axis? axis, bool percent, double axisLength, double labelSize)
    {
        if (axis?.LogBase is null && !percent && axisLength > 0 && AutoScale(dataMin, dataMax, axis, axisLength, labelSize) is { } auto)
        {
            var autoTicks = new List<double>();
            for (double v = auto.Min; v <= auto.Max + auto.Step / 1000 && v + auto.Step > v && autoTicks.Count <= MaxTicks; v += auto.Step) autoTicks.Add(v);
            return (auto.Min, auto.Max, autoTicks);
        }
        return AxisScale(dataMin, dataMax, positiveMin, axis, percent);
    }

    /// <summary>A guard against damaged scales (an infinite span, or a step below the value's precision) that would never end; no real axis has this many ticks.</summary>
    private const int MaxTicks = 10000;

    private static (double Min, double Max, List<double> Ticks) AxisScale(double dataMin, double dataMax, double positiveMin, Axis? axis, bool percent)
    {
        var ticks = new List<double>();
        if (axis?.LogBase is { } logBase && logBase >= 2 && logBase <= 1000 && !percent)   // c:logBase range 2..1000 (21.2.2.100)
        {
            double Log(double v) => Math.Log(v) / Math.Log(logBase);
            double low = axis.Min is > 0 ? axis.Min.Value : Math.Pow(logBase, Math.Floor(Log(positiveMin) + 1e-9));
            double high = axis.Max is > 0 ? axis.Max.Value : Math.Pow(logBase, Math.Ceiling(Log(Math.Max(dataMax, low)) - 1e-9));
            if (high <= low) high = low * logBase;
            for (double v = low; v <= high * (1 + 1e-9) && ticks.Count <= MaxTicks; v *= logBase) ticks.Add(v);
            return (low, high, ticks);
        }
        (double min, double max, double step) = Scale(dataMin, dataMax, axis, percent);
        for (double v = min; v <= max + step / 1000 && v + step > v && ticks.Count <= MaxTicks; v += step) ticks.Add(v);
        return (min, max, ticks);
    }

    /// <summary>The value axis of the groups not on the first value axis (c:axId of a valAx no other group uses; ODF chart:attached-axis secondary-y), or null.</summary>
    private Axis? SecondaryValueAxis(Axis? primary)
    {
        foreach (Group group in _groups)
            foreach (string id in group.AxisIds)
                if (_axes.TryGetValue(id, out Axis? axis) && !axis.IsCategory && axis != primary && !axis.Deleted) return axis;
        return null;
    }

    private void DrawAxes(StringBuilder svg, double left, double top, double right, double bottom, string textColor)
    {
        Axis? valueAxis = ValueAxis(), categoryAxis = CategoryAxis();
        bool horizontal = _groups.Any(g => g.Kind == Kind.Bar && g.Horizontal);
        bool xy = _groups[0].Kind is Kind.Scatter or Kind.Bubble;
        bool percent = _groups.Any(g => g.Grouping == "percentStacked");
        List<string> categories = CategoryLabels();
        int categoryCount = Math.Max(1, _groups.Max(g => g.Series.Max(s => s.Values.Count)));
        // secondary value axis: the groups attached to it get their own scale, labelled on the right
        Axis? secondaryAxis = xy || horizontal ? null : SecondaryValueAxis(valueAxis);
        List<Group> secondaryGroups = secondaryAxis is null ? new() : _groups.Where(g => g.AxisIds.Contains(secondaryAxis.Id)).ToList();
        List<Group> primaryGroups = _groups.Except(secondaryGroups).ToList();
        if (primaryGroups.Count == 0) { primaryGroups = secondaryGroups; secondaryGroups = new(); secondaryAxis = null; }
        percent = primaryGroups.Any(g => g.Grouping == "percentStacked");
        (double dataMin, double dataMax, double positiveMin) = DataRange(primaryGroups, categoryCount, includeErrorBars: !_libreOfficeLayout);
        double estimatedLength = Math.Max(1, horizontal ? right - left - 60 : bottom - top - 30);
        double axisLabelSize = valueAxis?.LabelSize ?? categoryAxis?.LabelSize ?? 9;
        (double vMin, double vMax, List<double> valueTicks) = AxisScale(dataMin, dataMax, positiveMin, valueAxis, percent, estimatedLength, axisLabelSize);
        // x range for scatter / bubble
        double xMin = 0, xMax = categoryCount, xStep = 1;
        if (xy)
        {
            var xs = _groups.SelectMany(g => g.Series).SelectMany(s => s.XValues.Count > 0 ? s.XValues : Enumerable.Range(1, s.Values.Count).Select(i => (double?)i)).Where(v => v is not null).Select(v => v!.Value).ToList();
            // Word scales the axis over the error bars' ends too (Word's PDF of word-charts-binary-v3: x 1..4 with a custom +0.8 on the
            // last point gives an x axis to 6, not 5)
            foreach (Series series in _groups.SelectMany(g => g.Series)) xs.AddRange(ErrorBarEnds(series, "x"));
            Axis? xAxis = _axes.Values.Where(a => !a.IsCategory).Skip(1).FirstOrDefault() ?? categoryAxis;
            (xMin, xMax, xStep) = Scale(xs.Count > 0 ? xs.Min() : 0, xs.Count > 0 ? xs.Max() : 1, xAxis, false);
            // x values that are text (1..n, Word): the axis is scaled as a value axis, with the headroom (v3: x up to 4.8 gives 0..6)
            if (!_libreOfficeLayout && xAxis is { LogBase: null } && _groups.SelectMany(g => g.Series).All(x => x.XValues.All(v => v is null)) && AutoScale(xs.Count > 0 ? xs.Min() : 0, xs.Count > 0 ? xs.Max() : 1, xAxis, Math.Max(1, right - left - 60), axisLabelSize) is { } autoX)
                (xMin, xMax, xStep) = (autoX.Min, autoX.Max, autoX.Step);
            if (xAxis is not null && !xAxis.IsCategory) categoryAxis = xAxis;
        }
        // room for tick labels and axis titles
        double labelSize = valueAxis?.LabelSize ?? categoryAxis?.LabelSize ?? 9;
        double valueTitleSize = valueAxis?.TitleSize ?? 10, categoryTitleSize = categoryAxis?.TitleSize ?? 10;
        string valueFormat = valueAxis?.FormatCode ?? (percent ? "0%" : _groups[0].Series[0].FormatCode ?? "General");
        var valueLabels = new List<(double Value, string Text)>();
        // the ticks carry the gridlines even when the labels are off (c:tickLblPos none: empty texts); a deleted axis has neither
        if (valueAxis is not { Deleted: true })
            foreach (double v in valueTicks) valueLabels.Add((v, valueAxis is { NoLabels: true } ? "" : FormatNumber(Math.Round(v, 10), valueFormat)));
        double valueLabelWidth = valueLabels.All(l => l.Text.Length == 0) ? 0 : valueLabels.Max(l => TextWidth(l.Text, labelSize)) + (_libreOfficeLayout ? 4 : 8.2);   // Word: 6.2 pt from the label to the plot (measured)
        double sMin = 0, sMax = 1;
        var secondaryLabels = new List<(double Value, string Text)>();
        if (secondaryAxis is not null && !secondaryAxis.NoLabels)
        {
            (double sDataMin, double sDataMax, double sPositiveMin) = DataRange(secondaryGroups, categoryCount);
            (sMin, sMax, List<double> secondaryTicks) = AxisScale(sDataMin, sDataMax, sPositiveMin, secondaryAxis, secondaryGroups.Any(g => g.Grouping == "percentStacked"), estimatedLength, axisLabelSize);
            string secondaryFormat = secondaryAxis.FormatCode ?? secondaryGroups[0].Series[0].FormatCode ?? "General";
            foreach (double v in secondaryTicks) secondaryLabels.Add((v, FormatNumber(Math.Round(v, 10), secondaryFormat)));
        }
        double secondaryLabelWidth = secondaryLabels.Count == 0 ? 0 : secondaryLabels.Max(l => TextWidth(l.Text, labelSize)) + 4;
        if (_plotLayout is not null && _plotLayoutInner)
        {
            // c:layoutTarget inner: the stored rectangle is the plot itself; the labels and titles go outside it
            double catLabelWidth = Math.Min((right - left) / 2, CategoryLabelWidth(categories, labelSize));
            left -= (valueAxis?.Title is not null && !horizontal ? valueTitleSize + 4 : 0) + (categoryAxis?.Title is not null && !categoryAxis.Deleted && horizontal ? categoryTitleSize + 4 : 0)
                    + (horizontal ? (categoryAxis is { NoLabels: true } ? 0 : catLabelWidth) : valueLabelWidth);
            bottom += (valueAxis?.Title is not null && horizontal ? valueTitleSize + 4 : 0) + (categoryAxis?.Title is not null && !categoryAxis.Deleted && !horizontal ? categoryTitleSize + 4 : 0)
                      + (horizontal ? (valueAxis is { NoLabels: true } ? 0 : labelSize + 4) : categoryAxis is { NoLabels: true } ? 0 : labelSize + 4);
            right += secondaryLabelWidth + (secondaryAxis?.Title is not null ? 14 : 0);
        }
        if (valueAxis?.Title is { } valueTitle) { double room = valueTitleSize + 4; if (horizontal) bottom -= room; else left += room; Text(svg, horizontal ? (left + right) / 2 : left - 6, horizontal ? bottom + valueTitleSize + 2 : (top + bottom) / 2, valueTitle, valueTitleSize, textColor, "middle", horizontal ? 0 : -90); }
        if (categoryAxis?.Title is { } categoryTitle && !categoryAxis.Deleted) { double room = categoryTitleSize + 4; if (horizontal) left += room; else bottom -= room; Text(svg, horizontal ? left - 6 : (left + right) / 2, horizontal ? (top + bottom) / 2 : bottom + categoryTitleSize + 2 + (_libreOfficeLayout ? categoryTitleSize * 0.5 : 0), categoryTitle, categoryTitleSize, textColor, "middle", horizontal ? -90 : 0); }
        if (secondaryAxis?.Title is { } secondaryTitle) { right -= 14; Text(svg, right + 6, (top + bottom) / 2, secondaryTitle, 10, textColor, "middle", 90); }
        right -= secondaryLabelWidth;
        // Word: a tick label centred on the plot's right edge (the last value of a horizontal bar chart's value axis, the last category
        // of an axis with the points on the ticks) stays 10.8 pt plus half its width from the chart's edge (measured on Word's PDF of
        // word-charts-binary-v3: the bar chart's "6" and the area chart's "Q4", both 4.8 pt further in than the plain 6 pt margin)
        if (!_libreOfficeLayout && _plotLayout is null && secondaryAxis is null && _legendPosition is not ("r" or "tr") && !xy && valueAxis is not { Reversed: true } && categoryAxis is not { Reversed: true })
        {
            string? edgeLabel = horizontal ? (valueLabels.Count > 0 ? valueLabels[valueLabels.Count - 1].Text : null)
                              : MidCategory() && categoryAxis is not { NoLabels: true } && categories.Count > 0 ? categories[categories.Count - 1] : null;
            if (edgeLabel is { Length: > 0 }) right = Math.Max(left + 1, right - (10.8 - 6 + TextWidth(edgeLabel, labelSize) / 2));   // the right edge so far is the chart edge less 6 pt
        }
        if (horizontal) { bottom -= valueAxis is { NoLabels: true } ? 0 : labelSize + 4; left += categoryAxis is { NoLabels: true } ? 0 : _plotLayout is not null && _plotLayoutInner ? Math.Min((right - left) / 2, CategoryLabelWidth(categories, labelSize)) : Math.Min((right - left) / 3, CategoryLabelWidth(categories, labelSize)); }
        CategoryLabelPlan? categoryPlan = null;
        if (!horizontal)
        {
            left += valueLabelWidth;
            bool drawnIn3D = _is3D && _groups.Count == 1 && _groups[0].Kind is Kind.Bar or Kind.Line or Kind.Area;   // the 3-D paths place their own labels
            if (!xy && !drawnIn3D && categoryAxis is not { NoLabels: true })
            {
                int slots = MidCategory() ? Math.Max(1, categoryCount - 1) : categoryCount;
                categoryPlan = PlanCategoryLabels(Enumerable.Range(0, categoryCount).Select(i => i < categories.Count ? categories[i] : (i + 1).ToString(CultureInfo.InvariantCulture)).ToList(),
                                                  Math.Max(1, right - left) / slots, labelSize, categoryAxis);
                bottom -= categoryPlan.Height;
            }
            else bottom -= categoryAxis is { NoLabels: true } ? 0 : labelSize + 4;
        }
        double plotWidth = Math.Max(1, right - left), plotHeight = Math.Max(1, bottom - top);
        string gridColor = DocxDrawingColor.Css(new XElement("c", new XElement("schemeClr", new XAttribute("val", "tx1"), new XElement("lumMod", new XAttribute("val", "15000")), new XElement("lumOff", new XAttribute("val", "85000")))), _theme) ?? "#D9D9D9";
        string? plotFill = PropertiesFill(_plotAreaProperties, null);
        if (_is3D && !xy && _groups.Count == 1 && _groups[0].Kind == Kind.Bar)
        {
            Draw3DColumns(svg, categoryCount, horizontal, left, top, right, bottom, ValueToAxisOf(valueAxis, vMin, vMax), valueLabels, categories, categoryAxis is not { NoLabels: true }, textColor, gridColor, plotFill, labelSize);
            return;
        }
        if (_is3D && !xy && !horizontal && _groups.Count == 1 && _groups[0].Kind is Kind.Line or Kind.Area)
        {
            Draw3DRows(svg, categoryCount, left, top, right, bottom, ValueToAxisOf(valueAxis, vMin, vMax), valueLabels, categories, categoryAxis is not { NoLabels: true }, textColor, gridColor, plotFill, labelSize);
            return;
        }
        if (_is3D && !horizontal)
        {
            // the 3-D look: back wall, side wall and floor as a box seen from the front right; the data stand on the front plane
            double depth = Math.Min(plotWidth, plotHeight) * 0.08;
            string wallFill = PropertiesFill(_backWallProperties, plotFill) ?? "none";
            string floorFill = PropertiesFill(_floorProperties, plotFill) ?? "none";
            svg.Append("<path d=\"M").Append(N(left + depth)).Append(' ').Append(N(top)).Append('H').Append(N(right)).Append('V').Append(N(bottom - depth)).Append('H').Append(N(left + depth)).Append("Z\" fill=\"").Append(wallFill).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
            svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(top + depth)).Append('L').Append(N(left + depth)).Append(' ').Append(N(top)).Append('V').Append(N(bottom - depth)).Append('L').Append(N(left)).Append(' ').Append(N(bottom)).Append("Z\" fill=\"").Append(wallFill).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
            svg.Append("<path class=\"word-chart-floor\" d=\"M").Append(N(left)).Append(' ').Append(N(bottom)).Append('L').Append(N(left + depth)).Append(' ').Append(N(bottom - depth)).Append('H').Append(N(right)).Append('L').Append(N(right - depth)).Append(' ').Append(N(bottom)).Append("Z\" fill=\"").Append(floorFill).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
            top += depth; right -= depth;
            plotWidth = Math.Max(1, right - left); plotHeight = Math.Max(1, bottom - top);
        }
        else if (plotFill is not null) svg.Append("<rect x=\"").Append(N(left)).Append("\" y=\"").Append(N(top)).Append("\" width=\"").Append(N(plotWidth)).Append("\" height=\"").Append(N(plotHeight)).Append("\" fill=\"").Append(plotFill).Append("\"/>");
        Func<double, double> ValueToAxis = AxisMap(valueAxis, vMin, vMax);
        Func<double, double> SecondaryToAxis = AxisMap(secondaryAxis, sMin, sMax);
        // gridlines and value labels
        foreach ((double value, string text) in valueLabels)
        {
            double t = ValueToAxis(value);
            if (horizontal)
            {
                double x = left + t * plotWidth;
                if (valueAxis?.Gridlines != false) svg.Append("<path d=\"M").Append(N(x)).Append(' ').Append(N(top)).Append('V').Append(N(bottom)).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
                Text(svg, x, bottom + labelSize + 2, text, labelSize, valueAxis?.LabelColor ?? textColor, "middle");
            }
            else
            {
                double y = bottom - t * plotHeight;
                if (valueAxis?.Gridlines != false) svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(y)).Append('H').Append(N(right)).Append("\" stroke=\"").Append(gridColor).Append("\" stroke-width=\"0.75\"/>");
                // LibreOffice centres the label's full text height on the tick (measured: baseline 0.08 x size above it); Word a third of the size below
                // Word: 6.2 pt from the plot with outside tick marks (the LibreOffice-written charts), 8.2 pt without (Word's own charts,
                // c:majorTickMark none) - both measured on Word's PDFs
                Text(svg, left - (_libreOfficeLayout ? labelSize * 0.7 : valueAxis?.NoTickMarks == true ? 8.2 : 6.2), y + (_libreOfficeLayout ? -0.08 * labelSize : labelSize / 3), text, labelSize, valueAxis?.LabelColor ?? textColor, "end");
            }
        }
        foreach ((double value, string text) in secondaryLabels)
            Text(svg, right + 4, bottom - SecondaryToAxis(value) * plotHeight + labelSize / 3, text, labelSize, secondaryAxis?.LabelColor ?? textColor, "start");
        // category labels
        if (!xy && categoryAxis is not { NoLabels: true })
            for (int i = 0; i < categoryCount; i++)
            {
                string label = i < categories.Count ? categories[i] : (i + 1).ToString(CultureInfo.InvariantCulture);
                double centre = MidCategory() ? (categoryCount == 1 ? 0.5 : (double)i / (categoryCount - 1)) : (i + 0.5) / categoryCount;
                if (horizontal) Text(svg, left - 4, bottom - centre * plotHeight + labelSize / 3, label, labelSize, categoryAxis?.LabelColor ?? textColor, "end");
                else if (categoryPlan is { } plan)
                {
                    if (i % plan.Rhythm != 0) continue;
                    double x = left + centre * plotWidth;
                    if (plan.Rotated) Text(svg, x, bottom + labelSize, label, labelSize, categoryAxis?.LabelColor ?? textColor, "end", -45);
                    else
                    {
                        double y = bottom + labelSize + 2 + (_libreOfficeLayout ? 0 : labelSize / 3) + (plan.Staggered && i % 2 == 1 ? labelSize * 1.2 : 0);   // Word: a third of the size lower (measured)
                        foreach (string line in plan.Lines[i]) { Text(svg, x, y, line, labelSize, categoryAxis?.LabelColor ?? textColor, "middle"); y += labelSize * 1.2; }
                    }
                }
                else Text(svg, left + centre * plotWidth, bottom + labelSize + 2, label, labelSize, categoryAxis?.LabelColor ?? textColor, "middle");
            }
        if (xy && categoryAxis is not { NoLabels: true })
            for (double x = xMin; x <= xMax + xStep / 1000; x += xStep)
                Text(svg, left + (x - xMin) / (xMax - xMin) * plotWidth, bottom + labelSize + 2, FormatNumber(Math.Round(x, 10), "General"), labelSize, categoryAxis?.LabelColor ?? textColor, "middle");
        int totalSeries = _groups.Sum(g => g.Series.Count);
        foreach (Group group in _groups)
        {
            Func<double, double> axis = secondaryGroups.Contains(group) ? SecondaryToAxis : ValueToAxis;
            (double gMin, double gMax) = secondaryGroups.Contains(group) ? (sMin, sMax) : (vMin, vMax);
            switch (group.Kind)
            {
                case Kind.Bar: DrawBars(svg, group, categoryCount, left, top, plotWidth, plotHeight, axis, totalSeries, textColor); break;
                case Kind.Line or Kind.Area: DrawLines(svg, group, categoryCount, left, top, plotWidth, plotHeight, axis, totalSeries, textColor, gMin, gMax); break;
                case Kind.Scatter or Kind.Bubble: DrawScatter(svg, group, left, top, plotWidth, plotHeight, axis, x => (x - xMin) / (xMax - xMin), totalSeries, textColor); break;
            }
        }
        // axis lines
        string axisColor = gridColor;
        svg.Append("<path d=\"M").Append(N(left)).Append(' ').Append(N(bottom)).Append('H').Append(N(right)).Append("\" stroke=\"").Append(axisColor).Append("\" stroke-width=\"0.75\"/>");
    }

    /// <summary>c:crossBetween midCat (21.2.2.33): the value axis crosses at the categories, so points and labels sit on the category ticks, first and last at the plot's edges (only when no bars share the plot).</summary>
    private bool MidCategory() => ValueAxis() is { MidCategory: true } && _groups.All(g => g.Kind is Kind.Line or Kind.Area);

    /// <summary>How a horizontal category axis shows its labels when they do not fit side by side.</summary>
    private sealed record CategoryLabelPlan(List<string[]> Lines, bool Staggered, bool Rotated, int Rhythm, double Height);

    /// <summary>
    /// Crowded category labels, as LibreOffice decides (chart2 VCartesianAxis.cxx:513-575, 720-960): line breaks first where allowed
    /// (ODF text:line-break; the text wrapped into the tick distance less 5 %, given up when a single word does not fit), then, when
    /// labels still overlap, staggering (only with loext:try-staggering-first) or a 45-degree rotation, then showing every n-th label
    /// (the "rhythm") until they no longer overlap. Word wraps category text at spaces into its slot the same way and otherwise turns
    /// the labels 45 degrees (measured on Word's PDF of the crowded fixture chart). Text widths are estimated at 0.55 x the size per character.
    /// </summary>
    private CategoryLabelPlan PlanCategoryLabels(List<string> labels, double slot, double size, Axis? axis)
    {
        double Width(string text) => TextWidth(text, size);
        double room = slot * 0.95;
        var single = labels.Select(l => new[] { l }).ToList();
        if (labels.All(l => Width(l) <= room)) return new CategoryLabelPlan(single, false, false, 1, size + 4);
        bool wrapAllowed = axis?.LineBreak != false;
        if (wrapAllowed && labels.All(l => l.Split(' ').All(word => Width(word) <= room)))
        {
            var wrapped = labels.Select(label =>
            {
                var lines = new List<string>();
                string current = "";
                foreach (string word in label.Split(' '))
                {
                    string candidate = current.Length == 0 ? word : current + " " + word;
                    if (Width(candidate) <= room || current.Length == 0) current = candidate;
                    else { lines.Add(current); current = word; }
                }
                if (current.Length > 0) lines.Add(current);
                return lines.ToArray();
            }).ToList();
            return new CategoryLabelPlan(wrapped, false, false, 1, wrapped.Max(l => l.Length) * size * 1.2 + 4);
        }
        if (_libreOfficeLayout && axis?.TryStaggeringFirst == true && labels.All(l => Width(l) <= room * 2))
            return new CategoryLabelPlan(single, true, false, 1, size * 2.4 + 4);
        // turned 45 degrees: neighbours overlap when the text height along the axis (size x sqrt 2) exceeds the slot
        int rhythm = Math.Max(1, (int)Math.Ceiling(size * Math.Sqrt(2) / Math.Max(1e-6, slot)));
        return new CategoryLabelPlan(single, false, true, rhythm, labels.Max(Width) * Math.Sqrt(0.5) + size + 4);
    }

    private double CategoryLabelWidth(List<string> categories, double labelSize) => categories.DefaultIfEmpty("").Max(c => TextWidth(c, labelSize)) + 6;

    /// <summary>
    /// A text's width. LibreOffice layouts keep the 0.55 x size per character they were measured with. Word's use the chart font's own
    /// advance widths (DocxChartFontWidths: the hmtx tables of Calibri, Cambria, Times New Roman and Arial; metric-compatible fonts map
    /// to them); characters outside ASCII take the font's average width.
    /// </summary>
    private double TextWidth(string text, double size)
    {
        if (_libreOfficeLayout) return text.Length * size * 0.55;
        var font = DocxChartFontWidths.Fonts[_chartFont];
        double em = 0;
        foreach (char c in text) em += c >= ' ' && c <= '~' ? font.Widths[c - ' '] : c >= '\u0590' && c <= '\u05FF' ? DocxChartFontWidths.Hebrew[c - '\u0590'] : font.Average;
        return em * size;
    }

    /// <summary>The width table of a typeface: Calibri (Office's theme font, also when the chart names none or a theme font "+mn-lt"), Cambria, Times New Roman / Liberation Serif / Tinos, Arial / Liberation Sans / Arimo / Helvetica.</summary>
    private static int ChartFontIndex(string? typeface)
    {
        string name = (typeface ?? "").ToLowerInvariant();
        if (name.StartsWith("cambria", StringComparison.Ordinal)) return 1;
        if (name.StartsWith("times", StringComparison.Ordinal) || name.Contains("liberation serif") || name.StartsWith("tinos", StringComparison.Ordinal)) return 2;
        if (name.StartsWith("arial", StringComparison.Ordinal) || name.Contains("liberation sans") || name.StartsWith("arimo", StringComparison.Ordinal) || name.StartsWith("helvetica", StringComparison.Ordinal)) return 3;
        return 0;
    }

    private static Func<double, double> ValueToAxisOf(Axis? axis, double min, double max) => AxisMap(axis, min, max);

    /// <summary>The first a:defRPr / a:rPr sz under an element (points), or null.</summary>
    private static double? FontSize(XElement? element)
    {
        if (element is null) return null;
        foreach (XElement run in element.DescendantsAndSelf().Where(e => e.Name.LocalName is "defRPr" or "rPr"))
            if (double.TryParse((string?)run.Attribute("sz"), NumberStyles.Float, CultureInfo.InvariantCulture, out double hundredths) && hundredths > 0) return hundredths / 100;
        return null;
    }

    private double TitleSize(int line) => _titleSizes.Count == 0 ? 14 : _titleSizes[Math.Min(line, _titleSizes.Count - 1)];

    /// <summary>Value to its place on an axis (0 = start, 1 = end): linear, or logarithmic (values at or below the start sit on it).</summary>
    private static Func<double, double> AxisMap(Axis? axis, double min, double max)
    {
        if (axis?.LogBase is >= 2 and <= 1000 && min > 0)
        {
            double low = Math.Log(min), span = Math.Log(max) - low;
            return value => value <= min ? 0 : (Math.Log(value) - low) / span;
        }
        return value => (value - min) / (max - min);
    }

    // ── error bars and trend lines (per series, after its points are drawn) ──

    /// <summary>
    /// A series' c:errBars (21.2.2.55): direction (y, or x for scatter), plus / minus / both, the amount by c:errValType - fixedVal,
    /// percentage of the value, stdDev x val and stdErr over the series' values (LibreOffice: population variance,
    /// chart2/source/tools/StatisticsHelper.cxx:45-70, 179-187; drawn at each point, VSeriesPlotter.cxx:940-1010), cust (c:plus / c:minus);
    /// end caps unless c:noEndCap. And its c:trendline (21.2.2.209): linear, exp, log, power, poly (c:order), movingAvg (c:period), fitted
    /// by least squares over (x, value) - category charts use x = 1..n as Office and LibreOffice do - drawn across the points' x range;
    /// c:dispEq / c:dispRSqr add the equation and R squared (1 - SSE / SST, PolynomialRegressionCurveCalculator.cxx:81; the "R^2 = " text,
    /// VSeriesPlotter.cxx:1536).
    /// </summary>
    private void DrawSeriesExtras(StringBuilder svg, Series series, List<(double Px, double Py, double X, double Y, int Index)> points, Func<double, double> toPy, Func<double, double> toPx, string color, string textColor, bool stacked = false)
    {
        // stacked groups: the error bars sit on each stacked top; no trend lines (Word and LibreOffice offer none on stacked charts)
        if (points.Count == 0) return;
        foreach (XElement bars in series.ErrorBars)
        {
            string direction = Value(bars.Element("errDir")) ?? "y";
            string type = Value(bars.Element("errBarType")) ?? "both";
            string valueType = Value(bars.Element("errValType")) ?? "fixedVal";
            double amount = Number(bars.Element("val")) ?? 0;
            bool caps = Value(bars.Element("noEndCap")) is not ("1" or "true");
            string barColor = bars.Element("spPr")?.Element("ln")?.Element("solidFill") is { } solid ? DocxDrawingColor.Css(solid, _theme) ?? "#000000" : "#000000";
            var values = (direction == "x" ? points.Select(p => p.X) : points.Select(p => p.Y)).ToList();
            double mean = values.Average(), deviation = Math.Sqrt(values.Sum(v => (v - mean) * (v - mean)) / values.Count);
            (List<double?> plus, _) = Numbers(bars.Element("plus"));
            (List<double?> minus, _) = Numbers(bars.Element("minus"));
            foreach (var p in points)
            {
                double basis = direction == "x" ? p.X : p.Y;
                double up = valueType switch
                {
                    "percentage" => Math.Abs(basis) * amount / 100, "stdDev" => deviation * amount, "stdErr" => deviation / Math.Sqrt(values.Count),
                    "cust" => Custom(plus, bars.Element("plus"), p.Index), _ => amount,
                };
                double down = valueType == "cust" ? Custom(minus, bars.Element("minus"), p.Index) : up;
                double high = type is "both" or "plus" ? basis + up : basis, low = type is "both" or "minus" ? basis - down : basis;
                if (direction == "x")
                {
                    double x1 = toPx(low), x2 = toPx(high);
                    svg.Append("<path class=\"word-chart-error-bar\" d=\"M").Append(N(x1)).Append(' ').Append(N(p.Py)).Append('H').Append(N(x2));
                    if (caps) svg.Append("M").Append(N(x1)).Append(' ').Append(N(p.Py - 3)).Append("v6M").Append(N(x2)).Append(' ').Append(N(p.Py - 3)).Append("v6");
                }
                else
                {
                    double y1 = toPy(low), y2 = toPy(high);
                    svg.Append("<path class=\"word-chart-error-bar\" d=\"M").Append(N(p.Px)).Append(' ').Append(N(y1)).Append('V').Append(N(y2));
                    if (caps) svg.Append("M").Append(N(p.Px - 3)).Append(' ').Append(N(y1)).Append("h6M").Append(N(p.Px - 3)).Append(' ').Append(N(y2)).Append("h6");
                }
                svg.Append("\" fill=\"none\" stroke=\"").Append(barColor).Append("\" stroke-width=\"0.75\"/>");
            }
        }
        foreach (XElement trend in stacked ? new List<XElement>() : series.Trendlines)
        {
            string type = Value(trend.Element("trendlineType")) ?? "linear";
            var xs = points.Select(p => p.X).ToList();
            var ys = points.Select(p => p.Y).ToList();
            string lineColor = trend.Element("spPr")?.Element("ln")?.Element("solidFill") is { } solid ? DocxDrawingColor.Css(solid, _theme) ?? color : color;
            var path = new StringBuilder();
            // c:backward / c:forward (21.2.2.9 / 21.2.2.70): the line extended by that much along x
            double xLow = xs.Min() - Math.Max(0, Number(trend.Element("backward")) ?? 0), xHigh = xs.Max() + Math.Max(0, Number(trend.Element("forward")) ?? 0);
            if (type == "movingAvg")
            {
                int period = Math.Max(2, (int)(Number(trend.Element("period")) ?? 2));
                for (int i = period - 1; i < points.Count; i++)
                    path.Append(path.Length == 0 ? 'M' : 'L').Append(N(points[i].Px)).Append(' ').Append(N(toPy(ys.Skip(i - period + 1).Take(period).Average())));
            }
            else if (Fit(type, (int)(Number(trend.Element("order")) ?? 2), xs, ys, Number(trend.Element("intercept"))) is { } fit)
            {
                for (int k = 0; k <= 40; k++)
                {
                    double x = xLow + (xHigh - xLow) * k / 40, y = fit.Evaluate(x);
                    if (double.IsNaN(y) || double.IsInfinity(y)) continue;
                    path.Append(path.Length == 0 ? 'M' : 'L').Append(N(toPx(x))).Append(' ').Append(N(toPy(y)));
                }
                var lines = new List<string>();
                string prefix = (string?)trend.Attribute("equationPrefix") ?? "y = ";
                if (Value(trend.Element("dispEq")) is "1" or "true") lines.Add(prefix + fit.Equation);
                if (Value(trend.Element("dispRSqr")) is "1" or "true") lines.Add("R² = " + Math.Round(fit.RSquared, 4).ToString("0.####", CultureInfo.InvariantCulture));
                for (int i = 0; i < lines.Count; i++)
                    Text(svg, toPx(xHigh) - 2, toPy(fit.Evaluate(xHigh)) - 6 - (lines.Count - 1 - i) * 11, lines[i], 9, textColor, "end");
            }
            if (path.Length > 0)
                svg.Append("<path class=\"word-chart-trendline\" d=\"").Append(path).Append("\" fill=\"none\" stroke=\"").Append(lineColor).Append("\" stroke-width=\"1.5\"/>");
        }
    }

    /// <summary>A custom error amount of a point: its own, or the one value of a constant list (an ODF constant error, constant="1"), else 0.</summary>
    /// <summary>The ends of a series' fixed, percentage and custom error bars in one direction (x: around the x values, else 1..n; y: around the values).</summary>
    private static IEnumerable<double> ErrorBarEnds(Series series, string direction)
    {
        foreach (XElement bars in series.ErrorBars)
        {
            if ((Value(bars.Element("errDir")) ?? "y") != direction) continue;
            string valueType = Value(bars.Element("errValType")) ?? "fixedVal";
            if (valueType is not ("fixedVal" or "percentage" or "cust")) continue;
            string type = Value(bars.Element("errBarType")) ?? "both";
            double amount = Number(bars.Element("val")) ?? 0;
            (List<double?> plus, _) = Numbers(bars.Element("plus"));
            (List<double?> minus, _) = Numbers(bars.Element("minus"));
            int count = series.Values.Count;
            for (int i = 0; i < count; i++)
            {
                double? basis = direction == "x" ? (series.XValues.Count > 0 ? (i < series.XValues.Count ? series.XValues[i] : null) : i + 1) : series.Values[i];
                if (basis is not { } b) continue;
                double up = valueType == "cust" ? Custom(plus, bars.Element("plus"), i) : valueType == "percentage" ? Math.Abs(b) * amount / 100 : amount;
                double down = valueType == "cust" ? Custom(minus, bars.Element("minus"), i) : up;
                if (type is "both" or "plus") yield return b + up;
                if (type is "both" or "minus") yield return b - down;
            }
        }
    }

    private static double Custom(List<double?> amounts, XElement? source, int index)
        => (string?)source?.Attribute("constant") == "1" && amounts.Count > 0 ? amounts[0] ?? 0 : index < amounts.Count ? amounts[index] ?? 0 : 0;

    /// <summary>A least-squares trend: its function, equation text and R squared.</summary>
    private sealed record TrendFit(Func<double, double> Evaluate, string Equation, double RSquared);

    /// <param name="intercept">c:intercept (21.2.2.83): the value at x = 0 the line is forced through (linear, polynomial, exponential;
    /// LibreOffice's ForceIntercept / InterceptValue, xmloff PropertyMaps.cxx:222-225).</param>
    private static TrendFit? Fit(string type, int order, List<double> xs, List<double> ys, double? intercept = null)
    {
        // linearised forms: y = a + b x (linear), ln y = ln a + b x (exp), y = a + b ln x (log), ln y = ln a + b ln x (power)
        bool logX = type is "log" or "power", logY = type is "exp" or "power";
        var pairs = xs.Zip(ys, (x, y) => (x, y)).Where(p => (!logX || p.x > 0) && (!logY || p.y > 0)).ToList();
        int degree = type == "poly" ? Math.Max(2, Math.Min(6, order)) : 1;
        if (pairs.Count <= degree) return null;
        double[] u = pairs.Select(p => logX ? Math.Log(p.x) : p.x).ToArray(), v = pairs.Select(p => logY ? Math.Log(p.y) : p.y).ToArray();
        double? fixedConstant = intercept is { } forced && type is "linear" or "poly" or "exp" ? (type == "exp" ? (forced > 0 ? Math.Log(forced) : null) : forced) : null;
        double[]? c = PolynomialLeastSquares(u, v, degree, fixedConstant);
        if (c is null) return null;
        double Poly(double t) { double r = 0; for (int i = degree; i >= 0; i--) r = r * t + c[i]; return r; }
        Func<double, double> evaluate = type switch
        {
            "exp" => x => Math.Exp(Poly(x)), "log" => x => x > 0 ? Poly(Math.Log(x)) : double.NaN,
            "power" => x => x > 0 ? Math.Exp(Poly(Math.Log(x))) : double.NaN, _ => Poly,
        };
        double mean = pairs.Average(p => p.y), total = pairs.Sum(p => (p.y - mean) * (p.y - mean)), error = pairs.Sum(p => Math.Pow(p.y - evaluate(p.x), 2));
        double rSquared = total > 0 ? 1 - error / total : 1;
        string F(double value) => Math.Round(value, 4).ToString("0.####", CultureInfo.InvariantCulture);
        string Signed(double value) => (value < 0 ? " - " : " + ") + F(Math.Abs(value));
        string equation = type switch
        {
            "exp" => F(Math.Exp(c[0])) + "e^(" + F(c[1]) + "x)",
            "log" => F(c[1]) + " ln(x)" + Signed(c[0]),
            "power" => F(Math.Exp(c[0])) + "x^" + F(c[1]),
            _ => string.Concat(Enumerable.Range(1, degree).Reverse().Select((power, k) => (k == 0 ? F(c[power]) : Signed(c[power])) + (power == 1 ? "x" : "x^" + power))) + Signed(c[0]),
        };
        return new TrendFit(evaluate, equation, rSquared);
    }

    /// <summary>Coefficients c[0..degree] of the least-squares polynomial (normal equations, Gaussian elimination with partial pivoting); null when singular.</summary>
    /// <param name="fixedConstant">A forced constant term: only the powers 1..degree are fitted, to y minus it.</param>
    private static double[]? PolynomialLeastSquares(double[] x, double[] y, int degree, double? fixedConstant = null)
    {
        int first = fixedConstant is null ? 0 : 1, n = degree + 1 - first;
        double constant = fixedConstant ?? 0;
        var a = new double[n, n + 1];
        for (int row = 0; row < n; row++)
        {
            for (int column = 0; column < n; column++) a[row, column] = x.Sum(t => Math.Pow(t, row + column + 2 * first));
            a[row, n] = x.Select((t, i) => Math.Pow(t, row + first) * (y[i] - constant)).Sum();
        }
        for (int pivot = 0; pivot < n; pivot++)
        {
            int best = pivot;
            for (int row = pivot + 1; row < n; row++) if (Math.Abs(a[row, pivot]) > Math.Abs(a[best, pivot])) best = row;
            if (Math.Abs(a[best, pivot]) < 1e-12) return null;
            for (int column = 0; column <= n; column++) (a[pivot, column], a[best, column]) = (a[best, column], a[pivot, column]);
            for (int row = 0; row < n; row++)
            {
                if (row == pivot) continue;
                double factor = a[row, pivot] / a[pivot, pivot];
                for (int column = pivot; column <= n; column++) a[row, column] -= factor * a[pivot, column];
            }
        }
        double[] solved = Enumerable.Range(0, n).Select(i => a[i, n] / a[i, i]).ToArray();
        return fixedConstant is null ? solved : new[] { constant }.Concat(solved).ToArray();
    }

    /// <summary>
    /// Stock charts: c:hiLowLines = a line from each category's lowest to highest series value; c:upDownBars = a box from the first
    /// series (open) to the last (close), white when it closes higher and black when lower unless c:upBars / c:downBars give a fill,
    /// c:gapWidth between boxes (default 150) - LibreOffice draws the same boxes (chart2/source/view/charttypes/CandleStickChart.cxx:226-245,
    /// white and black boxes, the range line from minimum to maximum).
    /// </summary>
    private void DrawStock(StringBuilder svg, Group group, int categories, double left, double top, double width, double height, Func<double, double> axis)
    {
        string lineColor = group.HiLowLines?.Element("spPr")?.Element("ln")?.Element("solidFill") is { } solid ? DocxDrawingColor.Css(solid, _theme) ?? "#000000" : "#000000";
        double gap = Number(group.UpDownBars?.Element("gapWidth")) ?? 150;
        double slot = width / categories, boxWidth = slot / (1 + gap / 100);
        for (int i = 0; i < categories; i++)
        {
            var values = group.Series.Select(s => i < s.Values.Count ? s.Values[i] : null).ToList();
            var present = values.Where(v => v is not null).Select(v => v!.Value).ToList();
            if (present.Count == 0) continue;
            double x = left + (i + 0.5) * slot;
            if (group.HiLowLines is not null)
                svg.Append("<path class=\"word-chart-hilow\" d=\"M").Append(N(x)).Append(' ').Append(N(top + height - axis(present.Min()) * height)).Append('V').Append(N(top + height - axis(present.Max()) * height))
                   .Append("\" stroke=\"").Append(lineColor).Append("\" stroke-width=\"0.75\"/>");
            if (group.UpDownBars is not null && group.Series.Count >= 4 && values[0] is { } open && values[values.Count - 1] is { } close)
            {
                bool up = close >= open;
                XElement? bar = group.UpDownBars.Element(up ? "upBars" : "downBars")?.Element("spPr");
                string fill = PropertiesFill(bar, up ? "#FFFFFF" : "#000000") ?? "none";
                double y1 = top + height - axis(Math.Max(open, close)) * height, y2 = top + height - axis(Math.Min(open, close)) * height;
                svg.Append("<rect class=\"word-chart-updown\" x=\"").Append(N(x - boxWidth / 2)).Append("\" y=\"").Append(N(y1)).Append("\" width=\"").Append(N(boxWidth)).Append("\" height=\"").Append(N(Math.Max(1, y2 - y1)))
                   .Append("\" fill=\"").Append(fill).Append("\" stroke=\"").Append(lineColor).Append("\" stroke-width=\"0.75\"/>");
            }
        }
    }

    private void DrawBars(StringBuilder svg, Group group, int categories, double left, double top, double width, double height, Func<double, double> axis, int totalSeries, string textColor)
    {
        bool stacked = group.Grouping is "stacked" or "percentStacked", percent = group.Grouping == "percentStacked";
        int count = group.Series.Count;
        double slot = 1.0 / categories, gap = group.GapWidth / 100, overlap = group.Overlap / 100;
        // bar width b: count bars overlapping by overlap * b, plus a gap of gap * b, fill the category slot
        double bars = stacked ? 1 : count - (count - 1) * overlap;
        double barWidth = slot / (bars + gap);
        var extras = new Dictionary<Series, List<(double Px, double Py, double X, double Y, int Index)>>();
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
                    Text(svg, group.Horizontal ? x + w + 2 : x + w / 2, group.Horizontal ? y + h / 2 + 3 : y - 2 + (_libreOfficeLayout ? -0.7 : 0.23) * _dataLabelSize, label, _dataLabelSize, textColor, group.Horizontal ? "start" : "middle");
                if (!group.Horizontal && (series.ErrorBars.Count > 0 || series.Trendlines.Count > 0 && !stacked))
                {
                    if (!extras.TryGetValue(series, out var list)) extras[series] = list = new();
                    list.Add((x + w / 2, top + height - axis(to) * height, i + 1, to, i));
                }
            }
        }
        foreach (KeyValuePair<Series, List<(double Px, double Py, double X, double Y, int Index)>> entry in extras)
            DrawSeriesExtras(svg, entry.Key, entry.Value, v => top + height - axis(v) * height, x => left + (x - 0.5) / categories * width, FillOf(group, entry.Key, -1, totalSeries).Css, textColor, stacked);
    }

    private void DrawLines(StringBuilder svg, Group group, int categories, double left, double top, double width, double height, Func<double, double> axis, int totalSeries, string textColor, double vMin, double vMax)
    {
        bool stacked = group.Grouping is "stacked" or "percentStacked", percent = group.Grouping == "percentStacked";
        var baseline = new double[categories];
        if (group.IsStock && (group.HiLowLines is not null || group.UpDownBars is not null)) DrawStock(svg, group, categories, left, top, width, height, axis);
        for (int s = 0; s < group.Series.Count; s++)
        {
            Series series = group.Series[s];
            var points = new List<(double X, double Y, int Index)>();
            var pointValues = new List<double>();
            var lower = new List<(double X, double Y)>();
            for (int i = 0; i < categories && i < series.Values.Count; i++)
            {
                if (series.Values[i] is not { } raw) continue;
                double total = group.Series.Sum(x => Math.Abs(i < x.Values.Count ? x.Values[i] ?? 0 : 0));
                double value = percent ? (total == 0 ? 0 : raw / total) : raw;
                double from = stacked ? baseline[i] : Math.Max(vMin, Math.Min(vMax, 0));
                double to = stacked ? baseline[i] + value : value;
                if (stacked) baseline[i] = to;
                double x = left + (MidCategory() ? (categories == 1 ? 0.5 : (double)i / (categories - 1)) : (i + 0.5) / categories) * width;
                points.Add((x, top + height - axis(to) * height, i));
                pointValues.Add(to);
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
                if (LabelText(group, series, p.Index) is { } label) Text(svg, p.X, p.Y - 4, label, _dataLabelSize, textColor, "middle");
            if (series.ErrorBars.Count > 0 || series.Trendlines.Count > 0)
                DrawSeriesExtras(svg, series, points.Select((p, k) => (p.X, p.Y, (double)(p.Index + 1), pointValues[k], p.Index)).ToList(),
                                 v => top + height - axis(v) * height, x => left + (MidCategory() ? (categories == 1 ? 0.5 : (x - 1) / (categories - 1)) : (x - 0.5) / categories) * width,
                                 group.Kind == Kind.Area ? FillOf(group, series, -1, totalSeries).Css : LineOf(series, totalSeries).Css, textColor, stacked);
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
            var data = new List<(double X, double Y)>();
            for (int i = 0; i < series.Values.Count; i++)
            {
                if (series.Values[i] is not { } y) continue;
                double x = i < series.XValues.Count && series.XValues[i] is { } xv ? xv : i + 1;
                points.Add((left + axisX(x) * width, top + height - axisY(y) * height, i));
                data.Add((x, y));
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
                if (LabelText(group, series, p.Index) is { } label) Text(svg, p.X, p.Y - 4, label, _dataLabelSize, textColor, "middle");
            if (series.ErrorBars.Count > 0 || series.Trendlines.Count > 0)
                DrawSeriesExtras(svg, series, points.Select((p, k) => (p.X, p.Y, data[k].X, data[k].Y, p.Index)).ToList(),
                                 v => top + height - axisY(v) * height, x => left + axisX(x) * width, css == "none" ? AutomaticColor(series.Index, totalSeries) : css, textColor);
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
                Text(svg, cx + at * Math.Cos(middle), cy + at * Math.Sin(middle) + 3, label, _dataLabelSize, textColor, "middle");
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
        for (double v = min + step; v <= max + step / 1000 && v + step > v; v += step)
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
