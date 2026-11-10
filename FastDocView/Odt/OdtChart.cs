using System.Globalization;
using System.Xml.Linq;
using FastDocView.DocxHtml;

namespace FastDocView.Odt;

/// <summary>
/// An OpenDocument chart (ODF 1.3 part 3 chapter 11, chart:chart in an embedded "Object N" sub-document's content.xml, or in
/// a flat file the office:document inside the draw:object) read into the DrawingML chart vocabulary that <see cref="DocxChart"/>
/// draws, so .odt and .docx charts share one model and one SVG renderer. Read the way LibreOffice imports it (xmloff/source/chart):
/// chart:class (SchXMLTools.cxx:136-149 - line, area, circle = pie, ring = doughnut, scatter, radar / filled-radar, bar,
/// stock, bubble; surface is drawn as bar there, :149, and here), the plot-area style's chart:vertical / stacked / percentage /
/// three-dimensional (PropertyMaps.cxx:95-111; 3-D is drawn as its 2-D kind, as .docx charts are), chart:gap-width /
/// chart:overlap (PropertyMaps.cxx:174-175; LibreOffice's own defaults 100 / 0, chart2/source/model/template/ColumnChartType.cxx:116-119),
/// series from chart:values-cell-range-address / chart:label-cell-address (SchXMLSeries2Context.cxx:324-328) over the chart's own
/// table:table (SchXMLTableContext.cxx), chart:domain = the x values of a scatter chart and, for a bubble chart, the y values
/// then the x values with chart:values = the sizes (SchXMLSeries2Context.cxx:469-471, 572-643); a series without a domain
/// takes the first series' (the "first first domain", :581); axis chart:categories, chart:title, chart:grid class="major"
/// (SchXMLAxisContext.cxx:575-594) and the axis style's chart:minimum / maximum / interval-major / reverse-direction
/// (PropertyMaps.cxx:168-172); chart:title + chart:subtitle, chart:legend chart:legend-position (SchXMLChartContext.cxx:1129-1140);
/// chart:wall's fill as the plot-area fill (SchXMLPlotAreaContext.cxx:467); data labels from chart:data-label-number (value /
/// percentage / value-and-percentage), chart:data-label-text (the category) and chart:data-label-series (PropertyMaps.cxx:247-248);
/// colours from the series' and data points' style:graphic-properties (draw:fill-color, svg:stroke-color), else
/// LibreOffice's hard-coded fallback series colours (chart2/source/tools/ConfigColorScheme.cxx:134-139).
/// Not drawn: the floor (3-D only), logarithmic scales, secondary axes (series on them share the primary), error bars,
/// regression curves, stock up / down bars (the stock series are drawn as lines, as .docx stock charts are).
/// </summary>
internal static class OdtChart
{
    private static readonly XNamespace Chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
    private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
    private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
    private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
    private static readonly XNamespace Style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
    private static readonly XNamespace Draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
    private static readonly XNamespace Svg = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";

    /// <summary>ConfigColorScheme.cxx:135-139: the series colours when a series' style has none.</summary>
    private static readonly string[] FallbackColors = ["9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF", "000080", "FF00FF", "00FFFF", "FFFF00"];

    /// <summary>The chart of a chart sub-document (its office:document-content or flat office:document root); null when it holds no chart:chart.</summary>
    public static DocxChart? Read(XElement document)
    {
        XElement? chart = document.Element(Office + "body")?.Element(Office + "chart")?.Element(Chart + "chart");
        if (chart is null) return null;
        var styles = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (XElement style in document.Element(Office + "automatic-styles")?.Elements(Style + "style") ?? [])
            if ((string?)style.Attribute(Style + "name") is { } name) styles[name] = style;
        return DocxChart.Read(ChartSpace(chart, styles), null);
    }

    // ── the chart document's own data table ───────────────────────────────────────────────

    /// <summary>The local table as a grid (row, column; 0-based, header rows and columns included): a cell's number or text.</summary>
    private sealed class DataTable
    {
        private readonly List<List<(double? Number, string Text)>> _rows = new();

        public DataTable(XElement? table)
        {
            if (table is null) return;
            foreach (XElement row in table.Descendants(Table + "table-row"))
            {
                if (row.Ancestors(Table + "table").FirstOrDefault() != table) continue;
                var cells = new List<(double?, string)>();
                foreach (XElement cell in row.Elements())
                {
                    if (cell.Name != Table + "table-cell" && cell.Name != Table + "covered-table-cell") continue;
                    int repeat = Math.Max(1, Math.Min(1024, (int?)cell.Attribute(Table + "number-columns-repeated") ?? 1));
                    double? number = double.TryParse((string?)cell.Attribute(Office + "value"), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : null;
                    string text = string.Join("\n", cell.Elements(Text + "p").Select(ParagraphText));
                    for (int i = 0; i < repeat; i++) cells.Add((number, text));
                }
                int rows = Math.Max(1, Math.Min(1024, (int?)row.Attribute(Table + "number-rows-repeated") ?? 1));
                for (int i = 0; i < rows; i++) _rows.Add(cells);
            }
        }

        public (double? Number, string Text) Cell(int row, int column)
            => row >= 0 && row < _rows.Count && column >= 0 && column < _rows[row].Count ? _rows[row][column] : (null, "");

        /// <summary>The cells of a range address ("local-table.$B$2:.$B$4"; the first range of a list), row by row.</summary>
        public List<(double? Number, string Text)> Range(string? address)
        {
            var cells = new List<(double?, string)>();
            if (string.IsNullOrWhiteSpace(address)) return cells;
            string first = FirstRange(address!.Trim());
            int colon = ColonOutsideQuotes(first);
            (int Row, int Column)? start = CellAddress(colon < 0 ? first : first.Substring(0, colon));
            (int Row, int Column)? end = colon < 0 ? start : CellAddress(first.Substring(colon + 1));
            if (start is null || end is null) return cells;
            for (int r = Math.Min(start.Value.Row, end.Value.Row); r <= Math.Max(start.Value.Row, end.Value.Row); r++)
                for (int c = Math.Min(start.Value.Column, end.Value.Column); c <= Math.Max(start.Value.Column, end.Value.Column); c++)
                    cells.Add(Cell(r, c));
            return cells;
        }

        /// <summary>The first range of a space-separated list (a quoted table name may hold spaces).</summary>
        private static string FirstRange(string text)
        {
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == (char)0x27) quoted = !quoted;
                else if (text[i] == ' ' && !quoted) return text.Substring(0, i);
            }
            return text;
        }

        private static int ColonOutsideQuotes(string text)
        {
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\'') quoted = !quoted;
                else if (text[i] == ':' && !quoted) return i;
            }
            return -1;
        }

        /// <summary>"table.$B$2" / ".$B$2" / "B2" to (row, column), 0-based.</summary>
        private static (int, int)? CellAddress(string text)
        {
            int dot = text.LastIndexOf('.');
            string cell = (dot >= 0 ? text.Substring(dot + 1) : text).Replace("$", "");
            int i = 0, column = 0;
            while (i < cell.Length && char.IsLetter(cell[i])) { column = column * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1); i++; }
            if (i == 0 || !int.TryParse(cell.Substring(i), NumberStyles.None, CultureInfo.InvariantCulture, out int row) || row < 1) return null;
            return (row - 1, column - 1);
        }
    }

    // ── conversion to the DrawingML chart vocabulary (namespace-free, as DocxChart reads it) ──

    private static XElement ChartSpace(XElement chart, Dictionary<string, XElement> styles)
    {
        XElement? StyleOf(XElement? element) => (string?)element?.Attribute(Chart + "style-name") is { } name && styles.TryGetValue(name, out XElement? style) ? style : null;
        string? ChartProperty(XElement? element, string name) => (string?)StyleOf(element)?.Element(Style + "chart-properties")?.Attribute(Chart + name);
        string? GraphicProperty(XElement? element, XNamespace ns, string name) => (string?)StyleOf(element)?.Element(Style + "graphic-properties")?.Attribute(ns + name);

        var data = new DataTable(chart.Element(Table + "table"));
        XElement? plotArea = chart.Element(Chart + "plot-area");
        string chartClass = Local((string?)chart.Attribute(Chart + "class")) ?? "bar";
        bool Flag(string name) => ChartProperty(plotArea, name) == "true";
        List<XElement> axes = plotArea?.Elements(Chart + "axis").ToList() ?? new();
        XElement? xAxis = axes.FirstOrDefault(a => (string?)a.Attribute(Chart + "dimension") == "x");
        XElement? yAxis = axes.FirstOrDefault(a => (string?)a.Attribute(Chart + "dimension") == "y");

        // categories: the x axis' chart:categories range (bar, line, area, pie, radar, stock)
        List<string> categories = data.Range((string?)xAxis?.Element(Chart + "categories")?.Attribute(Table + "cell-range-address")).Select(c => CellText(c)).ToList();

        var plot = new XElement("plotArea", new XElement("layout"));
        var groups = new Dictionary<string, XElement>(StringComparer.Ordinal);
        int index = 0;
        List<string>? firstDomains = null;
        foreach (XElement series in plotArea?.Elements(Chart + "series") ?? [])
        {
            string seriesClass = Local((string?)series.Attribute(Chart + "class")) ?? chartClass;
            if (chartClass == "stock") seriesClass = "stock";
            if (seriesClass == "surface") seriesClass = "bar";
            // chart:attached-axis secondary-y (SchXMLSeries2Context.cxx:334): the series goes in a group on the secondary value axis
            bool onSecondary = (string?)series.Attribute(Chart + "attached-axis") == "secondary-y";
            string groupKey = onSecondary ? seriesClass + "#2" : seriesClass;
            if (!groups.TryGetValue(groupKey, out XElement? group)) { group = Group(seriesClass, plotArea, axes, Flag, ChartProperty); groups[groupKey] = group; plot.Add(group); }

            var element = new XElement("ser", new XElement("idx", Val(index)), new XElement("order", Val(index)));
            string? labelAddress = (string?)series.Attribute(Chart + "label-cell-address");
            if (labelAddress is not null && data.Range(labelAddress) is { Count: > 0 } label) element.Add(new XElement("tx", new XElement("v", CellText(label[0]))));
            bool varies = seriesClass is "circle" or "ring";
            element.Add(Properties(series, GraphicProperty, index, seriesClass is "line" or "scatter" or "radar" or "stock", varies));

            // data points: their own fill (a pie's slices), chart:repeated counting points
            int point = 0;
            foreach (XElement dataPoint in series.Elements(Chart + "data-point"))
            {
                int repeat = Math.Max(1, (int?)dataPoint.Attribute(Chart + "repeated") ?? 1);
                if (GraphicProperty(dataPoint, Draw, "fill-color") is { } pointFill)
                    for (int i = 0; i < repeat; i++)
                        element.Add(new XElement("dPt", new XElement("idx", Val(point + i)), new XElement("spPr", Solid(pointFill))));
                point += repeat;
            }

            if ((Labels(series, ChartProperty) ?? Labels(plotArea, ChartProperty)) is { } labels)
            {
                if (SizeOf((string?)series.Attribute(Chart + "style-name"), styles) is { } labelSize) labels.Add(TextProperties(labelSize));
                element.Add(labels);
            }

            List<string> domains = series.Elements(Chart + "domain").Select(d => (string?)d.Attribute(Table + "cell-range-address") ?? "").ToList();
            if (domains.Count > 0 && firstDomains is null) firstDomains = domains;
            if (domains.Count == 0 && firstDomains is not null) domains = firstDomains;
            var values = data.Range((string?)series.Attribute(Chart + "values-cell-range-address"));
            if (seriesClass == "bubble")
            {
                // values = sizes; first domain = y, second = x (SchXMLSeries2Context.cxx:469-471, 572-643)
                var y = domains.Count > 0 ? data.Range(domains[0]) : new();
                var x = domains.Count > 1 ? data.Range(domains[1]) : new();
                element.Add(Numbers("xVal", x), Numbers("yVal", y), Numbers("bubbleSize", values));
            }
            else if (seriesClass == "scatter")
            {
                var x = domains.Count > 0 ? data.Range(domains[0]) : new();
                element.Add(Numbers("xVal", x), Numbers("yVal", values));
            }
            else
            {
                if (categories.Count > 0) element.Add(Strings("cat", categories));
                element.Add(Numbers("val", values));
            }
            if (varies)
            {
                // a pie / ring slice without its own colour: LibreOffice's colour of that point (varied per point, ConfigColorScheme.cxx:135-139)
                var own = new HashSet<string>(element.Elements("dPt").Select(d => (string?)d.Element("idx")?.Attribute("val") ?? ""));
                for (int i = 0; i < values.Count; i++)
                    if (!own.Contains(i.ToString(CultureInfo.InvariantCulture)))
                        element.Add(new XElement("dPt", new XElement("idx", Val(i)), new XElement("spPr", Solid(FallbackColors[i % FallbackColors.Length]))));
            }
            if (seriesClass is "line" or "scatter" or "radar" or "stock")
            {
                string? symbol = ChartProperty(series, "symbol-type") ?? ChartProperty(plotArea, "symbol-type");
                if (symbol is null or "none") element.Add(new XElement("marker", new XElement("symbol", new XAttribute("val", "none"))));
                if (ChartProperty(plotArea, "interpolation") is "cubic-spline" or "b-spline") element.Add(new XElement("smooth", new XAttribute("val", "1")));
            }
            AddErrorBars(element, series, styles, data);
            AddTrendlines(element, series, styles);
            group.Add(element);
            index++;
        }
        foreach (KeyValuePair<string, XElement> entry in groups)
        {
            XElement group = entry.Value;
            group.Add(new XElement("axId", Val(1)), new XElement("axId", Val(entry.Key.EndsWith("#2", StringComparison.Ordinal) ? 3 : 2)));
            if (group.Name.LocalName == "stockChart") AddStockParts(group, plotArea, GraphicProperty);
        }

        bool radialOnly = groups.Count > 0 && groups.Keys.All(k => k is "circle" or "ring");
        if (!radialOnly)
        {
            bool scatter = groups.Keys.Any(k => k is "scatter" or "bubble");
            plot.Add(Axis(xAxis, scatter ? "valAx" : "catAx", 1, 2, ChartProperty, styles));
            plot.Add(Axis(yAxis, "valAx", 2, 1, ChartProperty, styles));
            if (groups.Keys.Any(k => k.EndsWith("#2", StringComparison.Ordinal)))
            {
                XElement? secondary = axes.FirstOrDefault(a => (string?)a.Attribute(Chart + "dimension") == "y" && (string?)a.Attribute(Chart + "name") == "secondary-y");
                XElement axis = Axis(secondary, "valAx", 3, 1, ChartProperty, styles);
                axis.Add(new XElement("axPos", new XAttribute("val", "r")));
                plot.Add(axis);
            }
        }
        // chart:wall: the plot area's fill
        if (plotArea?.Element(Chart + "wall") is { } wall)
        {
            string? fill = GraphicProperty(wall, Draw, "fill");
            string? color = GraphicProperty(wall, Draw, "fill-color");
            plot.Add(new XElement("spPr", fill == "none" ? new XElement("noFill") : color is not null ? Solid(color) : new XElement("noFill")));
        }
        // chart:coordinate-region (the plot rectangle without its axis labels, as LibreOffice placed it) - else chart:plot-area's own
        // rectangle with them - as fractions of chart:chart's size, so the plot sits where LibreOffice prints it
        double chartWidth = Length((string?)chart.Attribute(Svg + "width")), chartHeight = Length((string?)chart.Attribute(Svg + "height"));
        XElement? region = plotArea?.Element(Chart + "coordinate-region");
        XElement? placed = region ?? plotArea;
        if (chartWidth > 0 && chartHeight > 0 && placed?.Attribute(Svg + "width") is not null && placed.Attribute(Svg + "height") is not null)
        {
            double Fraction(string name, double size) => Length((string?)placed.Attribute(Svg + name)) / size;
            plot.Element("layout")!.Add(new XElement("manualLayout",
                new XElement("layoutTarget", new XAttribute("val", region is not null ? "inner" : "outer")),
                new XElement("xMode", new XAttribute("val", "edge")), new XElement("yMode", new XAttribute("val", "edge")),
                new XElement("x", new XAttribute("val", N(Fraction("x", chartWidth)))), new XElement("y", new XAttribute("val", N(Fraction("y", chartHeight)))),
                new XElement("w", new XAttribute("val", N(Fraction("width", chartWidth)))), new XElement("h", new XAttribute("val", N(Fraction("height", chartHeight))))));
        }

        var chartElement = new XElement("chart");
        string? title = TitleText(chart.Element(Chart + "title"));
        string? subtitle = TitleText(chart.Element(Chart + "subtitle"));
        string? heading = title is null ? subtitle : subtitle is null ? title : title + "\n" + subtitle;
        var headingSizes = (title is null ? new List<double?>() : LineSizes(chart.Element(Chart + "title"), styles)).Concat(subtitle is null ? new List<double?>() : LineSizes(chart.Element(Chart + "subtitle"), styles)).ToList();
        if (heading is not null) chartElement.Add(RichTitle(heading, headingSizes));
        else chartElement.Add(new XElement("autoTitleDeleted", new XAttribute("val", "1")));
        // 3-D: chart:three-dimensional on the plot area (PropertyMaps.cxx:95): the wall and chart:floor (SchXMLPlotAreaContext.cxx:467-471) give the box look
        if (Flag("three-dimensional"))
        {
            chartElement.Add(View3D(plotArea, ChartProperty));
            foreach (XElement group in plot.Elements().Where(e => e.Name.LocalName is "barChart" or "lineChart" or "areaChart" or "pieChart").ToList())
                group.Name = group.Name.LocalName.Replace("Chart", "3DChart");
            if (plot.Element("spPr") is { } wallFill) chartElement.Add(new XElement("backWall", new XElement(wallFill)));
            if (plotArea?.Element(Chart + "floor") is { } floor && GraphicProperty(floor, Draw, "fill") != "none" && GraphicProperty(floor, Draw, "fill-color") is { } floorColor)
                chartElement.Add(new XElement("floor", new XElement("spPr", Solid(floorColor))));
        }
        chartElement.Add(plot);
        if (chart.Element(Chart + "legend") is { } legend)
        {
            string position = (string?)legend.Attribute(Chart + "legend-position") switch
            {
                "start" => "l", "top" => "t", "bottom" => "b", "top-end" => "tr", _ => "r",
            };
            var legendElement = new XElement("legend", new XElement("legendPos", new XAttribute("val", position)));
            if (SizeOf((string?)legend.Attribute(Chart + "style-name"), styles) is { } legendSize) legendElement.Add(TextProperties(legendSize));
            if (chartWidth > 0 && chartHeight > 0 && legend.Attribute(Svg + "x") is not null && legend.Attribute(Svg + "y") is not null)
                legendElement.Add(new XElement("layout", new XElement("manualLayout",
                    new XElement("xMode", new XAttribute("val", "edge")), new XElement("yMode", new XAttribute("val", "edge")),
                    new XElement("x", new XAttribute("val", N(Length((string?)legend.Attribute(Svg + "x")) / chartWidth))),
                    new XElement("y", new XAttribute("val", N(Length((string?)legend.Attribute(Svg + "y")) / chartHeight))))));
            chartElement.Add(legendElement);
        }
        var space = new XElement("chartSpace", new XAttribute("layoutRules", "libreoffice"), chartElement);
        if (chartElement.Element("title") is { } titleElement && title is not null && subtitle is not null) titleElement.SetAttributeValue("separateLines", "1");
        // LibreOffice's line and area charts put the points on the category ticks (exported to .docx as c:crossBetween midCat)
        if (groups.Keys.All(k => k.StartsWith("line", StringComparison.Ordinal) || k.StartsWith("area", StringComparison.Ordinal)) && groups.Count > 0)
            foreach (XElement valueAxis in plot.Elements("valAx")) valueAxis.Element("crossAx")!.AddBeforeSelf(new XElement("crossBetween", new XAttribute("val", "midCat")));
        // the chart area: chart:chart's own style fill (LibreOffice writes a white area as draw:fill-color)
        string? areaFill = GraphicProperty(chart, Draw, "fill-color");
        string? areaFillKind = GraphicProperty(chart, Draw, "fill");
        if (areaFillKind == "none") space.Add(new XElement("spPr", new XElement("noFill")));
        else if (areaFill is not null) space.Add(new XElement("spPr", Solid(areaFill)));
        return space;
    }

    private static XElement Group(string seriesClass, XElement? plotArea, List<XElement> axes, Func<string, bool> flag, Func<XElement?, string, string?> chartProperty)
    {
        string grouping = flag("percentage") ? "percentStacked" : flag("stacked") ? "stacked" : seriesClass is "line" or "area" ? "standard" : "clustered";
        if (seriesClass == "area" && grouping == "clustered") grouping = "standard";
        XElement group = seriesClass switch
        {
            "line" => new XElement("lineChart", new XElement("grouping", new XAttribute("val", grouping))),
            "area" => new XElement("areaChart", new XElement("grouping", new XAttribute("val", grouping))),
            "circle" => new XElement("pieChart", new XElement("varyColors", Val(1))),
            "ring" => new XElement("doughnutChart", new XElement("varyColors", Val(1))),
            "scatter" => new XElement("scatterChart", new XElement("scatterStyle", new XAttribute("val", "lineMarker"))),
            "radar" or "filled-radar" => new XElement("radarChart"),
            "stock" => new XElement("stockChart"),
            "bubble" => new XElement("bubbleChart"),
            _ => new XElement("barChart", new XElement("barDir", new XAttribute("val", flag("vertical") ? "bar" : "col")), new XElement("grouping", new XAttribute("val", grouping))),
        };
        if (group.Name.LocalName == "barChart")
        {
            // gap / overlap: the plot area's, else an axis style's (LibreOffice writes them on the y axis); defaults 100 / 0
            string? Find(string name) => chartProperty(plotArea, name) ?? axes.Select(a => chartProperty(a, name)).FirstOrDefault(v => v is not null);
            group.Add(new XElement("gapWidth", new XAttribute("val", Find("gap-width") ?? "100")));
            if (Find("overlap") is { } overlap) group.Add(new XElement("overlap", new XAttribute("val", overlap)));
            else if (grouping == "clustered") group.Add(new XElement("overlap", Val(0)));
        }
        if (group.Name.LocalName == "doughnutChart") group.Add(new XElement("holeSize", Val(50)));
        return group;
    }

    private static XElement Axis(XElement? axis, string kind, int id, int crossId, Func<XElement?, string, string?> chartProperty, Dictionary<string, XElement> styles)
    {
        var scaling = new XElement("scaling");
        if (chartProperty(axis, "reverse-direction") == "true") scaling.Add(new XElement("orientation", new XAttribute("val", "maxMin")));
        if (chartProperty(axis, "maximum") is { } max) scaling.Add(new XElement("max", new XAttribute("val", max)));
        if (chartProperty(axis, "minimum") is { } min) scaling.Add(new XElement("min", new XAttribute("val", min)));
        if (chartProperty(axis, "logarithmic") == "true") scaling.AddFirst(new XElement("logBase", new XAttribute("val", "10")));
        var element = new XElement(kind, new XElement("axId", Val(id)), scaling,
                                   new XElement("delete", Val(axis is null || chartProperty(axis, "display-label") == "false" && axis.Element(Chart + "title") is null ? 1 : 0)));
        if (axis?.Elements(Chart + "grid").Any(g => (string?)g.Attribute(Chart + "class") is null or "major") == true) element.Add(new XElement("majorGridlines"));
        if (TitleText(axis?.Element(Chart + "title")) is { } title) element.Add(RichTitle(title, LineSizes(axis!.Element(Chart + "title"), styles)));
        if (SizeOf((string?)axis?.Attribute(Chart + "style-name"), styles) is { } labelSize) element.Add(TextProperties(labelSize));
        // label wrapping and staggering preferences of the axis style (text:line-break, loext:try-staggering-first; VCartesianAxis.cxx:513-575)
        XNamespace textNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0", loext = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
        XElement? axisChartProperties = (string?)axis?.Attribute(Chart + "style-name") is { } axisStyle && styles.TryGetValue(axisStyle, out XElement? axisStyleElement) ? axisStyleElement.Element(Style + "chart-properties") : null;
        element.SetAttributeValue("lineBreak", (string?)axisChartProperties?.Attribute(textNs + "line-break") == "true" ? "1" : "0");
        element.SetAttributeValue("tryStaggeringFirst", (string?)axisChartProperties?.Attribute(loext + "try-staggering-first") == "true" ? "1" : "0");
        if (chartProperty(axis, "interval-major") is { } step) element.Add(new XElement("majorUnit", new XAttribute("val", step)));
        element.Add(new XElement("crossAx", Val(crossId)));
        return element;
    }

    /// <summary>A series' fill and line from its style (draw:fill-color, svg:stroke-color, svg:stroke-width, draw:stroke none); a line series' line takes the fill colour when it has no stroke colour.</summary>
    private static XElement Properties(XElement series, Func<XElement?, XNamespace, string, string?> graphic, int index, bool isLine, bool varies)
    {
        string? own = graphic(series, Draw, "fill-color");
        string fill = own ?? "#" + FallbackColors[index % FallbackColors.Length];
        // a varying (pie / ring) series without its own colour: no series fill, so each point takes its own
        var properties = varies && own is null ? new XElement("spPr") : new XElement("spPr", Solid(fill));
        string? stroke = graphic(series, Svg, "stroke-color");
        if (isLine)
        {
            var line = new XElement("ln");
            if (graphic(series, Draw, "stroke") == "none") line.Add(new XElement("noFill"));
            else line.Add(Solid(stroke ?? fill));
            if (graphic(series, Svg, "stroke-width") is { } width && OdtStyleCascade.Twips(width) is { } twips && twips > 0)
                line.SetAttributeValue("w", ((long)Math.Round(twips / 20.0 * 12700)).ToString(CultureInfo.InvariantCulture));
            properties.Add(line);
        }
        return properties;
    }

    /// <summary>Data labels (chart:data-label-number value / percentage / value-and-percentage, data-label-text = the category, data-label-series); null when the style says nothing.</summary>
    private static XElement? Labels(XElement? element, Func<XElement?, string, string?> chartProperty)
    {
        string? number = chartProperty(element, "data-label-number");
        string? text = chartProperty(element, "data-label-text");
        string? series = chartProperty(element, "data-label-series");
        if (number is null && text is null && series is null) return null;
        XElement Show(string name, bool on) => new(name, Val(on ? 1 : 0));
        return new XElement("dLbls",
            Show("showVal", number is "value" or "value-and-percentage"), Show("showPercent", number is "percentage" or "value-and-percentage"),
            Show("showCatName", text == "true"), Show("showSerName", series == "true"));
    }

    /// <summary>
    /// chart:error-indicator (SchXMLSeries2Context.cxx:735) and the series style's error properties (PropertyMaps.cxx:204-210):
    /// chart:error-category constant / percentage / standard-deviation / standard-error / error-margin (PropertyMaps.cxx:302-307);
    /// LibreOffice reads the positive amount from chart:error-lower-limit and the negative one from chart:error-upper-limit
    /// (PropertyMaps.cxx:205-206, kept as it is); chart:error-upper-indicator / lower-indicator choose the sides.
    /// </summary>
    private static void AddErrorBars(XElement element, XElement series, Dictionary<string, XElement> styles, DataTable data)
    {
        var sources = series.Elements(Chart + "error-indicator").Select(e => ((string?)e.Attribute(Chart + "style-name"), (string?)e.Attribute(Chart + "dimension") ?? "y")).ToList();
        if (sources.Count == 0) sources.Add(((string?)series.Attribute(Chart + "style-name"), "y"));
        foreach ((string? styleName, string dimension) in sources)
        {
            if (styleName is null || !styles.TryGetValue(styleName, out XElement? style) || style.Element(Style + "chart-properties") is not { } properties) continue;
            string? category = (string?)properties.Attribute(Chart + "error-category");
            if (category is null or "none") continue;
            double Amount(string name) => double.TryParse((string?)properties.Attribute(Chart + name), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0;
            bool upper = (string?)properties.Attribute(Chart + "error-upper-indicator") == "true", lower = (string?)properties.Attribute(Chart + "error-lower-indicator") == "true";
            if (!upper && !lower) continue;
            string type = upper && lower ? "both" : upper ? "plus" : "minus";
            var bars = new XElement("errBars", new XElement("errDir", new XAttribute("val", dimension == "x" ? "x" : "y")), new XElement("errBarType", new XAttribute("val", type)));
            switch (category)
            {
                case "percentage": bars.Add(new XElement("errValType", new XAttribute("val", "percentage")), new XElement("val", new XAttribute("val", N(Amount("error-percentage"))))); break;
                case "error-margin":
                    // LibreOffice: the percentage of the series' largest value, the same at every point (VSeriesPlotter.cxx:824-842)
                    double largest = (element.Element("val") ?? element.Element("yVal"))?.Descendants("v").Select(v => double.TryParse(v.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d : double.NegativeInfinity).DefaultIfEmpty(0).Max() ?? 0;
                    double margin = double.IsInfinity(largest) ? 0 : largest * Amount("error-margin") / 100;
                    bars.Add(new XElement("errValType", new XAttribute("val", "cust")), Constant("plus", margin), Constant("minus", margin));
                    break;
                case "standard-deviation": bars.Add(new XElement("errValType", new XAttribute("val", "stdDev")), new XElement("val", new XAttribute("val", "1"))); break;
                case "standard-error": bars.Add(new XElement("errValType", new XAttribute("val", "stdErr"))); break;
                case "cell-range":
                    // chart:error-upper-range / lower-range (PropertyMaps.cxx:239-240): the amounts per point from the chart's own table
                    bars.Add(new XElement("errValType", new XAttribute("val", "cust")),
                             Numbers("plus", data.Range((string?)properties.Attribute(Chart + "error-upper-range"))), Numbers("minus", data.Range((string?)properties.Attribute(Chart + "error-lower-range"))));
                    break;
                default:
                    bars.Add(new XElement("errValType", new XAttribute("val", "cust")), Constant("plus", Amount("error-lower-limit")), Constant("minus", Amount("error-upper-limit")));
                    break;
            }
            if (style.Element(Style + "graphic-properties")?.Attribute(Svg + "stroke-color") is { } stroke)
                bars.Add(new XElement("spPr", new XElement("ln", Solid(stroke.Value))));
            element.Add(bars);
        }
    }

    /// <summary>A constant error amount for every point (a c:numLit of one value, repeated by the reader per point).</summary>
    private static XElement Constant(string name, double value)
        => new(name, new XElement("numLit", new XElement("ptCount", Val(1)), new XElement("pt", new XAttribute("idx", 0), new XElement("v", value.ToString("R", CultureInfo.InvariantCulture)))), new XAttribute("constant", "1"));

    /// <summary>
    /// chart:regression-curve (SchXMLRegressionCurveObjectContext.cxx): chart:regression-type linear / logarithmic / exponential / power /
    /// polynomial / moving-average, the degree and period (chart: or loext: regression-max-degree / regression-period, PropertyMaps.cxx:213-229),
    /// chart:equation chart:display-equation / display-r-square (:133-136); LibreOffice writes its equations as "f(x) = ..." (VSeriesPlotter.cxx:1475-1540).
    /// </summary>
    private static void AddTrendlines(XElement element, XElement series, Dictionary<string, XElement> styles)
    {
        XNamespace loext = "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0";
        foreach (XElement curve in series.Elements(Chart + "regression-curve"))
        {
            XElement? style = (string?)curve.Attribute(Chart + "style-name") is { } name && styles.TryGetValue(name, out XElement? found) ? found : null;
            XElement? properties = style?.Element(Style + "chart-properties");
            string? Property(string attribute) => (string?)properties?.Attribute(Chart + attribute) ?? (string?)properties?.Attribute(loext + attribute);
            string? type = Property("regression-type") switch
            {
                "linear" => "linear", "logarithmic" => "log", "exponential" => "exp", "power" => "power", "polynomial" => "poly", "moving-average" => "movingAvg", _ => null,
            };
            if (type is null) continue;
            var trend = new XElement("trendline", new XAttribute("equationPrefix", "f(x) = "));
            if (style?.Element(Style + "graphic-properties")?.Attribute(Svg + "stroke-color") is { } stroke) trend.Add(new XElement("spPr", new XElement("ln", Solid(stroke.Value))));
            trend.Add(new XElement("trendlineType", new XAttribute("val", type)));
            if (type == "poly") trend.Add(new XElement("order", new XAttribute("val", Property("regression-max-degree") ?? "2")));
            if (type == "movingAvg") trend.Add(new XElement("period", new XAttribute("val", Property("regression-period") ?? "2")));
            // extrapolation and a forced intercept (PropertyMaps.cxx:222-234: regression-extrapolate-forward / backward, regression-force-intercept, regression-intercept-value)
            if (Property("regression-extrapolate-forward") is { } forward) trend.Add(new XElement("forward", new XAttribute("val", forward)));
            if (Property("regression-extrapolate-backward") is { } backward) trend.Add(new XElement("backward", new XAttribute("val", backward)));
            if (Property("regression-force-intercept") == "true") trend.Add(new XElement("intercept", new XAttribute("val", Property("regression-intercept-value") ?? "0")));
            XElement? equation = curve.Element(Chart + "equation");
            trend.Add(new XElement("dispRSqr", Val((string?)equation?.Attribute(Chart + "display-r-square") == "true" ? 1 : 0)),
                      new XElement("dispEq", Val((string?)equation?.Attribute(Chart + "display-equation") == "true" ? 1 : 0)));
            element.Add(trend);
        }
    }

    /// <summary>
    /// Stock parts (SchXMLPlotAreaContext.cxx:481-489): chart:stock-range-line = the high-low line; chart:stock-gain-marker / loss-marker =
    /// the open-close boxes (their fills), drawn when the series include an opening value (four series; chart:japanese-candle-stick,
    /// PropertyMaps.cxx:93). LibreOffice draws the stock series only as these marks (CandleStickChart.cxx:226-260), so the series lines go.
    /// </summary>
    private static void AddStockParts(XElement group, XElement? plotArea, Func<XElement?, XNamespace, string, string?> graphic)
    {
        foreach (XElement series in group.Elements("ser"))
        {
            series.Element("spPr")?.Remove();
            series.Add(new XElement("spPr", new XElement("ln", new XElement("noFill"))));
        }
        var hiLow = new XElement("hiLowLines");
        if (graphic(plotArea?.Element(Chart + "stock-range-line"), Svg, "stroke-color") is { } rangeColor) hiLow.Add(new XElement("spPr", new XElement("ln", Solid(rangeColor))));
        group.Add(hiLow);
        if (group.Elements("ser").Count() >= 4)
        {
            XElement Bars(string name, string marker)
            {
                var bars = new XElement(name);
                if (graphic(plotArea?.Element(Chart + marker), Draw, "fill-color") is { } fill) bars.Add(new XElement("spPr", Solid(fill)));
                return bars;
            }
            group.Add(new XElement("upDownBars", new XElement("gapWidth", Val(150)), Bars("upBars", "stock-gain-marker"), Bars("downBars", "stock-loss-marker")));
        }
    }

    /// <summary>
    /// The 3-D scene of the plot area (ODF dr3d attributes, read by xmloff's SdXML3DSceneAttributesHelper): dr3d:vpn is the direction
    /// toward the viewer, so its turn about the vertical axis is rotY and its height rotX (LibreOffice writes the docx view3D from the same
    /// scene: the fixture's vpn (0.416, 0.174, 0.893) exports as rotY 25, rotX 10); dr3d:projection parallel = no perspective, else
    /// chart2's default perspective (20, i.e. OOXML 40); chart:right-angled-axes (PropertyMaps.cxx:160); dr3d:ambient-color and the first
    /// enabled dr3d:light's diffuse colour and direction for the shading.
    /// </summary>
    private static XElement View3D(XElement? plotArea, Func<XElement?, string, string?> chartProperty)
    {
        XNamespace dr3d = "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
        var view = new XElement("view3D");
        (double X, double Y, double Z)? Vector(string? text)
        {
            if (text is null) return null;
            string[] parts = text.Trim('(', ')', ' ').Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 3) return null;
            double[] v = parts.Select(p => double.TryParse(p, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) ? d : double.NaN).ToArray();
            return v.Any(double.IsNaN) ? null : (v[0], v[1], v[2]);
        }
        if (Vector((string?)plotArea?.Attribute(dr3d + "vpn")) is { } vpn)
        {
            double length = Math.Sqrt(vpn.X * vpn.X + vpn.Y * vpn.Y + vpn.Z * vpn.Z);
            if (length > 0)
            {
                view.Add(new XElement("rotX", new XAttribute("val", N(Math.Asin(Math.Max(-1, Math.Min(1, vpn.Y / length))) * 180 / Math.PI))));
                view.Add(new XElement("rotY", new XAttribute("val", N(Math.Atan2(vpn.X, vpn.Z) * 180 / Math.PI))));
            }
        }
        view.Add(new XElement("rAngAx", Val(chartProperty(plotArea, "right-angled-axes") == "true" ? 1 : 0)));
        view.Add(new XElement("perspective", Val((string?)plotArea?.Attribute(dr3d + "projection") == "parallel" ? 0 : 40)));
        double Grey(string? css) => css is { Length: 7 } && int.TryParse(css.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int red) ? red / 255.0 : double.NaN;
        if (Grey((string?)plotArea?.Attribute(dr3d + "ambient-color")) is var ambient && !double.IsNaN(ambient)) view.SetAttributeValue("ambient", N(ambient));
        XElement? light = plotArea?.Elements(dr3d + "light").FirstOrDefault(l => (string?)l.Attribute(dr3d + "enabled") == "true");
        if (Grey((string?)light?.Attribute(dr3d + "diffuse-color")) is var diffuse && !double.IsNaN(diffuse)) view.SetAttributeValue("light", N(diffuse));
        if (Vector((string?)light?.Attribute(dr3d + "direction")) is { } lightDirection)
            view.SetAttributeValue("lightDirection", N(lightDirection.X) + " " + N(lightDirection.Y) + " " + N(lightDirection.Z));
        return view;
    }

    private static double Length(string? value) => value is not null && OdtStyleCascade.Twips(value) is { } twips ? twips / 20.0 : 0;

    private static string N(double value) => Math.Round(value, 6).ToString("0.######", CultureInfo.InvariantCulture);

    private static XElement Numbers(string name, List<(double? Number, string Text)> cells)
    {
        var cache = new XElement("numCache", new XElement("ptCount", Val(cells.Count)));
        for (int i = 0; i < cells.Count; i++)
            if (cells[i].Number is { } value)
                cache.Add(new XElement("pt", new XAttribute("idx", i), new XElement("v", value.ToString("R", CultureInfo.InvariantCulture))));
        return new XElement(name, new XElement("numRef", cache));
    }

    private static XElement Strings(string name, List<string> texts)
    {
        var cache = new XElement("strCache", new XElement("ptCount", Val(texts.Count)));
        for (int i = 0; i < texts.Count; i++) cache.Add(new XElement("pt", new XAttribute("idx", i), new XElement("v", texts[i])));
        return new XElement(name, new XElement("strRef", cache));
    }

    private static string CellText((double? Number, string Text) cell)
        => cell.Text.Length > 0 ? cell.Text : cell.Number is { } n ? n.ToString("G15", CultureInfo.InvariantCulture) : "";

    private static string? TitleText(XElement? title)
    {
        if (title is null) return null;
        string text = string.Join("\n", title.Elements(Text + "p").Select(ParagraphText));
        return text.Length > 0 ? text : null;
    }

    /// <summary>A text:p's text with white space collapsed (ODF 1.3 part 3 6.1.2: runs of spaces, tabs and line ends are one space; a pretty-printed flat file indents them).</summary>
    private static string ParagraphText(XElement paragraph)
        => System.Text.RegularExpressions.Regex.Replace(paragraph.Value, @"\s+", " ").Trim();

    /// <summary>A title in DrawingML rich text, each line with its size (a:rPr sz, hundredths of a point) when known.</summary>
    private static XElement RichTitle(string text, List<double?>? sizes = null)
        => new("title", new XElement("tx", new XElement("rich", text.Split('\n').Select((line, i) =>
            new XElement("p", new XElement("r", sizes is not null && i < sizes.Count && sizes[i] is { } size ? new XElement("rPr", new XAttribute("sz", N(size * 100))) : null, new XElement("t", line)))))));

    /// <summary>A text size from a style's style:text-properties fo:font-size (points), or null.</summary>
    private static double? SizeOf(string? styleName, Dictionary<string, XElement> styles)
    {
        XNamespace fo = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        string? size = styleName is not null && styles.TryGetValue(styleName, out XElement? style) ? (string?)style.Element(Style + "text-properties")?.Attribute(fo + "font-size") : null;
        return size is not null && size.EndsWith("pt", StringComparison.Ordinal) && double.TryParse(size.Substring(0, size.Length - 2), NumberStyles.Float, CultureInfo.InvariantCulture, out double points) ? points : null;
    }

    /// <summary>Each text:p's size: its first span's text style, else the title element's chart style.</summary>
    private static List<double?> LineSizes(XElement? title, Dictionary<string, XElement> styles)
    {
        double? own = SizeOf((string?)title?.Attribute(Chart + "style-name"), styles);
        return title?.Elements(Text + "p").Select(p => SizeOf((string?)p.Descendants(Text + "span").FirstOrDefault()?.Attribute(Text + "style-name"), styles) ?? own).ToList() ?? new List<double?>();
    }

    private static XElement TextProperties(double size) => new("txPr", new XElement("p", new XElement("pPr", new XElement("defRPr", new XAttribute("sz", N(size * 100))))));

    private static XElement Solid(string color) => new("solidFill", new XElement("srgbClr", new XAttribute("val", color.TrimStart('#').ToUpperInvariant())));

    private static XAttribute Val(int value) => new("val", value.ToString(CultureInfo.InvariantCulture));

    /// <summary>"chart:bar" to "bar" (the class is a QName in the chart namespace).</summary>
    private static string? Local(string? qualified) => qualified is null ? null : qualified.Substring(qualified.IndexOf(':') + 1);
}
