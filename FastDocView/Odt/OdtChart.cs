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
            if (!groups.TryGetValue(seriesClass, out XElement? group)) { group = Group(seriesClass, plotArea, axes, Flag, ChartProperty); groups[seriesClass] = group; plot.Add(group); }

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

            if ((Labels(series, ChartProperty) ?? Labels(plotArea, ChartProperty)) is { } labels) element.Add(labels);

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
            group.Add(element);
            index++;
        }
        foreach (XElement group in groups.Values) group.Add(new XElement("axId", Val(1)), new XElement("axId", Val(2)));

        bool radialOnly = groups.Count > 0 && groups.Keys.All(k => k is "circle" or "ring");
        if (!radialOnly)
        {
            bool scatter = groups.Keys.Any(k => k is "scatter" or "bubble");
            plot.Add(Axis(xAxis, scatter ? "valAx" : "catAx", 1, 2, ChartProperty, styles));
            plot.Add(Axis(yAxis, "valAx", 2, 1, ChartProperty, styles));
        }
        // chart:wall: the plot area's fill
        if (plotArea?.Element(Chart + "wall") is { } wall)
        {
            string? fill = GraphicProperty(wall, Draw, "fill");
            string? color = GraphicProperty(wall, Draw, "fill-color");
            plot.Add(new XElement("spPr", fill == "none" ? new XElement("noFill") : color is not null ? Solid(color) : new XElement("noFill")));
        }

        var chartElement = new XElement("chart");
        string? title = TitleText(chart.Element(Chart + "title"));
        string? subtitle = TitleText(chart.Element(Chart + "subtitle"));
        string? heading = title is null ? subtitle : subtitle is null ? title : title + "\n" + subtitle;
        if (heading is not null) chartElement.Add(RichTitle(heading));
        else chartElement.Add(new XElement("autoTitleDeleted", new XAttribute("val", "1")));
        chartElement.Add(plot);
        if (chart.Element(Chart + "legend") is { } legend)
        {
            string position = (string?)legend.Attribute(Chart + "legend-position") switch
            {
                "start" => "l", "top" => "t", "bottom" => "b", "top-end" => "tr", _ => "r",
            };
            chartElement.Add(new XElement("legend", new XElement("legendPos", new XAttribute("val", position))));
        }
        var space = new XElement("chartSpace", chartElement);
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
        var element = new XElement(kind, new XElement("axId", Val(id)), scaling,
                                   new XElement("delete", Val(axis is null || chartProperty(axis, "display-label") == "false" && axis.Element(Chart + "title") is null ? 1 : 0)));
        if (axis?.Elements(Chart + "grid").Any(g => (string?)g.Attribute(Chart + "class") is null or "major") == true) element.Add(new XElement("majorGridlines"));
        if (TitleText(axis?.Element(Chart + "title")) is { } title) element.Add(RichTitle(title));
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

    private static XElement RichTitle(string text)
        => new("title", new XElement("tx", new XElement("rich", text.Split('\n').Select(line => new XElement("p", new XElement("r", new XElement("t", line)))))));

    private static XElement Solid(string color) => new("solidFill", new XElement("srgbClr", new XAttribute("val", color.TrimStart('#').ToUpperInvariant())));

    private static XAttribute Val(int value) => new("val", value.ToString(CultureInfo.InvariantCulture));

    /// <summary>"chart:bar" to "bar" (the class is a QName in the chart namespace).</summary>
    private static string? Local(string? qualified) => qualified is null ? null : qualified.Substring(qualified.IndexOf(':') + 1);
}
