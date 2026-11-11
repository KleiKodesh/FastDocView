using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.DocBinary;

namespace FastDocView.DocxHtml;

/// <summary>
/// The chart of an embedded Excel 97-2003 chart object ("Excel.Chart.8": the OLE object Word 97-2003 .doc and RTF files hold
/// for a chart) read into the DrawingML chart vocabulary that <see cref="DocxChart"/> draws, so .doc / RTF charts share the
/// .docx / ODT renderer (the same adapter pattern as <see cref="Odt.OdtChart"/>). Read the way LibreOffice imports a BIFF8
/// chart substream (sc/source/filter/excel/xichart.cxx, record ids in sc/source/filter/inc/xlchart.hxx:214-737):
/// CHSERIES (0x1003) + CHSOURCELINK (0x1051) + CHSTRING (0x100D) for series names (xichart.cxx:744-768, 2126); CHDATAFORMAT
/// (0x1006, :1406) with CHAREAFORMAT / CHLINEFORMAT (:453-499; BIFF8 RGB then palette index) and CHESCHERFORMAT (0x1066,
/// :528-560: the OfficeArt property set Office 2007+ writes with the exact colours, fillColor 0x0181 / lineColor 0x01C0);
/// CHTYPEGROUP (0x1014, :2683) with CHBAR (overlap, gap, flags; :2233-2237 - the DrawingML overlap is the BIFF one negated, :2407),
/// CHLINE / CHAREA (stacked / percent flags), CHPIE (rotation, hole; :2248-2252), CHSCATTER, CHRADARLINE / CHRADARAREA, CHCHART3D
/// (0x103A, :2433-2490: rotation = rotY, elevation = rotX, right-angled axes unless REAL3D), CHLEGEND (0x1015, :2516: rect,
/// dock mode 0 bottom / 1 corner / 2 top / 3 right / 4 left / 7 not docked, xlchart.hxx:328-333); CHAXESSET / CHAXIS (0x101D,
/// type 0 = x, 1 = y; :3251), CHVALUERANGE (0x101F, :3092: min, max, major, minor, cross, auto flags xlchart.hxx:422-426),
/// CHTICK (0x101E: major tick mark kind first), CHAXISLINE (0x1021: 1 = major grid, xlchart.hxx:444-447) + CHLINEFORMAT;
/// CHTEXT (0x1025, :951) + CHOBJECTLINK (0x1027: 1 title, 2 y axis, 3 x axis, 4 data label, xlchart.hxx:506-512) + CHSTRING +
/// CHFONT (0x1026: an index into the workbook's FONT records, index 4 skipped as in every BIFF font table).
/// Values: LibreOffice resolves each CHSOURCELINK formula against the embedded worksheet; this reader takes the chart's own
/// cache after the chart records instead ([MS-XLS] 2.4.252 SIIndex 0x1065: 1 = values, 2 = categories, 3 = bubble sizes,
/// then NUMBER / LABEL / BLANK records with row = point and column = series - checked on Word 16's own files, where it holds
/// the same numbers as the worksheet), so no formula engine is needed.
/// Not drawn: number formats of the source cells, per-point data labels, trend lines and error bars.
/// </summary>
internal static class ExcelBinaryChart
{
    private const ushort Bof = 0x0809, Eof = 0x000A, Font = 0x0031, Number = 0x0203, Label = 0x0204, Continue = 0x003C;
    private const ushort ChChart = 0x1002, ChSeries = 0x1003, ChDataFormat = 0x1006, ChLineFormat = 0x1007, ChAreaFormat = 0x100A,
        ChString = 0x100D, ChTypeGroup = 0x1014, ChLegend = 0x1015, ChBar = 0x1017, ChLine = 0x1018, ChPie = 0x1019, ChArea = 0x101A,
        ChScatter = 0x101B, ChAxis = 0x101D, ChTick = 0x101E, ChValueRange = 0x101F, ChAxisLine = 0x1021, ChText = 0x1025,
        ChFont = 0x1026, ChObjectLink = 0x1027, ChBegin = 0x1033, ChEnd = 0x1034, ChChart3D = 0x103A, ChRadarLine = 0x103E,
        ChRadarArea = 0x1040, ChAxesSet = 0x1041, ChSourceLink = 0x1051, ChEscherFormat = 0x1066, ChSeriesIndex = 0x1065;

    /// <summary>
    /// The chart of an RTF \objdata (hex digits; [MS-OLEDS] 2.2.4 ObjectHeader, then NativeDataSize and the compound file -
    /// LibreOffice filter/source/msfilter/rtfutil.cxx:267-330 ExtractOLE2FromObjdata); null when it is not an Excel chart.
    /// </summary>
    public static DocxChart? FromObjectData(byte[] objectData)
    {
        try
        {
            int at = 8;                                                     // OLEVersion, FormatID
            string className = AnsiString(objectData, ref at);
            if (!className.StartsWith("Excel.Chart", StringComparison.Ordinal)) return null;
            AnsiString(objectData, ref at);                                 // TopicName
            AnsiString(objectData, ref at);                                 // ItemName
            if (at + 4 > objectData.Length) return null;
            int size = BitConverter.ToInt32(objectData, at);
            at += 4;
            if (size <= 0 || at + (long)size > objectData.Length) return null;
            using var compound = new CompoundFileReader(new MemoryStream(objectData, at, size, writable: false));
            using Stream? workbook = compound.OpenStream("Workbook");
            return workbook is null ? null : FromWorkbook(workbook);
        }
        catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException or ArgumentException or IOException) { return null; }
    }

    /// <summary>The chart of a .doc ObjectPool object storage ("_" + its id, LibreOffice ww8par4.cxx:342); null when it holds no BIFF8 chart.</summary>
    public static DocxChart? FromObjectPool(CompoundFileReader document, uint objectId)
    {
        try
        {
            using Stream? workbook = document.OpenStream("ObjectPool", "_" + objectId.ToString(CultureInfo.InvariantCulture), "Workbook");
            return workbook is null ? null : FromWorkbook(workbook);
        }
        catch (Exception exception) when (exception is InvalidDataException or EndOfStreamException or ArgumentException or IOException) { return null; }
    }

    private static string AnsiString(byte[] data, ref int at)
    {
        if (at + 4 > data.Length) throw new InvalidDataException("Short OLE object header.");
        int length = BitConverter.ToInt32(data, at);
        at += 4;
        if (length < 0 || at + (long)length > data.Length) throw new InvalidDataException("Bad OLE object header.");
        string text = Encoding.ASCII.GetString(data, at, length).TrimEnd('\0');
        at += length;
        return text;
    }

    // ── the record walk ─────────────────────────────────────────────────────────────────────

    private sealed class SeriesData
    {
        public string? Name;
        public string? Fill, Line;
        public bool NoFill, NoLine;
        public Dictionary<int, string> PointFills = new();
        public bool ShowValue, ShowPercent, ShowCategory;
    }

    private sealed class AxisData
    {
        public int Type;
        public double? Min, Max, Major;
        public bool Gridlines, NoTickMarks, Deleted;
        public double? LabelSize;
        public string? Title;
        public double? TitleSize;
    }

    private sealed class TextData
    {
        public int Link = -1, Series = -1;
        public string? Text;
        public double? Size;
        public string? Typeface;
        public ushort Flags;
    }

    private sealed class FontEntry
    {
        public double Size;
        public string Name = "";
    }

    /// <summary>Reads the first chart substream (BOF type 0x0020) of a BIFF8 workbook stream; null when there is none.</summary>
    internal static DocxChart? FromWorkbook(Stream workbook)
    {
        var fonts = new List<FontEntry>();
        var series = new List<SeriesData>();
        var axes = new List<AxisData>();
        var texts = new List<TextData>();
        string kind = "";
        bool horizontal = false, stacked = false, percent = false, is3D = false, real3D = false, varyColors = false;
        double gap = 150, overlap = 0, hole = 0, rotation = 0, elevation = 15, rotY = 20, eyeDistance = 30;
        string? legendPosition = null;
        (double X, double Y)? legendAt = null;
        double? legendSize = null;
        int cacheKind = 0;
        var values = new Dictionary<(int Point, int Series), double>();
        var categories = new Dictionary<int, string>();
        var bubbles = new Dictionary<(int Point, int Series), double>();
        var xNumbers = new Dictionary<int, double>();

        // the object each BEGIN opens: what a following format / font / string record belongs to
        var stack = new Stack<object?>();
        object? last = null;                                     // the record that a following BEGIN opens
        object? current = null;                                  // the innermost open object
        SeriesData? formatSeries = null;
        int formatPoint = -1;
        bool inChart = false, chartDone = false, inLegend = false;
        int axisLine = -1;
        string? labelTypeface = null;                            // the axis labels' font: the chart's text widths
        int substreamType = -1;

        var header = new byte[4];
        while (ReadFully(workbook, header, 4))
        {
            ushort id = BitConverter.ToUInt16(header, 0);
            int length = BitConverter.ToUInt16(header, 2);
            var body = new byte[length];
            if (!ReadFully(workbook, body, length)) break;
            if (id == Bof && length >= 4)
            {
                substreamType = BitConverter.ToUInt16(body, 2);
                if (substreamType == 0x0020 && !chartDone) inChart = true;
                continue;
            }
            if (id == Eof)
            {
                if (inChart) break;
                continue;
            }
            if (id == Font && length >= 16 && substreamType == 0x0005)
            {
                var font = new FontEntry { Size = BitConverter.ToUInt16(body, 0) / 20.0 };
                int count = body[14];
                bool wide = (body[15] & 1) != 0;
                if (16 + count * (wide ? 2 : 1) <= length)
                    font.Name = wide ? Encoding.Unicode.GetString(body, 16, count * 2) : Latin1(body, 16, count);
                fonts.Add(font);
                continue;
            }
            if (!inChart) continue;

            switch (id)
            {
                case ChBegin:
                    stack.Push(current);
                    current = last;
                    continue;
                case ChEnd:
                    if (current is TextData closed && closed.Link == 4 && closed.Series >= 0 && closed.Series < series.Count)
                    {
                        series[closed.Series].ShowValue |= (closed.Flags & 0x0004) != 0;
                        series[closed.Series].ShowPercent |= (closed.Flags & 0x1000) != 0;
                        series[closed.Series].ShowCategory |= (closed.Flags & 0x4000) != 0;
                    }
                    if (current is string marker && marker == "legend") inLegend = false;
                    current = stack.Count > 0 ? stack.Pop() : null;
                    if (stack.Count == 0) chartDone = true;
                    last = null;
                    continue;
                case ChChart:
                    last = "chart";
                    continue;
                case ChSeries:
                    var created = new SeriesData();
                    series.Add(created);
                    last = created;
                    continue;
                case ChString when length >= 5:
                    string text = UnicodeString(body, 2, 1);
                    if (current is SeriesData owner && owner.Name is null) owner.Name = text;
                    else if (current is TextData textData) textData.Text = text;
                    continue;
                case ChDataFormat when length >= 4:
                    formatPoint = BitConverter.ToUInt16(body, 0);
                    int seriesIndex = BitConverter.ToUInt16(body, 2);
                    formatSeries = current as SeriesData ?? (seriesIndex < series.Count ? series[seriesIndex] : null);
                    last = "dataformat";
                    continue;
                case ChAreaFormat when length >= 12 && current is "dataformat" && formatSeries is not null:
                {
                    bool automatic = (BitConverter.ToUInt16(body, 10) & 1) != 0;
                    bool none = BitConverter.ToUInt16(body, 8) == 0;
                    string color = Rgb(body, 0);
                    if (formatPoint == 0xFFFF)
                    {
                        if (none) formatSeries.NoFill = true;
                        else if (!automatic) formatSeries.Fill = color;
                    }
                    else if (!automatic && !none) formatSeries.PointFills[formatPoint] = color;
                    continue;
                }
                case ChLineFormat when length >= 10:
                {
                    int pattern = BitConverter.ToUInt16(body, 4);
                    bool automatic = (BitConverter.ToUInt16(body, 8) & 1) != 0;
                    if (current is "dataformat" && formatSeries is not null && formatPoint == 0xFFFF)
                    {
                        if (pattern == 5) formatSeries.NoLine = true;
                        else if (!automatic) formatSeries.Line = Rgb(body, 0);
                    }
                    else if (current is AxisData lined && axisLine == 1) lined.Gridlines = pattern != 5;
                    continue;
                }
                case ChEscherFormat when current is "dataformat" && formatSeries is not null:
                {
                    (string? fill, string? line) = EscherColors(body);
                    if (formatPoint == 0xFFFF)
                    {
                        if (fill is not null && !formatSeries.NoFill) formatSeries.Fill = fill;
                        if (line is not null && !formatSeries.NoLine) formatSeries.Line = line;
                    }
                    else if (fill is not null) formatSeries.PointFills[formatPoint] = fill;
                    continue;
                }
                case ChTypeGroup when length >= 18:
                    varyColors = (BitConverter.ToUInt16(body, 16) & 1) != 0;
                    last = "typegroup";
                    continue;
                case ChBar when length >= 6 && kind.Length == 0:
                    kind = "bar";
                    overlap = -BitConverter.ToInt16(body, 0);
                    gap = BitConverter.ToInt16(body, 2);
                    ushort barFlags = BitConverter.ToUInt16(body, 4);
                    horizontal = (barFlags & 1) != 0;
                    stacked = (barFlags & 2) != 0;
                    percent = (barFlags & 4) != 0;
                    continue;
                case ChLine when length >= 2 && kind.Length == 0:
                case ChArea when length >= 2 && kind.Length == 0:
                    kind = id == ChLine ? "line" : "area";
                    stacked = (body[0] & 1) != 0;
                    percent = (body[0] & 2) != 0;
                    continue;
                case ChPie when length >= 4 && kind.Length == 0:
                    rotation = BitConverter.ToUInt16(body, 0);
                    hole = BitConverter.ToUInt16(body, 2);
                    kind = hole > 0 ? "doughnut" : "pie";
                    continue;
                case ChScatter when kind.Length == 0:
                    kind = length >= 6 && (BitConverter.ToUInt16(body, 4) & 1) != 0 ? "bubble" : "scatter";
                    continue;
                case ChRadarLine when kind.Length == 0:
                case ChRadarArea when kind.Length == 0:
                    kind = id == ChRadarArea ? "filledRadar" : "radar";
                    continue;
                case ChChart3D when length >= 14:
                    is3D = true;
                    rotY = BitConverter.ToUInt16(body, 0);
                    elevation = BitConverter.ToInt16(body, 2);
                    eyeDistance = BitConverter.ToUInt16(body, 4);
                    real3D = (BitConverter.ToUInt16(body, 12) & 1) != 0;
                    continue;
                case ChLegend when length >= 18:
                {
                    int dock = body[16];
                    legendPosition = dock switch { 0 => "b", 1 => "tr", 2 => "t", 4 => "l", 3 => "r", _ => null };
                    double x = BitConverter.ToInt32(body, 0) / 4000.0, y = BitConverter.ToInt32(body, 4) / 4000.0;
                    double w = BitConverter.ToInt32(body, 8) / 4000.0, h = BitConverter.ToInt32(body, 12) / 4000.0;
                    if (legendPosition is null && x >= 0 && y >= 0 && x <= 1 && y <= 1)
                    {
                        // not docked (Office writes Word's legends so): its side from where its box sits, at its own place
                        legendPosition = y >= 0.75 ? "b" : y + h <= 0.25 ? "t" : x + w <= 0.25 ? "l" : "r";
                        legendAt = (x, y);
                    }
                    inLegend = true;
                    last = "legend";
                    continue;
                }
                case ChAxis when length >= 2:
                    var axis = new AxisData { Type = BitConverter.ToUInt16(body, 0) };
                    axes.Add(axis);
                    last = axis;
                    continue;
                case ChValueRange when length >= 42 && current is AxisData ranged:
                {
                    ushort flags = BitConverter.ToUInt16(body, 40);
                    if ((flags & 1) == 0) ranged.Min = BitConverter.ToDouble(body, 0);
                    if ((flags & 2) == 0) ranged.Max = BitConverter.ToDouble(body, 8);
                    if ((flags & 4) == 0) ranged.Major = BitConverter.ToDouble(body, 16);
                    continue;
                }
                case ChTick when length >= 1 && current is AxisData ticked:
                    ticked.NoTickMarks = body[0] is 0 or 1;            // none / inside: no mark outside the plot
                    if (length >= 3 && body[2] == 0) ticked.Deleted = true;   // label position none
                    continue;
                case ChAxisLine when length >= 2:
                    axisLine = BitConverter.ToUInt16(body, 0);
                    continue;
                case ChText when length >= 26:
                    var textRecord = new TextData { Flags = BitConverter.ToUInt16(body, 24) };
                    texts.Add(textRecord);
                    last = textRecord;
                    continue;
                case ChObjectLink when length >= 4 && current is TextData linked:
                    linked.Link = BitConverter.ToUInt16(body, 0);
                    linked.Series = BitConverter.ToUInt16(body, 2);
                    continue;
                case ChFont when length >= 2:
                {
                    int index = BitConverter.ToUInt16(body, 0);
                    int at = index >= 4 ? index - 1 : index;             // BIFF font tables skip index 4
                    if (at < 0 || at >= fonts.Count) continue;
                    double size = fonts[at].Size;
                    if (current is TextData sized) { sized.Size = size; sized.Typeface = fonts[at].Name; if (inLegend) legendSize = size; }
                    else if (current is AxisData axisFont) { axisFont.LabelSize = size; labelTypeface ??= fonts[at].Name; }
                    continue;
                }
                case ChSeriesIndex when length >= 2 && chartDone:
                    cacheKind = BitConverter.ToUInt16(body, 0);
                    continue;
                case Number when length >= 14 && chartDone:
                {
                    int row = BitConverter.ToUInt16(body, 0), column = BitConverter.ToUInt16(body, 2);
                    double value = BitConverter.ToDouble(body, 6);
                    if (cacheKind == 1) values[(row, column)] = value;
                    else if (cacheKind == 2) { if (!categories.ContainsKey(row)) categories[row] = value.ToString("R", CultureInfo.InvariantCulture); if (!xNumbers.ContainsKey(row)) xNumbers[row] = value; }
                    else if (cacheKind == 3) bubbles[(row, column)] = value;
                    continue;
                }
                case Label when length >= 9 && chartDone && cacheKind == 2:
                {
                    int row = BitConverter.ToUInt16(body, 0);
                    if (!categories.ContainsKey(row)) categories[row] = UnicodeString(body, 6, 2);
                    continue;
                }
                case Continue:
                    continue;
                default:
                    if (id is >= 0x1000 and < 0x1100) last = id;     // a record whose BEGIN block is not read
                    continue;
            }
        }
        if (series.Count == 0 || kind.Length == 0) return null;

        foreach (TextData text in texts)
            if (text.Link is 2 or 3 && text.Text is not null && axes.FirstOrDefault(a => a.Type == (text.Link == 2 ? 1 : 0)) is { } titled)
            {
                titled.Title = text.Text;
                titled.TitleSize = text.Size;
            }
        TextData? title = texts.FirstOrDefault(t => t.Link == 1);
        // the chart's text widths: the title's font when there is a title (Office 2007+ writes the theme font there - Calibri -
        // while the axis FONT records fall back to Arial), else the axis labels' font
        string? typeface = title?.Typeface is { Length: > 0 } titleFont ? titleFont : labelTypeface;
        return DocxChart.Read(ChartSpace(), null);

        XElement ChartSpace()
        {
            int pointCount = Math.Max(categories.Count == 0 ? 0 : categories.Keys.Max() + 1, values.Count == 0 ? 0 : values.Keys.Max(k => k.Point) + 1);
            string groupName = kind switch
            {
                "bar" => is3D ? "bar3DChart" : "barChart", "line" => is3D ? "line3DChart" : "lineChart", "area" => is3D ? "area3DChart" : "areaChart",
                "pie" => is3D ? "pie3DChart" : "pieChart", "doughnut" => "doughnutChart", "scatter" => "scatterChart", "bubble" => "bubbleChart",
                _ => "radarChart",
            };
            var group = new XElement(groupName);
            if (kind == "bar")
            {
                group.Add(new XElement("barDir", Val(horizontal ? "bar" : "col")));
                group.Add(new XElement("grouping", Val(percent ? "percentStacked" : stacked ? "stacked" : "clustered")));
            }
            else if (kind is "line" or "area") group.Add(new XElement("grouping", Val(percent ? "percentStacked" : stacked ? "stacked" : "standard")));
            else if (kind == "scatter") group.Add(new XElement("scatterStyle", Val("lineMarker")));
            else if (kind is "radar" or "filledRadar") group.Add(new XElement("radarStyle", Val(kind == "filledRadar" ? "filled" : "marker")));
            bool radial = kind is "pie" or "doughnut";
            group.Add(new XElement("varyColors", Val(varyColors || radial ? "1" : "0")));
            for (int s = 0; s < series.Count; s++)
            {
                SeriesData data = series[s];
                var element = new XElement("ser", new XElement("idx", Val(s)), new XElement("order", Val(s)));
                if (data.Name is not null) element.Add(new XElement("tx", new XElement("v", data.Name)));
                bool lined = kind is "line" or "scatter" or "radar";
                var shape = new XElement("spPr");
                if (lined)
                {
                    if (data.NoLine) shape.Add(new XElement("ln", new XElement("noFill")));
                    else if ((data.Line ?? data.Fill) is { } stroke) shape.Add(new XElement("ln", new XAttribute("w", "28575"), Solid(stroke)));
                }
                else
                {
                    if (data.NoFill) shape.Add(new XElement("noFill"));
                    else if (data.Fill is { } fill) shape.Add(Solid(fill));
                    shape.Add(new XElement("ln", new XElement("noFill")));
                }
                if (shape.HasElements) element.Add(shape);
                foreach (KeyValuePair<int, string> point in data.PointFills.OrderBy(p => p.Key))
                    element.Add(new XElement("dPt", new XElement("idx", Val(point.Key)), new XElement("spPr", Solid(point.Value))));
                if (data.ShowValue || data.ShowPercent || data.ShowCategory)
                    element.Add(new XElement("dLbls", new XElement("showVal", Val(data.ShowValue ? "1" : "0")), new XElement("showCatName", Val(data.ShowCategory ? "1" : "0")),
                                             new XElement("showPercent", Val(data.ShowPercent ? "1" : "0"))));
                var cells = new XElement("numCache", new XElement("ptCount", Val(pointCount)));
                for (int p = 0; p < pointCount; p++)
                    if (values.TryGetValue((p, s), out double v)) cells.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", v.ToString("R", CultureInfo.InvariantCulture))));
                if (kind is "scatter" or "bubble")
                {
                    var x = new XElement("numCache", new XElement("ptCount", Val(pointCount)));
                    for (int p = 0; p < pointCount; p++)
                        x.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", (xNumbers.TryGetValue(p, out double xv) ? xv : p + 1).ToString("R", CultureInfo.InvariantCulture))));
                    element.Add(new XElement("xVal", new XElement("numRef", x)), new XElement("yVal", new XElement("numRef", cells)));
                    if (kind == "bubble")
                    {
                        var sizes = new XElement("numCache", new XElement("ptCount", Val(pointCount)));
                        for (int p = 0; p < pointCount; p++)
                            if (bubbles.TryGetValue((p, s), out double b)) sizes.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", b.ToString("R", CultureInfo.InvariantCulture))));
                        element.Add(new XElement("bubbleSize", new XElement("numRef", sizes)));
                    }
                }
                else
                {
                    var names = new XElement("strCache", new XElement("ptCount", Val(pointCount)));
                    for (int p = 0; p < pointCount; p++)
                        names.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", categories.TryGetValue(p, out string? c) ? c : (p + 1).ToString(CultureInfo.InvariantCulture))));
                    element.Add(new XElement("cat", new XElement("strRef", names)), new XElement("val", new XElement("numRef", cells)));
                }
                if (lined && kind != "scatter") element.Add(new XElement("marker", new XElement("symbol", Val("none"))));
                group.Add(element);
            }
            if (kind == "bar") group.Add(new XElement("gapWidth", Val(N(gap))), new XElement("overlap", Val(N(overlap))));
            if (radial) group.Add(new XElement("firstSliceAng", Val(N(rotation))));
            if (kind == "doughnut") group.Add(new XElement("holeSize", Val(N(hole))));

            var plot = new XElement("plotArea", new XElement("layout"), group);
            if (!radial)
            {
                group.Add(new XElement("axId", Val(1)), new XElement("axId", Val(2)));
                AxisData? x = axes.FirstOrDefault(a => a.Type == 0), y = axes.FirstOrDefault(a => a.Type == 1);
                plot.Add(Axis(x, kind is "scatter" or "bubble" ? "valAx" : "catAx", 1, 2, horizontal ? "l" : "b"));
                plot.Add(Axis(y, "valAx", 2, 1, horizontal ? "b" : "l"));
            }

            var chart = new XElement("chart");
            if (title?.Text is { Length: > 0 } heading)
                chart.Add(new XElement("title", RichText(heading, title.Size)));
            else chart.Add(new XElement("autoTitleDeleted", Val("1")));
            if (is3D)
            {
                var view = new XElement("view3D", new XElement("rotX", Val(N(elevation))), new XElement("rotY", Val(N(rotY))));
                if (real3D) view.Add(new XElement("perspective", Val(N(eyeDistance * 2))));
                else view.Add(new XElement("rAngAx", Val("1")));
                chart.Add(view);
            }
            chart.Add(plot);
            if (legendPosition is not null)
            {
                var legend = new XElement("legend", new XElement("legendPos", Val(legendPosition)));
                if (legendSize is { } size) legend.Add(TextProperties(size));
                chart.Add(legend);
            }
            var space = new XElement("chartSpace", chart);
            if (typeface is { Length: > 0 })
                space.Add(new XElement("txPr", new XElement("p", new XElement("pPr", new XElement("defRPr", new XElement("latin", new XAttribute("typeface", typeface)))))));
            return space;
        }

        XElement Axis(AxisData? axis, string name, int id, int cross, string position)
        {
            var scaling = new XElement("scaling");
            if (axis?.Max is { } max) scaling.Add(new XElement("max", Val(N(max))));
            if (axis?.Min is { } min) scaling.Add(new XElement("min", Val(N(min))));
            var element = new XElement(name, new XElement("axId", Val(id)), scaling, new XElement("delete", Val(axis is null || axis.Deleted ? "1" : "0")),
                                       new XElement("axPos", Val(position)));
            if (axis?.Gridlines == true) element.Add(new XElement("majorGridlines"));
            if (axis?.Title is { } axisTitle) element.Add(new XElement("title", RichText(axisTitle, axis.TitleSize)));
            element.Add(new XElement("majorTickMark", Val(axis?.NoTickMarks == false ? "out" : "none")));
            if (axis?.LabelSize is { } labelSize) element.Add(TextProperties(labelSize));
            element.Add(new XElement("crossAx", Val(cross)));
            if (axis?.Major is { } major) element.Add(new XElement("majorUnit", Val(N(major))));
            return element;
        }
    }

    /// <summary>fillColor (0x0181) and lineColor (0x01C0) of an OfficeArtFOPT ([MS-ODRAW] 2.2.9; LibreOffice reads the record through its DFF importer, xichart.cxx:540-560); RGB only when the colour is not a scheme / system index.</summary>
    private static (string? Fill, string? Line) EscherColors(byte[] body)
    {
        if (body.Length < 8 || BitConverter.ToUInt16(body, 2) != 0xF00B) return (null, null);
        int count = BitConverter.ToUInt16(body, 0) >> 4;
        string? fill = null, line = null;
        for (int i = 0; i < count && 8 + i * 6 + 6 <= body.Length; i++)
        {
            int at = 8 + i * 6;
            int property = BitConverter.ToUInt16(body, at) & 0x3FFF;
            uint value = BitConverter.ToUInt32(body, at + 2);
            if ((value >> 24 & 0x1D) != 0) continue;           // palette, system, scheme or system-index colour: not an RGB
            string color = ((byte)value).ToString("X2") + ((byte)(value >> 8)).ToString("X2") + ((byte)(value >> 16)).ToString("X2");
            if (property == 0x0181) fill = color;
            else if (property == 0x01C0) line = color;
        }
        return (fill, line);
    }

    /// <summary>An XLUnicodeString-like run ([MS-XLS] 2.5.294): a count of <paramref name="countBytes"/> bytes, a flags byte (bit 0: UTF-16), the characters.</summary>
    private static string UnicodeString(byte[] body, int at, int countBytes)
    {
        if (at + countBytes + 1 > body.Length) return "";
        int count = countBytes == 1 ? body[at] : BitConverter.ToUInt16(body, at);
        at += countBytes;
        bool wide = (body[at++] & 1) != 0;
        count = Math.Min(count, (body.Length - at) / (wide ? 2 : 1));
        return wide ? Encoding.Unicode.GetString(body, at, count * 2) : Latin1(body, at, count);
    }

    private static string Latin1(byte[] body, int at, int count)
    {
        var text = new char[count];
        for (int i = 0; i < count; i++) text[i] = (char)body[at + i];
        return new string(text);
    }

    private static string Rgb(byte[] body, int at) => body[at].ToString("X2") + body[at + 1].ToString("X2") + body[at + 2].ToString("X2");

    private static bool ReadFully(Stream stream, byte[] buffer, int count)
    {
        int read = 0;
        while (read < count)
        {
            int n = stream.Read(buffer, read, count - read);
            if (n <= 0) return false;
            read += n;
        }
        return true;
    }

    private static XElement RichText(string text, double? size)
        => new("tx", new XElement("rich", text.Split('\n').Select(line =>
            new XElement("p", new XElement("r", size is { } s ? new XElement("rPr", new XAttribute("sz", N(s * 100))) : null, new XElement("t", line))))));

    private static XElement TextProperties(double size) => new("txPr", new XElement("p", new XElement("pPr", new XElement("defRPr", new XAttribute("sz", N(size * 100))))));

    private static XElement Solid(string color) => new("solidFill", new XElement("srgbClr", new XAttribute("val", color)));

    private static XAttribute Val(int value) => new("val", value.ToString(CultureInfo.InvariantCulture));

    private static XAttribute Val(string value) => new("val", value);

    private static string N(double value) => value.ToString("0.####", CultureInfo.InvariantCulture);
}
