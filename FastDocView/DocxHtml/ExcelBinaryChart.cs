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
        ChRadarArea = 0x1040, ChAxesSet = 0x1041, ChSourceLink = 0x1051, ChEscherFormat = 0x1066, ChSeriesIndex = 0x1065,
        ChLabelRange = 0x1020, ChSeriesParent = 0x104A, ChSeriesTrendLine = 0x104B, ChSeriesErrorBar = 0x105B, ChFormat = 0x104E, ChFramePos = 0x104F, Format = 0x041E, PaletteRecord = 0x0092;

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
        public Dictionary<int, ushort> PointLabels = new();       // point -> CHTEXT flags of its own label
        public int Parent;                                         // CHSERPARENT (1-based): a trend line / error bar series of that series
        public byte[]? TrendLine, ErrorBar;
        public List<double?>? CustomValues;                        // a custom error bar series' cached amounts
    }

    private sealed class AxisData
    {
        public int Type;
        public double? Min, Max, Major;
        public bool Gridlines, NoTickMarks, NoLabels;
        public double? LabelSize;
        public string? Title;
        public double? TitleSize;
        public int FormatIndex = -1;                               // CHFORMAT: the number format of the tick labels
        public string? LabelColor;
        public bool Between = true;
        /// <summary>Where the other axis crosses this one: at its maximum, or at a value (CHVALUERANGE cross / CHLABELRANGE category); null = automatic.</summary>
        public bool MaxCross;
        public double? CrossValue;
    }

    private sealed class TextData
    {
        public int Link = -1, Series = -1;
        public string? Text;
        public double? Size;
        public string? Typeface;
        public ushort Flags;
        public string? Color;                                      // CHTEXT RGB when not automatic (flags bit 0)
        public int Point = 0xFFFF;
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
        TextData? legendText = null;
        (double X, double Y, double W, double H)? plotRect = null;     // the axes set's CHFRAMEPOS (chart units / 4000)
        var formats = new Dictionary<int, string>();
        var palette = new List<string>();                          // PALETTE (0x0092): the workbook's colours 8.. (xistyle.cxx:169-190)
        string? language = null;                                  // the first lang="..." of the Office 2007+ DrawingML kept in CrtMlFrt records
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
            if (id == PaletteRecord && length >= 2 && substreamType == 0x0005)
            {
                int count = Math.Min(BitConverter.ToUInt16(body, 0), (length - 2) / 4);
                palette.Clear();
                for (int i = 0; i < count; i++) palette.Add(Rgb(body, 2 + i * 4));
                continue;
            }
            if (id == Format && length >= 5 && substreamType == 0x0005)
            {
                formats[BitConverter.ToUInt16(body, 0)] = UnicodeString(body, 2, 2);
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
                    if (current is TextData pointText && pointText.Link == 4 && pointText.Point != 0xFFFF && pointText.Series >= 0 && pointText.Series < series.Count)
                        series[pointText.Series].PointLabels[pointText.Point] = pointText.Flags;
                    else if (current is TextData closed && closed.Link == 4 && closed.Series >= 0 && closed.Series < series.Count)
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
                    string color = ColorOf(body, 0, length >= 14 ? BitConverter.ToUInt16(body, 12) : -1);
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
                        else if (!automatic) formatSeries.Line = ColorOf(body, 0, length >= 12 ? BitConverter.ToUInt16(body, 10) : -1);
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
                    axisLine = -1;
                    var axis = new AxisData { Type = BitConverter.ToUInt16(body, 0) };
                    axes.Add(axis);
                    last = axis;
                    continue;
                case ChValueRange when length >= 42 && current is AxisData ranged:
                {
                    ushort flags = BitConverter.ToUInt16(body, 40);
                    if ((flags & 1) == 0) ranged.Min = Sane(BitConverter.ToDouble(body, 0));
                    if ((flags & 2) == 0) ranged.Max = Sane(BitConverter.ToDouble(body, 8));
                    if ((flags & 4) == 0 && Sane(BitConverter.ToDouble(body, 16)) is > 0 and var major) ranged.Major = major;
                    if (ranged.Min is { } low && ranged.Max is { } high && high <= low) ranged.Min = ranged.Max = null;   // a damaged range: automatic
                    // where the category axis crosses (xichart.cxx:3092-3099 cross at 32, :3150-3161: EXC_CHVALUERANGE_MAXCROSS 0x0080
                    // overrides, EXC_CHVALUERANGE_AUTOCROSS 0x0010 = zero, a log scale stores the exponent)
                    ranged.MaxCross = (flags & 0x80) != 0;
                    if (!ranged.MaxCross && (flags & 0x10) == 0 && Sane(BitConverter.ToDouble(body, 32)) is { } cross) ranged.CrossValue = (flags & 0x20) != 0 ? Math.Pow(10, cross) : cross;
                    continue;
                }
                case ChTick when length >= 1 && current is AxisData ticked:
                    ticked.NoTickMarks = body[0] is 0 or 1;            // none / inside: no mark outside the plot
                    // the labels' colour unless automatic: rgb at 4 + 16, flags at 24 (EXC_CHTICK_AUTOCOLOR), BIFF8 palette index at 26 (xichart.cxx:3194-3209)
                    if (length >= 28 && (BitConverter.ToUInt16(body, 24) & 1) == 0) ticked.LabelColor = ColorOf(body, 20, BitConverter.ToUInt16(body, 26));
                    if (length >= 3 && body[2] == 0) ticked.NoLabels = true;   // label position none: the axis line stays
                    continue;
                case ChAxisLine when length >= 2:
                    axisLine = BitConverter.ToUInt16(body, 0);
                    continue;
                case ChText when length >= 26:
                    var textRecord = new TextData { Flags = BitConverter.ToUInt16(body, 24) };
                    // EXC_CHTEXT_AUTOCOLOR clear: the record's own colour, the BIFF8 palette index at 26 (xichart.cxx:951-968)
                    if ((textRecord.Flags & 1) == 0) textRecord.Color = ColorOf(body, 4, length >= 28 ? BitConverter.ToUInt16(body, 26) : -1);
                    texts.Add(textRecord);
                    last = textRecord;
                    continue;
                case ChObjectLink when length >= 4 && current is TextData linked:
                    linked.Link = BitConverter.ToUInt16(body, 0);
                    linked.Series = BitConverter.ToUInt16(body, 2);
                    if (length >= 6) linked.Point = BitConverter.ToUInt16(body, 4);   // wLinkVar2: one point's label; 0xFFFF the whole series
                    continue;
                case ChFont when length >= 2:
                {
                    int index = BitConverter.ToUInt16(body, 0);
                    int at = index >= 4 ? index - 1 : index;             // BIFF font tables skip index 4
                    if (at < 0 || at >= fonts.Count) continue;
                    double size = fonts[at].Size;
                    if (current is TextData sized) { sized.Size = size; sized.Typeface = fonts[at].Name; if (inLegend) { legendSize = size; legendText = sized; } }
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
                    if (row >= MaxPoints || column > 255 || Sane(value) is null) continue;   // a damaged cache record
                    if (cacheKind == 1) values[(row, column)] = value;
                    // categories / x values: only a drawn series' column - a trend line or error bar series has its own column in the
                    // cache too (Word's v3: the custom x error amounts), which is not the chart's categories
                    else if (cacheKind == 2 && IsDrawnSeriesColumn(column)) { if (!categories.ContainsKey(row)) categories[row] = value.ToString("R", CultureInfo.InvariantCulture); if (!xNumbers.ContainsKey(row)) xNumbers[row] = value; }
                    else if (cacheKind == 3) bubbles[(row, column)] = value;
                    continue;
                }
                case Label when length >= 9 && chartDone && cacheKind == 2:
                {
                    int row = BitConverter.ToUInt16(body, 0);
                    if (row >= MaxPoints || !IsDrawnSeriesColumn(BitConverter.ToUInt16(body, 2))) continue;
                    if (!categories.ContainsKey(row)) categories[row] = UnicodeString(body, 6, 2);
                    continue;
                }
                case ChSeriesParent when length >= 2 && current is SeriesData child:
                    child.Parent = BitConverter.ToUInt16(body, 0);
                    continue;
                case ChSeriesTrendLine when length >= 28 && current is SeriesData trend:
                    trend.TrendLine = body;
                    continue;
                case ChSeriesErrorBar when length >= 12 && current is SeriesData bars:
                    bars.ErrorBar = body;
                    continue;
                // CHLABELRANGE (xichart.cxx:2992-2998: flags at 6; EXC_CHLABELRANGE_BETWEEN 0x0001, xlchart.hxx):
                // cleared = the points sit on the category ticks (c:crossBetween midCat)
                case ChLabelRange when length >= 8 && current is AxisData ranged0:
                    ranged0.Between = (BitConverter.ToUInt16(body, 6) & 1) != 0;
                    // where the value axis crosses: the category number at 0, EXC_CHLABELRANGE_MAXCROSS 0x0002 overrides (xichart.cxx:2992-2998,
                    // 3058-3083; Excel does not move it in 3-D charts, so it is not written for them)
                    ranged0.MaxCross = (BitConverter.ToUInt16(body, 6) & 2) != 0;
                    if (!ranged0.MaxCross && BitConverter.ToUInt16(body, 0) is var category and > 1) ranged0.CrossValue = category;
                    continue;
                case ChFormat when length >= 2 && current is AxisData formatted:
                    formatted.FormatIndex = BitConverter.ToUInt16(body, 0);
                    continue;
                case ChFramePos when length >= 20:
                {
                    int topLeft = BitConverter.ToUInt16(body, 0), bottomRight = BitConverter.ToUInt16(body, 2);
                    double x = BitConverter.ToInt32(body, 4) / 4000.0, y = BitConverter.ToInt32(body, 8) / 4000.0;
                    double w = BitConverter.ToInt32(body, 12) / 4000.0, h = BitConverter.ToInt32(body, 16) / 4000.0;
                    // the plot area: modes 2 / 2 = chart units (xichart.cxx:4030-4046); the legend: its top left in chart units (:2610-2620)
                    if (current is "axesset" && topLeft == 2 && bottomRight == 2 && w > 0 && h > 0 && x >= 0 && y >= 0 && x + w <= 1.001 && y + h <= 1.001) plotRect = (x, y, w, h);
                    else if (current is "legend" && x >= 0 && y >= 0 && x <= 1 && y <= 1) legendAt = (x, y);
                    continue;
                }
                case ChAxesSet:
                    last = "axesset";
                    continue;
                case 0x08A5 when language is null:
                {
                    // CrtMlFrt ([MS-XLS] 2.4.70; continued in 0x08A4 records): Office's DrawingML text properties,
                    // whose a:endParaRPr lang is the chart's editing language (c:lang in the .docx Word writes from the same chart)
                    string xml = Latin1(body, 0, body.Length);
                    int at = xml.IndexOf("lang=\"", StringComparison.Ordinal);
                    if (at >= 0 && xml.IndexOf('"', at + 6) is int close && close > at + 6 && close - at < 20
                        && System.Text.RegularExpressions.Regex.IsMatch(xml.Substring(at + 6, close - at - 6), "^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$"))
                        language = xml.Substring(at + 6, close - at - 6);   // a BCP 47 tag only (damaged bytes are dropped)
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
        // custom error amounts: the error bar series' own cached values (their column of the value cache)
        for (int s = 0; s < series.Count; s++)
            if (series[s].ErrorBar is { } custom && custom[1] == 4)
                series[s].CustomValues = Enumerable.Range(0, MaxPointOf(values)).Select(p => values.TryGetValue((p, s), out double v) ? (double?)v : null).ToList();

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
                if (data.Parent > 0) continue;                           // a trend line / error bar series: drawn with its parent
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
                if (data.ShowValue || data.ShowPercent || data.ShowCategory || data.PointLabels.Count > 0)
                {
                    var labels = new XElement("dLbls");
                    foreach (KeyValuePair<int, ushort> own in data.PointLabels.OrderBy(p => p.Key))
                        labels.Add(new XElement("dLbl", new XElement("idx", Val(own.Key)), LabelFlags(own.Value)));
                    labels.Add(LabelFlags((ushort)((data.ShowValue ? 0x0004 : 0) | (data.ShowPercent ? 0x1000 : 0) | (data.ShowCategory ? 0x4000 : 0))));
                    element.Add(labels);
                }
                int parentNumber = s + 1;
                foreach (SeriesData child in series.Where(c => c.Parent == parentNumber && c.TrendLine is not null)) element.Add(TrendLineElement(child.TrendLine!));
                foreach (XElement errorBars in ErrorBarsElements(series.Where(c => c.Parent == parentNumber && c.ErrorBar is not null).ToList())) element.Add(errorBars);
                var cells = new XElement("numCache", new XElement("ptCount", Val(pointCount)));
                for (int p = 0; p < pointCount; p++)
                    if (values.TryGetValue((p, s), out double v)) cells.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", v.ToString("R", CultureInfo.InvariantCulture))));
                if (kind is "scatter" or "bubble")
                {
                    var x = new XElement("numCache", new XElement("ptCount", Val(pointCount)));
                    // text x values (no numbers in the category cache): none, so the points are 1..n as Word shows them
                    for (int p = 0; p < pointCount && xNumbers.Count > 0; p++)
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

            var layout = new XElement("layout");
            if (plotRect is { } rect && UsesSavedPlotPosition)
                // pies: the inner plot area; else with the axes (LibreOffice xichart.cxx:4036-4046)
                layout.Add(new XElement("manualLayout", new XElement("layoutTarget", Val(radial ? "inner" : "outer")),
                    new XElement("xMode", Val("edge")), new XElement("yMode", Val("edge")),
                    new XElement("x", Val(N(rect.X))), new XElement("y", Val(N(rect.Y))), new XElement("w", Val(N(rect.W))), new XElement("h", Val(N(rect.H)))));
            var plot = new XElement("plotArea", layout, group);
            if (!radial)
            {
                group.Add(new XElement("axId", Val(1)), new XElement("axId", Val(2)));
                AxisData? x = axes.FirstOrDefault(a => a.Type == 0), y = axes.FirstOrDefault(a => a.Type == 1);
                plot.Add(Axis(x, kind is "scatter" or "bubble" ? "valAx" : "catAx", 1, 2, horizontal ? "l" : "b"));
                plot.Add(Axis(y, "valAx", 2, 1, horizontal ? "b" : "l"));
            }

            var chart = new XElement("chart");
            if (title?.Text is { Length: > 0 } heading)
                chart.Add(new XElement("title", RichText(heading, title.Size, title.Color)));
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
                if (legendAt is { } at && UsesSavedPlotPosition)
                    legend.Add(new XElement("layout", new XElement("manualLayout", new XElement("xMode", Val("edge")), new XElement("yMode", Val("edge")),
                        new XElement("x", Val(N(at.X))), new XElement("y", Val(N(at.Y))))));
                if (legendSize is { } size) legend.Add(TextProperties(size, legendText?.Color));
                chart.Add(legend);
            }
            var space = new XElement("chartSpace", chart);
            if (language is not null) space.AddFirst(new XElement("lang", Val(language)));
            if (typeface is { Length: > 0 })
                space.Add(new XElement("txPr", new XElement("p", new XElement("pPr", new XElement("defRPr", new XElement("latin", new XAttribute("typeface", typeface)))))));
            return space;
        }

        // a BIFF8 palette index (xistyle.cxx:158-167, xlstyle.cxx:78-88 default table, xlstyle.hxx:205-216 system indices):
        // 8.. from the PALETTE record, else the default table; window text / chart text and auto = null (automatic)
        // a record's colour: BIFF8 keeps both an RGB and a palette index. LibreOffice reads only the index (xichart.cxx:453-499,
        // 951-968, 3194-3209), which Office 2007+ writes as the nearest of the 56 palette colours (measured: legend 00B050 has
        // index 0x39 = 339966) while the RGB is the exact colour Word draws. So the RGB, unless it is black and the index is not
        // (writers that leave the RGB empty) - then the palette colour as LibreOffice reads it.
        string ColorOf(byte[] record, int rgbAt, int index)
        {
            string rgb = Rgb(record, rgbAt);
            return rgb == "000000" && index >= 0 && PaletteColor(index) is { } fromPalette ? fromPalette : rgb;
        }

        string? PaletteColor(int index)
            => index >= 8 && index - 8 < palette.Count ? palette[index - 8] : index >= 0 && index < DefaultPalette.Length ? DefaultPalette[index] : index is 65 or 78 ? "FFFFFF" : null;

        bool IsDrawnSeriesColumn(int column) => column >= series.Count || series[column].Parent == 0;

        string? NumberFormat(int index) => formats.TryGetValue(index, out string? custom) ? custom : BuiltInFormat(index);

        XElement Axis(AxisData? axis, string name, int id, int cross, string position)
        {
            var scaling = new XElement("scaling");
            if (axis?.Max is { } max) scaling.Add(new XElement("max", Val(N(max))));
            if (axis?.Min is { } min) scaling.Add(new XElement("min", Val(N(min))));
            var element = new XElement(name, new XElement("axId", Val(id)), scaling, new XElement("delete", Val(axis is null ? "1" : "0")),
                                       new XElement("axPos", Val(position)));
            if (axis?.Gridlines == true) element.Add(new XElement("majorGridlines"));
            if (axis?.Title is { } axisTitle) element.Add(new XElement("title", RichText(axisTitle, axis.TitleSize)));
            element.Add(new XElement("majorTickMark", Val(axis?.NoTickMarks == false ? "out" : "none")));
            if (axis?.NoLabels == true) element.Add(new XElement("tickLblPos", Val("none")));
            if (axis is { FormatIndex: >= 0 } && NumberFormat(axis.FormatIndex) is { } code) element.Add(new XElement("numFmt", new XAttribute("formatCode", code), new XAttribute("sourceLinked", "0")));
            if (axis?.LabelSize is { } labelSize) element.Add(TextProperties(labelSize, axis.LabelColor));
            element.Add(new XElement("crossAx", Val(cross)));
            // c:crosses / c:crossesAt come from the crossing axis's record: a category axis crosses where the value axis's CHVALUERANGE
            // says, a value axis where the category axis's CHLABELRANGE says
            if (axes.FirstOrDefault(a => a.Type == (name == "catAx" ? 1 : 0)) is { } crossing && (name == "catAx" || axis?.Type == 1 && !is3D))
            {
                if (crossing.MaxCross) element.Add(new XElement("crosses", Val("max")));
                else if (crossing.CrossValue is { } at) element.Add(new XElement("crossesAt", Val(N(at))));
            }
            if (name == "valAx" && axes.FirstOrDefault(a => a.Type == 0) is { Between: false }) element.Add(new XElement("crossBetween", Val("midCat")));
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

    /// <summary>Excel's limit of points per series in a 2-D chart (32,000; Excel specifications and limits): a row past it is damage.</summary>
    private const int MaxPoints = 32000;

    /// <summary>Whether the saved plot rectangle (CHFRAMEPOS of the axes set) and legend position are applied: off - measured against Word's PDF (21 texts) they moved the texts further from Word's (median 8.5 / 7.5 pt with the outer rectangle, 16.9 / 19.6 pt with the inner one, 7.6 / 20.3 pt with the legend position alone) than the automatic layout does.</summary>
    private const bool UsesSavedPlotPosition = false;

    /// <summary>A finite number; null for NaN, infinity or a magnitude past 1e300 (damaged bytes).</summary>
    private static double? Sane(double value) => double.IsNaN(value) || double.IsInfinity(value) || Math.Abs(value) > 1e300 ? null : value;

    /// <summary>BIFF8's default colour table (LibreOffice xlstyle.cxx:44-48, 78-88): 0-7 and 8-15 the light EGA colours, 16-23 the dark ones, then the rest.</summary>
    private static readonly string[] DefaultPalette =
    [
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
        "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
        "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
        "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
        "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
        "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
        "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
    ];

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

    private static XElement RichText(string text, double? size, string? color = null)
        => new("tx", new XElement("rich", text.Split('\n').Select(line =>
            new XElement("p", new XElement("r", size is not null || color is not null
                ? new XElement("rPr", size is { } s ? new XAttribute("sz", N(s * 100)) : null, color is not null ? Solid(color) : null) : null, new XElement("t", line))))));

    /// <summary>A data label's show flags (CHTEXT, xlchart.hxx:474-483) as DrawingML c:showVal / c:showCatName / c:showPercent.</summary>
    private static XElement[] LabelFlags(ushort flags)
        => [new XElement("showVal", Val((flags & 0x0004) != 0 ? "1" : "0")), new XElement("showCatName", Val((flags & 0x4000) != 0 ? "1" : "0")),
            new XElement("showPercent", Val((flags & 0x1000) != 0 ? "1" : "0"))];

    /// <summary>
    /// CHSERTRENDLINE ([MS-XLS] 2.4.270 SerAuxTrend; LibreOffice xichart.cxx XclImpChSerTrendLine): type (0 polynomial, 1 exponential,
    /// 2 logarithmic, 3 power, 4 moving average), order / period, intercept (NaN = none), equation and R squared flags, forward and
    /// backward forecast - as c:trendline.
    /// </summary>
    private static XElement TrendLineElement(byte[] body)
    {
        int type = body[0], order = body[1];
        double intercept = BitConverter.ToDouble(body, 2);
        bool equation = body[10] != 0, rSquared = body[11] != 0;
        double forward = Sane(BitConverter.ToDouble(body, 12)) ?? 0, backward = Sane(BitConverter.ToDouble(body, 20)) ?? 0;
        string kindName = type switch { 0 => order <= 1 ? "linear" : "poly", 1 => "exp", 2 => "log", 3 => "power", 4 => "movingAvg", _ => "linear" };
        var element = new XElement("trendline", new XElement("trendlineType", Val(kindName)));
        if (kindName == "poly") element.Add(new XElement("order", Val(Math.Max(2, Math.Min(6, order)))));
        if (kindName == "movingAvg") element.Add(new XElement("period", Val(Math.Max(2, order))));
        if (forward > 0) element.Add(new XElement("forward", Val(N(forward))));
        if (backward > 0) element.Add(new XElement("backward", Val(N(backward))));
        if (Sane(intercept) is { } fixedIntercept) element.Add(new XElement("intercept", Val(N(fixedIntercept))));
        element.Add(new XElement("dispRSqr", Val(rSquared ? "1" : "0")), new XElement("dispEq", Val(equation ? "1" : "0")));
        return element;
    }

    /// <summary>
    /// The CHSERERRORBAR series of one parent ([MS-XLS] 2.4.271 SerAuxErrBar; LibreOffice xichart.cxx:1696-1704: direction 1 x+,
    /// 2 x-, 3 y+, 4 y-; source 1 percent, 2 fixed, 3 standard deviation, 4 custom, 5 standard error; end caps; the value): one
    /// record per direction, joined per axis as c:errBars - "both" when plus and minus have the same source and value, else one
    /// c:errBars for each. Custom amounts are the error series' own cached values (c:plus / c:minus).
    /// </summary>
    private static IEnumerable<XElement> ErrorBarsElements(List<SeriesData> bars)
    {
        foreach (bool yDirection in new[] { true, false })
        {
            var same = bars.Where(b => b.ErrorBar![0] is >= 1 and <= 4 && b.ErrorBar[1] is >= 1 and <= 5 && (b.ErrorBar[0] >= 3) == yDirection).ToList();
            SeriesData? plus = same.FirstOrDefault(b => b.ErrorBar![0] is 1 or 3), minus = same.FirstOrDefault(b => b.ErrorBar![0] is 2 or 4);
            bool joined = plus is not null && minus is not null && plus.ErrorBar![1] == minus.ErrorBar![1] && plus.ErrorBar[1] != 4
                          && BitConverter.ToDouble(plus.ErrorBar, 4).Equals(BitConverter.ToDouble(minus.ErrorBar, 4));
            if (joined) { yield return ErrorBarsOf(plus!, "both", yDirection); continue; }
            if (plus is not null) yield return ErrorBarsOf(plus, "plus", yDirection);
            if (minus is not null) yield return ErrorBarsOf(minus, "minus", yDirection);
        }
    }

    private static XElement ErrorBarsOf(SeriesData bar, string type, bool yDirection)
    {
        byte[] record = bar.ErrorBar!;
        string valueType = record[1] switch { 1 => "percentage", 3 => "stdDev", 4 => "cust", 5 => "stdErr", _ => "fixedVal" };
        if (valueType == "cust" && bar.CustomValues?.Any(v => v is not null) != true) valueType = "fixedVal";   // no cached amounts: a zero bar, not an empty list
        var element = new XElement("errBars", new XElement("errDir", Val(yDirection ? "y" : "x")), new XElement("errBarType", Val(type)),
            new XElement("errValType", Val(valueType)), new XElement("noEndCap", Val(record[2] != 0 ? "0" : "1")));
        if (valueType == "cust")
        {
            var cache = new XElement("numCache", new XElement("ptCount", Val(bar.CustomValues?.Count ?? 0)));
            for (int p = 0; p < (bar.CustomValues?.Count ?? 0); p++)
                if (bar.CustomValues![p] is { } amount) cache.Add(new XElement("pt", new XAttribute("idx", p), new XElement("v", amount.ToString("R", CultureInfo.InvariantCulture))));
            element.Add(new XElement(type == "minus" ? "minus" : "plus", new XElement("numRef", cache)));
            if (type == "both") element.Add(new XElement("minus", new XElement("numRef", new XElement(cache))));
        }
        else element.Add(new XElement("val", Val(N(Sane(BitConverter.ToDouble(record, 4)) ?? 0))));
        return element;
    }

    private static int MaxPointOf(Dictionary<(int Point, int Series), double> values) => values.Count == 0 ? 0 : values.Keys.Max(k => k.Point) + 1;


    /// <summary>Excel's built-in number formats seen on chart axes ([MS-XLS] 2.4.126 Format: the built-in ids are not stored); null when unknown.</summary>
    private static string? BuiltInFormat(int index) => index switch
    {
        0 => "General", 1 => "0", 2 => "0.00", 3 => "#,##0", 4 => "#,##0.00", 9 => "0%", 10 => "0.00%", 11 => "0.00E+00", _ => null,
    };

    private static XElement TextProperties(double size, string? color = null)
        => new("txPr", new XElement("p", new XElement("pPr", new XElement("defRPr", new XAttribute("sz", N(size * 100)), color is not null ? Solid(color) : null))));

    private static XElement Solid(string color) => new("solidFill", new XElement("srgbClr", new XAttribute("val", color)));

    private static XAttribute Val(int value) => new("val", value.ToString(CultureInfo.InvariantCulture));

    private static XAttribute Val(string value) => new("val", value);

    private static string N(double value) => value.ToString("0.####", CultureInfo.InvariantCulture);
}
