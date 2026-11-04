namespace FastDocView.Html;

/// <summary>What a custom-shape parameter is (com.sun.star.drawing.EnhancedCustomShapeParameterType).</summary>
internal enum WordShapeParameterKind : byte { Normal, Equation, Adjustment, Left, Top, Right, Bottom, Other }

/// <summary>One custom-shape parameter: a number (an integer one can be replaced by the stretch ratio), an equation or
/// adjustment index, or a keyword (svx EnhancedCustomShape2d::GetParameter).</summary>
internal readonly record struct WordShapeParameter(WordShapeParameterKind Kind, double Value, bool IsInteger)
{
    public static WordShapeParameter Integer(int value) => new(WordShapeParameterKind.Normal, value, true);
    public static WordShapeParameter Real(double value) => new(WordShapeParameterKind.Normal, value, false);
    public static WordShapeParameter EquationResult(int index) => new(WordShapeParameterKind.Equation, index, true);

    /// <summary>A preset table value: bit 30 set (EnhancedCustomShapeGeometry.hxx:88 MSO_I) = equation n (SetEnhancedCustomShapeParameter, EnhancedCustomShape2d.cxx:72-86).</summary>
    public static WordShapeParameter FromTable(int value) => (value & 0x40000000) != 0 && value > 0 ? EquationResult(value & 0xFFFF) : Integer(value);
}

/// <summary>com.sun.star.drawing.EnhancedCustomShapeSegmentCommand (ODF draw:enhanced-path letters M L C Z N F S T U A B W V X Y Q G H I J K).</summary>
internal enum WordPathCommand : byte
{
    MoveTo, LineTo, CurveTo, CloseSubpath, EndSubpath, NoFill, NoStroke, AngleEllipseTo, AngleEllipse, ArcTo, Arc, ClockwiseArcTo,
    ClockwiseArc, EllipticalQuadrantX, EllipticalQuadrantY, QuadraticCurveTo, ArcAngleTo, Darken, DarkenLess, Lighten, LightenLess, Unknown,
}

internal readonly record struct WordPathSegment(WordPathCommand Command, int Count);

/// <summary>How the angle-ellipse commands U / T read their angles (EnhancedCustomShape2d.cxx:2185-2343).</summary>
internal enum WordAngleEllipseMode : byte { Odf, Full, Binary }

/// <summary>
/// A custom shape's outline as LibreOffice computes it (svx/source/customshapes/EnhancedCustomShape2d.cxx, master read
/// 2026-11-04), shared by every format that carries LibreOffice's custom-shape model (OfficeArt presets and geometry,
/// ODF draw:enhanced-geometry):
/// - SetPathSize :650-732 - the coordinate space of each sub-path (sub-view sizes), the scale to the logic rectangle, and the
///   stretch ratios when path-stretchpoint-x / -y (nXRef / nYRef) are set;
/// - GetParameter :1011-1080 - an integer equal to the coordinate width / height is multiplied by the ratio; left / top are 0,
///   right / bottom the coordinate width / height;
/// - GetEnumFunc :874-895 - the formula keywords;
/// - CreateSubPath :2078-2698 - every segment command; checkClosed when a sub-path ends (#i76201#), the fill / stroke objects
///   (one object, or a fill and a stroke object when the shape sorts filled objects back or an open sub-path is stroked);
/// - CreatePathObj :2888-3040 - invisible objects dropped, the colour index (AdaptObjColor :2800-2886: per filled object)
///   and the filled objects sorted in front of the lines.
/// Works in hundredths of a millimetre like LibreOffice (logwidth / logheight are in that unit); the result is in points.
/// Ellipses and arcs are flattened to short lines (basegfx draws them as Bezier segments; the difference is sub-pixel).
/// </summary>
internal sealed class WordCustomShapeEngine
{
    public const double HundredthsPerPoint = 2540.0 / 72.0;

    public WordShapeParameter[] Coordinates = [];
    public WordPathSegment[] Segments = [];
    public (int Width, int Height)[] SubViewSizes = [];
    public double[] Adjust = [];
    /// <summary>Equation n's value; set by the owner (evaluated lazily against this engine's state, cached by the owner).</summary>
    public Func<int, double> Equation = _ => 0;
    public int CoordLeft, CoordTop, CoordWidthG = 21600, CoordHeightG = 21600;
    /// <summary>path-stretchpoint-x / -y (StretchX / StretchY); null = unset (0x80000000).</summary>
    public int? StretchX, StretchY;
    public double LogicWidth, LogicHeight;
    public bool Filled = true, Stroked = true, IsOoxml;
    public WordAngleEllipseMode EllipseMode;
    /// <summary>"mso-spt143" (WordArt RingOutside): its binary angle ellipses are in degrees already (:2257-2260).</summary>
    public bool AnglesInDegrees;
    /// <summary>SortFilledObjectsToBackByDefault (EnhancedCustomShapeGeometry.cxx:6215-6226).</summary>
    public bool SortFilledToBack;
    /// <summary>m_nColorData (EnhancedCustomShape2d.cxx:780-858): luminance changes per filled object, used unless <see cref="UsesBrightness"/>.</summary>
    public uint ColorData;
    /// <summary>OOXML shapes and ODF "non-primitive" use the path brightness (H / I / J / K) instead of the colour data (GetColorData :1095-1150).</summary>
    public bool UsesBrightness;
    public bool FlipHorizontal, FlipVertical;
    public WordShapeParameter[] TextFrames = [];

    private int _coordWidth, _coordHeight;
    private double _xScale, _yScale, _xRatio = 1, _yRatio = 1;

    public static WordAngleEllipseMode EllipseModeOf(string? typeName)
        => typeName is not null && typeName.StartsWith("mso", StringComparison.Ordinal) ? WordAngleEllipseMode.Binary
         : typeName is "ellipse" or "ring" or "smiley" or "sun" or "forbidden" or "flowchart-connector" or "flowchart-summing-junction" or "flowchart-or" or "cloud-callout"
             ? WordAngleEllipseMode.Full : WordAngleEllipseMode.Odf;

    /// <summary>SetPathSize (EnhancedCustomShape2d.cxx:650-732).</summary>
    public void SetPathSize(int index)
    {
        int width = 0, height = 0;
        if (index < SubViewSizes.Length) (width, height) = SubViewSizes[index];
        if (width != 0 && height != 0) { _coordWidth = width; _coordHeight = height; }
        else { _coordWidth = CoordWidthG; _coordHeight = CoordHeightG; }
        _xScale = _coordWidth == 0 ? 0 : LogicWidth / _coordWidth;
        _yScale = _coordHeight == 0 ? 0 : LogicHeight / _coordHeight;
        if (IsOoxml)
        {
            if (_coordWidth == 0) _xScale = width != 0 ? LogicWidth / width : 1;
            if (_coordHeight == 0) _yScale = height != 0 ? LogicHeight / height : 1;
        }
        _xRatio = 1;
        if (StretchX is not null && LogicHeight != 0)
        {
            _xRatio = LogicWidth / LogicHeight;
            if (_xRatio > 1) _xScale /= _xRatio; else _xRatio = 1;
        }
        _yRatio = 1;
        if (StretchY is not null && LogicWidth != 0)
        {
            _yRatio = LogicHeight / LogicWidth;
            if (_yRatio > 1) _yScale /= _yRatio; else _yRatio = 1;
        }
    }

    /// <summary>The formula keywords (GetEnumFunc :874-895).</summary>
    public double Keyword(string name) => name switch
    {
        "pi" => Math.PI,
        "left" => CoordLeft,
        "top" => CoordTop,
        "right" => ((double)CoordLeft + _coordWidth) * _xRatio,
        "bottom" => ((double)CoordTop + _coordHeight) * _yRatio,
        "xstretch" => StretchX ?? int.MinValue,
        "ystretch" => StretchY ?? int.MinValue,
        "hasstroke" => Stroked ? 1 : 0,
        "hasfill" => Filled ? 1 : 0,
        "width" => _coordWidth,
        "height" => _coordHeight,
        "logwidth" => LogicWidth,
        "logheight" => LogicHeight,
        _ => 0,
    };

    public double AdjustValue(int index) => index >= 0 && index < Adjust.Length ? Adjust[index] : 0;

    /// <summary>GetParameter (:1011-1080).</summary>
    public double Parameter(WordShapeParameter parameter, bool replaceWidth, bool replaceHeight)
    {
        switch (parameter.Kind)
        {
            case WordShapeParameterKind.Adjustment: return AdjustValue((int)parameter.Value);
            case WordShapeParameterKind.Equation: return Equation((int)parameter.Value);
            case WordShapeParameterKind.Normal:
            {
                double value = parameter.Value;
                if (parameter.IsInteger)
                {
                    if (replaceWidth && value == _coordWidth) value *= _xRatio;
                    else if (replaceHeight && value == _coordHeight) value *= _yRatio;
                }
                return value;
            }
            case WordShapeParameterKind.Right: return _coordWidth;
            case WordShapeParameterKind.Bottom: return _coordHeight;
            default: return 0;
        }
    }

    /// <summary>GetPointAsB2DPoint (:983-1002), in hundredths of a millimetre from the logic rectangle's corner.</summary>
    private (double X, double Y) Point(int pair, bool scale, bool replace)
    {
        double x = Parameter(Coordinates[pair * 2], replace, false) - CoordLeft;
        double y = Parameter(Coordinates[pair * 2 + 1], false, replace) - CoordTop;
        return scale ? (x * _xScale, y * _yScale) : (x, y);
    }

    private sealed class PathObject
    {
        public readonly List<List<(double X, double Y, bool IsCurve, double C1X, double C1Y, double C2X, double C2Y)>> Polygons = new();
        public readonly List<bool> Closed = new();
        public bool HasFill, HasStroke, IsLine, IsFillCopy;
        public double Brightness;
    }

    /// <summary>The shape's drawn objects (CreatePathObj), in points from the logic rectangle's corner; null when it has no coordinates (nothing is drawn).</summary>
    public List<WordShapePath>? Build()
    {
        if (Coordinates.Length < 2) return null;
        var objects = new List<PathObject>();
        int sourcePoint = 0, segmentIndex = 0, subPathIndex = 0;
        while (segmentIndex <= Segments.Length)
        {
            CreateSubPath(ref sourcePoint, ref segmentIndex, objects, subPathIndex);
            subPathIndex++;
        }
        // invisible objects (#i37011#) stay: they draw nothing, and the text-effect layout reads every sub-path group (bLineGeometryNeededOnly)
        // colour index (AdaptObjColor): every filled, non-line object in creation order takes the next one
        int colorCount = (int)(ColorData >> 28), colorIndex = 0;
        var result = new List<(PathObject source, WordShapePath path)>();
        foreach (PathObject o in objects)
        {
            var path = new WordShapePath { Fill = o.HasFill ? WordShapePathFill.Normal : WordShapePathFill.None, Stroke = o.HasStroke, IsFillCopy = o.IsFillCopy };
            if (o.HasFill && !o.IsLine)
            {
                int index = colorCount > 0 ? Math.Min(colorIndex, colorCount - 1) : colorIndex;
                if (UsesBrightness) path.Brightness = o.Brightness;
                else path.LuminancePercent = LuminanceChange(index);
                if (colorIndex < colorCount) colorIndex++;
            }
            const double ToPoints = 1 / HundredthsPerPoint;
            for (int p = 0; p < o.Polygons.Count; p++)
            {
                var points = o.Polygons[p];
                if (points.Count == 0) continue;
                var subpath = new WordShapeSubpath { StartX = points[0].X * ToPoints, StartY = points[0].Y * ToPoints, Closed = o.Closed[p] };
                for (int i = 1; i < points.Count; i++)
                {
                    var q = points[i];
                    subpath.Segments.Add(q.IsCurve
                        ? new WordShapeSegment(true, q.C1X * ToPoints, q.C1Y * ToPoints, q.C2X * ToPoints, q.C2Y * ToPoints, q.X * ToPoints, q.Y * ToPoints)
                        : WordShapeSegment.Line(q.X * ToPoints, q.Y * ToPoints));
                }
                path.Subpaths.Add(subpath);
            }
            result.Add((o, path));
        }
        // sort objects so that filled ones are in front (:2984-3001): the non-line objects first, then the lines, each in order
        if (SortFilledToBack && result.Count > 1)
            result = result.Where(r => !r.source.IsLine).Concat(result.Where(r => r.source.IsLine)).ToList();
        return result.Select(r => r.path).ToList();
    }

    /// <summary>GetLuminanceChange (:1082-1093): 4-bit signed entries of 10 percent.</summary>
    private int LuminanceChange(int index)
    {
        uint count = ColorData >> 28;
        if (count == 0) return 0;
        if (index >= count) index = (int)count - 1;
        int data = (int)(ColorData << ((1 + index) << 2));
        return (data >> 28) * 10;
    }

    /// <summary>The first text frame (GetTextRect :1152-1176) in points from the logic rectangle's corner, mirrored with the shape; null = the whole rectangle.</summary>
    public (double X, double Y, double Width, double Height)? TextRect()
    {
        if (TextFrames.Length < 4 || Coordinates is null) return null;
        SetPathSize(0);
        bool scale = !IsOoxml;
        (double X, double Y) Corner(int i)
        {
            double x = Parameter(TextFrames[i], true, false) - CoordLeft, y = Parameter(TextFrames[i + 1], false, true) - CoordTop;
            return scale ? (Math.Truncate(x * _xScale), Math.Truncate(y * _yScale)) : (Math.Truncate(x), Math.Truncate(y));
        }
        (double X, double Y) topLeft = Corner(0), bottomRight = Corner(2);
        double left = topLeft.X, top = topLeft.Y, right = bottomRight.X, bottom = bottomRight.Y;
        if (FlipHorizontal) { left = LogicWidth - 1 - bottomRight.X; right = LogicWidth - 1 - topLeft.X; }
        if (FlipVertical) { top = LogicHeight - 1 - bottomRight.Y; bottom = LogicHeight - 1 - topLeft.Y; }
        // tools::Rectangle width / height count both edges (+1)
        if (Math.Abs(right - left) + 1 <= 1 || Math.Abs(bottom - top) + 1 <= 1) return null;
        double x0 = Math.Min(left, right), y0 = Math.Min(top, bottom);
        double w = Math.Abs(right - left), h = Math.Abs(bottom - top);
        const double ToPoints = 1 / HundredthsPerPoint;
        return (x0 * ToPoints, y0 * ToPoints, w * ToPoints, h * ToPoints);
    }

    private void CreateSubPath(ref int sourcePoint, ref int segmentIndex, List<PathObject> objects, int index)
    {
        bool noFill = false, noStroke = false;
        double brightness = 0;
        var polygons = new List<List<(double X, double Y, bool IsCurve, double C1X, double C1Y, double C2X, double C2Y)>>();
        var closed = new List<bool>();
        var polygon = new List<(double X, double Y, bool IsCurve, double C1X, double C1Y, double C2X, double C2Y)>();
        SetPathSize(index);
        int coordinateCount = Coordinates.Length / 2;

        void Append((double X, double Y) p) => polygon.Add((p.X, p.Y, false, 0, 0, 0, 0));
        void AppendAll(IEnumerable<(double X, double Y)> points) { foreach (var p in points) Append(p); }
        void FinishOpen()
        {
            if (polygon.Count > 1) { bool isClosed = CheckClosed(polygon); polygons.Add(polygon); closed.Add(isClosed); }
            polygon = new();
        }

        if (Segments.Length == 0)
        {
            for (int i = 0; i < coordinateCount; i++) Append(Point(i, true, true));
            if (polygon.Count > 0) { polygons.Add(polygon); closed.Add(true); }
            polygon = new();
        }
        else
        {
            for (; segmentIndex < Segments.Length;)
            {
                WordPathCommand command = Segments[segmentIndex].Command;
                int count = Segments[segmentIndex++].Count;
                switch (command)
                {
                    case WordPathCommand.NoFill: noFill = true; break;
                    case WordPathCommand.NoStroke: noStroke = true; break;
                    case WordPathCommand.Darken: brightness = -0.4; break;
                    case WordPathCommand.DarkenLess: brightness = -0.2; break;
                    case WordPathCommand.Lighten: brightness = 0.4; break;
                    case WordPathCommand.LightenLess: brightness = 0.2; break;
                    case WordPathCommand.MoveTo:
                        FinishOpen();
                        if (sourcePoint < coordinateCount) Append(Point(sourcePoint++, true, true));
                        break;
                    case WordPathCommand.EndSubpath: break;
                    case WordPathCommand.CloseSubpath:
                        if (polygon.Count > 0)
                        {
                            if (polygon.Count > 1) { polygons.Add(polygon); closed.Add(true); }
                            polygon = new();
                        }
                        break;
                    case WordPathCommand.CurveTo:
                        for (int i = 0; i < count && sourcePoint + 2 < coordinateCount; i++)
                        {
                            var a = Point(sourcePoint++, true, true); var b = Point(sourcePoint++, true, true); var e = Point(sourcePoint++, true, true);
                            if (polygon.Count == 0) Append(a);       // basegfx appendBezierSegment needs a start point; LibreOffice asserts here
                            polygon.Add((e.X, e.Y, true, a.X, a.Y, b.X, b.Y));
                        }
                        break;
                    case WordPathCommand.AngleEllipse:
                    case WordPathCommand.AngleEllipseTo:
                        for (int i = 0; i < count && sourcePoint + 2 < coordinateCount; i++)
                        {
                            if (command == WordPathCommand.AngleEllipse) FinishOpen();
                            var centre = Point(sourcePoint, true, true);
                            double wr = Parameter(Coordinates[(sourcePoint + 1) * 2], true, false), hr = Parameter(Coordinates[(sourcePoint + 1) * 2 + 1], false, true);
                            double startAngle = Parameter(Coordinates[(sourcePoint + 2) * 2], false, false), endAngle = Parameter(Coordinates[(sourcePoint + 2) * 2 + 1], false, false);
                            sourcePoint += 3;
                            double scaledWr = wr * _xScale, scaledHr = hr * _yScale;
                            if (scaledWr == 0 && scaledHr == 0) { Append(centre); continue; }
                            AppendAll(AngleEllipse(centre, scaledWr, scaledHr, wr, hr, startAngle, endAngle));
                        }
                        break;
                    case WordPathCommand.QuadraticCurveTo:
                        for (int i = 0; i < count && sourcePoint + 1 < coordinateCount; i++)
                        {
                            if (polygon.Count > 0)
                            {
                                var previous = polygon[polygon.Count - 1];
                                var q = Point(sourcePoint++, true, true); var end = Point(sourcePoint++, true, true);
                                double c1x = (previous.X + q.X * 2) / 3, c1y = (previous.Y + q.Y * 2) / 3, c2x = (q.X * 2 + end.X) / 3, c2y = (q.Y * 2 + end.Y) / 3;
                                polygon.Add((end.X, end.Y, true, c1x, c1y, c2x, c2y));
                            }
                            else { sourcePoint++; Append(Point(sourcePoint++, true, true)); }
                        }
                        break;
                    case WordPathCommand.LineTo:
                        for (int i = 0; i < count && sourcePoint < coordinateCount; i++) Append(Point(sourcePoint++, true, true));
                        break;
                    case WordPathCommand.Arc or WordPathCommand.ClockwiseArc or WordPathCommand.ArcTo or WordPathCommand.ClockwiseArcTo:
                    {
                        bool clockwise = command is WordPathCommand.ClockwiseArc or WordPathCommand.ClockwiseArcTo;
                        bool implicitMove = command is WordPathCommand.Arc or WordPathCommand.ClockwiseArc;
                        int xor = clockwise ? 3 : 2;
                        for (int i = 0; i < count && sourcePoint + 3 < coordinateCount; i++)
                        {
                            if (implicitMove) FinishOpen();
                            var c1 = Truncated(Point(sourcePoint, true, true)); var c2 = Truncated(Point(sourcePoint + 1, true, true));
                            double left = Math.Min(c1.X, c2.X), right = Math.Max(c1.X, c2.X), top = Math.Min(c1.Y, c2.Y), bottom = Math.Max(c1.Y, c2.Y);
                            if (right - left + 1 != 0 && bottom - top + 1 != 0 && right != left && bottom != top)
                            {
                                var start = Truncated(Point(sourcePoint + xor, true, true)); var end = Truncated(Point(sourcePoint + (xor ^ 1), true, true));
                                var arc = Arc(left, top, right, bottom, start, end);
                                if (clockwise) arc.Reverse();
                                AppendAll(arc);
                            }
                            sourcePoint += 4;
                        }
                        break;
                    }
                    case WordPathCommand.ArcAngleTo:
                        for (int i = 0; i < count && sourcePoint + 1 < coordinateCount; i++)
                        {
                            var radii = Point(sourcePoint, false, false); var angles = Point(sourcePoint + 1, false, false);
                            double fwr = radii.X, fhr = radii.Y, start = angles.X, swing = angles.Y;
                            if (IsOoxml) swing = Math.Max(-360, Math.Min(360, swing));
                            if (polygon.Count > 0)
                            {
                                var arc = new List<(double X, double Y)>();
                                if (fwr == 0 && fhr == 0) arc.Add((0, 0));
                                else
                                {
                                    double end = start + swing;
                                    (double X, double Y) centre = (fwr, fhr);
                                    if (swing < 0) (start, end) = (end, start);
                                    double from = start;
                                    if (swing >= 360 || swing <= -360)
                                    {
                                        double to = from + 180;
                                        while (to < end)
                                        {
                                            arc.AddRange(EllipseSegment(centre, fwr, fhr, CircleAngle(fwr, fhr, from), CircleAngle(fwr, fhr, to)));
                                            from = to; to += 180;
                                        }
                                    }
                                    arc.AddRange(EllipseSegment(centre, fwr, fhr, CircleAngle(fwr, fhr, from), CircleAngle(fwr, fhr, end)));
                                    if (swing < 0) arc.Reverse();
                                    RemoveDoublePoints(arc);
                                }
                                for (int k = 0; k < arc.Count; k++) arc[k] = (arc[k].X * _xScale, arc[k].Y * _yScale);
                                var current = polygon[polygon.Count - 1];
                                double dx = current.X - arc[0].X, dy = current.Y - arc[0].Y;
                                foreach (var p in arc) Append((p.X + dx, p.Y + dy));
                            }
                            sourcePoint += 2;
                        }
                        break;
                    case WordPathCommand.EllipticalQuadrantX or WordPathCommand.EllipticalQuadrantY:
                        if (count > 0 && sourcePoint < coordinateCount)
                        {
                            (double X, double Y) start;
                            int i = 0;
                            if (sourcePoint > 0) start = Point(sourcePoint - 1, true, true);
                            else { start = Point(sourcePoint, true, true); Append(start); sourcePoint++; i++; }
                            bool isX = command == WordPathCommand.EllipticalQuadrantX;
                            for (; i < count && sourcePoint < coordinateCount; i++)
                            {
                                var end = Point(sourcePoint, true, true);
                                AppendAll(Quadrant(start, end, isX));
                                sourcePoint++;
                                isX = !isX;
                                start = end;
                            }
                        }
                        break;
                }
                if (command == WordPathCommand.EndSubpath) break;
            }
        }
        if (segmentIndex == Segments.Length) segmentIndex++;
        FinishOpen();
        if (polygons.Count == 0) return;

        bool allClosed = closed.All(c => c);
        bool forceTwo = !SortFilledToBack && !allClosed && !noStroke;
        if (forceTwo || SortFilledToBack)
        {
            if (Filled && !noFill)
            {
                var fill = new PathObject { HasFill = true, HasStroke = false, IsLine = false, IsFillCopy = true, Brightness = brightness };
                fill.Polygons.AddRange(polygons); fill.Closed.AddRange(closed.Select(_ => true));
                objects.Add(fill);
            }
            // a stroke object is made unless 'S'; one is kept (drawing nothing) even then, as the line geometry of a text effect
            {
                var stroke = new PathObject { HasFill = false, HasStroke = Stroked && !noStroke, IsLine = !allClosed, Brightness = brightness };
                stroke.Polygons.AddRange(polygons); stroke.Closed.AddRange(closed);
                objects.Add(stroke);
            }
        }
        else
        {
            var single = new PathObject { Brightness = brightness };
            single.Polygons.AddRange(polygons);
            if (noFill) { single.Closed.AddRange(closed); single.HasFill = false; single.IsLine = !allClosed; }
            else { single.Closed.AddRange(closed.Select(_ => true)); single.HasFill = Filled; }
            single.HasStroke = Stroked && !noStroke;
            objects.Add(single);
        }
    }

    /// <summary>basegfx::utils::checkClosed: a polygon whose last point equals its first is closed and loses the duplicate.</summary>
    private static bool CheckClosed(List<(double X, double Y, bool IsCurve, double C1X, double C1Y, double C2X, double C2Y)> polygon)
    {
        if (polygon.Count < 2) return false;
        var first = polygon[0]; var last = polygon[polygon.Count - 1];
        if (Math.Abs(first.X - last.X) > 1e-9 || Math.Abs(first.Y - last.Y) > 1e-9) return false;
        if (!last.IsCurve) polygon.RemoveAt(polygon.Count - 1);
        return true;
    }

    private static (double X, double Y) Truncated((double X, double Y) p) => (Math.Truncate(p.X), Math.Truncate(p.Y));

    private static void RemoveDoublePoints(List<(double X, double Y)> points)
    {
        for (int i = points.Count - 1; i > 0; i--)
            if (Math.Abs(points[i].X - points[i - 1].X) < 1e-9 && Math.Abs(points[i].Y - points[i - 1].Y) < 1e-9) points.RemoveAt(i);
    }

    private IEnumerable<(double X, double Y)> AngleEllipse((double X, double Y) centre, double scaledWr, double scaledHr, double wr, double hr, double startAngle, double endAngle)
    {
        switch (EllipseMode)
        {
            case WordAngleEllipseMode.Binary:
            {
                // angles in degrees * 2^16 (but RingOutside), the second a swing; 0 = right, counter-clockwise on screen (:2246-2297)
                double swing = endAngle;
                if (!AnglesInDegrees) { startAngle /= 65536.0; swing = endAngle / 65536.0; }
                startAngle = -startAngle; swing = -swing;
                endAngle = startAngle + swing;
                if (swing < 0) (startAngle, endAngle) = (endAngle, startAngle);
                var points = new List<(double X, double Y)>();
                double from = startAngle, to = from + 180;
                while (to < endAngle) { points.AddRange(EllipseSegment(centre, scaledWr, scaledHr, Normalized(from), Normalized(to))); from = to; to += 180; }
                points.AddRange(EllipseSegment(centre, scaledWr, scaledHr, Normalized(from), Normalized(endAngle)));
                if (swing < 0) points.Reverse();
                return points;
            }
            case WordAngleEllipseMode.Full:
                return EllipseFull(centre, scaledWr, scaledHr);
            default:
                if (Math.Abs(Math.Abs(endAngle - startAngle) - 360) < 1e-15)
                {
                    double s = CircleAngle(wr, hr, startAngle), h = CircleAngle(wr, hr, startAngle + 180), e = CircleAngle(wr, hr, endAngle);
                    return EllipseSegment(centre, scaledWr, scaledHr, s, h).Concat(EllipseSegment(centre, scaledWr, scaledHr, h, e)).ToList();
                }
                return EllipseSegment(centre, scaledWr, scaledHr, CircleAngle(wr, hr, startAngle), CircleAngle(wr, hr, endAngle)).ToList();
        }
    }

    /// <summary>basegfx createPolygonFromEllipse: a closed ellipse starting at angle 0 (right), in increasing angle (clockwise on screen).</summary>
    private static List<(double X, double Y)> EllipseFull((double X, double Y) centre, double rx, double ry)
        => EllipseSegment(centre, rx, ry, 0, 2 * Math.PI).ToList();

    /// <summary>basegfx createPolygonFromEllipseSegment: from <paramref name="start"/> to <paramref name="end"/> radians, increasing (y down).</summary>
    internal static List<(double X, double Y)> EllipseSegment((double X, double Y) centre, double rx, double ry, double start, double end)
    {
        if (end <= start) end += 2 * Math.PI;
        int steps = Math.Max(4, (int)Math.Ceiling((end - start) / (Math.PI / 32)));
        var points = new List<(double X, double Y)>(steps + 1);
        for (int i = 0; i <= steps; i++)
        {
            double t = start + (end - start) * i / steps;
            points.Add((centre.X + rx * Math.Cos(t), centre.Y + ry * Math.Sin(t)));
        }
        return points;
    }

    /// <summary>lcl_getNormalizedAngleRad (:2070-2076).</summary>
    private static double Normalized(double degrees)
    {
        double d = degrees % 360;
        if (d < 0) d += 360;
        return d * Math.PI / 180;
    }

    /// <summary>lcl_getNormalizedCircleAngleRad (:2051-2068).</summary>
    internal static double CircleAngle(double wr, double hr, double degrees)
    {
        double d = degrees % 360;
        if (d < 0) d += 360;
        if (d is 0 or 90 or 180 or 270) return d * Math.PI / 180;
        double x = hr * Math.Cos(d * Math.PI / 180), y = wr * Math.Sin(d * Math.PI / 180);
        if (x == 0 && y == 0) return 0;
        double result = Math.Atan2(y, x);
        return result < 0 ? result + 2 * Math.PI : result;
    }

    /// <summary>
    /// CreateArc (:2007-2049, the rectangle already normalised by the caller :2402) through tools::Polygon(rect, start, end,
    /// PolyStyle::Arc) (tools/source/generic/poly.cxx:246-330, ImplGetParameter :60-67): counter-clockwise on screen from the
    /// start ray to the end ray.
    /// </summary>
    private static List<(double X, double Y)> Arc(double left, double top, double right, double bottom, (double X, double Y) start, (double X, double Y) end)
    {
        var points = new List<(double X, double Y)>();
        double cx = (left + right) / 2, cy = (top + bottom) / 2, rx = cx - left, ry = cy - top;
        if (rx <= 0 || ry <= 0) return points;
        double ParameterOf((double X, double Y) p) { double angle = Math.Atan2(cy - p.Y, p.X - cx); return Math.Atan2(rx * Math.Sin(angle), ry * Math.Cos(angle)); }
        double from = ParameterOf(start), diff = ParameterOf(end) - from;
        if (diff <= 0) diff += 2 * Math.PI;
        int steps = Math.Max(16, (int)Math.Ceiling(diff / (Math.PI / 32)));
        for (int i = 0; i <= steps; i++)
        {
            double t = from + diff * i / steps;
            points.Add((cx + rx * Math.Cos(t), cy - ry * Math.Sin(t)));
        }
        return points;
    }

    /// <summary>ELLIPTICALQUADRANTX / Y (:2496-2582): a quarter ellipse from the current point, leaving along x (or y) first.</summary>
    private static List<(double X, double Y)> Quadrant((double X, double Y) from, (double X, double Y) to, bool isX)
    {
        double rx = Math.Abs(to.X - from.X), ry = Math.Abs(to.Y - from.Y);
        (double X, double Y) centre = isX ? (from.X, to.Y) : (to.X, from.Y);
        double half = Math.PI / 2, start, end;
        bool flip;
        if (isX)
        {
            if (to.X < from.X) { if (to.Y < from.Y) { start = half; end = Math.PI; flip = false; } else { start = Math.PI; end = 1.5 * Math.PI; flip = true; } }
            else { if (to.Y < from.Y) { start = 0; end = half; flip = true; } else { start = 1.5 * Math.PI; end = 2 * Math.PI; flip = false; } }
        }
        else
        {
            if (to.X < from.X) { if (to.Y < from.Y) { start = 1.5 * Math.PI; end = 2 * Math.PI; flip = true; } else { start = 0; end = half; flip = false; } }
            else { if (to.Y < from.Y) { start = Math.PI; end = 1.5 * Math.PI; flip = false; } else { start = half; end = Math.PI; flip = true; } }
        }
        var points = EllipseSegment(centre, rx, ry, start, end);
        if (flip) points.Reverse();
        return points;
    }
}
