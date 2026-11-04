namespace FastDocView.Html;

/// <summary>
/// Evaluates an OfficeArt outline - a preset from <see cref="WordOfficeArtShapeDefinitions"/> or a shape's own geometry
/// (pVertices / pSegmentInfo / pGuides) - into <see cref="WordShapePath"/>s in the shape's points, the way LibreOffice draws it:
/// the tables are merged into a custom-shape model as SdrObjCustomShape::MergeDefaultAttributes does (svx svdoashp.cxx:826-1050:
/// adjust defaults, view box, coordinates, segments via lcl_ShapeSegmentFromBinary :96-165, stretch points, text frames,
/// equations through EnhancedCustomShape2d::GetEquation) and drawn by <see cref="WordCustomShapeEngine"/>.
/// - Coordinates: a vertex whose value has the MSO_I mark (EnhancedCustomShapeGeometry.hxx:88) is equation n's result.
/// - Segments (preset tables, svdoashp.cxx:96-165): 0x00 line, 0x20 curve, 0x40 move, 0x60 close, 0x80 end, 0xa1 / 0xa2 angle
///   ellipse (to), 0xa3..0xa6 arc (to) / clockwise, 0xa7 / 0xa8 elliptical quadrant x / y, 0xaa no fill, 0xab no stroke.
/// </summary>
public static class WordOfficeArtGeometry
{
    /// <summary>A preset's paths for a box of <paramref name="width"/> x <paramref name="height"/> points; null when the type has no table.</summary>
    public static List<WordShapePath>? Preset(int shapeType, double width, double height, IReadOnlyList<int?> adjust)
        => PresetEngine(shapeType, width, height, adjust.Select(a => a is { } v ? (double?)v : null).ToList(), WordOfficeArtShapeDefinitions.NameOfType(shapeType), false, false, true, true)?.Build();

    /// <summary>
    /// The custom-shape model of a preset (MergeDefaultAttributes, svdoashp.cxx:826-1050) for a box in points: the given
    /// adjust values (null = the preset's default), the preset's view box, coordinates, segments, stretch points, text
    /// frames and equations; the angle-ellipse rule and colour data come from the type's ODF name (EnhancedCustomShape2d.cxx
    /// :2185-2210, :780-858).
    /// </summary>
    internal static WordCustomShapeEngine? PresetEngine(int shapeType, double width, double height, IReadOnlyList<double?> adjust, string? typeName,
                                                      bool flipHorizontal, bool flipVertical, bool filled, bool stroked)
    {
        if (WordOfficeArtShapeDefinitions.Of(shapeType) is not { } definition) return null;
        var engine = new WordCustomShapeEngine
        {
            CoordWidthG = definition.CoordinateWidth, CoordHeightG = definition.CoordinateHeight,
            LogicWidth = width * WordCustomShapeEngine.HundredthsPerPoint, LogicHeight = height * WordCustomShapeEngine.HundredthsPerPoint,
            FlipHorizontal = flipHorizontal, FlipVertical = flipVertical, Filled = filled, Stroked = stroked,
        };
        ApplyPreset(engine, definition, shapeType, typeName, adjust, hasCoordinates: false, hasSegments: false, hasEquations: false, hasTextFrames: false,
                    hasStretchX: false, hasStretchY: false, hasViewBox: false);
        return engine;
    }

    /// <summary>Fills what the shape does not set from its preset (MergeDefaultAttributes, svdoashp.cxx:826-1050).</summary>
    internal static void ApplyPreset(WordCustomShapeEngine engine, WordOfficeArtShapeDefinition definition, int shapeType, string? typeName, IReadOnlyList<double?> adjust,
                                     bool hasCoordinates, bool hasSegments, bool hasEquations, bool hasTextFrames, bool hasStretchX, bool hasStretchY, bool hasViewBox)
    {
        int count = Math.Max(adjust.Count, definition.Defaults.Length);
        engine.Adjust = new double[count];
        for (int i = 0; i < count; i++)
            engine.Adjust[i] = i < adjust.Count && adjust[i] is { } given ? given : i < definition.Defaults.Length ? definition.Defaults[i] : 0;
        if (!hasViewBox) { engine.CoordLeft = 0; engine.CoordTop = 0; engine.CoordWidthG = definition.CoordinateWidth; engine.CoordHeightG = definition.CoordinateHeight; }
        if (!hasCoordinates) engine.Coordinates = definition.Vertices.Select(WordShapeParameter.FromTable).ToArray();
        if (!hasSegments) engine.Segments = definition.Segments.Select(TableSegment).ToArray();
        if (!hasStretchX && definition.StretchX != int.MinValue) engine.StretchX = definition.StretchX;
        if (!hasStretchY && definition.StretchY != int.MinValue) engine.StretchY = definition.StretchY;
        if (!hasTextFrames) engine.TextFrames = definition.TextRectangles.Select(WordShapeParameter.FromTable).ToArray();
        if (!hasEquations) _ = new WordCustomShapeEquations(Formulas(definition.Calculations), engine);
        ApplyType(engine, shapeType, typeName);
    }

    /// <summary>The type-dependent drawing rules: angle-ellipse mode, sorting, colour data (EnhancedCustomShape2d.cxx:780-858).</summary>
    internal static void ApplyType(WordCustomShapeEngine engine, int? shapeType, string? typeName)
    {
        engine.EllipseMode = WordCustomShapeEngine.EllipseModeOf(typeName);
        engine.AnglesInDegrees = typeName == "mso-spt143";
        engine.IsOoxml = typeName is not null && typeName.StartsWith("ooxml-", StringComparison.Ordinal);
        int type = shapeType ?? 0x0FFF;                         // mso_sptNil (msdffdef.hxx:482-483)
        // SortFilledObjectsToBackByDefault (EnhancedCustomShapeGeometry.cxx:6215-6226): table bit for type 100, every type >= 0x100
        engine.SortFilledToBack = type >= 0x100 || type == 100;
        engine.UsesBrightness = engine.IsOoxml || type == 0;
        engine.ColorData = type switch
        {
            22 => 0x20400000, 16 => 0x302e0000, 189 => 0x502ce400, 190 => 0x702ce4ce, 191 => 0x602ce4c0, 192 => 0x702ce4c5,
            194 or 193 or 196 or 195 or 197 or 199 or 200 => 0x602ce4c0, 198 => 0x702ce4ec, 84 => 0x502ce400, 65 or 96 => 0x20e00000,
            103 or 102 or 104 or 105 => 0x20d00000, 54 or 53 or 108 or 107 or 97 or 98 => 0x30ee0000,
            0x0FFF => typeName switch
            {
                "col-60da8460" => 0x60ecc240u,
                "col-502ad400" => 0x502ce400u,
                { Length: > 4 } name when name.StartsWith("col-", StringComparison.Ordinal) && uint.TryParse(name.Substring(4), System.Globalization.NumberStyles.HexNumber, null, out uint data) => data,
                _ => 0u,
            },
            _ => 0u,
        };
    }

    internal static string[] Formulas(int[] calculations)
    {
        var formulas = new string[calculations.Length / 4];
        for (int i = 0; i < formulas.Length; i++)
            formulas[i] = WordCustomShapeFormula.OfficeArtEquation(calculations[i * 4], calculations[i * 4 + 1], calculations[i * 4 + 2], calculations[i * 4 + 3]);
        return formulas;
    }

    /// <summary>
    /// A shape's own geometry (FOPT pVertices 325, pSegmentInfo 326, pGuides 342; geoLeft..geoBottom 320-323): vertices as
    /// values - an equation reference is a vertex value whose high 16 bits are 0x8000 (SetEnhancedCustomShapeParameter,
    /// EnhancedCustomShape2d.cxx:72-86) - segments, guide records (flags, p1, p2, p3); angle ellipses use the binary-import rule
    /// (degrees in 16.16, a swing angle, EnhancedCustomShape2d.cxx:2246-2297).
    /// </summary>
    public static List<WordShapePath> Custom(IReadOnlyList<int> vertexValues, IReadOnlyList<ushort> segments, int[] guides,
                                             IReadOnlyList<int?> adjust, int left, int top, int right, int bottom, double width, double height)
    {
        var engine = new WordCustomShapeEngine
        {
            CoordLeft = left, CoordTop = top, CoordWidthG = right - left, CoordHeightG = bottom - top,
            LogicWidth = width * WordCustomShapeEngine.HundredthsPerPoint, LogicHeight = height * WordCustomShapeEngine.HundredthsPerPoint,
            EllipseMode = WordAngleEllipseMode.Binary,
            Adjust = Enumerable.Range(0, 10).Select(i => i < adjust.Count && adjust[i] is { } given ? (double)given : 0).ToArray(),
            Coordinates = vertexValues.Select(v => ((uint)v >> 16) == 0x8000 ? WordShapeParameter.EquationResult(v & 0xFFFF) : WordShapeParameter.Integer(v)).ToArray(),
            Segments = segments.Select(BinarySegment).ToArray(),
            UsesBrightness = true,
        };
        _ = new WordCustomShapeEquations(Formulas(guides), engine);
        return engine.Build() ?? new List<WordShapePath>();
    }

    /// <summary>A preset table's segment (svdoashp.cxx:96-165 lcl_ShapeSegmentFromBinary).</summary>
    internal static WordPathSegment TableSegment(ushort data)
    {
        int command = data >> 8, low = data & 0xFF;
        return command switch
        {
            0x00 => new(WordPathCommand.LineTo, Math.Max(1, low)),
            0x20 => new(WordPathCommand.CurveTo, Math.Max(1, low)),
            0x40 => new(WordPathCommand.MoveTo, Math.Max(1, low)),
            0x60 => new(WordPathCommand.CloseSubpath, 0),
            0x80 => new(WordPathCommand.EndSubpath, 0),
            0xA1 => new(WordPathCommand.AngleEllipseTo, low / 3),
            0xA2 => new(WordPathCommand.AngleEllipse, low / 3),
            0xA3 => new(WordPathCommand.ArcTo, low >> 2),
            0xA4 => new(WordPathCommand.Arc, low >> 2),
            0xA5 => new(WordPathCommand.ClockwiseArcTo, low >> 2),
            0xA6 => new(WordPathCommand.ClockwiseArc, low >> 2),
            0xA7 => new(WordPathCommand.EllipticalQuadrantX, low),
            0xA8 => new(WordPathCommand.EllipticalQuadrantY, low),
            0xAA => new(WordPathCommand.NoFill, 0),
            0xAB => new(WordPathCommand.NoStroke, 0),
            _ => new(WordPathCommand.Unknown, 0),
        };
    }

    /// <summary>
    /// A shape's own pSegmentInfo element (msdffimp.cxx:2267-2390: the top 3 bits the command - line, curve, move, close, end -
    /// and 13 bits the count; 5 / 6 = an escape whose next 5 bits pick extension (a line), angle ellipse to / angle ellipse, arc to,
    /// arc, clockwise arc to / arc, quadrant x / y, no fill, no stroke). RTF writes the same values (rtfsdrimport.cxx:531-590).
    /// </summary>
    private static WordPathSegment BinarySegment(ushort data)
    {
        int count = data & 0x1FFF;
        switch (data >> 13)
        {
            case 0: return new(WordPathCommand.LineTo, Math.Max(1, count));
            case 1: return new(WordPathCommand.CurveTo, Math.Max(1, count));
            case 2: return new(WordPathCommand.MoveTo, Math.Max(1, count));
            case 3: return new(WordPathCommand.CloseSubpath, 0);
            case 4: return new(WordPathCommand.EndSubpath, 0);
            case 5: case 6:
                return ((data >> 8) & 0x1F) switch
                {
                    0x0 => new(WordPathCommand.LineTo, Math.Max(1, count)),
                    0x1 => new(WordPathCommand.AngleEllipseTo, (data & 0xFF) / 3),
                    0x2 => new(WordPathCommand.AngleEllipse, (data & 0xFF) / 3),
                    0x3 => new(WordPathCommand.ArcTo, (data & 0xFF) >> 2),
                    0x4 => new(WordPathCommand.Arc, (data & 0xFF) >> 2),
                    0x5 => new(WordPathCommand.ClockwiseArcTo, (data & 0xFF) >> 2),
                    0x6 => new(WordPathCommand.ClockwiseArc, (data & 0xFF) >> 2),
                    0x7 => new(WordPathCommand.EllipticalQuadrantX, data & 0xFF),
                    0x8 => new(WordPathCommand.EllipticalQuadrantY, data & 0xFF),
                    0xA => new(WordPathCommand.NoFill, 0),
                    0xB => new(WordPathCommand.NoStroke, 0),
                    _ => new(WordPathCommand.Unknown, 0),
                };
            default: return new(WordPathCommand.Unknown, 0);
        }
    }
}
