using System.Globalization;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// A draw:custom-shape's draw:enhanced-geometry (ODF 1.3 part 3 10.6.2) read as LibreOffice reads it
/// (xmloff/source/draw/ximpcustomshape.cxx, master read 2026-11-04) into the shared custom-shape model
/// (<see cref="WordCustomShapeEngine"/>):
/// - parameters (GetNextParameter :194-430): '$n' adjustment, '?name' equation (resolved by name, unknown = 0, :1253-1318),
///   keywords left top right bottom xstretch ystretch hasstroke hasfill width height logwidth logheight (case-insensitive),
///   numbers (an integer unless it has '.' or an exponent), separated by spaces / commas;
/// - draw:enhanced-path (GetEnhancedPath :577-850, ODF 19.145): M L C Z N F S T U A B W V X Y Q G H I J K with their parameter
///   counts, implicit repetition, a moveto followed by more pairs continues as lineto, and the correction for curved arrows
///   written by old LibreOffice (mso-spt102..105 with two arcs in the first segment, :818-835);
/// - draw:equation draw:name / draw:formula (:1404-1428), draw:modifiers (GetAdjustmentValues :853-880: a non-number is
///   "default"), draw:text-areas, draw:view-box (integers), draw:sub-view-size, draw:path-stretchpoint-x / -y,
///   draw:mirror-horizontal / -vertical, draw:text-rotate-angle, draw:text-path*;
/// - what the shape does not set comes from its draw:type's preset (SdrObjCustomShape::MergeDefaultAttributes,
///   svdoashp.cxx:826-1050) - per property: adjustments, view box, coordinates, segments, stretch points, text frames,
///   equations. A type without a preset and without a path draws nothing (CreatePathObj :2890-2893). A "rectangle" /
///   "ooxml-rect" is drawn as a plain rectangle (CreateObject :3042-3055).
/// draw:handle and glue points do not change the drawing and are not read. 3-D extrusion (draw:extrusion="true") is drawn as
/// the flat shape.
/// </summary>
internal sealed class OdtEnhancedGeometry
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw;

    public string? Type;
    public bool MirrorHorizontal, MirrorVertical;
    public bool IsTextPath;
    public string TextPathMode = "normal";
    public bool TextPathScaleX, TextPathSameLetterHeights;
    public double TextRotateAngle;
    public bool IsRectangle;
    public WordCustomShapeEngine? Engine;

    /// <summary>The geometry of a custom shape of <paramref name="width"/> x <paramref name="height"/> points.</summary>
    public static OdtEnhancedGeometry Read(XElement? geometry, double width, double height, bool filled, bool stroked)
    {
        var result = new OdtEnhancedGeometry();
        if (geometry is null) return result;
        string? Attribute(string name) => (string?)geometry.Attribute(Draw + name);
        result.Type = Attribute("type");
        result.MirrorHorizontal = Attribute("mirror-horizontal") == "true";
        result.MirrorVertical = Attribute("mirror-vertical") == "true";
        result.IsTextPath = Attribute("text-path") == "true";
        result.TextPathMode = Attribute("text-path-mode") ?? "normal";
        result.TextPathScaleX = Attribute("text-path-scale") == "shape";
        result.TextPathSameLetterHeights = Attribute("text-path-same-letter-heights") == "true";
        if (double.TryParse(Attribute("text-rotate-angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out double textRotate)) result.TextRotateAngle = textRotate;

        int? presetType = result.Type is null ? null : WordOfficeArtShapeDefinitions.TypeOfName(result.Type);
        if (presetType == WordOfficeArtShapeTypes.Rectangle) { result.IsRectangle = true; return result; }

        var engine = new WordCustomShapeEngine
        {
            LogicWidth = width * WordCustomShapeEngine.HundredthsPerPoint, LogicHeight = height * WordCustomShapeEngine.HundredthsPerPoint,
            Filled = filled, Stroked = stroked, FlipHorizontal = result.MirrorHorizontal, FlipVertical = result.MirrorVertical,
        };
        // equations: names in document order; references by name become indices
        var formulas = new List<string>();
        var names = new List<string>();
        foreach (XElement equation in geometry.Elements(Draw + "equation"))
        {
            string name = (string?)equation.Attribute(Draw + "name") ?? "", formula = (string?)equation.Attribute(Draw + "formula") ?? "";
            if (name.Length == 0 && formula.Length == 0) continue;
            formulas.Add(formula); names.Add(name);
        }
        var indexOfName = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < names.Count; i++) indexOfName[names[i]] = i;

        bool hasViewBox = false;
        if (OdtSvgPath.ViewBox(Attribute("viewBox")) is { } box)
        {
            hasViewBox = true;
            engine.CoordLeft = (int)Math.Round(box.X, MidpointRounding.AwayFromZero); engine.CoordTop = (int)Math.Round(box.Y, MidpointRounding.AwayFromZero);
            engine.CoordWidthG = Math.Abs((int)Math.Round(box.Width, MidpointRounding.AwayFromZero)); engine.CoordHeightG = Math.Abs((int)Math.Round(box.Height, MidpointRounding.AwayFromZero));
        }
        if (Attribute("sub-view-size") is { } subView)
        {
            var numbers = subView.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(t => int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? (int?)v : null).TakeWhile(v => v is not null).Select(v => v!.Value).ToList();
            engine.SubViewSizes = Enumerable.Range(0, (numbers.Count + 1) / 2).Select(i => (numbers[i * 2], i * 2 + 1 < numbers.Count ? numbers[i * 2 + 1] : 0)).ToArray();
        }
        bool hasStretchX = int.TryParse(Attribute("path-stretchpoint-x"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int stretchX);
        bool hasStretchY = int.TryParse(Attribute("path-stretchpoint-y"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int stretchY);
        if (hasStretchX) engine.StretchX = stretchX;
        if (hasStretchY) engine.StretchY = stretchY;

        bool hasPath = false;
        if (Attribute("enhanced-path") is { } path)
        {
            (List<WordShapeParameter> coordinates, List<WordPathSegment> segments) = Path(path, result.Type, indexOfName);
            engine.Coordinates = coordinates.ToArray();
            engine.Segments = segments.ToArray();
            hasPath = true;
        }
        bool hasTextFrames = false;
        if (Attribute("text-areas") is { } areas)
        {
            var parameters = new List<WordShapeParameter>();
            int index = 0;
            while (true)
            {
                var four = new WordShapeParameter[4];
                bool complete = true;
                for (int k = 0; k < 4 && complete; k++) complete = NextParameter(areas, ref index, indexOfName, out four[k]);
                if (!complete) break;
                parameters.AddRange(four);
            }
            engine.TextFrames = parameters.ToArray();
            hasTextFrames = true;
        }
        var adjust = new List<double?>();
        if (Attribute("modifiers") is { } modifiers)
        {
            int index = 0;
            while (NextParameter(modifiers, ref index, indexOfName, out WordShapeParameter parameter))
                adjust.Add(parameter.Kind == WordShapeParameterKind.Normal ? parameter.Value : null);
        }
        bool hasEquations = formulas.Count > 0;
        if (hasEquations)
            _ = new WordCustomShapeEquations(formulas.Select(f => WordCustomShapeFormula.ResolveNames(f, indexOfName)).ToArray(), engine);

        if (presetType is { } type && WordOfficeArtShapeDefinitions.Of(type) is { } definition)
            WordOfficeArtGeometry.ApplyPreset(engine, definition, type, result.Type, adjust, hasPath, hasPath, hasEquations, hasTextFrames, hasStretchX, hasStretchY, hasViewBox);
        else
        {
            engine.Adjust = adjust.Select(a => a ?? 0).ToArray();
            WordOfficeArtGeometry.ApplyType(engine, presetType ?? (result.Type == "non-primitive" ? 0 : null), result.Type);
        }
        result.Engine = engine;
        return result;
    }

    /// <summary>GetEnhancedPath (ximpcustomshape.cxx:577-850).</summary>
    internal static (List<WordShapeParameter> coordinates, List<WordPathSegment> segments) Path(string value, string? type, IReadOnlyDictionary<string, int> indexOfName)
    {
        var coordinates = new List<WordShapeParameter>();
        var segments = new List<WordPathSegment>();
        int index = 0, parameterCount = 0, needed = 1;
        WordPathCommand latest = WordPathCommand.MoveTo;
        bool valid = true;
        while (valid && index < value.Length)
        {
            char c = value[index];
            (WordPathCommand command, int count)? letter = c switch
            {
                'M' => (WordPathCommand.MoveTo, 1), 'L' => (WordPathCommand.LineTo, 1), 'C' => (WordPathCommand.CurveTo, 3),
                'Z' => (WordPathCommand.CloseSubpath, 0), 'N' => (WordPathCommand.EndSubpath, 0), 'F' => (WordPathCommand.NoFill, 0),
                'S' => (WordPathCommand.NoStroke, 0), 'T' => (WordPathCommand.AngleEllipseTo, 3), 'U' => (WordPathCommand.AngleEllipse, 3),
                'A' => (WordPathCommand.ArcTo, 4), 'B' => (WordPathCommand.Arc, 4), 'G' => (WordPathCommand.ArcAngleTo, 2),
                'H' => (WordPathCommand.Darken, 0), 'I' => (WordPathCommand.DarkenLess, 0), 'J' => (WordPathCommand.Lighten, 0),
                'K' => (WordPathCommand.LightenLess, 0), 'W' => (WordPathCommand.ClockwiseArcTo, 4), 'V' => (WordPathCommand.ClockwiseArc, 4),
                'X' => (WordPathCommand.EllipticalQuadrantX, 1), 'Y' => (WordPathCommand.EllipticalQuadrantY, 1), 'Q' => (WordPathCommand.QuadraticCurveTo, 2),
                _ => null,
            };
            if (letter is { } l) { latest = l.command; needed = l.count; index++; }
            else if (c is '$' or '?' or (>= '0' and <= '9') or '.' or '-')
            {
                if (NextParameter(value, ref index, indexOfName, out WordShapeParameter first) && NextParameter(value, ref index, indexOfName, out WordShapeParameter second))
                {
                    coordinates.Add(first); coordinates.Add(second);
                    parameterCount++;
                }
                else valid = false;
            }
            else index++;
            if (parameterCount == 0 && needed == 0)
            {
                segments.Add(new WordPathSegment(latest, 0));
                needed = int.MaxValue;
            }
            else if (parameterCount >= needed)
            {
                if (latest == WordPathCommand.MoveTo)
                {
                    // ODF 1.2 19.145: further pairs after a moveto are linetos
                    segments.Add(new WordPathSegment(WordPathCommand.MoveTo, 1));
                    latest = WordPathCommand.LineTo;
                    needed = 1;
                }
                else if (segments.Count > 0 && segments[segments.Count - 1].Command == latest)
                    segments[segments.Count - 1] = new WordPathSegment(latest, segments[segments.Count - 1].Count + 1);
                else segments.Add(new WordPathSegment(latest, 1));
                parameterCount = 0;
            }
        }
        if (segments.Count > 0 && type is "mso-spt102" or "mso-spt103" or "mso-spt104" or "mso-spt105" && segments[0].Count == 2)
        {
            WordPathCommand first = segments[0].Command;
            segments[0] = new WordPathSegment(first, 1);
            segments.Insert(1, new WordPathSegment(first == WordPathCommand.ClockwiseArc ? WordPathCommand.ClockwiseArcTo : WordPathCommand.ArcTo, 1));
        }
        return (coordinates, segments);
    }

    private static readonly (string Name, WordShapeParameterKind Kind)[] Keywords =
    [
        ("left", WordShapeParameterKind.Left), ("top", WordShapeParameterKind.Top), ("right", WordShapeParameterKind.Right), ("bottom", WordShapeParameterKind.Bottom),
        ("xstretch", WordShapeParameterKind.Other), ("ystretch", WordShapeParameterKind.Other), ("hasstroke", WordShapeParameterKind.Other), ("hasfill", WordShapeParameterKind.Other),
        ("width", WordShapeParameterKind.Other), ("height", WordShapeParameterKind.Other), ("logwidth", WordShapeParameterKind.Other), ("logheight", WordShapeParameterKind.Other),
    ];

    /// <summary>GetNextParameter (ximpcustomshape.cxx:194-430).</summary>
    internal static bool NextParameter(string text, ref int index, IReadOnlyDictionary<string, int> indexOfName, out WordShapeParameter parameter)
    {
        parameter = default;
        if (index >= text.Length) return false;
        bool numberRequired = true, wholePositive = false;
        var kind = WordShapeParameterKind.Normal;
        if (text[index] == '$') { kind = WordShapeParameterKind.Adjustment; wholePositive = true; index++; }
        else if (text[index] == '?')
        {
            index++;
            numberRequired = false;
            int start = index;
            while (index < text.Length && WordCustomShapeFormula.IsNameChar(text[index])) index++;
            if (index == start) return false;
            string name = text.Substring(start, index - start);
            parameter = WordShapeParameter.EquationResult(indexOfName.TryGetValue(name, out int equation) ? equation : 0);
        }
        else if (text[index] > '9')
        {
            numberRequired = false;
            bool matched = false;
            foreach ((string name, WordShapeParameterKind keyword) in Keywords)
                if (string.Compare(text, index, name, 0, name.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    parameter = new WordShapeParameter(keyword, 0, true);
                    index += name.Length;
                    matched = true;
                    break;
                }
            if (!matched) return false;
        }
        if (numberRequired)
        {
            int start = index, eIndex = 0;
            bool e = false, eNumber = false, dot = false, end = false;
            while (index < text.Length)
            {
                char c = text[index];
                if (c == '.') { if (wholePositive || dot) return false; dot = true; }
                else if (c == '-')
                {
                    if (wholePositive) return false;
                    if (start == index) { }
                    else if (e) { if (eIndex + 1 == index) { } else if (eNumber) end = true; else return false; }
                }
                else if (c is 'e' or 'E') { if (wholePositive) end = true; else if (!e) { e = true; eIndex = index; } else end = true; }
                else if (c is >= '0' and <= '9') { if (e && !eNumber) eNumber = true; }
                else end = true;
                if (end) break;
                index++;
            }
            if (index == start) return false;
            string number = text.Substring(start, index - start);
            if (e || dot)
            {
                if (!double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out double real)) return false;
                parameter = new WordShapeParameter(kind, real, false);
            }
            else
            {
                if (!int.TryParse(number, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int whole)) return false;
                parameter = new WordShapeParameter(kind, whole, true);
            }
        }
        while (index < text.Length && text[index] is ' ' or ',') index++;
        return true;
    }
}
