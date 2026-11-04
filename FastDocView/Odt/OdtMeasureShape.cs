using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>A marker on a dimension line: its width (draw:marker-*-width) and its length along the line (points); (0, 0) = no marker.</summary>
internal readonly record struct OdtMeasureArrow(double Width, double Length)
{
    /// <summary>
    /// LibreOffice's arrow length: the marker polygon's height scaled to the marker width, halved when the marker is centred
    /// (svx/source/sdr/primitive2d/sdrmeasureprimitive2d.cxx:150-170).
    /// </summary>
    public static OdtMeasureArrow FromMarker(double width, double viewBoxWidth, double viewBoxHeight, bool centered)
    {
        if (width <= 0 || viewBoxWidth <= 0) return default;
        double length = viewBoxHeight * width / viewBoxWidth;
        return new OdtMeasureArrow(width, centered ? length * 0.5 : length);
    }
}

/// <summary>One drawn piece of a dimension line: the shape's start / end marker is drawn at its start / end only when the flag is set.</summary>
internal readonly record struct OdtMeasurePart(WordShapeSubpath Line, bool StartMarker, bool EndMarker);

/// <summary>A laid-out draw:measure: dimension-line pieces, extension (help) lines, and where the measure text sits.</summary>
internal sealed class OdtMeasureGeometry
{
    public readonly List<OdtMeasurePart> MainLines = new();
    public readonly List<WordShapeSubpath> HelpLines = new();
    /// <summary>Centre of the text block (points, the anchor's coordinates, like svg:x1).</summary>
    public double TextCenterX, TextCenterY;
    /// <summary>Clockwise rotation of the text block in degrees (CSS rotate), 0..360.</summary>
    public double TextRotationDegrees;
    /// <summary>
    /// The inputs of the dimension line and text placement, for the page to lay them out again with the label's measured size
    /// (KitveiHakodesh vue-frontend fastDocViewShapeText.ts, a port of <see cref="OdtMeasureShape.Layout"/>'s text-dependent part):
    /// space-separated numbers - start x, y, end x, y (anchor points), line distance, below (0/1), text rotated (0/1), horizontal
    /// and vertical position (0 automatic, 1 negative, 2 centred, 3 positive), start arrow width, length, end arrow width,
    /// length, line width, padding left, right, top, bottom.
    /// </summary>
    public string PageLayout = "";
}

/// <summary>
/// draw:measure (ODF 1.3 part 3 §10.3.12) laid out as LibreOffice draws it: SdrMeasurePrimitive2D::create2DDecomposition
/// (svx/source/sdr/primitive2d/sdrmeasureprimitive2d.cxx) for the geometry, SdrMeasureObj::TakeRepresentation
/// (svx/source/svdraw/svdomeas.cxx:73-176) and SdrModel::GetMetricString (svx/source/svdraw/svdmodel.cxx:1006-) for the text.
/// </summary>
internal static class OdtMeasureShape
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw, Svg = OdtStyleCascade.Svg, Fo = OdtStyleCascade.Fo, Text = OdtStyleCascade.Text;

    // Item-pool defaults (svx/source/svdraw/svdattr.cxx:288-290), in the pool's 1/100 mm.
    private const double DefaultLineDistance = 800 * 72 / 2540.0, DefaultGuideOverhang = 200 * 72 / 2540.0, DefaultGuideDistance = 100 * 72 / 2540.0;

    private enum TextPosition { Automatic, Negative, Centered, Positive }

    /// <summary>
    /// The measure text: the element's text:p content with each text:measure field filled in as LibreOffice fills it.
    /// LibreOffice drops its own pre-created fields on import and keeps only the file's paragraphs (xmloff/source/draw/ximpshap.cxx:2216-2243),
    /// so a draw:measure without text:p shows no text (empty string).
    /// </summary>
    public static string Label(XElement measure, Func<XName, string?> get)
    {
        var sb = new StringBuilder();
        bool first = true;
        foreach (XElement paragraph in measure.Elements().Where(e => e.Name == Text + "p" || e.Name == Text + "h"))
        {
            if (!first) sb.Append('\n');
            first = false;
            AppendInline(sb, paragraph, measure, get);
        }
        return sb.ToString();
    }

    private static void AppendInline(StringBuilder sb, XElement parent, XElement measure, Func<XName, string?> get)
    {
        foreach (XNode node in parent.Nodes())
        {
            if (node is XText text) { sb.Append(text.Value); continue; }
            if (node is not XElement e) continue;
            if (e.Name == Text + "s") sb.Append(' ', int.TryParse((string?)e.Attribute(Text + "c"), out int c) && c > 0 ? c : 1);
            else if (e.Name == Text + "tab") sb.Append('\t');
            else if (e.Name == Text + "line-break") sb.Append('\n');
            else if (e.Name == Text + "measure") sb.Append(Field(measure, get, (string?)e.Attribute(Text + "kind")));
            else AppendInline(sb, e, measure, get);
        }
    }

    /// <summary>One text:measure field: text:kind value / unit / gap (SdrMeasureFieldKind Value / Unit / Rotate90Blanks, svdomeas.cxx:89-172).</summary>
    public static string Field(XElement measure, Func<XName, string?> get, string? kind)
    {
        bool rotate90 = Bool(get(Draw + "parallel")) ?? false;
        switch (kind)
        {
            case "value":
                return Value(measure, get);
            case "unit":
                if (!(Bool(get(Draw + "show-unit")) ?? false)) return "";        // pool default false, svdattr.cxx:299
                return UnitString(MeasureUnit(get));
            case "gap":
                return rotate90 ? " " : "";                                        // svdomeas.cxx:163-170
            default:
                return "";
        }
    }

    /// <summary>The measured length as LibreOffice Writer prints it (svdomeas.cxx:91-150 + svdmodel.cxx:1006-1100).</summary>
    public static string Value(XElement measure, Func<XName, string?> get)
    {
        // Writer's drawing model counts in twips (sw/source/core/draw/drawdoc.cxx:42 SetScaleUnit(MapTwip)); a coordinate reaches it
        // as 1/100 mm from the importer (convertMeasureToCore), then the UNO layer converts to the pool metric (unoshape.cxx ForceMetricToItemPoolMetric).
        long x1 = Twips(measure.Attribute(Svg + "x1")), y1 = Twips(measure.Attribute(Svg + "y1"));
        long x2 = Twips(measure.Attribute(Svg + "x2")), y2 = Twips(measure.Attribute(Svg + "y2"));
        long length = GetLen(x2 - x1, y2 - y1);

        // UI unit: no Writer code sets the drawing model's UI unit, so SdrModel's own default FieldUnit::MM holds (svdmodel.cxx:127).
        const string uiUnit = "mm";
        string unit = MeasureUnit(get);
        if (unit == "") unit = uiUnit;                                             // svdomeas.cxx:96-97
        double factor = 1.0;
        if (unit != uiUnit) factor *= MillimetreTo(unit);                         // GetMapFactor(eModUIUnit, eMeasureUnit), svdomeas.cxx:101-105
        if (factor != 1.0) length = (long)(length * factor);                      // sal_Int32 = double: truncation, svdomeas.cxx:112-115

        int digits = 2;                                                            // SdrMeasureDecimalPlacesItem default 2 (include/svx/sxmtfitm.hxx:62)
        if (get(Draw + "decimal-places") is { } placesText && int.TryParse(placesText.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int places)) digits = places;
        string text = MetricString(length, digits);

        // SDRATTR_MEASUREAUTODECIMALPLACES: pool default true (svdattr.cxx:306), not in the ODF property map: trailing zeros go (svdomeas.cxx:118-146).
        if (text.IndexOf('.') >= 0)
        {
            int end = text.Length - 1;
            while (end >= 0 && text[end] == '0') end--;
            if (end >= 0 && text[end] == '.') end--;
            text = text.Substring(0, end + 1);
            if (text.Length == 0) text = "0";
        }
        return text;
    }

    private static long Twips(XAttribute? attribute)
    {
        double points = OdtShapeStyle.Length((string?)attribute) ?? 0;
        double hundredthMm = Math.Round(points * 2540 / 72, MidpointRounding.AwayFromZero);
        return (long)Math.Round(hundredthMm * 1440 / 2540, MidpointRounding.AwayFromZero);
    }

    /// <summary>svx/source/svdraw/svdtrans.cxx:402-425 GetLen: the rounded Euclidean length.</summary>
    private static long GetLen(long dx, long dy) => (long)Math.Round(Math.Sqrt((double)dx * dx + (double)dy * dy), MidpointRounding.AwayFromZero);

    /// <summary>
    /// SdrModel::GetMetricString(nVal, bNoUnitChars=true, nNumDigits) for a twip model and a millimetre UI unit (svdmodel.cxx:1006-1100):
    /// ImpSetUIUnit (svdmodel.cxx:872-929) gives twip→mm as 127/7200 (o3tl unit_conversion.hxx:177, reduced at :132-133), i.e. factor 127/72
    /// with the decimal mark 2 places in. Locale: en-US ('.' decimal, ',' thousands, leading zero) - LibreOffice takes these from the system locale.
    /// </summary>
    internal static string MetricString(long value, int digits)
    {
        bool negative = value < 0;
        double local = value * (127.0 / 72.0);
        if (negative) local = -local;
        int decimalMark = 2;
        if (decimalMark > digits) { local /= Math.Pow(10, decimalMark - digits); decimalMark = digits; }
        else if (decimalMark < digits) { local *= Math.Pow(10, digits - decimalMark); decimalMark = digits; }

        var sb = new StringBuilder(((int)(local + 0.5)).ToString(CultureInfo.InvariantCulture));
        if (decimalMark < 0) { sb.Append('0', -decimalMark); decimalMark = 0; }
        if (decimalMark > 0 && sb.Length <= decimalMark) sb.Insert(0, "0", decimalMark - sb.Length + 1);   // + the leading zero
        int beforeMark = sb.Length - decimalMark;
        if (decimalMark > 0) sb.Insert(beforeMark, '.');
        // isNumTrailingZeros: true -> kept here (the auto-decimal step of the caller trims them).
        if (beforeMark > 3)
            for (int i = beforeMark - 3; i > 0; i -= 3) sb.Insert(i, ',');
        if (negative) sb.Insert(0, '-');
        return sb.ToString();
    }

    /// <summary>draw:unit (xmloff/source/draw/sdpropls.cxx:706-718) as the unit's name; "" = automatic (FieldUnit::NONE, the default, include/svx/sxmuitm.hxx:32).</summary>
    private static string MeasureUnit(Func<XName, string?> get) => (get(Draw + "unit")?.Trim()) switch
    {
        "mm" => "mm", "cm" => "cm", "m" => "m", "km" => "km", "pt" => "pt", "pc" => "pc", "inch" => "inch", "ft" => "ft", "mi" => "mi",
        _ => "",
    };

    /// <summary>SdrModel::GetUnitString (svdmodel.cxx:971-1004).</summary>
    private static string UnitString(string unit) => (unit == "" ? "mm" : unit) switch
    {
        "mm" => "mm", "cm" => "cm", "m" => "m", "km" => "km", "pt" => "pt", "pc" => "pica", "inch" => "\"", "ft" => "ft", "mi" => "mile(s)",
        _ => "",
    };

    /// <summary>GetMapFactor(FieldUnit::MM, unit).X() (svdtrans.cxx:619-632): o3tl's millimetre-to-unit ratio.</summary>
    private static double MillimetreTo(string unit) => unit switch
    {
        "cm" => 0.1, "m" => 0.001, "km" => 0.000001, "pt" => 72 / 25.4, "pc" => 6 / 25.4, "inch" => 1 / 25.4, "ft" => 1 / 304.8, "mi" => 1 / 1609344.0,
        _ => 1.0,
    };

    /// <summary>
    /// The dimension line, extension lines and text placement. <paramref name="textWidth"/> / <paramref name="textHeight"/> = the
    /// unrotated text block's size in points (0 when there is no text); the arrows are the shape's markers (draw:marker-start / -end).
    /// </summary>
    public static OdtMeasureGeometry Layout(XElement measure, Func<XName, string?> get, double textWidth, double textHeight,
                                            OdtMeasureArrow startArrow, OdtMeasureArrow endArrow)
    {
        // Points (xmloff/source/draw/ximpshap.cxx:2159-2193, svg:x1 / y1 / x2 / y2).
        double sx = Length(measure.Attribute(Svg + "x1")), sy = Length(measure.Attribute(Svg + "y1"));
        double ex = Length(measure.Attribute(Svg + "x2")), ey = Length(measure.Attribute(Svg + "y2"));

        // Attributes (sdpropls.cxx:223-233; defaults svdattr.cxx:286-294, viewcontactofsdrmeasureobj.cxx:55-103).
        double distance = Length(get(Draw + "line-distance")) ?? DefaultLineDistance;
        double upper = Length(get(Draw + "guide-overhang")) ?? DefaultGuideOverhang;
        double lower = Length(get(Draw + "guide-distance")) ?? DefaultGuideDistance;
        double leftDelta = Length(get(Draw + "start-guide")) ?? 0;
        double rightDelta = Length(get(Draw + "end-guide")) ?? 0;
        bool below = get(Draw + "placing")?.Trim() == "below";                   // XMLNamedBoolPropertyHdl(below, above), sdpropls.cxx:1326-1329
        bool textRotation = Bool(get(Draw + "parallel")) ?? false;              // draw:parallel -> MeasureTextRotate90 as is (sdpropls.cxx:233)
        const bool textAutoAngle = true;                                         // SdrMeasureTextAutoAngleItem default (svx/inc/sxmtaitm.hxx:31); not in the ODF map
        TextPosition horizontal = (get(Draw + "measure-align")?.Trim()) switch
        {
            "left-outside" => TextPosition.Negative, "inside" => TextPosition.Centered, "right-outside" => TextPosition.Positive, _ => TextPosition.Automatic,
        };
        TextPosition vertical = (get(Draw + "measure-vertical-align")?.Trim()) switch
        {
            "above" => TextPosition.Negative, "center" => TextPosition.Centered, "below" => TextPosition.Positive, _ => TextPosition.Automatic,
        };
        double lineWidth = OdtShapeStyle.Length(get(Svg + "stroke-width")) ?? 0;
        double padLeft = Length(get(Fo + "padding-left")) ?? 0, padRight = Length(get(Fo + "padding-right")) ?? 0;
        double padTop = Length(get(Fo + "padding-top")) ?? 0, padBottom = Length(get(Fo + "padding-bottom")) ?? 0;

        var result = new OdtMeasureGeometry();
        result.PageLayout = string.Join(" ", new[] { sx, sy, ex, ey, distance, below ? 1 : 0, textRotation ? 1 : 0, (int)horizontal, (int)vertical,
                                                     startArrow.Width, startArrow.Length, endArrow.Width, endArrow.Length, lineWidth, padLeft, padRight, padTop, padBottom }
                                                   .Select(v => Math.Round(v, 4).ToString(CultureInfo.InvariantCulture)));
        double lineDx = ex - sx, lineDy = ey - sy;
        double length = Math.Sqrt(lineDx * lineDx + lineDy * lineDy);
        double angle = Math.Atan2(lineDy, lineDx);
        double cos = Math.Cos(angle), sin = Math.Sin(angle);
        (double X, double Y) ToAnchor(double x, double y) => (sx + x * cos - y * sin, sy + x * sin + y * cos);
        void Part(List<OdtMeasurePart>? main, double x1, double y1, double x2, double y2, bool left, bool right)
        {
            var s = new WordShapeSubpath();
            (s.StartX, s.StartY) = ToAnchor(x1, y1);
            var (px, py) = ToAnchor(x2, y2);
            s.Segments.Add(WordShapeSegment.Line(px, py));
            if (main != null) main.Add(new OdtMeasurePart(s, left && startArrow.Width > 0, right && endArrow.Width > 0));
            else result.HelpLines.Add(s);
        }

        // Text block range after its own rotation (sdrmeasureprimitive2d.cxx:93-127): a centred block, turned -90° when rotated.
        bool hasText = textWidth > 0 || textHeight > 0;
        double rangeWidth = textRotation ? textHeight : textWidth, rangeHeight = textRotation ? textWidth : textHeight;
        bool autoUpsideDown = false;
        if (hasText)
        {
            double testAngle = angle;
            if (textRotation)
            {
                testAngle -= Math.PI / 2;
                if (textAutoAngle && testAngle < -Math.PI) testAngle += 2 * Math.PI;
            }
            if (textAutoAngle && (testAngle > Math.PI / 4 || testAngle < -Math.PI * 0.75)) autoUpsideDown = true;
        }

        // :137-170
        double startW = startArrow.Width, startH = startArrow.Length, endW = endArrow.Width, endH = endArrow.Length;
        double spaceNeeded = startH + endH + (startW + endW) * 0.5;
        double outsideLength = (startH + endH + startW + endW) * 0.5;
        double halfLineWidth = lineWidth * 0.5;
        bool arrowsOutside = spaceNeeded > length, splitted = false;

        // :178-224
        if (vertical == TextPosition.Automatic) vertical = TextPosition.Negative;
        if (vertical == TextPosition.Centered) splitted = true;
        if (horizontal == TextPosition.Automatic)
        {
            horizontal = rangeWidth > length ? TextPosition.Negative : TextPosition.Centered;
            if (splitted) { if (rangeWidth + spaceNeeded > length) arrowsOutside = true; }
            else if (rangeWidth + startH + endH + (startW + endW) * 0.125 > length) arrowsOutside = true;
        }
        if (horizontal != TextPosition.Centered) arrowsOutside = true;
        if (below || (autoUpsideDown && !textRotation))
        {
            if (vertical == TextPosition.Negative) vertical = TextPosition.Positive;
            else if (vertical == TextPosition.Positive) vertical = TextPosition.Negative;
        }

        // Main line (:226-271).
        double mainY = below ? distance : -distance;
        if (arrowsOutside)
        {
            double lenLeft = outsideLength, lenRight = outsideLength;
            if (!splitted)
            {
                if (horizontal == TextPosition.Negative) lenLeft = startH + rangeWidth;
                else if (horizontal == TextPosition.Positive) lenRight = endH + rangeWidth;
            }
            Part(result.MainLines, -lenLeft, mainY, 0, mainY, false, true);
            Part(result.MainLines, length, mainY, length + lenRight, mainY, true, false);
            if (!splitted || horizontal != TextPosition.Centered) Part(result.MainLines, 0, mainY, length, mainY, false, false);
        }
        else if (splitted)
        {
            double half = (length - (rangeWidth + (startH + endH) * 0.25)) * 0.5;
            Part(result.MainLines, 0, mainY, half, mainY, true, false);
            Part(result.MainLines, length - half, mainY, length, mainY, false, true);
        }
        else Part(result.MainLines, 0, mainY, length, mainY, true, true);

        // Extension lines (:273-288).
        double topEdge = below ? upper + distance : -upper - distance;
        double bottomLeft = below ? lower - leftDelta : leftDelta - lower;
        double bottomRight = below ? lower - rightDelta : rightDelta - lower;
        Part(null, 0, bottomLeft, 0, topEdge, false, false);
        Part(null, length, bottomRight, length, topEdge, false, false);

        // Text position (:290-367): top-left of the text range in the line's frame.
        double textX, textY;
        if (horizontal == TextPosition.Negative)
        {
            textX = -(startH + rangeWidth + outsideLength * 0.18 + halfLineWidth);
            if (splitted) textX -= outsideLength - startH;
            textX -= padRight;
        }
        else if (horizontal == TextPosition.Positive)
        {
            textX = length + endH + outsideLength * 0.18 + halfLineWidth;
            if (splitted) textX += outsideLength - endH;
            textX += padLeft;
        }
        else textX = (length - rangeWidth) * 0.5 + (padLeft - padRight) / 2;
        if (vertical == TextPosition.Negative) textY = mainY - (rangeHeight + outsideLength * 0.10 + halfLineWidth) - padBottom;
        else if (vertical == TextPosition.Positive) textY = mainY + outsideLength * 0.10 + halfLineWidth + padTop;
        else textY = mainY - rangeHeight * 0.5 + (padTop - padBottom) / 2;

        // The block is centred on its origin, so turning it by 180° (:381-385) keeps its centre; the object matrix places it (:387-391).
        (result.TextCenterX, result.TextCenterY) = ToAnchor(textX + rangeWidth / 2, textY + rangeHeight / 2);
        double degrees = angle * 180 / Math.PI + (textRotation ? -90 : 0) + (autoUpsideDown ? 180 : 0);
        degrees %= 360;
        if (degrees < 0) degrees += 360;
        if (Math.Abs(degrees - 360) < 1e-9) degrees = 0;
        result.TextRotationDegrees = degrees;
        return result;
    }

    private static double Length(XAttribute? attribute) => OdtShapeStyle.Length((string?)attribute) ?? 0;
    private static double? Length(string? value) => OdtShapeStyle.Length(value);
    private static bool? Bool(string? value) => value?.Trim() switch { "true" => true, "false" => false, _ => null };
}
