using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>A user glue point of a shape (draw:glue-point, its draw:id) placed in the shape's anchor coordinates.</summary>
internal readonly record struct OdtConnectedGluePoint(int Id, OdtConnectorGluePoint Point);

/// <summary>
/// What a draw:connector needs of a shape it is glued to (draw:start-shape / draw:end-shape): the shape's snap rectangle in points in
/// its anchor's coordinates, its user glue points, and which coordinates those are (<see cref="FrameKey"/>: "p:N" = page N, "b:N" =
/// anchored to block N's paragraph, "" = a frame of its own - as-character, frame-anchored or inside a text box). Recorded at index
/// time, so a connector's block routes without reading the block the shape stands in.
/// </summary>
internal sealed record OdtConnectedShape(string FrameKey, (double X, double Y, double Width, double Height) SnapRect, OdtConnectedGluePoint[] GluePoints)
{
    /// <summary>The same shape moved by (dx, dy): its snap rectangle and glue points.</summary>
    public OdtConnectedShape Moved(double dx, double dy)
        => this with
        {
            SnapRect = (SnapRect.X + dx, SnapRect.Y + dy, SnapRect.Width, SnapRect.Height),
            GluePoints = GluePoints.Select(g => g with { Point = g.Point with { X = g.Point.X + dx, Y = g.Point.Y + dy } }).ToArray(),
        };
}

/// <summary>
/// The shapes of a document that carry an id (draw:id / xml:id), by id - the ones a connector may be glued to. LibreOffice writes an id only
/// on a shape something refers to (xmloff shapeexport.cxx: the id comes from the interface-to-identifier mapper the connector registered
/// it with), so the table stays small. Index-time data (<see cref="OdtBodyBlockScanner"/>), stored in the index (format 10).
/// </summary>
public sealed class OdtConnectedShapeTable
{
    private static readonly XNamespace Draw = OdtStyleCascade.Draw, Svg = OdtStyleCascade.Svg, TextNs = OdtStyleCascade.Text;
    private static readonly XNamespace XmlNs = XNamespace.Xml;
    private readonly Dictionary<string, OdtConnectedShape> _shapes;

    public static OdtConnectedShapeTable Empty { get; } = new(new Dictionary<string, OdtConnectedShape>(StringComparer.Ordinal));

    internal OdtConnectedShapeTable(Dictionary<string, OdtConnectedShape> shapes) => _shapes = shapes;

    public int Count => _shapes.Count;

    internal OdtConnectedShape? Find(string id) => _shapes.TryGetValue(id, out OdtConnectedShape? shape) ? shape : null;

    // ── index time ────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Records the shapes with an id in a drawing element captured during the index pass (the element itself and its descendants). Shapes reached
    /// through draw:g / draw:a only share <paramref name="frameKey"/> (a group's children are in its coordinates, ximpshap.cxx group context sets
    /// no transformation); a shape inside a text box has coordinates of its own (key "").
    /// </summary>
    internal static void Record(XElement root, string frameKey, OdtShapeBuilder builder, Dictionary<string, OdtConnectedShape> into)
    {
        foreach (XElement element in root.DescendantsAndSelf())
        {
            if (element.Name.Namespace != Draw) continue;
            string? drawId = (string?)element.Attribute(Draw + "id"), xmlId = (string?)element.Attribute(XmlNs + "id");
            if (string.IsNullOrEmpty(drawId) && string.IsNullOrEmpty(xmlId)) continue;
            if (SnapRectOf(element, builder) is not { } snap) continue;
            string key = frameKey;
            for (XElement? up = element == root ? null : element.Parent; up is not null && key.Length > 0; up = up.Parent)
            {
                if (up.Name != Draw + "g" && up.Name != Draw + "a") key = "";
                if (up == root) break;
            }
            var shape = new OdtConnectedShape(key, snap, GluePointsOf(element, snap));
            if (!string.IsNullOrEmpty(drawId)) into[drawId!] = shape;
            if (!string.IsNullOrEmpty(xmlId)) into[xmlId!] = shape;
        }
    }

    /// <summary>
    /// The snap rectangle in the anchor's points: a drawing shape's outlines through their transforms (what the renderer places, the
    /// rotated box's bounds as SdrTextObj::RecalcSnapRect, svdotext.cxx:1103-1111; a group = the union of its children); a draw:frame /
    /// draw:control = svg:x / y / width / height through draw:transform. A connector or measure is not a glue target here.
    /// </summary>
    private static (double X, double Y, double Width, double Height)? SnapRectOf(XElement element, OdtShapeBuilder builder)
    {
        string local = element.Name.LocalName;
        if (local is "connector" or "glue-point") return null;
        if (OdtShapeBuilder.IsShape(element))
        {
            List<OdtDrawnShape> drawn = builder.Build(element);
            return drawn.Count == 0 ? null : OdtShapeBuilder.SnapBounds(drawn);
        }
        if (local is not ("frame" or "control")) return null;
        double Length(string name) => OdtShapeStyle.Length((string?)element.Attribute(Svg + name)) ?? 0;
        double x = Length("x"), y = Length("y"), w = Length("width"), h = Length("height");
        WordAffine t = OdtShapeBuilder.Transform((string?)element.Attribute(Draw + "transform"));
        double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
        foreach ((double cx, double cy) in new[] { (x, y), (x + w, y), (x, y + h), (x + w, y + h) })
        {
            (double px, double py) = t.Apply(cx, cy);
            left = Math.Min(left, px); top = Math.Min(top, py); right = Math.Max(right, px); bottom = Math.Max(bottom, py);
        }
        return (left, top, right - left, bottom - top);
    }

    /// <summary>
    /// draw:glue-point children (SdXMLShapeContext::addGluePoint, ximpshap.cxx:247-320): draw:id (required), svg:x / svg:y read as lengths in
    /// 1/100 mm, draw:align (absent = relative: a percentage of the snap rectangle in 1/100 % from its centre), draw:escape-direction.
    /// </summary>
    private static OdtConnectedGluePoint[] GluePointsOf(XElement element, (double X, double Y, double Width, double Height) snap)
    {
        var points = new List<OdtConnectedGluePoint>();
        foreach (XElement glue in element.Elements(Draw + "glue-point"))
        {
            if (!int.TryParse((string?)glue.Attribute(Draw + "id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id)) continue;
            long Hmm(string name) => (long)Math.Round((OdtShapeStyle.Length((string?)glue.Attribute(Svg + name)) ?? 0) * 2540.0 / 72.0, MidpointRounding.AwayFromZero);
            OdtGlueEscape escape = (string?)glue.Attribute(Draw + "escape-direction") switch
            {
                "left" => OdtGlueEscape.Left, "right" => OdtGlueEscape.Right, "up" => OdtGlueEscape.Top, "down" => OdtGlueEscape.Bottom,
                "horizontal" => OdtGlueEscape.Horizontal, "vertical" => OdtGlueEscape.Vertical, _ => OdtGlueEscape.Smart,
            };
            points.Add(new OdtConnectedGluePoint(id, OdtConnectorRoute.UserGluePoint(snap, Hmm("x"), Hmm("y"), (string?)glue.Attribute(Draw + "align"), escape)));
        }
        return points.ToArray();
    }

    /// <summary>
    /// The coordinates a drawing element is in, as <see cref="OdtConnectedShape.FrameKey"/>, and its outermost drawing element in the block.
    /// <paramref name="isBodyLevel"/>: the outermost element is a block of its own (a page-anchored object at body level).
    /// </summary>
    internal static string FrameKeyOf(string? anchorType, string? anchorPage, bool isBodyLevel, int blockIndex)
    {
        if (isBodyLevel || anchorType == "page") return anchorPage is null ? "" : "p:" + anchorPage;   // no page number: no common origin known
        if (anchorType is "paragraph" or "char" && blockIndex >= 0) return "b:" + blockIndex.ToString(CultureInfo.InvariantCulture);
        return "";
    }

    /// <summary>The render-time twin of the index pass's key: the connector's outermost drawing ancestor in the block and its anchor.</summary>
    internal static (string Key, XElement Outermost) FrameOf(XElement drawing, int blockIndex)
    {
        XElement outermost = drawing;
        for (XElement? up = drawing.Parent; up is not null; up = up.Parent)
            if (up.Name.Namespace == Draw && Array.IndexOf(OdtBodyBlockScanner.BodyDrawingNames, "draw:" + up.Name.LocalName) >= 0) outermost = up;
        // only groups (draw:g, draw:a) from the connector up to the outermost drawing: the outermost's coordinates; else its own
        bool throughGroups = true;
        for (XElement? up = drawing.Parent; up is not null && up != outermost.Parent; up = up.Parent)
            if (up.Name != Draw + "g" && up.Name != Draw + "a") { throughGroups = false; break; }
        bool isBodyLevel = outermost.Parent is null;
        string key = throughGroups
            ? FrameKeyOf((string?)outermost.Attribute(TextNs + "anchor-type"), (string?)outermost.Attribute(TextNs + "anchor-page-number"), isBodyLevel, blockIndex)
            : "";
        return (key, outermost);
    }

    /// <summary>Whether a shape with this id stands inside the given element (same drawing, same coordinates when reached through groups).</summary>
    internal static bool Contains(XElement outermost, string id)
        => outermost.DescendantsAndSelf().Any(e => (string?)e.Attribute(Draw + "id") == id || (string?)e.Attribute(XmlNs + "id") == id);

    // ── storage ───────────────────────────────────────────────────────────────────────

    internal void Write(BinaryWriter writer)
    {
        string[] ids = _shapes.Keys.OrderBy(id => id, StringComparer.Ordinal).ToArray();
        IndexBinaryArrays.WriteCompressedStrings(writer, ids);
        IndexBinaryArrays.WriteCompressedStrings(writer, ids.Select(id => _shapes[id].FrameKey).ToArray());
        using var raw = new MemoryStream();
        using (var body = new BinaryWriter(raw, Encoding.UTF8, leaveOpen: true))
            foreach (string id in ids)
            {
                OdtConnectedShape shape = _shapes[id];
                body.Write(shape.SnapRect.X); body.Write(shape.SnapRect.Y); body.Write(shape.SnapRect.Width); body.Write(shape.SnapRect.Height);
                body.Write(shape.GluePoints.Length);
                foreach (OdtConnectedGluePoint glue in shape.GluePoints)
                {
                    body.Write(glue.Id); body.Write(glue.Point.X); body.Write(glue.Point.Y); body.Write((byte)glue.Point.Escape);
                }
            }
        WordIndexDeflatedBlob.Write(writer, raw.ToArray());
    }

    internal static OdtConnectedShapeTable Read(BinaryReader reader)
    {
        string[] ids = IndexBinaryArrays.ReadCompressedStrings(reader);
        string[] keys = IndexBinaryArrays.ReadCompressedStrings(reader);
        if (keys.Length != ids.Length) throw new InvalidDataException("OpenDocument index: connected shapes damaged.");
        var shapes = new Dictionary<string, OdtConnectedShape>(ids.Length, StringComparer.Ordinal);
        using var body = new BinaryReader(new MemoryStream(WordIndexDeflatedBlob.Read(reader)), Encoding.UTF8);
        for (int i = 0; i < ids.Length; i++)
        {
            var snap = (body.ReadDouble(), body.ReadDouble(), body.ReadDouble(), body.ReadDouble());
            int count = body.ReadInt32();
            if (count < 0 || count > body.BaseStream.Length) throw new InvalidDataException("OpenDocument index: connected shapes damaged.");
            var glue = new OdtConnectedGluePoint[count];
            for (int k = 0; k < count; k++)
                glue[k] = new OdtConnectedGluePoint(body.ReadInt32(), new OdtConnectorGluePoint(body.ReadDouble(), body.ReadDouble(), (OdtGlueEscape)body.ReadByte()));
            shapes[ids[i]] = new OdtConnectedShape(keys[i], snap, glue);
        }
        return new OdtConnectedShapeTable(shapes);
    }
}
