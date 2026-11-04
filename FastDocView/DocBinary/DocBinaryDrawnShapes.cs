using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>A .doc drawing object anchored in the text: its frame placement and drawn shapes, and the CP range of each shape's text box (by <see cref="WordShape.TextIndex"/>).</summary>
internal sealed record DocBinaryDrawing(WordPictureLayout Layout, List<WordShape> Shapes, List<(int start, int end)> TextRanges, bool IsPlainTextBox, bool GrowsWithText);

/// <summary>
/// The OfficeArt shape records of a .doc drawing, read from the Table stream only when the block that anchors them renders or is
/// searched ([MS-ODRAW] record layout as LibreOffice reads it, filter/source/msfilter/msdffimp.cxx, master 2026-11-03):
/// - FSP 0xF00A: recInstance = the shape type (MSOSPT), body spid u32 then flags u32 (include/filter/msfilter/escherex.hxx:83-99
///   ShapeFlag: Group 0x1, Child 0x2, FlipH 0x40, FlipV 0x80).
/// - FOPT 0xF00B: recInstance entries of u16 id (id &amp; 0x3FFF, dffpropset.cxx:1128) and u32 value; ids from include/svx/msdffdef.hxx:
///   Rotation 4 (16.16 degrees), lTxid 128, dxTextLeft..dyTextBottom 129-132 (EMU), FitTextToShape 191 (bit 2 = grow, msdffimp.cxx:1598),
///   pib 260, adjustValue / adjust2Value 327 / 328, fillColor 385, fillOpacity 386 (16.16), fill booleans 447 (fFilled = 0x10),
///   lineColor 448, lineOpacity 449, lineWidth 459 (EMU), lineDashing 462, lineStart / EndArrowhead 464 / 465, line booleans 511
///   (fLine = 0x08). A boolean group's high word marks which bits are set explicitly (dffpropset.cxx:1146-1157); a bit not set
///   explicitly takes the shape type's default (msdffimp.cxx:908-911 / 1320-1323 with IsCustomShapeFilledByDefault /
///   IsCustomShapeStrokedByDefault, see <see cref="WordOfficeArtShapeTypes"/>). Defaults: fill white (msdffimp.cxx:1149), line
///   black 9525 EMU (msdffimp.cxx:916, 1039), text insets 25 * 3600 / 13 * 3600 EMU (msdffimp.cxx:1474-1477).
/// - Colours (MSO_CLR): red in the low byte (msdffimp.cxx:3420-3600); a scheme or system colour (upper byte &amp; 0x19) has no
///   palette in a Word document and takes the property's default (msdffimp.cxx:3433-3454; the system-colour branch reads the
///   desktop's colours, which a viewer cannot know - the default is used).
/// - Groups: SpgrContainer 0xF003, "the first of which is the group shape itself" (msdffdef.hxx:57) with its FSPGR 0xF009 rect
///   (its coordinate space); each child's ChildAnchor 0xF00F is a rect "in units relative to the parent group" (msdffdef.hxx:64).
/// </summary>
internal static class DocBinaryDrawnShapes
{
    private const int PropertyRotation = 4, PropertyTextId = 128, PropertyTextLeft = 129, PropertyTextTop = 130, PropertyTextRight = 131,
                      PropertyTextBottom = 132, PropertyFitText = 191, PropertyPib = 260, PropertyAdjust = 327, PropertyAdjust2 = 328,
                      PropertyFillColor = 385, PropertyFillOpacity = 386, PropertyFillFlags = 447, PropertyLineColor = 448,
                      PropertyLineOpacity = 449, PropertyLineWidth = 459, PropertyLineDashing = 462, PropertyStartArrow = 464,
                      PropertyEndArrow = 465, PropertyLineFlags = 511;

    private sealed class ShapeRecord
    {
        public uint Spid;
        public int Type = -1;
        public uint Flags;
        public readonly Dictionary<int, uint> Properties = new();
        public (int left, int top, int right, int bottom)? ChildAnchor, GroupRect;
        public bool IsGroup => (Flags & 0x1) != 0;
    }

    /// <summary>The drawing whose container is at [<paramref name="start"/>, <paramref name="end"/>) (an SpContainer body, or an SpgrContainer body when <paramref name="isGroup"/>), in a frame of the FSPA's size.</summary>
    public static DocBinaryDrawing Read(Stream table, long start, long end, bool isGroup, WordPictureLayout layout, DocBinaryTextBoxStories stories)
    {
        var shapes = new List<WordShape>();
        var texts = new List<(int, int)>();
        ShapeRecord root;
        if (isGroup)
        {
            root = ReadGroup(table, start, end, 0, 0, layout.WidthPt, layout.HeightPt, shapes, texts, stories, depth: 0) ?? new ShapeRecord();
        }
        else
        {
            root = ReadShape(table, start, end);
            AddShape(root, 0, 0, layout.WidthPt, layout.HeightPt, shapes, texts, stories);
        }
        bool isPlain = !isGroup && shapes.Count == 1 && shapes[0].Geometry == WordShapeGeometry.Rectangle && shapes[0].TextIndex >= 0
                       && root.Type is WordOfficeArtShapeTypes.TextBox or WordOfficeArtShapeTypes.Rectangle
                       && shapes[0].RotationDegrees == 0 && !shapes[0].FlipHorizontal && !shapes[0].FlipVertical;
        bool grows = !isGroup && root.Properties.TryGetValue(PropertyFitText, out uint fit) && (fit & 2) != 0;
        return new DocBinaryDrawing(layout, shapes, texts, isPlain, grows);
    }

    /// <summary>The text-box ranges of every shape in a drawing (the search: each story part once, where its shape is anchored).</summary>
    public static IEnumerable<(int start, int end)> TextRanges(Stream table, long start, long end, bool isGroup, DocBinaryTextBoxStories stories)
    {
        var ranges = new List<(int, int)>();
        void Walk(long from, long to, int depth)
        {
            for (long position = from; position + 8 <= to;)
            {
                if (DocBinaryPictures.RecordHeader(table, position) is not { } record) return;
                long body = position + 8, bodyEnd = Math.Min(to, body + record.Length);
                if (record.Type == 0xF003 && depth < 16) Walk(body, bodyEnd, depth + 1);
                else if (record.Type == 0xF004 && ReadShape(table, body, bodyEnd).Properties.TryGetValue(PropertyTextId, out uint id) && stories.RangeOf(id) is { } range)
                    ranges.Add(range);
                position = body + record.Length;
            }
        }
        if (isGroup) Walk(start, end, 0);
        else if (ReadShape(table, start, end).Properties.TryGetValue(PropertyTextId, out uint id) && stories.RangeOf(id) is { } range) ranges.Add(range);
        return ranges;
    }

    private static ShapeRecord ReadShape(Stream table, long position, long end)
    {
        var shape = new ShapeRecord();
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(table, position) is not { } record) break;
            long body = position + 8;
            switch (record.Type)
            {
                case 0xF00A when record.Length >= 8:
                {
                    byte[] fsp = DocBinaryPictures.Read(table, body, 8);
                    if (fsp.Length == 8) { shape.Spid = BitConverter.ToUInt32(fsp, 0); shape.Flags = BitConverter.ToUInt32(fsp, 4); }
                    shape.Type = record.Instance;
                    break;
                }
                case 0xF00B:
                {
                    byte[] entries = DocBinaryPictures.Read(table, body, (int)Math.Min(record.Length, (long)record.Instance * 6));
                    for (int at = 0; at + 6 <= entries.Length; at += 6)
                        shape.Properties[BitConverter.ToUInt16(entries, at) & 0x3FFF] = BitConverter.ToUInt32(entries, at + 2);
                    break;
                }
                case 0xF00F or 0xF009 when record.Length >= 16:
                {
                    byte[] rect = DocBinaryPictures.Read(table, body, 16);
                    if (rect.Length < 16) break;
                    var value = (BitConverter.ToInt32(rect, 0), BitConverter.ToInt32(rect, 4), BitConverter.ToInt32(rect, 8), BitConverter.ToInt32(rect, 12));
                    if (record.Type == 0xF00F) shape.ChildAnchor = value; else shape.GroupRect = value;
                    break;
                }
            }
            position = body + record.Length;
        }
        return shape;
    }

    /// <summary>A group's children placed in its box; returns the group's own shape record (its first SpContainer).</summary>
    private static ShapeRecord? ReadGroup(Stream table, long position, long end, double x, double y, double width, double height,
                                          List<WordShape> shapes, List<(int, int)> texts, DocBinaryTextBoxStories stories, int depth)
    {
        ShapeRecord? group = null;
        int first = shapes.Count;
        (double left, double top, double right, double bottom) space = (0, 0, 1, 1);
        (double x, double y, double width, double height) Place((int left, int top, int right, int bottom)? anchor)
        {
            if (anchor is not { } a || space.right == space.left || space.bottom == space.top) return (x, y, width, height);
            double scaleX = width / (space.right - space.left), scaleY = height / (space.bottom - space.top);
            return (x + (Math.Min(a.left, a.right) - space.left) * scaleX, y + (Math.Min(a.top, a.bottom) - space.top) * scaleY,
                    Math.Abs(a.right - a.left) * scaleX, Math.Abs(a.bottom - a.top) * scaleY);
        }
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(table, position) is not { } record) break;
            long body = position + 8, bodyEnd = Math.Min(end, body + record.Length);
            if (record.Type == 0xF004)
            {
                ShapeRecord shape = ReadShape(table, body, bodyEnd);
                if (group is null)
                {
                    group = shape;
                    if (shape.GroupRect is { } rect) space = (rect.left, rect.top, rect.right, rect.bottom);
                }
                else
                {
                    (double cx, double cy, double cw, double ch) = Place(shape.ChildAnchor);
                    AddShape(shape, cx, cy, cw, ch, shapes, texts, stories);
                }
            }
            else if (record.Type == 0xF003 && group is not null && depth < 16)
            {
                // a nested group: its own first SpContainer carries its ChildAnchor (its box here) and its FSPGR
                long firstAt = body;
                if (DocBinaryPictures.RecordHeader(table, firstAt) is { Type: 0xF004 } own)
                {
                    ShapeRecord nested = ReadShape(table, firstAt + 8, Math.Min(bodyEnd, firstAt + 8 + own.Length));
                    (double cx, double cy, double cw, double ch) = Place(nested.ChildAnchor);
                    ReadGroup(table, body, bodyEnd, cx, cy, cw, ch, shapes, texts, stories, depth + 1);
                }
            }
            position = body + record.Length;
        }
        if (group is not null)
            WordShapeGroups.Turn(shapes, first, x, y, width, height, Rotation(group), (group.Flags & 0x40) != 0, (group.Flags & 0x80) != 0);
        return group;
    }

    private static double Rotation(ShapeRecord shape) => shape.Properties.TryGetValue(PropertyRotation, out uint value) ? (int)value / 65536.0 : 0;

    private static void AddShape(ShapeRecord record, double x, double y, double width, double height, List<WordShape> shapes, List<(int, int)> texts, DocBinaryTextBoxStories stories)
    {
        Dictionary<int, uint> p = record.Properties;
        int type = record.Type < 0 ? WordOfficeArtShapeTypes.Rectangle : record.Type;
        bool filled = Flag(p, PropertyFillFlags, 0x10) ?? WordOfficeArtShapeTypes.IsFilledByDefault(type);
        bool stroked = Flag(p, PropertyLineFlags, 0x08) ?? WordOfficeArtShapeTypes.IsStrokedByDefault(type);
        var shape = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = Rotation(record), FlipHorizontal = (record.Flags & 0x40) != 0, FlipVertical = (record.Flags & 0x80) != 0,
            FillColor = filled ? Color(p, PropertyFillColor, 0xFFFFFF) : null,
            LineColor = stroked ? Color(p, PropertyLineColor, 0) : null,
            LineWidthPt = (p.TryGetValue(PropertyLineWidth, out uint lineWidth) ? (int)lineWidth : 9525) / 12700.0,
            Dash = p.TryGetValue(PropertyLineDashing, out uint dashing) ? WordOfficeArtShapeTypes.Dash((int)dashing) : null,
            FillOpacity = p.TryGetValue(PropertyFillOpacity, out uint fillOpacity) ? fillOpacity / 65536.0 : 1,
            LineOpacity = p.TryGetValue(PropertyLineOpacity, out uint lineOpacity) ? lineOpacity / 65536.0 : 1,
            TextStyle = new WordTextBoxStyle(Emu(p, PropertyTextLeft, 25 * 3600), Emu(p, PropertyTextTop, 13 * 3600), Emu(p, PropertyTextRight, 25 * 3600),
                                             Emu(p, PropertyTextBottom, 13 * 3600), null, 0, false, null, 0, null,
                                             p.TryGetValue(PropertyFitText, out uint fit) && (fit & 2) != 0),
        };
        var adjust = new int?[10];
        for (int i = 0; i < 10; i++) adjust[i] = p.TryGetValue(PropertyAdjust + i, out uint value) ? (int)value : null;
        WordOfficeArtShapeTypes.Apply(shape, type, adjust);
        // line ends only when fArrowheadsOK (line booleans 0x10, set by default) - msdffimp.cxx:1063
        if (stroked && (Flag(p, PropertyLineFlags, 0x10) ?? true))
        {
            int? Get(int id) => p.TryGetValue(id, out uint v) ? (int)v : null;
            WordOfficeArtShapeTypes.ApplyLineEnds(shape, Get(PropertyStartArrow), Get(PropertyStartArrow + 2), Get(PropertyStartArrow + 3),
                                                  Get(PropertyEndArrow), Get(PropertyEndArrow + 3), Get(PropertyEndArrow + 4));
        }
        if (p.ContainsKey(PropertyPib)) { shape.Geometry = WordShapeGeometry.Picture; shape.PictureId = DocBinaryFloatingPictures.IdPrefix + record.Spid.ToString(System.Globalization.CultureInfo.InvariantCulture); }
        if (p.TryGetValue(PropertyTextId, out uint textId) && stories.RangeOf(textId) is { } range) { shape.TextIndex = texts.Count; texts.Add(range); }
        shapes.Add(shape);
    }

    /// <summary>A boolean of a group set explicitly (its "use" bit in the high word), or null for the default.</summary>
    private static bool? Flag(Dictionary<int, uint> properties, int group, uint bit)
        => properties.TryGetValue(group, out uint value) && (value & (bit << 16)) != 0 ? (value & bit) != 0 : null;

    private static double Emu(Dictionary<int, uint> properties, int id, int fallback) => (properties.TryGetValue(id, out uint value) ? (int)value : fallback) / 12700.0;

    private static string Color(Dictionary<int, uint> properties, int id, uint fallback)
    {
        uint value = properties.TryGetValue(id, out uint set) ? set : fallback;
        if ((value & 0xFE000000) == 0xFE000000) value &= 0x00FFFFFF;
        if (((value >> 24) & 0x19) != 0) value = fallback;
        return "#" + (value & 0xFF).ToString("x2") + ((value >> 8) & 0xFF).ToString("x2") + ((value >> 16) & 0xFF).ToString("x2");
    }
}
