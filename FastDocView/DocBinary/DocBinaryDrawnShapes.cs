using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>A .doc drawing object anchored in the text: its frame placement and drawn shapes, and the CP range of each shape's text box (by <see cref="WordShape.TextIndex"/>).</summary>
internal sealed record DocBinaryDrawing(WordPictureLayout Layout, List<WordShape> Shapes, List<DocBinaryShapeText> TextRanges, bool IsPlainTextBox, bool GrowsWithText);

/// <summary>A shape's text: a text-box story's CP range, or WordArt text held in the shape itself (gtextUNICODE) with its font and size.</summary>
internal readonly record struct DocBinaryShapeText(int Start, int End, string? Literal = null, string? FontFamily = null, double FontSizePt = 0, string? Color = null);

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
        /// <summary>Complex properties (FOPT id bit 0x8000): where their data lies in the Table stream and how long it is.</summary>
        public readonly Dictionary<int, (long at, int length)> Complex = new();
        public (int left, int top, int right, int bottom)? ChildAnchor, GroupRect;
        public bool IsGroup => (Flags & 0x1) != 0;
    }

    /// <summary>The drawing whose container is at [<paramref name="start"/>, <paramref name="end"/>) (an SpContainer body, or an SpgrContainer body when <paramref name="isGroup"/>), in a frame of the FSPA's size.</summary>
    public static DocBinaryDrawing Read(Stream table, long start, long end, bool isGroup, WordPictureLayout layout, DocBinaryTextBoxStories stories, DocBinaryFloatingPictures blips)
    {
        var shapes = new List<WordShape>();
        var texts = new List<DocBinaryShapeText>();
        ShapeRecord root;
        if (isGroup)
        {
            root = ReadGroup(table, start, end, 0, 0, layout.WidthPt, layout.HeightPt, shapes, texts, stories, blips, depth: 0) ?? new ShapeRecord();
        }
        else
        {
            root = ReadShape(table, start, end);
            AddShape(table, root, 0, 0, layout.WidthPt, layout.HeightPt, shapes, texts, stories, blips);
        }
        bool isPlain = !isGroup && shapes.Count == 1 && shapes[0].Geometry == WordShapeGeometry.Rectangle && shapes[0].TextIndex >= 0
                       && root.Type is WordOfficeArtShapeTypes.TextBox or WordOfficeArtShapeTypes.Rectangle
                       && shapes[0].RotationDegrees == 0 && !shapes[0].FlipHorizontal && !shapes[0].FlipVertical
                       && shapes[0].FillPaint is null && shapes[0].FillPictureId is null;
        bool grows = !isGroup && root.Properties.TryGetValue(PropertyFitText, out uint fit) && (fit & 2) != 0;
        return new DocBinaryDrawing(layout, shapes, texts, isPlain, grows);
    }

    /// <summary>
    /// The shape's pWrapPolygonVertices (899): an IMsoArray of x, y points in Word's 21600 space over the object, 8-byte int32
    /// pairs or 4-byte int16 pairs (0xFFF0 = truncated 8-byte) - read as LibreOffice reads it (msdffimp.cxx:5574-5605). A group's
    /// own shape is the first SpContainer of its SpgrContainer. Null when it has none or fewer than 3 points.
    /// </summary>
    public static List<(double X, double Y)>? WrapPolygon(Stream table, long start, long end, bool isGroup)
    {
        if (isGroup)
        {
            if (DocBinaryPictures.RecordHeader(table, start) is not { Type: 0xF004 } first) return null;
            end = Math.Min(end, start + 8 + first.Length);
            start += 8;
        }
        if (ReadArray(table, ReadShape(table, start, end), PropertyWrapPolygon) is not { } array) return null;
        int size = array.elementSize == 0xFFF0 ? 4 : array.elementSize;
        if (size is not (4 or 8)) return null;
        var points = new List<(double X, double Y)>(array.count);
        for (int i = 0; i < array.count && (i + 1) * size <= array.data.Length; i++)
            points.Add(size == 8 ? (BitConverter.ToInt32(array.data, i * 8), BitConverter.ToInt32(array.data, i * 8 + 4))
                                 : (BitConverter.ToInt16(array.data, i * 4), BitConverter.ToInt16(array.data, i * 4 + 2)));
        return points.Count > 2 ? points : null;
    }

    /// <summary>The text-box ranges of every shape in a drawing (the search: each story part once, where its shape is anchored).</summary>
    public static IEnumerable<DocBinaryShapeText> TextRanges(Stream table, long start, long end, bool isGroup, DocBinaryTextBoxStories stories)
    {
        var ranges = new List<DocBinaryShapeText>();
        void Walk(long from, long to, int depth)
        {
            for (long position = from; position + 8 <= to;)
            {
                if (DocBinaryPictures.RecordHeader(table, position) is not { } record) return;
                long body = position + 8, bodyEnd = Math.Min(to, body + record.Length);
                if (record.Type == 0xF003 && depth < 16) Walk(body, bodyEnd, depth + 1);
                else if (record.Type == 0xF004 && TextOf(table, ReadShape(table, body, bodyEnd), stories) is { } text)
                    ranges.Add(text);
                position = body + record.Length;
            }
        }
        if (isGroup) Walk(start, end, 0);
        else if (TextOf(table, ReadShape(table, start, end), stories) is { } text) ranges.Add(text);
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
                    // complex data follows the entries in their order (dffpropset.cxx:1111 nComplexDataFilePos, :1162-1214); an
                    // IMsoArray whose size counts only its elements is 6 bytes longer (its header, :1178-1194)
                    long complexAt = body + entries.Length, recordEnd = body + record.Length;
                    for (int at = 0; at + 6 <= entries.Length; at += 6)
                    {
                        int id = BitConverter.ToUInt16(entries, at), number = id & 0x3FFF;
                        uint value = BitConverter.ToUInt32(entries, at + 2);
                        shape.Properties[number] = value;
                        if ((id & 0x8000) == 0 || value == 0) continue;
                        long length = value;
                        if (IsArray(number) && DocBinaryPictures.Read(table, complexAt, 6) is { Length: 6 } header)
                        {
                            short elements = BitConverter.ToInt16(header, 0), reserved = BitConverter.ToInt16(header, 2), size = BitConverter.ToInt16(header, 4);
                            if (reserved >= elements)
                            {
                                if (size < 0) size = (short)(-size >> 2);
                                if ((long)size * elements == length) length += 6;
                                if (complexAt + length > recordEnd) length = 0;
                            }
                            else length = 0;
                        }
                        if (length == 0) { shape.Properties.Remove(number); continue; }
                        shape.Complex[number] = (complexAt, (int)Math.Min(int.MaxValue, length));
                        complexAt += length;
                    }
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
                                          List<WordShape> shapes, List<DocBinaryShapeText> texts, DocBinaryTextBoxStories stories, DocBinaryFloatingPictures? blips, int depth)
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
                    AddShape(table, shape, cx, cy, cw, ch, shapes, texts, stories, blips);
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
                    ReadGroup(table, body, bodyEnd, cx, cy, cw, ch, shapes, texts, stories, blips, depth + 1);
                }
            }
            position = body + record.Length;
        }
        if (group is not null)
            WordShapeGroups.Turn(shapes, first, x, y, width, height, Rotation(group), (group.Flags & 0x40) != 0, (group.Flags & 0x80) != 0);
        return group;
    }

    /// <summary>The IMsoArray properties LibreOffice checks for a missing header (dffpropset.cxx:1170-1176): pVertices 325, pSegmentInfo 326,
    /// fillShadeColors 407, lineDashStyle 463, pWrapPolygonVertices 899, connectorPoints 337, Handles 341, pFormulas 342, textRectangles 343 (include/svx/msdffdef.hxx).</summary>
    private static bool IsArray(int id) => id is 325 or 326 or 407 or 463 or 899 or 337 or 341 or 342 or 343;

    private const int PropertyWrapPolygon = 899, PropertyVertices = 325, PropertySegments = 326, PropertyGuides = 342, PropertyWordArtText = 192, PropertyWordArtSize = 195, PropertyWordArtFont = 197;

    /// <summary>
    /// The shape's own geometry as LibreOffice reads it (msdffimp.cxx:1625-1636 the coordinate space geoLeft..geoBottom, default
    /// 0, 0, 21600, 21600; :1956-1980 pFormulas - 8-byte records u16 flags, i16 p1, p2, p3; :2207-2265 pVertices - 8-byte int32
    /// pairs, or 4-byte int16 pairs (0xFFF0 = truncated 8-byte), unsigned for an arc (19); :2267-2390 pSegmentInfo u16 each).
    /// </summary>
    private static List<WordShapePath>? CustomGeometry(Stream table, ShapeRecord record, int?[] adjust, double width, double height)
    {
        if (ReadArray(table, record, PropertyVertices) is not { } vertices) return null;
        int elementSize = vertices.elementSize == 0xFFF0 ? 4 : vertices.elementSize;
        var values = new List<int>();
        for (int i = 0; i < vertices.count; i++)
        {
            int at = i * elementSize;
            if (elementSize == 8 && at + 8 <= vertices.data.Length) { values.Add(BitConverter.ToInt32(vertices.data, at)); values.Add(BitConverter.ToInt32(vertices.data, at + 4)); }
            else if (elementSize == 4 && at + 4 <= vertices.data.Length)
            {
                if (record.Type == 19) { values.Add(BitConverter.ToUInt16(vertices.data, at)); values.Add(BitConverter.ToUInt16(vertices.data, at + 2)); }
                else { values.Add(BitConverter.ToInt16(vertices.data, at)); values.Add(BitConverter.ToInt16(vertices.data, at + 2)); }
            }
        }
        var segments = new List<ushort>();
        if (ReadArray(table, record, PropertySegments) is { } segmentArray)
            for (int i = 0; i < segmentArray.count && i * 2 + 2 <= segmentArray.data.Length; i++) segments.Add(BitConverter.ToUInt16(segmentArray.data, i * 2));
        var guides = new List<int>();
        if (ReadArray(table, record, PropertyGuides) is { } guideArray && guideArray.count <= 128)
            for (int i = 0; i < guideArray.count && i * 8 + 8 <= guideArray.data.Length; i++)
            {
                guides.Add(BitConverter.ToUInt16(guideArray.data, i * 8));
                for (int k = 1; k <= 3; k++) guides.Add(BitConverter.ToInt16(guideArray.data, i * 8 + k * 2));
            }
        Dictionary<int, uint> p = record.Properties;
        int Geo(int id, int fallback) => p.TryGetValue(id, out uint v) ? (int)v : fallback;
        return WordOfficeArtGeometry.Custom(values, segments, guides.ToArray(), adjust, Geo(320, 0), Geo(321, 0), Geo(322, 21600), Geo(323, 21600), width, height);
    }

    /// <summary>An IMsoArray: element count, element size and the elements' bytes.</summary>
    private static (int count, int elementSize, byte[] data)? ReadArray(Stream table, ShapeRecord record, int id)
    {
        if (!record.Complex.TryGetValue(id, out var where) || where.length < 6) return null;
        byte[] bytes = DocBinaryPictures.Read(table, where.at, where.length);
        if (bytes.Length < 6) return null;
        return (BitConverter.ToUInt16(bytes, 0), BitConverter.ToUInt16(bytes, 4), bytes.AsSpan(6).ToArray());
    }

    private static string? ComplexString(Stream table, ShapeRecord record, int id)
    {
        if (!record.Complex.TryGetValue(id, out var where) || where.length < 2) return null;
        byte[] bytes = DocBinaryPictures.Read(table, where.at, where.length);
        string text = System.Text.Encoding.Unicode.GetString(bytes, 0, bytes.Length & ~1);
        int end = text.IndexOf('\0');
        return end >= 0 ? text.Substring(0, end) : text;
    }

    /// <summary>
    /// A shape's text: its text-box story (lTxid), else its WordArt string (gtextUNICODE 192 with gtextFont 197 and gtextSize 195,
    /// 16.16 points - msdffdef.hxx:121-125; LibreOffice puts that string in the shape as its text, msdffimp.cxx ImportGraphic /
    /// rtfsdrimport.cxx:949-957 alike).
    /// </summary>
    private static DocBinaryShapeText? TextOf(Stream table, ShapeRecord record, DocBinaryTextBoxStories stories)
    {
        if (record.Properties.TryGetValue(PropertyTextId, out uint id) && stories.RangeOf(id) is { } range) return new DocBinaryShapeText(range.start, range.end);
        if (ComplexString(table, record, PropertyWordArtText) is { Length: > 0 } wordArt)
            return new DocBinaryShapeText(0, 0, wordArt, ComplexString(table, record, PropertyWordArtFont),
                                          record.Properties.TryGetValue(PropertyWordArtSize, out uint size) ? size / 65536.0 : 36);
        return null;
    }

    /// <summary>The fill properties (msdffdef.hxx:173-200: fillType 384, fillColor 385, fillOpacity 386, fillBackColor 387, fillBackOpacity 388,
    /// fillBlip 390, fillWidth / fillHeight 393 / 394 (EMU), fillAngle 395, fillFocus 396, fillToRight / fillToBottom 399 / 400,
    /// fillShadeColors 407).</summary>
    private static WordOfficeArtFillProperties FillOf(Stream table, ShapeRecord record, bool stroked, DocBinaryFloatingPictures? blips)
    {
        Dictionary<int, uint> p = record.Properties;
        uint? Get(int id) => p.TryGetValue(id, out uint v) ? v : null;
        var fill = new WordOfficeArtFillProperties
        {
            Type = (int)(Get(384) ?? 0),
            Color = Color(p, PropertyFillColor, 0xFFFFFF, stroked),
            BackColor = Color(p, 387, 0xFFFFFF, stroked),
            HasOpacity = p.ContainsKey(PropertyFillOpacity),
            Opacity = (Get(PropertyFillOpacity) ?? 0x10000) / 65536.0,
            BackOpacity = (Get(388) ?? 0x10000) / 65536.0,
            AngleFix16 = (int)(Get(395) ?? 0),
            Focus = (int)(Get(396) ?? 0),
            ToRight = (int)(Get(399) ?? 0),
            ToBottom = (int)(Get(400) ?? 0),
            TileWidthPt = (int)(Get(393) ?? 0) / 12700.0,
            TileHeightPt = (int)(Get(394) ?? 0) / 12700.0,
        };
        if (ReadArray(table, record, 407) is { } shades && shades.elementSize == 8)
        {
            fill.ShadeColors = new List<(string, double)>();
            for (int i = 0; i < shades.count && i * 8 + 8 <= shades.data.Length; i++)
                fill.ShadeColors.Add((WordOfficeArtColor.Css(BitConverter.ToUInt32(shades.data, i * 8), PropertyFillColor, id => p.TryGetValue(id, out uint v) ? v : null, stroked),
                                      1 - BitConverter.ToInt32(shades.data, i * 8 + 4) / 65536.0));
        }
        if (Get(390) is { } blip && blips is not null)
        {
            if (fill.Type == 1 && blips.Dib((int)blip) is { } dib && WordOfficeArtFill.DibPixels(dib) is { } pattern)
                (fill.PatternPixels, fill.PatternWidth, fill.PatternHeight) = pattern;
            else fill.PictureId = blips.FillPictureId((int)blip);
        }
        return fill;
    }

    private static double Rotation(ShapeRecord shape) => shape.Properties.TryGetValue(PropertyRotation, out uint value) ? (int)value / 65536.0 : 0;

    private static void AddShape(Stream table, ShapeRecord record, double x, double y, double width, double height, List<WordShape> shapes, List<DocBinaryShapeText> texts, DocBinaryTextBoxStories stories, DocBinaryFloatingPictures? blips)
    {
        Dictionary<int, uint> p = record.Properties;
        int type = record.Type < 0 ? WordOfficeArtShapeTypes.Rectangle : record.Type;
        bool filled = Flag(p, PropertyFillFlags, 0x10) ?? WordOfficeArtShapeTypes.IsFilledByDefault(type);
        bool stroked = Flag(p, PropertyLineFlags, 0x08) ?? WordOfficeArtShapeTypes.IsStrokedByDefault(type);
        WordOfficeArtFillProperties? fill = filled ? FillOf(table, record, stroked, blips) : null;
        var shape = new WordShape
        {
            X = x, Y = y, Width = width, Height = height,
            RotationDegrees = Rotation(record), FlipHorizontal = (record.Flags & 0x40) != 0, FlipVertical = (record.Flags & 0x80) != 0,
            FillColor = filled ? Color(p, PropertyFillColor, 0xFFFFFF, stroked) : null,
            LineColor = stroked ? Color(p, PropertyLineColor, 0, stroked) : null,
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
        if (fill is not null) WordOfficeArtFill.Apply(shape, fill);
        if (record.Complex.ContainsKey(PropertyVertices) && CustomGeometry(table, record, adjust, width, height) is { } custom)
        {
            shape.Geometry = WordShapeGeometry.Path;
            shape.Paths = custom;
        }
        // line ends only when fArrowheadsOK (line booleans 0x10, set by default) - msdffimp.cxx:1063
        if (stroked && (Flag(p, PropertyLineFlags, 0x10) ?? true))
        {
            int? Get(int id) => p.TryGetValue(id, out uint v) ? (int)v : null;
            WordOfficeArtShapeTypes.ApplyLineEnds(shape, Get(PropertyStartArrow), Get(PropertyStartArrow + 2), Get(PropertyStartArrow + 3),
                                                  Get(PropertyEndArrow), Get(PropertyEndArrow + 3), Get(PropertyEndArrow + 4));
        }
        if (p.ContainsKey(PropertyPib)) { shape.Geometry = WordShapeGeometry.Picture; shape.PictureId = DocBinaryFloatingPictures.IdPrefix + record.Spid.ToString(System.Globalization.CultureInfo.InvariantCulture); }
        // a text effect (fGtext 0x4000 of the gtext booleans 255, msdffimp.cxx:2517): LibreOffice draws Fontwork in place of the
        // outline (svx EnhancedCustomShapeEngine.cxx:236-262), the text filled with the shape's fill colour
        bool textEffect = Flag(p, 255, 0x4000) ?? (p.TryGetValue(255, out uint gtextFlags) && (gtextFlags & 0x4000) != 0);
        DocBinaryShapeText? shapeText = TextOf(table, record, stories);
        if (textEffect && shapeText?.Literal is { } wordArt) WordOfficeArtFontwork.Apply(shape, wordArt, shapeText.Value.FontFamily, p.TryGetValue(194, out uint align) ? (int)align : 1);   // gtextAlign 194
        else if (shapeText is { } text)
        {
            shape.TextIndex = texts.Count;
            texts.Add(text);
        }
        if (textEffect && shapeText?.Literal is null) { shape.Geometry = WordShapeGeometry.Path; shape.Paths = []; }
        shapes.Add(shape);
    }

    /// <summary>A boolean of a group set explicitly (its "use" bit in the high word), or null for the default.</summary>
    private static bool? Flag(Dictionary<int, uint> properties, int group, uint bit)
        => properties.TryGetValue(group, out uint value) && (value & (bit << 16)) != 0 ? (value & bit) != 0 : null;

    private static double Emu(Dictionary<int, uint> properties, int id, int fallback) => (properties.TryGetValue(id, out uint value) ? (int)value : fallback) / 12700.0;

    private static string Color(Dictionary<int, uint> properties, int id, uint fallback, bool hasLine)
        => WordOfficeArtColor.Css(properties.TryGetValue(id, out uint set) ? set : fallback, id, other => properties.TryGetValue(other, out uint value) ? value : null, hasLine);
}
