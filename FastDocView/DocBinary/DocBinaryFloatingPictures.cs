using FastDocView.Html;

namespace FastDocView.DocBinary;

/// <summary>
/// Floating pictures of a .doc: a 0x08 character (fSpec) in the main text anchors a shape (LibreOffice ww8par.cxx:3656
/// "case 0x8: Read_GrafLayer"). Read as LibreOffice reads it (ww8graf.cxx:2605-2662 SwWW8ImplReader::Read_GrafLayer):
/// - PlcSpaMom (Table stream, FibRgFcLcb97 index 40; ww8scan.cxx:6117-6118 read order): a PLCF of CPs and 26-byte FSPAs
///   (ww8scan.cxx:1808-1811 "WW8PLCFspecial( pTableSt, fcPlcfspaMom, lcbPlcfspaMom, 26 )"), the entry at the anchor CP.
///   FSPA (ww8struc.hxx:903-964 WW8_FSPA / WW8_FSPA_SHADOW): spid, xaLeft, yaTop, xaRight, yaBottom (Int32 twips), then
///   flags u16: bHdr 0x0001, bx 0x0006, by 0x0018, wr 0x01E0, wrk 0x1E00, fRcaSimple 0x2000, fBelowText 0x4000.
/// - The drawing (Table stream at fcDggInfo, FibRgFcLcb97 index 50; ww8par.cxx:501 passes it to SvxMSDffManager with the
///   Data stream as the blip stream and the WordDocument stream as the second one): a DggContainer 0xF000 whose
///   BStoreContainer 0xF001 lists FBSEs 0xF007 (msdffimp.cxx:6020-6095: blip size 20 bytes into the FBSE body, foDelay 4
///   after it; foDelay 0 and a size smaller than the record = the blip is inside the FBSE), then DgContainers 0xF002, each
///   behind a one-byte label LibreOffice steps over (msdffimp.cxx:5990-6006 "nPos++"). A shape's SpContainer 0xF004 holds
///   its FSP 0xF00A (spid = first u32, msdffimp.cxx:6221-6226) and its FOPT 0xF00B: 6-byte entries u16 id / u32 value
///   (msdffimp.cxx:6238-6243), the property number being id &amp; 0x3FFF (dffpropset.cxx:1128); pib = 260
///   (msdffdef.hxx DFF_Prop_pib), the 1-based FBSE index (msdffimp.cxx:3843 / 6439-6455).
/// - The blip at foDelay: the Data stream first, then the WordDocument stream (msdffimp.cxx:6455-6473 GetBLIP).
/// Wrap from FSPA wr / wrk as ww8graf.cxx:2700-2745 maps them: 1 = text above and below, 3 = no wrap (through),
/// 0 / 2 / 4 / 5 = text on the sides; wrk 1 = text only on the left (the picture on the right); 4 / 5 along a contour. Size = xaRight - xaLeft by
/// yaBottom - yaTop. The browser has no page, so yaTop and the page / margin origin (bx, by) are not drawn; a picture
/// relative to the column (bx 2) or margin (bx 0) keeps xaLeft as its offset from the start of the line.
/// Only the main text's shapes (headers are not rendered here). A group, a shape without a picture or one with a text box
/// (lTxid) is a drawing: <see cref="Drawing"/> reads its records then (<see cref="DocBinaryDrawnShapes"/>).
/// The tables are read on the first floating picture: CPs (4 bytes a shape), spid -> pib, pib -> blip offset, spid -> the
/// shape's container offsets (a group: its SpgrContainer); no bytes, no shape properties.
/// </summary>
internal sealed class DocBinaryFloatingPictures
{
    private const int FspaSize = 26;
    private const int PropertyPib = 260;
    /// <summary>The id prefix that tells a floating picture from an inline one (whose id is a PICF offset, digits only).</summary>
    public const string IdPrefix = "s";

    private readonly Stream _tableStream;
    private readonly Stream _wordDocument;
    private readonly Stream? _dataStream;
    private readonly (uint offset, uint length) _plcSpa;
    private readonly (uint offset, uint length) _drawing;

    private int[]? _anchorCps;
    private Dictionary<uint, int>? _pibBySpid;
    private List<(uint offset, bool insideTable)>? _blips;
    private Dictionary<uint, (long start, long end, bool isGroup)>? _containerBySpid;
    private HashSet<uint>? _textBoxSpids;
    private readonly DocBinaryTextBoxStories _stories;

    /// <param name="isHeader">
    /// The header / footer stories' shapes (PlcSpaHdr, anchored at CPs counted from the header story's start - LibreOffice
    /// ww8par.cxx:3656 passes the story CP to Read_GrafLayer, ww8graf.cxx:2605-2620 seeks the header PLCF with it - with the header
    /// text-box story)
    /// instead of the main text's. The flow view does not show headers, so nothing renders these; the paged view reads them
    /// through <see cref="AnchorCps"/> / <see cref="Drawing"/> / <see cref="At"/> - they never reach the body or its search.
    /// </param>
    public DocBinaryFloatingPictures(Stream tableStream, Stream wordDocument, Stream? dataStream, DocBinaryFileInformationBlock fib, bool isHeader = false)
    {
        _tableStream = tableStream;
        _wordDocument = wordDocument;
        _dataStream = dataStream;
        _plcSpa = isHeader ? fib.HeaderShapeAnchors : fib.MainShapeAnchors;
        _drawing = fib.DrawingGroup;
        _stories = new DocBinaryTextBoxStories(tableStream, fib, isHeader);
    }

    /// <summary>The drawing anchored at <paramref name="cp"/> when it is not a lone picture <see cref="At"/> shows: shapes, groups, text boxes.</summary>
    public DocBinaryDrawing? Drawing(int cp)
    {
        if (Fspa(cp) is not { } fspa) return null;
        if (_pibBySpid is null) ReadDrawing();
        if (!_containerBySpid!.TryGetValue(fspa.spid, out var container)) return null;
        DocBinaryDrawing drawing = DocBinaryDrawnShapes.Read(_tableStream, container.start, container.end, container.isGroup, LayoutOf(fspa), _stories, this);
        return drawing with { Layout = WithContour(drawing.Layout, fspa.flags, container, () => WordShapeHtml.Outlines(drawing.Shapes)) };
    }

    /// <summary>The CPs that anchor a shape (main text or header story CPs, as this table was made for).</summary>
    public IReadOnlyList<int> AnchorCps
    {
        get
        {
            int[] cps = _anchorCps ??= ReadAnchorCps();
            return cps.Length < 2 ? [] : new ArraySegment<int>(cps, 0, cps.Length - 1);
        }
    }

    /// <summary>The text-box CP ranges of the drawings anchored in [<paramref name="startCp"/>, <paramref name="endCp"/>) (the search).</summary>
    public IEnumerable<DocBinaryShapeText> TextRangesIn(int startCp, int endCp)
    {
        int[] cps = _anchorCps ??= ReadAnchorCps();
        if (cps.Length < 2) yield break;
        int index = Array.BinarySearch(cps, 0, cps.Length - 1, startCp);
        if (index < 0) index = ~index;
        for (; index < cps.Length - 1 && cps[index] < endCp; index++)
        {
            if (Fspa(cps[index]) is not { } fspa) continue;
            if (_pibBySpid is null) ReadDrawing();
            if (!_containerBySpid!.TryGetValue(fspa.spid, out var container)) continue;
            foreach (DocBinaryShapeText range in DocBinaryDrawnShapes.TextRanges(_tableStream, container.start, container.end, container.isGroup, _stories))
                yield return range;
        }
    }

    private static WordPictureLayout LayoutOf((uint spid, int xaLeft, int yaTop, int xaRight, int yaBottom, int flags) fspa)
    {
        double widthPt = Math.Max(1, fspa.xaRight - (double)fspa.xaLeft) / 20, heightPt = Math.Max(1, fspa.yaBottom - (double)fspa.yaTop) / 20;
        int bx = (fspa.flags >> 1) & 3, wr = (fspa.flags >> 5) & 0xF, wrk = (fspa.flags >> 9) & 0xF;
        bool pageRelative = bx == 1 || (fspa.flags & 0x2000) != 0;
        WordPictureWrap wrap = wr switch { 1 => WordPictureWrap.TopAndBottom, 3 => WordPictureWrap.None, _ => WordPictureWrap.Square };
        WordPictureSide side = (wr is 2 or 4 && wrk == 1) ? WordPictureSide.Right : WordPictureSide.Left;
        double offsetPt = side == WordPictureSide.Left && !pageRelative ? Math.Max(0, fspa.xaLeft / 20.0) : 0;
        // wrk only counts for wr 2 / 4 (ww8graf.cxx:2730-2749): 1 = text only on the left, 2 = only on the right, 0 / 3 both / largest
        WordPictureTextSides sides = wr is 2 or 4 ? wrk switch { 1 => WordPictureTextSides.LeftOnly, 2 => WordPictureTextSides.RightOnly, _ => WordPictureTextSides.Both }
                                                  : WordPictureTextSides.Both;
        return new WordPictureLayout(widthPt, heightPt, wrap, side, 0, 0, 0, 0, offsetPt, sides);
    }

    /// <summary>
    /// A two-sided wrap floats to the side its alignment puts it. wr 4 / 5 (tight / through) wrap along a contour (ww8graf.cxx:2721-2725 bContour): the shape's pWrapPolygonVertices
    /// (<see cref="DocBinaryDrawnShapes.WrapPolygon"/>, ww8graf.cxx:2067-2128 scales it by the 21600 space), else
    /// <paramref name="outline"/> (a drawing's own shapes; a picture: its box, the layout as it is).
    /// </summary>
    private WordPictureLayout WithContour(WordPictureLayout layout, int flags, (long start, long end, bool isGroup) container, Func<IReadOnlyList<IReadOnlyList<(double X, double Y)>>?>? outline)
    {
        if (layout.Wrap == WordPictureWrap.Square && layout.TextSides == WordPictureTextSides.Both
            && DocBinaryDrawnShapes.HorizontalAlignment(_tableStream, container.start, container.end, container.isGroup) is var align and > 0)
            // placed by alignment (posh 1 left, 2 centre, 3 right, 4 inside, 5 outside - odd page taken): the side it floats to
            layout = layout with { Side = align is 3 or 5 ? WordPictureSide.Right : align == 2 ? WordPictureSide.Center : WordPictureSide.Left, OffsetPt = 0 };
        if ((flags >> 5 & 0xF) is not (4 or 5) || layout.Wrap != WordPictureWrap.Square) return layout;
        var contour = WordPictureContour.FromWordUnits(DocBinaryDrawnShapes.WrapPolygon(_tableStream, container.start, container.end, container.isGroup), layout.WidthPt, layout.HeightPt)
                      ?? outline?.Invoke();
        return contour is { Count: > 0 } ? layout with { Contour = contour } : layout;
    }

    /// <summary>The layout and id of the picture anchored at <paramref name="cp"/>; null when there is none we can show.</summary>
    public (WordPictureLayout layout, string pictureId)? At(int cp)
    {
        if (Fspa(cp) is not { } fspa) return null;
        if (Locate(fspa.spid) is null) return null;
        bool located = _containerBySpid!.TryGetValue(fspa.spid, out var container);
        if (located && container.isGroup) return null;                                                        // a group is a drawing
        if (_textBoxSpids!.Contains(fspa.spid)) return null;                                                  // a picture with a text box too
        return (located ? WithContour(LayoutOf(fspa), fspa.flags, container, null) : LayoutOf(fspa), IdPrefix + fspa.spid.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>The id prefix of a fill picture: the BStore index (1-based) of a shape's fillBlip.</summary>
    public const string FillIdPrefix = "sf";

    /// <summary>The picture id of BStore entry <paramref name="pib"/> as a fill (texture / picture fill); null when it is not a PNG / JPEG.</summary>
    internal string? FillPictureId(int pib)
        => LocateBlip(pib) is not null ? FillIdPrefix + pib.ToString(System.Globalization.CultureInfo.InvariantCulture) : null;

    /// <summary>The DIB of BStore entry <paramref name="pib"/> (a pattern fill's blip), or null.</summary>
    internal byte[]? Dib(int pib)
    {
        if (_pibBySpid is null) ReadDrawing();
        if (pib < 1 || pib > _blips!.Count) return null;
        (uint offset, bool insideTable) = _blips[pib - 1];
        IEnumerable<Stream?> streams = insideTable ? [_tableStream] : [_dataStream, _wordDocument];
        foreach (Stream? stream in streams)
            if (stream is not null && DocBinaryPictures.DibBlip(stream, offset) is { } dib)
                return DocBinaryPictures.Read(stream, dib.dataAt, dib.dataLength);
        return null;
    }

    /// <summary>The bytes of floating picture <paramref name="pictureId"/> (see <see cref="At"/>), or null.</summary>
    public WordDocumentPicture? Read(string pictureId)
    {
        if (pictureId.StartsWith(FillIdPrefix, StringComparison.Ordinal))
        {
            if (!int.TryParse(pictureId.Substring(FillIdPrefix.Length), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int pib)
                || LocateBlip(pib) is not { } fill) return null;
            byte[] fillBytes = DocBinaryPictures.Read(fill.stream, fill.dataAt, fill.dataLength);
            return fillBytes.Length == fill.dataLength ? WordPictureHtml.Picture(fill.extension, fillBytes) : null;
        }
        if (!pictureId.StartsWith(IdPrefix, StringComparison.Ordinal)
            || !uint.TryParse(pictureId.Substring(IdPrefix.Length), System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out uint spid))
            return null;
        if (Locate(spid) is not { } blip) return null;
        byte[] bytes = DocBinaryPictures.Read(blip.stream, blip.dataAt, blip.dataLength);
        return bytes.Length == blip.dataLength ? WordPictureHtml.Picture(blip.extension, bytes) : null;
    }

    private (uint spid, int xaLeft, int yaTop, int xaRight, int yaBottom, int flags)? Fspa(int cp)
    {
        int[] cps = _anchorCps ??= ReadAnchorCps();
        int index = Array.BinarySearch(cps, 0, Math.Max(0, cps.Length - 1), cp);
        if (index < 0) return null;
        long at = _plcSpa.offset + (long)cps.Length * 4 + (long)index * FspaSize;
        byte[] fspa = DocBinaryPictures.Read(_tableStream, at, FspaSize);
        if (fspa.Length < FspaSize) return null;
        uint spid = BitConverter.ToUInt32(fspa, 0);
        if (spid == 0) return null;                    // ww8graf.cxx:2649 "Where is the graphic (3)"
        return (spid, BitConverter.ToInt32(fspa, 4), BitConverter.ToInt32(fspa, 8), BitConverter.ToInt32(fspa, 12),
                BitConverter.ToInt32(fspa, 16), BitConverter.ToUInt16(fspa, 20));
    }

    private int[] ReadAnchorCps()
    {
        (uint offset, uint length) = _plcSpa;
        if (length < 4 + 4 + FspaSize || offset + (long)length > _tableStream.Length) return [];
        int count = (int)((length - 4) / (4 + FspaSize));
        byte[] bytes = DocBinaryPictures.Read(_tableStream, offset, (count + 1) * 4);
        if (bytes.Length < (count + 1) * 4) return [];
        var cps = new int[count + 1];
        for (int index = 0; index <= count; index++) cps[index] = BitConverter.ToInt32(bytes, index * 4);
        return cps;
    }

    private (Stream stream, long dataAt, int dataLength, string extension)? Locate(uint spid)
    {
        if (_pibBySpid is null) ReadDrawing();
        return _pibBySpid!.TryGetValue(spid, out int pib) ? LocateBlip(pib) : null;
    }

    private (Stream stream, long dataAt, int dataLength, string extension)? LocateBlip(int pib)
    {
        if (_pibBySpid is null) ReadDrawing();
        if (pib < 1 || pib > _blips!.Count) return null;
        (uint offset, bool insideTable) = _blips[pib - 1];
        IEnumerable<Stream?> streams = insideTable ? [_tableStream] : [_dataStream, _wordDocument];
        foreach (Stream? stream in streams)
            if (stream is not null && DocBinaryPictures.Blip(stream, offset) is { } blip)
                return (stream, blip.dataAt, blip.dataLength, blip.extension);
        return null;
    }

    /// <summary>One walk over the record headers of the drawing: the FBSEs' blip offsets and each shape's pib.</summary>
    private void ReadDrawing()
    {
        _pibBySpid = [];
        _blips = [];
        _containerBySpid = [];
        _textBoxSpids = [];
        (uint offset, uint length) = _drawing;
        Stream table = _tableStream;
        long end = Math.Min(table.Length, offset + (long)length);
        if (length < 8 || DocBinaryPictures.RecordHeader(table, offset) is not { Type: 0xF000 } dgg) return;
        long dggEnd = Math.Min(end, offset + 8 + dgg.Length);
        for (long position = offset + 8; position + 8 <= dggEnd;)
        {
            if (DocBinaryPictures.RecordHeader(table, position) is not { } record) break;
            if (record.Type == 0xF001) ReadBStore(position + 8, Math.Min(dggEnd, position + 8 + record.Length));
            position += 8 + record.Length;
        }
        for (long position = dggEnd; position + 8 <= end;)
        {
            if (DocBinaryPictures.RecordHeader(table, position) is not { Type: 0xF002 } drawing)
            {
                position++;                             // the drawing's one-byte label (msdffimp.cxx:6001 "nPos++")
                if (DocBinaryPictures.RecordHeader(table, position) is not { Type: 0xF002 } labelled) break;
                drawing = labelled;
            }
            ReadShapes(position + 8, Math.Min(end, position + 8 + drawing.Length), depth: 0);
            position += 8 + drawing.Length;
        }
    }

    private void ReadBStore(long position, long end)
    {
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(_tableStream, position) is not { } record) return;
            uint blipAt = 0;
            bool insideTable = false;
            if (record.Type == 0xF007 && record.Length >= 32)
            {
                byte[] body = DocBinaryPictures.Read(_tableStream, position + 8, 32);
                if (body.Length == 32)
                {
                    uint size = BitConverter.ToUInt32(body, 20), foDelay = BitConverter.ToUInt32(body, 28);
                    if (foDelay == 0 && size < record.Length) { blipAt = (uint)(position + 8 + 36); insideTable = true; }
                    else blipAt = foDelay;
                }
            }
            _blips!.Add((blipAt, insideTable));     // one entry per BStore record keeps pib numbering
            position += 8 + record.Length;
        }
    }

    /// <summary>Shapes of a DgContainer / SpgrContainer: each SpContainer's spid and pib; groups are walked for their members.</summary>
    private void ReadShapes(long position, long end, int depth)
    {
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(_tableStream, position) is not { } record) return;
            long body = position + 8, bodyEnd = Math.Min(end, body + record.Length);
            if (record.Type == 0xF003 && depth < 8)
            {
                ReadShapes(body, bodyEnd, depth + 1);
                // the group's own shape is its first SpContainer (msdffdef.hxx:57): an anchor on it names the whole group
                if (DocBinaryPictures.RecordHeader(_tableStream, body) is { Type: 0xF004 } own
                    && SpidOf(body + 8, Math.Min(bodyEnd, body + 8 + own.Length)) is { } groupSpid)
                    _containerBySpid![groupSpid] = (body, bodyEnd, true);
            }
            else if (record.Type == 0xF004 && ReadShape(body, bodyEnd) is { } spid && !_containerBySpid!.ContainsKey(spid))
                _containerBySpid[spid] = (body, bodyEnd, false);
            position = body + record.Length;
        }
    }

    private uint? SpidOf(long position, long end)
    {
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(_tableStream, position) is not { } record) return null;
            if (record.Type == 0xF00A && record.Length >= 4 && DocBinaryPictures.Read(_tableStream, position + 8, 4) is { Length: 4 } id) return BitConverter.ToUInt32(id, 0);
            position += 8 + record.Length;
        }
        return null;
    }

    /// <summary>One SpContainer: its spid, pib and whether it has a text box (lTxid 128); returns the spid.</summary>
    private uint? ReadShape(long position, long end)
    {
        bool hasText = false;
        uint? spid = null;
        int? pib = null;
        while (position + 8 <= end)
        {
            if (DocBinaryPictures.RecordHeader(_tableStream, position) is not { } record) return spid;
            long body = position + 8;
            if (record.Type == 0xF00A && record.Length >= 4)
                spid = BitConverter.ToUInt32(DocBinaryPictures.Read(_tableStream, body, 4) is { Length: 4 } id ? id : new byte[4], 0);
            else if (record.Type == 0xF00B)
            {
                int count = record.Instance;            // recInstance = the property count (msdffdef.hxx:45 DFF_msofbtOPT)
                byte[] entries = DocBinaryPictures.Read(_tableStream, body, (int)Math.Min(record.Length, (long)count * 6));
                for (int index = 0; index + 6 <= entries.Length; index += 6)
                {
                    int property = BitConverter.ToUInt16(entries, index) & 0x3FFF;
                    if (property == PropertyPib) pib = (int)BitConverter.ToUInt32(entries, index + 2);
                    else if (property == 128) hasText = true;              // lTxid (msdffdef.hxx:107)
                }
            }
            position = body + record.Length;
        }
        if (spid is { } shapeId && pib is { } blipIndex) _pibBySpid![shapeId] = blipIndex;
        if (spid is { } textShape && hasText) _textBoxSpids!.Add(textShape);
        return spid;
    }
}
