namespace FastDocView.Rtf;

/// <summary>What a picture-like group was (it is never decoded: a placeholder is shown).</summary>
internal enum RtfEmbeddedKind : byte { Picture, Object, Equation }

/// <summary>
/// Receives the content the <see cref="RtfParser"/> finds. The index pass and the HTML renderer are two sinks over
/// the same parser, so a block is segmented and formatted identically by both.
/// </summary>
internal abstract class RtfSink
{
    /// <summary>False: the parser reports text lengths only and skips decoding (the index pass, outside headings).</summary>
    public virtual bool WantsText => true;

    /// <summary>The audit counts control words the reader does not act on.</summary>
    public virtual bool WantsUnhandledWords => false;

    /// <summary>The document group '{' opened (the parser is just after it).</summary>
    public virtual void DocumentStarted() { }
    public virtual void Text(ReadOnlySpan<char> text, in RtfGroupState state) { }
    public virtual void TextLength(int length, in RtfGroupState state) { }
    /// <summary>\par (or \sect when <paramref name="isSectionEnd"/>).</summary>
    public virtual void ParagraphEnd(in RtfGroupState state, bool isSectionEnd) { }
    public virtual void CellEnd(in RtfGroupState state, bool isNested) { }
    public virtual void RowEnd(in RtfGroupState state, bool isNested) { }
    /// <summary>\page, \line, \column.</summary>
    public virtual void Break(RtfWord word, in RtfGroupState state) { }
    public virtual void FootnoteStart() { }
    public virtual void FootnoteEnd(bool isEndnote) { }
    /// <summary>\chftn — the automatic note number (in the text: the reference; in the note: its mark).</summary>
    public virtual void FootnoteMark(in RtfGroupState state) { }
    public virtual void ListTextStart() { }
    public virtual void ListTextEnd() { }
    public virtual void Embedded(RtfEmbeddedKind kind, in RtfGroupState state) { }
    /// <summary>Whether {\mmath} groups are read (<see cref="Equation"/>) instead of skipped as an equation.</summary>
    public virtual bool WantsMath => false;
    /// <summary>An equation: its {\mmath} group's bytes (RtfOfficeMath turns them into OMML).</summary>
    public virtual void Equation(byte[] group, in RtfGroupState state) { }
    /// <summary>Whether {\object} groups are read (<see cref="Object"/>) instead of skipped as an object.</summary>
    public virtual bool WantsObjects => false;
    /// <summary>An embedded object: its {\object} group's bytes (<see cref="RtfObjectChart"/> reads a chart from them).</summary>
    public virtual void Object(byte[] group, in RtfGroupState state) { }
    /// <summary>True: the parser reads \pict and \shp groups into <see cref="Picture"/> (the HTML renderer); false: a placeholder (index, search).</summary>
    public virtual bool WantsPictures => false;
    /// <summary>True: <see cref="RtfPicture.Data"/> is decoded (reading one picture's bytes); false: only its size and type.</summary>
    public virtual bool WantsPictureData => false;
    /// <summary>A {\pict} group, or a {\shp} holding one (then <see cref="RtfPicture.Placement"/> is set).</summary>
    public virtual void Picture(in RtfPicture picture, in RtfGroupState state) { }
    /// <summary>A {\shp} / {\shpgrp} in the text opened (not one inside a group).</summary>
    public virtual void ShapeStart(in RtfGroupState state) { }
    /// <summary>That shape's group closed, with its properties and children.</summary>
    public virtual void ShapeEnd(RtfShapePlacement shape, in RtfGroupState state) { }
    /// <summary>A {\shptxt} of <paramref name="shape"/> opened: its paragraphs follow (their ends come as <see cref="TextBoxParagraphEnd"/>).</summary>
    public virtual void TextBoxStart(RtfShapePlacement shape) { }
    public virtual void TextBoxEnd() { }
    /// <summary>\par, \sect, \cell or \row inside a text box: ends a paragraph of the box, never the block.</summary>
    public virtual void TextBoxParagraphEnd(in RtfGroupState state) { }
    /// <summary>\cell / \nestcell and \row / \nestrow inside a text box: its own table (never the block's); by default a paragraph end.</summary>
    public virtual void TextBoxCellEnd(in RtfGroupState state, bool isNested) => TextBoxParagraphEnd(state);
    public virtual void TextBoxRowEnd(in RtfGroupState state, bool isNested) => TextBoxParagraphEnd(state);
    /// <summary>A group closed; <paramref name="depth"/> is the depth after the close (0 = the document group).</summary>
    public virtual void GroupClosed(int depth) { }
    public virtual void UnhandledWord(ReadOnlySpan<byte> name) { }
    /// <summary>A page-setup word (\paperw \margl \margr \gutter, \sectd, \pgwsxn \marglsxn \margrsxn \guttersxn) in the main text.</summary>
    public virtual void PageSetup(RtfWord word, int value) { }
    /// <summary>A section note-numbering word (\sftnn* \saftnn* \sftnstart \saftnstart \sftnrst* \saftnrst*) in the main text.</summary>
    public virtual void SectionNoteNumbering(RtfWord word, int value) { }
    /// <summary>A destination group was skipped whole (header, info, bookmarks, ...).</summary>
    public virtual void DestinationSkipped(ReadOnlySpan<byte> name) { }
    /// <summary>A \field group opened at <paramref name="depth"/> (its group closes when GroupClosed reports a smaller depth).</summary>
    public virtual void FieldStart(int depth) { }
    /// <summary>The field's {\*\fldinst ...} text, when its group closed.</summary>
    public virtual void FieldInstruction(string instruction) { }
    /// <summary>The field's {\fldrslt ...} group opened: what follows is the shown result.</summary>
    public virtual void FieldResultStart() { }
    /// <summary>\fldlock / \fldpriv in the field's group (RTF 1.5 spec "Fields" &lt;fieldmod&gt;).</summary>
    public virtual void FieldModifier(RtfWord word) { }
    /// <summary>The field's {\*\formfield} data, when its group closed (reported before the code's <see cref="FieldInstruction"/>).</summary>
    public virtual void FormField(Fields.WordFormFieldState state) { }
    /// <summary>{\*\bkmkstart name} closed: a bookmark starts here.</summary>
    public virtual void BookmarkStart(string name) { }
    /// <summary>{\*\bkmkend name}: the bookmark's end (REF without a result shows the text between).</summary>
    public virtual void BookmarkEnd(string name) { }
}

/// <summary>A picture: shown size in twips (0 = unknown), the blip type's file extension ("" = a type a browser cannot show),
/// its bytes when the sink asked for them, for a floating shape its placement, and its crop (\piccrop*).</summary>
internal readonly record struct RtfPicture(int WidthTwips, int HeightTwips, string Extension, byte[]? Data, RtfShapePlacement? Placement = null,
                                           FastDocView.Html.WordPictureCrop Crop = default);

/// <summary>
/// A {\shp} or {\shpgrp} as the parser read it ("Word 97 RTF for Drawing Objects (Shapes)", techshelps.github.io/MSDN/SPECS/
/// newrtf/word97rtffordrawingobjectsshapes.htm; meanings and units there unless noted): the placement words, the {\sp} properties
/// it draws from, its picture (a pib) and, for a group, its children in z-order (bottom first).
/// </summary>
internal sealed class RtfShapePlacement
{
    public int Left, Top, Right, Bottom;
    /// <summary>\shpwr (0 = not given).</summary>
    public int Wrap;
    /// <summary>\shpwrk (0 both sides, 1 left, 2 right, 3 largest).</summary>
    public int WrapSide;
    /// <summary>\shpbxpage or posrelh 1 (rtfdispatchflag.cxx:1229-1232, rtfsdrimport.cxx:711-721): \shpleft counts from the page edge.</summary>
    public bool IsPageRelative;
    /// <summary>posrelh 3: the horizontal offset is from the anchor character (RTF 1.9.1 shape property posrelh: 0 margin, 1 page, 2 column, 3 character).</summary>
    public bool IsCharacterRelative;
    /// <summary>posh (0 = not given: placed at \shpleft).</summary>
    public int HorizontalPosition;
    /// <summary>posv (0 = not given: placed at \shptop; 1 top, 2 centre, 3 bottom, 4 inside, 5 outside).</summary>
    public int VerticalPosition;
    /// <summary>What \shptop counts from: \shpby* or posrelv (0 margin, 1 page, 2 paragraph - the default, 3 line).</summary>
    public int VerticalRelation = 2;
    /// <summary>dxWrapDistLeft, dyWrapDistTop, dxWrapDistRight, dyWrapDistBottom in EMU; left / right not given = Word's 9 pt
    /// (Word 16 leaves the default out of its RTF and reports 9 pt through COM - <see cref="DocBinary.DocBinaryDrawnShapes.DefaultWrapDistance"/>;
    /// LibreOffice's RTF import sets no default, rtfsdrimport.cxx:640-675).</summary>
    public long WrapLeft = (long)DocBinary.DocBinaryDrawnShapes.DefaultWrapDistance, WrapTop, WrapRight = (long)DocBinary.DocBinaryDrawnShapes.DefaultWrapDistance, WrapBottom;

    /// <summary>{\shpgrp}: a group whose children are placed in its groupLeft..groupBottom coordinates.</summary>
    public bool IsGroup;
    /// <summary>The parser's group depth of the shape's own group.</summary>
    public int Depth;
    public RtfShapePlacement? Parent;
    public readonly List<RtfShapePlacement> Children = new();
    /// <summary>shapeType (-1 = not given: a rectangle, LibreOffice rtfsdrimport.cxx:293-300 "Word seems to implement it as a rectangle").</summary>
    public int ShapeType = -1;
    /// <summary>fillColor / lineColor: 24-bit BGR (LibreOffice rtfsdrimport.cxx:459, 476 BGRToRGB); null = not given.</summary>
    public int? FillColor, LineColor, FillBackColor;
    /// <summary>lineWidth in EMU (null = Word's 0.75 pt, rtfsdrimport.cxx:384).</summary>
    public long? LineWidth;
    /// <summary>fFilled / fLine (Boolean 1 / 0); null = not given.</summary>
    public bool? Filled, HasLine;
    /// <summary>dxTextLeft, dyTextTop, dxTextRight, dyTextBottom in EMU (null = 91440 / 45720, rtfsdrimport.cxx:113-124).</summary>
    public long? TextLeft, TextTop, TextRight, TextBottom;
    /// <summary>A group's coordinate system (groupLeft..) and a child's box in it (relLeft..), rtfsdrimport.cxx:733-748, 1036-1053.</summary>
    public int? GroupLeft, GroupTop, GroupRight, GroupBottom, RelLeft, RelTop, RelRight, RelBottom;
    public bool FlipHorizontal, FlipVertical, FitsShapeToText;
    /// <summary>rotation: fixed 16.16 degrees, positive = clockwise (rtfsdrimport.cxx:245-249).</summary>
    public double RotationDegrees;
    /// <summary>lineStartArrowhead / Width / Length and lineEndArrowhead / Width / Length (MSOLINEEND / -WIDTH / -LENGTH; null = not given).</summary>
    public int? StartArrowType, StartArrowWidth, StartArrowLength, EndArrowType, EndArrowWidth, EndArrowLength;
    /// <summary>adjustValue .. adjust10Value (the shape type's coordinate space; null = its default).</summary>
    public readonly int?[] Adjust = new int?[10];
    /// <summary>lineDashing (MSO_LineDashing; null = solid), fillOpacity / lineOpacity (fixed 16.16, 65536 = opaque).</summary>
    public int? LineDashing;
    public long? FillOpacity, LineOpacity;
    /// <summary>The pib picture, when the renderer read one.</summary>
    public RtfPicture? Picture;
    /// <summary>fillBlip (texture / picture / pattern fills) and the fill properties (same numbers as .doc FOPT, msdffdef.hxx:173-200).</summary>
    public RtfPicture? FillPicture;
    public int FillType;
    public long? FillBackOpacity;
    public int FillAngle, FillFocus, FillToRight, FillToBottom;
    public long FillWidth, FillHeight;
    /// <summary>fillShadeColors: "8;n;(color,dist);..." pairs (an array as the RTF shapes page writes arrays).</summary>
    public List<(int color, int position)>? FillShadeColors;
    /// <summary>The shape's own geometry: pVerticies (x, y pairs), pSegmentInfo (u16 commands) and geoLeft..geoBottom (the RTF shapes page;
    /// LibreOffice rtfsdrimport.cxx:500-600 reads the same arrays).</summary>
    public List<int>? Vertices;
    /// <summary>pWrapPolygonVertices: "8;n;(x,y);..." in Word's 21600 space over the shape (LibreOffice rtfsdrimport.cxx:856-885 reads the same array); null = none.</summary>
    public List<(double X, double Y)>? WrapPolygon;
    public List<ushort>? Segments;
    public int GeoLeft, GeoTop, GeoRight = 21600, GeoBottom = 21600;
    /// <summary>The shape has a {\shptxt}; its rendered HTML (the renderer).</summary>
    public bool HasText;
    /// <summary>The text came from gtextUNICODE (WordArt), shown in gtextFont at gtextSize (16.16 points).</summary>
    public bool IsWordArt;
    /// <summary>fGtext ("True if the text effect properties (gtext*) are used", RTF shapes page): the shape is its text (a text path) - no outline drawn.</summary>
    public bool IsTextEffect;
    /// <summary>The gtextUNICODE string itself (drawn on the outlines by a text effect).</summary>
    public string? WordArtText;
    /// <summary>gtextAlign (1 = centre when not given).</summary>
    public int WordArtAlign = 1;
    public string? WordArtFont;
    public double WordArtSizePt;
    public string? TextHtml;
    /// <summary>The 3-D properties (OfficeArt ids 640-767, booleans gathered into 703 / 767 with their "use" bits), or null.</summary>
    public Dictionary<int, uint>? ThreeD;

    // [MS-ODRAW] 2.3.18-2.3.19 property names as Word writes them in RTF (\sn), by OfficeArt id
    private static readonly Dictionary<string, int> ThreeDNumbers = new()
    {
        ["c3DSpecularAmt"] = 640, ["c3DDiffuseAmt"] = 641, ["c3DShininess"] = 642, ["c3DExtrudeForward"] = 644, ["c3DExtrudeBackward"] = 645,
        ["c3DExtrusionColor"] = 647, ["c3DYRotationAngle"] = 704, ["c3DXRotationAngle"] = 705, ["c3DRotationCenterX"] = 710,
        ["c3DRotationCenterY"] = 711, ["c3DRotationCenterZ"] = 712, ["c3DRenderMode"] = 713, ["c3DTolerance"] = 714, ["c3DXViewpoint"] = 715,
        ["c3DYViewpoint"] = 716, ["c3DZViewpoint"] = 717, ["c3DOriginX"] = 718, ["c3DOriginY"] = 719, ["c3DSkewAngle"] = 720, ["c3DSkewAmount"] = 721,
        ["c3DAmbientIntensity"] = 722, ["c3DKeyX"] = 723, ["c3DKeyY"] = 724, ["c3DKeyZ"] = 725, ["c3DKeyIntensity"] = 726, ["c3DFillX"] = 727,
        ["c3DFillY"] = 728, ["c3DFillZ"] = 729, ["c3DFillIntensity"] = 730,
    };

    // the boolean groups: 3-D Object (703: fc3DLightFace 0, fc3DUseExtrusionColor 1, fc3DMetallic 2, f3D 3) and 3-D Style
    // (767: fc3DFillHarsh 0, fc3DKeyHarsh 1, fc3DParallel 2, fc3DRotationCenterAuto 3, fc3DConstrainRotation 4)
    private static readonly Dictionary<string, (int Id, int Bit)> ThreeDFlags = new()
    {
        ["fc3DLightFace"] = (703, 0), ["fc3DUseExtrusionColor"] = (703, 1), ["fc3DMetallic"] = (703, 2), ["f3D"] = (703, 3),
        ["fc3DFillHarsh"] = (767, 0), ["fc3DKeyHarsh"] = (767, 1), ["fc3DParallel"] = (767, 2), ["fc3DRotationCenterAuto"] = (767, 3),
        ["fc3DConstrainRotation"] = (767, 4),
    };

    public void SetProperty(string name, string value)
    {
        if (ThreeDNumbers.ContainsKey(name) || ThreeDFlags.ContainsKey(name))
        {
            if (!long.TryParse(value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long threeD)) return;
            ThreeD ??= new Dictionary<int, uint>();
            if (ThreeDNumbers.TryGetValue(name, out int id)) ThreeD[id] = (uint)threeD;
            else
            {
                (int group, int bit) = ThreeDFlags[name];
                uint flags = ThreeD.TryGetValue(group, out uint old) ? old : 0;
                flags |= 1u << (bit + 16);
                flags = threeD != 0 ? flags | (1u << bit) : flags & ~(1u << bit);
                ThreeD[group] = flags;
            }
            return;
        }
        if (name == "gtextFont") { WordArtFont = value; return; }
        if (name == "pVerticies")
        {
            Vertices = new List<int>();
            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(value, @"\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"))
            {
                Vertices.Add((int)long.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture));
                Vertices.Add((int)long.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture));
            }
            return;
        }
        if (name == "pWrapPolygonVertices")
        {
            WrapPolygon = new List<(double X, double Y)>();
            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(value, @"\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"))
                if (long.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long x)
                    && long.TryParse(match.Groups[2].Value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long y))
                    WrapPolygon.Add((x, y));
            return;
        }
        if (name == "pSegmentInfo")
        {
            string[] parts = value.Split(';');
            Segments = new List<ushort>();
            for (int i = 2; i < parts.Length; i++)
                if (long.TryParse(parts[i].Trim(), System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long segment))
                    Segments.Add((ushort)segment);
            return;
        }
        if (name == "fillShadeColors")
        {
            // "size;count;(a,b);(a,b)..." - the array form of the RTF shapes page ("{\sv 8;4;{0,0};...}")
            FillShadeColors = new List<(int, int)>();
            foreach (System.Text.RegularExpressions.Match match in System.Text.RegularExpressions.Regex.Matches(value, @"\(\s*(-?\d+)\s*,\s*(-?\d+)\s*\)"))
                FillShadeColors.Add(((int)(uint)long.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), int.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture)));
            return;
        }
        if (!long.TryParse(value, System.Globalization.NumberStyles.AllowLeadingSign, System.Globalization.CultureInfo.InvariantCulture, out long number)) return;
        int clamped = (int)Math.Max(int.MinValue, Math.Min(int.MaxValue, number));
        switch (name)
        {
            case "posh": HorizontalPosition = (int)Math.Max(0, Math.Min(5, number)); break;
            case "posrelh": IsPageRelative = number == 1; IsCharacterRelative = number == 3; break;
            case "posv": VerticalPosition = (int)Math.Max(0, Math.Min(5, number)); break;
            case "posrelv": VerticalRelation = (int)Math.Max(0, Math.Min(3, number)); break;
            case "dxWrapDistLeft": WrapLeft = number; break;
            case "dyWrapDistTop": WrapTop = number; break;
            case "dxWrapDistRight": WrapRight = number; break;
            case "dyWrapDistBottom": WrapBottom = number; break;
            case "shapeType": ShapeType = clamped; break;
            case "fillColor": FillColor = (int)(uint)number; break;
            case "lineColor": LineColor = (int)(uint)number; break;
            case "fillBackColor": FillBackColor = (int)(uint)number; break;
            case "lineWidth": LineWidth = number; break;
            case "fFilled": Filled = number != 0; break;
            case "fLine": HasLine = number != 0; break;
            case "dxTextLeft": TextLeft = number; break;
            case "dyTextTop": TextTop = number; break;
            case "dxTextRight": TextRight = number; break;
            case "dyTextBottom": TextBottom = number; break;
            case "groupLeft": GroupLeft = clamped; break;
            case "groupTop": GroupTop = clamped; break;
            case "groupRight": GroupRight = clamped; break;
            case "groupBottom": GroupBottom = clamped; break;
            case "relLeft": RelLeft = clamped; break;
            case "relTop": RelTop = clamped; break;
            case "relRight": RelRight = clamped; break;
            case "relBottom": RelBottom = clamped; break;
            case "fFlipH": FlipHorizontal = number != 0; break;
            case "fFlipV": FlipVertical = number != 0; break;
            case "fFitShapeToText": FitsShapeToText = number == 1; break;
            case "rotation": RotationDegrees = number / 65536.0; break;
            case "lineStartArrowhead": StartArrowType = clamped; break;
            case "lineStartArrowWidth": StartArrowWidth = clamped; break;
            case "lineStartArrowLength": StartArrowLength = clamped; break;
            case "lineEndArrowhead": EndArrowType = clamped; break;
            case "lineEndArrowWidth": EndArrowWidth = clamped; break;
            case "lineEndArrowLength": EndArrowLength = clamped; break;
            case "adjustValue": Adjust[0] = clamped; break;
            case "adjust2Value" or "adjust3Value" or "adjust4Value" or "adjust5Value" or "adjust6Value" or "adjust7Value" or "adjust8Value"
                or "adjust9Value" or "adjust10Value":
                Adjust[int.Parse(name.Substring(6, name.Length - 11), System.Globalization.CultureInfo.InvariantCulture) - 1] = clamped;
                break;
            case "lineDashing": LineDashing = clamped; break;
            case "fillType": FillType = clamped; break;
            case "fGtext": IsTextEffect = number != 0; break;
            case "gtextAlign": WordArtAlign = clamped; break;
            case "geoLeft": GeoLeft = clamped; break;
            case "geoTop": GeoTop = clamped; break;
            case "geoRight": GeoRight = clamped; break;
            case "geoBottom": GeoBottom = clamped; break;
            case "fillBackOpacity": FillBackOpacity = number; break;
            case "fillAngle": FillAngle = clamped; break;
            case "fillFocus": FillFocus = clamped; break;
            case "fillToRight": FillToRight = clamped; break;
            case "fillToBottom": FillToBottom = clamped; break;
            case "fillWidth": FillWidth = number; break;
            case "fillHeight": FillHeight = number; break;
            case "gtextSize": WordArtSizePt = number / 65536.0; break;
            case "fillOpacity": FillOpacity = number; break;
            case "lineOpacity": LineOpacity = number; break;
        }
    }
}
