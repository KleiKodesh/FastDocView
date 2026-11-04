using System.Text;

namespace FastDocView.Docx;

/// <summary>
/// Collapses every &lt;w:drawing&gt;, &lt;w:pict&gt; and &lt;w:object&gt; subtree in a block's raw XML
/// bytes before parsing, keeping only the text boxes inside it (&lt;w:txbxContent&gt;). A single
/// drawing can carry megabytes of DrawingML that the flow renderer shows as one placeholder, but
/// the text of a text box is document text and must survive. Matching is depth-aware: drawings DO
/// nest (a picture inside a text box inside a shape), and cutting at the first close tag broke the
/// XML of such blocks.
/// </summary>
internal static class DocxDrawingStripper
{
    private static readonly string[] CollapsedElementNames = ["w:drawing", "w:pict", "w:object"];
    private const string TextBoxContentName = "w:txbxContent";

    /// <summary>The bytes of [start, start+length) with drawings collapsed; the input range untouched when there are none.</summary>
    public static ArraySegment<byte> Strip(byte[] bytes, int start, int length)
    {
        var span = new ReadOnlySpan<byte>(bytes, start, length);
        bool hasAny = false;
        foreach (string name in CollapsedElementNames) if (IndexOfOpenTag(span, 0, name) >= 0) { hasAny = true; break; }
        if (!hasAny) return new ArraySegment<byte>(bytes, start, length);

        var output = new MemoryStream(length);
        int position = 0;
        while (position < span.Length)
        {
            (int elementStart, string? name) = NextCollapsedElement(span, position);
            if (name is null) { output.Write(bytes, start + position, span.Length - position); break; }
            output.Write(bytes, start + position, elementStart - position);
            int elementEnd = ElementEnd(span, elementStart, name);
            if (elementEnd < 0) { output.Write(bytes, start + elementStart, span.Length - elementStart); break; }

            // Keep the element (so the renderer still sees "a drawing here") with only its text boxes inside.
            WriteAscii(output, "<" + name + ">");
            if (name == "w:drawing") WritePictureLayout(output, span, elementStart, elementEnd);
            else WriteVmlPictureLayout(output, span, elementStart, elementEnd);
            CopyTextBoxes(output, bytes, start, span, elementStart, elementEnd);
            WriteAscii(output, "</" + name + ">");
            position = elementEnd;
        }
        return new ArraySegment<byte>(output.GetBuffer(), 0, (int)output.Length);
    }

    private static (int index, string? name) NextCollapsedElement(ReadOnlySpan<byte> span, int from)
    {
        int best = -1;
        string? bestName = null;
        foreach (string name in CollapsedElementNames)
        {
            int found = IndexOfOpenTag(span, from, name);
            if (found >= 0 && (best < 0 || found < best)) { best = found; bestName = name; }
        }
        return (best, bestName);
    }

    /// <summary>The element the renderer reads a picture's layout from (no namespace: it is ours, not WordprocessingML).</summary>
    public const string PictureLayoutElementName = "fastDocViewPicture";

    /// <summary>
    /// A DrawingML picture's layout, kept as one small element: wp:inline / wp:anchor (ECMA-376 20.4.2.8 / 20.4.2.3) with
    /// its dist*, behindDoc, wp:extent (20.4.2.7), the wrap type (20.4.2.15-20.4.2.18: wrapSquare / wrapTight / wrapThrough /
    /// wrapTopAndBottom / wrapNone) and its wrapText, wp:positionH (20.4.2.10: relativeFrom, align or posOffset) and the
    /// first a:blip r:embed (20.1.8.13). Only the drawing's first wp:inline / wp:anchor is read; the prefixes are the ones
    /// Word writes (wp:, a:, r:), as w: is for the drawing itself. The picture bytes are not here: the renderer reads them
    /// from the package by the relationship.
    /// </summary>
    private static void WritePictureLayout(MemoryStream output, ReadOnlySpan<byte> span, int from, int to)
    {
        ReadOnlySpan<byte> drawing = span.Slice(from, to - from);
        int frame = IndexOfOpenTag(drawing, 0, "wp:anchor");
        string kind = "anchor";
        int inline = IndexOfOpenTag(drawing, 0, "wp:inline");
        if (frame < 0 || (inline >= 0 && inline < frame)) { frame = inline; kind = "inline"; }
        if (frame < 0) return;
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"" + kind + "\"");
        void Copy(string attribute, ReadOnlySpan<byte> tag, string name)
        {
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(attribute).Append("=\"").Append(value).Append('"');
        }
        ReadOnlySpan<byte> frameTag = OpenTagAt(drawing, frame);
        foreach (string name in new[] { "distT", "distB", "distL", "distR", "behindDoc" }) Copy(name, frameTag, name);
        if (IndexOfOpenTag(drawing, frame, "wp:extent") is var extent and >= 0)
        {
            Copy("cx", OpenTagAt(drawing, extent), "cx");
            Copy("cy", OpenTagAt(drawing, extent), "cy");
        }
        if (kind == "anchor")
        {
            foreach (string wrap in new[] { "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom", "wrapNone" })
                if (IndexOfOpenTag(drawing, frame, "wp:" + wrap) is var at and >= 0)
                {
                    layout.Append(" wrap=\"").Append(wrap).Append('"');
                    Copy("wrapText", OpenTagAt(drawing, at), "wrapText");
                    break;
                }
            if (IndexOfOpenTag(drawing, frame, "wp:positionH") is var horizontal and >= 0)
            {
                Copy("relativeFrom", OpenTagAt(drawing, horizontal), "relativeFrom");
                int end = drawing.Slice(horizontal).IndexOf("</wp:positionH>"u8);
                ReadOnlySpan<byte> position = end < 0 ? default : drawing.Slice(horizontal, end);
                foreach ((string element, string attribute) in new[] { ("wp:align", "align"), ("wp:posOffset", "posOffset") })
                    if (IndexOfOpenTag(position, 0, element) is var at and >= 0 && position.Slice(at).IndexOf((byte)'>') is var open and >= 0
                        && position.Slice(at + open + 1).IndexOf((byte)'<') is var length and >= 0)
                        layout.Append(' ').Append(attribute).Append("=\"").Append(Encoding.UTF8.GetString(position.Slice(at + open + 1, length).ToArray()).Trim().Replace("\"", "&quot;")).Append('"');
            }
        }
        if (IndexOfOpenTag(drawing, frame, "a:blip") is var blip and >= 0) Copy("embed", OpenTagAt(drawing, blip), "r:embed");
        WriteTextBoxShape(layout, drawing);
        string shapes = ShapeTree(drawing, frame);
        if (shapes.Length == 0) layout.Append("/>");
        else layout.Append('>').Append(shapes).Append("</").Append(PictureLayoutElementName).Append('>');
        WriteUtf8(output, layout.ToString());
    }

    /// <summary>
    /// The DrawingML elements a shape or group is drawn from (ECMA-376 Part 1 20.1 / 20.4, wps = Word 2010 shapes, wpg = groups),
    /// copied without their prefixes (so the fragment needs no namespace declarations) and without anything else: wpg:wgp /
    /// wpg:grpSp with wpg:grpSpPr, wps:wsp with wps:spPr / wps:style / wps:bodyPr and an empty wps:txbx marker (its
    /// w:txbxContent is copied after the layout, in document order, as before), pic:pic with its a:blip and spPr; a:xfrm
    /// (off / ext / chOff / chExt), a:prstGeom / a:custGeom paths, fills, colours and their transforms, a:ln. Every other
    /// element (a:extLst, effects, 3-D, non-visual properties) is dropped with its subtree. Empty when the graphic holds no
    /// wps / wpg content (a plain picture keeps its single layout element).
    /// </summary>
    private static string ShapeTree(ReadOnlySpan<byte> drawing, int frame)
    {
        int data = IndexOfOpenTag(drawing, frame, "a:graphicData");
        if (data < 0) return "";
        int dataEnd = ElementEnd(drawing, data, "a:graphicData");
        if (dataEnd < 0) return "";
        ReadOnlySpan<byte> graphic = drawing.Slice(data, dataEnd - data);
        if (IndexOfOpenTag(graphic, 0, "wps:wsp") < 0 && IndexOfOpenTag(graphic, 0, "wpg:wgp") < 0 && IndexOfOpenTag(graphic, 0, "wpc:wpc") < 0) return "";
        var tree = new StringBuilder("<shapes>");
        int skipDepth = 0;
        var kept = new Stack<string>();
        int position = graphic.Slice(1).IndexOf((byte)'>') + 2;   // past the a:graphicData open tag
        while (position > 0 && position < graphic.Length)
        {
            int open = graphic.Slice(position).IndexOf((byte)'<');
            if (open < 0) break;
            int at = position + open;
            int close = graphic.Slice(at).IndexOf((byte)'>');
            if (close < 0) break;
            ReadOnlySpan<byte> tag = graphic.Slice(at + 1, close - 1);
            position = at + close + 1;
            if (tag.Length == 0 || tag[0] is (byte)'?' or (byte)'!') continue;
            if (tag[0] == (byte)'/')
            {
                if (skipDepth > 0) { skipDepth--; continue; }
                if (kept.Count == 0) break;   // </a:graphicData>
                tree.Append("</").Append(kept.Pop()).Append('>');
                continue;
            }
            bool selfClosing = tag[tag.Length - 1] == (byte)'/';
            if (selfClosing) tag = tag.Slice(0, tag.Length - 1);
            int nameEnd = 0;
            while (nameEnd < tag.Length && tag[nameEnd] is not ((byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n')) nameEnd++;
            string name = Encoding.ASCII.GetString(tag.Slice(0, nameEnd).ToArray());
            if (skipDepth > 0 || !KeptShapeElements.Contains(name))
            {
                if (!selfClosing) skipDepth++;
                continue;
            }
            tree.Append('<').Append(LocalName(name));
            AppendUnprefixedAttributes(tree, tag.Slice(nameEnd));
            if (name == "wps:txbx")
            {
                // the marker only: the text box content is copied (and stripped) after the layout
                tree.Append("/>");
                if (!selfClosing) skipDepth++;
                continue;
            }
            if (selfClosing) tree.Append("/>");
            else { tree.Append('>'); kept.Push(LocalName(name)); }
        }
        while (kept.Count > 0) tree.Append("</").Append(kept.Pop()).Append('>');   // a cut-off graphic still closes what it opened
        tree.Append("</shapes>");
        return tree.ToString();
    }

    private static readonly HashSet<string> KeptShapeElements = new(StringComparer.Ordinal)
    {
        "wpg:wgp", "wpg:grpSp", "wpg:grpSpPr", "wps:wsp", "wps:spPr", "wps:style", "wps:bodyPr", "wps:txbx", "wps:linkedTxbx",
        "pic:pic", "pic:blipFill", "pic:spPr", "a:blip",
        "a:xfrm", "a:off", "a:ext", "a:chOff", "a:chExt", "a:prstGeom", "a:avLst", "a:gd", "a:custGeom", "a:pathLst", "a:path",
        "a:moveTo", "a:lnTo", "a:cubicBezTo", "a:quadBezTo", "a:arcTo", "a:close", "a:pt",
        "a:solidFill", "a:noFill", "a:gradFill", "a:gsLst", "a:gs", "a:lin", "a:grpFill",
        "a:srgbClr", "a:schemeClr", "a:sysClr", "a:prstClr", "a:lumMod", "a:lumOff", "a:tint", "a:shade", "a:alpha",
        // the rest of the colour model (DocxDrawingColor), guides, gradient paths and custom dashes
        "a:scrgbClr", "a:hslClr", "a:alphaMod", "a:alphaOff", "a:lum", "a:hue", "a:hueMod", "a:hueOff", "a:sat", "a:satMod", "a:satOff",
        "a:red", "a:redMod", "a:redOff", "a:green", "a:greenMod", "a:greenOff", "a:blue", "a:blueMod", "a:blueOff",
        "a:gray", "a:comp", "a:inv", "a:gamma", "a:invGamma", "a:gdLst", "a:fillToRect", "a:custDash", "a:ds",
        // canvases (wpc), picture and pattern fills, the custom geometry's text rectangle
        "wpc:wpc", "wpc:bg", "wpc:whole", "a:blipFill", "a:stretch", "a:fillRect", "a:tile", "a:alphaModFix", "a:pattFill", "a:fgClr", "a:bgClr", "a:rect",
        "a:ln", "a:prstDash", "a:headEnd", "a:tailEnd", "a:lnRef", "a:fillRef", "a:spAutoFit",
    };

    private static string LocalName(string name) => name.IndexOf(':') is var colon and >= 0 ? name.Substring(colon + 1) : name;

    /// <summary>The tag's attributes without a prefix, plus r:embed as embed (the picture relationship); values stay escaped.</summary>
    private static void AppendUnprefixedAttributes(StringBuilder tree, ReadOnlySpan<byte> attributes)
    {
        int position = 0;
        while (position < attributes.Length)
        {
            while (position < attributes.Length && attributes[position] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') position++;
            int equals = attributes.Slice(position).IndexOf((byte)'=');
            if (equals < 0) return;
            string name = Encoding.ASCII.GetString(attributes.Slice(position, equals).ToArray()).Trim();
            int valueAt = position + equals + 1;
            while (valueAt < attributes.Length && attributes[valueAt] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n') valueAt++;
            if (valueAt >= attributes.Length || attributes[valueAt] is not ((byte)'"' or (byte)'\'')) return;
            byte quote = attributes[valueAt];
            int length = attributes.Slice(valueAt + 1).IndexOf(quote);
            if (length < 0) return;
            if (name == "r:embed") name = "embed";
            if (name.IndexOf(':') < 0 && name.Length > 0)
            {
                string value = Encoding.UTF8.GetString(attributes.Slice(valueAt + 1, length).ToArray());
                tree.Append(' ').Append(name).Append("=\"").Append(quote == (byte)'"' ? value : value.Replace("\"", "&quot;")).Append('"');
            }
            position = valueAt + 1 + length + 1;
        }
    }

    /// <summary>
    /// A wps:wsp text box's look (ECMA-376 Part 1 20.4 / 20.1, wps = Word 2010 shapes): wps:bodyPr lIns / tIns / rIns / bIns
    /// (EMU), vert, a:spAutoFit (LibreOffice oox textbodypropertiescontext.cxx:75-80 insets, :132 vert, :255-259 spAutoFit =
    /// auto-grow height), a:xfrm rot (60000ths of a degree, 20.1.7.6), and in wps:spPr the fill (a:solidFill / a:noFill) and
    /// a:ln (w in EMU, its own a:solidFill / a:noFill, a:prstDash). Only srgbClr colours are read; a theme colour (schemeClr)
    /// falls back to no fill and a black line, as the box still needs an edge.
    /// </summary>
    private static void WriteTextBoxShape(StringBuilder layout, ReadOnlySpan<byte> drawing)
    {
        int textBox = IndexOfOpenTag(drawing, 0, "wps:txbx");
        if (textBox < 0) return;
        layout.Append(" textBox=\"1\"");
        void Copy(string attribute, ReadOnlySpan<byte> tag, string name)
        {
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(attribute).Append("=\"").Append(value).Append('"');
        }
        if (IndexOfOpenTag(drawing, textBox, "wps:bodyPr") is var body and >= 0)
        {
            ReadOnlySpan<byte> bodyTag = OpenTagAt(drawing, body);
            foreach (string name in new[] { "lIns", "tIns", "rIns", "bIns", "vert" }) Copy(name, bodyTag, name);
            int bodyEnd = drawing.Slice(body).IndexOf("</wps:bodyPr>"u8);
            if (bodyEnd >= 0 && IndexOfOpenTag(drawing.Slice(body, bodyEnd), 0, "a:spAutoFit") >= 0) layout.Append(" autoFit=\"1\"");
        }
        int shapeProperties = IndexOfOpenTag(drawing, 0, "wps:spPr");
        if (shapeProperties < 0) return;
        int shapeEnd = drawing.Slice(shapeProperties).IndexOf("</wps:spPr>"u8);
        if (shapeEnd < 0) return;
        ReadOnlySpan<byte> spPr = drawing.Slice(shapeProperties, shapeEnd);
        if (IndexOfOpenTag(spPr, 0, "a:xfrm") is var transform and >= 0) Copy("rot", OpenTagAt(spPr, transform), "rot");
        int line = IndexOfOpenTag(spPr, 0, "a:ln");
        ReadOnlySpan<byte> fillPart = line < 0 ? spPr : spPr.Slice(0, line);
        int solid = IndexOfOpenTag(fillPart, 0, "a:solidFill");
        if (IndexOfOpenTag(fillPart, 0, "a:noFill") is var noFill and >= 0 && (solid < 0 || noFill < solid)) layout.Append(" fill=\"none\"");
        else if (solid >= 0 && IndexOfOpenTag(fillPart, solid, "a:srgbClr") is var fillColor and >= 0) Copy("fill", OpenTagAt(fillPart, fillColor), "val");
        if (line < 0) return;
        ReadOnlySpan<byte> lineTag = OpenTagAt(spPr, line);
        if (lineTag.Length > 0 && lineTag[lineTag.Length - 1] == (byte)'/') { Copy("lineW", lineTag, "w"); return; }
        int lineEnd = spPr.Slice(line).IndexOf("</a:ln>"u8);
        ReadOnlySpan<byte> ln = lineEnd < 0 ? spPr.Slice(line) : spPr.Slice(line, lineEnd);
        layout.Append(" line=\"1\"");
        Copy("lineW", lineTag, "w");
        if (IndexOfOpenTag(ln, 0, "a:noFill") >= 0) layout.Append(" lineFill=\"none\"");
        else if (IndexOfOpenTag(ln, 0, "a:srgbClr") is var lineColor and >= 0) Copy("lineFill", OpenTagAt(ln, lineColor), "val");
        if (IndexOfOpenTag(ln, 0, "a:prstDash") is var dash and >= 0) Copy("dash", OpenTagAt(ln, dash), "val");
    }

    /// <summary>
    /// An old-style VML picture (w:pict / w:object, ECMA-376 Part 4 VML v:imagedata) kept as the same small element
    /// with kind="vml": the r:id of the first v:imagedata (o:relid when that is what it carries - LibreOffice vmlshapecontext.cxx
    /// imagedata: "shapes in docx use r:id"), the style attribute of the shape that holds it (v:shape / v:rect / v:image:
    /// width, height, position, mso-position-horizontal, margin-left, mso-wrap-distance-*; vmlshapecontext.cxx setStyle) and
    /// w10:wrap type / side (vmlshapecontext.cxx NMSP_vmlWord | XML_wrap). A w:pict without an image (a shape, a text box)
    /// writes nothing - its text boxes are kept as before.
    /// </summary>
    private static void WriteVmlPictureLayout(MemoryStream output, ReadOnlySpan<byte> span, int from, int to)
    {
        ReadOnlySpan<byte> picture = span.Slice(from, to - from);
        int imageData = IndexOfOpenTag(picture, 0, "v:imagedata");
        if (IndexOfOpenTag(picture, 0, "v:group") is var group and >= 0) { WriteVmlGroupLayout(output, picture, group); return; }
        if (imageData < 0) { WriteVmlShapeLayout(output, picture); return; }
        ReadOnlySpan<byte> imageTag = OpenTagAt(picture, imageData);
        string? embed = AttributeValue(imageTag, "r:id") ?? AttributeValue(imageTag, "o:relid");
        if (embed is null) return;
        // the shape holding the image: the last v:shape / v:rect / v:image opened before it
        int shape = -1;
        foreach (string name in new[] { "v:shape", "v:rect", "v:image" })
        {
            for (int at = IndexOfOpenTag(picture, 0, name); at >= 0 && at < imageData; at = IndexOfOpenTag(picture, at + 1, name))
                if (at > shape) shape = at;
        }
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"vml\" embed=\"" + embed + "\"");
        if (shape >= 0 && AttributeValue(OpenTagAt(picture, shape), "style") is { } style) layout.Append(" style=\"").Append(style).Append('"');
        if (IndexOfOpenTag(picture, 0, "w10:wrap") is var wrap and >= 0)
        {
            ReadOnlySpan<byte> wrapTag = OpenTagAt(picture, wrap);
            if (AttributeValue(wrapTag, "type") is { } type) layout.Append(" wrap=\"").Append(type).Append('"');
            if (AttributeValue(wrapTag, "side") is { } side) layout.Append(" wrapText=\"").Append(side).Append('"');
        }
        layout.Append("/>");
        WriteUtf8(output, layout.ToString());
    }

    /// <summary>
    /// A VML shape without a picture (ECMA-376 Part 4 14.1: v:rect / v:roundrect / v:oval / v:line / v:shape) as the same
    /// layout element with kind="vml" and vmlShape = its element: style (size, position, wrap gaps, rotation, flip), fillcolor /
    /// filled, strokecolor / stroked / strokeweight, arcsize (roundrect), from / to (line), the v:fill / v:stroke child's
    /// on / color, and v:textbox inset with textBox="1" when it holds text; w10:wrap as for pictures. The attributes are the
    /// ones LibreOffice reads (oox vmlshapecontext.cxx:397-403 stroke / fill, vmltextboxcontext.cxx:186-211 inset).
    /// The first such element of the w:pict is the shape (a v:group: WriteVmlGroupLayout).
    /// </summary>
    private static void WriteVmlShapeLayout(MemoryStream output, ReadOnlySpan<byte> picture)
    {
        int shape = -1;
        string shapeName = "";
        foreach (string name in new[] { "v:rect", "v:roundrect", "v:oval", "v:line", "v:shape" })
            if (IndexOfOpenTag(picture, 0, name) is var at and >= 0 && (shape < 0 || at < shape)) { shape = at; shapeName = name; }
        if (shape < 0) return;
        ReadOnlySpan<byte> shapeTag = OpenTagAt(picture, shape);
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"vml\" vmlShape=\"" + shapeName.Substring(2) + "\"");
        int shapeEnd = ElementEnd(picture, shape, shapeName);
        AppendVmlShapeAttributes(layout, shapeTag, shapeEnd < 0 ? picture.Slice(shape) : picture.Slice(shape, shapeEnd - shape));
        void Copy(string attribute, ReadOnlySpan<byte> tag, string name)
        {
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(attribute).Append("=\"").Append(value).Append('"');
        }
        if (IndexOfOpenTag(picture, 0, "w10:wrap") is var wrap and >= 0)
        {
            ReadOnlySpan<byte> wrapTag = OpenTagAt(picture, wrap);
            Copy("wrap", wrapTag, "type");
            Copy("wrapText", wrapTag, "side");
        }
        layout.Append("/>");
        WriteUtf8(output, layout.ToString());
    }

    /// <summary>A VML shape's drawing attributes (see <see cref="WriteVmlShapeLayout"/>) from its open tag and body, onto an element being written.</summary>
    private static void AppendVmlShapeAttributes(StringBuilder layout, ReadOnlySpan<byte> shapeTag, ReadOnlySpan<byte> body)
    {
        void Copy(string attribute, ReadOnlySpan<byte> tag, string name)
        {
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(attribute).Append("=\"").Append(value).Append('"');
        }
        foreach (string name in new[] { "style", "fillcolor", "filled", "strokecolor", "stroked", "strokeweight", "arcsize", "from", "to" }) Copy(name, shapeTag, name);
        if (IndexOfOpenTag(body, 0, "v:fill") is var fill and >= 0) { Copy("fillOn", OpenTagAt(body, fill), "on"); Copy("fillColor2", OpenTagAt(body, fill), "color"); }
        if (IndexOfOpenTag(body, 0, "v:stroke") is var stroke and >= 0)
        {
            ReadOnlySpan<byte> strokeTag = OpenTagAt(body, stroke);
            Copy("dashstyle", strokeTag, "dashstyle"); Copy("strokeOn", strokeTag, "on"); Copy("endcap", strokeTag, "endcap");
            foreach (string end in new[] { "startarrow", "startarrowwidth", "startarrowlength", "endarrow", "endarrowwidth", "endarrowlength" }) Copy(end, strokeTag, end);
        }
        if (IndexOfOpenTag(body, 0, "v:textbox") is var textBox and >= 0 && IndexOfOpenTag(body, textBox, TextBoxContentName) >= 0)
        {
            layout.Append(" textBox=\"1\"");
            Copy("inset", OpenTagAt(body, textBox), "inset");
            Copy("insetmode", OpenTagAt(body, textBox), "insetmode");
            if (AttributeValue(OpenTagAt(body, textBox), "style") is { } textStyle && textStyle.Contains("mso-fit-shape-to-text:t")) layout.Append(" autoFit=\"1\"");
        }
    }

    private static readonly string[] VmlChildNames = ["v:group", "v:rect", "v:roundrect", "v:oval", "v:line", "v:shape", "v:image"];

    /// <summary>
    /// A w:pict holding a v:group (ECMA-376 Part 4 14.1.2.7): the layout element (kind="vml" vmlShape="group", the group's style
    /// and w10:wrap) with the group tree inside: a vgroup per v:group (style, coordorigin, coordsize) and a vshape per child
    /// shape (its element as vmlShape, the attributes of <see cref="AppendVmlShapeAttributes"/>, the v:imagedata r:id as
    /// embed), in document order, so the n-th text box is the n-th shape with textBox="1" (the w:txbxContent copies follow
    /// in the same order). LibreOffice reads the children in the group's coordinate system the same way (vmlshape.cxx
    /// GroupShape::implConvertAndInsert :1575-1585, lclGetAbsRect :105-116).
    /// </summary>
    private static void WriteVmlGroupLayout(MemoryStream output, ReadOnlySpan<byte> picture, int group)
    {
        var layout = new StringBuilder("<" + PictureLayoutElementName + " kind=\"vml\" vmlShape=\"group\"");
        if (AttributeValue(OpenTagAt(picture, group), "style") is { } style) layout.Append(" style=\"").Append(style).Append('"');
        if (IndexOfOpenTag(picture, 0, "w10:wrap") is var wrap and >= 0)
        {
            ReadOnlySpan<byte> wrapTag = OpenTagAt(picture, wrap);
            if (AttributeValue(wrapTag, "type") is { } type) layout.Append(" wrap=\"").Append(type).Append('"');
            if (AttributeValue(wrapTag, "side") is { } side) layout.Append(" wrapText=\"").Append(side).Append('"');
        }
        layout.Append('>');
        int end = ElementEnd(picture, group, "v:group");
        AppendVmlGroup(layout, end < 0 ? picture.Slice(group) : picture.Slice(group, end - group));
        layout.Append("</").Append(PictureLayoutElementName).Append('>');
        WriteUtf8(output, layout.ToString());
    }

    /// <summary>One v:group (its span from its open tag) as a vgroup element with its children.</summary>
    private static void AppendVmlGroup(StringBuilder layout, ReadOnlySpan<byte> group)
    {
        ReadOnlySpan<byte> tag = OpenTagAt(group, 0);
        layout.Append("<vgroup");
        foreach (string name in new[] { "style", "coordorigin", "coordsize" })
            if (AttributeValue(tag, name) is { } value) layout.Append(' ').Append(name).Append("=\"").Append(value).Append('"');
        layout.Append('>');
        int position = tag.Length + 1;
        while (position < group.Length)
        {
            int next = -1;
            string nextName = "";
            foreach (string name in VmlChildNames)
                if (IndexOfOpenTag(group, position, name) is var at and >= 0 && (next < 0 || at < next)) { next = at; nextName = name; }
            if (next < 0) break;
            int end = ElementEnd(group, next, nextName);
            ReadOnlySpan<byte> child = end < 0 ? group.Slice(next) : group.Slice(next, end - next);
            if (nextName == "v:group") AppendVmlGroup(layout, child);
            else
            {
                layout.Append("<vshape vmlShape=\"").Append(nextName.Substring(2)).Append('"');
                AppendVmlShapeAttributes(layout, OpenTagAt(child, 0), child);
                if (IndexOfOpenTag(child, 0, "v:imagedata") is var image and >= 0
                    && (AttributeValue(OpenTagAt(child, image), "r:id") ?? AttributeValue(OpenTagAt(child, image), "o:relid")) is { } embed)
                    layout.Append(" embed=\"").Append(embed).Append('"');
                layout.Append("/>");
            }
            if (end < 0) break;
            position = end;
        }
        layout.Append("</vgroup>");
    }


    private static void WriteUtf8(MemoryStream output, string text)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }

    /// <summary>The open tag starting at <paramref name="at"/> (up to its '&gt;').</summary>
    private static ReadOnlySpan<byte> OpenTagAt(ReadOnlySpan<byte> span, int at)
    {
        int end = span.Slice(at).IndexOf((byte)'>');
        return end < 0 ? default : span.Slice(at, end);
    }

    /// <summary>An attribute's raw value (still XML-escaped, so it can be written back as is), null when absent.</summary>
    private static string? AttributeValue(ReadOnlySpan<byte> tag, string name)
    {
        byte[] key = Encoding.ASCII.GetBytes(name + "=");
        int offset = 0;
        while (offset < tag.Length)
        {
            int found = tag.Slice(offset).IndexOf(key);
            if (found < 0) return null;
            int at = offset + found;
            int valueAt = at + key.Length;
            if (at > 0 && tag[at - 1] is (byte)' ' or (byte)'\t' or (byte)'\r' or (byte)'\n' && valueAt < tag.Length && tag[valueAt] is (byte)'"' or (byte)'\'')
            {
                byte quote = tag[valueAt];
                int length = tag.Slice(valueAt + 1).IndexOf(quote);
                if (length < 0) return null;
                string value = Encoding.UTF8.GetString(tag.Slice(valueAt + 1, length).ToArray());
                return quote == (byte)'"' ? value : value.Replace("\"", "&quot;");
            }
            offset = valueAt;
        }
        return null;
    }

    /// <summary>Copies every outermost &lt;w:txbxContent&gt;…&lt;/w:txbxContent&gt; inside [from, to).</summary>
    private static void CopyTextBoxes(MemoryStream output, byte[] bytes, int start, ReadOnlySpan<byte> span, int from, int to)
    {
        int position = from;
        while (position < to)
        {
            int textBox = IndexOfOpenTag(span, position, TextBoxContentName);
            if (textBox < 0 || textBox >= to) return;
            int textBoxEnd = ElementEnd(span, textBox, TextBoxContentName);
            if (textBoxEnd < 0 || textBoxEnd > to) return;
            // A text box's own drawings are collapsed the same way, recursively.
            ArraySegment<byte> inner = Strip(bytes, start + textBox, textBoxEnd - textBox);
            output.Write(inner.Array!, inner.Offset, inner.Count);
            position = textBoxEnd;
        }
    }

    /// <summary>End (exclusive) of the element whose open tag starts at <paramref name="openAt"/>, counting nested same-name elements; -1 if unterminated.</summary>
    private static int ElementEnd(ReadOnlySpan<byte> span, int openAt, string name)
    {
        int tagEnd = span.Slice(openAt).IndexOf((byte)'>');
        if (tagEnd < 0) return -1;
        if (span[openAt + tagEnd - 1] == (byte)'/') return openAt + tagEnd + 1;     // self-closing
        byte[] close = Encoding.ASCII.GetBytes("</" + name + ">");
        int depth = 1, position = openAt + tagEnd + 1;
        while (depth > 0)
        {
            int nextOpen = IndexOfOpenTag(span, position, name);
            int nextClose = span.Slice(position).IndexOf(close);
            if (nextClose < 0) return -1;
            nextClose += position;
            if (nextOpen >= 0 && nextOpen < nextClose)
            {
                int openEnd = span.Slice(nextOpen).IndexOf((byte)'>');
                if (openEnd < 0) return -1;
                if (span[nextOpen + openEnd - 1] != (byte)'/') depth++;
                position = nextOpen + openEnd + 1;
                continue;
            }
            depth--;
            position = nextClose + close.Length;
        }
        return position;
    }

    /// <summary>Index of "&lt;name" followed by a delimiter (so &lt;w:pict does not match &lt;w:pictureX), at or after <paramref name="from"/>.</summary>
    private static int IndexOfOpenTag(ReadOnlySpan<byte> span, int from, string name)
    {
        byte[] open = Encoding.ASCII.GetBytes("<" + name);
        int offset = from;
        while (offset < span.Length)
        {
            int found = span.Slice(offset).IndexOf(open);
            if (found < 0) return -1;
            int after = offset + found + open.Length;
            if (after >= span.Length || span[after] is (byte)' ' or (byte)'>' or (byte)'/' or (byte)'\t' or (byte)'\r' or (byte)'\n') return offset + found;
            offset = after;
        }
        return -1;
    }

    private static void WriteAscii(MemoryStream output, string text)
    {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        output.Write(bytes, 0, bytes.Length);
    }
}
