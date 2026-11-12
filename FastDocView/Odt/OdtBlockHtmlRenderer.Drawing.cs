using System.Globalization;
using System.Text;
using System.Xml.Linq;
using FastDocView.Html;

namespace FastDocView.Odt;

/// <summary>
/// Drawing shapes of an OpenDocument block: <see cref="OdtShapeBuilder"/> gives every shape's outline, paint and transform;
/// they are drawn as one frame placed like a picture (<see cref="Placement"/>), one SVG (<see cref="WordShapeHtml"/>), each
/// shape's text in its text area over it, or laid on the outlines as Fontwork.
/// </summary>
public sealed partial class OdtBlockHtmlRenderer
{
    private OdtShapeBuilder? _shapeBuilder;
    private readonly Dictionary<string, (double W, double H)?> _pictureSizes = new(StringComparer.Ordinal);

    /// <summary>The block file's root start tag (its office:version decides how unitless gradient angles read).</summary>
    public string? RootOpenTag { get; set; }

    /// <summary>Reads a picture by id (the document's <see cref="IWordWindowedDocument.ReadPicture"/>): a tiled fill without a size needs the picture's own size.</summary>
    public Func<string, WordDocumentPicture?>? Pictures { get; set; }

    /// <summary>A list style by name (the index keeps those of lists inside text boxes and shapes).</summary>
    public Func<string, OdtListStyle?>? ListStyles { get; set; }

    /// <summary>Picture ids of fill images stored in the styles (office:binary-data of a draw:fill-image): "f:" + its draw:name.</summary>
    public const string FillImagePrefix = "f:";

    /// <summary>The document's meta:generator (set by the document; null when it has none).</summary>
    public string? Generator { get; set; }

    /// <summary>How this document's unitless gradient angles read (<see cref="OdtGeneratorVersion.GradientAngleIsTenths"/>).</summary>
    private bool GradientAngleIsTenths => OdtGeneratorVersion.GradientAngleIsTenths(OfficeVersion, OdtGeneratorVersion.Of(Generator));

    private double OfficeVersion
    {
        get
        {
            if (RootOpenTag is null) return 1.3;
            int at = RootOpenTag.IndexOf("office:version=", StringComparison.Ordinal);
            if (at < 0) return 1.1;                     // no version: ODF 1.0 / 1.1 (the attribute is required from 1.2)
            int start = at + "office:version=".Length + 1, end = start < RootOpenTag.Length ? RootOpenTag.IndexOfAny(['"', '\''], start) : -1;
            return end > start && double.TryParse(RootOpenTag.Substring(start, end - start), NumberStyles.Float, CultureInfo.InvariantCulture, out double version) ? version : 1.3;
        }
    }

    /// <summary>The shapes connectors may be glued to, by id (the index's <see cref="OdtConnectedShapeTable"/>; null until the index is final - a first open draws glued ends as free points until then).</summary>
    internal Func<string, OdtConnectedShape?>? ConnectedShapes { get; set; }

    private OdtShapeBuilder ShapeBuilder
    {
        get
        {
            _shapeBuilder ??= new OdtShapeBuilder(_styles, GradientAngleIsTenths, FillPictureOf);
            _shapeBuilder.ConnectedShapes = ConnectedShapes;
            _shapeBuilder.CurrentBlockIndex = CurrentBlockIndex;
            _shapeBuilder.IsRecordingProvisionalGlue = !_isMeasuringConnectedShapes;
            return _shapeBuilder;
        }
    }

    /// <summary>
    /// Index time: records the shapes with an id in one block (a block from <see cref="OdtRandomAccessIndex.ConnectedShapeBlocks"/>) - each
    /// outermost drawing object in it with the coordinates it stands in (<see cref="OdtConnectedShapeTable.FrameOf"/>).
    /// </summary>
    private bool _isMeasuringConnectedShapes;

    /// <summary>Whether the block just rendered drew a glued connector without its shapes (<see cref="WordDocumentBlockHtml.IsProvisional"/>).</summary>
    internal bool LastBlockIsProvisional => _shapeBuilder?.UsedProvisionalGlue == true;

    /// <summary>Starts a block: nothing provisional drawn yet.</summary>
    internal void ResetProvisional() { if (_shapeBuilder is not null) _shapeBuilder.UsedProvisionalGlue = false; }

    internal void RecordConnectedShapes(XElement block, int blockIndex, Dictionary<string, OdtConnectedShape> into)
    {
        ConnectedShapes = null;          // the shapes are measured, never routed against a half-built table
        CurrentBlockIndex = blockIndex;
        _isMeasuringConnectedShapes = true;
        try
        {
            foreach (XElement element in block.DescendantsAndSelf())
            {
                if (element.Name.Namespace != DrawNs || Array.IndexOf(OdtBodyBlockScanner.BodyDrawingNames, "draw:" + element.Name.LocalName) < 0) continue;
                (string key, XElement outermost) = OdtConnectedShapeTable.FrameOf(element, blockIndex);
                if (outermost == element) OdtConnectedShapeTable.Record(element, key, ShapeBuilder, into);
            }
        }
        finally { _isMeasuringConnectedShapes = false; }
    }

    /// <summary>A draw:fill-image's picture id ("p:" + package entry, or "f:" + its name for inline data) and its own size.</summary>
    private (string Id, (double W, double H)? Size)? FillPictureOf(XElement fillImage)
    {
        string? id = null;
        if (fillImage.Element(OfficeNs + "binary-data") is not null && (string?)fillImage.Attribute(DrawNs + "name") is { Length: > 0 } name) id = FillImagePrefix + name;
        else if ((string?)fillImage.Attribute(XlinkNs + "href") is { Length: > 0 } href && !href.Contains(':'))
        {
            string part = href.StartsWith("./", StringComparison.Ordinal) ? href.Substring(2) : href;
            if (WordPictureHtml.ContentTypeOf(part) is not null) id = PackagePicturePrefix + part;
        }
        if (id is null) return null;
        return (id, PictureSizeOf(id));
    }

    /// <summary>A picture's own size in points from its header (<see cref="WordPictureSize"/>), read once per id (a bounded cache); null when unknown.</summary>
    private (double W, double H)? PictureSizeOf(string id)
    {
        if (!_pictureSizes.TryGetValue(id, out (double W, double H)? size))
        {
            size = Pictures?.Invoke(id) is { } picture ? WordPictureSize.Read(picture.Bytes) : null;
            if (_pictureSizes.Count > 256) _pictureSizes.Clear();
            _pictureSizes[id] = size;
        }
        return size;
    }

    private static bool IsDrawnShape(XElement element) => OdtShapeBuilder.IsShape(element);

    /// <summary>
    /// Contour wrapping (style:wrap-contour="true", ODF 1.3 part 3 20.401): text runs along the object's outline instead of its box.
    /// The outline: a frame's draw:contour-polygon / draw:contour-path (10.4.11.2 / .3, its svg:viewBox mapped to the frame), a
    /// shape's own outlines; the shared <see cref="WordPictureContour"/> turns it into shape-outside. style:wrap-contour-mode full
    /// (text also into the outline's indentations) cannot be laid out in CSS: a float keeps its text on one side (CSS Shapes 1,
    /// section 2) - drawn as outside. The layout unchanged when the object does not float or wraps its box.
    /// </summary>
    private WordPictureLayout WithContour(XElement element, WordPictureLayout layout, double width, double height, List<List<(double X, double Y)>>? outlines)
    {
        if (layout.Wrap != WordPictureWrap.Square || !(width > 0) || !(height > 0) || double.IsInfinity(width) || double.IsInfinity(height)) return layout;
        if (GraphicProperty(element, OdtStyleCascade.Style + "wrap-contour") != "true") return layout;
        if (element.Name == DrawNs + "frame")
        {
            outlines = null;
            XElement? contour = element.Element(DrawNs + "contour-polygon") ?? element.Element(DrawNs + "contour-path");
            if (contour is null || OdtSvgPath.ViewBox((string?)contour.Attribute(SvgNs + "viewBox")) is not { Width: > 0, Height: > 0 } box) return layout;
            double sx = width / box.Width, sy = height / box.Height;
            if (contour.Name.LocalName == "contour-polygon")
            {
                if (OdtSvgPath.Points((string?)contour.Attribute(DrawNs + "points")) is { Count: > 2 } points)
                    outlines = [points.Select(p => ((p.X - box.X) * sx, (p.Y - box.Y) * sy)).ToList()];
            }
            else if (OdtSvgPath.Parse((string?)contour.Attribute(SvgNs + "d")) is { Count: > 0 } paths)
                outlines = paths.Select(p => OdtShapeStyle.Flatten(p).Select(q => ((q.X - box.X) * sx, (q.Y - box.Y) * sy)).ToList()).ToList();
        }
        return outlines is { Count: > 0 } ? layout with { Contour = outlines } : layout;
    }

    /// <summary>
    /// A drawing shape (or group) as a frame where it is anchored (<see cref="Placement"/>): the frame is the snap rectangle of
    /// every outline (the shapes' boxes through their transforms - LibreOffice's position is the snap rectangle's), the SVG fills
    /// it, each shape's text sits in its text area. Only the shape's XML in the block being rendered is read (windowed).
    /// </summary>
    private void AppendShape(XElement element, StringBuilder html)
    {
        List<OdtDrawnShape> drawn = ShapeBuilder.Build(element);
        if (drawn.Count == 0) return;
        double left = double.MaxValue, top = double.MaxValue, right = double.MinValue, bottom = double.MinValue;
        void Take(double px, double py) { left = Math.Min(left, px); top = Math.Min(top, py); right = Math.Max(right, px); bottom = Math.Max(bottom, py); }
        foreach (OdtDrawnShape d in drawn)
        {
            WordAffine t = d.Shape.Transform ?? WordAffine.Identity;
            foreach ((double cx, double cy) in new[] { (0.0, 0.0), (d.Shape.Width, 0.0), (0.0, d.Shape.Height), (d.Shape.Width, d.Shape.Height) })
            {
                (double px, double py) = t.Apply(cx, cy);
                Take(px, py);
            }
            // a dimension line's pieces (laid out with the label estimate; the page sizes the frame again from the measured label)
            foreach ((WordShape piece, _, _) in d.MeasureLines)
                foreach (WordShapeSubpath sub in piece.Paths!.SelectMany(p => p.Subpaths))
                    foreach ((double qx, double qy) in OdtShapeStyle.Flatten(sub)) { (double px, double py) = (piece.Transform ?? WordAffine.Identity).Apply(qx, qy); Take(px, py); }
        }
        double width = right - left, height = bottom - top;
        WordAffine toFrame = WordAffine.Translate(-left, -top);
        string?[] growth = new string?[drawn.Count];
        for (int k = 0; k < drawn.Count; k++)
        {
            OdtDrawnShape d = drawn[k];
            d.Shape.Transform = WordAffine.Then(d.Shape.Transform ?? WordAffine.Identity, toFrame);
            d.TextTransform = WordAffine.Then(d.TextTransform, toFrame);
            d.Full = WordAffine.Then(d.Full, toFrame);
            foreach ((WordShape piece, _, _) in d.MeasureLines) piece.Transform = WordAffine.Then(piece.Transform ?? WordAffine.Identity, toFrame);
            if (d.Geometry is { IsTextPath: true } && d.HasText) ApplyFontwork(d);
            growth[k] = AutoGrowAttributes(d, k);
        }
        // a frame the page lays out again (a shape growing with its text, a dimension line): every shape's box on its <g>
        bool pageLayout = growth.Any(g => g is not null) || drawn.Any(d => d.MeasureLabel is not null);
        if (pageLayout)
            for (int k = 0; k < drawn.Count; k++)
                drawn[k].Shape.GroupAttributes = " data-shape-box=\"" + Css(drawn[k].Shape.Width) + " " + Css(drawn[k].Shape.Height) + "\"" + (growth[k] is null ? "" : " data-autogrow-id=\"" + k + "\"");
        WordPictureLayout layout = Placement(element, width, height, left, top);
        List<List<(double X, double Y)>> outlines = drawn.SelectMany(d => (d.Shape.Paths ?? []).SelectMany(p => p.Subpaths)
            .Select(s => OdtShapeStyle.Flatten(s).Select(q => (d.Shape.Transform ?? WordAffine.Identity).Apply(q.X, q.Y)).ToList())).ToList();
        int frameAt = html.Length;
        WordShapeHtml.OpenFrame(html, WithContour(element, layout, width, height, outlines), growsWithText: false);
        if (pageLayout) html.Insert(frameAt + "<span class=\"word-shape\"".Length, " data-shape-page=\"" + Css(width) + " " + Css(height) + "\"");
        WordShapeHtml.AppendGeometry(html, drawn.Select(d => d.Shape).ToList(), width, height, stretches: false);
        for (int k = 0; k < drawn.Count; k++)
        {
            OdtDrawnShape d = drawn[k];
            if (d.MeasureLabel is { } m)
            {
                // the dimension line and its label together: the page measures the label and lays both out again (fastDocViewShapeText.ts)
                WordAffine toPage = d.MeasureLines.Count > 0 ? d.MeasureLines[0].Shape.Transform ?? toFrame : toFrame;
                html.Append("<span class=\"fast-doc-view-measure\" data-measure=\"").Append(d.MeasurePageLayout).Append("\" data-measure-matrix=\"").Append(Matrix(toPage))
                    .Append("\" data-measure-markers=\"").Append(MarkerData(d.MeasureMarkers.Start)).Append('|').Append(MarkerData(d.MeasureMarkers.End))
                    .Append("\" style=\"position:absolute;left:0;top:0;width:100%;height:100%;pointer-events:none\">");
                if (d.MeasureLines.Count > 0) WordShapeHtml.AppendGeometry(html, d.MeasureLines.Select(p => p.Shape).ToList(), width, height, stretches: false);
                (double mx, double my) = toPage.Apply(m.X, m.Y);
                html.Append("<span class=\"word-shape-text\" style=\"position:absolute;left:").Append(Css(mx)).Append("pt;top:").Append(Css(my))
                    .Append("pt;white-space:nowrap;pointer-events:auto;transform:translate(-50%,-50%) rotate(").Append(Css(m.Rotation)).Append("deg);\">").Append(WordHtmlText.Encode(m.Text)).Append("</span></span>");
                continue;
            }
            if (d.HasText && d.Geometry is not { IsTextPath: true }) AppendShapeText(d, html, growth[k]);
        }
        WordShapeHtml.CloseFrame(html);
    }

    private static string Matrix(WordAffine t) => string.Join(" ", new[] { t.A, t.B, t.C, t.D, t.E, t.F }.Select(v => (Math.Round(v, 6) == 0 ? 0 : v).ToString("0.######", CultureInfo.InvariantCulture)));

    /// <summary>A line-end marker for the page: "width consumed;x,y x,y ...;..." (each polygon placed with its tip at the origin, pointing up), "" for none.</summary>
    private static string MarkerData(WordLineEndMarker? marker)
    {
        if (marker is null) return "";
        static string N(double v) => v.ToString("0.####", CultureInfo.InvariantCulture);
        string Polygon(List<(double X, double Y)> points) => string.Join(" ", points.Select(p => N(p.X) + "," + N(p.Y)));
        return N(marker.Width) + " " + N(marker.Consumed) + ";" + string.Join(";", new[] { marker.Points }.Concat(marker.MorePolygons ?? []).Select(Polygon));
    }

    /// <summary>
    /// A custom shape that grows with its text ("resize shape to fit text"), as data attributes on its text box for the page,
    /// which measures the text and resizes the shape (KitveiHakodesh vue-frontend fastDocViewShapeText.ts). LibreOffice's rules:
    /// - only draw:custom-shape (SdrObjCustomShape); a draw:rect / ellipse / ... is not a text frame (SdrTextObj::IsAutoGrowHeight
    ///   needs mbTextFrame, svx svdotext.cxx:248-252); draw:frame text boxes grow by fo:min-height (<see cref="AppendTextBox"/>).
    /// - SdrObjCustomShape::IsAutoGrowHeight / Width (svdoashp.cxx:2179-2194): both read draw:auto-grow-height (sdpropls.cxx:141);
    ///   the height grows unless the text is vertical and wraps, the width grows when horizontal text does not wrap
    ///   (fo:wrap-option no-wrap, sdpropls.cxx:157) or the text is vertical. draw:auto-grow-width is not read for custom shapes.
    ///   Absent, the item pool's default holds: on (svdattr.cxx:247). Nothing on the way in sets another value: the ODF import
    ///   maps only the attribute itself (sdpropls.cxx:141; ximpshap.cxx, ximpcustomshape.cxx, shapeimport.cxx and
    ///   XMLShapeStyleContext.cxx have no auto-grow default), and neither Writer's drawing model (sw DocumentDrawModelManager.cxx,
    ///   drawdoc.cxx) nor SdrModel (svdmodel.cxx) nor the custom shape's own code (svdoashp.cxx) sets one - master read 2026-11-04.
    /// - text fitted to the shape (draw:fit-to-size true / all) does not grow (SdrTextObj::AdjustTextFrameWidthAndHeight,
    ///   svdotxat.cxx:54-56 - the custom-shape version takes the same rule through IsFitToSize).
    /// - the text area takes the text's size (at least fo:min-width / fo:min-height, GetMinTextFrameWidth / Height, sdpropls.cxx:148-149)
    ///   plus the text distances (fo:padding-*), growing or shrinking; the shape changes by that much times shape / text-area size
    ///   (svdoashp.cxx:2324-2357, ImpCalculateTextFrame :2436-2464); it grows down for top-aligned text, up for bottom, both ways
    ///   for centre / justify (:2385-2418; draw:textarea-horizontal-align left grows right, right grows left, else both).
    /// - all of that in the shape's own unrotated rectangle; a rotated shape's rectangle then moves by the turned change of its
    ///   top left (:2420-2428) - the page applies the change in the shape's own coordinates before its transform (rotation,
    ///   shear, scale) and the mirroring about the new box's centre (EnhancedCustomShapeEngine.cxx:286-298); each shape of a group
    ///   grows on its own and the frame takes the new bounds.
    /// The outline is stretched to the new size (the page has no geometry engine).
    /// </summary>
    private static string? AutoGrowAttributes(OdtDrawnShape d, int id)
    {
        if (d.Element.Name != DrawNs + "custom-shape" || !d.HasText || d.Geometry is { IsTextPath: true }) return null;
        OdtShapeStyle style = d.Style;
        if ((style.Get(OdtStyleCascade.Draw + "auto-grow-height") ?? "true") != "true") return null;
        if (style.Get(OdtStyleCascade.Draw + "fit-to-size") is "true" or "all") return null;
        bool vertical = style.Get(OdtStyleCascade.Style + "writing-mode") is "tb-rl" or "tb";
        bool wraps = style.Get(OdtStyleCascade.Fo + "wrap-option") != "no-wrap";
        bool growHeight = !vertical || !wraps, growWidth = vertical || !wraps;
        double Pad(string side) => OdtShapeStyle.Length(style.Get(OdtStyleCascade.Fo + ("padding-" + side))) ?? OdtShapeStyle.Length(style.Get(OdtStyleCascade.Fo + "padding")) ?? 0;
        double Min(string name) => style.Get(OdtStyleCascade.Fo + name) is { } v && !v.Contains('%') ? OdtShapeStyle.Length(v) ?? 0 : 0;
        (double x, double y, double w, double h) = d.TextArea;
        string vAlign = style.Get(OdtStyleCascade.Draw + "textarea-vertical-align") ?? "top";
        string hAlign = style.Get(OdtStyleCascade.Draw + "textarea-horizontal-align") ?? "justify";
        static string Numbers(params double[] values) => string.Join(" ", values.Select(v => v.ToString("0.######", CultureInfo.InvariantCulture)));
        return " data-autogrow-id=\"" + id + "\" data-autogrow=\"" + (growWidth ? "w" : "") + (growHeight ? "h" : "") + "\" data-autogrow-text=\"" + Numbers(x, y, w, h)
             + "\" data-autogrow-shape=\"" + Numbers(d.Shape.Width, d.Shape.Height) + "\" data-autogrow-pad=\"" + Numbers(Pad("left"), Pad("top"), Pad("right"), Pad("bottom"))
             + "\" data-autogrow-min=\"" + Numbers(Min("min-width"), Min("min-height")) + "\" data-autogrow-align=\"" + WordHtmlText.Encode(vAlign) + " " + WordHtmlText.Encode(hAlign)
             + "\" data-autogrow-matrix=\"" + Matrix(d.Full) + "\" data-autogrow-mirror=\"" + (d.Geometry?.MirrorHorizontal == true ? 1 : 0) + " " + (d.Geometry?.MirrorVertical == true ? 1 : 0)
             + "\" data-autogrow-turn=\"" + Numbers(Math.Atan2(d.Full.B, d.Full.A)) + "\" data-autogrow-text-rotate=\"" + Numbers(d.Geometry?.TextRotateAngle ?? 0) + "\"";
    }

    /// <summary>
    /// A shape's paragraphs in its text area (ODF 1.3 part 3 20.x / LibreOffice sdpropls.cxx:147-161): fo:padding-* are the text
    /// distances (pool default 0, svdattr.cxx:249-252), draw:textarea-vertical-align top / middle / bottom / justify (pool default
    /// top, sdtaitm.hxx), draw:textarea-horizontal-align left / center / right shrink the text block to its lines and place it,
    /// justify (block, the pool default) spans the area; style:writing-mode tb-rl sets the text vertical; draw:text-rotate-angle
    /// turns the text further. The text turns with the shape (not with its shear or mirroring). The text starts from the graphic
    /// style's character formatting.
    /// </summary>
    private void AppendShapeText(OdtDrawnShape d, StringBuilder html, string? pageAttributes = null)
    {
        OdtShapeStyle style = d.Style;
        double Pad(string side) => OdtShapeStyle.Length(style.Get(OdtStyleCascade.Fo + ("padding-" + side))) ?? OdtShapeStyle.Length(style.Get(OdtStyleCascade.Fo + "padding")) ?? 0;
        (double x, double y, double w, double h) = d.TextArea;
        WordAffine transform = WordAffine.Then(WordAffine.Translate(x, y), d.TextTransform);
        if (d.Geometry is { TextRotateAngle: not 0 } g)
        {
            // draw:text-rotate-angle: degrees counter-clockwise about the text area's centre
            WordAffine turn = WordAffine.Then(WordAffine.Then(WordAffine.Translate(-w / 2, -h / 2), WordAffine.Rotate(-g.TextRotateAngle * Math.PI / 180)), WordAffine.Translate(w / 2, h / 2));
            transform = WordAffine.Then(turn, transform);
        }
        string justify = style.Get(OdtStyleCascade.Draw + "textarea-vertical-align") switch { "middle" => "center", "bottom" => "flex-end", "justify" => "space-between", _ => "flex-start" };
        string align = style.Get(OdtStyleCascade.Draw + "textarea-horizontal-align") switch { "left" => "flex-start", "center" => "center", "right" => "flex-end", _ => "stretch" };
        var css = new StringBuilder("position:absolute;left:0;top:0;box-sizing:border-box;margin:0;border:0;overflow:visible;display:flex;flex-direction:column;max-width:none;");
        css.Append("width:").Append(Css(Math.Max(w, 0))).Append("pt;height:").Append(Css(Math.Max(h, 0))).Append("pt;");
        css.Append("padding:").Append(Css(Pad("top"))).Append("pt ").Append(Css(Pad("right"))).Append("pt ").Append(Css(Pad("bottom"))).Append("pt ").Append(Css(Pad("left"))).Append("pt;");
        css.Append("justify-content:").Append(justify).Append(";align-items:").Append(align).Append(';');
        if (!transform.IsIdentity) css.Append("transform-origin:0 0;transform:").Append(transform.Css).Append(';');
        if (style.Get(OdtStyleCascade.Style + "writing-mode") is "tb-rl" or "tb") css.Append("writing-mode:vertical-rl;");
        if (style.Get(OdtStyleCascade.Fo + "wrap-option") == "no-wrap") css.Append("white-space:nowrap;");
        WordRunFormatting text = _styles.GraphicTextFormatting((string?)d.Element.Attribute(DrawNs + "style-name"));
        string textCss = text.ToCss(_styles.DefaultParagraph.Formatting.RunFormatting);
        html.Append("<span class=\"text-box word-shape-text\"").Append(pageAttributes).Append(" style=\"").Append(css).Append(textCss).Append("\">");
        html.Append("<span style=\"display:block;").Append(align == "stretch" ? "" : "width:fit-content;max-width:100%;").Append("\">");
        AppendOwnParagraphs(d.Element, html);
        html.Append("</span></span>");
    }

    /// <summary>
    /// Fontwork (draw:text-path="true", ODF 1.3 part 3 19.172): the shape's text laid on its outlines by the shared text-effect
    /// layout (<see cref="WordOfficeArtFontwork"/>, LibreOffice EnhancedCustomShapeFontWork.cxx), filled with the shape's fill and
    /// outlined with its line (the glyph outlines are path objects with the shape's attributes, EnhancedCustomShapeFontWork.cxx).
    /// Alignment: draw:fit-to-size true / all / shape stretch every line to its outline (TextFitToSize PROPORTIONAL / ALLLINES,
    /// :543-563), else draw:textarea-horizontal-align (SdrTextHorzAdjust, :800-830: left, centre, right; block = from the start).
    /// LibreOffice reads draw:text-path-mode but its Fontwork drawing does not use it (only ScaleX and SameLetterHeights).
    /// </summary>
    private void ApplyFontwork(OdtDrawnShape d)
    {
        var lines = new List<string>();
        string? font = null;
        foreach (XElement paragraph in d.Element.Descendants().Where(e => e.Name == TextNs + "p" || e.Name == TextNs + "h"))
        {
            lines.Add(WordXmlSearchTextPlain(paragraph));
            if (font is null)
            {
                string? paragraphStyle = (string?)paragraph.Attribute(TextNs + "style-name");
                WordRunFormatting run = _styles.ResolveParagraph(paragraphStyle).Formatting.RunFormatting;
                if (paragraph.Descendants(TextNs + "span").FirstOrDefault() is { } span && (string?)span.Attribute(TextNs + "style-name") is { } spanStyle)
                    run = _styles.ApplySpan(run, spanStyle);
                font = run.FontName;
            }
        }
        font ??= _styles.GraphicTextFormatting((string?)d.Element.Attribute(DrawNs + "style-name")).FontName;
        string fit = d.Style.Get(OdtStyleCascade.Draw + "fit-to-size") ?? "false";
        int align = fit is "true" or "all" or "shape" ? 0 : d.Style.Get(OdtStyleCascade.Draw + "textarea-horizontal-align") switch { "center" => 1, "right" => 3, _ => 2 };
        WordShape shape = d.Shape;
        (string Definitions, string Paint)? paint = shape.FillPainter?.Invoke(0);
        string? stroke = shape.LineColor;
        double strokeWidth = shape.LineWidthPt;
        WordOfficeArtFontwork.Apply(shape, string.Join("\n", lines), font, align);
        if (paint is { } p) { shape.FillDefinitions = p.Definitions; shape.TextOnPathColor = p.Paint; }
        else if (shape.FillColor is null && shape.FillPictureId is null) shape.TextOnPathColor = "none";
        shape.FillPainter = null;
        if (stroke is not null) { shape.TextOnPathStroke = stroke; shape.TextOnPathStrokeWidth = strokeWidth; }
    }

    /// <summary>A paragraph's text as the search sees it (spaces, tabs, line breaks as spaces, ruby text and hidden text left out).</summary>
    private static string WordXmlSearchTextPlain(XElement paragraph)
    {
        var text = new StringBuilder();
        void Walk(XElement element)
        {
            foreach (XNode node in element.Nodes())
            {
                if (node is XText t) { text.Append(t.Value); continue; }
                if (node is not XElement e) continue;
                if (e.Name.Namespace != TextNs) continue;
                switch (e.Name.LocalName)
                {
                    case "s": text.Append(' ', int.TryParse((string?)e.Attribute(TextNs + "c"), out int c) && c > 0 ? Math.Min(c, 1000) : 1); break;
                    case "tab": text.Append('\t'); break;
                    case "line-break": text.Append(' '); break;
                    case "ruby-text": case "hidden-text": case "note": case "bookmark-start": break;
                    default: Walk(e); break;
                }
            }
        }
        Walk(paragraph);
        return text.ToString();
    }

    /// <summary>
    /// A text box frame's fill when it is more than a colour (gradient, hatch, bitmap: ODF draw:fill on the frame's graphic style,
    /// Writer frames since LibreOffice 4.4): an SVG layer under the text, filling the box.
    /// </summary>
    private void AppendFrameFill(XElement frame, StringBuilder html, double width, double height)
    {
        var style = new OdtShapeStyle(_styles, (string?)frame.Attribute(DrawNs + "style-name"), GradientAngleIsTenths);
        if (style.Get(OdtStyleCascade.Draw + "fill") is not ("gradient" or "hatch" or "bitmap")) return;
        var shape = new WordShape { Geometry = WordShapeGeometry.Path, Width = width, Height = height, Transform = WordAffine.Identity };
        var path = new WordShapePath { Stroke = false };
        path.Subpaths.Add(OdtShapeBuilder.RoundedRectangle(width, height, 0));
        shape.Paths = [path];
        style.ApplyFill(shape, (0, 0, width, height), FillPictureOf);
        html.Append("<span aria-hidden=\"true\" style=\"position:absolute;left:0;top:0;width:100%;height:100%;z-index:-1;pointer-events:none\">");
        WordShapeHtml.AppendGeometry(html, [shape], width, height, stretches: true);
        html.Append("</span>");
    }

    /// <summary>
    /// A list inside a text box or shape (not part of the body's numbering, which the index counts): each item's label from the
    /// list's style (text:style-name, inherited by nested lists; ODF 1.3 part 3 5.3.1), counted within the list - a nested list
    /// restarts its level; a list style the document does not have gives bullets.
    /// </summary>
    private void RenderBoxList(XElement list, StringBuilder html, string? inheritedStyle, int level, int[] counters)
    {
        string? styleName = (string?)list.Attribute(TextNs + "style-name") ?? inheritedStyle;
        OdtListStyle? listStyle = styleName is null ? null : _styles.ListStyle(styleName) ?? ListStyles?.Invoke(styleName);
        int depth = Math.Min(level, OdtListStyle.MaxLevels - 1);
        for (int i = depth; i < counters.Length; i++) counters[i] = int.MinValue;
        foreach (XElement item in list.Elements())
        {
            bool isItem = item.Name == TextNs + "list-item", isHeader = item.Name == TextNs + "list-header";
            if (!isItem && !isHeader) continue;
            if (isItem)
            {
                OdtListStyle.Level? definition = listStyle?.LevelAt(depth);
                int start = definition?.StartValue ?? 1;
                if (int.TryParse((string?)item.Attribute(TextNs + "start-value"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int given)) counters[depth] = given;
                else counters[depth] = counters[depth] == int.MinValue ? start : counters[depth] + 1;
                for (int i = 0; i < depth; i++) if (counters[i] == int.MinValue) counters[i] = listStyle?.LevelAt(i).StartValue ?? 1;
            }
            bool first = true;
            foreach (XElement child in item.Elements())
            {
                if (child.Name == TextNs + "list") { RenderBoxList(child, html, styleName, level + 1, counters); first = false; continue; }
                if (child.Name != TextNs + "p" && child.Name != TextNs + "h") { RenderBoxBlock(child, html); first = false; continue; }
                string? savedLabel = _listLabel; int savedLevel = _listLevel; string savedLayout = _listLayout;
                string? label = null;
                if (isItem && first)
                    label = listStyle is null ? "•" : listStyle.FormatLabel(depth, counters.Select(c => c == int.MinValue ? 0 : c).ToArray());
                _listLabel = label; _listLevel = depth; _listLayout = listStyle?.LevelAt(depth).Layout ?? "";
                RenderParagraph(child, html, isTopLevel: true);
                _listLabel = savedLabel; _listLevel = savedLevel; _listLayout = savedLayout;
                first = false;
            }
        }
    }

    /// <summary>A block inside a text box or shape: paragraphs, lists (<see cref="RenderBoxList"/>), tables, sections.</summary>
    private void RenderBoxBlock(XElement child, StringBuilder html)
    {
        if (child.Name == TextNs + "p" || child.Name == TextNs + "h") RenderParagraph(child, html, isTopLevel: false);
        else if (child.Name == TextNs + "list") RenderBoxList(child, html, null, 0, new int[OdtListStyle.MaxLevels]);
        else if (child.Name == TableNs + "table") RenderTable(child, html);
        else if (child.Name == TextNs + "section") foreach (XElement inner in child.Elements()) RenderBoxBlock(inner, html);
    }

    /// <summary>True when the paragraph holds a table in a text box: a &lt;table&gt; may not sit inside a &lt;p&gt; (the HTML parser closes the paragraph).</summary>
    private static bool HoldsBoxTable(XElement paragraph)
        => paragraph.Descendants(TableNs + "table").Any(table => table.Ancestors().TakeWhile(a => a != paragraph).Any(a => a.Name == DrawNs + "text-box"));
}
