using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace FastDocView.Html;

/// <summary>
/// Office Math (OMML, ECMA-376 Part 1 22.1) as MathML Core, which Chromium / WebView2 draw natively. The input is an
/// m:oMathPara or m:oMath element (from .docx, from RTF {\mmath} groups rebuilt as OMML, or from a .doc picture's
/// equationXML property); elements are matched by local name. The mapping follows ECMA-376 22.1.2 element by element and
/// Microsoft's OMML2MML.XSL (Office's own OMML-to-MathML stylesheet), checked against LibreOffice's OMML import
/// (starmath/source/ooxmlimport.cxx: handleF 604, handleNary 739, handleD 511, handleAcc 355, handleBar 447, handleR 863,
/// handleRad 993, handleGroupChr 681, handleM 716, handleEqArr 588). MathML Core has no menclose, no mathvariant other
/// than "normal" and no columnalign: a border box is a CSS-bordered mrow, styled letters are mapped to the Mathematical
/// Alphanumeric Symbols (as MathML Core's own text-transform does), and column alignment is CSS text-align on the mtd.
/// </summary>
public sealed class WordOfficeMathHtml
{
    /// <summary>A math run's w:rPr (null: none) → CSS for its token (colour, size, bold of normal text...); null = none.</summary>
    private readonly Func<XElement?, string?>? _runCss;
    private bool _splitsAlignment;     // rendering the direct runs of an m:eqArr row: '&' is an alignment point, not text

    private int _rowDepth;             // rows open: 1 = the equation's own row (where a display equation may break)

    private const char CellBreak = '\u0001';
    private const char ForcedBreak = '\u0002';     // m:brk on a run of the equation's own row: a manual line break before it

    private readonly WordOfficeMathSettings _settings;
    private bool _isDisplay;           // rendering a display equation (m:oMathPara)
    private int _scriptDepth;          // script arguments open (m:sub, m:sup, m:lim, m:deg): operators there lose their spacing

    private WordOfficeMathHtml(Func<XElement?, string?>? runCss, WordOfficeMathSettings settings) { _runCss = runCss; _settings = settings; }

    /// <summary>Whether an element is an Office Math zone (m:oMath / m:oMathPara).</summary>
    public static bool IsMathZone(XElement element) => element.Name.LocalName is "oMath" or "oMathPara" && element.Name.NamespaceName == MathNamespace;

    public const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>
    /// Appends an m:oMathPara (display math: one block line per m:oMath, aligned by m:jc, else the document's m:defJc) or an
    /// inline m:oMath, with the document's math settings (<see cref="WordOfficeMathSettings"/>; null: the defaults).
    /// </summary>
    public static void Append(StringBuilder html, XElement zone, Func<XElement?, string?>? runCss, WordOfficeMathSettings? settings = null)
    {
        settings ??= WordOfficeMathSettings.Default;
        var converter = new WordOfficeMathHtml(runCss, settings);
        // m:mathFont (22.1.2.61): the font equations are drawn in (a run's own w:rFonts still wins)
        string font = settings.Font is { } face ? "font-family:'" + WordHtmlText.CssFontName(face) + "';" : "";
        if (zone.Name.LocalName == "oMathPara")
        {
            // m:oMathParaPr/m:jc (22.1.2.51): left, right, center, centerGroup; without it the document's m:defJc (22.1.2.25)
            string? jc = Val(zone.Element(M("oMathParaPr"))?.Element(M("jc"))) ?? settings.DefaultJustification;
            string align = jc switch { "left" => "left", "right" => "right", _ => "center" };
            converter._isDisplay = true;
            var block = new StringBuilder("display:block;text-align:").Append(align);
            // m:lMargin / m:rMargin (22.1.2.56 / 22.1.2.89): the display equation's own margins; m:wrapIndent (22.1.2.120): a
            // broken equation's continuation lines start that far in (a hanging indent), or with m:wrapRight sit at the right
            double indent = settings.WrapRight ? 0 : settings.WrapIndentPt;
            block.Append(";padding-left:").Append(Points(settings.LeftMarginPt + indent)).Append("pt;padding-right:").Append(Points(settings.RightMarginPt)).Append("pt");
            if (indent > 0) block.Append(";text-indent:-").Append(Points(indent)).Append("pt");
            // m:wrapRight and m:brkBin repeat depend on where the browser wraps: the page finishes them once the lines are known
            // (the app's fastDocViewMath.ts: continuation lines moved to the right edge, the repeated operator shown at a line start)
            string marks = (settings.WrapRight ? " data-math-wrap-right=\"\"" : "") + (settings.BreakBinary == "repeat" ? " data-math-break=\"repeat\"" : "");
            foreach (XElement math in zone.Elements())
            {
                if (math.Name.LocalName != "oMath") continue;
                // dir ltr: an equation reads left to right in a right-to-left paragraph too (its pieces must not flow from the right)
                html.Append("<span class=\"omath-para\" dir=\"ltr\"").Append(marks).Append(" style=\"").Append(block).Append("\">");
                converter.AppendBreakable(html, converter.Row(math), font);
                html.Append("</span>");
            }
            return;
        }
        html.Append("<math class=\"omath\" dir=\"ltr\"").Append(font.Length > 0 ? " style=\"" + font + "\"" : "").Append('>')
            .Append(converter.Row(zone).Replace(ForcedBreak.ToString(), "")).Append("</math>");
    }

    /// <summary>The operator a "repeat" break starts the next line with: itself, or for a minus what m:brkBinSub says.</summary>
    private string RepeatedOperator(string op) => op is "-" or "\u2212" ? _settings.BreakBinarySubtraction == "-+" ? "+" : op : op;

    /// <summary>For brkBinSub "+-" a broken minus ends its line as a plus (the minus starts the next line); null: unchanged.</summary>
    private string? LineEndOperator(string op) => op is "-" or "−" && _settings.BreakBinarySubtraction == "+-" ? "+" : null;

    /// <summary>The math font's line height (em): read from its file, else the values of the fonts Word offers, else 1.2.</summary>
    private double MathLineHeight
    {
        get
        {
            string font = _settings.Font ?? "Cambria Math";
            return WordMathFontReader.LineHeight(font) ?? (font.Equals("Cambria Math", StringComparison.OrdinalIgnoreCase) ? 1.173 : 1.2);
        }
    }

    private static string Points(double points) => points.ToString("0.###", CultureInfo.InvariantCulture);

    /// <summary>
    /// A display equation that can break like Word's: Word breaks a display equation too long for the line at a binary or
    /// relation operator of its own row - before it, after it, or both (m:brkBin before / after / repeat, 22.1.2.15) - and at a
    /// run marked m:brk (22.1.2.16). MathML Core has no line breaking, so the equation's row is cut there into consecutive
    /// math elements the browser may wrap between; m:brk is a forced break. Nested structures are never cut. m:dispDef off
    /// (22.1.2.27) draws the display equation with the inline defaults (no display style).
    /// </summary>
    private void AppendBreakable(StringBuilder html, string row, string font)
    {
        // Row() returns one mrow element; each piece reopens it with its attributes (an m:argSz scriptlevel stays on every piece)
        int first = row.IndexOf('>');
        string rowTag = row.StartsWith("<mrow", StringComparison.Ordinal) && row.EndsWith("</mrow>", StringComparison.Ordinal) && first > 0 ? row.Substring(0, first + 1) : "<mrow>";
        string inner = rowTag.Length > 6 || row.StartsWith("<mrow>", StringComparison.Ordinal) ? row.Substring(rowTag.Length, row.Length - rowTag.Length - 7) : row;
        string open = "<math class=\"omath\" dir=\"ltr\"" + (_settings.DisplayDefaults ? " displaystyle=\"true\" style=\"math-style:normal;" + font + "\"" : font.Length > 0 ? " style=\"" + font + "\"" : "") + ">" + rowTag;
        // brkBin (22.1.2.15): before (the default) - the operator starts the next line; after - it ends the line; repeat - it
        // ends the line and starts the next again (as brkBinSub says for a minus, 22.1.2.16: "--", "-+" or "+-")
        bool after = _settings.BreakBinary is "after" or "repeat", repeats = _settings.BreakBinary == "repeat";
        var cuts = new List<(int at, bool forced)>();
        var repeated = new Dictionary<int, string>();     // cut -> the operator element the next line starts with (repeat)
        int depth = 0;
        for (int i = 0; i < inner.Length; i++)
        {
            char c = inner[i];
            if (c == ForcedBreak && depth == 0) { if (i > 0) cuts.Add((i, true)); continue; }
            if (c != '<') continue;
            int close = inner.IndexOf('>', i);
            if (close < 0) break;
            bool isEnd = inner[i + 1] == '/', isEmpty = inner[close - 1] == '/';
            if (!isEnd && depth == 0 && string.CompareOrdinal(inner, i, "<mo", 0, 3) == 0 && inner[i + 3] is '>' or ' ' && !isEmpty)
            {
                int textEnd = inner.IndexOf("</mo>", close, StringComparison.Ordinal);
                string op = textEnd > close ? System.Net.WebUtility.HtmlDecode(inner.Substring(close + 1, textEnd - close - 1)) : "";
                if (i > 0 && IsBreakOperator(op))
                {
                    int at = after ? textEnd + 5 : i;
                    if (at < inner.Length)
                    {
                        cuts.Add((at, false));
                        if (repeats) repeated[at] = "<mo class=\"omath-repeat\" form=\"infix\" style=\"display:none\"" + (LineEndOperator(op) is { } end ? " data-line-end=\"" + WordHtmlText.Encode(end) + "\"" : "") + ">" + WordHtmlText.Encode(RepeatedOperator(op)) + "</mo>";
                    }
                }
            }
            if (isEnd) depth--; else if (!isEmpty) depth++;
            i = close;
        }
        int start = 0;
        html.Append(open);
        foreach ((int at, bool forced) in cuts)
        {
            if (at <= start) continue;      // nothing before this cut (an operator right after a forced break)
            html.Append(inner, start, at - start).Append("</mrow></math>");
            html.Append(forced ? "<br>" : "\u200B").Append(open);
            if (!forced && repeated.TryGetValue(at, out string? again)) html.Append(again);
            start = forced ? at + 1 : at;
            // the operator now opens a math element: keep it an infix operator (its spacing as in the unbroken row)
            if (!after && start + 3 < inner.Length && string.CompareOrdinal(inner, start, "<mo", 0, 3) == 0 && inner[start + 3] is '>' or ' ') { html.Append("<mo form=\"infix\""); start += 3; }
        }
        html.Append(inner, start, inner.Length - start).Append("</mrow></math>");
    }

    /// <summary>Binary and relation operators Word breaks before (Unicode Technical Note 28 section 3.3; the common ones).</summary>
    private static bool IsBreakOperator(string op) => op is "+" or "-" or "\u2212" or "\u00B1" or "\u2213" or "=" or "<" or ">" or "\u2264" or "\u2265" or "\u2260"
        or "\u2248" or "\u2261" or "\u00D7" or "\u22C5" or "\u00B7" or "\u2192" or "\u21D2" or "\u2208" or "\u2282" or "\u2286" or "\u222A" or "\u2229";

    /// <summary>
    /// The text an equation shows, for search: what <see cref="Append"/> draws - m:t text (an m:eqArr row's own '&' are
    /// alignment marks, not text), w:sym characters, m:br as a space; deletions hidden, a plain w:r's w:t read.
    /// </summary>
    public static string SearchText(XElement zone)
    {
        var text = new StringBuilder();
        AppendSearchText(text, zone, inEquationArrayRow: false);
        return text.ToString();
    }

    private static void AppendSearchText(StringBuilder text, XElement element, bool inEquationArrayRow)
    {
        foreach (XElement child in element.Elements())
        {
            string name = child.Name.LocalName;
            if (child.Name.NamespaceName != MathNamespace)
            {
                if (name is "ins" or "moveTo" or "sdt" or "sdtContent" or "smartTag" or "customXml" or "hyperlink") AppendSearchText(text, child, false);
                else if (name == "r") foreach (XElement t in child.Elements().Where(e => e.Name.LocalName == "t")) text.Append(t.Value);
                continue;
            }
            if (name == "r") { text.Append(RunText(child, inEquationArrayRow)); continue; }
            if (name.EndsWith("Pr", StringComparison.Ordinal)) continue;
            AppendSearchText(text, child, name == "e" && element.Name.LocalName == "eqArr");
        }
    }

    /// <summary>A math run's characters as drawn (m:t, m:br a space, w:sym its character); '&' dropped in an eqArr row's own runs.</summary>
    private static string RunText(XElement run, bool dropsAlignment)
    {
        var text = new StringBuilder();
        foreach (XElement child in run.Elements())
        {
            if (child.Name.LocalName == "t") text.Append(dropsAlignment && child.Name.NamespaceName == MathNamespace ? child.Value.Replace("&", "") : child.Value);
            else if (child.Name.LocalName == "br") text.Append(' ');
            else if (child.Name.LocalName == "sym" && child.Attributes().FirstOrDefault(a => a.Name.LocalName == "char")?.Value is { } code
                     && int.TryParse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int symbol))
                text.Append(SymbolCharacter(symbol));
        }
        return text.ToString();
    }

    /// <summary>A w:sym character (22.1 runs take WordprocessingML's w:sym, 17.3.3.30): a symbol-font code F0xx is drawn as its low byte.</summary>
    internal static string SymbolCharacter(int symbol)
    {
        int code = symbol >= 0xF000 && symbol <= 0xF0FF ? symbol - 0xF000 : symbol;
        return code is > 0 and <= 0x10FFFF and (< 0xD800 or > 0xDFFF) ? char.ConvertFromUtf32(code) : "\uFFFD";
    }

    /// <summary>A math run's w:rPr read without a style cascade (a .doc picture's equation XML, RTF): its w:color only.</summary>
    public static string? RunColorCss(XElement? runProperties)
    {
        string? color = runProperties?.Elements().FirstOrDefault(e => e.Name.LocalName == "color") is { } element ? Val(element) : null;
        return color is { Length: 6 } && color.All(Uri.IsHexDigit) ? "color:#" + color.ToUpperInvariant() : null;
    }

    private static XName M(string name) => XName.Get(name, MathNamespace);

    private static string? Val(XElement? element)
        => element?.Attributes().FirstOrDefault(a => a.Name.LocalName == "val")?.Value;

    /// <summary>An on/off property (22.9.2.7 ST_OnOff): present without a value = on.</summary>
    private static bool IsOn(XElement? element) => element is not null && Val(element) is null or "1" or "on" or "true";

    private static XElement? Child(XElement element, string name) => element.Elements().FirstOrDefault(e => e.Name.LocalName == name);

    private static XElement? Property(XElement element, string properties, string name)
        => Child(element, properties) is { } pr ? Child(pr, name) : null;

    /// <summary>An argument (m:e, m:num, m:sub...) or any container: its children in one mrow.</summary>
    private string Row(XElement? container)
    {
        if (container is null) return "<mrow></mrow>";
        // m:argPr/m:argSz (22.1.2.6): -2..2 steps smaller / larger than the argument's own size = MathML scriptlevel steps
        var html = new StringBuilder("<mrow");
        if (int.TryParse(Val(Property(container, "argPr", "argSz")), NumberStyles.Integer, CultureInfo.InvariantCulture, out int size) && size is >= -2 and <= 2 and not 0)
            html.Append(" scriptlevel=\"").Append(size < 0 ? "+" + (-size) : "-" + size).Append('"');
        html.Append('>');
        bool isScript = container.Name.LocalName is "sub" or "sup" or "lim" or "deg";
        _rowDepth++;
        if (isScript) _scriptDepth++;
        try { foreach (XElement child in container.Elements()) AppendElement(html, child); }
        finally { _rowDepth--; if (isScript) _scriptDepth--; }
        return html.Append("</mrow>").ToString();
    }

    private void AppendElement(StringBuilder html, XElement element)
    {
        string name = element.Name.LocalName;
        if (element.Name.NamespaceName != MathNamespace)
        {
            // WordprocessingML inside math: insertions shown, deletions hidden (Word's final view); a plain run's text as text
            switch (name)
            {
                case "ins": case "moveTo": case "sdt": case "sdtContent": case "smartTag": case "customXml": case "hyperlink":
                    foreach (XElement child in element.Elements()) AppendElement(html, child);
                    break;
                case "r":
                    string text = string.Concat(element.Elements().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
                    if (text.Length > 0) AppendToken(html, "mtext", text, _runCss?.Invoke(Child(element, "rPr")), null);
                    break;
            }
            return;
        }
        bool splits = _splitsAlignment;
        if (name != "r") _splitsAlignment = false;     // only an eqArr row's own runs carry its alignment points
        try { AppendMathElement(html, element, name); }
        finally { _splitsAlignment = splits; }
    }

    private void AppendMathElement(StringBuilder html, XElement element, string name)
    {
        switch (name)
        {
            case "r": AppendRun(html, element); break;
            case "oMath": case "box": html.Append(Row(element)); break;    // m:box (22.1.2.13): a grouping with no visible mark
            case "f": AppendFraction(html, element); break;
            case "rad":
            {
                // 22.1.2.86: m:degHide or an empty m:deg = a square root
                XElement? degree = Child(element, "deg");
                if (IsOn(Property(element, "radPr", "degHide")) || degree is null || !degree.Elements().Any(e => e.Name.LocalName != "ctrlPr" && e.Name.LocalName != "argPr"))
                    html.Append("<msqrt>").Append(Row(Child(element, "e"))).Append("</msqrt>");
                else html.Append("<mroot>").Append(Row(Child(element, "e"))).Append(Row(degree)).Append("</mroot>");
                break;
            }
            case "sSup": html.Append("<msup>").Append(Row(Child(element, "e"))).Append(Row(Child(element, "sup"))).Append("</msup>"); break;
            case "sSub": html.Append("<msub>").Append(Row(Child(element, "e"))).Append(Row(Child(element, "sub"))).Append("</msub>"); break;
            case "sSubSup":
                html.Append("<msubsup>").Append(Row(Child(element, "e"))).Append(Row(Child(element, "sub"))).Append(Row(Child(element, "sup"))).Append("</msubsup>");
                break;
            case "sPre":     // 22.1.2.99: scripts before the base
                html.Append("<mmultiscripts>").Append(Row(Child(element, "e"))).Append("<mprescripts/>")
                    .Append(Row(Child(element, "sub"))).Append(Row(Child(element, "sup"))).Append("</mmultiscripts>");
                break;
            case "nary": AppendNary(html, element); break;
            case "d": AppendDelimiter(html, element); break;
            case "func":     // 22.1.2.39: the name, a function application, the argument
                html.Append("<mrow>").Append(Row(Child(element, "fName"))).Append("<mo>&#x2061;</mo>").Append(Row(Child(element, "e"))).Append("</mrow>");
                break;
            case "limLow": html.Append("<munder>").Append(Row(Child(element, "e"))).Append(Row(Child(element, "lim"))).Append("</munder>"); break;
            case "limUpp": html.Append("<mover>").Append(Row(Child(element, "e"))).Append(Row(Child(element, "lim"))).Append("</mover>"); break;
            case "acc":
            {
                // 22.1.2.1: m:chr (default U+0302 combining circumflex) over the base; a combining mark is drawn as its spacing form
                string accent = SpacingAccent(Val(Property(element, "accPr", "chr")) ?? "̂");
                html.Append("<mover accent=\"true\">").Append(Row(Child(element, "e"))).Append("<mo stretchy=\"true\">").Append(WordHtmlText.Encode(accent)).Append("</mo></mover>");
                break;
            }
            case "bar":
                // 22.1.2.7: m:pos top = an overbar, bot (the default, ooxmlimport.cxx:450) = an underbar
                if (Val(Property(element, "barPr", "pos")) == "top")
                    html.Append("<mover accent=\"true\">").Append(Row(Child(element, "e"))).Append("<mo stretchy=\"true\">&#x203E;</mo></mover>");
                else html.Append("<munder accentunder=\"true\">").Append(Row(Child(element, "e"))).Append("<mo stretchy=\"true\">_</mo></munder>");
                break;
            case "borderBox": AppendBorderBox(html, element); break;
            case "eqArr":
            {
                // 22.1.2.34: one row per m:e; '&' in a row's runs = alignment points (cells alternate right / left, as Word aligns them)
                html.Append("<mtable class=\"omath-eqarr\">");
                foreach (XElement row in element.Elements().Where(e => e.Name.LocalName == "e"))
                {
                    html.Append("<mtr>");
                    string[] cells = AlignedRow(row).Split(CellBreak);
                    for (int i = 0; i < cells.Length; i++)
                    {
                        string align = cells.Length == 1 ? "center" : i % 2 == 0 ? "right" : "left";
                        html.Append("<mtd style=\"text-align:").Append(align).Append(";padding:0\"><mrow>").Append(cells[i]).Append("</mrow></mtd>");
                    }
                    html.Append("</mtr>");
                }
                html.Append("</mtable>");
                break;
            }
            case "m": AppendMatrix(html, element); break;
            case "groupChr":
            {
                // 22.1.2.41: m:chr (default U+23DF bottom curly bracket) under (m:pos bot, the default) or over the base
                XElement? pr = Child(element, "groupChrPr");
                string character = (pr is null ? null : Val(Child(pr, "chr"))) ?? "⏟";
                bool isTop = (pr is null ? null : Val(Child(pr, "pos"))) == "top";
                string tag = isTop ? "mover" : "munder";
                html.Append('<').Append(tag).Append('>').Append(Row(Child(element, "e"))).Append("<mo stretchy=\"true\">").Append(WordHtmlText.Encode(character)).Append("</mo></").Append(tag).Append('>');
                break;
            }
            case "phant":
            {
                // 22.1.2.81: m:show off hides the base but keeps its space; zeroWid / zeroAsc / zeroDesc remove a dimension
                XElement? pr = Child(element, "phantPr");
                string body = Row(Child(element, "e"));
                bool show = pr is null || Child(pr, "show") is not { } showElement || IsOn(showElement);
                if (!show) body = "<mphantom>" + body + "</mphantom>";
                var padded = new StringBuilder();
                if (pr is not null && IsOn(Child(pr, "zeroWid"))) padded.Append(" width=\"0\"");
                if (pr is not null && IsOn(Child(pr, "zeroAsc"))) padded.Append(" height=\"0\"");
                if (pr is not null && IsOn(Child(pr, "zeroDesc"))) padded.Append(" depth=\"0\"");
                html.Append(padded.Length > 0 ? "<mpadded" + padded + ">" + body + "</mpadded>" : body);
                break;
            }
            // properties (…Pr, m:ctrlPr, m:argPr) draw nothing; an argument met on its own is a row
            case "e": case "num": case "den": case "sub": case "sup": case "deg": case "lim": case "fName": html.Append(Row(element)); break;
        }
    }

    private string AlignedRow(XElement row)
    {
        bool splits = _splitsAlignment;
        _splitsAlignment = true;
        try
        {
            var html = new StringBuilder();
            foreach (XElement child in row.Elements()) AppendElement(html, child);
            return html.ToString();
        }
        finally { _splitsAlignment = splits; }
    }

    private string FractionStyle => (_isDisplay && _settings.SmallFractions ? " displaystyle=\"false\"" : "") + " style=\"padding-inline:0\"";

    private void AppendFraction(StringBuilder html, XElement element)
    {
        // 22.1.2.117 ST_FType: bar (stacked, the default), skw (skewed), lin (linear a/b), noBar (stacked, no line: binomial)
        string num = Row(Child(element, "num")), den = Row(Child(element, "den"));
        switch (Val(Property(element, "fPr", "type")))
        {
            // Word sets the slash of a linear / skewed fraction tight (math-v2.pdf: "a/b" 20.3 pt); MathML's '/' has 1/6 em each side
            case "lin": html.Append("<mrow>").Append(num).Append("<mo lspace=\"0\" rspace=\"0\">/</mo>").Append(den).Append("</mrow>"); break;
            case "skw":
                // MathML Core has no bevelled fraction: the numerator raised, a slash, the denominator lowered
                html.Append("<mrow><msup><mrow></mrow>").Append(num).Append("</msup><mo stretchy=\"false\" lspace=\"0\" rspace=\"0\">/</mo><msub><mrow></mrow>").Append(den).Append("</msub></mrow>");
                break;
            // Word draws an inline equation's fraction small (its numerator and denominator at script size: math-v2.pdf, 8.4 pt
            // in 12 pt text) - MathML's own inline rule - and a display equation's at full size unless m:smallFrac (22.1.2.98).
            // MathML Core pads a fraction by 1px each side (its user-agent sheet); Word's bar is as wide as its widest part.
            case "noBar": html.Append("<mfrac linethickness=\"0\"").Append(FractionStyle).Append('>').Append(num).Append(den).Append("</mfrac>"); break;
            default: html.Append("<mfrac").Append(FractionStyle).Append('>').Append(num).Append(den).Append("</mfrac>"); break;
        }
    }

    private void AppendNary(StringBuilder html, XElement element)
    {
        // 22.1.2.70: m:chr (default U+222B integral); m:limLoc undOvr (limits under / over) or subSup (as scripts); without
        // it Word's defaults (m:intLim subSup for integrals, m:naryLim undOvr for the rest, 22.1.2.49 / 22.1.2.72)
        XElement? pr = Child(element, "naryPr");
        string character = (pr is null ? null : Val(Child(pr, "chr"))) ?? "∫";
        string? location = pr is null ? null : Val(Child(pr, "limLoc"));
        bool isIntegral = character.Length == 1 && character[0] is >= '∫' and <= '∳';
        // without m:limLoc: the document's m:intLim (integrals) / m:naryLim (the rest), 22.1.2.49 / 22.1.2.72
        bool isUnderOver = (location ?? (isIntegral ? _settings.IntegralLimits : _settings.NaryLimits)) == "undOvr";
        bool hasSub = !(pr is not null && IsOn(Child(pr, "subHide"))), hasSup = !(pr is not null && IsOn(Child(pr, "supHide")));
        // m:grow (22.1.2.40, off by default for n-ary operators): the operator grows to its argument's height - a stretchy mo
        // in the row of the operator and its argument
        bool grows = pr is not null && Child(pr, "grow") is { } grow && IsOn(grow);
        // undOvr limits move beside the operator in an inline equation, as Word draws them (math-v2.pdf: an inline sum is 15.8 pt
        // tall, limits as scripts): MathML's movablelimits does exactly that outside display style
        string op = "<mo largeop=\"true\"" + (grows ? " stretchy=\"true\" symmetric=\"true\"" : "") + ">" + WordHtmlText.Encode(character) + "</mo>";
        if (grows) { var stretched = new StringBuilder(); AppendStretched(stretched, op, character, true); op = stretched.ToString(); }
        string sub = Row(Child(element, "sub")), sup = Row(Child(element, "sup"));
        html.Append("<mrow>");
        if (hasSub && hasSup) html.Append(isUnderOver ? "<munderover>" : "<msubsup>").Append(op).Append(sub).Append(sup).Append(isUnderOver ? "</munderover>" : "</msubsup>");
        else if (hasSub) html.Append(isUnderOver ? "<munder>" : "<msub>").Append(op).Append(sub).Append(isUnderOver ? "</munder>" : "</msub>");
        else if (hasSup) html.Append(isUnderOver ? "<mover>" : "<msup>").Append(op).Append(sup).Append(isUnderOver ? "</mover>" : "</msup>");
        else html.Append(op);
        html.Append(Row(Child(element, "e"))).Append("</mrow>");
    }

    private void AppendDelimiter(StringBuilder html, XElement element)
    {
        // 22.1.2.24: m:begChr "(" / m:sepChr "|" / m:endChr ")" by default; an empty value draws nothing
        XElement? pr = Child(element, "dPr");
        string begin = (pr is null ? null : Val(Child(pr, "begChr"))) ?? "(";
        string separator = (pr is null ? null : Val(Child(pr, "sepChr"))) ?? "|";
        string end = (pr is null ? null : Val(Child(pr, "endChr"))) ?? ")";
        // m:grow (22.1.2.40): the delimiters grow with their content; on unless turned off (LibreOffice: handleD reads it, ooxmlimport.cxx:511)
        // Word grows them only past one line of text: Chromium stretching a Cambria Math fence to a plain row picks a larger
        // variant (9 pt for "(" at 12 pt against Word's 5.3: math-v2.pdf "(a+b)" 38.0 pt wide, ours 51.4 stretched, 38.3 not)
        var arguments = element.Elements().Where(e => e.Name.LocalName == "e").Select(Row).ToList();
        bool isTall = arguments.Any(IsTallerThanALine);
        string stretchy = pr is not null && Child(pr, "grow") is { } grow && !IsOn(grow) || !isTall ? "false" : "true";
        html.Append("<mrow>");
        if (begin.Length > 0) AppendStretched(html, "<mo fence=\"true\" form=\"prefix\" stretchy=\"" + stretchy + "\">" + WordHtmlText.Encode(begin) + "</mo>", begin, stretchy == "true");
        bool first = true;
        foreach (string argument in arguments)
        {
            if (!first && separator.Length > 0) AppendStretched(html, "<mo separator=\"true\" stretchy=\"" + stretchy + "\">" + WordHtmlText.Encode(separator) + "</mo>", separator, stretchy == "true");
            html.Append(argument);
            first = false;
        }
        if (end.Length > 0) AppendStretched(html, "<mo fence=\"true\" form=\"postfix\" stretchy=\"" + stretchy + "\">" + WordHtmlText.Encode(end) + "</mo>", end, stretchy == "true");
        html.Append("</mrow>");
    }

    /// <summary>
    /// A stretched operator drawn at the width of the glyph Word picks. Word draws a grown delimiter or n-ary operator with the
    /// font's OpenType MATH vertical size variants, else its glyph assembly (MathVariants); Chromium stretches with the same
    /// table but its stretched glyph box is wider (measured: Cambria Math "(" 9.0 pt against Word's 7.3 around three rows).
    /// LibreOffice Math uses neither: SmMathSymbolNode::AdaptToY (starmath/source/node.cxx) scales the glyph height and keeps
    /// its width. So the operator carries the font's variants (height:width em, then the widest assembly
    /// part; read from the installed font's MATH table by WordMathFontReader, else Html/WordMathFontVariants.cs) and the page (the app's fastDocViewMath.ts)
    /// narrows it (negative margins, centred on the glyph) to the variant Word picks for the height the browser stretched it to. A font
    /// without a table there keeps the browser's own width.
    /// </summary>
    private void AppendStretched(StringBuilder html, string mo, string character, bool stretches)
    {
        string font = _settings.Font ?? "Cambria Math";
        // the installed font's own MATH table, else the tables generated for the fonts Word offers (a host without font files)
        var table = WordMathFontReader.Variants(font) ?? (WordMathFontVariants.Fonts.TryGetValue(font, out var generated) ? generated : null);
        if (!stretches || table is null || !table.TryGetValue(character, out var variants))
        {
            html.Append(mo);
            return;
        }
        // on the mo itself: Chromium stops stretching an operator wrapped in an mpadded (measured: a "(" around ten rows stayed
        // one line tall); the page narrows the mo with negative margins instead
        bool isLarge = character.Length == 1 && character[0] is >= '∀' and <= '⋿' or >= '⨀' and <= '⫿';
        var data = new StringBuilder(" class=\"omath-stretch\"").Append(isLarge ? " data-stretch=\"nary\"" : "").Append(" data-variants=\"");
        for (int i = 0; i < variants.heights.Length; i++)
            data.Append(i > 0 ? " " : "").Append(Points(variants.heights[i])).Append(':').Append(Points(variants.widths[i]));
        data.Append('|').Append(Points(variants.assemblyWidth)).Append('"');
        html.Append(mo.Insert(3, data.ToString()));
    }

    /// <summary>Whether rendered MathML stacks vertically (fractions, limits, tables, radicals, scripts): taller than a text line.</summary>
    private static bool IsTallerThanALine(string mathMl)
        => mathMl.Contains("<mfrac", StringComparison.Ordinal) || mathMl.Contains("<mtable", StringComparison.Ordinal) || mathMl.Contains("<munder", StringComparison.Ordinal)
        || mathMl.Contains("<mover", StringComparison.Ordinal) || mathMl.Contains("<msqrt", StringComparison.Ordinal) || mathMl.Contains("<mroot", StringComparison.Ordinal)
        || mathMl.Contains("<msup", StringComparison.Ordinal) || mathMl.Contains("<msub", StringComparison.Ordinal) || mathMl.Contains("<mmultiscripts", StringComparison.Ordinal)
        || mathMl.Contains("largeop", StringComparison.Ordinal);

    private void AppendBorderBox(StringBuilder html, XElement element)
    {
        // 22.1.2.11: a box drawn around the base; m:hideTop/Bot/Left/Right drop a side, m:strikeH/V/BLTR/TLBR strike through
        XElement? pr = Child(element, "borderBoxPr");
        bool Has(string name) => pr is not null && IsOn(Child(pr, name));
        var css = new StringBuilder("padding:0.1em;border:1px solid currentColor");
        if (Has("hideTop")) css.Append(";border-top:none");
        if (Has("hideBot")) css.Append(";border-bottom:none");
        if (Has("hideLeft")) css.Append(";border-left:none");
        if (Has("hideRight")) css.Append(";border-right:none");
        var strikes = new List<string>();
        const string line = "transparent calc(50% - 0.5px),currentColor calc(50% - 0.5px),currentColor calc(50% + 0.5px),transparent calc(50% + 0.5px)";
        if (Has("strikeH")) strikes.Add("linear-gradient(to bottom," + line + ")");
        if (Has("strikeV")) strikes.Add("linear-gradient(to right," + line + ")");
        if (Has("strikeBLTR")) strikes.Add("linear-gradient(to top left," + line + ")");
        if (Has("strikeTLBR")) strikes.Add("linear-gradient(to top right," + line + ")");
        if (strikes.Count > 0) css.Append(";background-image:").Append(string.Join(",", strikes));
        html.Append("<mrow class=\"omath-border-box\" style=\"").Append(css).Append("\">").Append(Row(Child(element, "e"))).Append("</mrow>");
    }

    private void AppendMatrix(StringBuilder html, XElement element)
    {
        // 22.1.2.60: m:mr rows of m:e cells; m:mcs/m:mc/m:mcPr (m:count columns, m:mcJc left / center / right) per column group
        var alignments = new List<string>();
        if (Property(element, "mPr", "mcs") is { } columns)
        {
            foreach (XElement column in columns.Elements().Where(e => e.Name.LocalName == "mc"))
            {
                XElement? pr = Child(column, "mcPr");
                int count = int.TryParse(pr is null ? null : Val(Child(pr, "count")), NumberStyles.Integer, CultureInfo.InvariantCulture, out int c) ? Math.Max(1, Math.Min(c, 64)) : 1;
                string align = (pr is null ? null : Val(Child(pr, "mcJc"))) switch { "left" => "left", "right" => "right", _ => "center" };
                for (int i = 0; i < count; i++) alignments.Add(align);
            }
        }
        html.Append("<mtable class=\"omath-matrix\">");
        foreach (XElement row in element.Elements().Where(e => e.Name.LocalName == "mr"))
        {
            // Word sets a matrix's rows one line of the math font apart (m:rSpRule 0, single spacing, ECMA-376 22.1.2.90; its
            // PDFs: 1.171 em in Cambria Math, 1.2 em in DejaVu Math TeX Gyre = each font's hhea line height): a row is at
            // least that tall
            html.Append("<mtr style=\"height:").Append(Points(MathLineHeight)).Append("em\">");
            int index = 0;
            foreach (XElement cell in row.Elements().Where(e => e.Name.LocalName == "e"))
            {
                string align = index < alignments.Count ? alignments[index] : "center";
                // Word's default column gap (m:cGp 1 = a gap of about half an em, matched to Word's PDF of math-v2) as padding on each side
                html.Append("<mtd style=\"text-align:").Append(align).Append(";padding:0 0.25em\">").Append(Row(cell)).Append("</mtd>");
                index++;
            }
            html.Append("</mtr>");
        }
        html.Append("</mtable>");
    }

    /// <summary>A combining accent (U+0300..U+036F, U+20D0..U+20FF) as the spacing character MathML draws over a base.</summary>
    internal static string SpacingAccent(string accent) => accent switch
    {
        "̀" => "`", "́" => "´", "̂" => "ˆ", "̃" => "˜", "̄" => "¯", "̅" => "‾",
        "̆" => "˘", "̇" => "˙", "̈" => "¨", "̊" => "˚", "̌" => "ˇ",
        "⃐" => "↼", "⃑" => "⇀", "⃖" => "←", "⃗" => "→", "⃡" => "↔", "⃛" => "⋯",
        _ => accent,
    };

    // ── runs ──

    private enum Script : byte { Roman, Script, Fraktur, DoubleStruck, SansSerif, Monospace }

    private void AppendRun(StringBuilder html, XElement run)
    {
        // 22.1.2.91 m:rPr: m:nor = normal (non-math) text; m:sty p / b / i / bi (default: italic letters, 22.1.2.105 / 22.1.2.111);
        // m:scr roman / script / fraktur / double-struck / sans-serif / monospace; m:lit only affects linear-format input
        XElement? mathProperties = Child(run, "rPr");
        string? css = _runCss?.Invoke(run.Elements().FirstOrDefault(e => e.Name.LocalName == "rPr" && e.Name.NamespaceName != MathNamespace));
        var text = new StringBuilder();
        foreach (XElement child in run.Elements())
        {
            if (child.Name.LocalName == "t") text.Append(child.Value);
            else if (child.Name.LocalName == "br") text.Append(' ');
            else if (child.Name.LocalName == "sym" && child.Attributes().FirstOrDefault(a => a.Name.LocalName == "char")?.Value is { } code
                     && int.TryParse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int symbol))
                text.Append(SymbolCharacter(symbol));
        }
        if (text.Length == 0) return;
        if (_rowDepth == 1 && mathProperties is not null && Child(mathProperties, "brk") is not null) html.Append(ForcedBreak);
        if (mathProperties is not null && IsOn(Child(mathProperties, "nor")))
        {
            AppendSplit(html, text.ToString(), part => AppendToken(html, "mtext", part, css, null));
            return;
        }
        string? style = mathProperties is null ? null : Val(Child(mathProperties, "sty"));
        Script script = (mathProperties is null ? null : Val(Child(mathProperties, "scr"))) switch
        {
            "script" => Script.Script, "fraktur" => Script.Fraktur, "double-struck" => Script.DoubleStruck,
            "sans-serif" => Script.SansSerif, "monospace" => Script.Monospace, _ => Script.Roman,
        };
        bool bold = style is "b" or "bi", italic = style is null or "i" or "bi";
        bool isScript = _scriptDepth > 0;
        AppendSplit(html, text.ToString(), part => AppendMathText(html, part, script, bold, italic, style == "p", css, isScript));
    }

    /// <summary>Text split at an eqArr row's alignment points ('&' there is a mark, not a character).</summary>
    private void AppendSplit(StringBuilder html, string text, Action<string> append)
    {
        if (!_splitsAlignment || text.IndexOf('&') < 0) { append(text); return; }
        string[] parts = text.Split('&');
        for (int i = 0; i < parts.Length; i++)
        {
            if (i > 0) html.Append(CellBreak);
            if (parts[i].Length > 0) append(parts[i]);
        }
    }

    /// <summary>Math text as tokens: numbers (digits with inner . or ,) mn, letters mi, the rest mo; spaces kept as space.</summary>
    private static void AppendMathText(StringBuilder html, string text, Script script, bool bold, bool italic, bool isPlain, string? css, bool isScript = false)
    {
        bool isStyled = bold || script != Script.Roman;
        for (int i = 0; i < text.Length;)
        {
            char c = text[i];
            int end = i + 1;
            if (char.IsHighSurrogate(c) && end < text.Length) end++;
            if (char.IsDigit(c))
            {
                while (end < text.Length && (char.IsDigit(text[end]) || text[end] is '.' or ',' && end + 1 < text.Length && char.IsDigit(text[end + 1]))) end++;
                AppendToken(html, "mn", Styled(text.Substring(i, end - i), script, bold, false), css, isStyled ? "normal" : null);
            }
            else if (char.IsLetter(c) || char.IsSurrogate(c))
            {
                // a plain (m:sty p) run of letters is one upright identifier (sin, lim); otherwise every letter is a variable
                if (isPlain) while (end < text.Length && char.IsLetter(text[end])) end++;
                string letters = text.Substring(i, end - i);
                if (isStyled || italic && !isPlain && letters.Length > 1) AppendToken(html, "mi", Styled(letters, script, bold, italic), css, "normal");
                else AppendToken(html, "mi", letters, css, !italic || isPlain ? "normal" : null);
            }
            else if (c == ' ') html.Append("<mspace width=\"0.25em\"></mspace>");
            // in a script (limits, sub- and superscripts) Word sets operators without their spacing (as TeX does in script
            // style: math-v2.pdf draws "i=1" under a sum 14.5 pt wide at 8.4 pt, no room around '=')
            else if (isScript) AppendToken(html, "mo", text.Substring(i, end - i), css, null, " lspace=\"0\" rspace=\"0\"");
            else AppendToken(html, "mo", text.Substring(i, end - i), css, null);
            i = end;
        }
    }

    private static void AppendToken(StringBuilder html, string tag, string text, string? css, string? mathVariant, string? attributes = null)
    {
        html.Append('<').Append(tag).Append(attributes);
        if (mathVariant is not null) html.Append(" mathvariant=\"").Append(mathVariant).Append('"');
        if (!string.IsNullOrEmpty(css)) html.Append(" style=\"").Append(css).Append('"');
        html.Append('>').Append(WordHtmlText.Encode(text)).Append("</").Append(tag).Append('>');
    }

    // ── Mathematical Alphanumeric Symbols (Unicode U+1D400..U+1D7FF; the letters encoded earlier in Letterlike Symbols are the holes) ──

    private static readonly Dictionary<int, int> LetterlikeHoles = new()
    {
        [0x1D455] = 0x210E,                                                                   // italic h
        [0x1D49D] = 0x212C, [0x1D4A0] = 0x2130, [0x1D4A1] = 0x2131, [0x1D4A3] = 0x210B, [0x1D4A4] = 0x2110, [0x1D4A7] = 0x2112,
        [0x1D4A8] = 0x2133, [0x1D4AD] = 0x211B, [0x1D4BA] = 0x212F, [0x1D4BC] = 0x210A, [0x1D4C4] = 0x2134,   // script B E F H I L M R e g o
        [0x1D506] = 0x212D, [0x1D50B] = 0x210C, [0x1D50C] = 0x2111, [0x1D515] = 0x211C, [0x1D51D] = 0x2128,   // fraktur C H I R Z
        [0x1D53A] = 0x2102, [0x1D53F] = 0x210D, [0x1D545] = 0x2115, [0x1D547] = 0x2119, [0x1D548] = 0x211A,
        [0x1D549] = 0x211D, [0x1D551] = 0x2124,                                               // double-struck C H N P Q R Z
    };

    /// <summary>A MathML mathvariant (MathML 3 section 3.2.2) applied to text as Mathematical Alphanumeric Symbols; null for an unknown variant.</summary>
    public static string? StyledText(string text, string mathVariant) => mathVariant switch
    {
        "bold" => Styled(text, Script.Roman, true, false),
        "italic" => Styled(text, Script.Roman, false, true),
        "bold-italic" => Styled(text, Script.Roman, true, true),
        "double-struck" => Styled(text, Script.DoubleStruck, false, false),
        "script" => Styled(text, Script.Script, false, false),
        "bold-script" => Styled(text, Script.Script, true, false),
        "fraktur" => Styled(text, Script.Fraktur, false, false),
        "bold-fraktur" => Styled(text, Script.Fraktur, true, false),
        "sans-serif" => Styled(text, Script.SansSerif, false, false),
        "bold-sans-serif" => Styled(text, Script.SansSerif, true, false),
        "sans-serif-italic" => Styled(text, Script.SansSerif, false, true),
        "sans-serif-bold-italic" => Styled(text, Script.SansSerif, true, true),
        "monospace" => Styled(text, Script.Monospace, false, false),
        _ => null,
    };

    private static string Styled(string text, Script script, bool bold, bool italic)
    {
        int? latin = (script, bold, italic) switch
        {
            (Script.Roman, true, false) => 0x1D400, (Script.Roman, false, true) => 0x1D434, (Script.Roman, true, true) => 0x1D468,
            (Script.Script, false, _) => 0x1D49C, (Script.Script, true, _) => 0x1D4D0,
            (Script.Fraktur, false, _) => 0x1D504, (Script.Fraktur, true, _) => 0x1D56C,
            (Script.DoubleStruck, _, _) => 0x1D538,
            (Script.SansSerif, false, false) => 0x1D5A0, (Script.SansSerif, true, false) => 0x1D5D4,
            (Script.SansSerif, false, true) => 0x1D608, (Script.SansSerif, true, true) => 0x1D63C,
            (Script.Monospace, _, _) => 0x1D670,
            _ => null,
        };
        int? digits = script switch
        {
            Script.Roman when bold => 0x1D7CE, Script.DoubleStruck => 0x1D7D8, Script.SansSerif => bold ? 0x1D7EC : 0x1D7E2, Script.Monospace => 0x1D7F6, _ => null,
        };
        int? greek = (script, bold, italic) switch
        {
            (Script.Roman, true, false) => 0x1D6A8, (Script.Roman, false, true) => 0x1D6E2, (Script.Roman, true, true) => 0x1D71C, _ => null,
        };
        var result = new StringBuilder();
        foreach (char c in text)
        {
            int mapped = c;
            if (latin is { } l && c is >= 'A' and <= 'Z') mapped = l + (c - 'A');
            else if (latin is { } s && c is >= 'a' and <= 'z') mapped = s + 26 + (c - 'a');
            else if (digits is { } d && c is >= '0' and <= '9') mapped = d + (c - '0');
            else if (greek is { } g && c is >= 'Α' and <= 'Ω' && c != '΢') mapped = g + (c - 0x391);
            else if (greek is { } h && c is >= 'α' and <= 'ω') mapped = h + 26 + (c - 0x3B1);
            if (LetterlikeHoles.TryGetValue(mapped, out int hole)) mapped = hole;
            if (mapped == c) result.Append(c); else result.Append(char.ConvertFromUtf32(mapped));
        }
        return result.ToString();
    }
}
