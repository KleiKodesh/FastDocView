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

    private const char CellBreak = '\u0001';

    private WordOfficeMathHtml(Func<XElement?, string?>? runCss) => _runCss = runCss;

    /// <summary>Whether an element is an Office Math zone (m:oMath / m:oMathPara).</summary>
    public static bool IsMathZone(XElement element) => element.Name.LocalName is "oMath" or "oMathPara" && element.Name.NamespaceName == MathNamespace;

    public const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    /// <summary>Appends an m:oMathPara (display math: one block line per m:oMath, aligned by m:jc) or an inline m:oMath.</summary>
    public static void Append(StringBuilder html, XElement zone, Func<XElement?, string?>? runCss)
    {
        var converter = new WordOfficeMathHtml(runCss);
        if (zone.Name.LocalName == "oMathPara")
        {
            // m:oMathParaPr/m:jc (22.1.2.51): left, right, center, centerGroup (the default; a group's lines centred as one block)
            string? jc = Val(zone.Element(M("oMathParaPr"))?.Element(M("jc")));
            string align = jc switch { "left" => "left", "right" => "right", _ => "center" };
            foreach (XElement math in zone.Elements())
            {
                if (math.Name.LocalName != "oMath") continue;
                html.Append("<span class=\"omath-para\" style=\"display:block;text-align:").Append(align).Append("\">");
                html.Append("<math class=\"omath\" dir=\"ltr\" displaystyle=\"true\" style=\"math-style:normal\">");
                html.Append(converter.Row(math));
                html.Append("</math></span>");
            }
            return;
        }
        html.Append("<math class=\"omath\" dir=\"ltr\">").Append(converter.Row(zone)).Append("</math>");
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
        var html = new StringBuilder("<mrow>");
        foreach (XElement child in container.Elements()) AppendElement(html, child);
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

    private void AppendFraction(StringBuilder html, XElement element)
    {
        // 22.1.2.117 ST_FType: bar (stacked, the default), skw (skewed), lin (linear a/b), noBar (stacked, no line: binomial)
        string num = Row(Child(element, "num")), den = Row(Child(element, "den"));
        switch (Val(Property(element, "fPr", "type")))
        {
            case "lin": html.Append("<mrow>").Append(num).Append("<mo>/</mo>").Append(den).Append("</mrow>"); break;
            case "skw":
                // MathML Core has no bevelled fraction: the numerator raised, a slash, the denominator lowered
                html.Append("<mrow><msup><mrow></mrow>").Append(num).Append("</msup><mo stretchy=\"false\">/</mo><msub><mrow></mrow>").Append(den).Append("</msub></mrow>");
                break;
            case "noBar": html.Append("<mfrac linethickness=\"0\">").Append(num).Append(den).Append("</mfrac>"); break;
            default: html.Append("<mfrac>").Append(num).Append(den).Append("</mfrac>"); break;
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
        bool isUnderOver = location is null ? !isIntegral : location == "undOvr";
        bool hasSub = !(pr is not null && IsOn(Child(pr, "subHide"))), hasSup = !(pr is not null && IsOn(Child(pr, "supHide")));
        string op = "<mo largeop=\"true\"" + (isUnderOver ? " movablelimits=\"false\"" : "") + ">" + WordHtmlText.Encode(character) + "</mo>";
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
        html.Append("<mrow>");
        if (begin.Length > 0) html.Append("<mo fence=\"true\" form=\"prefix\" stretchy=\"true\">").Append(WordHtmlText.Encode(begin)).Append("</mo>");
        bool first = true;
        foreach (XElement argument in element.Elements().Where(e => e.Name.LocalName == "e"))
        {
            if (!first && separator.Length > 0) html.Append("<mo separator=\"true\" stretchy=\"true\">").Append(WordHtmlText.Encode(separator)).Append("</mo>");
            html.Append(Row(argument));
            first = false;
        }
        if (end.Length > 0) html.Append("<mo fence=\"true\" form=\"postfix\" stretchy=\"true\">").Append(WordHtmlText.Encode(end)).Append("</mo>");
        html.Append("</mrow>");
    }

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
            html.Append("<mtr>");
            int index = 0;
            foreach (XElement cell in row.Elements().Where(e => e.Name.LocalName == "e"))
            {
                string align = index < alignments.Count ? alignments[index] : "center";
                html.Append("<mtd style=\"text-align:").Append(align).Append("\">").Append(Row(cell)).Append("</mtd>");
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
                text.Append((symbol >= 0xF000 ? symbol - 0xF000 : symbol) is var symbolCode && symbolCode is >= 0 and <= 0x10FFFF and (< 0xD800 or > 0xDFFF) ? char.ConvertFromUtf32(symbolCode) : "�");
        }
        if (text.Length == 0) return;
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
        AppendSplit(html, text.ToString(), part => AppendMathText(html, part, script, bold, italic, style == "p", css));
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
    private static void AppendMathText(StringBuilder html, string text, Script script, bool bold, bool italic, bool isPlain, string? css)
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
            else AppendToken(html, "mo", text.Substring(i, end - i), css, null);
            i = end;
        }
    }

    private static void AppendToken(StringBuilder html, string tag, string text, string? css, string? mathVariant)
    {
        html.Append('<').Append(tag);
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
