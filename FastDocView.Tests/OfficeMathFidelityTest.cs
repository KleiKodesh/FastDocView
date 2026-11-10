using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Equations (Office Math, ECMA-376 Part 1 22.1) as MathML: test-data/word-math/math-v2.docx / .rtf / .doc, saved by Word 16
/// from one document with a paragraph per OMML construct (scripts/make-math-base-docx.py, then make-math-fixtures.ps1 adds an
/// equation in a table cell, a footnote and a text box; Word's PDF of it is math-v2.pdf). .docx keeps m:oMath, .rtf
/// {\mmath} groups, .doc a picture per equation whose OfficeArt equationXML property (MS-ODRAW 2.3.2.10, opid 0x030C) holds the
/// OMML; all three must give the same MathML structure and their equation text must be found by search.
/// </summary>
public static class OfficeMathFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Equations: OMML as MathML (.docx, .rtf, .doc saved by Word 16)");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-math"));
        foreach (string extension in new[] { ".docx", ".rtf", ".doc" })
        {
            string path = Path.Combine(folder, "math-v2" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-math/math-v2{extension} present"); continue; }
            WordSaved(path, extension);
        }
        string odtFolder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "odt-math"));
        foreach (string extension in new[] { ".odt", ".fodt" })
        {
            string path = Path.Combine(odtFolder, "odt-math-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/odt-math/odt-math-v1{extension} present"); continue; }
            LibreOfficeSaved(path, extension);
        }
        RtfFieldEquations();
        ArgumentSize(Path.Combine(folder, "math-base.docx"));
        MathSettings(Path.Combine(folder, "math-v2.rtf"));
    }

    private static readonly System.Xml.Linq.XNamespace MathNs = Html.WordOfficeMathHtml.MathNamespace;

    private static string Render(string omml, Html.WordOfficeMathSettings settings)
    {
        var zone = System.Xml.Linq.XElement.Parse("<m:oMathPara xmlns:m=\"" + MathNs.NamespaceName + "\">" + omml + "</m:oMathPara>");
        var html = new StringBuilder();
        Html.WordOfficeMathHtml.Append(html, zone.Elements().Count() == 1 && zone.Elements().First().Name.LocalName == "oMathPara" ? zone.Elements().First() : zone, null, settings);
        return System.Net.WebUtility.HtmlDecode(html.ToString());
    }

    private static string InlineRender(string omml, Html.WordOfficeMathSettings settings)
    {
        var zone = System.Xml.Linq.XElement.Parse("<m:oMath xmlns:m=\"" + MathNs.NamespaceName + "\">" + omml + "</m:oMath>");
        var html = new StringBuilder();
        Html.WordOfficeMathHtml.Append(html, zone, null, settings);
        return System.Net.WebUtility.HtmlDecode(html.ToString());
    }

    /// <summary>
    /// m:mathPr (22.1.2.62) read from settings.xml and applied: mathFont, brkBin after, smallFrac, dispDef off, margins,
    /// wrapIndent / wrapRight, defJc, intLim / naryLim; and the RTF {\mmathPr} of Word 16's math-v2.rtf read into the same settings.
    /// </summary>
    private static void MathSettings(string rtfPath)
    {
        var settings = Html.WordOfficeMathSettings.FromOoxml(System.Xml.Linq.XElement.Parse(
            "<m:mathPr xmlns:m=\"" + MathNs.NamespaceName + "\"><m:mathFont m:val=\"STIX Two Math\"/><m:brkBin m:val=\"after\"/><m:smallFrac/><m:dispDef m:val=\"0\"/>"
            + "<m:lMargin m:val=\"360\"/><m:rMargin m:val=\"720\"/><m:defJc m:val=\"left\"/><m:wrapIndent m:val=\"720\"/><m:intLim m:val=\"undOvr\"/><m:naryLim m:val=\"subSup\"/></m:mathPr>"));
        const string run = "<m:r><m:t>a+b</m:t></m:r>";
        const string fraction = "<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>y</m:t></m:r></m:den></m:f>";
        string display = Render("<m:oMath>" + run + fraction + "</m:oMath>", settings);
        TestReport.Check(display.Contains("text-align:left;padding-left:54pt;padding-right:36pt;text-indent:-36pt", StringComparison.Ordinal), "m:mathPr: defJc left, lMargin 18 pt + wrapIndent 36 pt as a hanging indent, rMargin 36 pt");
        TestReport.Check(display.Contains("font-family:'STIX Two Math'", StringComparison.Ordinal) && !display.Contains("displaystyle=\"true\"", StringComparison.Ordinal), "m:mathPr: mathFont on the equation; dispDef off = no display style");
        TestReport.Check(display.Contains("<mfrac displaystyle=\"false\"", StringComparison.Ordinal), "m:mathPr: smallFrac = small fractions in display equations");
        TestReport.Check(display.Contains("<mo>+</mo></mrow></math>\u200B", StringComparison.Ordinal), "m:mathPr: brkBin after = the break point after the operator");
        string sum = InlineRender("<m:nary><m:naryPr><m:chr m:val=\"\u2211\"/></m:naryPr><m:sub/><m:sup/><m:e/></m:nary><m:nary><m:sub/><m:sup/><m:e/></m:nary>", settings);
        TestReport.Check(sum.Contains("<msubsup><mo largeop=\"true\">\u2211", StringComparison.Ordinal) && sum.Contains("<munderover><mo largeop=\"true\">\u222B", StringComparison.Ordinal), "m:mathPr: naryLim subSup, intLim undOvr as the defaults without m:limLoc");
        string defaults = Render("<m:oMath>" + run + fraction + "</m:oMath>", Html.WordOfficeMathSettings.Default);
        TestReport.Check(defaults.Contains("text-align:center;padding-left:72pt;padding-right:0pt;text-indent:-72pt", StringComparison.Ordinal) && defaults.Contains("displaystyle=\"true\"", StringComparison.Ordinal)
            && defaults.Contains("\u200B<math", StringComparison.Ordinal) && defaults.Contains("<mo form=\"infix\">+", StringComparison.Ordinal), "m:mathPr defaults: centred, wrapIndent 1 inch, display style, break before the operator");
        string right = Render("<m:oMath>" + run + "</m:oMath>", Html.WordOfficeMathSettings.FromOoxml(System.Xml.Linq.XElement.Parse("<m:mathPr xmlns:m=\"" + MathNs.NamespaceName + "\"><m:wrapRight/></m:mathPr>")));
        TestReport.Check(right.Contains("padding-left:0pt", StringComparison.Ordinal) && right.Contains("text-align-last:right", StringComparison.Ordinal), "m:mathPr: wrapRight = continuation lines at the right, no indent");
        // the RTF header's {\mmathPr} (Word 16: \mmathFont34 = Cambria Math, \mdefJc1, \mwrapIndent1440)
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-math-settings-rtf") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(rtfPath, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(b => b.Html));
        TestReport.Check(html.Contains("<math class=\"omath\" dir=\"ltr\" style=\"font-family:'Cambria Math';\">", StringComparison.Ordinal)
            && html.Contains("padding-left:72pt", StringComparison.Ordinal), ".rtf: {\\mmathPr} read: the math font and wrap indent applied");
    }

    /// <summary>m:argSz (22.1.2.6): Word 16 drops it from m:num when it re-saves, so it is checked on the hand-written base file.</summary>
    private static void ArgumentSize(string path)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-math-base") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(b => b.Html)).Replace(" style=\"font-family:'Cambria Math';\"", "");
        TestReport.Check(html.Contains("<mrow scriptlevel=\"+1\"><mi>s</mi>", StringComparison.Ordinal), "math-base.docx: m:argSz -1: the argument one script level smaller");
    }

    /// <summary>
    /// Hand-written RTF: an equation inside a field's code is not shown and not searched (the parser reads a field instruction
    /// as no text destination), one in the field's result is both - the renderer and both search readers agree.
    /// </summary>
    private static void RtfFieldEquations()
    {
        const string math = @"{\mmath{\*\moMath{\mr\mscr0\msty2 WORD}}}";
        string rtf = @"{\rtf1\ansi\deff0{\fonttbl{\f0 Arial;}}" + "\n"
            + @"\pard before {\field{\*\fldinst { QUOTE " + math.Replace("WORD", "codeeq") + @" }}{\fldrslt {" + math.Replace("WORD", "resulteq") + @"}}} after\par" + "\n}";
        string path = Path.Combine(TestReport.OutputDirectory, "math-field-equations-v1.rtf");
        File.WriteAllText(path, rtf);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-math-field-rtf") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(b => System.Net.WebUtility.HtmlDecode(b.Html)));
        TestReport.Check(!html.Contains("<mi>c</mi><mi>o</mi><mi>d</mi><mi>e</mi>", StringComparison.Ordinal), ".rtf: an equation in a field's code is not shown");
        TestReport.Check(html.Contains("<mi>r</mi><mi>e</mi><mi>s</mi><mi>u</mi><mi>l</mi><mi>t</mi>", StringComparison.Ordinal), ".rtf: an equation in a field's result is shown");
        TestReport.Check(document.Search("codeeq", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 0, ".rtf: an equation in a field's code is not searched");
        TestReport.Check(document.Search("resulteq", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length > 0, ".rtf: an equation in a field's result is searched");
        if (document is Rtf.RtfWindowedDocument windowed)
        {
            TestReport.Check(windowed.SearchByBlockReparse("codeeq", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 0
                && windowed.SearchByBlockReparse("resulteq", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length > 0, ".rtf: the per-block reference search agrees");
        }
    }

    /// <summary>
    /// LibreOffice Math objects (scripts/make-odt-math.py: one formula per paragraph; LibreOffice's PDF is odt-math-v1.pdf):
    /// the MathML of "Object N/content.xml" (.odt) or inline in the draw:object (.fodt) drawn instead of the preview picture,
    /// the StarMath annotation not shown, mstyle variants and colours mapped, the formula text searched.
    /// </summary>
    private static void LibreOfficeSaved(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-math-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => System.Net.WebUtility.HtmlDecode(block.Html)).ToList();
        void Expect(string label, Func<string, bool> check, string what)
        {
            string block = html.FirstOrDefault(b => Strip(b).TrimStart().StartsWith("formula " + label + ":", StringComparison.Ordinal)) ?? "";
            TestReport.Check(block.Length > 0 && block.Contains("<math class=\"odf-formula\"", StringComparison.Ordinal) && !block.Contains("word-picture", StringComparison.Ordinal) && check(block),
                $"{format}: formula {label}: {what}{(block.Length == 0 ? " (block missing)" : "")}");
        }
        Expect("fraction", b => b.Contains("<mfrac><mi>a</mi><mi>b</mi></mfrac>", StringComparison.Ordinal) && !b.Contains("a over b", StringComparison.Ordinal), "a fraction, the StarMath annotation hidden");
        Expect("root", b => b.Contains("<msqrt>", StringComparison.Ordinal) && b.Contains("<mroot>", StringComparison.Ordinal), "square and cube roots");
        Expect("scripts", b => b.Contains("<msup>", StringComparison.Ordinal) && b.Contains("<msub>", StringComparison.Ordinal) && b.Contains("<msubsup>", StringComparison.Ordinal), "scripts");
        Expect("sum", b => b.Contains("<munderover>", StringComparison.Ordinal) && b.Contains("\u2211", StringComparison.Ordinal), "a sum with limits");
        Expect("integral", b => b.Contains("\u222B", StringComparison.Ordinal), "an integral");
        Expect("matrix", b => b.Contains("<mtable", StringComparison.Ordinal) && b.Contains("<mtd>", StringComparison.Ordinal), "a matrix");
        Expect("styles", b => b.Contains("\U0001D401", StringComparison.Ordinal) && b.Contains("color:red", StringComparison.Ordinal) && b.Contains("<mtext>text here</mtext>", StringComparison.Ordinal), "bold mapped, colour as CSS, text");
        Expect("accents", b => b.Contains("<mover", StringComparison.Ordinal), "accents over");
        WordDocumentSearchBatch result = document.Search("text here", WordDocumentSearchOptions.Default, 0, document.BlockCount);
        TestReport.Check(result.HitBlocks.Length > 0, $"{format}: search finds formula text");
    }

    private static string Strip(string html) => WordFieldsFidelityTest.Strip(html);

    private static void WordSaved(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-math-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount).ToList();
        // the math font every Word math run names (w:rFonts Cambria Math) is left out of the structure checks
        List<string> html = blocks.Select(block => System.Net.WebUtility.HtmlDecode(block.Html).Replace(" style=\"font-family:'Cambria Math';\"", "")).ToList();
        if (Environment.GetEnvironmentVariable("MATH_DUMP") is { Length: > 0 } dump) File.WriteAllText(Path.Combine(dump, "math" + format + ".html"), string.Join("\n", html) + "\nNOTES " + string.Concat(blocks.SelectMany(b => b.Notes).Select(n => n.Html)));
        string notes = string.Concat(blocks.SelectMany(block => block.Notes).Select(note => System.Net.WebUtility.HtmlDecode(note.Html))).Replace(" style=\"font-family:'Cambria Math';\"", "");
        string Block(string start) => html.FirstOrDefault(block => Strip(block).TrimStart().StartsWith(start, StringComparison.Ordinal)) ?? "";
        void Expect(string label, Func<string, bool> check, string what)
        {
            string block = Block("case " + label + ":");
            TestReport.Check(block.Length > 0 && check(block), $"{format}: {label}: {what}{(block.Length == 0 ? " (block missing)" : "")}");
        }
        static bool Has(string block, params string[] parts) => parts.All(part => block.Contains(part, StringComparison.Ordinal));
        static bool Math(string block) => block.Contains("<math class=\"omath\"", StringComparison.Ordinal);

        Expect("frac bar", b => Math(b) && Has(b, "<mfrac style=\"padding-inline:0\"><mrow><mi>a</mi></mrow><mrow><mi>b</mi></mrow></mfrac>"), "a stacked fraction");
        Expect("frac skw", b => Has(b, "<msup><mrow></mrow><mrow><mi>a</mi></mrow></msup><mo stretchy=\"false\" lspace=\"0\" rspace=\"0\">/</mo>"), "a skewed fraction");
        Expect("frac lin", b => Has(b, "<mi>a</mi></mrow><mo lspace=\"0\" rspace=\"0\">/</mo><mrow><mi>b</mi>") && !b.Contains("<mfrac", StringComparison.Ordinal), "a linear fraction");
        Expect("frac noBar", b => Has(b, "<mfrac linethickness=\"0\" style=\"padding-inline:0\">"), "a fraction without a bar");
        Expect("sqrt", b => Has(b, "<msqrt><mrow><mi>x</mi></mrow></msqrt>"), "a square root (degHide)");
        Expect("root", b => Has(b, "<mroot><mrow><mi>y</mi></mrow><mrow><mn>3</mn></mrow></mroot>"), "a cube root");
        Expect("sup", b => Has(b, "<msup><mrow><mi>x</mi></mrow><mrow><mn>2</mn></mrow></msup>"), "a superscript");
        Expect("sub", b => Has(b, "<msub><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow></msub>"), "a subscript");
        Expect("subsup", b => Has(b, "<msubsup><mrow><mi>x</mi></mrow><mrow><mi>i</mi></mrow><mrow><mn>2</mn></mrow></msubsup>"), "both scripts");
        Expect("pre", b => Has(b, "<mmultiscripts><mrow><mi>Y</mi></mrow><mprescripts/><mrow><mn>1</mn></mrow><mrow><mi>n</mi></mrow></mmultiscripts>"), "pre-scripts");
        Expect("integral", b => Has(b, "<msubsup><mo largeop=\"true\">∫</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow></msubsup>"), "an integral with limits as scripts");
        Expect("sum", b => Has(b, "<munderover><mo largeop=\"true\">∑</mo>"), "a sum with limits under and over");
        Expect("product", b => Has(b, "<msub><mo largeop=\"true\">∏</mo><mrow><mi>k</mi></mrow></msub>") && !b.Contains("<msubsup", StringComparison.Ordinal), "a product, upper limit hidden");
        Expect("delim", b => Has(b, ">[</mo>", "separator=\"true\" stretchy=\"false\">;</mo>", "form=\"postfix\" stretchy=\"false\">}</mo>"), "begin / separator / end characters");
        Expect("paren", b => Has(b, "form=\"prefix\" stretchy=\"false\">(</mo>", "form=\"postfix\" stretchy=\"false\">)</mo>"), "default parentheses, not stretched around one line");
        Expect("func", b => Has(b, "<mi mathvariant=\"normal\">sin</mi>", "<mo>⁡</mo>", "<mi>θ</mi>"), "a function: upright name, apply, argument");
        Expect("limlow", b => Has(b, "<munder><mrow><mi mathvariant=\"normal\">lim</mi></mrow><mrow><mi>n</mi><mo lspace=\"0\" rspace=\"0\">→</mo><mi>∞</mi></mrow></munder>") || Has(b, "<munder><mrow><mi mathvariant=\"normal\">lim</mi></mrow><mrow><mi>n</mi><mo lspace=\"0\" rspace=\"0\">→</mo><mo lspace=\"0\" rspace=\"0\">∞</mo></mrow></munder>"), "a lower limit");
        Expect("limupp", b => Has(b, "<mover><mrow><mi>x</mi></mrow><mrow><mi>d</mi><mi>e</mi><mi>f</mi></mrow></mover>"), "an upper limit");
        Expect("acc", b => Has(b, "<mover accent=\"true\"><mrow><mi>a</mi></mrow><mo stretchy=\"true\">˜</mo></mover>"), "a tilde accent (combining U+0303 drawn spacing)");
        Expect("acc hat", b => Has(b, "<mo stretchy=\"true\">ˆ</mo>"), "the default accent: a circumflex");
        Expect("bar top", b => Has(b, "<mover accent=\"true\"><mrow><mi>z</mi></mrow><mo stretchy=\"true\">‾</mo></mover>"), "an overbar");
        Expect("bar bot", b => Has(b, "<munder accentunder=\"true\"><mrow><mi>w</mi></mrow>"), "an underbar (the default)");
        Expect("box", b => Has(b, "<mrow><mrow><mi>q</mi></mrow></mrow>") || Has(b, "<mi>q</mi>"), "a box: just its content");
        Expect("borderbox", b => Has(b, "class=\"omath-border-box\" style=\"padding:0.1em;border:1px solid currentColor\"", "<msup>"), "a bordered box");
        Expect("eqarr", b => Has(b, "<mtable class=\"omath-eqarr\">", "<mtd style=\"text-align:right;padding:0\"><mrow><mi>x</mi></mrow></mtd><mtd style=\"text-align:left;padding:0\"><mrow><mo>=</mo><mn>1</mn></mrow></mtd>") && !Strip(b).Contains('&'), "an equation array aligned at '&'");
        Expect("matrix", b => Has(b, "<mtable class=\"omath-matrix\"><mtr><mtd style=\"text-align:center;padding:0 0.25em\"><mrow><mn>1</mn></mrow></mtd>"), "a 2x2 matrix");
        Expect("groupchr", b => Has(b, "<munder><mrow><mi>a</mi><mo>+</mo>", "<mo stretchy=\"true\">⏟</mo></munder>"), "a group character under (the default)");
        Expect("groupchr top", b => Has(b, "<mover><mrow><mi>k</mi></mrow><mo stretchy=\"true\">⏞</mo></mover>"), "a group character over");
        Expect("phant", b => Has(b, "<mphantom><mrow><mi>b</mi><mi>b</mi><mi>b</mi></mrow></mphantom>"), "a hidden phantom keeps its space");
        Expect("styles", b => Has(b, "<mi mathvariant=\"normal\">P</mi>", "\U0001D401", "<mi>I</mi>", "\U0001D47F"), "plain, bold, italic, bold italic");
        Expect("scripts", b => Has(b, "ℝ", "\U0001D509", "ℒ", "\U0001D61A", "\U0001D67C"), "double-struck, fraktur, script, sans-serif, monospace letters");
        Expect("normal text", b => Has(b, "<mtext>if x</mtext>", "<mi>l</mi><mi>i</mi><mi>t</mi>"), "normal text as text, m:lit ignored");
        // Word 16's RTF writes a math w:sym as a SYMBOL field with code 63 ('?'): the character is lost in the file itself
        if (format != ".rtf") Expect("symbol", b => Has(b, "<mo>\u2200</mo>"), "a w:sym in a math run drawn as its character");
        Expect("grow", b => Has(b, "stretchy=\"true\" symmetric=\"true\">\u222B</mo>", "form=\"prefix\" stretchy=\"false\">(</mo>"), "m:grow: the integral grows, the delimiters (grow off) do not");
        Expect("colour", b => Has(b, "#FF0000") && Has(b, "<mi") , "a coloured run");

        string all = string.Join("\n", html);
        TestReport.Check(Has(all, "omath-para\" style=\"display:block;text-align:left"), $"{format}: oMathPara jc left");
        TestReport.Check(Has(all, "omath-para\" style=\"display:block;text-align:right"), $"{format}: oMathPara jc right");
        // Word 16 joins the two m:oMath of one m:oMathPara into one when it saves (math-v2.docx has one): a centred line
        TestReport.Check(Has(all, "omath-para\" style=\"display:block;text-align:center;padding-left:72pt"), $"{format}: oMathPara centred (centerGroup, the default)");
        TestReport.Check(html.Any(b => b.Contains("<mi>t</mi><mi>a</mi><mi>b</mi><mi>l</mi><mi>e</mi>", StringComparison.Ordinal) && b.Contains("<mfrac", StringComparison.Ordinal)), $"{format}: an equation in a table cell");
        TestReport.Check(notes.Contains("<msup><mrow><mi>x</mi></mrow><mrow><mn>2</mn></mrow></msup>", StringComparison.Ordinal), $"{format}: an equation in a footnote");
        TestReport.Check(html.Any(b => b.Contains("<mi>b</mi><mi>o</mi><mi>x</mi>", StringComparison.Ordinal) && b.Contains("<msqrt>", StringComparison.Ordinal)), $"{format}: an equation in a text box");
        TestReport.Check(html.All(b => !b.Contains("[image]", StringComparison.Ordinal)), $"{format}: no equation left as an unshown picture");

        string breaks = html.FirstOrDefault(b => b.Contains("<mi>b</mi><mi>r</mi><mi>k</mi>", StringComparison.Ordinal)) ?? "";
        // Word 16 drops m:brk from the equation XML it keeps in a .doc
        if (format != ".doc") TestReport.Check(breaks.Contains("</math><br><math", StringComparison.Ordinal), $"{format}: m:brk: a manual line break in a display equation");
        string longEquation = html.FirstOrDefault(b => b.Contains("<mi>l</mi><mi>o</mi><mi>n</mi><mi>g</mi>", StringComparison.Ordinal)) ?? "";
        TestReport.Check(longEquation.Split(new[] { "<math class=\"omath\"" }, StringSplitOptions.None).Length >= 40 && longEquation.Contains("\u200B<math", StringComparison.Ordinal),
            $"{format}: a long display equation can break before its operators (one piece per term)");
        foreach (string query in new[] { "x=1", "y=22" })
            TestReport.Check(document.Search(query, WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length > 0, $"{format}: search reads an equation array row without its '&' alignment marks (\"{query}\")");
        TestReport.Check(document.Search("&=", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 0, $"{format}: search does not find the hidden '&'");
        foreach (string query in format == ".rtf" ? new[] { "tablemath", "notemath", "boxmath", "sin", "red" } : new[] { "tablemath", "notemath", "boxmath", "sin", "red", "forall\u2200" })
        {
            WordDocumentSearchBatch result = document.Search(query, WordDocumentSearchOptions.Default, 0, document.BlockCount);
            TestReport.Check(result.HitBlocks.Length > 0, $"{format}: search finds equation text \"{query}\"");
        }
    }
}
