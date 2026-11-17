namespace FastDocView.Tests;

/// <summary>
/// Phonetic guide in .doc and RTF: test-data/word-ruby/ruby-v1 (Word 16, make-ruby-fixtures.ps1) saves each guide as
/// "EQ \* jcN \* "Font:Arial" \* hpsN \o(\s\up 11(guideN),baseN)"; it is drawn as HTML ruby like .docx's w:ruby, the guide
/// in its font at hps / 2 pt (Word's PDF: ArialMT 10 / 8 / 10 / 12 / 10 pt over the base), ruby-align from jc.
/// ODT bullet fonts named by style:font-name: test-data/odt-bullet-fonts/bullet-font-name-v1 (.fodt hand-written, .odt packed by
/// scripts/make-odt-bullet-font-name.py) - the font-face name resolves to its svg:font-family as LibreOffice does
/// (xmlnumi.cxx:645-694; its PDF draws Wingdings U+F0A7 and SymbolMT U+F0B7).
/// </summary>
public static class RubyAndBulletFontFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Ruby in .doc / RTF (EQ phonetic guide); ODT bullet style:font-name resolved through font-face-decls");
        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));
        foreach (string extension in new[] { ".doc", ".rtf", ".docx" })
        {
            string path = Path.Combine(root, "word-ruby", "ruby-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-ruby/ruby-v1{extension} present"); continue; }
            Ruby(path, extension);
        }
        Escapes(root);
        foreach (string extension in new[] { ".fodt", ".odt" })
        {
            string path = Path.Combine(root, "odt-bullet-fonts", "bullet-font-name-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/odt-bullet-fonts/bullet-font-name-v1{extension} present"); continue; }
            BulletFontNames(path, extension);
        }
        TestReport.Check(Fields.WordFieldRuby.ToHtml(@" \* jc0 \o(\s\up 11(a),b)") == "<ruby>b<rt>a</rt></ruby>"
                         && Fields.WordFieldRuby.ToHtml(@" \o(\s\up 11(a),b)") is null && Fields.WordFieldRuby.ToHtml(@" \* jc0 \f(1,2)") is null,
                         "EQ: only Word's phonetic-guide form (\\* options, then \\o with \\s\\up) is ruby");
    }

    private static void Ruby(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-ruby-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string[] sizes = ["10", "8", "10", "12", "10"];
        string?[] aligns = [null, "space-between", "space-around", "start", "end"];
        for (int i = 0; i < 5; i++)
        {
            string block = blocks.FirstOrDefault(html => html.Contains("ruby case " + i)) ?? "";
            string alignCss = aligns[i] is { } align ? "ruby-align:" + align + ";" : "";
            bool isRuby = block.Contains("<ruby>base" + i + "<rt") && block.Contains("guide" + i) && !block.Contains("field-equation");
            // .docx carries the guide's size and font on its runs, its alignment on w:rubyAlign; .doc / RTF all three on the EQ switches
            bool isStyled = format == ".docx"
                ? block.Contains("<rt" + (alignCss.Length > 0 ? " style=\"" + alignCss + "\"" : "") + "><span style=\"" + (sizes[i] == "12" ? "" : "font-size:" + sizes[i] + "pt;") + "font-family:'Arial';\">guide" + i + "</span></rt>")   // 12 pt: the paragraph's size, no run size
                : block.Contains("<rt style=\"font-size:" + sizes[i] + "pt;font-family:'Arial';" + alignCss + "\">guide" + i + "</rt>");
            TestReport.Check(isRuby && isStyled, $"{format}: ruby case {i}: base{i} with guide{i} above as HTML ruby ({sizes[i]} pt Arial{(aligns[i] is null ? "" : ", " + aligns[i])})");
        }
    }

    /// <summary>
    /// Brackets, commas, semicolons, backslashes and spaces in a guide or base, and unquoted multi-word fonts, read as Word's EQ
    /// parser does (Word 16 PDFs ruby-escapes-v1.pdf / ruby-escapes-probe-v1.pdf: the guide ends at the first comma, every ")"
    /// is dropped, a backslash escapes nothing).
    /// </summary>
    private static void Escapes(string root)
    {
        // (file, case, guide as Word's PDF draws it, base as Word's PDF draws it, guide font)
        string b = "\\";
        (string File, int Case, string Guide, string Base, string Font)[] cases =
        [
            ("ruby-escapes-v1.doc", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-v1.doc", 1, "g", "1,base,1", "Courier New"),
            ("ruby-escapes-v1.doc", 2, "g;2", "base;2", "Arial"), ("ruby-escapes-v1.doc", 3, "g" + b + "3", "base" + b + "3", "Arial"),
            ("ruby-escapes-v1.doc", 4, "g4(", "base4(", "Calibri"), ("ruby-escapes-v1.doc", 5, "g 5 x", "base 5", "Calibri"),
            ("ruby-escapes-v1.rtf", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-v1.rtf", 1, "g", "1,base,1", "Courier New"),
            ("ruby-escapes-v1.rtf", 2, "g;2", "base;2", "Arial"), ("ruby-escapes-v1.rtf", 3, "g" + b + "3", "base" + b + "3", "Arial"),
            ("ruby-escapes-v1.rtf", 4, "g4(", "base4(", "Calibri"), ("ruby-escapes-v1.rtf", 5, "g 5 x", "base 5", "Calibri"),
            ("ruby-escapes-probe-v1.rtf", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-probe-v1.rtf", 1, "g", "1,base,1", "Courier New"),
            ("ruby-escapes-probe-v1.rtf", 2, "a" + b, "b,c" + b + ",d", "Arial"), ("ruby-escapes-probe-v1.rtf", 3, "e" + b + "f", "g" + b + "h", "Arial"),
            ("ruby-escapes-probe-v1.rtf", 4, "i" + b + "(j", "k" + b + b + "l", "Calibri"),
        ];
        foreach (var group in cases.GroupBy(c => c.File))
        {
            string path = Path.Combine(root, "word-ruby", group.Key);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-ruby/{group.Key} present"); continue; }
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-ruby-escapes-" + group.Key.Replace('.', '-')) };
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
            foreach (var c in group)
            {
                string block = blocks.FirstOrDefault(html => html.Contains("ruby escape " + c.Case + ":")) ?? "";
                string expected = "<ruby>" + Html.WordHtmlText.Encode(c.Base) + "<rt style=\"font-size:10pt;font-family:'" + c.Font + "';\">"
                                  + Html.WordHtmlText.Encode(c.Guide) + "</rt></ruby>";
                TestReport.Check(block.Contains(expected), $"{group.Key}: ruby escape {c.Case}: guide and base as Word's PDF draws them, {c.Font}");
            }
        }
        TestReport.Check(Fields.WordFieldRuby.ToHtml(@" \* jc0 \* Font:MS Mincho \* hps20 \o(\s\up 11(a),b)") == "<ruby>b<rt style=\"font-size:10pt;font-family:'MS Mincho';\">a</rt></ruby>",
                         "EQ ruby: an unquoted multi-word Font: name runs to the next switch");
        TestReport.Check(Fields.WordFieldRuby.ToHtml(@" \* jc0 \o\ac(\s\up 11(a),b) \* MERGEFORMAT") == "<ruby>b<rt>a</rt></ruby>"
                         && Fields.WordFieldRuby.ToHtml(@" \* jc0 \o(x,\s\up 3(a),b)") is null && Fields.WordFieldRuby.ToHtml(@" \* jc0 \o(\s\up 11(a,b") is null,
                         "EQ ruby: text after \\o's last bracket is not base; \\s\\up must be \\o's first element; a code cut short stays an equation");
    }

    private static void BulletFontNames(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-bullet-font-name-" + format.TrimStart('.')) };
        for (int open = 0; open < 2; open++)     // the first open builds the index, the second reads the list styles from it
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            List<string> blocks = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
            string wingdings = blocks.FirstOrDefault(html => html.Contains("font-name Wingdings")) ?? "";
            string symbol = blocks.FirstOrDefault(html => html.Contains("font-name Symbol")) ?? "";
            string pass = open == 0 ? "first open" : "reopen";
            TestReport.Check(wingdings.Contains("") && wingdings.Contains("font-family:'Wingdings'") && !wingdings.Contains("BulletFace"),
                             $"{format} ({pass}): a bullet in font-face \"BulletFaceW\" draws in its family Wingdings, code kept");
            TestReport.Check(symbol.Contains("•") && !symbol.Contains("") && !symbol.Contains("BulletFace"),
                             $"{format} ({pass}): a bullet in font-face \"BulletFaceS\" (Symbol) recoded to U+2022 as a Symbol bullet");
        }
    }
}
