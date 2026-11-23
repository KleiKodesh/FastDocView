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
        Edges(root);
        NoteInField(root);
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
        // (file, case, guide as Word's PDF draws it, base line as Word's PDF draws it, guide font). In the base line, the part
        // before "|" is the guide's rest after a comma, drawn in the guide's font and size (Word's PDF: in the guide font).
        string b = "\\";
        (string File, int Case, string Guide, string Base, string Font)[] cases =
        [
            ("ruby-escapes-v1.doc", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-v1.doc", 1, "g", "1|,base,1", "Courier New"),
            ("ruby-escapes-v1.doc", 2, "g;2", "base;2", "Arial"), ("ruby-escapes-v1.doc", 3, "g" + b + "3", "base" + b + "3", "Arial"),
            ("ruby-escapes-v1.doc", 4, "g4(", "base4(", "Calibri"), ("ruby-escapes-v1.doc", 5, "g 5 x", "base 5", "Calibri"),
            ("ruby-escapes-v1.rtf", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-v1.rtf", 1, "g", "1|,base,1", "Courier New"),
            ("ruby-escapes-v1.rtf", 2, "g;2", "base;2", "Arial"), ("ruby-escapes-v1.rtf", 3, "g" + b + "3", "base" + b + "3", "Arial"),
            ("ruby-escapes-v1.rtf", 4, "g4(", "base4(", "Calibri"), ("ruby-escapes-v1.rtf", 5, "g 5 x", "base 5", "Calibri"),
            ("ruby-escapes-probe-v1.rtf", 0, "g(0", "base(0", "Times New Roman"), ("ruby-escapes-probe-v1.rtf", 1, "g", "1|,base,1", "Courier New"),
            ("ruby-escapes-probe-v1.rtf", 2, "a" + b, "b|,c" + b + ",d", "Arial"), ("ruby-escapes-probe-v1.rtf", 3, "e" + b + "f", "g" + b + "h", "Arial"),
            ("ruby-escapes-probe-v1.rtf", 4, "i" + b + "(j", "k" + b + b + "l", "Calibri"),
            ("ruby-docx-commas-v1.doc", 0, "g", "1|,base", "Courier New"), ("ruby-docx-commas-v1.doc", 1, "g", "b,1", "Courier New"),
            ("ruby-docx-commas-v1.doc", 2, "g", "1,2|,base", "Courier New"), ("ruby-docx-commas-v1.doc", 3, "g", "1|,b,2", "Courier New"),
            ("ruby-docx-commas-v1.doc", 4, "abcd", "ef", "Courier New"), ("ruby-docx-commas-v1.doc", 5, "ab", "cdef", "Courier New"),
            ("ruby-docx-commas-v1.doc", 6, "a(b", "c(d", "Courier New"), ("ruby-docx-commas-v1.doc", 7, "g", "1|,base,1", "Courier New"),
        ];
        foreach (var group in cases.GroupBy(c => c.File))
        {
            List<string> blocks = RenderAll(root, group.Key);
            foreach (var c in group)
            {
                string block = blocks.FirstOrDefault(html => html.Contains("ruby escape " + c.Case + ":") || html.Contains("ruby comma " + c.Case + ":")) ?? "";
                string font = "font-size:10pt;font-family:'" + c.Font + "';";
                string[] baseParts = c.Base.Split('|');
                string baseHtml = baseParts.Length == 2 ? "<span style=\"" + font + "\">" + Html.WordHtmlText.Encode(baseParts[0]) + "</span>" + Html.WordHtmlText.Encode(baseParts[1])
                                                        : Html.WordHtmlText.Encode(c.Base);
                string expected = "<ruby>" + baseHtml + "<rt style=\"" + font + "\">" + Html.WordHtmlText.Encode(c.Guide) + "</rt></ruby>";
                TestReport.Check(block.Contains(expected), $"{group.Key}: ruby case {c.Case}: guide and base as Word's PDF draws them, {c.Font}");
            }
        }
        // .docx: Word prints its own w:ruby through the same rules, a little differently (Word 16 PDFs ruby-escapes-v1.docx.pdf
        // in the generator's run / ruby-docx-commas-v1.docx.pdf): each part ends at its first ")", and a guide comma puts
        // "rest,rest" on the base line in place of the base (the rests in the guide's font)
        (string File, int Case, string Guide, string Base)[] docxCases =
        [
            ("ruby-escapes-v1.docx", 0, "g(0", "base(0"), ("ruby-escapes-v1.docx", 1, "g", "1,1"), ("ruby-escapes-v1.docx", 2, "g;2", "base;2"),
            ("ruby-escapes-v1.docx", 3, "g" + b + "3", "base" + b + "3"), ("ruby-escapes-v1.docx", 4, "g", "base"), ("ruby-escapes-v1.docx", 5, "g 5 x", "base 5"),
            ("ruby-docx-commas-v1.docx", 0, "g", "1,1"), ("ruby-docx-commas-v1.docx", 1, "g", "b,1"), ("ruby-docx-commas-v1.docx", 2, "g", "1,2,1,2"),
            ("ruby-docx-commas-v1.docx", 3, "g", "1,1"), ("ruby-docx-commas-v1.docx", 4, "ab", "ef"), ("ruby-docx-commas-v1.docx", 5, "ab", "cd"),
            ("ruby-docx-commas-v1.docx", 6, "a(b", "c(d"), ("ruby-docx-commas-v1.docx", 7, "g", "1,1"),
        ];
        foreach (var group in docxCases.GroupBy(c => c.File))
        {
            List<string> blocks = RenderAll(root, group.Key);
            foreach (var c in group)
            {
                string block = blocks.FirstOrDefault(html => html.Contains("ruby escape " + c.Case + ":") || html.Contains("ruby comma " + c.Case + ":")) ?? "";
                int start = block.IndexOf("<ruby>", StringComparison.Ordinal), rt = block.IndexOf("<rt", start + 1, StringComparison.Ordinal),
                    end = block.IndexOf("</ruby>", start + 1, StringComparison.Ordinal);
                string baseText = start < 0 || rt < 0 ? "" : PlainText(block.Substring(start + 6, rt - start - 6));
                string guideText = rt < 0 || end < 0 ? "" : PlainText(block.Substring(rt, end - rt));
                // a guide split by a comma: its rests on the base line are in the guide's font (Courier New in both files)
                bool restInGuideFont = !c.Base.StartsWith("1", StringComparison.Ordinal) || start >= 0 && rt > start && block.Substring(start, rt - start).Contains("Courier New");
                TestReport.Check(guideText == c.Guide && baseText == c.Base && restInGuideFont,
                                 $"{group.Key}: ruby case {c.Case}: guide {c.Guide} over {c.Base} as Word's PDF draws them (got {guideText} over {baseText})");
            }
        }
        TestReport.Check(Fields.WordFieldRuby.ToHtml(@" \* jc0 \* Font:MS Mincho \* hps20 \o(\s\up 11(a),b)") == "<ruby>b<rt style=\"font-size:10pt;font-family:'MS Mincho';\">a</rt></ruby>",
                         "EQ ruby: an unquoted multi-word Font: name runs to the next switch");
        TestReport.Check(Fields.WordFieldRuby.ToHtml(@" \* jc0 \o\ac(\s\up 11(a),b) \* MERGEFORMAT") == "<ruby>b<rt>a</rt></ruby>"
                         && Fields.WordFieldRuby.ToHtml(@" \* jc0 \o(x,\s\up 3(a),b)") is null && Fields.WordFieldRuby.ToHtml(@" \* jc0 \o(\s\up 11(a,b") is null,
                         "EQ ruby: text after \\o's last bracket is not base; \\s\\up must be \\o's first element; a code cut short stays an equation");
    }

    /// <summary>
    /// test-data/word-ruby/ruby-edges-v1 (.docx hand-written by make-ruby-edges.py, .doc / .rtf saved from it by Word, a PDF of
    /// each) and ruby-edges-semicolon-v1.rtf: a tab, a note mark and a bookmark after a ")" in the base; a note mark and a
    /// bookmark in a base whose guide has a comma; a ruby with a ")" ending its paragraph; ";" in guide and base. Word's PDFs
    /// (the same for all three formats): "abc d", "abc" + mark 1, "abcy", "g" over "3,base" + mark 2, "g" over "4,basew",
    /// "end of linex", "g;6" over "b;6" (the semicolon probe: over ";b;6"). Search and GetText read what is drawn.
    /// </summary>
    private static void Edges(string root)
    {
        (int Case, string Guide, string Base)[] expected =
        [
            (0, "g0", "abc d"), (1, "g1", "abc1"), (2, "g2", "abcy"), (3, "g", "3,base2"), (4, "g", "4,basew"), (5, "g5", "end of linex"), (6, "g;6", "b;6"),
        ];
        foreach (string file in new[] { "ruby-edges-v1.docx", "ruby-edges-v1.doc", "ruby-edges-v1.rtf", "ruby-edges-semicolon-v1.rtf" })
        {
            List<string> blocks = RenderAll(root, file);
            foreach (var c in expected)
            {
                string block = blocks.FirstOrDefault(html => html.Contains("ruby edge " + c.Case + ":")) ?? "";
                int start = block.IndexOf("<ruby>", StringComparison.Ordinal), rt = start < 0 ? -1 : block.IndexOf("<rt", start, StringComparison.Ordinal),
                    end = start < 0 ? -1 : block.IndexOf("</ruby>", start, StringComparison.Ordinal);
                string baseHtml = start < 0 || rt < 0 ? "" : block.Substring(start + 6, rt - start - 6);
                string baseText = System.Text.RegularExpressions.Regex.Replace(PlainText(baseHtml), @"\s+", " ").Trim();
                string guideText = rt < 0 || end < 0 ? "" : PlainText(block.Substring(rt, end - rt));
                string wantBase = file.Contains("semicolon") && c.Case == 6 ? ";b;6" : c.Base;
                bool extra = c.Case switch
                {
                    0 => baseHtml.Contains('\t'),                                                    // the tab is a tab in every format
                    1 or 3 => baseHtml.Contains("note-reference", StringComparison.Ordinal),        // the note's mark, numbered as Word does
                    2 => block.Contains("EdgeTwo", StringComparison.Ordinal) || file != "ruby-edges-v1.docx",
                    4 => block.Contains("EdgeFour", StringComparison.Ordinal) || file != "ruby-edges-v1.docx",
                    5 => !System.Text.RegularExpressions.Regex.IsMatch(baseHtml, @"linex\s+<"),      // trailing spaces at the paragraph end are not drawn
                    _ => true,
                };
                TestReport.Check(guideText == c.Guide && baseText == wantBase && extra,
                                 $"{file}: ruby edge {c.Case}: \"{c.Guide}\" over \"{wantBase}\" as Word's PDF draws it (got \"{guideText}\" over \"{baseText}\")"
                                 + (Environment.GetEnvironmentVariable("FASTDOCVIEW_RUBY_DUMP") is null ? "" : " " + block));
            }
        }
        // the notes inside rubies keep Word's numbering and come with their blocks, in order (the pane)
        foreach (string file in new[] { "ruby-edges-v1.docx", "ruby-edges-v1.doc", "ruby-edges-v1.rtf" })
        {
            var notePolicy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-ruby-notes-" + file.Replace('.', '-')) };
            using IWordWindowedDocument noted = WordWindowedDocumentOpener.Open(Path.Combine(root, "word-ruby", file), notePolicy);
            List<WordDocumentNoteHtml> notes = noted.RenderBlocks(0, noted.BlockCount).SelectMany(block => block.Notes).ToList();
            TestReport.Check(notes.Count == 2 && notes[0].Number == 1 && notes[0].Label == "1" && PlainText(notes[0].Html).Contains("note one")
                             && notes[1].Number == 2 && notes[1].Label == "2" && PlainText(notes[1].Html).Contains("note two"),
                             $"{file}: the notes inside rubies numbered 1 and 2 in order, each with its text ({notes.Count} note(s))");
        }
        // search and GetText read the .docx base as drawn
        string docx = Path.Combine(root, "word-ruby", "ruby-edges-v1.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-ruby-edges-search") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(docx, policy);
        foreach (string query in new[] { "abcy", "3,base", "4,basew", "end of linex" })
        {
            WordDocumentSearchBatch hits = document.Search(query, WordDocumentSearchOptions.Default, 0, document.BlockCount);
            TestReport.Check(hits.HitBlocks.Length == 1, $"ruby-edges-v1.docx: search finds \"{query}\" as drawn ({hits.HitBlocks.Length} block(s))");
        }
        TestReport.Check(document.Search("ab)c", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 0,
                         "ruby-edges-v1.docx: search does not find a \")\" Word does not draw");
        string text = WordDocumentText.GetText(document);
        TestReport.Check(text.Contains("3,base") && text.Contains("abcy") && !text.Contains("ab)c"), "ruby-edges-v1.docx: GetText reads the bases as drawn");
    }

    /// <summary>
    /// test-data/word-ruby/note-in-field-v1.rtf (hand-written): a note inside a plain EQ field's code (a fraction, not a
    /// ruby), inside a field's result, inside a non-EQ field's code, and one in body text for comparison. LibreOffice
    /// resolves {ootnote} wherever it stands, field code included (rtfdispatchdestination.cxx:265-317). Each note keeps
    /// its mark and its number: in the EQ numerator, between the result's runs, after a non-EQ field's result.
    /// </summary>
    private static void NoteInField(string root)
    {
        const string file = "note-in-field-v1.rtf";
        List<string> blocks = RenderAll(root, file);
        string Block(string label) => blocks.FirstOrDefault(html => html.Contains(label, StringComparison.Ordinal)) ?? "";
        static string Mark(int n) => "data-note-number=\"" + n + "\">" + n + "</sup>";
        string a = Block("case A:"), b = Block("case B:"), c = Block("case C:"), e = Block("case E:");
        int numerator = a.IndexOf("<mfrac><mrow>", StringComparison.Ordinal), denominator = a.IndexOf("</mrow><mrow>", StringComparison.Ordinal);
        int markA = a.IndexOf(Mark(1), StringComparison.Ordinal);
        TestReport.Check(numerator >= 0 && markA > numerator && markA < denominator, $"{file}: a note in a plain EQ field's code keeps its mark in place (the fraction's numerator)");
        TestReport.Check(b.Contains("res<sup", StringComparison.Ordinal) && b.Contains(Mark(2) + "ult", StringComparison.Ordinal), $"{file}: a note in a field's result is drawn between the result's runs");
        TestReport.Check(c.Contains("merged<sup", StringComparison.Ordinal) && c.Contains(Mark(3), StringComparison.Ordinal) && !c.Contains("name", StringComparison.Ordinal),
                         $"{file}: a note in a non-EQ field's code is drawn after the field's result; the code stays hidden");
        TestReport.Check(e.Contains(Mark(4), StringComparison.Ordinal), $"{file}: the body note after them is numbered 4");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-note-in-field") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(Path.Combine(root, "word-ruby", file), policy);
        List<WordDocumentNoteHtml> notes = document.RenderBlocks(0, document.BlockCount).SelectMany(block => block.Notes).ToList();
        TestReport.Check(notes.Count == 4 && Enumerable.Range(0, 4).All(i => notes[i].Number == i + 1 && PlainText(notes[i].Html).Contains("note " + "ABCE"[i])),
                         $"{file}: the notes pane lists the four notes in order, each with its text ({notes.Count} note(s))");
        TestReport.Check(document.SearchNotes("note A", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.Length == 1
                         && document.SearchNotes("note C", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.Length == 1
                         && document.Search("note A", WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 0,
                         $"{file}: notes inside field codes are found by the notes search, not in the body");
    }

    private static List<string> RenderAll(string root, string file)
    {
        string path = Path.Combine(root, "word-ruby", file);
        if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-ruby/{file} present"); return []; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-ruby-" + file.Replace('.', '-')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    private static string PlainText(string html) => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]*>", ""));

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
