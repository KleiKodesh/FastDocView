using System.Text;
using FastDocView.DocBinary;
using FastDocView.Rtf;

namespace FastDocView.Tests;

/// <summary>
/// Word fields in .doc and .rtf: test-data/word-fields/fields-v1.doc / .rtf, saved by Word 16 from one document with a
/// field per paragraph (make-fields-fixtures.ps1 next to them; Word's own results printed by the script), plus hand-written
/// RTF for what Word never saves (fields without a result, \fldlock, \fldpriv, a result across paragraphs, counters across
/// blocks). Expected values hand-worked from Word's results; the renderer and both search readers of each format must agree.
/// </summary>
public static class WordFieldsOtherFormatsTest
{
    public static void Run()
    {
        TestReport.Step("Word fields: .doc and .rtf (Word-saved fixtures, hand-written RTF)");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-fields");
        foreach (string extension in new[] { ".doc", ".rtf", ".docx" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "fields-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-fields/fields-v1{extension} present"); continue; }
            WordSaved(path, extension);
        }
        HandWrittenRtf();
        DocListContexts(Path.GetFullPath(Path.Combine(folder, "fields-lists-v1.doc")));
    }

    /// <summary>
    /// .doc REF \n \w \r on a three-level list "%1)" "%2)" "%3)" (fields-lists-v1.doc, make-fields-lists-fixture.ps1): the
    /// fields carry Word's results, which the renderer shows; the same codes are computed here through the document's own
    /// field context (the span scan's list contexts) as a field Word saved without a result or dirty would be. Word 16's
    /// results (COM, 2026-11-04): inside 2)b) \r "i)"; from outside the list \n "i)", \w "2)b)i)", \r "i)".
    /// </summary>
    private static void DocListContexts(string path)
    {
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-fields/fields-lists-v1.doc present"); return; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields-lists-doc") };
        using var document = (DocBinaryWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        List<string> text = document.RenderBlocks(0, document.BlockCount).Select(block => Strip(block.Html).Trim()).ToList();
        int Block(string part) => text.FindIndex(block => block.Contains(part, StringComparison.Ordinal));
        Fields.WordFieldContext context = document.Renderer.FieldContext!;
        string Computed(string code, int block)
        {
            context.CurrentBlock = block;
            return Fields.WordFieldEvaluator.Compute(Fields.WordFieldInstruction.Parse(code), context) is { } output ? output.Text ?? "" : "";
        }
        int outside = Block("refs "), inside = Block("sub b ");
        string fromOutside = string.Join(" | ", Computed("REF deep \\n", outside), Computed("REF deep \\w", outside), Computed("REF deep \\r", outside));
        TestReport.Check(outside >= 0 && fromOutside == "i) | 2)b)i) | i)", $".doc REF \\n / \\w / \\r from outside the list (Word 16: i), 2)b)i), i)): [{fromOutside}]");
        TestReport.Check(inside >= 0 && Computed("REF deep \\r", inside) == "i)", $".doc REF \\r from inside 2)b) (Word 16: i)): [{Computed("REF deep \\r", inside)}]");
        TestReport.Check(Computed("REF deep \\w", Block("one")) == "2)b)i)", ".doc REF \\w from inside 1): the full number");
    }

    private static string Strip(string html) => WordFieldsFidelityTest.Strip(html);

    private static void WordSaved(string path, string format)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields-" + format.TrimStart('.')) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Block(string start) => html.FirstOrDefault(block => Strip(block).TrimStart().StartsWith(start, StringComparison.Ordinal)) ?? "";
        void Expect(string start, Func<string, bool> check, string what)
        {
            string block = Block(start);
            TestReport.Check(block.Length > 0 && check(block), $"{format}: {what} (got [{Strip(block)}])");
        }
        Expect("cached ", b => Strip(b) == "cached Target text" && b.Contains("href=\"#bookmark-Target\""), "REF \\h: the cached result, linked");
        Expect("page ", b => !Strip(b).Contains('1') && b.Contains("field-page-placeholder"), "PAGE: a placeholder, not Word's page number");
        Expect("see ", b => b.Contains("field-page-placeholder") && b.Contains("href=\"#bookmark-Target\""), "PAGEREF \\h: a linked placeholder");
        Expect("sum ", b => Strip(b) == "sum 14", "= 2+3*4: 14");
        Expect("Figure 1", b => Strip(b) == "Figure 1", "SEQ: its result");
        Expect("Figure ii", b => Strip(b) == "Figure ii", "SEQ \\* roman: its result");
        Expect("locked ", b => Strip(b) == "locked 2", "locked formula: its result");
        Expect("author ", b => Strip(b) == "author Moshe Cohen", "AUTHOR");
        Expect("client ", b => Strip(b) == "client Acme Ltd", "DOCPROPERTY Client");
        Expect("box ", b => b.Contains("☒") && !b.Contains("FORMCHECKBOX"), "FORMCHECKBOX checked");
        Expect("drop ", b => Strip(b) == "drop Blueberry", "FORMDROPDOWN: the selected entry");
        Expect("macro ", b => Strip(b) == "macro Press here ", "MACROBUTTON: its text as written, trailing space too (Word 16 shows \"Click here now \")");
        Expect("eq ", b => b.Contains("<mfrac>"), "EQ: MathML");
        Expect("auto ", b => Strip(b) == "auto 1.", "AUTONUM: 1.");
        Expect("list ", b => Strip(b) == "list 1)", "LISTNUM: 1)");
        Expect("index", b => Strip(b) == "index", "XE: hidden");
        Expect("merge ", b => Strip(b) == "merge «FirstName»", "MERGEFIELD");
        Expect("quote ", b => Strip(b) == "quote QUOTED TEXT", "QUOTE \\* Upper");
        Expect("nested ", b => Strip(b) == "nested nestedyes", "IF: its result");
        Expect("link ", b => b.Contains("href=\"#bookmark-Target\"") && Strip(b) == "link Target", "HYPERLINK \\l");
        SearchAgrees(document, format, [("Target text", 2), ("Press here", 1), ("Blueberry", 1), ("QUOTED TEXT", 1), ("nestedyes", 1),
                                        ("hiddenentry", 0), ("MACROBUTTON", 0), ("nestedno", 0), ("FORMCHECKBOX", 0), ("Moshe Cohen", 1), ("auto 1.", 1), ("list 1)", 1)]);
    }

    /// <summary>The one-pass search and the per-block reference read the same blocks, and the expected number of them.</summary>
    private static void SearchAgrees(IWordWindowedDocument document, string format, (string query, int blocks)[] queries)
    {
        foreach ((string query, int expected) in queries)
        {
            int[] onePass = Hits(document, query, reference: false), reference = Hits(document, query, reference: true);
            TestReport.Check(onePass.SequenceEqual(reference) && onePass.Length == expected,
                             $"{format} search \"{query}\": {expected} block(s), one pass = reference ([{string.Join(",", onePass)}] / [{string.Join(",", reference)}])");
        }
    }

    private static int[] Hits(IWordWindowedDocument document, string query, bool reference)
    {
        var hits = new List<int>();
        int from = 0;
        while (true)
        {
            WordDocumentSearchBatch batch = document switch
            {
                RtfWindowedDocument rtf when reference => rtf.SearchByBlockReparse(query, WordDocumentSearchOptions.Default, from, 4),
                DocBinaryWindowedDocument doc when reference => doc.SearchByBlockRanges(query, WordDocumentSearchOptions.Default, from, 4),
                _ => document.Search(query, WordDocumentSearchOptions.Default, from, 4),
            };
            hits.AddRange(batch.HitBlocks);
            if (batch.ReachedEnd) break;
            from = batch.NextBlock;
        }
        return hits.ToArray();
    }

    private static void HandWrittenRtf()
    {
        // fields Word never saves this way: no \fldrslt (computed), \fldlock with a stale result, \fldpriv, a dropdown form field,
        // SEQ counting across many blocks (the counters ride the index's anchors), a hyperlink result across paragraphs
        var rtf = new StringBuilder(@"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0 Arial;}}{\stylesheet{\s0 Normal;}{\s5 Quote Style;}}{\info{\title Hand title}{\author Rivka}{\creatim\yr2024\mo3\dy5\hr10\min21}}{\*\userprops {\propname Client}\proptype30{\staticval Acme RTF}}{\*\docvar {Var1}{RtfVariable}}");
        rtf.Append(@"\pard {\*\bkmkstart Target}Target text{\*\bkmkend Target} not in the bookmark\par ");
        rtf.Append(@"\pard\s5 quoted line\par ");
        rtf.Append(@"\pard sref {\field{\*\fldinst STYLEREF ""Quote Style"" }}\par ");
        rtf.Append(@"\pard calc {\field{\*\fldinst = 6*7 }}\par ");
        rtf.Append(@"\pard locked {\field\fldlock{\*\fldinst = 1+1 }{\fldrslt staleresult}}\par ");
        rtf.Append(@"\pard private {\field\fldpriv{\*\fldinst PRIVATE data }{\fldrslt binarydata}}\par ");
        rtf.Append(@"\pard props {\field{\*\fldinst TITLE }}/{\field{\*\fldinst AUTHOR }}/{\field{\*\fldinst DOCPROPERTY Client }}/{\field{\*\fldinst DOCVARIABLE Var1 }}/{\field{\*\fldinst CREATEDATE \\@ ""yyyy-MM-dd"" }}\par ");
        rtf.Append(@"\pard drop {\field{\*\fldinst FORMDROPDOWN {\*\formfield{\fftype2\ffres1{\*\ffname D1}{\*\ffl Red}{\*\ffl Green}}}}{\fldrslt }}\par ");
        rtf.Append(@"\pard check {\field{\*\fldinst FORMCHECKBOX {\*\formfield{\fftype1\ffres25\ffdefres1}}}{\fldrslt }}\par ");
        rtf.Append(@"\pard ref {\field{\*\fldinst REF Target }}\par ");
        rtf.Append(@"\pard nested {\field{\*\fldinst IF {\field{\*\fldinst = 2+2 }{\fldrslt 4}} = 4 ""yesnested"" ""nonested"" }}\par ");
        rtf.Append(@"\pard page {\field{\*\fldinst PAGE }{\fldrslt 57}}\par ");
        rtf.Append(@"\pard {\field{\*\fldinst HYPERLINK \\l ""Target"" }{\fldrslt first part\par second part}} after\par ");
        // enough text between the counted fields that the later ones fall past an index anchor (RtfRandomAccessIndex.AnchorSpanBytes)
        string filler = new string('x', 900);
        for (int i = 1; i <= 120; i++) rtf.Append(@"\pard Figure {\field{\*\fldinst SEQ Figure }" + (i % 2 == 0 ? "" : @"{\fldrslt " + i + "}") + "} " + filler + @"\par ");
        rtf.Append(@"\pard last line\par }");
        string path = Path.Combine(TestReport.OutputDirectory, "fields-handwritten-v1.rtf");
        File.WriteAllText(path, rtf.ToString(), new UTF8Encoding(false));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields-rtf-hand") };
        string sidecar = policy.IndexPathFor(path);
        if (File.Exists(sidecar)) File.Delete(sidecar);
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        string Block(string start) => html.FirstOrDefault(block => Strip(block).TrimStart().StartsWith(start, StringComparison.Ordinal)) ?? "";
        void Expect(string start, Func<string, bool> check, string what)
        {
            string block = Block(start);
            TestReport.Check(block.Length > 0 && check(block), $".rtf (hand-written): {what} (got [{Strip(block)}])");
        }
        Expect("calc ", b => Strip(b) == "calc 42", "a field without \\fldrslt is computed");
        Expect("sref ", b => Strip(b) == "sref quoted line", "STYLEREF of any style: the nearest paragraph in \\s5 (by its stylesheet name)");
        Expect("locked ", b => Strip(b) == "locked staleresult", "\\fldlock keeps its result");
        Expect("private ", b => Strip(b).Trim() == "private", "\\fldpriv: the result is not shown");
        Expect("props ", b => Strip(b) == "props Hand title/Rivka/Acme RTF/RtfVariable/2024-03-05", "{\\info}, \\userprops and \\docvar properties");
        Expect("drop ", b => Strip(b) == "drop Green", "\\formfield drop-down: \\ffres picks the entry");
        Expect("check ", b => b.Contains("☒"), "\\formfield checkbox: \\ffres 25 = the default (\\ffdefres1)");
        Expect("ref ", b => Strip(b) == "ref Target text", "REF without a result: the bookmarked text");
        Expect("nested ", b => Strip(b) == "nested yesnested", "a nested field's result is part of the code");
        Expect("page ", b => !Strip(b).Contains("57") && b.Contains("field-page-placeholder"), "PAGE: a placeholder");
        Expect("second part", b => b.Contains("href=\"#bookmark-Target\"") && b.Contains("second part</a>") && !b.Contains("after</a>"), "a hyperlink result across paragraphs stays a link in the next block");
        for (int i = 1; i <= 120; i += 17)
        {
            string expected = "Figure " + i;
            TestReport.Check(html.Any(block => Strip(block).StartsWith(expected + " x", StringComparison.Ordinal)), $".rtf (hand-written): SEQ #{i} reads {expected} (results and computed alike, across index anchors)");
        }
        SearchAgrees(document, ".rtf (hand-written)", [("calc 42", 1), ("staleresult", 1), ("binarydata", 0), ("Acme RTF", 1), ("Green", 1), ("yesnested", 1), ("nonested", 0),
                                                       ("page 57", 0), ("Figure 120 ", 1), ("Figure 64 ", 1), ("PRIVATE", 0)]);
    }
}
