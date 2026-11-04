using System.Globalization;
using System.Text;
using FastDocView.Docx;
using FastDocView.Fields;

namespace FastDocView.Tests;

/// <summary>
/// Word fields (Fields/*): switches and formulas against Word 16's own results (COM automation on the dev machine,
/// 2026-11-04: Fields.Add + Fields.Update, Range.Text of each result - the values below are what Word wrote), then the .docx
/// renderer and both .docx search readers on hand-written documents. .doc / .rtf: <see cref="WordFieldsOtherFormatsTest"/>.
/// </summary>
public static class WordFieldsFidelityTest
{
    public static void Run()
    {
        TestReport.Step("Word fields: switches, formulas, docx rendering and search");
        Switches();
        Formulas();
        Instructions();
        DocxRendering();
        DocxLookups();
        DocxListContexts();
        DocxListLookupsAcrossSpans();
        DocxTableOfContentsFarHeading();
    }

    private static readonly WordFieldContext English = new() { Culture = CultureInfo.GetCultureInfo("en-US"), Now = new DateTime(2024, 3, 5, 14, 7, 9) };

    private static string Field(string code, WordFieldContext? context = null, int? counted = null)
        => WordFieldEvaluator.Compute(WordFieldInstruction.Parse(code), context ?? English, counted) is { } output ? output.Text ?? output.Html ?? "" : "";

    private static void Check(string code, string expected)
    {
        string actual = Field(code);
        TestReport.Check(actual == expected, $"{code} → [{expected}] (got [{actual}])");
    }

    private static void Switches()
    {
        // \* number and text formats (Word 16)
        Check("= 21 \\* CardText", "twenty-one");
        Check("= 21 \\* OrdText", "twenty-first");
        Check("= 21 \\* Ordinal", "21st");
        Check("= 5 \\* Ordinal", "5th");
        Check("= 12 \\* Ordinal", "12th");
        Check("= 23 \\* Ordinal", "23rd");
        Check("= 101 \\* OrdText", "one hundred first");
        Check("= 112 \\* OrdText", "one hundred twelfth");
        Check("= 1234 \\* CardText", "one thousand two hundred thirty-four");
        Check("= 0 \\* CardText", "zero");
        Check("= 1000000 \\* CardText", WordFieldFormat.NumberNotRepresentable);
        Check("= 1234.5 \\* DollarText", "one thousand two hundred thirty-four and 50/100");   // Word's "and" is localized: Hebrew UI wrote "ו-"
        Check("= 0.25 \\* DollarText", "zero and 25/100");
        Check("= 255 \\* Hex", "FF");
        Check("= 28 \\* ALPHABETIC", "BB");
        Check("= 1 \\* alphabetic", "a");
        Check("= 52 \\* alphabetic", "zz");
        Check("= 53 \\* alphabetic", "aaa");
        Check("= 1999 \\* roman", "mcmxcix");
        Check("= 4000 \\* roman", "mmmm");
        Check("= 2.5 \\* roman", "iii");
        Check("= 1.5 \\* Arabic", "2");
        Check("= 21 \\* ArabicDash", "- 21 -");
        Check("= 15 \\* hebrew1", "טו");
        Check("= 115 \\* hebrew1", "קטו");
        Check("= 270 \\* hebrew1", "רע");
        Check("= 272 \\* hebrew1", "רעב");
        Check("= 392 \\* hebrew1", "שצב");
        Check("= 393 \\* hebrew1", WordFieldFormat.NumberNotRepresentable);
        Check("= 3 \\* hebrew2", "ג");
        Check("= 22 \\* hebrew2", "ת");
        Check("= 23 \\* hebrew2", "תא");
        Check("= 44 \\* hebrew2", "תת");
        Check("= 45 \\* hebrew2", "תתא");
        Check("= 21 \\* CardText \\* Upper", "TWENTY-ONE");
        Check("= 21 \\* CardText \\* FirstCap", "Twenty-one");
        Check("= 21 \\* CardText \\* Caps", "Twenty-One");
        Check("QUOTE \"abc def\" \\* FirstCap", "Abc def");
        Check("QUOTE \"abc def\" \\* Caps", "Abc Def");
        Check("QUOTE \"ABC\" \\* Lower", "abc");
        Check("= 2 \\# \"00\" \\* roman", "ii");                         // a \* number format wins over \#
        // \# numeric pictures (Word 16)
        Check("= 1234.567 \\# \"#,##0.00\"", "1,234.57");
        Check("= -5 \\# \"0.0;(0.0)\"", "(5.0)");
        Check("= 0.256 \\# \"0%\"", "0%");
        Check("= 3 \\# \"00\"", "03");
        Check("= 1234.5 \\# \"$#,##0.00\"", "$1,234.50");
        Check("= 2.5 \\# \"0\"", "3");
        Check("= -2.5 \\# \"0\"", "-3");
        Check("= 12 \\# \"x##\"", " 12");
        Check("= 1234 \\# \"#,##0\"", "1,234");
        Check("= 0.5 \\# \"#.##\"", " .5 ");
        Check("= 1.005 \\# \"0.00\"", "1.01");
        Check("= 2.675 \\# \"0.00\"", "2.68");
        Check("= 1234.5 \\# \"#,##0.00;-#,##0.00;'zero'\"", "1,234.50");
        Check("= 0 \\# \"#,##0.00;-#,##0.00;'zero'\"", "zero");
        Check("= 5 \\# \"'Total: '0\"", "Total: 5");
        Check("= 10.5 \\# 0", "11");
        Check("= 1234.567 \\# \"0.0\"", "1234.6");
        Check("= 5 \\# \"+0;-0\"", "+5");
        Check("= -5 \\# \"+0;-0\"", "-5");
        Check("= 5 \\# \"##%\"", " 5%");
        Check("= 5 \\# \"0.00 'USD'\"", "5.00 USD");
        Check("= -1234.5 \\# \"#,##0.00\"", "-1,234.50");
        Check("= -0.4 \\# \"0\"", "0");
        Check("= 1234.5678 \\# \",0.0\"", "1,234.6");
        Check("= 12345 \\# \"00,000\"", "12,345");
        // \@ date pictures (en-US names; Word takes names from the system: a Hebrew system wrote Hebrew names)
        Check("QUOTE \"2024-03-05 14:07\" \\@ \"dddd, MMMM d, yyyy h:mm am/pm\"", "Tuesday, March 5, 2024 2:07 pm");
        Check("QUOTE \"2024-03-05\" \\@ \"d/M/yyyy\"", "5/3/2024");
        Check("QUOTE \"2024-03-05 14:07:09\" \\@ \"HH:mm:ss\"", "14:07:09");
        Check("QUOTE \"2024-03-05\" \\@ \"dd.MM.yy 'week'\"", "05.03.24 week");
        Check("QUOTE \"2024-03-05\" \\@ \"ddd MMM yy\"", "Tue Mar 24");
        Check("DATE \\@ \"yyyy-MM-dd\"", "2024-03-05");
        Check("TIME \\@ \"H:mm\"", "14:07");
        var hebrew = new WordFieldContext { Culture = CultureInfo.GetCultureInfo("he-IL") };
        string hebrewDate = Field("QUOTE \"2024-03-05\" \\@ \"dddd d MMMM yyyy\"", hebrew);
        TestReport.Check(hebrewDate == "יום שלישי 5 מרץ 2024", $"date names follow the system culture (Word 16 on a Hebrew system: יום שלישי 5 מרץ 2024; got {hebrewDate})");
    }

    private static void Formulas()
    {
        Check("= 1/7", "0.14");
        Check("= 2/3*1000", "666.67");
        Check("= 1/8", "0.13");
        Check("= 1234567", "1234567");
        Check("= 0.001", "0.001");
        Check("= 0.001*3", "0.003");
        Check("= 1.2345/3", "0.4115");
        Check("= 10/3+0.001", "3.334");
        Check("= 2.50", "2.5");
        Check("= 3.0", "3");
        Check("= 0.1+0.2", "0.3");
        Check("= 1,000 + 1", "1,001");
        Check("= 2^10", "1024");
        Check("= -2^2", "4");
        Check("= 2+3*4", "14");
        Check("= (2+3)*4", "20");
        Check("= 5%", "0.05");
        Check("= 10 > 5", "1");
        Check("= 3 = 3", "1");
        Check("= 3 <> 3", "0");
        Check("= TRUE", "1");
        Check("= MOD(7,3)", "1");
        Check("= ROUND(2.567,2)", "2.57");
        Check("= IF(1>2,5,6)", "6");
        Check("= AVERAGE(1,2,3,4)", "2.5");
        Check("= PRODUCT(2,3,4)", "24");
        Check("= MAX(1,5,3)", "5");
        Check("= COUNT(1,2,3)", "3");
        Check("= ABS(-4)", "4");
        Check("= AND(1,0)", "0");
        Check("= OR(1,0)", "1");
        Check("= NOT(0)", "1");
        Check("= SIGN(-3)", "-1");
        Check("= INT(-2.5)", "-2");
        Check("= SUM(1.5,2.25)", "3.75");
        Check("= DEFINED(xyz)", "0");
        Check("= 1/0", WordFieldExpression.ZeroDivide);
        Check("= 3 +", WordFieldExpression.UnexpectedEnd);
        Check("= xyz + 1", WordFieldExpression.UndefinedBookmark + "XYZ");
        Check("= 7 MOD 3", WordFieldExpression.SyntaxError + "MOD");
        Check("= 1e3", WordFieldExpression.SyntaxError + "1E3");
        Check("= SUM(1;2)", WordFieldExpression.SyntaxError + ";");
        // IF / COMPARE / QUOTE / SYMBOL / MERGEFIELD / buttons / counters (Word 16)
        Check("IF 3 > 2 \"yes\" \"no\"", "yes");
        Check("IF \"abc\" = \"ABC\" \"yes\" \"no\"", "no");
        Check("IF \"abc\" <> \"ab?\" \"yes\" \"no\"", "no");
        Check("IF 1 = 1 \"a b\" \"c\"", "a b");
        Check("IF 1 = 2 \"a\" ", "");
        Check("COMPARE 3 > 2", "1");
        Check("COMPARE \"abc\" = \"a*\"", "1");
        Check("COMPARE \"abc\" = \"a?c\"", "1");
        Check("COMPARE \"a\" = \"b\"", "0");
        Check("QUOTE \"a\" \"b\"", "ab");
        Check("QUOTE 65", "A");
        Check("MERGEFIELD Name", "«Name»");
        Check("MACROBUTTON NoMacro Click here now", "Click here now");
        Check("GOTOBUTTON bm Go there", "Go");
        Check("TITLE \"x\"", "x");
        var counters = new WordFieldSequenceCounters();
        string Count(string code) => Field(code, English, WordFieldEvaluator.Count(WordFieldInstruction.Parse(code), counters));
        string sequence = string.Join("|", new[] { "SEQ Fig", "SEQ Fig", "SEQ Fig \\c", "SEQ Fig \\h", "SEQ Fig", "SEQ Fig \\r 10", "SEQ Fig \\* roman" }.Select(Count));
        TestReport.Check(sequence == "1|2|2||4|10|xi", $"SEQ counts as Word 16 (1|2|2||4|10|xi; got {sequence})");
        counters = new WordFieldSequenceCounters();
        string autonumbers = string.Join("|", new[] { "AUTONUM", "AUTONUM", "AUTONUM \\s :", "AUTONUMLGL", "AUTONUMOUT" }.Select(Count));
        TestReport.Check(autonumbers == "1.|2.|3:|4.|5.", $"AUTONUM* count as Word 16 (1.|2.|3:|4.|5.; got {autonumbers})");
        counters = new WordFieldSequenceCounters();
        string listNumbers = string.Join("|", new[] { "LISTNUM", "LISTNUM", "LISTNUM \\l 2", "LISTNUM NumberDefault \\l 1 \\s 5", "LISTNUM LegalDefault", "LISTNUM OutlineDefault" }.Select(Count));
        TestReport.Check(listNumbers == "1)|2)|a)|5)|6.|VII.", $"LISTNUM counts as Word 16 (1)|2)|a)|5)|6.|VII.; got {listNumbers})");
        string symbol = WordFieldEvaluator.Compute(WordFieldInstruction.Parse("SYMBOL 65 \\f Arial \\s 12"), English).Html ?? "";
        TestReport.Check(symbol.Contains("font-family:'Arial'") && symbol.Contains("font-size:12pt") && symbol.Contains(">A<"), "SYMBOL: the character in its font and size");
        string wingding = WordFieldEvaluator.Compute(WordFieldInstruction.Parse("SYMBOL 0x4A \\f Wingdings"), English).Html ?? "";
        TestReport.Check(wingding.Contains("\uF04A"), "SYMBOL in a symbol font maps to U+F0xx like w:sym");
        string equation = WordFieldEvaluator.Compute(WordFieldInstruction.Parse("EQ \\f(1,2) + \\r(3,x) \\b\\bc\\[(y) \\s\\up8(2)"), English).Html ?? "";
        TestReport.Check(equation.Contains("<mfrac><mrow><mtext>1</mtext></mrow><mrow><mtext>2</mtext></mrow></mfrac>") && equation.Contains("<mroot><mrow><mtext>x</mtext></mrow><mrow><mtext>3</mtext></mrow></mroot>")
                         && equation.Contains(">[</mo>") && equation.Contains(">]</mo>") && equation.Contains("voffset=\"8pt\""), "EQ: fraction, radical with index, mirrored brackets, raised text as MathML");
        string integral = WordFieldEvaluator.Compute(WordFieldInstruction.Parse("EQ \\i\\su(i=1,n,i)"), English).Html ?? "";
        TestReport.Check(integral.Contains("<munderover>") && integral.Contains("∑"), "EQ \\i\\su: a sum with limits");
        TestReport.Check(WordFieldEvaluator.Compute(WordFieldInstruction.Parse("PAGE"), English).Html!.Contains("field-page-placeholder"), "PAGE: a placeholder");
        TestReport.Check(WordFieldEvaluator.Compute(WordFieldInstruction.Parse("XE \"entry\""), English).IsEmpty, "XE: nothing shows");
        var properties = new WordFieldContext
        {
            DocumentProperty = name => name switch { "AUTHOR" => "Moshe", "NUMWORDS" => "65", "Client" => "Acme", _ => null },
            DocumentDate = name => name == "CREATEDATE" ? new DateTime(2024, 3, 5, 10, 21, 0) : null,
            Culture = CultureInfo.GetCultureInfo("en-US"), FileName = "a.docx", FilePath = "C:\\x\\a.docx",
            BookmarkText = name => name == "Amount" ? "1,250.50 NIS" : name == "Chapter" ? "The chapter" : null,
            Table = new Table([["1", "2", ""], ["3", "4", ""]], 0, 2),
        };
        TestReport.Check(Field("AUTHOR", properties) == "Moshe" && Field("NUMWORDS \\* CardText", properties) == "sixty-five" && Field("DOCPROPERTY Client", properties) == "Acme"
                         && Field("CREATEDATE \\@ \"d MMM yyyy\"", properties) == "5 Mar 2024" && Field("FILENAME", properties) == "a.docx" && Field("FILENAME \\p", properties) == "C:\\x\\a.docx"
                         && Field("INFO Author", properties) == "Moshe", "document property fields (AUTHOR, NUMWORDS, DOCPROPERTY, CREATEDATE, FILENAME, INFO)");
        TestReport.Check(Field("REF Chapter \\* Upper", properties) == "THE CHAPTER" && Field("REF Missing", properties) == WordFieldEvaluator.ReferenceNotFound
                         && Field("= Amount * 2", properties) == "2501", "REF text and a bookmark in a formula (its leading number)");
        TestReport.Check(Field("= SUM(LEFT)", properties) == "3" && Field("= SUM(A1:B2)", properties) == "10" && Field("= A2*B2", properties) == "12" && Field("= SUM(R1C1:R2C1)", properties) == "4",
                         "table formulas: LEFT, A1:B2, cell product, R1C1 ranges");
    }

    private sealed class Table(string[][] cells, int row, int column) : IWordFieldTable
    {
        public int RowCount => cells.Length;
        public int ColumnCount(int r) => cells[r].Length;
        public string? CellText(int r, int c) => r < cells.Length && c < cells[r].Length ? cells[r][c] : null;
        public int CurrentRow => row;
        public int CurrentColumn => column;
    }

    private static void Instructions()
    {
        var parsed = WordFieldInstruction.Parse(" HYPERLINK \"http://a/b c\" \\l \"Part 2\" \\o \"tip \\\"q\\\"\" \\* MERGEFORMAT ");
        TestReport.Check(parsed.Name == "HYPERLINK" && parsed.Arguments.Single() == "http://a/b c" && parsed.SwitchArgument("\\l") == "Part 2" && parsed.SwitchArgument("\\o") == "tip \"q\""
                         && parsed.GeneralFormats.Single() == "MERGEFORMAT", "instruction tokens: quoted argument, switch arguments, escaped quote, general switch (DomainMapper_Impl.cxx lcl_ExtractToken)");
        var seq = WordFieldInstruction.Parse("SEQ Figure \\r 3 \\c \\* ARABIC");
        TestReport.Check(seq.Arguments.Single() == "Figure" && seq.SwitchArgument("\\r") == "3" && seq.Has("\\c") && seq.SwitchArgument("\\c") == "", "SEQ: \\r takes an argument, \\c does not");
        TestReport.Check(WordFieldInstruction.Parse("=2+3 \\# \"0.0\"").FormulaExpression.Trim() == "2+3" && WordFieldInstruction.Parse("=2+3 \\# \"0.0\"").NumericPicture == "0.0",
                         "formula: the expression stops at its first general switch");
    }

    // ---------------------------------------------------------------- .docx

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string Begin(string attributes = "", string inner = "") => $"<w:r><w:fldChar w:fldCharType=\"begin\"{attributes}>{inner}</w:fldChar></w:r>";
    private static string Code(string code) => $"<w:r><w:instrText xml:space=\"preserve\">{code}</w:instrText></w:r>";
    private const string Separate = "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>";
    private const string End = "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";
    private static string Complex(string code, string? result, string beginAttributes = "") => Begin(beginAttributes) + Code(code) + (result is null ? "" : Separate + Run(result)) + End;

    private static void DocxRendering()
    {
        string table = "<w:tbl><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
                     + "<w:tr><w:tc><w:p>" + Run("1") + "</w:p></w:tc><w:tc><w:p>" + Run("2") + "</w:p></w:tc><w:tc><w:p>" + Complex(" =SUM(LEFT) ", "3") + "</w:p></w:tc></w:tr>"
                     + "<w:tr><w:tc><w:p>" + Run("10") + "</w:p></w:tc><w:tc><w:p>" + Run("20") + "</w:p></w:tc><w:tc><w:p>" + Complex(" = SUM(ABOVE) ", null) + "</w:p></w:tc></w:tr></w:tbl>";
        string checkBox = "<w:ffData><w:name w:val=\"Check1\"/><w:enabled/><w:checkBox><w:sizeAuto/><w:default w:val=\"0\"/><w:checked/></w:checkBox></w:ffData>";
        string dropDown = "<w:ffData><w:name w:val=\"Drop1\"/><w:ddList><w:result w:val=\"1\"/><w:listEntry w:val=\"Red\"/><w:listEntry w:val=\"Blueberry\"/></w:ddList></w:ffData>";
        string body =
            /* 0 */ "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"Target\"/>" + Run("Target text") + "<w:bookmarkEnd w:id=\"1\"/></w:p>"
            /* 1 */ + "<w:p>" + Run("cached ") + Complex(" REF Target \\h ", "Cached ref") + Run(" end") + "</w:p>"
            /* 2 */ + "<w:p>" + Run("page ") + Complex(" PAGE ", "57") + Run(" of ") + Complex(" NUMPAGES ", "99") + Run(" see ") + Complex(" PAGEREF Target \\h ", "31") + "</w:p>"
            /* 3 */ + "<w:p>" + Begin() + Code(" IF ") + Complex(" = 2+2 ", "4") + Code(" = 4 \"nestedyes\" \"nestedno\" ") + End + "</w:p>"
            /* 4 */ + "<w:p>" + Run("Figure ") + Complex(" SEQ Figure ", "1") + "</w:p>"
            /* 5 */ + "<w:p>" + Run("Figure ") + Complex(" SEQ Figure ", null) + Run(" and ") + Complex(" SEQ Figure \\* roman ", null) + "</w:p>"
            /* 6 */ + "<w:p>" + Complex(" = 1+1 ", "staleresult", " w:dirty=\"true\"") + Run(" | ") + Complex(" = 1+1 ", "lockedresult", " w:dirty=\"true\" w:fldLock=\"true\"") + "</w:p>"
            /* 7 */ + table
            /* 8 */ + "<w:p>" + Complex(" AUTHOR ", null) + Run(" / ") + Complex(" DOCPROPERTY Client ", null) + Run(" / ") + Complex(" DOCVARIABLE Var1 ", null) + "</w:p>"
            /* 9 */ + "<w:p>" + Begin("", checkBox) + Code(" FORMCHECKBOX ") + End + Begin("", dropDown) + Code(" FORMDROPDOWN ") + End + "</w:p>"
            /* 10 */ + "<w:p>" + Complex(" MACROBUTTON NoMacro Press here ", null) + Run(" ") + Complex(" GOTOBUTTON Target Go ", null) + "</w:p>"
            /* 11 */ + "<w:p>" + Run("index") + Complex(" XE \"hiddenentry\" ", null) + Complex(" TC \"hiddentoc\" ", null) + "</w:p>"
            /* 12 */ + "<w:p>" + Complex(" EQ \\f(1,2) ", null) + "</w:p>"
            /* 13 */ + "<w:p>" + Complex(" MERGEFIELD FirstName ", null) + "<w:fldSimple w:instr=\" = 6*7 \"/>" + "<w:fldSimple w:instr=\" QUOTE simplecached \">" + Run("simplecached") + "</w:fldSimple></w:p>"
            /* 14 */ + "<w:p>" + Complex(" REF Target ", null) + Run(" | ") + Complex(" HYPERLINK \\l \"Target\" ", "go to target") + "</w:p>"
            /* 15 */ + "<w:p>" + Begin() + Code(" QUOTE \"multi") + "</w:p>"
            /* 16 */ + "<w:p>" + Code("paragraph\" ") + Separate + Run("multiresult") + End + Run(" after") + "</w:p>"
            /* 17 */ + "<w:p>" + Run("Figure ") + Complex(" SEQ Figure ", null) + "</w:p>";
        const string coreProperties = "<?xml version=\"1.0\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:creator>Moshe Cohen</dc:creator></cp:coreProperties>";
        const string customProperties = "<?xml version=\"1.0\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Client\"><vt:lpwstr>Acme Ltd</vt:lpwstr></property></Properties>";
        var parts = new Dictionary<string, byte[]>
        {
            ["docProps/core.xml"] = Encoding.UTF8.GetBytes(coreProperties),
            ["docProps/custom.xml"] = Encoding.UTF8.GetBytes(customProperties),
        };
        string settings = "<w:docVars><w:docVar w:name=\"Var1\" w:val=\"VariableValue\"/></w:docVars>";
        List<string> html = DocxFidelityTest.RenderAll("fields-v1", body, settings: settings, binaryParts: parts);
        TestReport.Check(html[1].Contains("Cached ref") && html[1].Contains("href=\"#bookmark-Target\"") && !html[1].Contains("REF"), "REF \\h: the cached result shows, as a link, the code does not");
        TestReport.Check(!html[2].Contains("57") && !html[2].Contains("99") && !html[2].Contains(">31<") && html[2].Split("field-page-placeholder").Length == 4
                         && html[2].Contains("<a class=\"link\" href=\"#bookmark-Target\"") , "PAGE / NUMPAGES / PAGEREF: placeholders (PAGEREF \\h still links), page numbers not shown");
        TestReport.Check(html[3].Contains("nestedyes") && !html[3].Contains(">4<") && !html[3].Contains("nestedno"), "nested field in a code: its result is part of the code; IF without a result computed");
        TestReport.Check(html[4].Contains(">1<") || html[4].Contains("1"), "SEQ with a cached result shows it");
        TestReport.Check(html[5].Contains("Figure 2") && html[5].Contains(" iii"), $"SEQ without a result counts on from the index's counters (2, then iii): {Strip(html[5])}");
        TestReport.Check(Strip(html[17]).Contains("Figure 4"), $"SEQ counters reach a later block (Figure 4): {Strip(html[17])}");
        TestReport.Check(html[6].Contains("2") && !html[6].Contains("staleresult") && html[6].Contains("lockedresult"), "dirty field recomputed, dirty + locked field keeps its result");
        TestReport.Check(Strip(html[7]).EndsWith("10203"), $"table formula: SUM(ABOVE) over the cell above (3): {Strip(html[7])}");
        TestReport.Check(html[8].Contains("Moshe Cohen") && html[8].Contains("Acme Ltd") && html[8].Contains("VariableValue"), "AUTHOR / DOCPROPERTY / DOCVARIABLE from the package's properties");
        TestReport.Check(html[9].Contains("☒") && html[9].Contains("Blueberry") && !html[9].Contains("Red"), "form fields: checked box, selected drop-down entry");
        TestReport.Check(html[10].Contains("Press here") && html[10].Contains("href=\"#bookmark-Target\"") && html[10].Contains(">Go<"), "MACROBUTTON text, GOTOBUTTON link");
        TestReport.Check(!html[11].Contains("hidden"), "XE / TC: nothing shows");
        TestReport.Check(html[12].Contains("<mfrac>"), "EQ drawn as MathML");
        TestReport.Check(Strip(html[13]) == "«FirstName»42simplecached", "MERGEFIELD, empty fldSimple formula, fldSimple cached result: " + Strip(html[13]));
        TestReport.Check(html[14].Contains("Target text") && html[14].Contains(">go to target</a>"), "REF without a result: the bookmark's text; HYPERLINK \\l links");
        TestReport.Check(!html[15].Contains("multi") && html[16].Contains("multiresult") && !html[16].Contains("paragraph"), "a code across paragraphs stays hidden (index-time open-field state)");

        // search: the byte scanner and the reference reader agree, and find what shows, never codes / page numbers / hidden fields
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-fields-v1.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields") };
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        (string query, int[] expectedBlocks)[] queries =
        [
            ("Cached ref", [1]), ("REF", []), ("57", []), ("99", []), ("nestedyes", [3]), ("nestedno", []), ("Figure 2", [5]), ("iii", [5]), ("Figure 4", [17]),
            ("staleresult", []), ("lockedresult", [6]), ("Moshe Cohen", [8]), ("Acme Ltd", [8]), ("VariableValue", [8]), ("Blueberry", [9]), ("Press here", [10]),
            ("hiddenentry", []), ("«FirstName»", [13]), ("42", [13]), ("Target text", [0, 14]), ("multiresult", [16]), ("paragraph", []), ("MERGEFIELD", []),
        ];
        foreach ((string query, int[] expected) in queries)
        {
            int[] bytes = Hits(document, query, DocxWindowedDocument.DocxSearchTextMode.Bytes);
            int[] reference = Hits(document, query, DocxWindowedDocument.DocxSearchTextMode.ReferenceXmlReader);
            TestReport.Check(bytes.SequenceEqual(expected) && reference.SequenceEqual(expected),
                             $"search \"{query}\": blocks [{string.Join(",", expected)}] (bytes [{string.Join(",", bytes)}], reference [{string.Join(",", reference)}])");
        }
    }

    /// <summary>
    /// Fields computed from the index: STYLEREF, NOTEREF, REF \n \p, TOC / INDEX / TOA built inside their block, BARCODE,
    /// ADVANCE \u (expected values hand-worked: Word 16 wrote REF \n of list label "1." as 1, REF \n \p as "1 above").
    /// </summary>
    private static void DocxLookups()
    {
        const string styles = """
            <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
            <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
            """;
        const string numbering = """
            <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
            """;
        const string footnotes = """
            <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
            <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
            <w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t>the note</w:t></w:r></w:p></w:footnote>
            """;
        string Heading(int level, string text) => $"<w:p><w:pPr><w:pStyle w:val=\"Heading{level}\"/></w:pPr>{Run(text)}</w:p>";
        string body =
            /* 0 */ "<w:p>" + Complex(" TOC \\o \"1-2\" ", null) + "</w:p>"
            /* 1 */ + Heading(1, "Chapter one")
            /* 2 */ + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:bookmarkStart w:id=\"5\" w:name=\"Item\"/>" + Run("Listed item") + "<w:bookmarkEnd w:id=\"5\"/></w:p>"
            /* 3 */ + "<w:p>" + Run("see ") + Complex(" REF Item \\n ", null) + Run(" / ") + Complex(" REF Item \\n \\p ", null) + Run(" / ") + Complex(" REF Later \\p ", null) + "</w:p>"
            /* 4 */ + "<w:p>" + Run("running ") + Complex(" STYLEREF \"heading 1\" ", null) + Run(" / ") + Complex(" STYLEREF 2 ", null) + Run(" / ") + Complex(" STYLEREF \"No such\" ", null) + "</w:p>"
            /* 5 */ + Heading(2, "Section one-one")
            /* 6 */ + "<w:p><w:bookmarkStart w:id=\"6\" w:name=\"Note1\"/>" + Run("noted") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r><w:bookmarkEnd w:id=\"6\"/></w:p>"
            /* 7 */ + "<w:p>" + Run("note ") + Complex(" NOTEREF Note1 ", null) + Run(" / ") + Complex(" NOTEREF Note1 \\f ", null) + "</w:p>"
            /* 8 */ + "<w:p>" + Run("term") + Complex(" XE \"Apple:Green\" ", null) + Complex(" XE \"Banana\" ", null) + Complex(" TA \\l \"Smith v. Jones\" \\s \"Smith\" \\c 1 ", null) + "</w:p>"
            /* 9 */ + "<w:p>" + Complex(" INDEX \\e \" - \" ", null) + "</w:p>"
            /* 10 */ + "<w:p>" + Complex(" TOA \\c 1 ", null) + "</w:p>"
            /* 11 */ + "<w:p>" + Run("zip ") + Complex(" BARCODE \"55555-1237\" \\u ", null) + "</w:p>"
            /* 12 */ + "<w:p>" + Run("base") + Complex(" ADVANCE \\u 3 ", null) + Run("raised") + "</w:p>"
            /* 13 */ + "<w:p><w:bookmarkStart w:id=\"7\" w:name=\"Later\"/>" + Run("later") + "<w:bookmarkEnd w:id=\"7\"/></w:p>";
        List<string> html = DocxFidelityTest.RenderAll("fields-lookups-v1", body, styles, numbering, footnotes: footnotes);
        TestReport.Check(Strip(html[0]).Contains("Chapter one") && Strip(html[0]).Contains("Section one-one") && html[0].Split("toc-line").Length == 3 && html[0].Contains("field-page-placeholder"),
                         $"TOC without a result: built from the headings inside its block, page numbers as placeholders ({Strip(html[0])})");
        TestReport.Check(Strip(html[3]) == "see 1 / 1 above / below", $"REF \\n / \\n \\p / \\p (Word 16: 1, 1 above): {Strip(html[3])}");
        TestReport.Check(Strip(html[4]) == "running Chapter one / Section one-one / " + WordFieldEvaluator.StyleNotFound("No such"),
                         $"STYLEREF: back to the nearest heading 1, forward to heading 2 when none is before, Word's error for a style no paragraph has: {Strip(html[4])}");
        TestReport.Check(Strip(html[7]) == "note 1 / 1" && html[7].Contains("note-reference"), $"NOTEREF: the note's mark; \\f as a note reference: {Strip(html[7])}");
        TestReport.Check(Strip(html[9]).Contains("Apple") && Strip(html[9]).Contains("Green") && Strip(html[9]).Contains("Banana") && html[9].Contains("index-line") && Strip(html[9]).Contains(" - "),
                         $"INDEX without a result: built from the XE entries, levels and the \\e separator: {Strip(html[9])}");
        TestReport.Check(Strip(html[10]).Contains("Smith v. Jones"), $"TOA \\c 1: the TA long citations of category 1: {Strip(html[10])}");
        TestReport.Check(html[11].Contains("<svg class=\"field-barcode\"") && html[11].Contains("aria-label=\"5555512372\"") && html[11].Split("<rect").Length - 1 == 52,
                         "BARCODE: POSTNET with check digit 2 (the worked example 55555-1237), 52 bars");
        TestReport.Check(html[12].Contains("position:relative;top:-3pt") && html[12].IndexOf("base") < html[12].IndexOf("top:-3pt") && html[12].IndexOf("top:-3pt") < html[12].IndexOf("raised"),
                         "ADVANCE \\u 3: the text after it is raised 3pt");
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-fields-lookups-v1.docx");
        using var document = (DocxWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields-lookups") });
        foreach ((string query, int[] expected) in new (string, int[])[] { ("1 above", [3]), ("Chapter one", [0, 1, 4]), ("Section one-one", [0, 4, 5]), ("Smith v. Jones", [10]), ("Banana", [9]) })
        {
            int[] bytes = Hits(document, query, DocxWindowedDocument.DocxSearchTextMode.Bytes), reference = Hits(document, query, DocxWindowedDocument.DocxSearchTextMode.ReferenceXmlReader);
            TestReport.Check(bytes.SequenceEqual(expected) && reference.SequenceEqual(expected), $"search \"{query}\": [{string.Join(",", expected)}] (bytes [{string.Join(",", bytes)}], reference [{string.Join(",", reference)}])");
        }
    }

    /// <summary>
    /// REF \n \r \w on a three-level list and \p in Hebrew text, against Word 16 (COM, 2026-11-04: levels "%1)" "%2)" "%3)" at
    /// 2) b) i) → \n i), \w 2)b)i), \r i) from inside 2)b) and from outside the list; Hebrew \p "מעל"), TOC \h block links,
    /// BARCODE \f FIM patterns.
    /// </summary>
    private static void DocxListContexts()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="1">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>
              <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3)"/></w:lvl>
            </w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
            """;
        const string styles = """<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>""";
        string Item(int level, string text, string extra = "") => $"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"{level}\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>{Run(text)}{extra}</w:p>";
        string body =
            /* 0 */ "<w:p>" + Complex(" TOC \\o \"1-1\" \\h ", null) + "</w:p>"
            /* 1 */ + "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>" + Run("Heading A") + "</w:p>"
            /* 2 */ + Item(0, "one")
            /* 3 */ + Item(0, "two")
            /* 4 */ + Item(1, "sub a")
            /* 5 */ + Item(1, "sub b ", Complex(" REF deep \\r ", null))
            /* 6 */ + "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"2\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:bookmarkStart w:id=\"9\" w:name=\"deep\"/>" + Run("subsub") + "<w:bookmarkEnd w:id=\"9\"/></w:p>"
            /* 7 */ + "<w:p>" + Complex(" REF deep \\n ", null) + Run(" | ") + Complex(" REF deep \\w ", null) + Run(" | ") + Complex(" REF deep \\r ", null) + "</w:p>"
            /* 8 */ + "<w:p>" + Complex(" BARCODE \"1\" \\f A ", null) + Complex(" BARCODE \"1\" \\f C ", null) + "</w:p>";
        List<string> html = DocxFidelityTest.RenderAll("fields-lists-v1", body, styles, numbering);
        TestReport.Check(Strip(html[7]) == "i) | 2)b)i) | i)", $"REF \\n / \\w / \\r from outside the list (Word 16: i), 2)b)i), i)): {Strip(html[7])}");
        TestReport.Check(Strip(html[5]).EndsWith("sub b i)"), $"REF \\r from inside 2)b): the shared levels left out (Word 16: i)): {Strip(html[5])}");
        TestReport.Check(html[0].Contains("href=\"#block-1\""), "TOC \\h: the entry links to its heading's block");
        TestReport.Check(html[8].Contains("aria-label=\"FIM A\"") && html[8].Split("<rect").Length - 1 == 5 + 6, "BARCODE \\f A / C: FIM A (5 bars, 110010011), FIM C (6 bars, 110101011)");
        var hebrew = new WordFieldContext { IsRightToLeftText = true, CurrentBlock = 5, BookmarkBlock = _ => 2, ListLabelOfBlock = _ => "1." };
        TestReport.Check(Field("REF x \\p", hebrew) == "מעל" && Field("REF x \\n \\p", hebrew) == "1 מעל", "\\p in Hebrew text: מעל (Word 16 wrote it for a Hebrew field)");
        hebrew.CurrentBlock = 1;
        TestReport.Check(Field("REF x \\p", hebrew) == "מתחת", "\\p in Hebrew text below: מתחת");
    }

    /// <summary>
    /// Can a list-number lookup re-enter a block scan in progress (the RTF reader's crash)? A .docx of ~20 index spans whose
    /// REF \w fields each point to a list item in a span far from their own, more spans than the table caches: rendering a
    /// window and searching (fields computed while the search pieces are fed) make every lookup scan, evict and rescan other
    /// spans mid-render. Each REF must still show its target's full number (counters hand-worked: levels "%1)" "%2)" "%3)").
    /// </summary>
    private static void DocxListLookupsAcrossSpans()
    {
        const string numbering = """
            <w:abstractNum w:abstractNumId="1">
              <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1)"/></w:lvl>
              <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl>
              <w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%3)"/></w:lvl>
            </w:abstractNum>
            <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
            """;
        int[] pattern = [0, 1, 2, 2, 1, 2];               // 1) a) i) ii) b) i)
        string[] romans = ["i", "ii"];
        const int Items = 12060, Every = 402;        // Every a multiple of the pattern: item % Every == 2 is a level-3 item
        string padding = new string('x', 300);
        var body = new StringBuilder();
        var counters = new int[3];
        var fullNumbers = new Dictionary<string, string>();   // bookmark -> its full number
        var references = new List<(int block, string bookmark)>();
        int block = 0;
        for (int item = 0; item < Items; item++)
        {
            int level = pattern[item % pattern.Length];
            counters[level]++;
            for (int deeper = level + 1; deeper < 3; deeper++) counters[deeper] = 0;
            string bookmark = "";
            if (level == 2 && item % Every == 2)
            {
                bookmark = "bm" + item;
                fullNumbers[bookmark] = $"{counters[0]}){(char)('a' + counters[1] - 1)}){romans[counters[2] - 1]})";
            }
            body.Append("<w:p><w:pPr><w:numPr><w:ilvl w:val=\"").Append(level).Append("\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>");
            if (bookmark.Length > 0) body.Append("<w:bookmarkStart w:id=\"").Append(item).Append("\" w:name=\"").Append(bookmark).Append("\"/>");
            body.Append(Run("item " + item + " " + padding));
            if (bookmark.Length > 0) body.Append("<w:bookmarkEnd w:id=\"").Append(item).Append("\"/>");
            body.Append("</w:p>");
            block++;
            if (item % Every == Every - 1)
            {
                // a plain paragraph referring to bookmarks far behind and far ahead (other spans)
                foreach (int target in new[] { (item + Items / 2) % Items, (item + Items / 3) % Items })
                {
                    string targetBookmark = "bm" + (target / Every * Every + 2);
                    body.Append("<w:p>").Append(Run("ref ")).Append(Complex($" REF {targetBookmark} \\w ", null)).Append("</w:p>");
                    references.Add((block++, targetBookmark));
                }
            }
        }
        List<string> inMemory = DocxFidelityTest.RenderAll("fields-list-spans-v1", body.ToString(), "", numbering);
        int wrong = references.Count(reference => Strip(inMemory[reference.block]) != "ref " + fullNumbers[reference.bookmark]);
        TestReport.Check(wrong == 0, $"REF \\w across the document, one in-memory pass: {references.Count - wrong} of {references.Count} as Word numbers them");

        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-fields-list-spans-v1.docx");
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-fields-list-spans"), ForcedMode = WordIndexMode.Sidecar };
        using (DocxWindowedDocument.Open(path, policy)) { }
        using var document = DocxWindowedDocument.Open(path, policy);       // reopened: the compact table, spans rescanned on demand
        TestReport.Check(document.Index.Anchors.Length > 12, $"the fixture spans more anchors ({document.Index.Anchors.Length}) than the table caches (8)");
        var random = new Random(5);
        wrong = 0;
        foreach ((int referenceBlock, string bookmark) in references.OrderBy(_ => random.Next()))
        {
            int first = Math.Max(0, referenceBlock - 3);
            IReadOnlyList<WordDocumentBlockHtml> window = document.RenderBlocks(first, Math.Min(8, document.BlockCount - first));
            if (Strip(window[referenceBlock - first].Html) != "ref " + fullNumbers[bookmark]) wrong++;
        }
        TestReport.Check(wrong == 0, $"REF \\w from the compact table, windows in random order (each lookup scans and evicts other spans): {references.Count - wrong} of {references.Count} right");
        string sought = fullNumbers[references[^1].bookmark];
        int[] expected = references.Where(reference => fullNumbers[reference.bookmark].Contains(sought)).Select(reference => reference.block).ToArray();
        int[] bytes = Hits(document, sought, DocxWindowedDocument.DocxSearchTextMode.Bytes), referenceHits = Hits(document, sought, DocxWindowedDocument.DocxSearchTextMode.ReferenceXmlReader);
        TestReport.Check(bytes.SequenceEqual(expected) && referenceHits.SequenceEqual(expected),
                         $"search \"{sought}\" (computed REF \\w, lookups mid-piece): [{string.Join(",", expected)}] (bytes [{string.Join(",", bytes)}], reference [{string.Join(",", referenceHits)}])");
    }

    /// <summary>
    /// A TOC \h whose heading is far below it (block 601): the live check (vue-demo playwright/fieldsCheck.mjs on
    /// fidelity-v1-fields-toc-far-v1.docx) clicks the entry with the heading off screen and expects the jump to bring it in.
    /// </summary>
    private static void DocxTableOfContentsFarHeading()
    {
        const string styles = """<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>""";
        var body = new StringBuilder("<w:p>" + Complex(" TOC \\o \"1-1\" \\h ", null) + "</w:p>");
        for (int i = 1; i <= 600; i++) body.Append("<w:p>").Append(Run("filler paragraph " + i)).Append("</w:p>");
        body.Append("<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>").Append(Run("Far heading")).Append("</w:p>");
        body.Append("<w:p>").Append(Run("after the heading")).Append("</w:p>");
        List<string> html = DocxFidelityTest.RenderAll("fields-toc-far-v1", body.ToString(), styles);
        TestReport.Check(html[0].Contains("href=\"#block-601\"") && Strip(html[0]).Contains("Far heading"), "TOC \\h: the entry links to its heading 601 blocks below");
    }

    private static int[] Hits(DocxWindowedDocument document, string query, DocxWindowedDocument.DocxSearchTextMode mode)
    {
        var hits = new List<int>();
        int from = 0;
        while (true)
        {
            WordDocumentSearchBatch result = document.Search(query, WordDocumentSearchOptions.Default, from, 3, mode);
            hits.AddRange(result.HitBlocks);
            if (result.ReachedEnd) break;
            from = result.NextBlock;
        }
        return hits.ToArray();
    }

    internal static string Strip(string html) => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", ""));
}
