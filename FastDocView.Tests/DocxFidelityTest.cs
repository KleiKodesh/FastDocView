using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Rendering fidelity checks against Word's behaviour, one tiny synthetic .docx per construct (the
/// expected results were observed in Word 16 — COM read-back or PDF export — or read in the cited spec /
/// LibreOffice source). Each case builds its package here from body / styles / numbering XML strings.
/// </summary>
public static partial class DocxFidelityTest
{
    private const string WordNamespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"";

    public static void Run()
    {
        TestReport.Step("Docx fidelity: toggle properties, scripts, underline / strike, shading, tabs, lists, tables, revisions, links");
        foreach (Action check in Checks()) check();
    }

    private static IEnumerable<Action> Checks()
    {
        yield return ToggleProperties;
        yield return ScriptProperties;
        yield return UnderlineAndStrike;
        yield return ListNumbering;
        yield return ListLabelsAndParagraphMark;
        yield return ShadingPatterns;
        yield return Symbols;
        yield return CompatibilitySettings;
        yield return TableConditionalFormatting;
    }

    /// <summary>Writes a minimal package and returns every block's HTML.</summary>
    internal static List<string> RenderAll(string name, string body, string styles = "", string? numbering = null,
                                           string extraDocumentRelationships = "", string? settings = null)
    {
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-" + name + ".docx");
        if (File.Exists(path)) File.Delete(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            void Add(string entry, string text)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entry).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Add("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Add("word/_rels/document.xml.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdS\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>"
                + (numbering is null ? "" : "<Relationship Id=\"rIdN\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>")
                + (settings is null ? "" : "<Relationship Id=\"rIdT\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>")
                + extraDocumentRelationships + "</Relationships>");
            Add("word/styles.xml", $"<?xml version=\"1.0\"?><w:styles {WordNamespaces}><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>{styles}</w:styles>");
            if (numbering is not null) Add("word/numbering.xml", $"<?xml version=\"1.0\"?><w:numbering {WordNamespaces}>{numbering}</w:numbering>");
            if (settings is not null) Add("word/settings.xml", $"<?xml version=\"1.0\"?><w:settings {WordNamespaces}>{settings}</w:settings>");
            Add("word/document.xml", $"<?xml version=\"1.0\"?><w:document {WordNamespaces}><w:body>{body}</w:body></w:document>");
        }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        return document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
    }

    internal static string Paragraph(string? paragraphStyle, string runProperties, string text, string paragraphProperties = "")
        => "<w:p>" + (paragraphStyle is null && paragraphProperties.Length == 0 ? "" : "<w:pPr>" + (paragraphStyle is null ? "" : $"<w:pStyle w:val=\"{paragraphStyle}\"/>") + paragraphProperties + "</w:pPr>")
         + $"<w:r><w:rPr>{runProperties}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static bool IsBold(string html) => html.Contains("font-weight:bold");

    /// <summary>
    /// §17.7.3 toggle properties as Word 16 resolves them (COM read-back of a synthetic document, 2026-10-28):
    /// last value wins inside a basedOn chain, table / paragraph / character styles XOR, docDefaults on stays on,
    /// direct formatting absolute; b and bCs (i and iCs) are separate properties chosen by w:rtl.
    /// </summary>
    private static void ToggleProperties()
    {
        const string styles = """
            <w:style w:type="paragraph" w:styleId="P1"><w:name w:val="P1"/><w:rPr><w:b/><w:i/></w:rPr></w:style>
            <w:style w:type="paragraph" w:styleId="P2"><w:name w:val="P2"/><w:basedOn w:val="P1"/><w:rPr><w:b/></w:rPr></w:style>
            <w:style w:type="character" w:styleId="C1"><w:name w:val="C1"/><w:rPr><w:b/><w:iCs/></w:rPr></w:style>
            <w:style w:type="character" w:styleId="C2"><w:name w:val="C2"/><w:basedOn w:val="C1"/><w:rPr><w:b/></w:rPr></w:style>
            <w:style w:type="character" w:styleId="C3"><w:name w:val="C3"/><w:rPr><w:b w:val="0"/></w:rPr></w:style>
            <w:style w:type="table" w:styleId="T1"><w:name w:val="T1"/><w:rPr><w:b/></w:rPr></w:style>
            """;
        string body = Paragraph("P1", "", "p1")                                   // 0 bold
                    + Paragraph("P1", "<w:rStyle w:val=\"C1\"/>", "p1c1")          // 1 not bold (XOR)
                    + Paragraph("P2", "", "p2")                                   // 2 bold (chain: last wins)
                    + Paragraph(null, "<w:rStyle w:val=\"C2\"/>", "c2")            // 3 bold (chain: last wins)
                    + Paragraph("P1", "<w:rStyle w:val=\"C3\"/>", "p1c3")          // 4 bold (false contributes nothing)
                    + Paragraph("P1", "<w:rStyle w:val=\"C1\"/><w:b/>", "direct")  // 5 bold (direct absolute)
                    + Paragraph("P1", "<w:rStyle w:val=\"C1\"/><w:rtl/>", "rtl")   // 6 rtl: bCs off → not bold; iCs from C1 → italic
                    + Paragraph(null, "<w:bCs/>", "bcs")                           // 7 no rtl: bCs ignored → not bold
                    + "<w:tbl><w:tblPr><w:tblStyle w:val=\"T1\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid><w:tr><w:tc>"
                    + Paragraph("P1", "", "t-p1") + Paragraph(null, "", "t-normal") + Paragraph(null, "<w:rStyle w:val=\"C1\"/>", "t-c1")
                    + "</w:tc></w:tr></w:tbl>";
        List<string> html = RenderAll("toggles", body, styles);
        TestReport.Check(IsBold(html[0]) && !html[1].Contains("<span style=\"font-weight:bold") && html[1].Contains("font-weight:normal"), "toggle: bold paragraph style XOR bold character style = not bold");
        TestReport.Check(IsBold(html[2]) && IsBold(html[3]), "toggle: inside a basedOn chain the last value wins (paragraph and character styles)");
        TestReport.Check(!html[4].Contains("font-weight:normal"), "toggle: b=0 in a character style leaves the paragraph style's bold on");
        TestReport.Check(!html[5].Contains("font-weight:normal"), "toggle: direct <w:b/> is absolute");
        TestReport.Check(html[6].Contains("font-weight:normal") && !html[6].Contains("font-style:normal"),
                         "w:rtl run: bold from bCs (off), italic from iCs (C1 on, P1's i ignored)");
        TestReport.Check(!IsBold(html[7]), "bCs without w:rtl does not bold the run");
        string table = html[8];
        int p1 = table.IndexOf("t-p1", StringComparison.Ordinal), normal = table.IndexOf("t-normal", StringComparison.Ordinal), c1 = table.IndexOf("t-c1", StringComparison.Ordinal);
        string Before(int end) => table.Substring(0, end).Substring(table.Substring(0, end).LastIndexOf("<p", StringComparison.Ordinal));
        TestReport.Check(!Before(p1).Contains("font-weight:bold") && Before(normal).Contains("font-weight:bold") && Before(c1).Contains("font-weight:normal"),
                         "toggle: table style bold XOR paragraph / character style bold");
        List<string> defaults = RenderAll("toggles-defaults", Paragraph("P1", "<w:rStyle w:val=\"C1\"/>", "x"),
                                          "<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault></w:docDefaults>" + styles);
        TestReport.Check(!defaults[0].Contains("font-weight:normal") && IsBold(defaults[0]), "toggle: bold in docDefaults stays on (paragraph XOR character style does not turn it off)");
    }

    /// <summary>
    /// Which property set a run uses (Word 16 PDF export, 2026-10-28): w:rtl or w:cs → cs font / szCs / bCs for the
    /// whole run; neither → ascii / sz / b even for Hebrew letters and in a bidi paragraph; CJK → the eastAsia face.
    /// </summary>
    private static void ScriptProperties()
    {
        const string fonts = "<w:rFonts w:ascii=\"Courier New\" w:hAnsi=\"Courier New\" w:eastAsia=\"MS Mincho\" w:cs=\"Times New Roman\"/><w:sz w:val=\"40\"/><w:szCs w:val=\"20\"/>";
        const string hebrew = "שלום";
        List<string> html = RenderAll("scripts",
            Paragraph(null, fonts + "<w:b/>", "abc " + hebrew)                                  // 0 ascii, 20pt, bold
          + Paragraph(null, fonts + "<w:rtl/><w:bCs/>", "abc " + hebrew)                        // 1 cs, 10pt, bold
          + Paragraph(null, fonts + "<w:cs/>", "abc " + hebrew)                                 // 2 cs, 10pt
          + Paragraph(null, fonts + "<w:b/>", "abc " + hebrew, "<w:bidi/>")                     // 3 bidi paragraph: still ascii, 20pt, bold
          + Paragraph(null, fonts, "abc 漢字 " + hebrew));                               // 4 CJK in MS Mincho
        TestReport.Check(html[0].Contains("font-family:'Courier New'") && html[0].Contains("font-size:20pt") && IsBold(html[0]), "no rtl / cs: Hebrew text uses the ascii font, sz and b");
        TestReport.Check(html[1].Contains("font-family:'Times New Roman'") && html[1].Contains("font-size:10pt") && IsBold(html[1]), "w:rtl: cs font, szCs, bCs");
        TestReport.Check(html[2].Contains("font-family:'Times New Roman'") && html[2].Contains("font-size:10pt"), "w:cs: cs font and szCs");
        TestReport.Check(html[3].Contains("font-family:'Courier New'") && html[3].Contains("font-size:20pt") && IsBold(html[3]), "w:bidi paragraph alone does not switch runs to the cs set");
        TestReport.Check(html[4].Contains("<span style=\"font-family:'MS Mincho'\">漢字</span>"), "CJK characters switch to the eastAsia face");
    }

    /// <summary>
    /// Underline / strike as Word 16 draws them (PDF export of a synthetic document, 2026-10-28): spaces between words and
    /// tabs underlined, the spaces ending the paragraph not; "words" skips spaces; a bare &lt;w:u/&gt; draws nothing; a style's
    /// underline switched off by the run (w:u none) really disappears; dstrike + a patterned underline keep both patterns.
    /// </summary>
    private static void UnderlineAndStrike()
    {
        const string styles = """<w:style w:type="paragraph" w:styleId="U"><w:name w:val="U"/><w:rPr><w:u w:val="single"/><w:strike/></w:rPr></w:style>""";
        static string Run(string properties, string text) => $"<w:r><w:rPr>{properties}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        List<string> html = RenderAll("underline", 
            "<w:p>" + Run("<w:u w:val=\"single\"/>", "abc   ") + "</w:p>"                                        // 0 trailing spaces at paragraph end: not underlined
          + "<w:p>" + Run("<w:u w:val=\"single\"/>", "abc   ") + Run("", "def") + "</w:p>"                     // 1 spaces before more text: underlined
          + "<w:p>" + Run("<w:u w:val=\"words\"/>", "ab cd ef") + "</w:p>"                                // 2 words
          + "<w:p>" + Run("<w:u w:val=\"single\"/>", "abc") + "<w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:tab/></w:r>" + Run("", "x") + "</w:p>"   // 3 tab underlined
          + "<w:p>" + Run("<w:u/>", "bare") + "</w:p>"                                                       // 4 bare w:u: nothing
          + "<w:p><w:pPr><w:pStyle w:val=\"U\"/></w:pPr>" + Run("<w:u w:val=\"none\"/><w:strike w:val=\"0\"/>", "off") + Run("", "on") + "</w:p>"   // 5
          + "<w:p>" + Run("<w:u w:val=\"dotted\"/><w:dstrike/>", "both") + "</w:p>", styles);                 // 6
        const string line = "<span style=\"text-decoration:underline;\">";
        TestReport.Check(html[0].Contains(line + "abc</span>   "), "underline stops before the spaces that end the paragraph");
        TestReport.Check(html[1].Contains(line + "abc   </span>"), "spaces followed by more text are underlined");
        TestReport.Check(html[2].Contains(line + "ab</span> " + line + "cd</span>&#160;" + line + "ef</span>") || html[2].Contains(line + "ab</span> " + line + "cd</span> " + line + "ef</span>"),
                         "w:u words: no line under spaces (breaking or non-breaking)");
        TestReport.Check(html[3].Contains(line + "\t</span>"), "an underlined tab is underlined");
        TestReport.Check(!html[4].Contains("underline"), "a bare <w:u/> (no w:val) draws no underline");
        TestReport.Check(!html[5].Contains("<p style=\"text-decoration") && !html[5].Contains("text-decoration:underline;\">off") && html[5].Contains(line + "on</span>")
                         && html[5].Contains("line-through;\">") , "a style's underline / strike switched off by a run disappears; other runs keep them on the run");
        TestReport.Check(html[6].Contains("text-decoration:line-through;text-decoration-style:double;") && html[6].Contains("text-decoration:underline;text-decoration-style:dotted;"),
                         "dstrike and a dotted underline keep their own patterns (separate elements)");
    }
}
