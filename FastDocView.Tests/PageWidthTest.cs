using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Page text width per section (IWordWindowedDocument.Sections): page width minus left / right margins
/// and gutter, from .docx w:sectPr, .doc PlcfSed / Sepx, .rtf \paperw \margl \margr (+ per-section
/// forms), .odt / .fodt master pages; the Word default (Letter, 1 in margins = 9360 twips) without one.
/// </summary>
public static class PageWidthTest
{
    private const string WordNamespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"";

    public static void Run(IEnumerable<string> realPaths)
    {
        TestReport.Step("Page text width per section");
        Docx();
        Rtf();
        Odt();
        foreach (string path in realPaths.Where(File.Exists))
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy());
            while (!document.IsIndexComplete) document.ScanAhead(100_000);
            IReadOnlyList<WordDocumentSection> sections = document.Sections;
            TestReport.Check(sections.Count > 0 && sections[0].FirstBlockIndex == 0
                             && sections.All(section => section.TextWidthTwips >= WordDocumentPageSetup.MinimumTextWidthTwips),
                             $"{Path.GetExtension(path)} sample: {sections.Count} section(s): {Describe(sections.Take(5).ToList())}");
        }
    }

    private static WordIndexPolicy Policy(WordIndexMode? mode = null)
        => new() { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-page-width"), ForcedMode = mode };

    private static string Describe(IReadOnlyList<WordDocumentSection> sections) => string.Join(" ", sections.Select(s => $"{s.FirstBlockIndex}:{s.TextWidthTwips}"));

    private static void Expect(string what, IReadOnlyList<WordDocumentSection> sections, params (int first, int width)[] expected)
        => TestReport.Check(sections.Select(s => (s.FirstBlockIndex, s.TextWidthTwips)).SequenceEqual(expected), $"{what}: {Describe(sections)}");

    private static IReadOnlyList<WordDocumentSection> DocxSections(string name, string body)
    {
        string path = Path.Combine(TestReport.OutputDirectory, "page-width-" + name + ".docx");
        if (File.Exists(path)) File.Delete(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            void Add(string entry, string text)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entry).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("[Content_Types].xml", "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Add("_rels/.rels", "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Add("word/document.xml", $"<?xml version=\"1.0\"?><w:document {WordNamespaces}><w:body>{body}</w:body></w:document>");
        }
        IReadOnlyList<WordDocumentSection> inMemory, sidecar;
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy())) inMemory = document.Sections;
        // the same through a saved sidecar index, then reopened from it
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy(WordIndexMode.Sidecar))) { }
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy(WordIndexMode.Sidecar))) sidecar = document.Sections;
        TestReport.Check(inMemory.SequenceEqual(sidecar), $"docx {name}: the sidecar index keeps the sections");
        return inMemory;
    }

    private static string P(string text, string sectPr = "") => $"<w:p>{(sectPr.Length > 0 ? "<w:pPr>" + sectPr + "</w:pPr>" : "")}<w:r><w:t>{text}</w:t></w:r></w:p>";

    private static void Docx()
    {
        Expect("docx without page setup: Word default", DocxSections("none", P("a") + P("b")), (0, 9360));
        Expect("docx A4, 1 in margins", DocxSections("a4", P("a") + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:gutter=\"0\"/></w:sectPr>"), (0, 9026));
        Expect("docx gutter narrows the text", DocxSections("gutter", P("a") + "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:left=\"1440\" w:right=\"1440\" w:gutter=\"720\"/></w:sectPr>"), (0, 8640));
        Expect("docx two sections (a pPr sectPr ends the first at its paragraph)",
               DocxSections("two", P("a") + P("b", "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:left=\"1800\" w:right=\"1800\"/></w:sectPr>") + P("c")
                                   + "<w:sectPr><w:pgSz w:w=\"16838\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/></w:sectPr>"),
               (0, 8640), (2, 13958));
        Expect("docx a section break in a content control counts; one in a table cell does not",
               DocxSections("sdt", "<w:sdt><w:sdtContent>" + P("a", "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:left=\"1800\" w:right=\"1800\"/></w:sectPr>") + "</w:sdtContent></w:sdt>"
                                   + "<w:tbl><w:tr><w:tc>" + P("x", "<w:sectPr><w:pgSz w:w=\"30000\"/></w:sectPr>") + "</w:tc></w:tr></w:tbl>" + P("b")
                                   + "<w:sectPr><w:pgSz w:w=\"11906\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/></w:sectPr>"),
               (0, 8640), (1, 9026));
        Expect("docx a w:sectPrChange's old values are ignored",
               DocxSections("change", P("a") + "<w:sectPr><w:pgSz w:w=\"11906\"/><w:pgMar w:left=\"1440\" w:right=\"1440\"/><w:sectPrChange w:id=\"1\"><w:sectPr><w:pgSz w:w=\"20000\"/></w:sectPr></w:sectPrChange></w:sectPr>"),
               (0, 9026));
    }

    private static IReadOnlyList<WordDocumentSection> OpenText(string name, string extension, string content)
    {
        string path = Path.Combine(TestReport.OutputDirectory, "page-width-" + name + extension);
        File.WriteAllText(path, content, new UTF8Encoding(false));
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy());
        while (!document.IsIndexComplete) document.ScanAhead(100_000);
        return document.Sections;
    }

    private static void Rtf()
    {
        Expect("rtf defaults (12240 - 2 x 1800)", OpenText("rtf-default", ".rtf", "{\\rtf1\\ansi a\\par b\\par}"), (0, 8640));
        Expect("rtf document setup, then a section of its own; a style's \\pgwsxn is ignored",
               OpenText("rtf-sections", ".rtf", "{\\rtf1\\ansi\\paperw11906\\margl1440\\margr1440{\\stylesheet{\\s1\\pgwsxn1000 x;}}a\\par b\\sect\\sectd\\pgwsxn12240\\marglsxn1800\\margrsxn1800\\guttersxn360 c\\par}"),
               (0, 9026), (2, 8280));
        Expect("rtf a \\sect inside a table row ends the section with the row",
               OpenText("rtf-row-sect", ".rtf", @"{\rtf1\ansi\paperw11906\margl1440\margr1440 \trowd\cellx1000\intbl a\cell\sect\row\pard\sectd\marglsxn1800\margrsxn1800\pgwsxn12240 b\par}"),
               (0, 9026), (1, 8640));
        Expect("text file: Word default", OpenText("text", ".txt", "a\r\nb\r\n"), (0, 9360));
    }

    private const string OdfNamespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" " +
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
        "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\"";

    private static void Odt()
    {
        string flat = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + $"<office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles>"
            + "<style:style style:name=\"P2\" style:family=\"paragraph\" style:master-page-name=\"Landscape\"/>"
            + "<style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"21cm\" fo:page-height=\"29.7cm\" fo:margin-left=\"2cm\" fo:margin-right=\"2cm\"/></style:page-layout>"
            + "<style:page-layout style:name=\"pm2\"><style:page-layout-properties fo:page-width=\"11in\" fo:margin-left=\"1in\" fo:margin-right=\"0.5in\"/></style:page-layout>"
            + "</office:automatic-styles>"
            + "<office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\"/><style:master-page style:name=\"Landscape\" style:page-layout-name=\"pm2\"/></office:master-styles>"
            + "<office:body><office:text><text:p>a</text:p><text:p text:style-name=\"P2\">b</text:p><text:p>c</text:p></office:text></office:body></office:document>";
        // 21 cm = 11906 twips (rounded), 2 cm = 1134; 11 in - 1.5 in = 13680
        Expect("fodt master pages (Standard, then a paragraph naming Landscape)", OpenText("odt", ".fodt", flat), (0, 9638), (1, 13680));
    }
}
