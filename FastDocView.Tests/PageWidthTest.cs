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
        Margins();
        foreach (string path in realPaths.Where(File.Exists))
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, Policy());
            while (!document.IsIndexComplete) document.ScanAhead(100_000);
            IReadOnlyList<WordDocumentSection> sections = document.Sections;
            TestReport.Check(sections.Count > 0 && sections[0].FirstBlockIndex == 0
                             && sections.All(section => section.TextWidthTwips >= WordDocumentPageSetup.MinimumTextWidthTwips)
                             && sections.All(section => section.TopMarginTwips >= 0 && section.BottomMarginTwips >= 0),
                             $"{Path.GetExtension(path)} sample: {sections.Count} section(s): {Describe(sections.Take(5).ToList())} margins {DescribeMargins(sections.Take(5).ToList())}");
        }
    }

    private static WordIndexPolicy Policy(WordIndexMode? mode = null)
        => new() { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-page-width"), ForcedMode = mode };

    private static string Describe(IReadOnlyList<WordDocumentSection> sections) => string.Join(" ", sections.Select(s => $"{s.FirstBlockIndex}:{s.TextWidthTwips}"));

    private static string DescribeMargins(IReadOnlyList<WordDocumentSection> sections) => string.Join(" ", sections.Select(s => $"{s.FirstBlockIndex}:{s.TopMarginTwips}/{s.BottomMarginTwips}"));

    private static void ExpectMargins(string what, IReadOnlyList<WordDocumentSection> sections, params (int first, int top, int bottom)[] expected)
        => TestReport.Check(sections.Select(s => (s.FirstBlockIndex, s.TopMarginTwips, s.BottomMarginTwips)).SequenceEqual(expected), $"{what}: {DescribeMargins(sections)}");

    /// <summary>Top / bottom page margins per section: one sample per format, the Word default without page setup,
    /// negative docx / rtf values taken as their absolute value, a gutter at the top added to the top margin.</summary>
    private static void Margins()
    {
        ExpectMargins("docx margins without page setup: 1 in", DocxSections("margins-none", P("a")), (0, 1440, 1440));
        ExpectMargins("docx w:top / w:bottom (negative = same distance)",
                      DocxSections("margins", P("a") + "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:top=\"-1000\" w:bottom=\"2000\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>"),
                      (0, 1000, 2000));
        ExpectMargins("docx two sections with their own margins (a width-equal section still splits)",
                      DocxSections("margins-two", P("a", "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:top=\"720\" w:bottom=\"720\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>") + P("b")
                                                  + "<w:sectPr><w:pgSz w:w=\"12240\"/><w:pgMar w:top=\"1800\" w:bottom=\"900\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>"),
                      (0, 720, 720), (1, 1800, 900));
        ExpectMargins("rtf defaults (\\margt / \\margb 1440)", OpenText("rtf-margins-default", ".rtf", "{\\rtf1\\ansi a\\par}"), (0, 1440, 1440));
        ExpectMargins("rtf document \\margt, a section's \\margbsxn (negative), gutter at the top (\\gutterprl) added to the top",
                      OpenText("rtf-margins", ".rtf", @"{\rtf1\ansi\gutterprl\gutter360\margt1000\margb1100 a\par b\sect\sectd\margbsxn-1700 c\par}"),
                      (0, 1360, 1100), (2, 1360, 1700));
        string flat = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + $"<office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles>"
            + "<style:style style:name=\"P2\" style:family=\"paragraph\" style:master-page-name=\"Other\"/>"
            + "<style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"21cm\" fo:margin-left=\"2cm\" fo:margin-right=\"2cm\" fo:margin-top=\"1in\" fo:margin-bottom=\"0.5in\"/></style:page-layout>"
            + "<style:page-layout style:name=\"pm2\"><style:page-layout-properties fo:page-width=\"21cm\" fo:margin=\"2cm\"/></style:page-layout>"
            + "</office:automatic-styles>"
            + "<office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\"/><style:master-page style:name=\"Other\" style:page-layout-name=\"pm2\"/></office:master-styles>"
            + "<office:body><office:text><text:p>a</text:p><text:p text:style-name=\"P2\">b</text:p></office:text></office:body></office:document>";
        ExpectMargins("fodt fo:margin-top / -bottom, then fo:margin (2 cm = 1134)", OpenText("odt-margins", ".fodt", flat), (0, 1440, 720), (1, 1134, 1134));
        ExpectMargins("rtf negative \\margt / \\margb: the absolute value; no \\gutterprl: the gutter stays off the top",
                      OpenText("rtf-margins-negative", ".rtf", @"{\rtf1\ansi\gutter360\margt-900\margb-800 a\par}"), (0, 900, 800));
        ExpectMargins("odt (zip) styles.xml page layouts: margin-top / -bottom, then a master page from content",
                      OdtZipSections("odt-zip-margins",
                          "<office:automatic-styles>"
                          + "<style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"8.5in\" fo:margin-left=\"1in\" fo:margin-right=\"1in\" fo:margin-top=\"0.75in\" fo:margin-bottom=\"1.25in\"/></style:page-layout>"
                          + "<style:page-layout style:name=\"pm2\"><style:page-layout-properties fo:page-width=\"8.5in\" fo:margin-left=\"1in\" fo:margin-right=\"1in\" fo:margin-top=\"2in\" fo:margin-bottom=\"0.5in\"/></style:page-layout>"
                          + "</office:automatic-styles><office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\"/><style:master-page style:name=\"Other\" style:page-layout-name=\"pm2\"/></office:master-styles>",
                          "<office:automatic-styles><style:style style:name=\"P2\" style:family=\"paragraph\" style:master-page-name=\"Other\"/></office:automatic-styles>"
                          + "<office:body><office:text><text:p>a</text:p><text:p text:style-name=\"P2\">b</text:p></office:text></office:body>"),
                      (0, 1080, 1800), (1, 2880, 720));
        // where the file is silent: Word's defaults (Letter, 1 in margins), not LibreOffice's page style
        string bare = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + $"<office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles><style:style style:name=\"P2\" style:family=\"paragraph\" style:master-page-name=\"Other\"/>"
            + "<style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"21cm\" fo:margin-left=\"2cm\" fo:margin-right=\"2cm\"/></style:page-layout>"
            + "<style:page-layout style:name=\"pm2\"><style:page-layout-properties fo:page-width=\"21cm\" fo:margin-top=\"1in\"/></style:page-layout>"
            + "</office:automatic-styles>"
            + "<office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\"/><style:master-page style:name=\"Other\" style:page-layout-name=\"pm2\"/></office:master-styles>"
            + "<office:body><office:text><text:p>a</text:p><text:p text:style-name=\"P2\">b</text:p></office:text></office:body></office:document>";
        ExpectMargins("fodt page layout without fo:margin-top / -bottom: Word's 1 in; only fo:margin-top set: bottom 1 in",
                      OpenText("odt-margins-missing", ".fodt", bare), (0, 1440, 1440), (1, 1440, 1440));
        // pm2 has a page width but no side margins: Word's 1 in each side (11906 - 2880 = 9026)
        Expect("fodt missing fo:margin-left / -right: Word's 1 in", OpenText("odt-margins-missing", ".fodt", bare), (0, 9638), (1, 9026));
        string noWidth = bare.Replace(" fo:page-width=\"21cm\" fo:margin-top=\"1in\"", " fo:margin-top=\"1in\"");
        Expect("fodt layout without fo:page-width: Letter (12240 - 2 x 1 in = 9360)", OpenText("odt-no-page-width", ".fodt", noWidth), (0, 9638), (1, 9360));
        string noLayout = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
            + $"<office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:body><office:text><text:p>a</text:p></office:text></office:body></office:document>";
        ExpectMargins("fodt without any page layout: Word's 1 in", OpenText("odt-margins-no-layout", ".fodt", noLayout), (0, 1440, 1440));
        Expect("fodt without any page layout: Letter, 1 in (9360)", OpenText("odt-margins-no-layout", ".fodt", noLayout), (0, 9360));
        DocSepx();
        ExpectMargins("text file: 1 in", OpenText("text-margins", ".txt", "a\r\n"), (0, 1440, 1440));
        string docSample = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "apache-poi", "47304.doc");
        if (File.Exists(docSample))
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(docSample, Policy());
            while (!document.IsIndexComplete) document.ScanAhead(100_000);
            IReadOnlyList<WordDocumentSection> sections = document.Sections;
            TestReport.Check(sections.Count > 0 && sections.All(section => section.TopMarginTwips > 0 && section.BottomMarginTwips > 0),
                             $"doc sample 47304.doc margins (Sepx dyaTop / dyaBottom): {DescribeMargins(sections)}");
        }
        else TestReport.Check(false, "doc sample test-data/apache-poi/47304.doc not found");
    }

    /// <summary>.doc section sprms (a Sepx grpprl fed straight to the page-setup reader): dyaTop / dyaBottom, negative ones,
    /// and a gutter at the top (Dop iGutterPos) added to the top margin and taken off the width.</summary>
    private static void DocSepx()
    {
        static (int width, int left, int top, int bottom) Setup(bool isGutterAtTop, params (ushort opcode, short value)[] sprms)
        {
            var grpprl = new List<byte>();
            foreach ((ushort opcode, short value) in sprms) { grpprl.AddRange(BitConverter.GetBytes(opcode)); grpprl.AddRange(BitConverter.GetBytes(value)); }
            var stream = new MemoryStream();
            stream.Write(BitConverter.GetBytes((short)grpprl.Count), 0, 2);
            stream.Write(grpprl.ToArray(), 0, grpprl.Count);
            return FastDocView.DocBinary.DocBinarySectionTable.PageSetupOf(stream, 0, isGutterAtTop, null);
        }
        const ushort Top = 0x9023, Bottom = 0x9024, Gutter = 0xB025;
        var plain = Setup(false, (Top, 1000), (Bottom, 1200));
        TestReport.Check(plain.top == 1000 && plain.bottom == 1200, $"doc dyaTop / dyaBottom: {plain.top}/{plain.bottom}");
        var negative = Setup(false, (Top, -1500), (Bottom, -700));
        TestReport.Check(negative.top == 1500 && negative.bottom == 700, $"doc negative dyaTop / dyaBottom: the absolute value: {negative.top}/{negative.bottom}");
        var defaults = Setup(false);
        TestReport.Check(defaults.top == 1440 && defaults.bottom == 1440, $"doc no sprms: Sep defaults: {defaults.top}/{defaults.bottom}");
        var gutterLeft = Setup(false, (Top, 1000), (Gutter, 500));
        var gutterTop = Setup(true, (Top, 1000), (Gutter, 500));
        TestReport.Check(gutterLeft.top == 1000 && gutterTop.top == 1500 && gutterTop.width == gutterLeft.width + 500,
                         $"doc gutter: left keeps top {gutterLeft.top}, at top adds it {gutterTop.top}; widths {gutterLeft.width} / {gutterTop.width}");
    }

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
        return OpenWithSidecarCheck(path, name);
    }

    /// <summary>Sections from an in-memory open, checked equal to those from a saved sidecar index reopened (not for .txt).</summary>
    private static IReadOnlyList<WordDocumentSection> OpenWithSidecarCheck(string path, string name)
    {
        static IReadOnlyList<WordDocumentSection> Read(string path, WordIndexPolicy policy)
        {
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            while (!document.IsIndexComplete) document.ScanAhead(100_000);
            return document.Sections;
        }
        IReadOnlyList<WordDocumentSection> inMemory = Read(path, Policy());
        if (Path.GetExtension(path) is ".txt") return inMemory;
        Read(path, Policy(WordIndexMode.Sidecar));                                   // builds and saves the sidecar
        IReadOnlyList<WordDocumentSection> sidecar = Read(path, Policy(WordIndexMode.Sidecar));   // reopens from it
        TestReport.Check(inMemory.SequenceEqual(sidecar), $"{Path.GetExtension(path)} {name}: the sidecar index keeps the sections (margins {DescribeMargins(sidecar)})");
        return inMemory;
    }

    /// <summary>A zipped .odt from content.xml / styles.xml bodies (mimetype stored first, as ODF requires).</summary>
    private static IReadOnlyList<WordDocumentSection> OdtZipSections(string name, string styles, string content)
    {
        string path = Path.Combine(TestReport.OutputDirectory, "page-width-" + name + ".odt");
        if (File.Exists(path)) File.Delete(path);
        using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
        {
            void Add(string entry, string text, CompressionLevel level = CompressionLevel.Optimal)
            {
                using var writer = new StreamWriter(zip.CreateEntry(entry, level).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
            Add("META-INF/manifest.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">"
                + "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>"
                + "<manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>");
            Add("styles.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles {OdfNamespaces}>{styles}</office:document-styles>");
            Add("content.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {OdfNamespaces}>{content}</office:document-content>");
        }
        return OpenWithSidecarCheck(path, name);
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
