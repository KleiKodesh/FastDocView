using System.Text;
using System.Xml.Linq;
using FastDocView.Docx;
using FastDocView.DocxHtml;
using FastDocView.Html;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// The lazy readers of the cold first open against the whole parse they stand in for: the .docx styles.xml slicer
/// (another prefix, a style without w:styleId, a non-UTF-8 encoding, basedOn chains, defaults, latent styles), the
/// .odt style / list-style names cut from bytes (decoded as the XML parse decodes them, not as HTML), the on-demand
/// relationship map of a Word 2003 XML file (looked up from many threads at once), and Word 2003 XML files without notes
/// or with an unclosed element before the body's end.
/// </summary>
public static class ColdOpenLazyPartsTest
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public static void Run()
    {
        TestReport.Step("Cold first open: lazy styles, names and relationships equal the whole parse");
        DocxStylesSlices();
        OdtNamesFromBytes();
        OnDemandRelationshipsFromThreads();
        Word2003Edges();
    }

    private static void DocxStylesSlices()
    {
        string Styles(string body, string declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>")
            => declaration + "<w:styles xmlns:w=\"" + W + "\" xmlns:x=\"" + W + "\">" + body + "</w:styles>";

        // a chain, defaults and latent styles: the slices resolve as the whole parse does
        string chain = Styles(
            "<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val=\"30\"/></w:rPr></w:rPrDefault></w:docDefaults>"
            + "<w:latentStyles w:defQFormat=\"0\"><w:lsdException w:name=\"heading 1\" w:uiPriority=\"9\"/><w:lsdException w:name=\"Title\"/></w:latentStyles>"
            + "<w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:pPr><w:jc w:val=\"center\"/></w:pPr></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Mid\"><w:name w:val=\"Mid\"/><w:basedOn w:val=\"Normal\"/><w:pPr><w:spacing w:after=\"240\"/></w:pPr></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Top\"><w:name w:val=\"Heading &amp; more\"/><w:basedOn w:val=\"Mid\"/><w:rPr><w:b/></w:rPr></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Loop1\"><w:basedOn w:val=\"Loop2\"/></w:style><w:style w:type=\"paragraph\" w:styleId=\"Loop2\"><w:basedOn w:val=\"Loop1\"/></w:style>"
            + "<w:style w:type=\"character\" w:default=\"1\" w:styleId=\"DefaultFont\"><w:name w:val=\"Default Paragraph Font\"/></w:style>"
            + "<w:style w:type=\"paragraph\" w:styleId=\"Mid\"><w:name w:val=\"Mid again\"/><w:basedOn w:val=\"Normal\"/></w:style>");
        (DocxStyleCascade? sliced, DocxStyleCascade whole) = Both(chain);
        TestReport.Check(sliced is not null, "styles.xml with a basedOn chain, defaults and latent styles is cut");
        if (sliced is not null)
        {
            bool same = true;
            foreach (string? id in new[] { "Normal", "Mid", "Top", "Loop1", null, "Missing" })
                same &= Describe(sliced, id) == Describe(whole, id);
            same &= sliced.StyleIdByName("heading & more") == whole.StyleIdByName("heading & more") && whole.StyleIdByName("heading & more") == "Top";
            same &= sliced.StyleIdByName("Mid again") == "Mid";
            TestReport.Check(same, "sliced cascade = whole cascade: chain, loop, defaults, repeated id, entity-encoded name");
        }

        // a style under another prefix bound to the main namespace: the whole parse reads it, so must the cascade
        string prefixed = Styles("<w:style w:type=\"paragraph\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>"
                                 + "<x:style x:type=\"paragraph\" x:styleId=\"Big\"><x:name x:val=\"Big one\"/></x:style>");
        TestReport.Check(Cascade(prefixed).StyleIdByName("Big one") == "Big", "a style written under another prefix of the main namespace is found");

        // the style id under another prefix on a w:style element
        string idPrefixed = Styles("<w:style w:type=\"paragraph\" x:styleId=\"Other\"><w:name w:val=\"Other one\"/></w:style>");
        TestReport.Check(Cascade(idPrefixed).StyleIdByName("Other one") == "Other", "a w:style whose styleId has another prefix is found");

        // a styles.xml in ISO-8859-1: a non-ASCII name decoded as declared
        string latin = Styles("<w:style w:type=\"paragraph\" w:styleId=\"Cafe\"><w:name w:val=\"Caf\u00e9\"/></w:style>", "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>");
        byte[] latinBytes = Encoding.Latin1.GetBytes(latin);
        DocxStyleCascade latinCascade = DocxStylesPartSlices.TryCut(latinBytes) is { } latinSlices ? new DocxStyleCascade(latinSlices) : new DocxStyleCascade(XDocument.Load(new MemoryStream(latinBytes)));
        TestReport.Check(latinCascade.StyleIdByName("Caf\u00e9") == "Cafe", "styles.xml declared ISO-8859-1: names read in that encoding");
    }

    private static DocxStyleCascade Cascade(string xml)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        return DocxStylesPartSlices.TryCut(bytes) is { } slices ? new DocxStyleCascade(slices) : new DocxStyleCascade(XDocument.Parse(xml));
    }

    private static (DocxStyleCascade? sliced, DocxStyleCascade whole) Both(string xml)
        => (DocxStylesPartSlices.TryCut(Encoding.UTF8.GetBytes(xml)) is { } slices ? new DocxStyleCascade(slices) : null, new DocxStyleCascade(XDocument.Parse(xml)));

    private static string Describe(DocxStyleCascade cascade, string? id)
    {
        WordParagraphFormatting p = cascade.ResolveParagraphStyle(id);
        return $"{p.Alignment}|{p.SpaceAfterTwips}|{p.RunFormatting.IsBold}|{p.RunFormatting.FontSizeHalfPoints}|{cascade.HeadingLevelOf(id)}";
    }

    private static void OdtNamesFromBytes()
    {
        const string Root = "<office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\""
                            + " xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:fo=\"urn:oasis:names:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\">";
        string xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + Root + "<office:styles>"
                     + "<style:style style:name=\"Tab\tLiteral\" style:family=\"paragraph\"><style:paragraph-properties fo:margin-left=\"1cm\"/></style:style>"
                     + "<style:style style:name=\"Tab&#9;Reference\" style:family=\"paragraph\"/>"
                     + "<style:style style:name=\"Two\r\nLines\" style:family=\"paragraph\"/>"
                     + "<style:style style:name=\"Amp&amp;&lt;x&gt;\" style:family=\"paragraph\"/>"
                     + "<text:list-style style:name=\"List\tOne\"><text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"*\"/></text:list-style>"
                     + "</office:styles></office:document-styles>";
        byte[] bytes = Encoding.UTF8.GetBytes(xml);
        var expected = XDocument.Parse(xml, LoadOptions.None).Descendants()
            .Where(element => element.Name.LocalName is "style" or "list-style")
            .Select(element => (string)element.Attribute(XName.Get("name", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"))!).ToList();
        var found = new List<string>();
        using (var view = new OdtStylesListStyleFilterStream(new MemoryStream(bytes), (name, _) => found.Add(name), (family, name, _) => found.Add(name)))
            view.CopyTo(Stream.Null);
        found.Sort(StringComparer.Ordinal);
        expected.Sort(StringComparer.Ordinal);
        TestReport.Check(found.SequenceEqual(expected, StringComparer.Ordinal),
                         $"ODT style and list-style names cut from bytes = the XML parse's ({string.Join(" | ", found.Select(Visible))} vs {string.Join(" | ", expected.Select(Visible))})");
    }

    private static string Visible(string text) => text.Replace("\t", "\\t").Replace("\r", "\\r").Replace("\n", "\\n");

    private static void OnDemandRelationshipsFromThreads()
    {
        // pictures and links are looked up outside the document's read lock (PicturePartsOf in ReadPicture): many threads at once
        var map = new DocxRelationshipMapWithFallback<string>(new Dictionary<string, string> { ["rIdListed"] = "listed" },
                                                              (string id, out string value) => (value = id.StartsWith("rIdL", StringComparison.Ordinal) ? "target:" + id : null!) is not null);
        int wrong = 0, failed = 0;
        var threads = Enumerable.Range(0, 8).Select(t => new Thread(() =>
        {
            try
            {
                for (int i = 0; i < 20000; i++)
                {
                    string id = "rIdL" + t + "-" + i;
                    if (!map.TryGetValue(id, out string? value) || value != "target:" + id) Interlocked.Increment(ref wrong);
                    if (!map.TryGetValue("rIdListed", out string? listed) || listed != "listed") Interlocked.Increment(ref wrong);
                }
            }
            catch (Exception) { Interlocked.Increment(ref failed); }
        })).ToList();
        threads.ForEach(thread => thread.Start());
        threads.ForEach(thread => thread.Join());
        TestReport.Check(wrong == 0 && failed == 0 && !map.TryGetValue("rIdOther", out _), $"on-demand relationships looked up from 8 threads at once ({wrong} wrong, {failed} threads failed)");
    }

    private static void Word2003Edges()
    {
        const string Head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><?mso-application progid=\"Word.Document\"?>"
                            + "<w:wordDocument xmlns:w=\"http://schemas.microsoft.com/office/word/2003/wordml\" xmlns:wx=\"http://schemas.microsoft.com/office/word/2003/auxHint\">";
        string folder = Path.Combine(TestReport.OutputDirectory, "word2003-edges-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            // no notes at all: the notes parts exist but are empty; no notes are shown, the body renders
            string plain = Path.Combine(folder, "no-notes.xml");
            File.WriteAllText(plain, Head + "<w:body><wx:sect><w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:hlink w:dest=\"https://example.com/a?b=1&amp;c=2\"><w:r><w:t>Link</w:t></w:r></w:hlink></w:p></wx:sect></w:body></w:wordDocument>", new UTF8Encoding(false));
            foreach ((string name, WordIndexPolicy policy) in Policies(folder))
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(plain, policy);
                while (!document.IsIndexComplete) document.ScanAhead(1000);
                string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
                TestReport.Check(document.BlockCount == 2 && html.Contains("First") && html.Contains("https://example.com/a?b=1&amp;c=2") && document.NoteCounts == (0, 0),
                                 $"Word 2003 XML without notes ({name}): 2 blocks, the link's target, no notes ({document.BlockCount} blocks, notes {document.NoteCounts})");
            }

            // an element left open before the body's end: the block stops at the body's end, the next blocks are not lost
            string unclosed = Path.Combine(folder, "unclosed.xml");
            File.WriteAllText(unclosed, Head + "<w:body><w:p><w:r><w:t>Kept</w:t></w:r></w:p><w:p><w:r><w:t>Open</w:t></w:r></w:body></w:wordDocument>", new UTF8Encoding(false));
            string outcome;
            try
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(unclosed, Policies(folder)[0].policy);
                while (!document.IsIndexComplete) document.ScanAhead(1000);
                string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
                outcome = html.Contains("Kept") && !html.Contains("wordDocument") ? "ok" : "lost text: " + html;
            }
            catch (InvalidDataException exception) { outcome = "clear error: " + exception.Message; }
            catch (Exception exception) { outcome = "crash: " + exception.GetType().Name + " " + exception.Message; }
            TestReport.Check(outcome == "ok" || outcome.StartsWith("clear error", StringComparison.Ordinal), $"Word 2003 XML with an element open at the body's end ({outcome})");
        }
        finally { try { Directory.Delete(folder, recursive: true); } catch (IOException) { } }
    }

    private static List<(string name, WordIndexPolicy policy)> Policies(string folder) =>
    [
        ("in memory", new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(folder, "cache") }),
        ("sidecar", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = Path.Combine(folder, "cache") }),
        ("sidecar reopen", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = Path.Combine(folder, "cache") }),
        ("on demand", new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = Path.Combine(folder, "cache-demand"), IndexesOnDemand = true }),
    ];
}
