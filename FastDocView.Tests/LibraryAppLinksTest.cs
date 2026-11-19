using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Links into a library app (Otzaria, Zayit, KitveiHakodesh) are marked data-app-link in every reader, so the host opens
/// them inside itself and never hands them to a browser: the shapes these apps write (Otzaria's
/// lib/core/external_uri_router.dart and lib/utils/text/html_link_handler.dart: otzaria://open/book|pdf/, inline-link,
/// book://title#heading; Zayit's core/deeplink/ZayitDeepLink.kt and its seforim://line/ citations; this app's
/// kitveihakodeshapp:// and seforimapp://), in upper case too, and a Hebrew title that is not a valid URI host.
/// Formats: .docx (w:hyperlink and a HYPERLINK field), .odt and .fodt (text:a), .rtf (HYPERLINK field), .txt (an a tag
/// in the text) and .doc (HYPERLINK fields: test-data/doc-links/doc-app-links-v1.doc, which Word (invisible, SaveAs2
/// as Word 97-2003) wrote from this test's .docx; bump the version when the link list changes).
/// </summary>
public static class LibraryAppLinksTest
{
    /// <summary>A Hebrew book title (not a valid URI host), written as code points.</summary>
    private static readonly string Title = new(new[] { (char)0x05d1, (char)0x05e8, (char)0x05d0, (char)0x05e9, (char)0x05d9, (char)0x05ea });

    private static readonly string[] Links =
    [
        "otzaria://open/book/12",
        "zayit://book/4/line/99",
        "OTZARIA://open/book/12?index=3&mark",
        "otzaria://open/pdf/120?index=17",
        "seforim://line/5",
        "kitveihakodeshapp://book/7",
        "seforimapp://book/7/?index=2",
        "otzaria://inline-link?path=books%2Fx.txt&index=4",
        "book://" + Title + "#" + Title,
    ];

    public static void Run()
    {
        TestReport.Step("Library app links (Otzaria, Zayit, KitveiHakodesh) are marked in every format");
        Check("docx", WriteDocx());
        Check("odt", WriteOdt());
        Check("fodt", WriteFlatOdt());
        Check("rtf", WriteRtf());
        Check("txt", WriteText());
        string doc = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "doc-links", "doc-app-links-v1.doc"));
        TestReport.Check(File.Exists(doc), "test-data/doc-links/doc-app-links-v1.doc present");
        if (File.Exists(doc)) Check("doc", doc);
    }

    private static void Check(string format, string path)
    {
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        var html = new StringBuilder();
        foreach (var block in document.RenderBlocks(0, document.BlockCount)) html.Append(block.Html).Append('\n');
        string all = html.ToString();
        for (int i = 0; i < Links.Length; i++)
        {
            int at = all.IndexOf("applink" + i + "<", StringComparison.Ordinal);
            int open = at < 0 ? -1 : all.LastIndexOf("<a ", at, StringComparison.Ordinal);
            string tag = open < 0 ? "" : all.Substring(open, all.IndexOf('>', open) - open + 1);
            string scheme = Links[i].Substring(0, Links[i].IndexOf(':')).ToLowerInvariant();
            TestReport.Check(tag.Contains("data-app-link=\"1\"") && !tag.Contains("target=\"_blank\"") && tag.ToLowerInvariant().Contains("href=\"" + scheme + ":"),
                             $"{format}: the {scheme}: link #{i} is marked as an app link ({Ascii(tag)})");
        }
        TestReport.Check(all.Contains("href=\"https://example.org/x\"") && all.Contains("target=\"_blank\""), $"{format}: a web link still opens outside");
    }

    private static string Ascii(string text)
    {
        var result = new StringBuilder(text.Length);
        foreach (char ch in text) result.Append(ch < 128 ? ch.ToString() : $"\\u{(int)ch:x4}");
        return result.ToString();
    }

    private static string Xml(string text) => System.Security.SecurityElement.Escape(text)!;

    private static string WriteDocx()
    {
        var body = new StringBuilder();
        var relationships = new StringBuilder();
        for (int i = 0; i < Links.Length; i++)
        {
            // even ones as w:hyperlink (a relationship), odd ones as a HYPERLINK field: both paths of the .docx reader
            if (i % 2 == 0)
            {
                relationships.Append($"<Relationship Id=\"rIdL{i}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"{Xml(Links[i])}\" TargetMode=\"External\"/>");
                body.Append($"<w:p><w:hyperlink r:id=\"rIdL{i}\"><w:r><w:t>applink{i}</w:t></w:r></w:hyperlink></w:p>");
            }
            else
                body.Append($"<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"{Xml(Links[i])}\" </w:instrText></w:r>"
                          + $"<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>applink{i}</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>");
        }
        relationships.Append("<Relationship Id=\"rIdWeb\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.org/x\" TargetMode=\"External\"/>");
        body.Append("<w:p><w:hyperlink r:id=\"rIdWeb\"><w:r><w:t>web</w:t></w:r></w:hyperlink></w:p>");
        string path = Path.Combine(TestReport.OutputDirectory, "app-links-v1.docx");
        WriteZip(path, [
            ("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>"),
            ("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"),
            ("word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + relationships + "</Relationships>"),
            ("word/document.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>" + body + "<w:sectPr/></w:body></w:document>"),
        ]);
        return path;
    }

    /// <summary>A zip of text entries (stored when the name is "mimetype", as ODF requires), replacing any earlier file.</summary>
    private static void WriteZip(string path, (string name, string text)[] entries)
    {
        using var stream = new FileStream(path, FileMode.Create);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create);
        foreach ((string name, string text) in entries)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name, name == "mimetype" ? CompressionLevel.NoCompression : CompressionLevel.Optimal).Open(), new UTF8Encoding(false));
            writer.Write(text);
        }
    }

    private static string OdtBody()
    {
        var body = new StringBuilder();
        for (int i = 0; i < Links.Length; i++) body.Append($"<text:p><text:a xlink:type=\"simple\" xlink:href=\"{Xml(Links[i])}\">applink{i}</text:a></text:p>");
        return body.Append("<text:p><text:a xlink:type=\"simple\" xlink:href=\"https://example.org/x\">web</text:a></text:p>").ToString();
    }

    private const string OdfNamespaces = "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" "
                                       + "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" office:version=\"1.3\"";

    internal static string WriteOdt()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "app-links-v1.odt");
        WriteZip(path, [
            ("mimetype", "application/vnd.oasis.opendocument.text"),
            ("content.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-content {OdfNamespaces}><office:body><office:text>{OdtBody()}</office:text></office:body></office:document-content>"),
            ("META-INF/manifest.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">"
                + "<manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"),
        ]);
        return path;
    }

    private static string WriteFlatOdt()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "app-links-v1.fodt");
        File.WriteAllText(path, $"<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document {OdfNamespaces} office:mimetype=\"application/vnd.oasis.opendocument.text\"><office:body><office:text>{OdtBody()}</office:text></office:body></office:document>",
                          new UTF8Encoding(false));
        return path;
    }

    private static string WriteRtf()
    {
        // RTF is 8-bit: the Hebrew title goes percent-encoded (UTF-8 bytes)
        var rtf = new StringBuilder(@"{\rtf1\ansi{\fonttbl{\f0 Arial;}}" + "\n");
        for (int i = 0; i < Links.Length; i++)
        {
            var address = new StringBuilder();
            foreach (byte b in Encoding.UTF8.GetBytes(Links[i])) address.Append(b < 128 ? ((char)b).ToString() : $"%{b:X2}");
            rtf.Append(@"\pard {\field{\*\fldinst HYPERLINK """).Append(address).Append(@""" }{\fldrslt applink").Append(i).Append(@"}}\par").Append('\n');
        }
        rtf.Append(@"\pard {\field{\*\fldinst HYPERLINK ""https://example.org/x"" }{\fldrslt web}}\par" + "\n}");
        string path = Path.Combine(TestReport.OutputDirectory, "app-links-v1.rtf");
        File.WriteAllText(path, rtf.ToString(), Encoding.ASCII);
        return path;
    }

    private static string WriteText()
    {
        var text = new StringBuilder();
        for (int i = 0; i < Links.Length; i++) text.Append($"<a href=\"{Links[i].Replace("&", "&amp;")}\">applink{i}</a>\r\n");
        text.Append("<a href=\"https://example.org/x\">web</a>\r\n");
        string path = Path.Combine(TestReport.OutputDirectory, "app-links-v1.txt");
        File.WriteAllText(path, text.ToString(), new UTF8Encoding(true));
        return path;
    }
}
