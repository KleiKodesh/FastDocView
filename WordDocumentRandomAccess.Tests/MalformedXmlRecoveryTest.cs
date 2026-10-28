using System.Text;
using System.Xml.Linq;
using WordDocumentRandomAccess.Docx;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// WordDocumentXml.LoadUntilFirstError — what LibreOffice keeps from a malformed .docx part in repair
/// mode (libxml2 without XML_PARSE_RECOVER: every SAX event before the first fatal error, none after)
/// — one case per error class, plus the scanners' &lt;!DOCTYPE&gt; skipping.
/// </summary>
public static class MalformedXmlRecoveryTest
{
    public static void Run()
    {
        TestReport.Step("Malformed XML: content before the first error is kept (LibreOffice repair mode)");

        // mismatched end tag (math-malformed_xml.docx): "a" is delivered, the error stops everything after it
        XElement? mismatched = Load("<r><t>+</t><t>a</sPre></r><after/>");
        TestReport.Check(mismatched is not null && mismatched.Elements("t").Select(t => t.Value).SequenceEqual(["+", "a"]) && !mismatched.Descendants("after").Any(),
                         $"mismatched end tag: text before it kept, nothing after — {mismatched}");

        // duplicate attribute (tdf165348): the element with the error is not delivered (libxml2 errs in xmlParseStartTag2 before startElementNs)
        XElement? duplicate = Load("<p><pPr><before/><jc val=\"center\" val=\"center\"/></pPr><r><t>test</t></r></p>");
        TestReport.Check(duplicate is not null && duplicate.Element("pPr")?.Element("before") is not null && !duplicate.Descendants("jc").Any() && !duplicate.Descendants("r").Any(),
                         $"duplicate attribute: its element and everything after are dropped — {duplicate}");

        // undeclared entity (ExternalEntityInText.docx): never resolved; nothing after the reference
        XElement? entity = Load("<p><t>site: &quot;&testent;&quot;</t><t>after</t></p>");
        TestReport.Check(entity is not null && !entity.Value.Contains("after") && entity.Element("t") is not null,
                         $"undeclared entity: not resolved, nothing after it — {entity}");

        // unclosed elements at end of input: closed where the input ends
        XElement? unclosed = Load("<p><r><t>abc");
        TestReport.Check(unclosed?.Element("r")?.Element("t")?.Value == "abc", $"unclosed at end: elements closed, text kept — {unclosed}");

        // garbage byte (a control character is not an XML Char): content before it kept
        XElement? garbage = Load("<p><t>ok</t>\u0001<t>no</t></p>");
        TestReport.Check(garbage is not null && garbage.Elements("t").Select(t => t.Value).SequenceEqual(["ok"]), $"illegal character: content before it kept — {garbage}");

        // namespaces survive (block fragments are parsed inside the part's root tag)
        XElement? namespaced = Load("<w:document xmlns:w=\"urn:w\" xmlns=\"urn:d\"><w:p><w:t>x</w:t></w:q></w:document>");
        XNamespace w = "urn:w";
        TestReport.Check(namespaced?.Element(w + "p")?.Element(w + "t")?.Value == "x", $"namespaced elements keep their namespace — {namespaced}");

        // <!DOCTYPE> with an internal subset holding '>' / ']' in a literal and in a comment ends only at its own '>'
        const string doctype = "OCTYPE x [<!ENTITY e \"a>]b\"> <!-- ]> --> <!ENTITY f SYSTEM 'http://x/'>]>";
        var skipper = new DocxDocumentTypeDeclarationSkipper();
        int endAt = -1;
        byte[] bytes = Encoding.ASCII.GetBytes(doctype);
        for (int i = 0; i < bytes.Length; i++) if (skipper.Feed(bytes[i])) { endAt = i; break; }
        TestReport.Check(endAt == bytes.Length - 1, $"<!DOCTYPE …> internal subset skipped to its own '>' (ended at {endAt} of {bytes.Length - 1})");

        // the body scanner: a DOCTYPE before the root, and a comment holding "D", "'" and "[" between blocks
        const string body = "<?xml version=\"1.0\"?><!DOCTYPE d [<!ENTITY e SYSTEM \"http://x/\">]>"
                          + "<w:document xmlns:w=\"urn:w\"><w:body><w:p/><!-- Don't [x] Draft --><w:p/><w:p/></w:body></w:document>";
        foreach (int chunkSize in new[] { 1, 7, 4096 })
        {
            var scanner = new DocxBodyBlockScanner();
            byte[] all = Encoding.UTF8.GetBytes(body);
            for (int i = 0; i < all.Length; i += chunkSize)
            {
                int count = Math.Min(chunkSize, all.Length - i);
                scanner.Feed(all.AsSpan(i, count).ToArray(), count);
            }
            scanner.Finish();
            TestReport.Check(scanner.BlockStartOffsets.Count == 3 && scanner.BodyEndOffset > 0,
                             $"body scanner, chunks of {chunkSize}: DOCTYPE and a comment with D/'/[ skipped — {scanner.BlockStartOffsets.Count} blocks");
        }
    }

    private static XElement? Load(string xml)
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        using var reader = WordDocumentXml.CreateReader(stream, tolerateUndeclaredPrefixes: false, ignoreWhitespace: true);
        return WordDocumentXml.LoadUntilFirstError(reader);
    }
}
