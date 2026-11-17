using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using FastDocView;
using FastDocView.Odt;

namespace FastDocView.Tests;

/// <summary>
/// OpenDocument fields as LibreOffice prints them. The fixture saves a wrong text ("XX") in every field; LibreOffice 25.8
/// (soffice --convert-to pdf, 2026-11-17) printed the values below - what it recomputes, and what it keeps (author, sender,
/// input, statistics). Each paragraph is "key [field]"; the check reads the text between the brackets. Also: an on-demand first
/// open whose early references point at late targets shows nothing and is provisional until the index is complete.
/// </summary>
public static class OdtFieldsFidelityTest
{
    private const string Head = """
        <?xml version="1.0" encoding="UTF-8"?>
        <office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
        <office:meta><dc:title>RealTitle</dc:title><dc:subject>RealSubject</dc:subject><meta:initial-creator>RealCreator</meta:initial-creator><dc:creator>RealModifier</dc:creator><meta:creation-date>2020-03-04T05:06:07</meta:creation-date><meta:user-defined meta:name="Prop1">RealProp</meta:user-defined><meta:keyword>KeyA</meta:keyword><dc:description>RealComment</dc:description></office:meta>
        <office:automatic-styles><number:date-style style:name="D1" number:language="en" number:country="US"><number:day-of-week number:style="long"/><number:text>, </number:text><number:day/><number:text> </number:text><number:month number:textual="true" number:style="long"/><number:text> </number:text><number:year number:style="long"/></number:date-style></office:automatic-styles>
        <office:body><office:text>
        <text:variable-decls><text:variable-decl office:value-type="float" text:name="V"/><text:variable-decl office:value-type="string" text:name="S"/></text:variable-decls>
        <text:sequence-decls><text:sequence-decl text:display-outline-level="0" text:name="Illustration"/><text:sequence-decl text:display-outline-level="1" text:name="Table"/></text:sequence-decls>
        <text:user-field-decls><text:user-field-decl office:value-type="float" office:value="7" text:name="U"/></text:user-field-decls>
        <text:h text:outline-level="1">ChapOne</text:h>

        """;

    /// <summary>(key, field markup, what LibreOffice printed; null = checked apart).</summary>
    private static readonly (string Key, string Field, string? Printed)[] Rows =
    [
        ("seq1", "<text:sequence text:ref-name=\"refIllustration0\" text:name=\"Illustration\" text:formula=\"ooow:Illustration+1\" style:num-format=\"1\">XX</text:sequence>", "1"),
        ("seq2", "<text:sequence text:ref-name=\"refIllustration1\" text:name=\"Illustration\" text:formula=\"ooow:Illustration+1\" style:num-format=\"1\">XX</text:sequence>", "2"),
        ("seqref-value", "<text:sequence-ref text:reference-format=\"value\" text:ref-name=\"refIllustration1\">XX</text:sequence-ref>", "2"),
        ("seqref-catval", "<text:sequence-ref text:reference-format=\"category-and-value\" text:ref-name=\"refIllustration1\">XX</text:sequence-ref>", "2"),
        ("seqref-page", "<text:sequence-ref text:reference-format=\"page\" text:ref-name=\"refIllustration1\">XX</text:sequence-ref>", null),
        ("bmref-text", "<text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"bm\">XX</text:bookmark-ref>", "BMTEXT"),
        ("bmref-chapter", "<text:bookmark-ref text:reference-format=\"chapter\" text:ref-name=\"bm\">XX</text:bookmark-ref>", ""),
        ("bmref-direction", "<text:bookmark-ref text:reference-format=\"direction\" text:ref-name=\"bm\">XX</text:bookmark-ref>", "below"),
        ("refref-text", "<text:reference-ref text:reference-format=\"text\" text:ref-name=\"rm\">XX</text:reference-ref>", "RMTEXT"),
        ("chapter-name", "<text:chapter text:display=\"name\" text:outline-level=\"1\">XX</text:chapter>", "ChapOne"),
        ("varset", "<text:variable-set text:name=\"V\" text:formula=\"ooow:3+4\" office:value-type=\"float\" office:value=\"1\">XX</text:variable-set>", "7"),
        ("varget", "<text:variable-get text:name=\"V\">XX</text:variable-get>", "7"),
        ("userget", "<text:user-field-get text:name=\"U\">XX</text:user-field-get>", "7"),
        ("expr", "<text:expression text:formula=\"ooow:2*V\" office:value-type=\"float\" office:value=\"1\">XX</text:expression>", "14"),
        ("date-nofix", "<text:date text:date-value=\"2001-01-01\">XX</text:date>", null),
        ("date-fixed", "<text:date text:fixed=\"true\" text:date-value=\"2001-01-01\">XX</text:date>", "1/1/01"),
        ("title", "<text:title>XX</text:title>", "RealTitle"),
        ("subject", "<text:subject>XX</text:subject>", "RealSubject"),
        ("author-nofix", "<text:author-name>XX</text:author-name>", "XX"),
        ("author-fixed", "<text:author-name text:fixed=\"true\">XX</text:author-name>", "XX"),
        ("initial-creator", "<text:initial-creator>XX</text:initial-creator>", "RealCreator"),
        ("creation-date", "<text:creation-date>XX</text:creation-date>", "03/04/2020"),
        ("user-defined", "<text:user-defined text:name=\"Prop1\">XX</text:user-defined>", "RealProp"),
        ("page-number", "<text:page-number text:select-page=\"current\">XX</text:page-number>", null),
        ("cond", "<text:conditional-text text:condition=\"ooow:1==1\" text:string-value-if-true=\"T\" text:string-value-if-false=\"F\" text:current-value=\"false\">XX</text:conditional-text>", "T"),
        ("textinput", "<text:text-input>XX</text:text-input>", "XX"),
        ("sender", "<text:sender-firstname>XX</text:sender-firstname>", "XX"),
        ("noteref", "<text:note-ref text:note-class=\"footnote\" text:reference-format=\"text\" text:ref-name=\"ftn0\">XX</text:note-ref>", "1"),
        ("chapter-number", "<text:chapter text:display=\"number\" text:outline-level=\"1\">XX</text:chapter>", ""),
        ("modifier-date", "<text:modification-date>XX</text:modification-date>", ""),
        ("time-nofix", "<text:time text:time-value=\"01:02:03\">XX</text:time>", null),
        ("date-styled", "<text:date text:fixed=\"true\" style:data-style-name=\"D1\" text:date-value=\"2001-02-03\">XX</text:date>", "Saturday, 3 February 2001"),
        ("strset", "<text:variable-set text:name=\"S\" office:value-type=\"string\">abc</text:variable-set>", "abc"),
        ("strget", "<text:variable-get text:name=\"S\">XX</text:variable-get>", "abc"),
        ("hidden-true", "<text:hidden-text text:condition=\"ooow:1\" text:string-value=\"HT\" text:is-hidden=\"true\"/>", ""),
        ("hidden-false", "<text:hidden-text text:condition=\"ooow:0\" text:string-value=\"HF\" text:is-hidden=\"false\"/>", "HF"),
        ("table1", "<text:sequence text:ref-name=\"refTable0\" text:name=\"Table\" text:formula=\"ooow:Table+1\" style:num-format=\"A\">XX</text:sequence>", "A"),
        ("keywords", "<text:keywords>XX</text:keywords>", "KeyA"),
        ("description", "<text:description>XX</text:description>", "RealComment"),
        ("seqref-text", "<text:sequence-ref text:reference-format=\"text\" text:ref-name=\"refIllustration9\">XX</text:sequence-ref>", "Illustration 3: Cap text"),
        ("seqref-catval2", "<text:sequence-ref text:reference-format=\"category-and-value\" text:ref-name=\"refIllustration9\">XX</text:sequence-ref>", "Illustration 3"),
        ("seqref-caption", "<text:sequence-ref text:reference-format=\"caption\" text:ref-name=\"refIllustration9\">XX</text:sequence-ref>", "Cap text"),
        ("seqref-direction", "<text:sequence-ref text:reference-format=\"direction\" text:ref-name=\"refIllustration0\">XX</text:sequence-ref>", "above"),
        ("missing", "<text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"nowhere\">XX</text:bookmark-ref>", OdtFieldComputer.ReferenceNotFound),
    ];

    private const string Tail = """
        <text:p>Illustration <text:sequence text:ref-name="refIllustration9" text:name="Illustration" text:formula="ooow:Illustration+1" style:num-format="1">XX</text:sequence>: Cap text</text:p>
        <text:h text:outline-level="1">ChapTwo</text:h>
        <text:p>table2 [<text:sequence text:ref-name="refTable1" text:name="Table" text:formula="ooow:Table+1" style:num-format="A">XX</text:sequence>]</text:p>
        <text:p><text:bookmark-start text:name="bm"/>BMTEXT<text:bookmark-end text:name="bm"/> and <text:reference-mark-start text:name="rm"/>RMTEXT<text:reference-mark-end text:name="rm"/> note<text:note text:id="ftn0" text:note-class="footnote"><text:note-citation>9</text:note-citation><text:note-body><text:p>n</text:p></text:note-body></text:note></text:p>
        </office:text></office:body></office:document>
        """;

    public static string FixturePath => Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-fields.fodt");

    public static string WriteFixture()
    {
        var document = new StringBuilder(Head);
        foreach ((string key, string field, _) in Rows) document.Append("<text:p>").Append(key).Append(" [").Append(field).Append("]</text:p>\n");
        document.Append(Tail);
        string text = document.ToString();
        if (!File.Exists(FixturePath) || File.ReadAllText(FixturePath) != text) File.WriteAllText(FixturePath, text, new UTF8Encoding(false));
        return FixturePath;
    }

    public static void Run()
    {
        TestReport.Step("OpenDocument fields as LibreOffice prints them (sequences, variables, references, chapter, dates, document information)");
        string path = WriteFixture();
        List<string> html = OdtDrawingFidelityTest.Render(path);
        var shown = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (string block in html)
        {
            string text = WebUtility.HtmlDecode(Regex.Replace(block, "<[^>]+>", ""));
            Match match = Regex.Match(text, @"^\s*(\S+) \[(.*)\]\s*$", RegexOptions.Singleline);
            if (match.Success) shown[match.Groups[1].Value] = match.Groups[2].Value;
        }
        foreach ((string key, _, string? printed) in Rows)
        {
            if (printed is null) continue;
            string got = shown.TryGetValue(key, out string? value) ? value : "(missing)";
            TestReport.Check(got == printed, $"{key}: '{Ascii(got)}' (LibreOffice printed '{Ascii(printed)}')");
        }
        foreach (string key in new[] { "seqref-page", "page-number" })
            TestReport.Check(shown.TryGetValue(key, out string? saved) && saved == "XX", $"{key}: page-bound, the saved text until the paged view ('{Ascii(saved ?? "")}')");
        TestReport.Check(shown.TryGetValue("table2", out string? table2) && table2 == "A", $"table2: a sequence numbered per chapter restarts at the next heading ('{Ascii(table2 ?? "")}')");
        string today = DateTime.Now.ToString("M/d/yy", System.Globalization.CultureInfo.InvariantCulture);
        TestReport.Check(shown.TryGetValue("date-nofix", out string? date) && date == today, $"date-nofix: today '{Ascii(date ?? "")}'");
        TestReport.Check(shown.TryGetValue("time-nofix", out string? time) && Regex.IsMatch(time, @"^\d\d:\d\d:\d\d$") && time != "01:02:03", $"time-nofix: now '{Ascii(time ?? "")}'");

        TestReport.Step("OpenDocument fields on demand: early references to late targets are provisional until the index is complete");
        RunLateTargets(Path.Combine(TestReport.OutputDirectory, "synthetic-odt-late-field-targets-v1.fodt"));
    }

    private static void RunLateTargets(string path)
    {
        if (!File.Exists(path))
        {
            var body = new StringBuilder(Head);
            body.Append("<text:p>Early <text:bookmark-ref text:reference-format=\"text\" text:ref-name=\"late\">old</text:bookmark-ref> and <text:sequence-ref text:reference-format=\"value\" text:ref-name=\"lateSeq\">old</text:sequence-ref> end.</text:p>\n");
            for (int i = 0; i < 20_000; i++) body.Append("<text:p>Filler paragraph ").Append(i).Append(".</text:p>\n");
            body.Append("<text:p><text:bookmark-start text:name=\"late\"/>Target words<text:bookmark-end text:name=\"late\"/> <text:sequence text:ref-name=\"lateSeq\" text:name=\"Illustration\" text:formula=\"ooow:Illustration+1\">9</text:sequence></text:p>\n");
            body.Append("</office:text></office:body></office:document>");
            File.WriteAllText(path, body.ToString(), new UTF8Encoding(false));
        }
        string cache = Path.Combine(TestReport.OutputDirectory, "index-cache-odt-fields");
        var onDemand = new WordIndexPolicy { CacheDirectory = cache, ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        if (File.Exists(onDemand.IndexPathFor(path))) File.Delete(onDemand.IndexPathFor(path));
        using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, onDemand))
        {
            WordDocumentBlockHtml early = document.RenderBlocks(1, 1)[0];
            TestReport.Check(!document.IsIndexComplete && early.IsProvisional && !early.Html.Contains(OdtFieldComputer.ReferenceNotFound) && !early.Html.Contains("Target words") && !early.Html.Contains("old"),
                             $"before completion: provisional, nothing shown (complete={document.IsIndexComplete}, provisional={early.IsProvisional})");
            TestReport.Check(!document.RenderBlocks(2, 1)[0].IsProvisional, "a block without such fields is not provisional");
            while (!document.IsIndexComplete) document.ScanAhead(5_000);
            WordDocumentBlockHtml final = document.RenderBlocks(1, 1)[0];
            TestReport.Check(!final.IsProvisional && final.Html.Contains("Target words") && final.Html.Contains(" and 1 end"), $"after completion: the bookmark text and the sequence number, final (provisional={final.IsProvisional})");
        }
        // a reopen from the saved index computes the same (field values travel in the span anchors' scan state)
        using (IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(path, onDemand))
        {
            TestReport.Check(reopened.IsIndexComplete, "reopen: the saved index is used");
            string late = reopened.RenderBlocks(reopened.BlockCount - 1, 1)[0].Html;
            TestReport.Check(late.Contains(">1<") || Regex.Replace(late, "<[^>]+>", "").EndsWith(" 1"), $"reopen: the late sequence is 1 ({Ascii(Regex.Replace(late, "<[^>]+>", ""))})");
            TestReport.Check(reopened.RenderBlocks(1, 1)[0].Html.Contains("Target words"), "reopen: the early reference resolves at once");
        }
    }

    private static string Ascii(string text)
    {
        var builder = new StringBuilder();
        foreach (char c in text) builder.Append(c < 128 ? c.ToString() : $"\\u{(int)c:X4}");
        return builder.ToString();
    }
}
