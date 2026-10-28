using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Writes an OpenDocument text with a known structure, as a package (.odt: mimetype stored first,
/// manifest, styles.xml with notes configuration, content.xml) or flat (.fodt, the same parts in one
/// office:document). A fixed prefix exercises lists (nested, display-levels, start-value, list-header,
/// continue-numbering, continue-list, bullets), a table with spans and a footnote in a cell, footnotes
/// and endnotes; then <c>count</c> repeating paragraphs ("פסקה N", RTL Hebrew), every 50th a heading
/// ("סעיף N"), every 10th (i % 10 == 5) with a footnote, every 200th (i % 200 == 199) a table.
/// </summary>
public static class SyntheticOdtGenerator
{
    public const int PrefixBlockCount = 13;
    private const string Paragraph = "פסקה";       // "paragraph"
    private const string Chapter = "פרק";              // "chapter"
    private const string Section = "סעיף";        // "section"

    /// <summary>The index-time list label expected per prefix block (null = none).</summary>
    public static readonly Dictionary<int, string?> ExpectedPrefixLabels = new()
    {
        [0] = null, [1] = null, [2] = "1.", [3] = "2.", [4] = "2.a)", [5] = "7.", [6] = null, [7] = null,
        [8] = "8.", [9] = "1.", [10] = "9.", [11] = null, [12] = "•",
    };

    public static int ExpectedBlockCount(int count) => PrefixBlockCount + count;
    public static int ExpectedHeadings(int count) => 1 + Enumerable.Range(0, count).Count(i => i % 50 == 0 && i % 200 != 199);
    public static int ExpectedFootnotes(int count) => 2 + Enumerable.Range(0, count).Count(i => i % 10 == 5 && i % 200 != 199);
    public const int ExpectedEndnotes = 1;

    private const string Namespaces =
        "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" " +
        "xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" " +
        "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" " +
        "xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" office:version=\"1.3\"";

    private const string FontFaces = """<office:font-face-decls><style:font-face style:name="David CLM" svg:font-family="'David CLM'"/></office:font-face-decls>""";

    private const string NamedStyles = """
        <office:styles>
          <style:default-style style:family="paragraph"><style:paragraph-properties fo:text-align="start" style:writing-mode="rl-tb"/><style:text-properties style:font-name="David CLM" style:font-name-complex="David CLM" fo:font-size="12pt" style:font-size-complex="13pt"/></style:default-style>
          <style:style style:name="Standard" style:family="paragraph"/>
          <style:style style:name="Heading_20_1" style:display-name="Heading 1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-top="0.1665in" fo:margin-bottom="0.0835in"/><style:text-properties fo:font-size="16pt" fo:font-weight="bold" style:font-weight-complex="bold" fo:color="#2f5496"/></style:style>
          <style:style style:name="Emphasis" style:family="text"><style:text-properties fo:font-style="italic" style:font-style-complex="italic"/></style:style>
          <text:list-style style:name="L1">
            <text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"/>
            <text:list-level-style-number text:level="2" style:num-suffix=")" style:num-format="a" text:display-levels="2"/>
          </text:list-style>
          <text:list-style style:name="L2"><text:list-level-style-bullet text:level="1" text:bullet-char="•"/></text:list-style>
          <text:notes-configuration text:note-class="footnote" style:num-format="1" text:start-value="0" text:footnotes-position="page" text:start-numbering-at="document"/>
          <text:notes-configuration text:note-class="endnote" style:num-format="i" text:start-value="0"/>
        </office:styles>
        """;

    private const string AutomaticStyles = """
        <office:automatic-styles>
          <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:text-align="justify" fo:margin-bottom="0.0835in"/></style:style>
          <style:style style:name="T1" style:family="text" style:parent-style-name="Emphasis"><style:text-properties style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"/></style:style>
          <style:style style:name="Table1" style:family="table"><style:table-properties style:width="6in" table:align="center" style:writing-mode="rl-tb"/></style:style>
          <style:style style:name="Table1.A" style:family="table-column"><style:table-column-properties style:column-width="2in"/></style:style>
          <style:style style:name="Table1.A1" style:family="table-cell"><style:table-cell-properties fo:padding="0.04in" fo:border="0.5pt solid #000000" fo:background-color="#eeeeee" style:vertical-align="middle"/></style:style>
        </office:automatic-styles>
        """;

    public static void WritePackage(string path, int count)
    {
        using var file = File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);
        // ODF 1.3 part 2 §3.3: "mimetype" is the first file of the zip, stored, no extra field
        using (var writer = new StreamWriter(zip.CreateEntry("mimetype", CompressionLevel.NoCompression).Open(), new UTF8Encoding(false)))
            writer.Write("application/vnd.oasis.opendocument.text");
        using (var writer = new StreamWriter(zip.CreateEntry("META-INF/manifest.xml").Open(), new UTF8Encoding(false)))
            writer.Write("""<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>""");
        using (var writer = new StreamWriter(zip.CreateEntry("styles.xml").Open(), new UTF8Encoding(false)))
            writer.Write($"""<?xml version="1.0" encoding="UTF-8"?><office:document-styles {Namespaces}>{FontFaces}{NamedStyles}<office:master-styles/></office:document-styles>""");
        using (var writer = new StreamWriter(zip.CreateEntry("content.xml").Open(), new UTF8Encoding(false), 1 << 16))
        {
            writer.Write($"""<?xml version="1.0" encoding="UTF-8"?><office:document-content {Namespaces}><office:scripts/>{FontFaces}{AutomaticStyles}<office:body><office:text>""");
            WriteBody(writer, count);
            writer.Write("</office:text></office:body></office:document-content>");
        }
    }

    public static void WriteFlat(string path, int count)
    {
        using var writer = new StreamWriter(path, false, new UTF8Encoding(false), 1 << 16);
        writer.Write($"""<?xml version="1.0" encoding="UTF-8"?><office:document {Namespaces} office:mimetype="application/vnd.oasis.opendocument.text"><office:meta/>{FontFaces}{NamedStyles}{AutomaticStyles}<office:master-styles/><office:body><office:text>""");
        WriteBody(writer, count);
        writer.Write("</office:text></office:body></office:document>");
    }

    private static int _footnote, _endnote;

    private static string Footnote() => $"<text:note text:id=\"ftn{++_footnote}\" text:note-class=\"footnote\"><text:note-citation>{_footnote}</text:note-citation><text:note-body><text:p text:style-name=\"Standard\">footnote body {_footnote}</text:p></text:note-body></text:note>";

    private static void WriteBody(TextWriter writer, int count)
    {
        _footnote = 0;
        _endnote = 0;
        writer.Write("<text:sequence-decls/><!-- a comment with don't and > inside --><?processing instruction?>");
        // 0-1: heading, paragraph with a footnote and an endnote, a span with a text style
        writer.Write($"<text:h text:style-name=\"Heading_20_1\" text:outline-level=\"1\">{Chapter} 1</text:h>");
        writer.Write($"<text:p text:style-name=\"P1\">{Paragraph} <text:span text:style-name=\"T1\">styled</text:span>{Footnote()} and<text:s text:c=\"2\"/>end<text:note text:id=\"edn1\" text:note-class=\"endnote\"><text:note-citation>i</text:note-citation><text:note-body><text:p>endnote body 1</text:p></text:note-body></text:note></text:p>");
        // 2-6: list1 — items, nested list (display-levels 2), start-value, list-header
        writer.Write("<text:list xml:id=\"list1\" text:style-name=\"L1\">");
        writer.Write("<text:list-item><text:p>item a</text:p></text:list-item>");
        writer.Write("<text:list-item><text:p>item b</text:p><text:list><text:list-item><text:p>sub item</text:p></text:list-item></text:list></text:list-item>");
        writer.Write("<text:soft-page-break/>");
        writer.Write("<text:list-item text:start-value=\"7\"><text:p>item seven</text:p></text:list-item>");
        writer.Write("<text:list-header><text:p>header paragraph</text:p></text:list-header>");
        writer.Write("</text:list>");
        // 7: plain paragraph inside a section
        writer.Write("<text:section text:name=\"Section1\"><text:p>between lists</text:p></text:section>");
        // 8: continue-numbering continues list1 (same style, last processed list)
        writer.Write("<text:list text:style-name=\"L1\" text:continue-numbering=\"true\"><text:list-item><text:p>continued eight</text:p></text:list-item></text:list>");
        // 9: a new list restarts
        writer.Write("<text:list xml:id=\"list3\" text:style-name=\"L1\"><text:list-item><text:p>new list one</text:p></text:list-item></text:list>");
        // 10: continue-list="list1" continues list1 by id
        writer.Write("<text:list text:style-name=\"L1\" text:continue-list=\"list1\"><text:list-item><text:p>continued nine</text:p></text:list-item></text:list>");
        // 11: table with a column span, a covered cell, cell style and a footnote in a cell
        writer.Write("<table:table table:name=\"Table1\" table:style-name=\"Table1\"><table:table-column table:style-name=\"Table1.A\" table:number-columns-repeated=\"3\"/>");
        writer.Write("<table:table-row><table:table-cell table:style-name=\"Table1.A1\" table:number-columns-spanned=\"2\"><text:p>spanned</text:p></table:table-cell><table:covered-table-cell/><table:table-cell><text:p>cell c" + Footnote() + "</text:p></table:table-cell></table:table-row>");
        writer.Write("<table:table-row><table:table-cell><text:p>r2a</text:p></table:table-cell><table:table-cell><text:p>r2b</text:p></table:table-cell><table:table-cell><text:p>r2c</text:p></table:table-cell></table:table-row></table:table>");
        // 12: bullet list
        writer.Write("<text:list text:style-name=\"L2\"><text:list-item><text:p>bullet</text:p></text:list-item></text:list>");
        writer.Write("<text:tracked-changes><text:changed-region text:id=\"c1\"><text:deletion><text:p>deleted text never indexed</text:p></text:deletion></text:changed-region></text:tracked-changes>");

        for (int i = 0; i < count; i++)
        {
            int blockIndex = PrefixBlockCount + i;
            if (i % 200 == 199)
                writer.Write($"<table:table table:name=\"T{i}\"><table:table-column/><table:table-row><table:table-cell><text:p>{Paragraph} {blockIndex} table</text:p></table:table-cell></table:table-row></table:table>");
            else if (i % 50 == 0)
                writer.Write($"<text:h text:style-name=\"Heading_20_1\" text:outline-level=\"2\">{Section} {blockIndex}</text:h>");
            else if (i % 10 == 5)
                writer.Write($"<text:p text:style-name=\"P1\">{Paragraph} {blockIndex}{Footnote()} text text text</text:p>");
            else
                writer.Write($"<text:p text:style-name=\"P1\">{Paragraph} {blockIndex} <text:span text:style-name=\"T1\">some styled words</text:span> and more text to give the paragraph a realistic length.</text:p>");
        }
    }
}
