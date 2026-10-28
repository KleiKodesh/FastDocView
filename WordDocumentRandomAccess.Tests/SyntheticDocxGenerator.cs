using System.IO.Compression;
using System.Text;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Writes a .docx with a known structure so index results can be checked exactly: every paragraph
/// carries its ordinal ("פסקה N"), every 50th is a Heading1, every 200th block is a 2x2 table.
/// Streams the body straight into the zip entry, so a multi-GB document costs no memory.
/// </summary>
public static class SyntheticDocxGenerator
{
    private const string StylesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:cs="David"/><w:sz w:val="22"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault>
            <w:pPrDefault><w:pPr><w:spacing w:after="120"/><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>
          <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="both"/></w:pPr></w:style>
          <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="32"/><w:color w:val="2F5496"/></w:rPr></w:style>
          <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>
        </w:styles>
        """;

    private const string NumberingXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:abstractNum w:abstractNumId="0">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2"/></w:lvl>
          </w:abstractNum>
          <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="hebrew1"/><w:lvlText w:val="%1."/></w:lvl>
          </w:abstractNum>
          <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
          <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
        </w:numbering>
        """;

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
          <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
          <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
        </Types>
        """;

    private const string RootRelsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
        </Relationships>
        """;

    private const string DocumentRelsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
          <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
        </Relationships>
        """;

    public static int ExpectedBlockCount(int blockCount) => blockCount + 1;   // + the trailing w:sectPr

    /// <param name="noteFillerCharacters">Extra text per note — makes the notes parts as large as a test needs (streamed, never built in memory).</param>
    public static void Write(string path, int blockCount, int noteFillerCharacters = 0)
    {
        using var file = File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);
        AddText(zip, "[Content_Types].xml", ContentTypesXml);
        AddText(zip, "_rels/.rels", RootRelsXml);
        AddText(zip, "word/_rels/document.xml.rels", DocumentRelsXml);
        AddText(zip, "word/styles.xml", StylesXml);
        AddText(zip, "word/numbering.xml", NumberingXml);
        WriteNotes(zip, "word/footnotes.xml", "footnotes", "footnote", blockCount, FootnoteEvery, "הערה", noteFillerCharacters);
        WriteNotes(zip, "word/endnotes.xml", "endnotes", "endnote", blockCount, EndnoteEvery, "הערת סיום", noteFillerCharacters);

        ZipArchiveEntry document = zip.CreateEntry("word/document.xml", CompressionLevel.Optimal);
        using var writer = new StreamWriter(document.Open(), new UTF8Encoding(false), 1 << 16);
        writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
        writer.Write("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><w:body>");
        var random = new Random(12345);
        for (int i = 0; i < blockCount; i++)
        {
            if (i % 200 == 199) { WriteTable(writer, i); continue; }
            if (i % 50 == 0) writer.Write($"<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr><w:r><w:t>כותרת {i / 50 + 1} — פסקה {i}</w:t></w:r></w:p>");
            else if (i % 10 == 3) writer.Write($"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>פסקה {i} פריט רשימה</w:t></w:r></w:p>");
            else if (i % 10 == 4) writer.Write($"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr><w:r><w:t>פסקה {i} תת-פריט</w:t></w:r></w:p>");
            else if (i % 10 == 7) writer.Write($"<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr></w:pPr><w:r><w:t>פסקה {i} סעיף עברי</w:t></w:r></w:p>");
            else if (i % FootnoteEvery == 1) writer.Write($"<w:p><w:r><w:t xml:space=\"preserve\">פסקה {i} עם הערה</w:t></w:r><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"{NoteId(i)}\"/></w:r></w:p>");
            else if (i % EndnoteEvery == 2) writer.Write($"<w:p><w:r><w:t xml:space=\"preserve\">פסקה {i} עם הערת סיום</w:t></w:r><w:r><w:endnoteReference w:id=\"{NoteId(i)}\"/></w:r></w:p>");
            else WriteParagraph(writer, i, random);
        }
        writer.Write("<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/></w:sectPr></w:body></w:document>");
    }

    public const int FootnoteEvery = 20;     // paragraphs 1, 21, 41 … reference footnotes 1, 2, 3 …
    public const int EndnoteEvery = 100;     // paragraphs 2, 102, 202 … reference endnotes i, ii, iii …

    /// <summary>Note ids are not their numbers in a real file either: offset so the test catches a mix-up.</summary>
    private static int NoteId(int paragraphIndex) => 1000 + paragraphIndex;

    /// <summary>footnotes.xml / endnotes.xml with Word's two separator notes and one note per referencing paragraph, streamed into the entry.</summary>
    private static void WriteNotes(ZipArchive zip, string entryName, string root, string element, int blockCount, int every, string text, int fillerCharacters)
    {
        ZipArchiveEntry entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false), 1 << 16);
        writer.Write($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:{root} xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");
        writer.Write($"<w:{element} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{element}>");
        writer.Write($"<w:{element} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{element}>");
        int offset = element == "footnote" ? 1 : 2;
        var random = new Random(offset);
        for (int i = offset; i < blockCount; i += every)
        {
            writer.Write($"<w:{element} w:id=\"{NoteId(i)}\"><w:p><w:r><w:{element}Ref/></w:r><w:r><w:t xml:space=\"preserve\"> {text} של פסקה {i}</w:t></w:r></w:p>");
            for (int written = 0; written < fillerCharacters;)
            {
                // filler as separate paragraphs of mixed words, so it compresses like real text, not like a run of one letter
                writer.Write("<w:p><w:r><w:t xml:space=\"preserve\">");
                for (int word = 0; word < 60 && written < fillerCharacters; word++)
                {
                    string sample = SampleWords[random.Next(SampleWords.Length)];
                    writer.Write(sample); writer.Write(' ');
                    written += sample.Length + 1;
                }
                writer.Write("</w:t></w:r></w:p>");
            }
            writer.Write($"</w:{element}>");
        }
        writer.Write($"</w:{root}>");
    }

    private static void WriteParagraph(TextWriter writer, int index, Random random)
    {
        writer.Write("<w:p><w:r><w:t xml:space=\"preserve\">פסקה ");
        writer.Write(index);
        writer.Write(" </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>מודגש</w:t></w:r><w:r><w:t xml:space=\"preserve\"> ");
        int words = 8 + random.Next(40);
        for (int w = 0; w < words; w++) { writer.Write(SampleWords[random.Next(SampleWords.Length)]); writer.Write(' '); }
        writer.Write("</w:t></w:r><w:r><w:rPr><w:rStyle w:val=\"Emphasis\"/></w:rPr><w:t>English text ");
        writer.Write(index);
        writer.Write("</w:t></w:r></w:p>");
    }

    private static void WriteTable(TextWriter writer, int index)
    {
        writer.Write("<w:tbl><w:tblPr><w:bidiVisual/></w:tblPr>");
        for (int row = 0; row < 2; row++)
        {
            writer.Write("<w:tr>");
            for (int column = 0; column < 2; column++)
                writer.Write($"<w:tc><w:p><w:r><w:t>טבלה {index} תא {row},{column}</w:t></w:r></w:p></w:tc>");
            writer.Write("</w:tr>");
        }
        writer.Write("</w:tbl>");
    }

    private static readonly string[] SampleWords =
        ["בראשית", "ברא", "אלהים", "את", "השמים", "ואת", "הארץ", "והארץ", "היתה", "תהו", "ובהו", "וחשך", "על", "פני", "תהום", "ורוח", "מרחפת", "המים"];

    private static void AddText(ZipArchive zip, string name, string content)
    {
        ZipArchiveEntry entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
