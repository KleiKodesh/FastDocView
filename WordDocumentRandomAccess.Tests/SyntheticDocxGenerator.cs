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

    /// <summary>An Office-like theme: minor latin Calibri, empty a:cs, Hebrew supplemental face Arial; accent1 4472C4 (Word's 2016 palette).</summary>
    private const string ThemeXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>
          <a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
            <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
          <a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Times New Roman"/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hebr" typeface="Arial"/></a:minorFont></a:fontScheme>
          <a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
        </a:themeElements></a:theme>
        """;

    /// <summary>settings.xml with the theme hooks: bidi font language Hebrew (picks the "Hebr" supplemental face) and accent2 remapped onto accent6.</summary>
    private const string SettingsXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:themeFontLang w:val="en-US" w:bidi="he-IL"/>
          <w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent6" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
        </w:settings>
        """;

    private const string ContentTypesXml = """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
          <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
          <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
          <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
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
          <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
          <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        </Relationships>
        """;

    public static int ExpectedBlockCount(int blockCount) => blockCount + 1;   // + the trailing w:sectPr

    /// <param name="noteFillerCharacters">Extra text per note — makes the notes parts as large as a test needs (streamed, never built in memory).</param>
    /// <param name="withThemeSamples">Also writes theme1.xml + settings.xml and turns paragraphs 45-49 into the theme / underline / emphasis samples (see <see cref="ThemeSampleParagraph"/>).</param>
    public static void Write(string path, int blockCount, int noteFillerCharacters = 0, bool withThemeSamples = false)
    {
        using var file = File.Create(path);
        using var zip = new ZipArchive(file, ZipArchiveMode.Create);
        AddText(zip, "[Content_Types].xml", ContentTypesXml);
        AddText(zip, "_rels/.rels", RootRelsXml);
        AddText(zip, "word/_rels/document.xml.rels", DocumentRelsXml);
        AddText(zip, "word/styles.xml", StylesXml);
        AddText(zip, "word/numbering.xml", NumberingXml);
        if (withThemeSamples) { AddText(zip, "word/settings.xml", SettingsXml); AddText(zip, "word/theme/theme1.xml", ThemeXml); }
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
            else if (withThemeSamples && ThemeSampleParagraph(i) is { } sample) writer.Write(sample);
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

    public const int ThemeFontParagraph = 45, ThemeColorParagraph = 46, UnderlineParagraph = 47, EmphasisParagraph = 48, MappedColorParagraph = 49, UnderlinedParagraphMarkParagraph = 55;

    /// <summary>
    /// The formatting samples DocxThemeAndUnderlineTest checks: a theme complex-script font over an explicit one,
    /// a theme color with a tint over a wrong explicit val, a colored double underline, an emphasis mark, a remapped accent.
    /// </summary>
    private static string? ThemeSampleParagraph(int index) => index switch
    {
        ThemeFontParagraph => $"<w:p><w:r><w:rPr><w:rFonts w:cs=\"David\" w:cstheme=\"minorBidi\"/><w:rtl/></w:rPr><w:t>פסקה {index} גופן ערכת נושא</w:t></w:r></w:p>",
        ThemeColorParagraph => $"<w:p><w:r><w:rPr><w:color w:val=\"FF0000\" w:themeColor=\"accent1\" w:themeTint=\"99\"/><w:shd w:val=\"clear\" w:fill=\"FF0000\" w:themeFill=\"accent6\" w:themeFillShade=\"80\"/></w:rPr><w:t>פסקה {index} צבע ערכת נושא</w:t></w:r></w:p>",
        UnderlineParagraph => $"<w:p><w:r><w:rPr><w:u w:val=\"double\" w:color=\"00B050\"/></w:rPr><w:t>פסקה {index} קו תחתון כפול</w:t></w:r><w:r><w:rPr><w:u w:val=\"none\"/></w:rPr><w:t> ללא קו</w:t></w:r><w:r><w:rPr><w:u w:val=\"wavyHeavy\" w:themeColor=\"text2\"/></w:rPr><w:t> גלי</w:t></w:r></w:p>",
        EmphasisParagraph => $"<w:p><w:r><w:rPr><w:em w:val=\"underDot\"/></w:rPr><w:t>פסקה {index} סימן הדגשה</w:t></w:r><w:r><w:rPr><w:em w:val=\"comma\"/></w:rPr><w:t> מעל</w:t></w:r></w:p>",
        UnderlinedParagraphMarkParagraph => $"<w:p><w:pPr><w:rPr><w:u w:val=\"single\"/></w:rPr></w:pPr><w:r><w:rPr><w:u w:val=\"dotted\"/></w:rPr><w:t>פסקה {index} מנוקד בתוך פסקה מקווקוות</w:t></w:r></w:p>",
        MappedColorParagraph => $"<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"8\" w:color=\"FF0000\" w:themeColor=\"accent2\"/></w:pBdr></w:pPr><w:r><w:rPr><w:color w:val=\"FF0000\" w:themeColor=\"background1\" w:themeShade=\"80\"/></w:rPr><w:t>פסקה {index} מיפוי צבעים</w:t></w:r></w:p>",
        _ => null,
    };

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
