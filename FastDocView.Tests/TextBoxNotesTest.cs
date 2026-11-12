using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// Notes inside text boxes: an ODF text:note in a draw:frame / draw:text-box (nested too) or in a shape's text, and a .docx
/// w:footnoteReference / w:endnoteReference in w:txbxContent (DrawingML and VML), are references in the body like any other:
/// numbered in document order, their bodies in the notes pane, found by notes search and never by body search.
/// LibreOffice drops such notes (sw thints.cxx SwTextNode::InsertHint, RES_TXTATR_FTN: a footnote outside the body or a split
/// fly is destroyed) and refuses the file when one is in a shape's text; Word's UI does not insert them either - we show what
/// the file holds, in the order Word's draft view reads it.
/// </summary>
public static class TextBoxNotesTest
{
    public static void Run()
    {
        TestReport.Step("Notes inside text boxes and shapes (.fodt, .docx)");
        string directory = Path.Combine(TestReport.OutputDirectory, "textbox-notes");
        Directory.CreateDirectory(directory);
        string fodt = Path.Combine(directory, "textbox-notes-v1.fodt"), docx = Path.Combine(directory, "textbox-notes-v1.docx");
        File.WriteAllText(fodt, Fodt, new UTF8Encoding(false));
        WriteDocx(docx);
        Check(fodt, ["NOTEONE", "NOTETWO", "NOTEFOUR", "NOTEFIVE", "NOTEART", "NOTETAIL", "NOTEGRPA", "NOTEGRPB", "NOTEGRPEND", "NOTESIX"], ["NOTETHREE"], ["Boxed", "Nested", "InShape", "GroupText"],
              "f1 f2 e1 f3 f4 f6 f7 f9 f10");   // Fontwork draws plain text: its notes (f5, f8 in a group) have no mark, the next ones keep their numbers
        Check(docx, ["NOTEONE", "NOTETWO", "NOTEFOUR", "NOTEFIVE", "NOTESIX"], ["NOTETHREE"], ["Boxed", "InVml"], "f1 f2 e1 f3 f4 f5");
    }

    private static void Check(string path, string[] footnotes, string[] endnotes, string[] boxedWords, string marks)
    {
        string name = Path.GetFileName(path);
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") });
        string body = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        string bodyMarks = string.Join(" ", System.Text.RegularExpressions.Regex.Matches(body, @"data-note-kind=""(f|e)\w*"" data-note-number=""(\d+)""")
            .Select(match => match.Groups[1].Value + match.Groups[2].Value));
        TestReport.Check(bodyMarks == marks && document.NoteReferences.Count == footnotes.Length + endnotes.Length,
            $"{name}: note marks in the body {bodyMarks} (expected {marks}), {document.NoteReferences.Count} notes indexed (expected {footnotes.Length + endnotes.Length})");
        IReadOnlyList<WordDocumentNoteHtml> notes = document.RenderNotesForDisplay(0, document.NoteReferences.Count);
        string[] footnoteOrder = notes.Where(note => note.Kind == WordDocumentNoteKind.Footnote).Select(note => footnotes.FirstOrDefault(note.Html.Contains) ?? "?").ToArray();
        string[] endnoteOrder = notes.Where(note => note.Kind == WordDocumentNoteKind.Endnote).Select(note => endnotes.FirstOrDefault(note.Html.Contains) ?? "?").ToArray();
        int[] footnoteNumbers = notes.Where(note => note.Kind == WordDocumentNoteKind.Footnote).Select(note => note.Number).ToArray();
        TestReport.Check(footnoteOrder.SequenceEqual(footnotes) && endnoteOrder.SequenceEqual(endnotes) && footnoteNumbers.SequenceEqual(Enumerable.Range(1, footnotes.Length)),
            $"{name}: notes pane in document order, numbered 1..{footnotes.Length} ({string.Join(",", footnoteOrder)} | {string.Join(",", endnoteOrder)})");
        foreach (string word in new[] { "NOTETWO", "NOTETHREE", "NOTEFIVE" })
        {
            WordDocumentSearchBatch inNotes = document.SearchNotes(word, WordDocumentSearchOptions.Default, 0, document.NoteReferences.Count);
            WordDocumentSearchBatch inBody = document.Search(word, WordDocumentSearchOptions.Default, 0, document.BlockCount);
            TestReport.Check(inNotes.HitBlocks.Length == 1 && inBody.HitBlocks.Length == 0, $"{name}: \"{word}\" (a boxed note) found in one note, not in the body");
        }
        foreach (string word in boxedWords)
            TestReport.Check(document.Search(word, WordDocumentSearchOptions.Default, 0, document.BlockCount).HitBlocks.Length == 1, $"{name}: boxed text \"{word}\" found in the body");
    }

    private const string Fodt = @"<?xml version=""1.0"" encoding=""UTF-8""?>
<office:document xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" xmlns:text=""urn:oasis:names:tc:opendocument:xmlns:text:1.0"" xmlns:draw=""urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"" xmlns:svg=""urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"" xmlns:fo=""urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"" xmlns:style=""urn:oasis:names:tc:opendocument:xmlns:style:1.0"" office:version=""1.3"" office:mimetype=""application/vnd.oasis.opendocument.text"">
<office:body><office:text>
<text:p>Alpha<text:note text:id=""n1"" text:note-class=""footnote""><text:note-citation>1</text:note-citation><text:note-body><text:p>NOTEONE</text:p></text:note-body></text:note> before box.</text:p>
<text:p>Anchor<draw:frame draw:name=""F1"" text:anchor-type=""paragraph"" svg:width=""5cm"" svg:height=""3cm""><draw:text-box><text:p>Boxed<text:note text:id=""n2"" text:note-class=""footnote""><text:note-citation>2</text:note-citation><text:note-body><text:p>NOTETWO</text:p></text:note-body></text:note> text<draw:frame draw:name=""F2"" text:anchor-type=""as-char"" svg:width=""3cm"" svg:height=""1cm""><draw:text-box><text:p>Nested<text:note text:id=""n3"" text:note-class=""endnote""><text:note-citation>i</text:note-citation><text:note-body><text:p>NOTETHREE</text:p></text:note-body></text:note></text:p></draw:text-box></draw:frame></text:p></draw:text-box></draw:frame> after<text:note text:id=""n4"" text:note-class=""footnote""><text:note-citation>3</text:note-citation><text:note-body><text:p>NOTEFOUR</text:p></text:note-body></text:note>.</text:p>
<text:p>Shape<draw:custom-shape text:anchor-type=""paragraph"" svg:width=""4cm"" svg:height=""2cm""><text:p>InShape<text:note text:id=""n5"" text:note-class=""footnote""><text:note-citation>4</text:note-citation><text:note-body><text:p>NOTEFIVE</text:p></text:note-body></text:note></text:p><draw:enhanced-geometry svg:viewBox=""0 0 21600 21600"" draw:type=""rectangle""/></draw:custom-shape> end.</text:p>
<text:p>Art<draw:custom-shape text:anchor-type=""as-char"" svg:width=""4cm"" svg:height=""1cm""><text:p>ARTWORD<text:note text:id=""n7"" text:note-class=""footnote""><text:note-citation>5</text:note-citation><text:note-body><text:p>NOTEART</text:p></text:note-body></text:note></text:p><draw:enhanced-geometry svg:viewBox=""0 0 21600 21600"" draw:type=""fontwork-plain-text"" draw:text-path=""true"" draw:enhanced-path=""M 0 0 L 21600 0 N M 0 21600 L 21600 21600 N""/></draw:custom-shape> tail<text:note text:id=""n8"" text:note-class=""footnote""><text:note-citation>6</text:note-citation><text:note-body><text:p>NOTETAIL</text:p></text:note-body></text:note></text:p>
<text:p>Group<draw:g text:anchor-type=""as-char""><draw:custom-shape svg:x=""0cm"" svg:y=""0cm"" svg:width=""4cm"" svg:height=""1cm""><text:p>GroupText<text:note text:id=""n9"" text:note-class=""footnote""><text:note-citation>7</text:note-citation><text:note-body><text:p>NOTEGRPA</text:p></text:note-body></text:note></text:p><draw:enhanced-geometry svg:viewBox=""0 0 21600 21600"" draw:type=""rectangle""/></draw:custom-shape><draw:custom-shape svg:x=""0cm"" svg:y=""1cm"" svg:width=""4cm"" svg:height=""1cm""><text:p>ARTTWO<text:note text:id=""n10"" text:note-class=""footnote""><text:note-citation>8</text:note-citation><text:note-body><text:p>NOTEGRPB</text:p></text:note-body></text:note></text:p><draw:enhanced-geometry svg:viewBox=""0 0 21600 21600"" draw:type=""fontwork-plain-text"" draw:text-path=""true"" draw:enhanced-path=""M 0 0 L 21600 0 N M 0 21600 L 21600 21600 N""/></draw:custom-shape></draw:g> end<text:note text:id=""n11"" text:note-class=""footnote""><text:note-citation>9</text:note-citation><text:note-body><text:p>NOTEGRPEND</text:p></text:note-body></text:note></text:p>
<text:p>Omega<text:note text:id=""n6"" text:note-class=""footnote""><text:note-citation>10</text:note-citation><text:note-body><text:p>NOTESIX</text:p></text:note-body></text:note></text:p>
</office:text></office:body></office:document>";

    private const string WordNamespaces = "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" xmlns:v=\"urn:schemas-microsoft-com:vml\"";

    private static string Run(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
    private static string Footnote(int id) => $"<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:footnoteReference w:id=\"{id}\"/></w:r>";
    private static string Endnote(int id) => $"<w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:endnoteReference w:id=\"{id}\"/></w:r>";

    /// <summary>A DrawingML text box (wps:txbx / w:txbxContent) anchored in the paragraph.</summary>
    private static string TextBox(string paragraphContent) =>
        "<w:r><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\" simplePos=\"0\" relativeHeight=\"1\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\">"
        + "<wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>"
        + "<wp:extent cx=\"1800000\" cy=\"900000\"/><wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/><wp:wrapSquare wrapText=\"bothSides\"/><wp:docPr id=\"1\" name=\"Box1\"/>"
        + "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"><wps:wsp><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1800000\" cy=\"900000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></wps:spPr>"
        + $"<wps:txbx><w:txbxContent><w:p>{paragraphContent}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

    private static string Notes(string kind, (int Id, string Text)[] notes)
    {
        string reference = kind == "footnote" ? "footnoteRef" : "endnoteRef";
        var xml = new StringBuilder($"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:{kind}s {WordNamespaces}>");
        xml.Append($"<w:{kind} w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:{kind}><w:{kind} w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:{kind}>");
        foreach ((int id, string text) in notes)
            xml.Append($"<w:{kind} w:id=\"{id}\"><w:p><w:r><w:rPr><w:vertAlign w:val=\"superscript\"/></w:rPr><w:{reference}/></w:r>{Run(" " + text)}</w:p></w:{kind}>");
        return xml.Append($"</w:{kind}s>").ToString();
    }

    private static void WriteDocx(string path)
    {
        string body = $"<w:p>{Run("Alpha")}{Footnote(1)}{Run(" before box.")}</w:p>"
            + $"<w:p>{Run("Anchor")}{TextBox(Run("Boxed") + Footnote(2) + Run(" text") + Endnote(1))}{Run(" after")}{Footnote(3)}</w:p>"
            + $"<w:p>{Run("Vml")}<w:r><w:pict><v:shape style=\"width:100pt;height:50pt\"><v:textbox><w:txbxContent><w:p>{Run("InVml")}{Footnote(4)}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"
            + $"<w:p>{Run("Omega")}{Footnote(5)}</w:p>";
        File.Delete(path);
        using ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Create);
        void Add(string name, string content) { using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(content); }
        Add("[Content_Types].xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/><Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/><Override PartName=\"/word/endnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml\"/></Types>");
        Add("_rels/.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
        Add("word/_rels/document.xml.rels", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/></Relationships>");
        Add("word/document.xml", $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document {WordNamespaces}><w:body>{body}<w:sectPr/></w:body></w:document>");
        Add("word/footnotes.xml", Notes("footnote", [(1, "NOTEONE"), (2, "NOTETWO"), (3, "NOTEFOUR"), (4, "NOTEFIVE"), (5, "NOTESIX")]));
        Add("word/endnotes.xml", Notes("endnote", [(1, "NOTETHREE")]));
    }
}
