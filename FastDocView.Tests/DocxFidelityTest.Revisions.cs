namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    private const string BoxBorders = "<w:pBdr><w:top w:val=\"single\" w:sz=\"8\" w:color=\"000000\"/><w:left w:val=\"single\" w:sz=\"8\" w:color=\"000000\"/><w:bottom w:val=\"single\" w:sz=\"8\" w:color=\"000000\"/><w:right w:val=\"single\" w:sz=\"8\" w:color=\"000000\"/><w:between w:val=\"single\" w:sz=\"4\" w:color=\"FF0000\"/></w:pBdr><w:spacing w:before=\"120\" w:after=\"120\"/>";

    /// <summary>
    /// Paragraph border groups (§17.3.1.24 pBdr, between §17.3.1.5): consecutive paragraphs with the same borders and indents are
    /// one box — inner top / bottom borders dropped, the between border drawn, the spacing inside the box. Across blocks (body)
    /// and among siblings (a cell).
    /// </summary>
    private static void BorderGroups()
    {
        string body = Paragraph(null, "", "one", BoxBorders) + Paragraph(null, "", "two", BoxBorders) + Paragraph(null, "", "three", BoxBorders)
                    + Paragraph(null, "", "alone", BoxBorders.Replace("<w:spacing", "<w:ind w:left=\"720\"/><w:spacing"))
                    + "<w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc>" + Paragraph(null, "", "c1", BoxBorders) + Paragraph(null, "", "c2", BoxBorders) + "</w:tc></w:tr></w:tbl>";
        List<string> html = RenderAll("border-groups", body);
        TestReport.Check(html[0].Contains("border-top:1pt solid #000000") && !html[0].Contains("border-bottom") && html[0].Contains("padding:1pt 4pt 6pt;"), "first of a group: top border, no bottom border, space-after inside the box");
        TestReport.Check(html[1].Contains("border-top:0.5pt solid #FF0000") && !html[1].Contains("border-bottom") && !html[1].Contains("margin-top"), "middle: the between border on top, no bottom border, no margins");
        TestReport.Check(html[2].Contains("border-top:0.5pt solid #FF0000") && html[2].Contains("border-bottom:1pt solid #000000"), "last of a group: between on top, its own bottom border");
        TestReport.Check(html[3].Contains("border-top:1pt solid #000000") && html[3].Contains("border-bottom:1pt solid #000000"), "a different indent starts its own box");
        TestReport.Check(html[4].Contains("#FF0000"), "paragraphs in a cell group too");
    }

    /// <summary>
    /// Tracked changes in Word's final view (Open-Xml-PowerTools RevisionProcessor): a deleted paragraph mark joins the paragraph to the
    /// next one, which keeps its own properties; deleted rows (trPr/w:del) and cells (tcPr/w:cellDel) disappear.
    /// </summary>
    private static void DeletedMarksRowsAndCells()
    {
        string deletedMark = "<w:pPr><w:rPr><w:del w:id=\"1\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>";
        string body = $"<w:p>{deletedMark}<w:r><w:t xml:space=\"preserve\">first </w:t></w:r></w:p>"
                    + "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>second</w:t></w:r></w:p>"
                    + "<w:tbl><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>"
                    + "<w:tr><w:trPr><w:del w:id=\"2\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"/></w:trPr><w:tc><w:p><w:r><w:t>gone-row</w:t></w:r></w:p></w:tc><w:tc><w:p/></w:tc></w:tr>"
                    + "<w:tr><w:tc><w:tcPr><w:cellDel w:id=\"3\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"/></w:tcPr><w:p><w:r><w:t>gone-cell</w:t></w:r></w:p></w:tc>"
                    + $"<w:tc><w:p>{deletedMark}<w:r><w:t xml:space=\"preserve\">in </w:t></w:r></w:p><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        List<string> html = RenderAll("revisions", body);
        TestReport.Check(html[0].Length == 0, "a paragraph whose mark is deleted renders nothing in its own block");
        TestReport.Check(html[1].Contains("text-align:center") && html[1].Contains(">first second<") || (html[1].Contains("first ") && html[1].IndexOf("first ", StringComparison.Ordinal) < html[1].IndexOf("second", StringComparison.Ordinal) && html[1].Contains("text-align:center")),
                         "...and its text starts the next paragraph, which keeps its own properties");
        TestReport.Check(!html[2].Contains("gone-row") && !html[2].Contains("gone-cell"), "deleted rows and cells are gone");
        TestReport.Check(html[2].Contains("in cell") && html[2].Split("<p").Length - 1 == 1, "a deleted mark inside a cell joins its sibling paragraph");
    }

    /// <summary>
    /// Search sees what is rendered: a tracked move's source (w:moveFrom, hidden in Word's final view) is neither shown nor
    /// found; a ruby (§17.3.3.25) shows its base with the guide text above it (HTML ruby) and only the base is searched - in the
    /// body and in a note.
    /// </summary>
    private static void MoveFromAndRubySearch()
    {
        const string Move = "w:id=\"7\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"";
        string ruby = "<w:r><w:ruby><w:rubyPr><w:rubyAlign w:val=\"center\"/><w:hps w:val=\"10\"/><w:hpsRaise w:val=\"18\"/><w:hpsBaseText w:val=\"20\"/><w:lid w:val=\"ja-JP\"/></w:rubyPr>"
                    + "<w:rt><w:r><w:t>guidetext</w:t></w:r></w:rt><w:rubyBase><w:r><w:t>basetext</w:t></w:r></w:rubyBase></w:ruby></w:r>";
        string body = $"<w:p><w:moveFrom {Move}><w:r><w:t>movedaway</w:t></w:r></w:moveFrom><w:r><w:t xml:space=\"preserve\">kept </w:t></w:r>{ruby}</w:p>"
                    + $"<w:p><w:moveTo {Move.Replace("7", "8")}><w:r><w:t>movedhere</w:t></w:r></w:moveTo><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
        string note = $"<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r><w:moveFrom {Move.Replace("7", "9")}><w:r><w:t>notemoved</w:t></w:r></w:moveFrom>{ruby.Replace("basetext", "notebase").Replace("guidetext", "noteguide")}</w:p></w:footnote>";
        var notes = new List<string>();
        int Hits(IWordWindowedDocument document, string query) => document.Search(query, WordDocumentSearchOptions.Default, 0, 100).HitCounts.Sum();
        // The notes are searched apart from the text (SearchNotes), over the note as it is rendered.
        int NoteHits(IWordWindowedDocument document, string query) => document.SearchNotes(query, WordDocumentSearchOptions.Default, 0, 100).HitCounts.Sum();
        var found = new Dictionary<string, int>();
        var foundInNotes = new Dictionary<string, int>();
        List<string> html = RenderAll("movefrom-ruby", body, footnotes: note, noteHtml: notes,
            inspect: document =>
            {
                foreach (string q in new[] { "movedaway", "movedhere", "basetext", "guidetext", "kept basetext", "notemoved", "notebase", "noteguide" }) found[q] = Hits(document, q);
                foreach (string q in new[] { "notemoved", "notebase", "basetext" }) foundInNotes[q] = NoteHits(document, q);
            });
        TestReport.Check(!html[0].Contains("movedaway") && html[1].Contains("movedhere"), "moveFrom hidden, moveTo shown");
        TestReport.Check(found["movedaway"] == 0 && found["movedhere"] == 1, $"search: moveFrom not found, moveTo found ({found["movedaway"]}, {found["movedhere"]})");
        TestReport.Check(html[0].Contains("<ruby>") && html[0].Contains("basetext") && html[0].Contains("<rt>") && html[0].Contains("guidetext"), "a ruby renders its base with its guide text (HTML ruby)");
        TestReport.Check(found["basetext"] == 1 && found["kept basetext"] == 1 && found["guidetext"] == 0, $"search: the ruby base is found in line, the guide is not ({found["basetext"]}, {found["kept basetext"]}, {found["guidetext"]})");
        TestReport.Check(notes.Count == 1 && !notes[0].Contains("notemoved") && notes[0].Contains("<ruby>") && notes[0].Contains("noteguide"), "in a note: moveFrom hidden, ruby rendered");
        TestReport.Check(found["notemoved"] == 0 && found["notebase"] == 0 && found["noteguide"] == 0, $"search: a note's text is never a hit on the text ({found["notemoved"]}, {found["notebase"]}, {found["noteguide"]})");
        TestReport.Check(foundInNotes["notemoved"] == 0 && foundInNotes["notebase"] == 1 && foundInNotes["basetext"] == 0,
                         $"notes search: moveFrom not found, the ruby base found, the text's words not ({foundInNotes["notemoved"]}, {foundInNotes["notebase"]}, {foundInNotes["basetext"]})");
    }
}
