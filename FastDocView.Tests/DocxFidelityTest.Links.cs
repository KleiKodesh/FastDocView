namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>w:hyperlink (r:id / w:anchor), HYPERLINK complex and simple fields (LibreOffice's switch reading); only http(s), mailto and in-document anchors become links.</summary>
    private static void Hyperlinks()
    {
        const string relationships = """
            <Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/a?b=1&amp;c=2" TargetMode="External"/>
            <Relationship Id="rIdL2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>
            """;
        static string Text(string text, string properties = "") => $"<w:r><w:rPr>{properties}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        string body = "<w:p><w:hyperlink r:id=\"rIdL1\" w:tooltip=\"Tip\">" + Text("site", "<w:b/>") + "</w:hyperlink></w:p>"
                    + "<w:p><w:hyperlink r:id=\"rIdL2\">" + Text("bad") + "</w:hyperlink></w:p>"
                    + "<w:p><w:bookmarkStart w:id=\"0\" w:name=\"Part 2\"/>" + Text("target") + "<w:bookmarkEnd w:id=\"0\"/><w:hyperlink w:anchor=\"Part 2\">" + Text("jump") + "</w:hyperlink></w:p>"
                    + "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"mailto:a@b.org\" \\o \"mail\" </w:instrText></w:r>"
                    + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Text("write") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" + Text(" after") + "</w:p>"
                    + "<w:p><w:fldSimple w:instr=\" HYPERLINK \\l &quot;Part 2&quot; \">" + Text("simple") + "</w:fldSimple></w:p>"
                    + "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> HYPERLINK \"file:///c:/x.txt\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Text("file") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        List<string> html = RenderAll("links", body, extraDocumentRelationships: relationships);
        TestReport.Check(html[0].Contains("<a class=\"link\" href=\"https://example.org/a?b=1&amp;c=2\" title=\"Tip\" rel=\"noopener noreferrer\" target=\"_blank\">site</a>"), "w:hyperlink r:id → an external link inside the run's span, with its tooltip");
        TestReport.Check(!html[1].Contains("<a ") && html[1].Contains("bad"), "a javascript: target renders as plain text");
        TestReport.Check(html[2].Contains("id=\"bookmark-Part%202\"") && html[2].Contains("href=\"#bookmark-Part%202\""), "w:anchor links to the bookmark's element");
        TestReport.Check(html[3].Contains("href=\"mailto:a@b.org\" title=\"mail\"") && html[3].Contains(">write</a>") && !html[3].Contains("after</a>"), "HYPERLINK field: only its result is the link");
        TestReport.Check(html[4].Contains("href=\"#bookmark-Part%202\""), "fldSimple HYPERLINK l-switch → an in-document link");
        TestReport.Check(!html[5].Contains("<a "), "a file: link is not followed");
    
        // a HYPERLINK field whose result runs across two paragraphs (the open field is recorded at index time per block)
        string spanning = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> HYPERLINK \"https://example.org/x\" </w:instrText></w:r>"
                        + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Text("part one") + "</w:p>"
                        + "<w:p>" + Text("part two") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" + Text(" plain") + "</w:p>";
        List<string> across = RenderAll("links-spanning", spanning);
        TestReport.Check(across[1].Contains("href=\"https://example.org/x\"") && across[1].Contains(">part two</a>") && !across[1].Contains("plain</a>"),
                         "a HYPERLINK field result that runs into the next paragraph stays a link there (index-time open-field state)");
    
        // review findings: a field begun in a paragraph whose mark is deleted (joined into the next block) closes there;
        // a deleted field end is not counted at index time; a text box inside a link is not inside the <a>
        string deletedMark = "<w:pPr><w:rPr><w:del w:id=\"9\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>";
        string joined = $"<w:p>{deletedMark}<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> HYPERLINK \"https://example.org/j\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Text("in ") + "</w:p>"
                      + "<w:p>" + Text("link") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>" + Text(" out") + "</w:p>";
        List<string> joinedHtml = RenderAll("links-joined", joined);
        TestReport.Check(joinedHtml[1].Contains(">link</a>") && !joinedHtml[1].Contains("out</a>"), "a link begun in a joined paragraph ends where its field ends");
        string deletedEnd = "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText> HYPERLINK \"https://example.org/d\" </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + Text("a")
                          + "<w:del w:id=\"7\" w:author=\"a\" w:date=\"2026-01-01T00:00:00Z\"><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:del></w:p>"
                          + "<w:p>" + Text("b") + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>";
        TestReport.Check(RenderAll("links-deleted-end", deletedEnd)[1].Contains(">b</a>"), "a deleted field end does not close the link for the next block");
        string boxInLink = "<w:p><w:hyperlink r:id=\"rIdL1\"><w:r><w:t>x</w:t><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></wp:inline></w:drawing></w:r></w:hyperlink></w:p>";
        string boxHtml = RenderAll("links-textbox", boxInLink, extraDocumentRelationships: relationships)[0];
        // a link inside a footnote: its r:id is in footnotes.xml.rels (the note part's own relationships)
        const string footnoteRelationships = """
            <Relationship Id="rIdN1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/note" TargetMode="External"/>
            """;
        string noteBody = "<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r><w:hyperlink r:id=\"rIdN1\">" + Text("in note") + "</w:hyperlink></w:p></w:footnote>";
        var notes = new List<string>();
        RenderAll("links-footnote", "<w:p>" + Text("text") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>",
                  footnotes: noteBody, footnoteRelationships: footnoteRelationships, noteHtml: notes);
        TestReport.Check(notes.Count == 1 && notes[0].Contains("href=\"https://example.org/note\"") && notes[0].Contains(">in note</a>"),
                         "a link inside a footnote resolves through footnotes.xml.rels");
        TestReport.Check(boxHtml.Split("<a ").Length - 1 == 1 && !boxHtml.Contains("<a class=\"link\" href=\"https://example.org/a?b=1&amp;c=2\" title") , "a text box inside a link is not wrapped in another link");
    }


}
