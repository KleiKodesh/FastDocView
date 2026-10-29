namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    /// <summary>
    /// A right-to-left run (w:rtl) lays out its contents right to left, punctuation included (ECMA-376 1st edition, rtl
    /// "Right To Left Text"): in a left-to-right paragraph it is an RTL embedding, so a closing parenthesis, colon,
    /// quote, period or apostrophe at either end of Hebrew text stays on the Hebrew side and mirrors as Word draws it.
    /// In a right-to-left paragraph the paragraph's own base direction already does that: no extra markup.
    /// </summary>
    private static void RunDirection()
    {
        const string hebrew = "\u05D0\u05D1\u05D2 \u05D3\u05D4";
        foreach (string punctuation in new[] { "(", ")", ":", "\"", ".", "'" })
        {
            string text = punctuation + hebrew + punctuation;
            string run = $"<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space=\"preserve\">{System.Security.SecurityElement.Escape(text)}</w:t></w:r>";
            List<string> blocks = RenderAll("run-direction-" + (int)punctuation[0],
                $"<w:p><w:r><w:t xml:space=\"preserve\">ltr </w:t></w:r>{run}</w:p><w:p><w:pPr><w:bidi/></w:pPr>{run}</w:p>");
            TestReport.Check(blocks[0].Contains("unicode-bidi:embed;direction:rtl;"),
                $"ltr paragraph: rtl run with '{punctuation}' at both ends is an rtl embedding");
            TestReport.Check(blocks[1].Contains("dir=\"rtl\"") && !blocks[1].Contains("unicode-bidi"),
                $"rtl paragraph: rtl run with '{punctuation}' needs no embedding");
        }
        // a text box's own (right-to-left) paragraph must not leave its direction behind for the outer paragraph's runs
        string box = "<w:r><w:drawing><wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\"><w:txbxContent>"
            + "<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:rPr><w:rtl/></w:rPr><w:t>\u05D0</w:t></w:r></w:p></w:txbxContent></wp:inline></w:drawing></w:r>";
        List<string> boxed = RenderAll("run-direction-text-box",
            $"<w:p><w:r><w:t xml:space=\"preserve\">ltr </w:t></w:r>{box}<w:r><w:rPr><w:rtl/></w:rPr><w:t>\u05D1):</w:t></w:r></w:p>");
        TestReport.Check(boxed[0].IndexOf("unicode-bidi:embed;direction:rtl;", boxed[0].IndexOf("text-box", StringComparison.Ordinal), StringComparison.Ordinal) > 0,
            "ltr paragraph: an rtl run after a text box with an rtl paragraph is still an rtl embedding");
        List<string> plain = RenderAll("run-direction-ltr", $"<w:p><w:r><w:t xml:space=\"preserve\">(abc) {"\u05D0\u05D1"}:</w:t></w:r></w:p>");
        TestReport.Check(!plain[0].Contains("unicode-bidi"), "ltr paragraph: a run without w:rtl is not embedded");
    }

    /// <summary>RTF: \rtlch "the character data following this control word will be treated as a right-to-left run" (RTF 1.5 spec).</summary>
    private static void RtfRunDirection()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-run-direction.rtf");
        // \'e0\'e1 = two Hebrew letters in cp1255; "):" at the end of the right-to-left run
        File.WriteAllText(path, @"{\rtf1\ansi\ansicpg1255{\fonttbl{\f0\fcharset177 Arial;}}
\pard\ltrpar\f0 ltr {\rtlch \'e0\'e1 (\'e0\'e1):}\par
\pard\rtlpar\f0 {\rtlch \'e0\'e1 (\'e0\'e1):}\par
}");
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, 2);
        TestReport.Check(blocks[0].Html.Contains("unicode-bidi:embed;direction:rtl;"), @"rtf: \rtlch run in an \ltrpar paragraph is an rtl embedding");
        TestReport.Check(!blocks[1].Html.Contains("unicode-bidi"), @"rtf: \rtlch run in an \rtlpar paragraph needs no embedding");
    }
}
