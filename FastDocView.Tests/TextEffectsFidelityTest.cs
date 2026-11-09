using FastDocView.DocBinary;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// Character scale in every format (.docx w:w, .doc sprmCCharScale from test-data/word-text-effects/text-effects-v1.doc made
/// by Word 16, RTF \charscalex, ODT style:text-scale) and the Word 2010 text effects w14:textFill (solid with alpha, scheme
/// colour, gradient), w14:glow and w14:shadow, from synthetic .docx.
/// </summary>
public static class TextEffectsFidelityTest
{
    private const string W14 = "xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\"";
    private const string ScaleBox = "<span class=\"char-scale\" style=\"display:inline-flex;justify-content:center;width:calc-size(max-content,size * ";

    public static void Run()
    {
        TestReport.Step("Text effects: character scale (docx, doc, rtf, odt), w14 text fill, glow, shadow");
        WordHtmlTextScale();
        DocxScale();
        DocScale();
        RtfScale();
        OdtScale();
        DocxTextFill();
        DocxGlowAndShadow();
    }

    private static void WordHtmlTextScale()
    {
        var html = new System.Text.StringBuilder();
        WordHtmlText.AppendScaled(html, "ab cd", 100);
        TestReport.Check(html.ToString() == "ab cd", "scale 100: text unchanged");
        html.Clear();
        WordHtmlText.AppendScaled(html, "a<b>x y</b>c d<br>e", 80);
        string s = html.ToString();
        html.Clear();
        WordHtmlText.AppendScaled(html, "ab</a><a>cd", 80);
        TestReport.Check(html.ToString().StartsWith(ScaleBox) && html.ToString().Contains("ab</span></span></a>"), "scale: a stray close tag ends the word box first: " + html);
        html.Clear();
        WordHtmlText.AppendScaled(html, "a<b>x y</b>c d<br>e", 80);
        s = html.ToString();
        TestReport.Check(s.Split(ScaleBox).Length - 1 == 2 && s.Contains("a<b>x y</b>c</span></span><span class=\"char-scale-space\"") && s.Contains(">"+" "+"</span></span><wbr>") && s.Contains("d<br>e</span></span>"),
                         "scale: one box per word, a space inside an open tag does not split, <br> is void: " + s);
    }

    private static void DocxScale()
    {
        static string Run(string properties, string text) => $"<w:r><w:rPr>{properties}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";
        List<string> html = DocxFidelityTest.RenderAll("char-scale",
            "<w:p>" + Run("<w:w w:val=\"80\"/>", "two words") + "</w:p>"
          + "<w:p>" + Run("<w:w w:val=\"150\"/><w:u w:val=\"single\"/>", "wide") + "</w:p>"
          + "<w:p>" + Run("<w:w w:val=\"700\"/>", "bad") + "</w:p>"
          + "<w:p>" + Run("<w:w w:val=\"80\"/>", string.Join(" ", Enumerable.Range(0, 80).Select(i => "word" + i))) + "</w:p>");   // 3: wraps (the live check)
        TestReport.Check(html[0].Contains(ScaleBox + "0.8);text-decoration:inherit\"><span style=\"display:inline-block;flex:none;transform:scaleX(0.8);text-decoration:inherit\">two</span></span><span class=\"char-scale-space\""),
                         "docx w:w 80: each word in its own scaled box");
        TestReport.Check(html[1].Contains("text-decoration:underline;\">" + ScaleBox + "1.5)"), "docx w:w 150 inside the run's underline");
        TestReport.Check(!html[2].Contains("char-scale"), "docx w:w 700 (out of 1..600) = 100 (DomainMapper.cxx:2496)");
    }

    private static void DocScale()
    {
        string path = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-text-effects", "text-effects-v1.doc"));
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-text-effects/text-effects-v1.doc present"); return; }
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-text-effects-doc") };
        using var document = (DocBinaryWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        List<string> html = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html).ToList();
        TestReport.Check(html.Count >= 3 && html[0].Contains(ScaleBox + "0.8)") && html[0].Split(ScaleBox).Length - 1 == 3, ".doc sprmCCharScale 80 (Word 16 Font.Scaling): three scaled words");
        TestReport.Check(html.Count >= 3 && !html[1].Contains("char-scale"), ".doc scale 100: plain text");
        TestReport.Check(html.Count >= 3 && html[2].Contains(ScaleBox + "1.5)"), ".doc scale 150");
    }

    private static void RtfScale()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-char-scale.rtf");
        File.WriteAllText(path, @"{\rtf1\ansi{\fonttbl{\f0 Arial;}}
\pard\f0 {\charscalex80 narrow text}\par
\pard\f0 {\charscalex80\plain normal}\par
}");
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, 2);
        TestReport.Check(blocks[0].Html.Split(ScaleBox).Length - 1 == 2 && blocks[0].Html.Contains(ScaleBox + "0.8)"), @"rtf \charscalex80: two scaled words");
        TestReport.Check(!blocks[1].Html.Contains("char-scale"), @"rtf \plain resets \charscalex");
    }

    private static void OdtScale()
    {
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-char-scale.fodt");
        File.WriteAllText(path, "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
            + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
            + "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:automatic-styles><style:style style:name=\"T1\" style:family=\"text\"><style:text-properties style:text-scale=\"60%\"/></style:style></office:automatic-styles>"
            + "<office:body><office:text><text:p>plain <text:span text:style-name=\"T1\">thin text</text:span></text:p></office:text></office:body></office:document>");
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = document.RenderBlocks(0, 1)[0].Html;
        TestReport.Check(html.Split(ScaleBox).Length - 1 == 2 && html.Contains(ScaleBox + "0.6)") && html.Contains("plain "), "odt style:text-scale 60%: the span's two words scaled, the rest not: " + html);
    }

    private static string Run(string properties, string text) => $"<w:r><w:rPr>{properties}</w:rPr><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static void DocxTextFill()
    {
        List<string> html = DocxFidelityTest.RenderAll("text-fill",
            "<w:p>" + Run($"<w14:textFill {W14}><w14:solidFill><w14:srgbClr w14:val=\"FF0000\"><w14:alpha w14:val=\"40000\"/></w14:srgbClr></w14:solidFill></w14:textFill>", "alpha") + "</w:p>"
          + "<w:p>" + Run($"<w14:textFill {W14}><w14:solidFill><w14:schemeClr w14:val=\"accent1\"><w14:lumMod w14:val=\"75000\"/></w14:schemeClr></w14:solidFill></w14:textFill>", "scheme") + "</w:p>"
          + "<w:p>" + Run($"<w14:textFill {W14}><w14:gradFill><w14:gsLst><w14:gs w14:pos=\"0\"><w14:srgbClr w14:val=\"FF0000\"/></w14:gs>"
                          + "<w14:gs w14:pos=\"100000\"><w14:srgbClr w14:val=\"0000FF\"><w14:alpha w14:val=\"50000\"/></w14:srgbClr></w14:gs></w14:gsLst><w14:lin w14:ang=\"5400000\" w14:scaled=\"0\"/></w14:gradFill></w14:textFill>", "linear") + "</w:p>"
          + "<w:p>" + Run($"<w14:textFill {W14}><w14:gradFill><w14:gsLst><w14:gs w14:pos=\"0\"><w14:srgbClr w14:val=\"FFFF00\"/></w14:gs>"
                          + "<w14:gs w14:pos=\"100000\"><w14:srgbClr w14:val=\"008000\"/></w14:gs></w14:gsLst><w14:path w14:path=\"circle\"><w14:fillToRect w14:l=\"50000\" w14:t=\"50000\" w14:r=\"50000\" w14:b=\"50000\"/></w14:path></w14:gradFill></w14:textFill>", "radial") + "</w:p>"
          + "<w:p>" + Run($"<w:highlight w:val=\"yellow\"/><w14:textFill {W14}><w14:gradFill><w14:gsLst><w14:gs w14:pos=\"0\"><w14:srgbClr w14:val=\"FF0000\"/></w14:gs>"
                          + "<w14:gs w14:pos=\"100000\"><w14:srgbClr w14:val=\"0000FF\"/></w14:gs></w14:gsLst><w14:lin w14:ang=\"0\" w14:scaled=\"0\"/></w14:gradFill></w14:textFill>", "highlighted") + "</w:p>");
        TestReport.Check(html[0].Contains("color:rgba(255,0,0,0.6);"), "w14:textFill solid + w14:alpha 40% transparency = rgba opacity 0.6: " + html[0]);
        // no theme part in the fixture: accent1 resolves to nothing, so the run keeps the default colour rather than a wrong one
        TestReport.Check(!html[1].Contains("transparent"), "w14:textFill schemeClr without a theme: no fill css");
        TestReport.Check(html[2].Contains("background-image:linear-gradient(180deg,#FF0000 0%,rgba(0,0,255,0.5) 100%),none;-webkit-background-clip:text,border-box;background-clip:text,border-box;color:transparent;"),
                         "w14:gradFill lin 90 deg = CSS 180deg (top to bottom), stops with alpha: " + html[2]);
        TestReport.Check(html[4].Contains("background-clip:text,border-box;") && html[4].Contains("background-color:"), "gradient fill + highlight: the highlight keeps the border box: " + html[4]);
        TestReport.Check(html[3].Contains("radial-gradient(circle farthest-corner at 50% 50%,#FFFF00 0%,#008000 100%)"), "w14:gradFill path circle = radial gradient from fillToRect's centre: " + html[3]);
    }

    private static void DocxGlowAndShadow()
    {
        List<string> html = DocxFidelityTest.RenderAll("glow-shadow",
            "<w:p>" + Run($"<w14:glow {W14} w14:rad=\"127000\"><w14:srgbClr w14:val=\"00B0F0\"><w14:alpha w14:val=\"60000\"/></w14:srgbClr></w14:glow>", "glow") + "</w:p>"
          + "<w:p>" + Run($"<w14:shadow {W14} w14:blurRad=\"38100\" w14:dist=\"25400\" w14:dir=\"2700000\" w14:sx=\"100000\" w14:sy=\"100000\" w14:kx=\"0\" w14:ky=\"0\" w14:algn=\"tl\">"
                          + "<w14:srgbClr w14:val=\"000000\"><w14:alpha w14:val=\"60000\"/></w14:srgbClr></w14:shadow>", "shadow") + "</w:p>"
          + "<w:p>" + Run("<w:shadow/>", "legacy") + "</w:p>");
        TestReport.Check(html[0].Contains("text-shadow:0 0 5pt rgba(0,176,240,0.4),0 0 10pt rgba(0,176,240,0.4);"), "w14:glow rad 10pt, 60% transparent: " + html[0]);
        TestReport.Check(html[1].Contains("text-shadow:1.41pt 1.41pt 3pt rgba(0,0,0,0.4);"), "w14:shadow dist 2pt dir 45 deg (down-right), blur 3pt, 60% transparent: " + html[1]);
        TestReport.Check(html[2].Contains("text-shadow:1px 1px 1px rgba(0,0,0,.45);"), "legacy w:shadow keeps the generic shadow");
    }
}
