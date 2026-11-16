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
        TestReport.Step("Text effects: character scale (docx, doc, rtf, odt), w14 text fill, glow, shadow, OpenType typography, reflection");
        WordHtmlTextScale();
        DocxScale();
        DocScale();
        RtfScale();
        OdtScale();
        DocxTextFill();
        DocxGlowAndShadow();
        DocxOpenTypeTypography();
        DocxReflection();
        WordScriptShaping();
        OdtFontFeatures();
        TestReport.Check(FastDocView.Markup.WebArchiveStylesheet.SanitizeInline("font-variant-ligatures:no-common-ligatures;font-variant-numeric:oldstyle-nums;position:absolute")
                         == "font-variant-ligatures:no-common-ligatures;font-variant-numeric:oldstyle-nums;", "web archive: ligature and number CSS kept, as the browser prints it");
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

    /// <summary>ODT: LibreOffice's font-name features ("Name:-liga&amp;ss05", vcl FeatureParser) as font-feature-settings; a plain name
    /// keeps the browser's (= HarfBuzz's) defaults.</summary>
    private static void OdtFontFeatures()
    {
        TestReport.Check(FastDocView.Odt.OdtStyleCascade.FontWithoutFeatures("Gabriola:-liga&ss05&calt=0&lang=he&kern[1:3]", out string? css) == "Gabriola"
                         && css == "'liga' 0,'ss05' 1,'calt' 0,'kern' 1", "LibreOffice font features parsed: " + css);
        TestReport.Check(FastDocView.Odt.OdtStyleCascade.FontWithoutFeatures("Arial", out string? none) == "Arial" && none is null, "plain font name: no features");
        string path = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-font-features.fodt");
        File.WriteAllText(path, "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" "
            + "xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" "
            + "xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\" "
            + "xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\" office:version=\"1.3\" office:mimetype=\"application/vnd.oasis.opendocument.text\">"
            + "<office:font-face-decls><style:font-face style:name=\"Gabriola:-liga\" svg:font-family=\"Gabriola:-liga\"/></office:font-face-decls>"
            + "<office:automatic-styles><style:style style:name=\"T1\" style:family=\"text\"><style:text-properties style:font-name=\"Gabriola:-liga\"/></style:style></office:automatic-styles>"
            + "<office:body><office:text><text:p>office <text:span text:style-name=\"T1\">office</text:span></text:p></office:text></office:body></office:document>");
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = document.RenderBlocks(0, 1)[0].Html;
        TestReport.Check(html.Contains("font-feature-settings:'liga' 0;") && html.Contains("Gabriola") && !html.Contains("Gabriola:") && !html.Contains("font-variant-ligatures")
                         && !document.Stylesheet.Contains("font-variant-ligatures"), "odt Gabriola:-liga: family + liga off, nothing else: " + html);
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
    /// <summary>w14 ligatures / numForm / numSpacing / cntxtAlts / stylisticSets, as Word 16 writes them (Font.Ligatures,
    /// NumberForm, NumberSpacing, ContextualAlternates, StylisticSet; Word's docDefaults carry standardContextual).</summary>
    private static void DocxOpenTypeTypography()
    {
        List<string> html = DocxFidelityTest.RenderAll("opentype-v1",
            "<w:p>" + Run("", "plain") + "</w:p>"
          + "<w:p>" + Run($"<w14:ligatures {W14} w14:val=\"standard\"/>", "office") + "</w:p>"
          + "<w:p>" + Run($"<w14:ligatures {W14} w14:val=\"all\"/><w14:cntxtAlts {W14}/>", "all") + "</w:p>"
          + "<w:p>" + Run($"<w14:numForm {W14} w14:val=\"oldStyle\"/><w14:numSpacing {W14} w14:val=\"tabular\"/>", "0123") + "</w:p>"
          + "<w:p>" + Run($"<w14:stylisticSets {W14}><w14:styleSet w14:id=\"2\"/><w14:styleSet w14:id=\"3\" w14:val=\"0\"/><w14:styleSet w14:id=\"14\"/></w14:stylisticSets>", "Gabriola") + "</w:p>"
          + "<w:p>" + Run($"<w:rtl/><w14:numForm {W14} w14:val=\"oldStyle\"/>", "123") + "</w:p>"
          + $"<w:p><w:pPr><w:rPr><w14:ligatures {W14} w14:val=\"standardContextual\"/></w:rPr></w:pPr>" + Run($"<w14:ligatures {W14} w14:val=\"standardContextual\"/>", "a") + Run("<w:smallCaps/>", "b") + "</w:p>",
            styles: $"<w:docDefaults><w:rPrDefault><w:rPr><w14:numForm {W14} w14:val=\"lining\"/></w:rPr></w:rPrDefault></w:docDefaults>");
        TestReport.Check(html[0].Contains("font-variant-ligatures:no-common-ligatures no-contextual;"), "no w14:ligatures: Word draws no ligatures and no contextual alternates: " + html[0]);
        TestReport.Check(html[0].Contains("font-variant-numeric:lining-nums;"), "docDefaults numForm lining reaches the paragraph: " + html[0]);
        TestReport.Check(html[1].Contains("font-variant-ligatures:common-ligatures no-contextual;font-feature-settings:'liga' 1,'clig' 0;"), "standard = liga only: " + html[1]);
        TestReport.Check(html[2].Contains("font-variant-ligatures:common-ligatures discretionary-ligatures historical-ligatures contextual;") && !html[2].Contains("'liga'"), "all + cntxtAlts: " + html[2]);
        TestReport.Check(html[3].Contains("font-variant-numeric:oldstyle-nums tabular-nums;"), "oldStyle + tabular: " + html[3]);
        TestReport.Check(html[4].Contains("font-feature-settings:'ss02' 1,'ss14' 1;"), "stylistic sets 2 and 14 on, 3 off: " + html[4]);
        TestReport.Check(!html[5].Contains("oldstyle-nums"), "a right-to-left run keeps the browser's shaping (Word ignores numForm there): " + html[5]);
        TestReport.Check(html[6].Contains("font-variant:small-caps;font-variant-ligatures:no-common-ligatures no-contextual;"), "small caps shorthand: the ligature setting is written again after it: " + html[6]);
    }

    /// <summary>w14:reflection as Word 16 writes its "tight reflection, touching" preset: per-word mirror boxes.</summary>
    private static void DocxReflection()
    {
        const string Preset = "w14:blurRad=\"6350\" w14:stA=\"60000\" w14:stPos=\"0\" w14:endA=\"900\" w14:endPos=\"58000\" w14:dist=\"0\" w14:dir=\"5400000\" w14:fadeDir=\"5400000\" w14:sx=\"100000\" w14:sy=\"-100000\" w14:kx=\"0\" w14:ky=\"0\" w14:algn=\"bl\"";
        List<string> html = DocxFidelityTest.RenderAll("reflection-v1",
            "<w:p>" + Run($"<w14:reflection {W14} {Preset}/>", "two words") + "</w:p>"
          + "<w:p>" + Run($"<w:w w:val=\"80\"/><w14:reflection {W14} {Preset.Replace("w14:dist=\"0\"", "w14:dist=\"50800\"")}/>", "scaled") + "</w:p>");
        const string Box = "<span class=\"text-reflection\" style=\"display:inline-block;text-box:trim-both text alphabetic;text-decoration:inherit;-webkit-box-reflect:below 0 linear-gradient(to top,rgba(0,0,0,0.6) 0%,rgba(0,0,0,0.009) 58%)\">";
        TestReport.Check(html[0].Contains(Box + "two</span> " + Box + "words</span>"), "reflection: one mirror box per word, the space between them: " + html[0]);
        TestReport.Check(html[1].Contains("scaleX(0.8);text-box:trim-both text alphabetic;text-decoration:inherit;-webkit-box-reflect:below 4pt linear-gradient(") && html[1].Contains("class=\"char-scale\""), "reflection + scale: on the scaled word box, dist 4pt: " + html[1]);
        var plain = new System.Text.StringBuilder();
        WordHtmlText.AppendScaled(plain, "a<b>x y</b>", 100, "below 0");
        TestReport.Check(plain.ToString() == "<span class=\"text-reflection\" style=\"display:inline-block;text-box:trim-both text alphabetic;text-decoration:inherit;-webkit-box-reflect:below 0\">a<b>x y</b></span>",
                         "reflection: a space inside an open tag does not split the box: " + plain);
    }
    /// <summary>Word's shaping per script (Word 16 XPS of test-data/word-text-effects/w14-scripts-v1.docx: the same glyphs with
    /// w14:ligatures none and all + cntxtAlts): Hebrew without ligatures, Arabic with them; .doc and RTF get the same default.</summary>
    private static void WordScriptShaping()
    {
        var html = new System.Text.StringBuilder();
        string arabic = "\u0627\u0644\u0644\u0647 \u0628\u0633\u0645";
        TestReport.Check(WordHtmlText.WithWordScriptShaping("abc \u05D0") == "abc \u05D0", "no Arabic letter: text unchanged");
        const string Normal = WordHtmlText.ArabicShapingSpan;
        TestReport.Check(WordHtmlText.WithWordScriptShaping("a " + arabic + " b") == "a " + Normal + arabic.Replace(" ", "</span> " + Normal) + "</span> b",
                         "each Arabic word gets the browser's ligatures back, one span per word");
        html.Clear();
        WordHtmlText.AppendScaled(html, WordHtmlText.WithWordScriptShaping(arabic), 80);
        TestReport.Check(html.ToString().Split("class=\"char-scale\"").Length - 1 == 2, "scaled Arabic: still one box per word: " + html);
        TestReport.Check(WordHtmlText.WithWordScriptShaping("<i title=\"\u0627\">x</i>") == "<i title=\"\u0627\">x</i>", "Arabic inside a tag is left alone");
        List<string> docx = DocxFidelityTest.RenderAll("script-shaping-v2",
            "<w:p>" + Run($"<w:rtl/><w14:ligatures {W14} w14:val=\"all\"/><w14:cntxtAlts {W14}/>", "\u05D5\u05C2") + "</w:p>"
          + "<w:p>" + Run("<w:rtl/>", arabic) + "</w:p>"
          + "<w:p>" + Run($"<w14:ligatures {W14} w14:val=\"all\"/><w14:cntxtAlts {W14}/>", "fi וׂ") + "</w:p>"
          + "<w:p>" + Run($"<w14:ligatures {W14} w14:val=\"standard\"/>", "fi " + arabic) + "</w:p>"
          + "<w:p>" + Run("", "fi וׂ") + "</w:p>");
        TestReport.Check(!docx[0].Contains("common-ligatures discretionary") && !docx[0].Contains(" contextual;"), "Hebrew run: w14:ligatures all + cntxtAlts ignored, as Word does: " + docx[0]);
        TestReport.Check(docx[1].Contains(Normal + arabic.Substring(0, 4) + "</span>"), "Arabic run keeps its ligatures");
        TestReport.Check(docx[2].Contains("common-ligatures discretionary") && docx[2].Contains(WordHtmlText.HebrewShapingSpan + "וׂ</span>") && !docx[2].Contains(WordHtmlText.HebrewShapingSpan + "fi"),
                         "Hebrew letters in a Latin run (no w:rtl) with ligatures all: by their script, without ligatures; the Latin keeps them");
        TestReport.Check(docx[3].Contains("'liga' 1,'clig' 0") && docx[3].Contains(Normal + arabic.Substring(0, 4) + "</span>"),
                         "Arabic letters in a Latin run with liga apart from clig: the feature list is set back, ligatures kept: " + docx[3]);
        TestReport.Check(!docx[4].Contains(WordHtmlText.HebrewShapingSpan), "Hebrew letters in a Latin run with Word's default: no extra span: " + docx[4]);
        TestReport.Check(WordHtmlText.WithWordScriptShaping("א x", isHebrewReset: true) == WordHtmlText.HebrewShapingSpan + "א</span> x", "Hebrew stretch: one span per word");
        string rtfPath = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-shaping.rtf");
        File.WriteAllText(rtfPath, @"{\rtf1\ansi{\fonttbl{\f0 Arial;}}\pard\f0 office\par}");
        using IWordWindowedDocument rtf = WordWindowedDocumentOpener.Open(rtfPath);
        TestReport.Check(rtf.Stylesheet.Contains("font-variant-ligatures:no-common-ligatures no-contextual;"), "RTF: Word's no-ligature default in the document style sheet");
        string docPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-text-effects", "text-effects-v1.doc"));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-text-effects-doc") };
        using IWordWindowedDocument doc = WordWindowedDocumentOpener.Open(docPath, policy);
        TestReport.Check(doc.Stylesheet.Contains("font-variant-ligatures:no-common-ligatures no-contextual;"), ".doc: Word's no-ligature default in the document style sheet");
    }
}
