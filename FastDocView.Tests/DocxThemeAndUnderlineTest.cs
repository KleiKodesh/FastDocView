using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Theme fonts and colors (theme1.xml + settings.xml clrSchemeMapping / themeFontLang), underline
/// patterns and colors, emphasis marks — checked on the rendered CSS of the generator's sample
/// paragraphs. Expected values follow ECMA-376's own worked tint / shade examples (per-component
/// RGB: tint C' = 255 − (255 − C)·T/255, shade C' = C·S/255).
/// </summary>
public static class DocxThemeAndUnderlineTest
{
    public static void Run()
    {
        TestReport.Step("Theme fonts / colors, underline styles, emphasis marks (synthetic docx with theme1.xml)");
        string path = Path.Combine(TestReport.OutputDirectory, "synthetic-theme-v1-120.docx");
        SyntheticDocxGenerator.Write(path, 120, withThemeSamples: true);
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);

        // cstheme="minorBidi" + themeFontLang bidi=he-IL → the minor font's "Hebr" face (Arial), not the explicit cs="David" and not the a:cs (empty)
        string themeFont = Html(document, SyntheticDocxGenerator.ThemeFontParagraph);
        TestReport.Check(themeFont.Contains("font-family:'Arial'"), "cstheme=minorBidi resolves to the theme's Hebrew face (Arial)");
        TestReport.Check(Regex.IsMatch(themeFont, "<span style=\"[^\"]*font-family:'Arial';\">") && !Regex.IsMatch(themeFont, "<span[^>]*David"), "the theme attribute wins over the explicit cs face (the w:rtl run takes szCs and the theme's Hebrew face)");

        // accent1 4472C4 with tint 99 → 8FAADC (val FF0000 is ignored); shd themeFill accent6 70AD47 shade 80 → 385724
        string themeColor = Html(document, SyntheticDocxGenerator.ThemeColorParagraph);
        TestReport.Check(themeColor.Contains("color:#8FAADC;"), "themeColor accent1 + themeTint 99 → #8FAADC");
        TestReport.Check(!themeColor.Contains("#FF0000"), "explicit val / fill ignored when the theme resolves");
        TestReport.Check(themeColor.Contains("background-color:#385724;"), "shd themeFill accent6 + themeFillShade 80 → #385724");

        // underline: double + explicit color; "none" stays off; wavyHeavy + themeColor text2 (→ t2 → dark2 44546A)
        string underline = Html(document, SyntheticDocxGenerator.UnderlineParagraph);
        TestReport.Check(underline.Contains("text-decoration:underline;text-decoration-style:double;text-decoration-color:#00B050;"), "u double + color → text-decoration-style / -color");
        TestReport.Check(underline.Contains("</span> ללא קו<span"), "u none renders as plain text (no underline span)");
        TestReport.Check(underline.Contains("text-decoration-style:wavy;text-decoration-thickness:0.15em;text-decoration-color:#44546A;"), "u wavyHeavy + themeColor text2 → wavy, thick, #44546A");

        // paragraph mark underlined (single) + run dotted: the run must re-declare the line, or the style has nothing to apply to
        string nested = Html(document, SyntheticDocxGenerator.UnderlinedParagraphMarkParagraph);
        TestReport.Check(nested.Contains("<span style=\"text-decoration:underline;text-decoration-style:dotted;\">"), "a differently patterned underline inside an underlined paragraph re-declares text-decoration");

        // emphasis marks
        string emphasis = Html(document, SyntheticDocxGenerator.EmphasisParagraph);
        TestReport.Check(emphasis.Contains("text-emphasis:filled dot;text-emphasis-position:under right;"), "em underDot → text-emphasis filled dot, under");
        TestReport.Check(emphasis.Contains("text-emphasis:filled sesame;") && !emphasis.Contains("sesame;text-emphasis-position"), "em comma → filled sesame above");

        // clrSchemeMapping: background1 → bg1 → light1 FFFFFF shade 80 → 808080; accent2 remapped to accent6 (70AD47) on a paragraph border
        string mapped = Html(document, SyntheticDocxGenerator.MappedColorParagraph);
        TestReport.Check(mapped.Contains("color:#808080;"), "themeColor background1 + themeShade 80 → #808080 through clrSchemeMapping");
        TestReport.Check(mapped.Contains("#70AD47"), "pBdr themeColor accent2 follows clrSchemeMapping accent2=accent6");

        // a document without a theme part still renders explicit colors and fonts unchanged
        string plainPath = Path.Combine(TestReport.OutputDirectory, "synthetic-theme-v1-none-60.docx");
        SyntheticDocxGenerator.Write(plainPath, 60);
        using IWordWindowedDocument plain = WordWindowedDocumentOpener.Open(plainPath, policy);
        TestReport.Check(plain.RenderBlocks(0, 1)[0].Html.Contains("color:#2F5496;"), "no theme part: explicit style color kept");
    }

    private static string Html(IWordWindowedDocument document, int blockIndex) => document.RenderBlocks(blockIndex, 1)[0].Html;
}
