using System.Xml.Linq;
using FastDocView.DocxHtml;

namespace FastDocView.Tests;

public static partial class DocxFidelityTest
{
    private static DocxCompatibilitySettings Compat(string compatChildren)
        => new(XDocument.Parse($"<w:settings {WordNamespaces}><w:compat>{compatChildren}</w:compat></w:settings>"));

    private const string WordUri = "http://schemas.microsoft.com/office/word";

    /// <summary>settings.xml w:compat read like LibreOffice SettingsTable (GetWordCompatibilityMode: highest of the repeated values above 10; absent = 12).</summary>
    private static void CompatibilitySettings()
    {
        TestReport.Check(DocxCompatibilitySettings.Default.CompatibilityMode == 12, "no compatibilityMode → 12");
        TestReport.Check(Compat($"<w:compatSetting w:name=\"compatibilityMode\" w:uri=\"{WordUri}\" w:val=\"14\"/><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"{WordUri}\" w:val=\"15\"/><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"{WordUri}\" w:val=\"14\"/>").CompatibilityMode == 15,
                         "repeated compatibilityMode: the highest wins");
        TestReport.Check(Compat($"<w:compatSetting w:name=\"compatibilityMode\" w:uri=\"{WordUri}\" w:val=\"10\"/>").CompatibilityMode == 12, "compatibilityMode 10 is not valid (11 is the first)");
        DocxCompatibilitySettings options = Compat($"<w:ulTrailSpace/><w:doNotExpandShiftReturn/><w:doNotUseHTMLParagraphAutoSpacing/><w:compatSetting w:name=\"overrideTableStyleFontSizeAndJustification\" w:uri=\"{WordUri}\" w:val=\"1\"/>");
        TestReport.Check(options.UnderlinesTrailingSpaces && options.DoesNotExpandShiftReturn && options.DoesNotUseHtmlParagraphAutoSpacing && options.OverridesTableStyleFontSizeAndJustification,
                         "compat on/off options and overrideTableStyleFontSizeAndJustification read");
        List<string> html = RenderAll("compat-ultrailspace", "<w:p><w:r><w:rPr><w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\">abc   </w:t></w:r></w:p>",
                                      settings: "<w:compat><w:ulTrailSpace/></w:compat>");
        TestReport.Check(html[0].Contains("<span style=\"text-decoration:underline;\">abc   </span>"), "ulTrailSpace: the spaces ending the paragraph are underlined too");
    }
}
