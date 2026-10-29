using System.IO.Compression;
using System.Text;

namespace FastDocView.Tests;

/// <summary>
/// List label layout in RTF and OpenDocument (found by the LibreOffice PDF comparison, NOTES.md): RTF \listtext keeps its
/// tab as the tab suffix and \leveljc places the label (LibreOffice reads \leveljc, not \leveljcn); ODF label-alignment
/// levels indent the paragraph with fo:margin-left / fo:text-indent and the text after a listtab starts at
/// text:list-tab-stop-position.
/// </summary>
public static class ListLabelFormatsTest
{
    public static void Run()
    {
        TestReport.Step("List label layout: RTF \\listtext / \\leveljc, ODF label-alignment levels");
        string rtfPath = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-rtf-list-layout.rtf");
        File.WriteAllText(rtfPath, @"{\rtf1\ansi{\fonttbl{\f0 Arial;}}
{\*\listtable{\list\listtemplateid1{\listlevel\levelnfc0\leveljc2\leveljcn0\levelfollow0\levelstartat1{\leveltext\'02\'00.;}{\levelnumbers\'01;}\fi-360\li720}\listid7}}
{\*\listoverridetable{\listoverride\listid7\listoverridecount0\ls1}}
{\listtext\pard\plain 1.\tab}\pard\fi-360\li720\ls1 first\par
}");
        using (IWordWindowedDocument rtf = WordWindowedDocumentOpener.Open(rtfPath))
        {
            string html = rtf.RenderBlocks(0, 1)[0].Html;
            TestReport.Check(html.Contains("margin-inline-start:36pt;") && html.Contains("min-width:18pt;"), @"rtf: \listtext tab reaches the hanging indent (" + html + ")");
            TestReport.Check(html.Contains("justify-content:flex-end") && !html.Contains("1.\t"), @"rtf: \leveljc2 ends the label at its position, the tab is not drawn as text");
        }

        string odtPath = Path.Combine(TestReport.OutputDirectory, "fidelity-v1-odt-list-layout.odt");
        if (File.Exists(odtPath)) File.Delete(odtPath);
        using (var zip = ZipFile.Open(odtPath, ZipArchiveMode.Create))
        {
            void Add(string name, string text, CompressionLevel level = CompressionLevel.Optimal)
            {
                using var writer = new StreamWriter(zip.CreateEntry(name, level).Open(), new UTF8Encoding(false));
                writer.Write(text);
            }
            Add("mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
            Add("content.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
                  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3">
                  <office:automatic-styles>
                    <text:list-style style:name="L1">
                      <text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1">
                        <style:list-level-properties text:list-level-position-and-space-mode="label-alignment">
                          <style:list-level-label-alignment text:label-followed-by="listtab" text:list-tab-stop-position="1in" fo:text-indent="-0.25in" fo:margin-left="1in"/>
                        </style:list-level-properties>
                      </text:list-level-style-number>
                    </text:list-style>
                  </office:automatic-styles>
                  <office:body><office:text><text:list text:style-name="L1"><text:list-item><text:p>first</text:p></text:list-item></text:list></office:text></office:body>
                </office:document-content>
                """);
        }
        using (IWordWindowedDocument odt = WordWindowedDocumentOpener.Open(odtPath))
        {
            while (!odt.IsIndexComplete) odt.ScanAhead(1000);
            string html = odt.RenderBlocks(0, 1)[0].Html;
            TestReport.Check(html.Contains("margin-inline-start:72pt;") && html.Contains("text-indent:-18pt;") && html.Contains("min-width:18pt;"),
                             "odt: label-alignment level indents the paragraph, the listtab reaches the tab stop (" + html + ")");
        }
    }
}
