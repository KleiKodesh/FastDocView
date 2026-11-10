using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Word 6 / 95 list levels: test-data/word-lists/legacy-lists-v1.docx / .doc / .rtf (make-legacy-lists.py + make-legacy-lists-fixture.ps1:
/// w:legacy levels Word 16 re-saved as .docx, .doc - LVLF flag 0x40 with dxaSpace / dxaIndent - and .rtf - \levelold \levelspace
/// \levelindent). Word 16 puts the text at the furthest of the start indent, label start + legacyIndent and label end + legacySpace,
/// with no tab (legacy-lists-v1.expected.txt: e.g. left 1440 hanging 720, legacyIndent 1000 = text 50 pt after the label start;
/// legacySpace 1000 = 50 pt after the label's end). The label box reaches max(hanging, legacyIndent) and its text is padded by legacySpace.
/// </summary>
public static class LegacyListsFidelityTest
{
    // (legacySpace, legacyIndent) per level, as make-legacy-lists.py writes them; null = an ordinary level
    private static readonly (int space, int indent)?[] Variants = [null, (0, 0), (0, 1000), (1000, 0), (100, 1000), (0, 200), (700, 0), (0, 1500)];
    private static readonly int[] Hangings = [720, 0];

    public static void Run()
    {
        TestReport.Step("Lists: Word 6 levels (w:legacy, .doc LVLF fWord6, RTF levelold) - label reach and spacing as Word 16 draws them");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-lists");
        foreach (string extension in new[] { ".docx", ".doc", ".rtf" })
        {
            string path = Path.GetFullPath(Path.Combine(folder, "legacy-lists-v1" + extension));
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/word-lists/legacy-lists-v1{extension} present"); continue; }
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-legacy-lists-" + extension.TrimStart('.')) };
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            List<string> markers = document.RenderBlocks(0, document.BlockCount).Select(block => block.Html)
                .Where(html => WordFieldsFidelityTest.Strip(html).Contains("legacy ")).Select(Marker).ToList();
            TestReport.Check(markers.Count == Variants.Length * Hangings.Length * 2, $"{extension}: {Variants.Length * Hangings.Length * 2} list paragraphs (got {markers.Count})");
            if (markers.Count != Variants.Length * Hangings.Length * 2) continue;
            for (int h = 0; h < Hangings.Length; h++)
                for (int v = 0; v < Variants.Length; v++)
                {
                    string marker = markers[(h * Variants.Length + v) * 2];
                    if (Variants[v] is not { } legacy) { TestReport.Check(!marker.Contains("<span style=\"padding-inline-end:") || marker.Contains("0.25em"), $"{extension}: ordinary level keeps its tab (got [{marker}])"); continue; }
                    int reach = Math.Max(Hangings[h], legacy.indent);
                    bool reachOk = reach == 0 ? !marker.Contains("min-width:") : marker.Contains("min-width:" + Pt(reach) + "pt;");
                    bool spaceOk = legacy.space == 0 ? !marker.Contains("padding-inline-end:") : marker.Contains("padding-inline-end:" + Pt(legacy.space) + "pt");
                    TestReport.Check(reachOk && spaceOk && !marker.Contains("0.25em"),
                                     $"{extension}: hanging {Hangings[h]}, legacySpace {legacy.space} legacyIndent {legacy.indent}: label reaches {Pt(reach)} pt, then {Pt(legacy.space)} pt (got [{marker}])");
                }
        }
    }

    private static string Pt(int twips) => (twips / 20.0).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static string Marker(string html)
    {
        Match m = Regex.Match(html, "<span class=\"list-marker\"[^>]*>(<span[^>]*>)?");
        return m.Success ? m.Value : "";
    }
}
