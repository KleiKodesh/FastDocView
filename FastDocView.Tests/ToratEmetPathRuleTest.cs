using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>The path rule: a .txt under a folder whose name holds "ToratEmet" is a Torat Emet book with no detection; the index remembers it.</summary>
public static class ToratEmetPathRuleTest
{
    public static void Run()
    {
        TestReport.Step("Torat Emet path rule");
        TestReport.Check(ToratEmetPathRule.Applies(@"C:\x\ToratEmetInstall\Books\a.txt") && ToratEmetPathRule.Applies(@"C:\x\TORATEMET\a.txt")
                         && ToratEmetPathRule.Applies(@"C:\x\mytoratemetfolder\sub\a.txt") && !ToratEmetPathRule.Applies(@"C:\x\Torat\a.txt") && !ToratEmetPathRule.Applies(null),
                         "the rule: any level, any case, nothing else");
        string root = Path.Combine(TestReport.OutputDirectory, "torat-path-rule-v1");
        string under = Path.Combine(root, "ToratEmetUserData", "MyBooks");
        string elsewhere = Path.Combine(root, "plain-folder");
        Directory.CreateDirectory(under); Directory.CreateDirectory(elsewhere);
        Encoding ansi = MarkupTextEncoding.FromCodePage(MarkupTextEncoding.HebrewWindows);

        string markedText = "$ Title line\r\n# Hash\r\n! text <b>bold</b><span>s</span>\r\nplain line\r\n";   // no &Key= line: detection says no
        string noMarksText = "first line\r\n<span>x</span> second <b>b</b>\r\nthird\r\n";                      // no heading mark at all
        string keyLineText = "&UniqueId=5&ForcedBookName=Forced\r\nbody line\r\n";                           // line 1 is metadata, no heading
        var cases = new (string name, string text)[] { ("marked", markedText), ("nomarks", noMarksText), ("keyline", keyLineText) };
        foreach ((string name, string text) in cases)
        {
            string inRule = Path.Combine(under, name + ".txt"), outside = Path.Combine(elsewhere, name + ".txt");
            File.WriteAllBytes(inRule, ansi.GetBytes(text)); File.WriteAllBytes(outside, ansi.GetBytes(text));
            if (name != "keyline")
            {
                using var plain = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(outside, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
                TestReport.Check(plain.ToratEmet is null && plain.Headings.Count == 0, $"{name}: the same text elsewhere stays a plain .txt");
            }
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(root, "index-cache"), ForcedMode = mode };
                string indexPath = policy.IndexPathFor(inRule);
                if (File.Exists(indexPath)) File.Delete(indexPath);
                for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
                {
                    string what = $"{name} {mode}{(pass == 1 ? " reopened from the index" : "")}";
                    using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(inRule, policy);
                    TestReport.Check(document.ToratEmet is not null, $"{what}: parsed as Torat Emet without detection");
                    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
                    if (name == "marked")
                    {
                        TestReport.Check(document.BlockCount == 4 && blocks[0].Html == "<div class=\"text-line te-heading\"><h2>Title line</h2></div>"
                                         && blocks[1].Html.Contains("<h3>Hash</h3>") && blocks[2].Html == "<div class=\"text-line\">text <b>bold</b>s</div>",
                                         $"{what}: block 0 is the first line, marks become headings and are hidden ({blocks[0].Html})");
                        TestReport.Check(string.Join(",", document.Headings.Select(h => h.Level + ":" + h.BlockIndex)) == "2:0,3:1", $"{what}: heading levels");
                    }
                    else if (name == "nomarks")
                    {
                        TestReport.Check(document.BlockCount == 3 && blocks[0].Html == "<div class=\"text-line\">first line</div>"
                                         && blocks[1].Html == "<div class=\"text-line\">x second <b>b</b></div>" && document.Headings.Count == 0,
                                         $"{what}: no marks: plain lines drawn by the Torat Emet renderer ({blocks[1].Html})");
                    }
                    else
                    {
                        TestReport.Check(document.DocumentTitle == "Forced" && document.BlockCount == 2 && blocks[0].Html.Contains("Forced") && blocks[1].Html.Contains("body line"),
                                         $"{what}: line 1 is the head, no heading needed");
                    }
                    TestReport.Check(name == "keyline" || document.Search("line", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.Length > 0, $"{what}: search sees the first line");
                }
            }
        }
    }
}
