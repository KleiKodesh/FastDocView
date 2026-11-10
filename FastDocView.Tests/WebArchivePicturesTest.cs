using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Pictures in a Word web archive (test-data/word-mhtml/pictures-v1.mht, saved by Word 16 as "Single File Web Page" with a
/// 40x20 PNG and a 30x30 JPEG): each &lt;img&gt; is drawn once (the VML copy is skipped) with its size, and reads back from its
/// archive part by the src - PNG and JPEG bytes with the right type; an unknown id reads nothing.
/// </summary>
public static class WebArchivePicturesTest
{
    private static readonly Regex Picture = new("<img[^>]*data-word-picture=\"([^\"]+)\"[^>]*width=\"([0-9]+)\" height=\"([0-9]+)\"", RegexOptions.CultureInvariant);

    public static void Run()
    {
        TestReport.Step("Web archive pictures (Word 16 .mht)");
        string path = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-mhtml", "pictures-v1.mht");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-mhtml/pictures-v1.mht present"); return; }
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        var found = new List<(string id, string width, string height)>();
        for (int first = 0; first < document.BlockCount; first += 100)
            foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, 100))
            {
                foreach (Match match in Picture.Matches(block.Html)) found.Add((match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value));
                TestReport.Check(!block.Html.Contains("[image]", StringComparison.Ordinal), $"block {block.BlockIndex}: no [image] placeholder");
            }
        TestReport.Check(found.Count == 2, $"2 pictures drawn once each ({found.Count})");
        if (found.Count != 2) return;
        TestReport.Check(found[0].width == "40" && found[0].height == "20" && found[1].width == "30" && found[1].height == "30", "sizes from the <img> width / height");
        WordDocumentPicture? png = document.ReadPicture(found[0].id), jpeg = document.ReadPicture(found[1].id);
        TestReport.Check(png is { ContentType: "image/png" } && png.Bytes.Length > 8 && png.Bytes[0] == 0x89 && png.Bytes[1] == 0x50, $"first picture reads back as PNG ({png?.Bytes.Length ?? 0} bytes)");
        TestReport.Check(jpeg is { ContentType: "image/jpeg" } && jpeg.Bytes.Length > 3 && jpeg.Bytes[0] == 0xFF && jpeg.Bytes[1] == 0xD8, $"second picture reads back as JPEG ({jpeg?.Bytes.Length ?? 0} bytes)");
        TestReport.Check(document.ReadPicture("mh:missing.png") is null && document.ReadPicture("s1") is null, "an unknown id reads nothing");
    }
}
