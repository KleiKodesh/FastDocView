using System.Text.RegularExpressions;

namespace FastDocView.Tests;

/// <summary>
/// Floating .doc pictures (0x08 + PlcSpaMom FSPA + OfficeArt drawing, DocBinaryFloatingPictures) on Apache POI's test
/// files: each named picture is drawn with its FSPA size and wrap and reads back as a PNG; inline pictures keep their ids.
/// </summary>
public static class DocBinaryFloatingPicturesTest
{
    private static readonly Regex Picture = new("<img[^>]*data-word-picture=\"(s[0-9]+)\"[^>]*style=\"([^\"]*)\"", RegexOptions.CultureInvariant);

    public static void Run()
    {
        TestReport.Step(".doc floating pictures (0x08 / FSPA / OfficeArt) on test-data/apache-poi");
        string folder = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "apache-poi");
        // file -> (floating pictures expected, a style fragment one of them must carry)
        (string file, int count, string style)[] cases =
        [
            ("PngPicture.doc", 1, "width:141.7pt"),
            ("FloatingPictures.doc", 4, "display:block"),      // two text-above-and-below, two square
            ("pictures_escher.doc", 2, "width:179.95pt"),
        ];
        foreach ((string file, int count, string style) in cases)
        {
            string path = Path.Combine(folder, file);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data/apache-poi/{file} present"); continue; }
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            var found = new List<(string id, string css)>();
            for (int first = 0; first < document.BlockCount; first += 100)
                foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, 100))
                    foreach (Match match in Picture.Matches(block.Html)) found.Add((match.Groups[1].Value, match.Groups[2].Value));
            TestReport.Check(found.Count == count, $"{file}: {found.Count} floating picture(s), expected {count}");
            TestReport.Check(found.Exists(picture => picture.css.Contains(style, StringComparison.Ordinal)), $"{file}: a picture styled '{style}'");
            foreach ((string id, _) in found)
            {
                WordDocumentPicture? picture = document.ReadPicture(id);
                bool isPng = picture is { ContentType: "image/png", Bytes.Length: > 8 } && picture.Bytes[0] == 0x89 && picture.Bytes[1] == 0x50;
                TestReport.Check(isPng, $"{file}: picture {id} reads back as PNG ({picture?.Bytes.Length ?? 0} bytes)");
            }
            TestReport.Check(document.ReadPicture("s999999") is null && document.ReadPicture("s") is null, $"{file}: an unknown shape id reads nothing");
        }
    }
}
