using System.Globalization;
using System.Text.RegularExpressions;
using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Picture effects (.docx) and shapes in web archives (.mht), against Word 16's PDFs in test-data/word-picture-effects
/// (make-effects-fixture.ps1 + make_variants.py: swatch-v1.png with Word's own effects - grey, black and white, washout,
/// brightness / contrast, transparent colour, artistic effects, soft edges - and XML-only variants: duotone, transparency,
/// thresholds, grey + brightness, colour change, soft edge 20 pt, 3-D rotations). Word's PDF draws each effect into the
/// picture; the expected pixel values below are read from those images (2026-11-17).
/// </summary>
public static class PictureEffectsFidelityTest
{
    private static string Folder => Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "word-picture-effects");

    public static void Run()
    {
        TestReport.Step("Picture effects: .docx against Word's PDF; shapes in web archives");
        Docx();
        Cameras();
        WebArchive();
    }

    /// <summary>
    /// Soft edges, duotone clipping and perspective cameras against Word's PDF of effects-3d-v1.docx (make_3d.py). Word's
    /// numbers: soft edge 10 pt half transparent 8.45 pt in (alpha 0.47 at 8 pt); duotone 202020 / C0504D satMod 300 % mixes to
    /// 255, 0, 0 (a saturation clipped at 100 % gives 255, 15, 15); perspectiveFront / perspectiveRelaxed draw at a camera
    /// 1557 pt away for fov 45 (645 pt for fov 90) - the browser's projection of these was measured against Word's: within
    /// 0.5 pt (fov 90: 1.5 pt), the same edge nearer.
    /// </summary>
    private static void Cameras()
    {
        string path = Path.Combine(Folder, "effects-3d-v1.docx");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-picture-effects/effects-3d-v1.docx present"); return; }
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        string[] pictures = Regex.Split(html, "(?=<svg class=\"word-picture-effect\")|(?<=<img[^>]*>)").Where(part => part.Contains("<img")).ToArray();
        TestReport.Check(pictures.Length == 16, $"16 pictures ({pictures.Length})");
        if (pictures.Length != 16) return;
        Match stop = Regex.Match(pictures[1], @"rgba\(0,0,0,([0-9.]+)\) 8pt");
        TestReport.Check(stop.Success && Math.Abs(double.Parse(stop.Groups[1].Value, CultureInfo.InvariantCulture) - 0.47) < 0.03, $"soft edge 10 pt: alpha at 8 pt {stop.Groups[1].Value} (Word 0.47)");
        TestReport.Check(pictures[5].Contains("-0.0267 -0.0898 -0.0091 0 0.1255"), "duotone with satMod 300 %: Word's clipping order (green / blue mix to 0)");
        TestReport.Check(pictures[8].Contains("perspective(1557.1677pt) rotateX(0deg) rotateY(-30deg)") && pictures[10].Contains("perspective(645pt)"),
                         "perspective cameras: fov 45 and 90 at Word's distances");
        TestReport.Check(pictures[11].Contains("transform:rotateX(0deg) rotateY(-40deg) rotateZ(0deg)"), "orthographic camera: no perspective");
    }

    private static void Docx()
    {
        string path = Path.Combine(Folder, "effects-xml-v1.docx");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-picture-effects/effects-xml-v1.docx present"); return; }
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        // one segment per picture: its own svg filter (written just before it) and its <img>
        string[] pictures = Regex.Split(html, "(?=<svg class=\"word-picture-effect\")|(?<=<img[^>]*>)").Where(part => part.Contains("<img")).ToArray();
        TestReport.Check(pictures.Length == 22, $"22 pictures ({pictures.Length})");
        if (pictures.Length != 22) return;
        TestReport.Check(!pictures[0].Contains("filter:url(") && !pictures[0].Contains("word-picture-effect"), "no effect: a plain picture");
        TestReport.Check(pictures[1].Contains("0.2126 0.7152 0.0722 0 0 0.2126") && pictures[1].Contains("filter:url(#"), "grey: Rec. 709 luminance (Word: red 112 -> 23, green 136 -> 97)");
        TestReport.Check(Step(pictures[2]) is { } half && Math.Abs(half - 127.5) < 1, "black and white 50 %: step at luminance 128");
        TestReport.Check(Step(pictures[15]) is { } quarter && Math.Abs(quarter - 63.75) < 1 && Step(pictures[16]) is { } threeQuarters && Math.Abs(threeQuarters - 191.25) < 1,
                         "black and white 25 / 75 %: steps at 64 / 191");
        // brightness / contrast against Word's pixels (washout: 0 -> 205, 112 -> 238; b 40 % c -40 %: 0 -> 132, 112 -> 199; b -40 % c 60 %: 174 -> 65, 249 -> 252)
        foreach ((int index, int input, int word) in new[] { (3, 0, 205), (3, 112, 238), (4, 0, 132), (4, 112, 199), (5, 174, 65), (5, 249, 252) })
            TestReport.Check(Transfer(pictures[index], input) is { } value && Math.Abs(value - word) <= 1, $"brightness / contrast picture {index}: {input} -> {Transfer(pictures[index], input):0.#} (Word {word})");
        TestReport.Check(pictures[6].Contains("flood-opacity=\"0\""), "transparent colour: the exact colour drawn transparent");
        TestReport.Check(pictures[7].Contains("data-word-picture=\"m:rId5\"") && pictures[8].Contains("data-word-picture=\"m:rId7\""),
                         "artistic effects: Word's own drawing of the effect (a:blip r:embed), not the original (a14:imgLayer)");
        TestReport.Check(pictures[10].Contains("mask-image:") && pictures[10].Contains(" 20pt,") && pictures[19].Contains(" 40pt,"), "soft edges 10 / 20 pt: faded over 2 radii");
        TestReport.Check(pictures[12].Contains("<feColorMatrix") && pictures[13].Contains("0.0525 0.1767 0.0178 0 0.7529"),
                         "duotone: luminance mixed from the first colour (C00000) to the second (FFFF00)");
        TestReport.Check(pictures[14].Contains("opacity:0.5;"), "transparency 50 %: opacity");
        TestReport.Check(pictures[17].IndexOf("0.2126 0.7152", StringComparison.Ordinal) is var grey and >= 0
                         && pictures[17].IndexOf("<feComponentTransfer>", StringComparison.Ordinal) > grey, "grey then brightness: document order");
        TestReport.Check(pictures[18].Contains("flood-color=\"rgb(0,0,255)\""), "colour change: green to blue");
        TestReport.Check(pictures[20].Contains("transform:rotateX(0deg) rotateY(-40deg)"), "3-D rotation, orthographic camera: CSS rotation without perspective (Word: 180 pt shown 139 pt wide at 40 degrees)");
        TestReport.Check(pictures[21].Contains("perspective(1557.1677pt) rotateX(-30deg)"), "3-D rotation, perspective camera (perspectiveRelaxed, fov 45): Word's camera distance");
    }

    /// <summary>The black-and-white step (the input luminance 0..255 from which the output is white) of a picture's filter.</summary>
    private static double? Step(string picture)
    {
        Match transfer = Regex.Match(picture, "feFuncR type=\"linear\" slope=\"([^\"]+)\" intercept=\"([^\"]+)\"");
        if (!transfer.Success) return null;
        double slope = double.Parse(transfer.Groups[1].Value, CultureInfo.InvariantCulture), intercept = double.Parse(transfer.Groups[2].Value, CultureInfo.InvariantCulture);
        return -intercept / slope * 255;
    }

    /// <summary>A pixel value 0..255 through the picture's brightness / contrast transfer, clamped as the filter clamps.</summary>
    private static double? Transfer(string picture, int input)
    {
        Match transfer = Regex.Match(picture, "feFuncR type=\"linear\" slope=\"([^\"]+)\" intercept=\"([^\"]+)\"");
        if (!transfer.Success) return null;
        double slope = double.Parse(transfer.Groups[1].Value, CultureInfo.InvariantCulture), intercept = double.Parse(transfer.Groups[2].Value, CultureInfo.InvariantCulture);
        return Math.Max(0, Math.Min(1, slope * input / 255 + intercept)) * 255;
    }

    /// <summary>
    /// A Word web archive with shapes (vml-shapes-v1.mht: rectangle with text, oval, text box, line, a grey picture): a browser has
    /// no VML, so it shows what Word wrote for it inside &lt;![if !vml]&gt; - the shapes laid out in a table (filled cells
    /// holding the shapes' text, pictures of the oval and the line) and the effect drawn into the picture's copy. That is what
    /// is shown; the VML is skipped.
    /// </summary>
    private static void WebArchive()
    {
        string path = Path.Combine(Folder, "vml-shapes-v1.mht");
        if (!File.Exists(path)) { TestReport.Check(false, "test-data/word-picture-effects/vml-shapes-v1.mht present"); return; }
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        TestReport.Check(html.Contains("Rectangle text") && html.Contains("Text box content"), "web archive: the shapes' text is shown");
        TestReport.Check(html.Contains("bgcolor=\"#FF8000\"") || html.Contains("background:#FF8000"), "web archive: the rectangle's fill (Word's fallback cell)");
        TestReport.Check(html.Contains("vml-shapes-v1.files/image001.png") && html.Contains("vml-shapes-v1.files/image002.png"), "web archive: the oval and the line as Word's pictures of them");
        TestReport.Check(html.Contains("vml-shapes-v1.files/image004.jpg") && !html.Contains("vml-shapes-v1.files/image003.png"),
                         "web archive: the grey picture is Word's copy with the effect drawn in (not the VML's original)");
        TestReport.Check(!Regex.IsMatch(html, "<v:(rect|oval|shape|line)\\b"), "web archive: no VML reaches the page");
    }
}
