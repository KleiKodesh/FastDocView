using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FastDocView.DocxHtml;

namespace FastDocView.Tests;

/// <summary>
/// SmartArt laid out from its data model and layout definition (--fidelity, seconds): test-data/smartart-layouts/ holds one
/// Word 16 .docx per built-in layout (scripts/make-smartart-layouts.ps1, with Word's PDF of each) and a copy without its drawing
/// part (scripts/strip-smartart-drawing.py). Word's saved drawing part is what Word prints, so the shapes DocxDiagramLayout
/// lays out from the original's data, layout, colors and quickStyle parts are matched to the drawing's shapes by presentation
/// point (dsp:sp modelId): each layout listed in <see cref="Layouts"/> must place every shape within the tolerance; the rest
/// are reported only. The stripped copy must render shapes through the document path (no empty frame) and its text be found by
/// search.
/// </summary>
public static class SmartArtLayoutFidelityTest
{
    /// <summary>The layouts the engine matches Word for (box edges within <see cref="TolerancePoints"/>); the others are reported, not checked.</summary>
    private static readonly string[] Layouts = ["block-list", "process", "pyramid", "vertical-list", "chevron", "cycle", "hierarchy", "org-chart"];
    private const double TolerancePoints = 4;
    private const double EmuPerPoint = 12700;

    public static void Run()
    {
        TestReport.Step("SmartArt without a drawing part: laid out from data + layout, matched to Word's drawing");
        string folder = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data", "smartart-layouts"));
        if (!Directory.Exists(folder)) { TestReport.Check(false, "test-data/smartart-layouts present"); return; }
        foreach (string path in Directory.GetFiles(folder, "*-v1.docx").OrderBy(p => p, StringComparer.Ordinal))
        {
            string name = Path.GetFileNameWithoutExtension(path).Replace("-v1", "");
            (double worst, int matched, int expected, string detail) = Compare(path);
            bool required = Layouts.Contains(name);
            string summary = $"{name}: {matched} of {expected} shapes laid out, worst edge {worst.ToString("0.0", CultureInfo.InvariantCulture)} pt";
            if (required) TestReport.Check(matched == expected && worst <= TolerancePoints, summary + (matched == expected && worst <= TolerancePoints ? "" : " " + detail));
            else Console.WriteLine("   INFO  " + summary + " (not yet matched: " + detail + ")");
            string stripped = Path.Combine(folder, Path.GetFileNameWithoutExtension(path) + "-nodrawing.docx");
            if (required && File.Exists(stripped)) CheckDocument(stripped, name);
        }
    }

    /// <summary>The laid-out shapes against the drawing part's, by modelId: the largest difference of any box edge, in points.</summary>
    private static (double Worst, int Matched, int Expected, string Detail) Compare(string path)
    {
        using ZipArchive zip = ZipFile.OpenRead(path);
        XElement Part(string name) => DocxShapeGeometry.WithoutNamespaces(XDocument.Load(zip.GetEntry(name)!.Open()).Root!);
        XElement document = XDocument.Load(zip.GetEntry("word/document.xml")!.Open()).Root!;
        XElement extent = document.Descendants().First(e => e.Name.LocalName == "extent");
        double width = double.Parse((string)extent.Attribute("cx")!, CultureInfo.InvariantCulture), height = double.Parse((string)extent.Attribute("cy")!, CultureInfo.InvariantCulture);
        XElement? laidOut = DocxDiagramLayout.Lay(Part("word/diagrams/data1.xml"), Part("word/diagrams/layout1.xml"), Part("word/diagrams/colors1.xml"), Part("word/diagrams/quickStyle1.xml"), width, height);
        var ours = Boxes(laidOut);
        var words = Boxes(Part("word/diagrams/drawing1.xml").Element("spTree"));
        double worst = 0;
        int matched = 0;
        var detail = new List<string>();
        foreach ((string id, (double X, double Y, double W, double H) word) in words)
        {
            if (!ours.TryGetValue(id, out var mine)) { detail.Add("missing " + Shorten(id)); continue; }
            matched++;
            double difference = new[] { Math.Abs(mine.X - word.X), Math.Abs(mine.Y - word.Y), Math.Abs(mine.X + mine.W - word.X - word.W), Math.Abs(mine.Y + mine.H - word.Y - word.H) }.Max();
            if (difference > TolerancePoints) detail.Add($"{Shorten(id)} ours {Box(mine)} word {Box(word)}");
            worst = Math.Max(worst, difference);
        }
        return (worst, matched, words.Count, string.Join("; ", detail.Take(4)));
    }

    private static Dictionary<string, (double X, double Y, double W, double H)> Boxes(XElement? tree)
    {
        var boxes = new Dictionary<string, (double, double, double, double)>();
        foreach (XElement shape in tree?.Elements("sp") ?? [])
        {
            // Word's empty helper shapes (connectors drawn as custom geometry, no fill) are compared only when they have a preset
            XElement? transform = shape.Element("spPr")?.Element("xfrm");
            if (transform is null || shape.Element("spPr")?.Element("prstGeom") is null) continue;
            double Value(string element, string attribute) => double.Parse((string?)transform.Element(element)?.Attribute(attribute) ?? "0", CultureInfo.InvariantCulture) / EmuPerPoint;
            boxes[(string?)shape.Attribute("modelId") ?? ""] = (Value("off", "x"), Value("off", "y"), Value("ext", "cx"), Value("ext", "cy"));
        }
        return boxes;
    }

    /// <summary>The copy without a drawing part through the document: shapes drawn in the frame, the node texts shown and searchable.</summary>
    private static void CheckDocument(string path, string name)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-smartart-" + name) };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        int shapes = Regex.Matches(html, "<path ").Count;
        TestReport.Check(shapes > 0, $"{name} without its drawing part: {shapes} shape paths drawn in the frame");
        // search (the one text path, GetText too) finds the node texts once, as with a drawing part
        using ZipArchive zip = ZipFile.OpenRead(path);
        XElement data = DocxShapeGeometry.WithoutNamespaces(XDocument.Load(zip.GetEntry("word/diagrams/data1.xml")!.Open()).Root!);
        foreach (string text in data.Element("ptLst")!.Elements("pt").Where(p => p.Attribute("type") is null).SelectMany(p => p.Descendants("t")).Select(t => t.Value).Where(t => t.Length > 0).Distinct().Take(3))
        {
            int hits = 0;
            for (int from = 0; from < document.BlockCount;)
            {
                WordDocumentSearchBatch batch = document.Search(text, WordDocumentSearchOptions.Default, from, 4);
                hits += batch.HitCounts.Sum();
                if (batch.ReachedEnd) break;
                from = batch.NextBlock;
            }
            TestReport.Check(hits == 1, $"{name} without its drawing part: search finds \"{text}\" once ({hits})");
        }
    }

    private static string Box((double X, double Y, double W, double H) box) => $"({box.X:0.0},{box.Y:0.0} {box.W:0.0}x{box.H:0.0})";
    private static string Shorten(string id) => id.Length > 9 ? id.Substring(1, 8) : id;
}
