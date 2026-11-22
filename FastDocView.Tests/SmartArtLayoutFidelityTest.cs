using System.Globalization;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using FastDocView.DocxHtml;
using FastDocView.Html;

namespace FastDocView.Tests;

/// <summary>
/// SmartArt laid out from its data model and layout definition (--fidelity, seconds): test-data/smartart-layouts/ holds one
/// Word 16 .docx per built-in layout (scripts/make-smartart-layouts.ps1, with Word's PDF of each) and a copy without its drawing
/// part (scripts/strip-smartart-drawing.py). Word's saved drawing part is what Word prints, so the shapes DocxDiagramLayout
/// lays out from the original's data, layout, colors and quickStyle parts are matched to the drawing's shapes by presentation
/// point (dsp:sp modelId): each layout listed in <see cref="Layouts"/> must place every shape within the tolerance and give every
/// text Word's autofit font size; the rest
/// are reported only. The stripped copy must render shapes through the document path (no empty frame) and its text be found by
/// search.
/// </summary>
public static class SmartArtLayoutFidelityTest
{
    /// <summary>The layouts the engine matches Word for (box edges within <see cref="TolerancePoints"/>); the others are reported, not checked.</summary>
    private static readonly string[] Layouts = ["block-list", "process", "pyramid", "vertical-list", "chevron", "cycle", "hierarchy", "org-chart", "horizontal-list", "stacked-list", "pyramid-4-levels",
                                                // Hebrew (Arial from the theme's Hebr font, niqqud without width), mixed Hebrew / English, a Times New Roman theme, a theme a:cs of David (Word still uses the Hebr font)
                                                "hebrew-block-list", "hebrew-niqqud-block-list", "hebrew-vertical-list", "hebrew-process", "hebrew-horizontal-list", "mixed-block-list", "times-block-list", "hebrew-david-block-list"];
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
            (double worst, int matched, int expected, string detail, string fonts) = Compare(path);
            // the autofit font size of every text shape equals Word's (font metrics, kerning, text rectangles, op="equ" groups)
            if (Layouts.Contains(name)) TestReport.Check(fonts.Length == 0, $"{name}: font sizes as Word's{(fonts.Length == 0 ? "" : " " + fonts)}");
            else if (fonts.Length > 0) Console.WriteLine($"   INFO  {name}: font sizes differ from Word's: {fonts}");
            bool required = Layouts.Contains(name);
            string summary = $"{name}: {matched} of {expected} shapes laid out, worst edge {worst.ToString("0.0", CultureInfo.InvariantCulture)} pt";
            if (required) TestReport.Check(matched == expected && worst <= TolerancePoints, summary + (matched == expected && worst <= TolerancePoints ? "" : " " + detail));
            else Console.WriteLine("   INFO  " + summary + " (not yet matched: " + detail + ")");
            string stripped = Path.Combine(folder, Path.GetFileNameWithoutExtension(path) + "-nodrawing.docx");
            if (required && File.Exists(stripped)) CheckDocument(stripped, name);
        }
        RunDeepData();
        RunFontMetrics();
    }

    /// <summary>
    /// The installed font files read at run time agree with the tables generated from them (scripts/chart-font-widths.py):
    /// Calibri's printable ASCII and kern pairs, Arial's Hebrew letters; niqqud is a non-spacing mark. Skipped off Windows.
    /// </summary>
    private static void RunFontMetrics()
    {
        TestReport.Step("Font metrics from the installed font files (WordFontMetrics)");
        if (WordFontFile.Find("Calibri") is null || WordFontFile.Find("Arial") is null) { Console.WriteLine("   INFO  Calibri / Arial not installed: skipped"); return; }
        WordFontMetrics calibri = WordFontMetrics.Of("Calibri")!, arial = WordFontMetrics.Of("Arial")!;
        double[] widths = DocxChartFontWidths.Fonts.First(f => f.Name == "Calibri").Widths;
        double worst = Enumerable.Range(' ', 95).Max(c => Math.Abs(calibri.Advance(c) - widths[c - ' ']));
        TestReport.Check(worst < 0.0001, $"Calibri ASCII advances as the generated table (worst {worst:0.00000} em)");
        double worstKern = DocxCalibriKerning.All().Max(p => Math.Abs(calibri.Kerning(p.Left, p.Right) - p.Em));
        TestReport.Check(worstKern < 0.0001, $"Calibri kern pairs as the generated table (worst {worstKern:0.00000} em)");
        double worstHebrew = Enumerable.Range(0x05D0, 27).Max(c => Math.Abs(arial.Advance(c) - DocxChartFontWidths.Hebrew[c - 0x0590]));
        TestReport.Check(worstHebrew < 0.0001 && arial.Has(0x05D0), $"Arial Hebrew letters as the generated table (worst {worstHebrew:0.00000} em)");
        TestReport.Check(WordFontMetrics.IsNonSpacingMark(0x05B8) && WordFontMetrics.IsNonSpacingMark(0x05BC) && !WordFontMetrics.IsNonSpacingMark(0x05D0), "niqqud (U+05B8, U+05BC) are non-spacing marks");
    }

    /// <summary>
    /// Hostile data: a 200,000-deep parOf chain (the old recursive calcMaxDepth overflowed the stack on it), a parOf cycle and a
    /// diamond. maxDepth / depth walk the whole real structure, no cap.
    /// </summary>
    private static void RunDeepData()
    {
        TestReport.Step("SmartArt data: deep parOf chain, cycle and diamond (maxDepth / depth without recursion)");
        const int deep = 200_000;
        var points = new XElement("ptLst", new XElement("pt", new XAttribute("modelId", "p0"), new XAttribute("type", "doc")));
        var connections = new XElement("cxnLst");
        for (int i = 1; i <= deep; i++)
        {
            points.Add(new XElement("pt", new XAttribute("modelId", "p" + i)));
            connections.Add(new XElement("cxn", new XAttribute("modelId", "c" + i), new XAttribute("srcId", "p" + (i - 1)), new XAttribute("destId", "p" + i),
                                         new XAttribute("parTransId", "t" + i), new XAttribute("sibTransId", "s" + i)));
        }
        var chain = new DocxDiagramLayout.DataModel(new XElement("dataModel", points, connections));
        TestReport.Check(chain.MaxDepth("p0") == deep, $"deep chain: maxDepth {chain.MaxDepth("p0")} (want {deep})");
        TestReport.Check(chain.Depth("p" + deep) == deep, $"deep chain: depth of the last point {chain.Depth("p" + deep)} (want {deep})");

        static XElement Cxn(string from, string to) => new("cxn", new XAttribute("srcId", from), new XAttribute("destId", to), new XAttribute("parTransId", "t"), new XAttribute("sibTransId", "s"));
        var cycle = new DocxDiagramLayout.DataModel(new XElement("dataModel", new XElement("cxnLst", Cxn("a", "b"), Cxn("b", "c"), Cxn("c", "a"))));
        TestReport.Check(cycle.MaxDepth("a") == 3 && cycle.Depth("a") == 2, $"cycle: maxDepth {cycle.MaxDepth("a")} (3), depth {cycle.Depth("a")} (2), no hang");
        var diamond = new DocxDiagramLayout.DataModel(new XElement("dataModel", new XElement("cxnLst", Cxn("r", "x"), Cxn("r", "y"), Cxn("x", "z"), Cxn("y", "z"), Cxn("z", "w"))));
        TestReport.Check(diamond.MaxDepth("r") == 3 && diamond.MaxDepth("w") == 0, $"diamond: maxDepth {diamond.MaxDepth("r")} (3), leaf {diamond.MaxDepth("w")} (0)");
    }

    /// <summary>The laid-out shapes against the drawing part's, by modelId: the largest difference of any box edge, in points.</summary>
    private static (double Worst, int Matched, int Expected, string Detail, string Fonts) Compare(string path)
    {
        using ZipArchive zip = ZipFile.OpenRead(path);
        XElement Part(string name) => DocxShapeGeometry.WithoutNamespaces(XDocument.Load(zip.GetEntry(name)!.Open()).Root!);
        XElement document = XDocument.Load(zip.GetEntry("word/document.xml")!.Open()).Root!;
        XElement extent = document.Descendants().First(e => e.Name.LocalName == "extent");
        double width = double.Parse((string)extent.Attribute("cx")!, CultureInfo.InvariantCulture), height = double.Parse((string)extent.Attribute("cy")!, CultureInfo.InvariantCulture);
        // the document's theme and settings: the text is measured in the theme's fonts (minor Latin, complex-script for Hebrew)
        XDocument? Document(string name) => zip.GetEntry(name) is { } entry ? XDocument.Load(entry.Open()) : null;
        var theme = new DocxTheme(Document("word/theme/theme1.xml"), Document("word/settings.xml"));
        XElement? laidOut = DocxDiagramLayout.Lay(Part("word/diagrams/data1.xml"), Part("word/diagrams/layout1.xml"), Part("word/diagrams/colors1.xml"), Part("word/diagrams/quickStyle1.xml"), width, height, theme);
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
        // the autofit font size: each shape's first run at Word's size (whole points); by text where Word writes a text shape's
        // text into the geometry it lies over (the pyramid's levelTx in its level's trapezoid)
        var ourSizes = FontSizes(laidOut);
        var ourSizesByText = ourSizes.Values.GroupBy(v => v.Text).ToDictionary(g => g.Key, g => g.First().Size);
        var fonts = new List<string>();
        foreach ((string id, (string size, string text)) in FontSizes(Part("word/diagrams/drawing1.xml").Element("spTree")))
        {
            string? mine = ourSizes.TryGetValue(id, out var own) ? own.Size : ourSizesByText.TryGetValue(text, out string? byText) ? byText : null;
            if (mine != size) fonts.Add($"{Shorten(id)} {mine ?? "none"} word {size}");
        }
        return (worst, matched, words.Count, string.Join("; ", detail.Take(4)), string.Join("; ", fonts.Take(3)));
    }

    private static Dictionary<string, (string Size, string Text)> FontSizes(XElement? tree)
    {
        var sizes = new Dictionary<string, (string, string)>();
        foreach (XElement shape in tree?.Elements("sp") ?? [])
            if (shape.Descendants().FirstOrDefault(e => e.Name.LocalName == "rPr" && e.Attribute("sz") is not null) is { } run)
                sizes[(string?)shape.Attribute("modelId") ?? ""] = ((string)run.Attribute("sz")!, string.Concat(shape.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value)));
        return sizes;
    }

    private static Dictionary<string, (double X, double Y, double W, double H)> Boxes(XElement? tree)
    {
        var boxes = new Dictionary<string, (double, double, double, double)>();
        foreach (XElement shape in tree?.Elements("sp") ?? [])
        {
            // preset shapes and the custom-geometry connectors Word writes for hierarchies (their boxes)
            XElement? transform = shape.Element("spPr")?.Element("xfrm");
            if (transform is null || shape.Element("spPr")?.Elements().Any(e => e.Name.LocalName is "prstGeom" or "custGeom") is not true) continue;
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
        // search (the one text path, GetText too) finds each node text as often as the node texts hold it (once, or inside a longer one too), as with a drawing part
        using ZipArchive zip = ZipFile.OpenRead(path);
        XElement data = DocxShapeGeometry.WithoutNamespaces(XDocument.Load(zip.GetEntry("word/diagrams/data1.xml")!.Open()).Root!);
        var texts = data.Element("ptLst")!.Elements("pt").Where(p => p.Attribute("type") is null).Select(p => string.Concat(p.Descendants("r").Select(r => (string?)r.Element("t") ?? ""))).Where(t => t.Length > 0).Distinct().ToList();
        foreach (string text in texts.Take(3))
        {
            int expected = texts.Sum(t => Regex.Matches(t, Regex.Escape(text)).Count);
            int hits = 0;
            for (int from = 0; from < document.BlockCount;)
            {
                WordDocumentSearchBatch batch = document.Search(text, WordDocumentSearchOptions.Default, from, 4);
                hits += batch.HitCounts.Sum();
                if (batch.ReachedEnd) break;
                from = batch.NextBlock;
            }
            TestReport.Check(hits == expected, $"{name} without its drawing part: search finds \"{Ascii(text)}\" {expected} time(s) ({hits})");
        }
    }

    /// <summary>The text with every character outside ASCII as U+XXXX (console output stays ASCII).</summary>
    private static string Ascii(string text) => string.Concat(text.Select(c => c < 128 ? c.ToString() : $"U+{(int)c:X4}"));

    private static string Box((double X, double Y, double W, double H) box) => $"({box.X:0.0},{box.Y:0.0} {box.W:0.0}x{box.H:0.0})";
    private static string Shorten(string id) => id.Length > 9 ? id.Substring(1, 8) : id;
}
