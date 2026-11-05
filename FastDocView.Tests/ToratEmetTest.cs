using System.Diagnostics;
using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// Torat Emet book files read as .txt: synthetic cases (detection, head, marks, heading levels, the tag whitelist, the on-demand and cached
/// opens, a mark split across two fed pieces, plain .txt unchanged) and, with a folder, a fixed-seed sample of the real books.
/// Hebrew is written as \u escapes; output stays ASCII.
/// </summary>
public static class ToratEmetTest
{
    private const string Name = "ספר בדיקה";     // "test book"
    private const string Letter = "A";

    private static readonly Encoding Hebrew1255 = MarkupTextEncoding.FromCodePage(MarkupTextEncoding.HebrewWindows);

    private static string BookLines(string head = "") =>
        "&UniqueId=7&LastLevelIndex=2&PutNewLinesAsIs=1&ForcedBookName=" + Name + "&TextSource=Src&RavMechaber=Auth&CosmeticsType=#wp=//x.bmp#rep3=<b115>=<b style='a=b&c'>#rep4=<span style='open &GroupId=g\r\n"
        + "// editor comment\r\n[General]\r\nKey=value\r\n<B>&LastLevelIndex=2&TextSource=Dup</B>\r\n\r\n"
        + head
        + "$ Book Title\r\n\r\n^ Parsha\r\n# Hash title\r\n@ At title\r\n~ Chapter <B>bold</B>\r\n! {" + Letter + "}\r\n"
        + "text <B>bold<SMALL>x</SMALL></B> <90>tail <SZ:150%%>x <TF>t <span style='a'>sp</span> <script>js</script> a < b <!-- hidden -->end<BR>\r\n"
        + "// body slash line\r\n<B>open across\r\nnext line\r\n</B>closed\r\n";

    public static void Run()
    {
        TestReport.Step("Torat Emet: synthetic book");
        string directory = TestReport.OutputDirectory;
        string path = Path.Combine(directory, "torat-emet-synthetic-v1.txt");
        File.WriteAllBytes(path, Hebrew1255.GetBytes(BookLines()));
        foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
        {
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
            {
                using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                VerifySynthetic(document, $"{mode}{(pass == 1 ? " from the cached index" : "")}");
            }
        }

        TestReport.Step("Torat Emet: a head with a text line before the first heading shows that line");
        string textHead = Path.Combine(directory, "torat-emet-text-in-head-v1.txt");
        File.WriteAllBytes(textHead, Hebrew1255.GetBytes("&UniqueId=1\r\n// c\r\nA real line of text\r\n$ T\r\nbody\r\n"));
        using (var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(textHead, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory }))
        {
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
            TestReport.Check(document.ToratEmet is not null && document.BlockCount == 4 && blocks[1].Html.Contains("A real line of text") && !blocks[0].Html.Contains("UniqueId"),
                             "metadata lines are hidden; the first line that is text is shown");
        }

        RunNotTheFormat(directory);
        ToratEmetWidenedTest.Run(directory);
        RunSplitPieces(directory);
        RunOnDemand(directory);
        RunTagExamples();
    }

    private static void VerifySynthetic(MarkupWindowedDocument document, string what)
    {
        TestReport.Check(document.ToratEmet is not null, $"{what}: detected as a Torat Emet book");
        TestReport.Check(document.DocumentTitle == Name, $"{what}: ForcedBookName is the document title");
        TestReport.Check(document.ToratEmet!.Get("CosmeticsType")!.StartsWith("#wp=//x.bmp#rep3=<b115>=<b style='a=b&c'>", StringComparison.Ordinal)
                         && document.ToratEmet.Get("GroupId") == "g" && document.ToratEmet.Get("RavMechaber") == "Auth" && document.ToratEmet.Get("Key") == "value",
                         $"{what}: metadata parsed ('=' and '&' inside a value stay)");
        IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
        TestReport.Check(blocks[0].Html.Contains(Name) && blocks[0].Html.Contains("Auth") && blocks[0].Html.Contains("Src")
                         && !blocks[0].Html.Contains("UniqueId") && !blocks[0].Html.Contains("editor comment") && !blocks[0].Html.Contains("Dup") && !blocks[0].Html.Contains("[General]"),
                         $"{what}: block 0 is the small metadata header, the head lines are not shown");
        TestReport.Check(string.Join(",", document.Headings.Select(h => $"{h.Level}:{h.BlockIndex}")) == "1:1,3:3,2:4,3:5,4:6", $"{what}: heading levels ({string.Join(",", document.Headings.Select(h => h.Level + ":" + h.BlockIndex))})");
        TestReport.Check(document.Headings[0].Text == Name && document.Headings[4].Text == "Chapter bold", $"{what}: the TOC root is the book name; a heading's text has no mark and no tags");
        TestReport.Check(blocks[1].Html == "<div class=\"text-line te-heading\"><h1>Book Title</h1></div>", $"{what}: $ is an h1 ({blocks[1].Html})");
        TestReport.Check(blocks[4].Html.Contains("<h2>Hash title</h2>") && blocks[5].Html.Contains("<h3>At title</h3>") && blocks[3].Html.Contains("<h3>Parsha</h3>")
                         && blocks[6].Html.Contains("<h4>Chapter <b>bold</b></h4>"), $"{what}: # h2, @ and ^ h3, ~ h4");
        TestReport.Check(blocks[7].Html.StartsWith("<div class=\"text-line\">{" + Letter + "} text <b>bold", StringComparison.Ordinal), $"{what}: the exclamation mark and its space are stripped, the label merged with the next line by one space ({blocks[7].Html})");
        string text = blocks[7].Html;
        TestReport.Check(text.Contains("<b>bold<small>x</small></b>") && text.EndsWith("end<br></div>", StringComparison.Ordinal)
                         && !text.Contains("<span") && !text.Contains("<script") && !text.Contains("&lt;90") && !text.Contains("<sz") && !text.Contains("hidden")
                         && text.Contains("tail ") && text.Contains("sp") && text.Contains("a &lt; b"),
                         $"{what}: whitelist tags drawn, other tags dropped with their text kept, a stray '<' is text ({text})");
        TestReport.Check(blocks[8].Html.Contains("// body slash line"), $"{what}: a body '//' line stays visible");
        TestReport.Check(blocks[9].Html.Contains("<b>open across</b>") && blocks[10].Html.Contains("<b>next line</b>") && blocks[11].Html.Contains("closed") && !blocks[11].Html.Contains("<b>closed"),
                         $"{what}: a tag open across lines carries to the next line");
        WordDocumentSearchBatch hits = document.Search("Chapter", WordDocumentSearchOptions.Default, 0, 100);
        TestReport.Check(hits.HitBlocks.SequenceEqual([6]), $"{what}: search finds the heading text, never the hidden head");
        TestReport.Check(document.Search("UniqueId", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.Length == 0, $"{what}: search does not match the hidden metadata");
    }

    private static void RunNotTheFormat(string directory)
    {
        TestReport.Step("Torat Emet: plain .txt behaves exactly as before");
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        var cases = new (string name, string text)[]
        {
            ("unknown-key", "&foo=1&bar=2\r\n$ Title\r\n# x\r\n! y <b>b</b><span>s</span>\r\n"),
            ("no-heading", "&UniqueId=3\r\n" + string.Concat(Enumerable.Repeat("plain line\r\n", 250)) + "$ Title too far\r\n"),
            ("no-metadata", "# Title\r\n<b>b</b><span>s</span>\r\n"),
        };
        foreach ((string name, string text) in cases)
        {
            string path = Path.Combine(directory, $"torat-emet-not-{name}-v1.txt");
            File.WriteAllBytes(path, Hebrew1255.GetBytes(text));
            using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
            string[] lines = text.Split(["\r\n"], StringSplitOptions.None).Where(l => l.Length > 0).ToArray();
            TestReport.Check(document.ToratEmet is null && document.DocumentTitle is null && document.Headings.Count == 0 && document.BlockCount == lines.Length,
                             $"{name}: not detected, one block per line, no headings");
            TestReport.Check(blocks.Select((b, i) => (b, i)).All(p => p.b.Html == MarkupBlockRenderer.RenderLine(document.OpenTagsOf(p.i), lines[p.i])) && blocks[0].Html.Contains(lines[0].Substring(0, 3)),
                             $"{name}: every line drawn by the unchanged .txt renderer (marks and tags as before)");
        }
    }

    /// <summary>The mark of a heading line sits on the last byte of the first fed piece, its space on the first byte of the next.</summary>
    private static void RunSplitPieces(string directory)
    {
        TestReport.Step("Torat Emet: a heading mark and its space in two different fed pieces");
        foreach (int boundary in new[] { MarkupTextEncoding.DetectionSampleBytes, MarkupTextEncoding.DetectionSampleBytes + MarkupRandomAccessIndexBuilder.ChunkSize })
        {
            var text = new StringBuilder("&UniqueId=1&ForcedBookName=X\r\n$ T\r\n");
            while (text.Length < boundary - 80) text.Append("filler line of ascii text\r\n");
            text.Append(new string('f', boundary - 1 - text.Length - 2)).Append("\r\n");     // the next line starts at boundary - 1
            text.Append("# Split heading\r\nafter\r\n@ Second\r\n");
            string path = Path.Combine(directory, $"torat-emet-split-{boundary}-v1.txt");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(text.ToString()));
            using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            WordDocumentHeading? split = document.Headings.FirstOrDefault(h => h.Text == "Split heading");
            TestReport.Check(split is { Level: 2 } && document.Headings.Any(h => h.Text == "Second" && h.Level == 3) && document.Headings[0].Text == "X",
                             $"mark and space split at byte {boundary}: heading found without the mark ({string.Join("|", document.Headings.Select(h => h.Level + h.Text))})");
            IReadOnlyList<WordDocumentBlockHtml> rendered = document.RenderBlocks(split!.BlockIndex, 1);
            TestReport.Check(rendered[0].Html == "<div class=\"text-line te-heading\"><h2>Split heading</h2></div>", "the split heading renders as an h2");
        }
    }

    /// <summary>The on-demand first open (sidecar, index built while windows are read) ends with the same index as the one-pass open.</summary>
    private static void RunOnDemand(string directory)
    {
        TestReport.Step("Torat Emet: on-demand first open");
        var text = new StringBuilder(BookLines());
        for (int i = 0; i < 40_000; i++) text.Append(i % 500 == 0 ? "# Part " + i + "\r\n" : "line " + i + " <B>x</B>\r\n");
        string path = Path.Combine(directory, "torat-emet-ondemand-v1.txt");
        File.WriteAllBytes(path, Hebrew1255.GetBytes(text.ToString()));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        string indexPath = policy.IndexPathFor(path);
        if (File.Exists(indexPath)) File.Delete(indexPath);
        using var onDemand = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        TestReport.Check(!onDemand.IsIndexComplete && onDemand.ToratEmet is not null, "first open returns on a partial index that already knows the head");
        string first = onDemand.RenderBlocks(0, 2)[0].Html;
        TestReport.Check(first.Contains(Name) && !first.Contains("UniqueId"), "the partial index draws the header");
        while (!onDemand.IsIndexComplete) onDemand.ScanAhead(5000);
        using var whole = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
        TestReport.Check(onDemand.BlockCount == whole.BlockCount && onDemand.Headings.Select(h => (h.BlockIndex, h.Level, h.Text)).SequenceEqual(whole.Headings.Select(h => (h.BlockIndex, h.Level, h.Text))),
                         $"on-demand index = one-pass index ({whole.BlockCount} blocks, {whole.Headings.Count} headings)");
        using var cached = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
        TestReport.Check(cached.Decision.HasCachedIndex && cached.ToratEmet?.HeadEndOffset == whole.ToratEmet!.HeadEndOffset && cached.Headings.Count == whole.Headings.Count,
                         "the saved sidecar carries the head and the heading levels");
    }

    private static void RunTagExamples()
    {
        string html = ToratEmetFormat.RenderInline("<B>", "a</B><i>b</u></i><sup>1</sup><sub>2</sub><u>u</u><big>g</big><br><A HREF=\"javascript:x\">l</A><img src=x onerror=y>z<div onclick=\"x\">d</div>");
        TestReport.Check(html == "<b>a</b><i>b</i><sup>1</sup><sub>2</sub><u>u</u><big>g</big><br>lzd", $"whitelist: only b i u big small sup sub br, bare, no attributes ({html})");
    }

    // -- real books -------------------------------------------------------------------------

    /// <summary>
    /// The books in <paramref name="folder"/>: a fixed-seed sample (or every file with <paramref name="isFullPass"/>, index only: statistics) opened, rendered
    /// from the start, the middle and the end, and checked: detected, a heading list, no mark or foreign tag in the drawn HTML. Light: below normal
    /// priority, one file at a time.
    /// </summary>
    public static void RunReal(string folder, int sample, bool isFullPass)
    {
        try { Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.Idle; } catch { }
        string[] all = Directory.EnumerateFiles(folder, "*.txt", SearchOption.AllDirectories).OrderBy(p => p, StringComparer.Ordinal).ToArray();
        var shuffle = new Random(2026);
        string[] files = isFullPass ? all : all.OrderBy(_ => shuffle.Next()).Take(sample).ToArray();
        TestReport.Step($"Torat Emet real books: {files.Length} of {all.Length} files{(isFullPass ? " (statistics pass, index only)" : "")}");
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        int detected = 0, notDetected = 0, failures = 0, done = 0;
        long blocks = 0;
        var levels = new long[10];
        var codePages = new Dictionary<int, int>();
        var undetected = new List<string>();
        var watch = Stopwatch.StartNew();
        foreach (string path in files)
        {
            if (++done % 100 == 0) Console.WriteLine($"   [{watch.Elapsed:mm\\:ss}] {done} / {files.Length}");
            try
            {
                using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                codePages[document.Index.CodePage] = codePages.GetValueOrDefault(document.Index.CodePage) + 1;
                if (document.ToratEmet is null) { notDetected++; undetected.Add(Path.GetFileName(path)); continue; }
                detected++;
                blocks += document.BlockCount;
                foreach (WordDocumentHeading heading in document.Headings) levels[heading.Level]++;
                if (isFullPass) continue;
                if ((document.Headings.Count > 0 ? false : !document.ToratEmet.Metadata.All(m => m.Key.Equals("wrap", StringComparison.OrdinalIgnoreCase) || m.Key.Equals("rav", StringComparison.OrdinalIgnoreCase)))) { failures++; Console.WriteLine($"   FAIL no headings: {Ascii(path)} headings {document.Headings.Count}: {string.Join(",", document.Headings.Take(4).Select(h => h.Level + ":" + h.BlockIndex))} head end {document.ToratEmet.HeadEndOffset}"); }
                foreach (int from in new[] { 0, document.BlockCount / 2, Math.Max(0, document.BlockCount - 40) })
                {
                    foreach (WordDocumentBlockHtml block in document.RenderBlocks(from, 40))
                    {
                        string html = block.Html;
                        bool bad = html.Contains("<span") || html.Contains("<script") || html.Contains("<img") || html.Contains("<a ") || html.Contains("&lt;90") || html.Contains("&lt;sz")
                                   || (block.BlockIndex == 0 ? html.Contains("UniqueId") : "!$#~@^".Any(mark => html.StartsWith("<div class=\"text-line\">" + mark + " ", StringComparison.Ordinal)));
                        if (bad) { failures++; Console.WriteLine("   cond " + string.Join(",", new[] { "span", "script", "img", "a ", "&lt;90", "&lt;sz", ">! ", ">$ ", ">~ ", ">@ ", ">^ " }.Where(html.Contains))); Console.WriteLine($"   FAIL block {block.BlockIndex} of {Ascii(path)}: {Ascii(html.Length > 160 ? html.Substring(0, 160) : html)}"); break; }
                    }
                }
            }
            catch (Exception exception) { failures++; Console.WriteLine($"   CRASH {Ascii(path)}: {exception.GetType().Name} {Ascii(exception.Message)}"); }
        }
        Console.WriteLine($"   detected {detected}, not detected {notDetected} ({Ascii(string.Join(", ", undetected.Take(12)))}), {blocks:N0} blocks, {watch.Elapsed:mm\\:ss}");
        Console.WriteLine($"   headings by level: h1 {levels[1]:N0}, h2 {levels[2]:N0}, h3 {levels[3]:N0}, h4 {levels[4]:N0}; code pages: {string.Join(", ", codePages.Select(p => p.Key + ":" + p.Value))}");
        TestReport.Check(failures == 0, $"Torat Emet real books: {failures} failures");
    }

    private static string Ascii(string text)
    {
        var builder = new StringBuilder(text.Length);
        foreach (char c in text) builder.Append(c is >= ' ' and <= '~' ? c : '?');
        return builder.ToString();
    }
}
