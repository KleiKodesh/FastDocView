using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// Torat Emet detection without a title line or without line-1 metadata, heading levels (dollar 1, hash 2, at and caret 3, tilde 4), and label
/// lines (exclamation mark and space) merged into the next line: synthetic cases, the boundary cases between fed pieces, the first-open path.
/// Text is ASCII; the file encoding is a single-byte code page.
/// </summary>
public static class ToratEmetWidenedTest
{
    private static readonly Encoding Hebrew1255 = MarkupTextEncoding.FromCodePage(MarkupTextEncoding.HebrewWindows);

    public static void Run(string directory)
    {
        RunDetection(directory);
        RunLabels(directory);
        RunLabelBoundaries(directory);
    }

    private static string Path1(string directory, string name) => Path.Combine(directory, $"torat-emet-widened-{name}-v2.txt");

    private static string Lines(params string[] lines) => string.Join("\r\n", lines) + "\r\n";

    // -- detection ---------------------------------------------------------------------------

    private static void RunDetection(string directory)
    {
        TestReport.Step("Torat Emet: detection without a title line, and without line-1 metadata");
        string plain150 = string.Concat(Enumerable.Repeat("plain\r\n", 150));
        var cases = new (string name, string text, bool isDetected, bool hasHead, string headings)[]
        {
            ("wrap-marks", Lines("&wrap=1", "@ Part", "~ Chapter", "text"), true, true, "3:1,4:2"),
            ("wrap-upper", Lines("&WRAP=1", "# T", "@ A", "text"), true, true, "2:1,3:2"),
            ("wrap-tilde-only", Lines("&wrap=1", "~ C", "text"), true, true, "4:1"),
            ("rav-at-only", Lines("&rav=X", "@ C", "text"), true, true, "3:1"),
            ("wrap-caret", Lines("&wrap=1", "^ C", "text"), true, true, "3:1"),
            ("wrap-no-marks", Lines("&wrap=1", "plain", "more plain"), true, true, ""),
            ("wrap-late-mark", "&wrap=1\r\n" + plain150 + "@ Late\r\n", true, true, "3:151"),
            ("key-late-dollar", "&UniqueId=3\r\n" + plain150 + "$ Late\r\n", true, true, "1:151"),
            ("key-dollar-title", Lines("&UniqueId=3", "$ T", "text"), true, true, "1:1"),
            ("dollar-hash", Lines("$ Title", "text", "# Sub", "more"), true, false, "1:0,2:2"),
            ("dollar-tilde", Lines("$ Title", "~ Chapter", "text"), true, false, "1:0,4:1"),
            ("dollar-at", Lines("$ Title", "text", "@ A"), true, false, "1:0,3:2"),
            // plain .txt regressions
            ("md-title", Lines("# Title", "some text <b>b</b>", "more"), false, false, ""),
            ("md-title-sub", Lines("# Title", "text", "## not a mark", "more"), false, false, ""),
            ("dollar-only", Lines("$ Title", "text", "! label"), false, false, ""),
            ("dollar-far", string.Concat(Enumerable.Repeat("plain line\r\n", 120)) + "$ Title\r\n# x\r\n", false, false, ""),
            ("unknown-key", Lines("&foo=1&bar=2", "$ Title", "# x", "@ y"), false, false, ""),
            ("unknown-word", Lines("&word=1", "@ x", "~ y"), false, false, ""),
            ("plain-note", Lines("A plain note", "with two lines & a=b"), false, false, ""),
            ("key-no-mark", "&UniqueId=3\r\n" + string.Concat(Enumerable.Repeat("plain line\r\n", 20)), false, false, ""),
            ("key-bang-only", Lines("&UniqueId=3", "! x", "text"), false, false, ""),
            ("one-line-wrap", "&wrap=1\r\n", false, false, ""),
            ("marks-too-far", "&UniqueId=3\r\n" + string.Concat(Enumerable.Repeat("plain\r\n", 250)) + "$ Late\r\n", false, false, ""),
        };
        var sidecar = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = WordIndexMode.Sidecar };
        var onDemandPolicy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache-ondemand"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        foreach ((string name, string text, bool isDetected, bool hasHead, string headings) in cases)
        {
            string path = Path1(directory, name);
            File.WriteAllBytes(path, Hebrew1255.GetBytes(text));
            foreach (WordIndexPolicy policy in new[] { sidecar, onDemandPolicy })
                if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
            using var whole = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            long wholeEnd = whole.ToratEmet?.HeadEndOffset ?? -2;
            TestReport.Check(whole.ToratEmet is not null == isDetected, $"{name}: {(isDetected ? "detected" : "stays plain .txt")}");
            if (!isDetected) { TestReport.Check(whole.Headings.Count == 0 && whole.DocumentTitle is null, $"{name}: no headings, no title"); }
            else
            {
                TestReport.Check(whole.ToratEmet!.HasHead == hasHead && whole.ToratEmet.BookName is null && whole.DocumentTitle is null, $"{name}: head {(hasHead ? "hidden" : "empty")}, no book name (the file name is the title)");
                string got = string.Join(",", whole.Headings.Select(h => $"{h.Level}:{h.BlockIndex}"));
                TestReport.Check(got == headings, $"{name}: the headings that exist, no root ({got})");
                string first = whole.RenderBlocks(0, 1)[0].Html;
                TestReport.Check(hasHead ? !first.Contains("wrap") && !first.Contains("UniqueId") && first.Contains("te-meta") : !first.Contains("te-meta"),
                                 $"{name}: line 1 is the hidden head or an ordinary line");
            }
            using (var saved1 = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, sidecar)) { }
            using var saved = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, sidecar);
            TestReport.Check((saved.ToratEmet?.HeadEndOffset ?? -2) == wholeEnd && saved.Headings.Count == whole.Headings.Count, $"{name}: the saved index decides the same");
            using var onDemand = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, onDemandPolicy);
            long partialEnd = onDemand.ToratEmet?.HeadEndOffset ?? -2;
            while (!onDemand.IsIndexComplete) onDemand.ScanAhead(5000);
            TestReport.Check(partialEnd == wholeEnd && (onDemand.ToratEmet?.HeadEndOffset ?? -2) == wholeEnd, $"{name}: the partial first-open index decides the same, before and after it finishes");
        }
    }

    // -- label lines -------------------------------------------------------------------------

    /// <summary>The blocks a text must produce: a label line merges with the next line unless that is missing, blank, or a mark and a space.</summary>
    private static List<string> ExpectedLines(string[] lines)
    {
        var result = new List<string>();
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            if (line.StartsWith("! ", StringComparison.Ordinal) && i + 1 < lines.Length && lines[i + 1].Length > 0
                && !(lines[i + 1].Length >= 2 && lines[i + 1][1] == ' ' && "$#@^~!".IndexOf(lines[i + 1][0]) >= 0))
            { result.Add("!M" + line.Substring(2) + " " + lines[i + 1]); i++; }
            else result.Add(line);
        }
        return result;
    }

    private static string ExpectedHtml(string expected) =>
        expected.StartsWith("!M", StringComparison.Ordinal) ? ToratEmetFormat.RenderLine("", expected.Substring(2)) : ToratEmetFormat.RenderLine("", expected);

    private static void CheckBlocks(MarkupWindowedDocument document, int firstBlock, List<string> expected, string what)
    {
        bool ok = document.BlockCount == firstBlock + expected.Count;
        if (ok)
        {
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(firstBlock, expected.Count);
            for (int i = 0; i < expected.Count && ok; i++) ok = blocks[i].Html == ExpectedHtml(expected[i]);
        }
        TestReport.Check(ok, $"{what}: {expected.Count} blocks as expected (found {document.BlockCount - firstBlock})");
    }

    private static void RunLabels(string directory)
    {
        TestReport.Step("Torat Emet: a label line (exclamation mark) merges with the next line");
        var cases = new (string name, string[] lines)[]
        {
            ("basic", ["$ T", "! {A}", "verse text", "after"]),
            ("two-labels-text", ["$ T", "! {A}", "verse one", "! {B}", "verse two"]),
            ("label-last", ["$ T", "text", "! {A}"]),
            ("label-then-blank", ["$ T", "! {A}", "", "text"]),
            ("label-then-heading", ["$ T", "! {A}", "~ Chapter", "text"]),
            ("label-then-dollar", ["$ T", "! {A}", "# H", "text"]),
            ("label-then-label", ["$ T", "! {A}", "! {B}", "verse"]),
            ("label-then-lone-mark", ["$ T", "! {A}", "~x not a heading", "text"]),
            ("label-then-single-mark-char", ["$ T", "! {A}", "~", "text"]),
            ("label-then-bang-word", ["$ T", "! {A}", "!x word", "text"]),
            ("label-then-tag", ["$ T", "! {A}", "<b>bold</b> text", "x"]),
            ("bang-no-space", ["$ T", "!{A}", "text"]),
            ("label-then-spaces", ["$ T", "! {A}", "   indented", "x"]),
        };
        foreach ((string name, string[] lines) in cases)
        {
            foreach ((string ending, string tag) in new[] { ("\r\n", "crlf"), ("\n", "lf"), ("\r", "cr") })
            {
                string text = "&UniqueId=1&ForcedBookName=X\r\n" + string.Join(ending, lines) + ending;
                if (name == "label-last" && tag == "lf") text = text.TrimEnd('\n');     // no line break after the very last line
                string path = Path1(directory, $"label-{name}-{tag}");
                File.WriteAllBytes(path, Hebrew1255.GetBytes(text));
                var onDemandPolicy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache-ondemand"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
                if (File.Exists(onDemandPolicy.IndexPathFor(path))) File.Delete(onDemandPolicy.IndexPathFor(path));
                List<string> expected = ExpectedLines(lines);
                using var whole = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
                CheckBlocks(whole, 1, expected, $"{name} {tag}");
                using var onDemand = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, onDemandPolicy);
                while (!onDemand.IsIndexComplete) onDemand.ScanAhead(5000);
                CheckBlocks(onDemand, 1, expected, $"{name} {tag} (first-open path)");
            }
        }

        string pathBasic = Path1(directory, "label-basic-crlf");
        using (var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(pathBasic, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory }))
        {
            TestReport.Check(document.RenderBlocks(2, 1)[0].Html == "<div class=\"text-line\">{A} verse text</div>", "the label, one space and the verse are ONE line, no mark and no break");
            TestReport.Check(document.Search("{A} verse", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.SequenceEqual([2]), "search matches the merged text across the label and the verse");
            TestReport.Check(document.Search("! {A}", WordDocumentSearchOptions.Default, 0, 100).HitBlocks.Length == 0, "search never matches the hidden mark");
        }
        string pathAlone = Path1(directory, "label-label-then-heading-crlf");
        using (var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(pathAlone, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory }))
            TestReport.Check(document.RenderBlocks(2, 1)[0].Html == "<div class=\"text-line\">{A}</div>" && document.Headings.Any(h => h.Text == "Chapter" && h.Level == 4 && h.BlockIndex == 3), "a label before a heading stays alone; the heading is the next block");

        TestReport.Step("Torat Emet: a long line after a label is split, the label merges with the first piece only");
        string longText = "&UniqueId=1&ForcedBookName=X\r\n$ T\r\n! {A}\r\n" + string.Join(" ", Enumerable.Repeat("word", 20000)) + "\r\nend\r\n";
        string longPath = Path1(directory, "label-long-line");
        File.WriteAllBytes(longPath, Hebrew1255.GetBytes(longText));
        using (var whole = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(longPath, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory }))
        {
            string firstPiece = whole.RenderBlocks(2, 1)[0].Html;
            TestReport.Check(firstPiece.StartsWith("<div class=\"text-line\">{A} word word", StringComparison.Ordinal), $"{firstPiece.Substring(0, Math.Min(80, firstPiece.Length))} {whole.BlockCount} blocks; the label heads the first piece of the long line");
        }
    }

    // -- boundaries ----------------------------------------------------------------------------

    /// <summary>
    /// The label line is moved so that every byte of "label line break, next line's first two bytes" in turn sits on the end of the first fed piece
    /// (the detection sample, and the first chunk after it): the blocks must not depend on where the pieces end.
    /// </summary>
    private static void RunLabelBoundaries(string directory)
    {
        TestReport.Step("Torat Emet: labels across fed-piece boundaries (first open and one pass)");
        var tails = new (string name, string[] lines)[]
        {
            ("text", ["! {A}", "verse x", "tail"]),
            ("heading", ["! {A}", "~ Chapter", "tail"]),
            ("mark-char", ["! {A}", "~x y", "tail"]),
            ("blank", ["! {A}", "", "tail"]),
            ("label-label", ["! {A}", "! {B}", "tail"]),
            ("last", ["! {A}"]),
        };
        int checks = 0, failures = 0;
        foreach (int boundary in new[] { MarkupTextEncoding.DetectionSampleBytes, MarkupTextEncoding.DetectionSampleBytes + MarkupRandomAccessIndexBuilder.ChunkSize })
        foreach ((string tailName, string[] tail) in tails)
        foreach (string ending in new[] { "\r\n", "\n" })
        for (int shift = -3; shift <= 9; shift++)
        {
            // the label line starts at (boundary - shift) in file bytes: shift 0 = the label's first byte is the first of the next piece
            var text = new StringBuilder("&UniqueId=1&ForcedBookName=X" + ending + "$ T" + ending);
            while (text.Length < boundary - 80) text.Append("filler line of ascii text" + ending);
            int pad = boundary - shift - text.Length - ending.Length;
            if (pad < 0) { text.Length = text.Length - 30; pad = boundary - shift - text.Length - ending.Length; }
            text.Append(new string('f', pad)).Append(ending);
            var fileLines = new List<string>();
            string prefix = text.ToString();
            string body = string.Join(ending, tail) + (tailName == "last" ? "" : ending);
            string path = Path1(directory, $"bound-{boundary}-{tailName}-{(ending.Length == 2 ? "crlf" : "lf")}-{shift + 3}");
            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(prefix + body));
            string[] lines = (prefix + body).Split([ending], StringSplitOptions.None).ToArray();
            if (lines[^1].Length == 0) lines = lines.Take(lines.Length - 1).ToArray();
            List<string> expected = ExpectedLines(lines.Skip(1).ToArray());      // everything after the hidden head line
            var onDemandPolicy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache-ondemand"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
            if (File.Exists(onDemandPolicy.IndexPathFor(path))) File.Delete(onDemandPolicy.IndexPathFor(path));
            using var whole = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory });
            using var onDemand = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, onDemandPolicy);
            while (!onDemand.IsIndexComplete) onDemand.ScanAhead(5000);
            foreach ((MarkupWindowedDocument document, string how) in new[] { (whole, "one pass"), (onDemand, "first open") })
            {
                checks++;
                string detail = "";
                bool ok = document.BlockCount == 1 + expected.Count;
                if (ok)
                {
                    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(1, expected.Count);
                    // the filler block text is long and not interesting: compare the last blocks (the label region) and the count
                    int from = Math.Max(0, expected.Count - 4);
                    for (int i = from; i < expected.Count && ok; i++) { ok = blocks[i].Html == ExpectedHtml(expected[i]); if (!ok) detail = string.Join(" | ", blocks.Skip(from).Select(x => x.Html.Substring(23))) + " <> " + ExpectedHtml(expected[i]); }
                }
                if (!ok) { failures++; Console.WriteLine($"   FAIL boundary {boundary} {tailName} shift {shift} {how}: {document.BlockCount} blocks, expected {1 + expected.Count} {detail}"); }
            }
        }
        TestReport.Check(failures == 0, $"label boundary cases: {checks} checks, {failures} failures");
    }
}
