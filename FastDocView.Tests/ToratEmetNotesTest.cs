using System.Text;
using FastDocView.Markup;

namespace FastDocView.Tests;

/// <summary>
/// Torat Emet note lines (see <see cref="ToratEmetNotes"/>): a listed plain line right after the title moves from the text into the header line.
/// The Ramban shape, notes with blank and comment lines between, a title-only book that keeps every line, unlisted lines, the line limit, the three open
/// paths (one pass, on-demand first open, cached reopen) agreeing, and scan-piece boundaries. Hebrew is written as \u escapes; output stays ASCII.
/// </summary>
public static class ToratEmetNotesTest
{
    private const string NoteOne = "הספר אינו שלם ואינו מוגה";   // "the book is incomplete and unvetted"
    private const string NoteTwo = "באור הענין";
    private const string Unlisted = "שלום עליכם ורבותי";
    private const string Dot = " · ";
    private static readonly Encoding Hebrew1255 = MarkupTextEncoding.FromCodePage(MarkupTextEncoding.HebrewWindows);

    public static void Run(string directory)
    {
        TestReport.Step("Torat Emet: note lines move into the header");
        foreach (bool isAsIs in new[] { false, true })
        {
            string head = "&UniqueId=1&PutNewLinesAsIs=" + (isAsIs ? "1" : "0") + "&TextSource=Src&ForcedBookName=Nm\r\n";
            string name = isAsIs ? "asis" : "drop";
            // the Ramban shape: head, blank, title, note, blank, section heading, text
            Verify(directory, "ramban-" + name, head + "\r\n$ T\r\n" + NoteOne + "\r\n\r\n^ P\r\ntext\r\n", ["Nm", "Src", NoteOne], "T|P|text");
            Verify(directory, "ramban-bold-" + name, head + "\r\n$ T\r\n<b>" + NoteOne + ".</b>\r\n\r\n^ P\r\ntext\r\n", ["Nm", "Src", NoteOne + "."], "T|P|text");   // tags and a trailing full stop do not stop a match
            Verify(directory, "tight-" + name, head + "$ T\r\n" + NoteOne + "\r\n^ P\r\ntext\r\n", ["Nm", "Src", NoteOne], "T|P|text");
            // two notes with blank and comment lines between them
            Verify(directory, "two-" + name, head + "$ T\r\n\r\n" + NoteOne + "\r\n\r\n// c\r\n" + NoteTwo + "\r\n\r\n# H\r\ntext\r\n", ["Nm", "Src", NoteOne, NoteTwo], isAsIs ? "T|// c||H|text" : "T|// c|H|text");   // the comment is text; a blank line after the last note is a block when blank lines are kept
            // the fifth non-blank line is still a note; the sixth is text
            Verify(directory, "fifth-" + name, head + "$ T\r\n// 1\r\n// 2\r\n// 3\r\n// 4\r\n" + NoteOne + "\r\n# H\r\ntext\r\n", ["Nm", "Src", NoteOne], "T|// 1|// 2|// 3|// 4|H|text");
            Verify(directory, "sixth-" + name, head + "$ T\r\n// 1\r\n// 2\r\n// 3\r\n// 4\r\n// 5\r\n" + NoteOne + "\r\n# H\r\ntext\r\n", ["Nm", "Src"], "T|// 1|// 2|// 3|// 4|// 5|" + NoteOne + "|H|text");
            // text that is not a listed note stays text, and ends the notes
            Verify(directory, "unlisted-" + name, head + "$ T\r\n" + Unlisted + "\r\n# H\r\ntext\r\n", ["Nm", "Src"], "T|" + Unlisted + "|H|text");
            Verify(directory, "unlisted-first-" + name, head + "$ T\r\n" + Unlisted + "\r\n" + NoteOne + "\r\n# H\r\ntext\r\n", ["Nm", "Src"], "T|" + Unlisted + "|" + NoteOne + "|H|text");
            // a title-only book: the whole book is text after the title - nothing moves, even a note-looking line
            Verify(directory, "title-only-" + name, head + "$ T\r\n" + NoteOne + "\r\nbody a\r\nbody b\r\n", ["Nm", "Src"], "T|" + NoteOne + "|body a|body b");
            Verify(directory, "title-only-blank-" + name, head + "$ T\r\n\r\n" + NoteOne + "\r\n\r\n" + NoteTwo + "\r\n", ["Nm", "Src"], "T|" + NoteOne + (isAsIs ? "||" : "|") + NoteTwo);
            // the same line after text is never a note
            Verify(directory, "late-" + name, head + "$ T\r\nbody\r\n" + NoteOne + "\r\n# H\r\ntext\r\n", ["Nm", "Src"], "T|body|" + NoteOne + "|H|text");
        }
        RunPieceSplits();
        RunBoundaries(directory);
    }

    /// <summary>Opens the book three ways (one pass, on-demand first open, cached reopen) and checks the header parts and the plain text of every block after the header.</summary>
    private static void Verify(string directory, string name, string text, string[] headerParts, string expectedBlocks)
    {
        string path = Path.Combine(directory, "torat-emet-notes-" + name + "-v1.txt");
        File.WriteAllBytes(path, Hebrew1255.GetBytes(text));
        var onePass = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        var sidecar = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = WordIndexMode.Sidecar, IndexesOnDemand = true };
        if (File.Exists(sidecar.IndexPathFor(path))) File.Delete(sidecar.IndexPathFor(path));
        string? reference = null;
        foreach ((string what, WordIndexPolicy policy, bool isOnDemand) in new[] { ("one pass", onePass, false), ("on demand", sidecar, true), ("cached reopen", sidecar, false) })
        {
            using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
            if (isOnDemand) while (!document.IsIndexComplete) document.ScanAhead(5000);
            if (what == "cached reopen") TestReport.Check(document.Decision.HasCachedIndex, $"{name}: the reopen reads the sidecar");
            IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
            string header = Plain(blocks[0].Html);
            string rest = string.Join("|", blocks.Skip(1).Select(b => Plain(b.Html)));
            string expectedHeader = string.Join(Dot, headerParts);
            TestReport.Check(header == expectedHeader && rest == expectedBlocks, $"{name}, {what}: header [{Ascii(header)}] blocks [{Ascii(rest)}] expected [{Ascii(expectedHeader)}] [{Ascii(expectedBlocks)}]");
            string signature = Ascii(header + "#" + rest) + "#" + document.BlockCount + "#" + string.Join(",", document.Headings.Select(h => h.BlockIndex + ":" + h.Level));
            if (reference is null) reference = signature;
            else TestReport.Check(signature == reference, $"{name}, {what}: same blocks and headings as the one-pass open");
        }
    }

    private static string Plain(string html) => System.Net.WebUtility.HtmlDecode(System.Text.RegularExpressions.Regex.Replace(html, "<[^>]*>", "")).Trim();

    private static string Ascii(string text)
    {
        var result = new StringBuilder();
        foreach (char c in text) result.Append(c < 128 ? c.ToString() : "\\u" + ((int)c).ToString("X4"));
        return result.ToString();
    }

    /// <summary>The scanner alone, fed in two pieces split at every offset and in small pieces (a span rescan reads arbitrary pieces): the same blocks and headings as one feed.</summary>
    private static void RunPieceSplits()
    {
        foreach (bool isAsIs in new[] { false, true })
        foreach (string eol in new[] { "\r\n", "\n", "\r" })
        {
            string text = ("&UniqueId=1&PutNewLinesAsIs=" + (isAsIs ? "1" : "0") + "&ForcedBookName=Nm\r\n$ T\r\n\r\n" + NoteOne + "\r\n\r\n// c\r\n" + NoteTwo + "\r\n\r\n# H\r\ntext a\r\n\r\ntext b\r\n").Replace("\r\n", eol);
            byte[] bytes = Hebrew1255.GetBytes(text);
            ToratEmetHead head = ToratEmetFormat.TryReadHead(bytes, 0, Hebrew1255)!;
            string whole = Scan(head, bytes, bytes.Length);
            int bad = 0;
            for (int split = 1; split < bytes.Length; split++) if (Scan(head, bytes, split) != whole) bad++;
            for (int size = 1; size <= 9; size++) if (ScanPieces(head, bytes, size) != whole) bad++;
            TestReport.Check(head.Notes.Count == 2 && bad == 0, $"scanner pieces (as-is={isAsIs}, eol {eol.Length} byte(s), {bytes.Length} bytes): {bad} different results over every split and small pieces; {whole}");
        }
    }

    private static string Scan(ToratEmetHead head, byte[] bytes, int split)
    {
        var scanner = new MarkupBlockScanner(MarkupBlockMode.Lines, 1, false, Hebrew1255, head);
        scanner.Begin(0);
        scanner.Feed(bytes, 0, split, 0);
        scanner.Feed(bytes, split, bytes.Length - split, split);
        scanner.Finish(bytes.Length);
        return Describe(scanner);
    }

    private static string ScanPieces(ToratEmetHead head, byte[] bytes, int size)
    {
        var scanner = new MarkupBlockScanner(MarkupBlockMode.Lines, 1, false, Hebrew1255, head);
        scanner.Begin(0);
        for (int at = 0; at < bytes.Length; at += size) scanner.Feed(bytes, at, Math.Min(size, bytes.Length - at), at);
        scanner.Finish(bytes.Length);
        return Describe(scanner);
    }

    private static string Describe(MarkupBlockScanner scanner) =>
        string.Join(",", scanner.BlockStartOffsets.Select((start, i) => start + "/" + scanner.BlockTextLengths[i])) + ";" + string.Join(",", scanner.Headings.Select(h => h.BlockIndex + ":" + h.Level));

    /// <summary>A long comment line before the note moves it across the 4 KB .. 32 KB read boundaries of the rescans: the document is the same however the pieces fall.</summary>
    private static void RunBoundaries(string directory)
    {
        int bad = 0, count = 0;
        foreach (int boundary in new[] { 4096, 8192, 16384, 32768 })
        foreach (int shift in Enumerable.Range(-12, 14))
        {
            string prefix = "&UniqueId=1&PutNewLinesAsIs=0&TextSource=Src&ForcedBookName=Nm\r\n$ T\r\n// ";
            string text = prefix + new string('c', boundary + shift - prefix.Length - 2) + "\r\n" + NoteOne + "\r\n\r\n# H\r\ntext\r\n";
            string path = Path.Combine(directory, $"torat-emet-notes-boundary-{boundary}-{shift}-v1.txt");
            File.WriteAllBytes(path, Hebrew1255.GetBytes(text));
            foreach (WordIndexMode mode in new[] { WordIndexMode.InMemory, WordIndexMode.Sidecar })
            {
                var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(directory, "index-cache"), ForcedMode = mode };
                if (File.Exists(policy.IndexPathFor(path))) File.Delete(policy.IndexPathFor(path));
                for (int pass = 0; pass < (mode == WordIndexMode.Sidecar ? 2 : 1); pass++)
                {
                    using var document = (MarkupWindowedDocument)WordWindowedDocumentOpener.Open(path, policy);
                    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(0, document.BlockCount);
                    count++;
                    if (!(blocks.Count == 5 && Plain(blocks[0].Html).EndsWith(NoteOne, StringComparison.Ordinal) && Plain(blocks[3].Html) == "H" && Plain(blocks[4].Html) == "text")) bad++;
                }
            }
        }
        TestReport.Check(bad == 0, $"note across read boundaries: {count} opens, {bad} wrong");
    }
}
