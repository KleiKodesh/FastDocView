using System.Security.Cryptography;
using System.Text;
using FastDocView.DocBinary;
using FastDocView.Odt;
using FastDocView.Rtf;

namespace FastDocView.Tests;

/// <summary>
/// Writes everything a reader can observe about a document into one text file (a "golden dump"): every block's
/// kind and rendered HTML hash (notes included), the headings, note references, sections, bookmarks (every name
/// resolved through FindBookmark) and the hits of one search. Two dumps of the same file from different index
/// layouts must be byte-identical — the proof that an index change lost nothing. The text-length estimates go to
/// a second file (they may legitimately differ for spans the compact table has not scanned).
/// </summary>
public static class GoldenDumpTest
{
    public static void Run(string outputDirectory, IEnumerable<string> paths)
    {
        Directory.CreateDirectory(outputDirectory);
        foreach (string path in paths)
        {
            TestReport.Step($"Golden dump: {Path.GetFileName(path)}");
            string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-golden");
            var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            var watch = System.Diagnostics.Stopwatch.StartNew();
            string dumpText, lengthsText;
            using (IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy))
            {
                if (document is OdtWindowedDocument odt) odt.CompleteIndex();
                Console.WriteLine($"   cold open {watch.ElapsedMilliseconds} ms, {document.BlockCount:N0} blocks, sidecar {(File.Exists(indexPath) ? new FileInfo(indexPath).Length : 0):N0} bytes");
                (dumpText, lengthsText) = Dump(document);
            }
            // the same from a reopen over the saved sidecar: what the file stores must give the same observable output
            using (IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(path, policy))
            {
                if (reopened is OdtWindowedDocument odt) odt.CompleteIndex();
                (string reopenedDump, _) = Dump(reopened);
                TestReport.Check(reopenedDump == dumpText, "a reopen from the sidecar dumps the same");
            }
            string stem = Path.Combine(outputDirectory, Path.GetFileName(path));
            File.WriteAllText(stem + ".golden.txt", dumpText);
            File.WriteAllText(stem + ".lengths.txt", lengthsText);
            Console.WriteLine($"   wrote {stem}.golden.txt ({dumpText.Length:N0} chars), total {watch.ElapsedMilliseconds} ms");
        }
    }

    private static (string dump, string lengths) Dump(IWordWindowedDocument document)
    {
        {
            var dump = new StringBuilder();
            var lengths = new StringBuilder();
            dump.Append("blocks ").Append(document.BlockCount).Append(" complete ").Append(document.IsIndexComplete).Append(" known ").Append(document.KnownBlockCount).Append(" rtl ").Append(document.IsRightToLeft).Append('\n');
            dump.Append("stylesheet ").Append(Hash(document.Stylesheet)).Append('\n');
            string? query = null;
            const int window = 64;
            for (int first = 0; first < document.BlockCount; first += window)
            {
                foreach (WordDocumentBlockHtml block in document.RenderBlocks(first, window))
                {
                    dump.Append(block.BlockIndex).Append(' ').Append((int)block.Kind).Append(' ').Append(Hash(block.Html));
                    foreach (WordDocumentNoteHtml note in block.Notes) dump.Append(" n").Append((int)note.Kind).Append(':').Append(note.Number).Append(':').Append(Hash(note.Label + "\u0001" + note.Html));
                    dump.Append('\n');
                    query ??= QueryFrom(block.Html);
                }
            }
            foreach (WordDocumentHeading heading in document.Headings) dump.Append("heading ").Append(heading.BlockIndex).Append(' ').Append(heading.Level).Append(' ').Append(Hash(heading.Text)).Append('\n');
            dump.Append("notes ").Append(document.NoteCounts.footnotes).Append(' ').Append(document.NoteCounts.endnotes).Append('\n');
            foreach (WordDocumentNoteReference reference in document.NoteReferences) dump.Append("noteref ").Append((int)reference.Kind).Append(' ').Append(reference.Number).Append(' ').Append(reference.BlockIndex).Append('\n');
            int noteCount = document.NoteReferences.Count;
            for (int first = 0; first < noteCount; first += 200)
                foreach (WordDocumentNoteHtml note in document.RenderNotes(first, 200)) dump.Append("note ").Append((int)note.Kind).Append(' ').Append(note.Number).Append(' ').Append(Hash(note.Label + "\u0001" + note.Html)).Append('\n');
            foreach (WordDocumentSection section in document.Sections) dump.Append("section ").Append(section.FirstBlockIndex).Append(' ').Append(section.TextWidthTwips).Append('\n');
            foreach (string name in BookmarkNames(document).Distinct(StringComparer.Ordinal))
                dump.Append("bookmark ").Append(Hash(name)).Append(' ').Append(document.FindBookmark(name)?.ToString() ?? "-").Append('\n');
            dump.Append("bookmark-missing ").Append(document.FindBookmark("no-such-bookmark-golden")?.ToString() ?? "-").Append('\n');
            if (query is not null)
            {
                dump.Append("query ").Append(Hash(query)).Append('\n');
                int from = 0;
                while (true)
                {
                    WordDocumentSearchBatch batch = document.Search(query, WordDocumentSearchOptions.Default, from, 5000);
                    for (int i = 0; i < batch.HitBlocks.Length; i++) dump.Append("hit ").Append(batch.HitBlocks[i]).Append(' ').Append(batch.HitCounts[i]).Append('\n');
                    if (batch.ReachedEnd || batch.NextBlock <= from) break;
                    from = batch.NextBlock;
                }
            }
            dump.Append("changed ").Append(document.HasSourceChanged).Append('\n');
            for (int block = 0; block < document.BlockCount; block++) lengths.Append(document.EstimatedTextLength(block)).Append('\n');
            return (dump.ToString(), lengths.ToString());
        }
    }

    private static IEnumerable<string> BookmarkNames(IWordWindowedDocument document) => document switch
    {
        DocBinaryWindowedDocument doc => doc.Bookmarks.Names,
        OdtWindowedDocument odt => odt.Index.Bookmarks.Names,
        RtfWindowedDocument rtf => rtf.Index.Bookmarks.Names,
        Docx.DocxWindowedDocument docx => docx.Index.BookmarkNames,
        _ => [],
    };

    /// <summary>The first word of at least three letters in the block's text — a deterministic search query for any file.</summary>
    private static string? QueryFrom(string html)
    {
        string text = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ");
        text = System.Net.WebUtility.HtmlDecode(text);
        int start = -1;
        for (int i = 0; i <= text.Length; i++)
        {
            bool letter = i < text.Length && char.IsLetter(text[i]);
            if (letter && start < 0) start = i;
            else if (!letter && start >= 0)
            {
                if (i - start >= 3) return text.Substring(start, Math.Min(i - start, 6));
                start = -1;
            }
        }
        return null;
    }

    private static string Hash(string text)
    {
        using var sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(text));
        var builder = new StringBuilder(24);
        for (int i = 0; i < 12; i++) builder.Append(hash[i].ToString("x2"));
        return builder.ToString();
    }
}
