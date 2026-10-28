using WordDocumentRandomAccess.DocBinary;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// .doc list labels: opens each binary .doc, reports its list tables (LSTF / LFO counts), how many
/// blocks the index gave a label, and checks that rendering those blocks emits the list marker.
/// When none of the given files has a list and a fallback list of paths is supplied, the first .doc
/// files in it are scanned (counts only, no names) until one with labels is found.
/// </summary>
public static class DocBinaryListLabelsTest
{
    public static void Run(IEnumerable<string> paths, string? fallbackListPath, int fallbackMaxFiles = 400)
    {
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        int filesWithLabels = 0;
        foreach (string path in paths)
        {
            TestReport.Step($"List labels: {Path.GetFileName(path)}");
            if (Inspect(path, policy, verbose: true) > 0) filesWithLabels++;
        }
        if (filesWithLabels == 0 && fallbackListPath is not null && File.Exists(fallbackListPath))
        {
            TestReport.Step($"List labels: scanning up to {fallbackMaxFiles} .doc files from the corpus list for one with lists");
            int scanned = 0;
            foreach (string path in File.ReadAllLines(fallbackListPath))
            {
                if (!path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) || !File.Exists(path)) continue;
                if (++scanned > fallbackMaxFiles) break;
                int labelled;
                try { labelled = Inspect(path, policy, verbose: false); }
                catch (Exception) { continue; }                   // detect / open failures belong to the corpus audit
                if (labelled == 0) continue;
                Console.WriteLine($"   corpus file #{scanned}: {labelled:N0} labelled blocks");
                filesWithLabels++;
                break;
            }
        }
        TestReport.Check(filesWithLabels > 0, "at least one .doc rendered list markers");
    }

    /// <summary>Returns the number of blocks with a label; checks the rendered HTML of the first few of them.</summary>
    private static int Inspect(string path, WordIndexPolicy policy, bool verbose)
    {
        if (WordWindowedDocumentOpener.DetectFormat(path) != WordWindowedDocumentOpener.WordFileFormat.DocBinary) return 0;
        using var document = DocBinaryWindowedDocument.Open(path, policy);
        DocBinaryBlockIndex index = document.BlockIndex;
        int labelled = 0;
        var firstLabelled = new List<int>();
        for (int block = 0; block < index.BlockCount; block++)
        {
            if (index.BlockListLabelIds[block] < 0) continue;
            labelled++;
            if (firstLabelled.Count < 8) firstLabelled.Add(block);
        }
        if (verbose)
        {
            Console.WriteLine($"   list tables: {document.ListDefinitions.Lists.Count} LSTF, {document.ListDefinitions.Overrides.Count} LFO; " +
                              $"{labelled:N0} of {index.BlockCount:N0} blocks labelled, {index.ListLabels.Length} distinct labels");
            int pieces = document.PieceTable.Pieces.Count, piecesWithPrm = document.PieceTable.Pieces.Count(p => p.PrmGrpprl is not null);
            Console.WriteLine($"   pieces: {pieces:N0}, with Prm: {piecesWithPrm:N0}, Prc grpprls in Clx: {document.PieceTable.PrcGrpprls.Count}");
            var frequency = new Dictionary<int, int>();
            foreach (int id in index.BlockListLabelIds) if (id >= 0) frequency[id] = frequency.TryGetValue(id, out int n) ? n + 1 : 1;
            IEnumerable<string> top = frequency.OrderByDescending(f => f.Value).Take(8).Select(f => $"{Escape(index.ListLabels[f.Key])} x{f.Value}");
            Console.WriteLine($"   most frequent labels: {string.Join(", ", top)}");
            string twinPath = Path.ChangeExtension(path, ".docx");
            if (File.Exists(twinPath))
            {
                using var twin = Docx.DocxWindowedDocument.Open(twinPath, policy);
                int twinLabelled = twin.Index.BlockListLabelIds.Count(id => id >= 0);
                Console.WriteLine($"   .docx twin: {twinLabelled:N0} of {twin.BlockCount:N0} blocks labelled, {twin.Index.ListLabels.Length} distinct labels");
            }
            foreach (int block in firstLabelled.Take(4))
            {
                string html = document.RenderBlocks(block, 1)[0].Html;
                string text = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", "");
                Console.WriteLine($"   block {block}: {Escape(text.Length > 60 ? text.Substring(0, 60) : text)}");
            }
        }
        if (labelled == 0) return 0;
        int rendered = 0;
        foreach (int block in firstLabelled)
        {
            string html = document.RenderBlocks(block, 1)[0].Html;
            if (html.Contains("class=\"list-marker\"", StringComparison.Ordinal)) rendered++;
        }
        TestReport.Check(rendered == firstLabelled.Count, $"{rendered} of {firstLabelled.Count} sampled labelled blocks render a list marker");
        return labelled;
    }

    /// <summary>Non-ASCII as \uXXXX so the console (and filters in front of it) see plain ASCII.</summary>
    private static string Escape(string text)
    {
        var builder = new System.Text.StringBuilder(text.Length);
        foreach (char c in text) { if (c < 0x7F && c >= ' ') builder.Append(c); else builder.Append("\\u").Append(((int)c).ToString("X4")); }
        return builder.ToString();
    }
}
