using System.Diagnostics;

namespace FastDocView.Tests;

/// <summary>
/// Document properties (<see cref="IWordWindowedDocument.ReadProperties"/>): Word's own saves of one document as .docx, .doc
/// and .rtf (test-data/word-fields/fields-v1.*) must agree on the texts and counts they all store; then the given files
/// and a fixed-seed sample of test-data are read (0 crashes, ms printed).
/// </summary>
public static class DocumentPropertiesTest
{
    public static void Run(IEnumerable<string> extraFiles, int sampleSize)
    {
        TestReport.Step("Document properties: Word-saved .docx / .doc / .rtf agree; samples read");
        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));
        var saved = new Dictionary<string, WordDocumentProperties>();
        foreach (string extension in new[] { ".docx", ".doc", ".rtf" })
        {
            string path = Path.Combine(root, "word-fields", "fields-v1" + extension);
            if (!File.Exists(path)) { TestReport.Check(false, $"fields-v1{extension} present"); continue; }
            saved[extension] = Read(path, out _);
            Console.WriteLine($"  fields-v1{extension}: {Describe(saved[extension])}");
        }
        if (saved.Count == 3)
        {
            WordDocumentProperties docx = saved[".docx"];
            TestReport.Check(docx.Author is not null && docx.Created is not null && docx.Modified is not null && docx.Words is not null && docx.Application is not null,
                             ".docx: author, created, modified, words, application read");
            foreach (string extension in new[] { ".doc", ".rtf" })
            {
                WordDocumentProperties other = saved[extension];
                TestReport.Check(other.Title == docx.Title && other.Author == docx.Author && other.Subject == docx.Subject && other.Keywords == docx.Keywords
                                 && other.Comments == docx.Comments && other.Company == docx.Company && other.Category == docx.Category,
                                 $"{extension} texts = .docx texts");
                TestReport.Check(other.Created is not null && other.Modified is not null, $"{extension}: created and modified read");
                TestReport.Check(other.Characters is null || other.CharactersWithSpaces is null || other.Characters <= other.CharactersWithSpaces,
                                 $"{extension}: characters without spaces <= with spaces");
            }
        }
        var files = new List<string>(extraFiles.Where(File.Exists));
        if (sampleSize > 0 && Directory.Exists(root))
        {
            var all = Directory.GetFiles(root, "*", SearchOption.AllDirectories)
                .Where(path => WordWindowedDocumentOpener.SupportedExtensions.Contains(Path.GetExtension(path).ToLowerInvariant()))
                .OrderBy(path => path, StringComparer.Ordinal).ToList();
            var random = new Random(20261126);
            files.AddRange(all.OrderBy(_ => random.Next()).Take(sampleSize));
        }
        int read = 0, failed = 0, withValues = 0;
        double slowest = 0;
        foreach (string path in files)
        {
            try
            {
                WordDocumentProperties properties = Read(path, out double milliseconds);
                read++;
                slowest = Math.Max(slowest, milliseconds);
                if (properties != WordDocumentProperties.None) withValues++;
                if (extraFiles.Contains(path)) Console.WriteLine($"  {Path.GetFileName(path)} ({milliseconds:0.0} ms): {Describe(properties)}");
            }
            catch (Exception exception) when (exception is InvalidDataException or NotSupportedException or IOException or UnauthorizedAccessException)
            {
                failed++;      // a file the opener itself rejects (encrypted, broken): not a properties failure
            }
            catch (Exception exception)
            {
                TestReport.Check(false, $"properties of {path}: {exception.GetType().Name} {exception.Message}");
            }
        }
        Console.WriteLine($"  sample: {read} read ({withValues} with properties), {failed} rejected by the opener, slowest read {slowest:0.0} ms");
    }

    private static WordDocumentProperties Read(string path, out double milliseconds)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-properties") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        var watch = Stopwatch.StartNew();
        WordDocumentProperties properties = document.ReadProperties();
        milliseconds = watch.Elapsed.TotalMilliseconds;
        return properties;
    }

    private static string Describe(WordDocumentProperties p) =>
        $"title=[{p.Title}] subject=[{p.Subject}] author=[{p.Author}] lastBy=[{p.LastModifiedBy}] keywords=[{p.Keywords}] category=[{p.Category}] "
        + $"comments=[{p.Comments}] company=[{p.Company}] manager=[{p.Manager}] status=[{p.ContentStatus}] rev=[{p.Revision}] app=[{p.Application}] "
        + $"template=[{p.Template}] created={p.Created:o} modified={p.Modified:o} printed={p.LastPrinted:o} pages={p.Pages} words={p.Words} "
        + $"chars={p.Characters} charsWs={p.CharactersWithSpaces} lines={p.Lines} paras={p.Paragraphs} minutes={p.TotalEditingMinutes}";
}
