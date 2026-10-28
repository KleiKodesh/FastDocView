namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Our reader must never stop Word from opening or saving a document it has open. On a copy of each
/// sample (the samples themselves are never touched): open it with the library and render windows,
/// then do what Word does — open the file exclusively for writing, and save by writing a temp file
/// and replacing the original — and check both succeed, that the document reports the change,
/// and that it still renders (from the version it was opened on).
/// </summary>
public static class WordCoexistenceTest
{
    public static void Run(IEnumerable<string> samplePaths)
    {
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-coexistence") };
        foreach (string sample in samplePaths)
        {
            TestReport.Step($"Word can open and save while we read: {Path.GetFileName(sample)}");
            string copy = Path.Combine(TestReport.OutputDirectory, "coexistence" + Path.GetExtension(sample));
            File.Copy(sample, copy, overwrite: true);

            foreach (WordIndexMode mode in new[] { WordIndexMode.Sidecar, WordIndexMode.InMemory })
            {
                File.Copy(sample, copy, overwrite: true);
                File.SetLastWriteTimeUtc(copy, DateTime.UtcNow.AddMinutes(-5));
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(copy, new WordIndexPolicy { CacheDirectory = policy.CacheDirectory, ForcedMode = mode });
                document.RenderBlocks(0, 40);
                document.RenderBlocks(document.BlockCount / 2, 40);
                TestReport.Check(!document.HasSourceChanged, $"{mode}: unchanged file reports no change");

                // 1. Word opens the file for editing with no sharing for writers.
                bool canOpenExclusively;
                try { using var exclusive = new FileStream(copy, FileMode.Open, FileAccess.ReadWrite, FileShare.None); canOpenExclusively = true; }
                catch (IOException) { canOpenExclusively = false; }
                TestReport.Check(canOpenExclusively, $"{mode}: the file can be opened exclusively for writing while our document is open");

                // 2. Word saves: write a temp file next to it, then replace the original.
                string temporary = copy + ".save.tmp";
                File.Copy(sample, temporary, overwrite: true);
                using (var append = new FileStream(temporary, FileMode.Append)) append.WriteByte(0);   // a different version
                bool canReplace;
                try { File.Replace(temporary, copy, null); canReplace = true; }
                catch (IOException) { canReplace = false; }
                TestReport.Check(canReplace, $"{mode}: the file can be replaced (Word's save) while our document is open");
                TestReport.Check(document.HasSourceChanged || mode == WordIndexMode.InMemory, $"{mode}: the document reports that the file changed");

                // 3. A document that still reads the file must refuse to render the new version with the old index.
                bool isRefused = false;
                try { document.RenderBlocks(0, 5); }
                catch (WordDocumentSourceChangedException) { isRefused = true; }
                bool readsFile = !(mode == WordIndexMode.InMemory && Path.GetExtension(copy) == ".docx");
                TestReport.Check(isRefused == readsFile, readsFile ? $"{mode}: render after Word's save is refused (host reopens)" : $"{mode}: in-memory body keeps rendering the opened version");

                // 4. Reopening picks up the new version.
                using IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(copy, new WordIndexPolicy { CacheDirectory = policy.CacheDirectory, ForcedMode = mode });
                TestReport.Check(!reopened.HasSourceChanged && reopened.RenderBlocks(0, 5).Count > 0, $"{mode}: reopening renders the saved version");
            }
            File.Delete(copy);
        }
    }
}
