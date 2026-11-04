namespace FastDocView.Tests;

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
                // Another process may hold the freshly written files for a moment (an antivirus scanning them): a failed
                // replace names the holders (Restart Manager). Ours must never be one of them — then the replace is
                // retried while only others hold the files, as Word itself retries a save.
                bool canReplace = false, isHeldByUs = false;
                var waited = System.Diagnostics.Stopwatch.StartNew();
                for (int attempt = 1; ; attempt++)
                {
                    try { File.Replace(temporary, copy, null); canReplace = true; break; }
                    catch (IOException error)
                    {
                        List<(int processId, string name)> holders = FileLockHolders.Of(copy, temporary);
                        isHeldByUs |= holders.Any(holder => holder.processId == Environment.ProcessId);
                        Console.WriteLine($"   replace attempt {attempt} refused after {waited.ElapsedMilliseconds} ms (0x{error.HResult:X8} {error.Message.Trim()}); "
                            + $"held by: {(holders.Count == 0 ? "nobody now" : string.Join(", ", holders.Select(holder => $"{holder.name} pid {holder.processId}")))}");
                        if (isHeldByUs || waited.ElapsedMilliseconds > 10_000) break;
                        Thread.Sleep(100);
                    }
                }
                TestReport.Check(!isHeldByUs, $"{mode}: our process holds no handle that blocks Word's save");
                TestReport.Check(canReplace, $"{mode}: the file can be replaced (Word's save) while our document is open");
                TestReport.Check(document.HasSourceChanged || mode == WordIndexMode.InMemory, $"{mode}: the document reports that the file changed");

                // 3. A document that still reads the file must refuse to render the new version with the old index.
                bool isRefused = false;
                try { document.RenderBlocks(0, 5); }
                catch (WordDocumentSourceChangedException) { isRefused = true; }
                // in-memory mode holds the text only for zip packages, whose parts it inflates once (DocxWindowedDocument / OdtWindowedDocument
                // OpenFromStream); .doc, flat .fodt, RTF, text and MHTML always read windows from the file
                bool readsFile = !(mode == WordIndexMode.InMemory && WordWindowedDocumentOpener.DetectFormat(sample)
                                   is WordWindowedDocumentOpener.WordFileFormat.Docx or WordWindowedDocumentOpener.WordFileFormat.OpenDocument);
                TestReport.Check(isRefused == readsFile, readsFile ? $"{mode}: render after Word's save is refused (host reopens)" : $"{mode}: in-memory body keeps rendering the opened version");

                // 4. Reopening picks up the new version.
                using IWordWindowedDocument reopened = WordWindowedDocumentOpener.Open(copy, new WordIndexPolicy { CacheDirectory = policy.CacheDirectory, ForcedMode = mode });
                TestReport.Check(!reopened.HasSourceChanged && reopened.RenderBlocks(0, 5).Count > 0, $"{mode}: reopening renders the saved version");
            }
            File.Delete(copy);
        }
    }

    /// <summary>
    /// The environment check behind the retry above: the same copy + save-by-replace <paramref name="rounds"/> times with
    /// no document of ours open (control) and with one open, counting refusals and naming who held the files.
    /// </summary>
    public static void Probe(string sample, int rounds)
    {
        string copy = Path.Combine(TestReport.OutputDirectory, "coexistence-probe" + Path.GetExtension(sample));
        var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-coexistence"), ForcedMode = WordIndexMode.Sidecar };
        foreach (bool withDocument in new[] { false, true })
        {
            TestReport.Step($"Replace probe, {rounds} rounds, {(withDocument ? "our document open" : "control: nothing of ours open")}: {Path.GetFileName(sample)}");
            int refused = 0, refusedByUs = 0;
            var holderNames = new Dictionary<string, int>();
            var errorKinds = new Dictionary<string, int>();
            for (int round = 0; round < rounds; round++)
            {
                File.Copy(sample, copy, overwrite: true);
                IWordWindowedDocument? document = withDocument ? WordWindowedDocumentOpener.Open(copy, policy) : null;
                try
                {
                    document?.RenderBlocks(0, 40);
                    string temporary = copy + ".save.tmp";
                    File.Copy(sample, temporary, overwrite: true);
                    using (var append = new FileStream(temporary, FileMode.Append)) append.WriteByte(0);
                    try { File.Replace(temporary, copy, null); }
                    catch (IOException error)
                    {
                        refused++;
                        string kind = $"0x{error.HResult:X8} {error.Message.Trim()}";
                        errorKinds[kind] = errorKinds.GetValueOrDefault(kind) + 1;
                        List<(int processId, string name)> holders = FileLockHolders.Of(copy, temporary);
                        if (holders.Any(holder => holder.processId == Environment.ProcessId)) refusedByUs++;
                        foreach ((int _, string name) in holders) holderNames[name] = holderNames.GetValueOrDefault(name) + 1;
                        if (holders.Count == 0) holderNames["(released before the query)"] = holderNames.GetValueOrDefault("(released before the query)") + 1;
                        File.Delete(temporary);
                    }
                }
                finally { document?.Dispose(); }
                if ((round + 1) % 10 == 0) Console.WriteLine($"   {round + 1}/{rounds}: {refused} refused so far");
            }
            Console.WriteLine($"   refused {refused} of {rounds}, with our process among the holders {refusedByUs}; holders: {(holderNames.Count == 0 ? "-" : string.Join(", ", holderNames.Select(pair => $"{pair.Key} x{pair.Value}")))}"
                + (errorKinds.Count == 0 ? "" : $"; errors: {string.Join(" | ", errorKinds.Select(pair => $"{pair.Key} x{pair.Value}"))}"));
            TestReport.Check(refusedByUs == 0, $"{(withDocument ? "document open" : "control")}: our process never held the files when a replace was refused");
        }
        File.Delete(copy);
    }
}
