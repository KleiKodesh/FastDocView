namespace FastDocView.Tests;

/// <summary>
/// One check per corpus failure fixed on 2026-10-28 (test-data/): files that crashed or were
/// misdetected now open and render, and files that cannot be read fail with a typed, precise error.
/// Runs with --corpus when test-data is the corpus folder.
/// </summary>
public static class CorpusRegressionTest
{
    private static readonly string[] MustOpen =
    [
        @"apache-tika\testODTStyles3.odt",                                   // deflated mimetype entry
        @"libreoffice\sw\qa\extras\layout\data\tdf134298.ott",                // mimetype not the first entry
        @"libreoffice\sw\qa\extras\ooxmlexport\data\tdf104713_undefinedStyles.docx",   // main part word/trial.xml via _rels/.rels
        @"apache-poi\MultipleBodyBug.docx",                                   // three w:body elements
        @"libreoffice\sw\qa\core\text\data\ruby.fodt",                        // undeclared namespace prefix
        @"libreoffice\sw\qa\extras\odfimport\data\linked_ole.fodt",           // root tag attributes on separate lines
        @"apache-poi\testCroppedPictures.doc",                                // STD running past the style sheet
        @"libreoffice\sw\qa\core\data\ww8\pass\forcepoint50-grfanchor-1.doc", // FKP crun larger than the page holds
        @"libreoffice\sw\qa\core\data\ww8\pass\ofz38011-1.doc",
        @"libreoffice\sw\qa\core\data\ww8\pass\ofz18534-1.doc",
        @"libreoffice\sw\qa\core\data\ww8\pass\hang-8.doc",                   // piece-table CPs out of order
        @"libreoffice\sw\qa\core\data\ww8\pass\ofz47205-1.doc",
        @"libreoffice\sw\qa\core\data\ww8\pass\ofz18554-1.doc",               // note text range past the stream
        @"apache-poi\clusterfuzz-testcase-minimized-POIHWPFFuzzer-4951943183990784.doc",
        @"apache-poi\clusterfuzz-testcase-minimized-POIHWPFFuzzer-6610789829836800.doc",
        @"apache-poi\clusterfuzz-testcase-minimized-POIHWPFFuzzer-5440721166139392.doc",   // bin table outside the table stream: read without it
    ];

    private static readonly (string file, Type exception, string messagePart)[] MustFailClearly =
    [
        (@"apache-poi\clusterfuzz-testcase-POIHWPFFuzzer-5696094627495936.doc", typeof(InvalidDataException), "name length"),
        (@"libreoffice\sw\qa\core\data\ww8\pass\crash-2.doc", typeof(InvalidDataException), "piece table"),
        (@"libreoffice\sw\qa\core\data\ww8\fail\hang-2.doc", typeof(InvalidDataException), "piece table"),
        (@"libreoffice\sw\qa\core\data\odt\fail\tdf102018.odt", typeof(InvalidDataException), "office:text body"),   // styles.xml is repaired (content before the first error); the body never closes
        (@"apache-poi\52117.doc", typeof(NotSupportedException), "Word 6.0 / Word 95"),
        (@"apache-poi\word2.doc", typeof(NotSupportedException), "Word for Windows 2.0"),
        (@"apache-tika\testWORDML_tableInsideParagraph.doc", typeof(NotSupportedException), "Word 2003 XML"),
        (@"libreoffice\sw\qa\extras\ooxmlexport\data\Encrypted_MSO2007_abc.docx", typeof(WordDocumentEncryptedException), "Encrypted"),
    ];

    /// <summary>
    /// Malformed .docx parts keep what precedes the first error (WordDocumentXml.LoadUntilFirstError —
    /// LibreOffice's repair mode); each file: (relative path, text the first block must show, text it must not).
    /// </summary>
    private static readonly (string file, string mustShow, string? mustNotShow)[] MalformedDocxKeepContentBeforeError =
    [
        // <m:t>a</m:sPre>: the paragraph and its math up to "a" are kept (math itself is not rendered yet)
        (@"libreoffice\sw\qa\extras\ooxmlimport\data\math-malformed_xml.docx", "<p", null),
        // duplicate w:val on w:jc inside w:pPr: the paragraph exists, its run "test" comes after the error
        (@"libreoffice\sw\qa\extras\ooxmlimport\data\tdf165348_broken_package.docx", "<p", "test"),
        // <!DOCTYPE> with an external entity: the scanner steps over it; &testent; is never resolved
        (@"apache-poi\ExternalEntityInText.docx", "<p", "poi.apache.org"),
    ];

    public static void Run(string testDataFolder)
    {
        TestReport.Step("Corpus regressions (fixed 2026-10-28)");
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "corpus-index-cache") };
        foreach (string relative in MustOpen)
        {
            string path = Path.Combine(testDataFolder, relative);
            string? error = null;
            try
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                document.RenderBlocks(0, 40);
            }
            catch (Exception exception) { error = $"{exception.GetType().Name}: {exception.Message}"; }
            TestReport.Check(error is null, $"opens + renders: {relative}{(error is null ? "" : "  — " + error)}");
        }
        foreach ((string relative, string mustShow, string? mustNotShow) in MalformedDocxKeepContentBeforeError)
        {
            string path = Path.Combine(testDataFolder, relative);
            string? html = null, error = null;
            try
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                html = string.Concat(document.RenderBlocks(0, 40).Select(block => block.Html));
            }
            catch (Exception exception) { error = $"{exception.GetType().Name}: {exception.Message}"; }
            TestReport.Check(html is not null && html.Contains(mustShow) && (mustNotShow is null || !html.Contains(mustNotShow)),
                             $"malformed XML keeps the content before the error: {relative}{(error is null ? "  — " + Shorten(html) : "  — " + error)}");
        }
        foreach ((string relative, Type expected, string messagePart) in MustFailClearly)
        {
            string path = Path.Combine(testDataFolder, relative);
            Exception? caught = null;
            try { using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy); document.RenderBlocks(0, 40); }
            catch (Exception exception) { caught = exception; }
            TestReport.Check(caught is not null && caught.GetType() == expected && caught.Message.Contains(messagePart),
                             $"fails with {expected.Name} \"{messagePart}\": {relative}{(caught is null ? "  — opened" : "  — " + caught.GetType().Name + ": " + caught.Message)}");
        }
    }

    /// <summary>Open failures a host shows as a translated message: (relative path, expected class).</summary>
    private static readonly (string file, WordDocumentOpenFailure failure)[] OpenFailureClasses =
    [
        (@"libreoffice\sw\qa\extras\ooxmlexport\data\Encrypted_MSO2007_abc.docx", WordDocumentOpenFailure.Encrypted),
        (@"apache-poi\PasswordProtected.doc", WordDocumentOpenFailure.Encrypted),
        (@"apache-tika\testODTEncrypted.odt", WordDocumentOpenFailure.Encrypted),
        (@"libreoffice\sw\qa\core\data\ww8\pass\crash-2.doc", WordDocumentOpenFailure.Corrupt),
        (@"apache-poi\word2.doc", WordDocumentOpenFailure.UnsupportedFormat),
        (@"apache-poi\no-such-file-exists.docx", WordDocumentOpenFailure.NotFound),
    ];

    /// <summary>Open failures classify for the host (encrypted, corrupt, unsupported, missing).</summary>
    public static void RunOpenFailureClasses(string testDataFolder)
    {
        TestReport.Step("Open failures classified for the host");
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "corpus-index-cache") };
        foreach ((string relative, WordDocumentOpenFailure expected) in OpenFailureClasses)
        {
            Exception? caught = null;
            try { using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(Path.Combine(testDataFolder, relative), policy); document.RenderBlocks(0, 40); }
            catch (Exception exception) { caught = exception; }
            WordDocumentOpenFailure? actual = caught is null ? null : WordDocumentOpenFailures.Classify(caught);
            TestReport.Check(actual == expected, $"classified {expected}: {relative}  — {(caught is null ? "opened" : actual + " / " + caught.GetType().Name + ": " + caught.Message)}");
        }
    }

    /// <summary>
    /// One bad block never fails its window: a real document's window is rendered through the guard with block 2 made
    /// to throw a content error - the other blocks render, block 2 is the placeholder, reported once; a changed file
    /// (WordDocumentSourceChangedException) still escapes.
    /// </summary>
    public static void RunMalformedBlock(string testDataFolder)
    {
        TestReport.Step("A malformed block becomes a placeholder, its window still renders");
        string path = Path.Combine(testDataFolder, @"apache-poi\MultipleBodyBug.docx");
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "corpus-index-cache") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
        int count = Math.Min(document.BlockCount, 6);
        var guard = new WordDocumentBlockRenderGuard(path);
        int reports = 0;
        void OnFailed(string file, int block, Exception cause) { if (file == path) reports++; }
        WordDocumentBlockRenderGuard.BlockRenderFailed += OnFailed;
        try
        {
            IReadOnlyList<WordDocumentBlockHtml> RenderWithBadBlock(int first, int blocks)
            {
                if (first <= 2 && 2 < first + blocks) throw new InvalidDataException("synthetic malformed block 2");
                return document.RenderBlocks(first, blocks);
            }
            for (int round = 0; round < 2; round++)
            {
                IReadOnlyList<WordDocumentBlockHtml> window = guard.Render(0, count, document.BlockCount, RenderWithBadBlock, _ => WordDocumentBlockKind.Paragraph);
                TestReport.Check(window.Count == count && window.Select(b => b.BlockIndex).SequenceEqual(Enumerable.Range(0, count)),
                                 $"round {round}: every block of the window is returned ({window.Count}/{count})");
                TestReport.Check(count > 2 && window[2].Html == WordDocumentBlockRenderGuard.PlaceholderHtml
                                 && window.Where(b => b.BlockIndex != 2).All(b => b.Html != WordDocumentBlockRenderGuard.PlaceholderHtml),
                                 $"round {round}: only block 2 is the placeholder");
            }
            TestReport.Check(reports == 1, $"the bad block is reported once ({reports})");
            bool escaped = false;
            try { guard.Render(0, count, document.BlockCount, (_, _) => throw new WordDocumentSourceChangedException(path), _ => WordDocumentBlockKind.Paragraph); }
            catch (WordDocumentSourceChangedException) { escaped = true; }
            TestReport.Check(escaped, "a changed file is never turned into placeholders");
        }
        finally { WordDocumentBlockRenderGuard.BlockRenderFailed -= OnFailed; }
    }

    private static string Shorten(string? html) => html is null ? "" : html.Length > 300 ? html.Substring(0, 300) + "…" : html;
}
