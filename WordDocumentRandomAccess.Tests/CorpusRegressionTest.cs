namespace WordDocumentRandomAccess.Tests;

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
        (@"libreoffice\sw\qa\extras\ooxmlimport\data\math-malformed_xml.docx", typeof(InvalidDataException), "not well-formed"),
        (@"libreoffice\sw\qa\core\data\odt\fail\tdf102018.odt", typeof(InvalidDataException), "styles.xml"),
        (@"apache-poi\52117.doc", typeof(NotSupportedException), "Word 6.0 / Word 95"),
        (@"apache-poi\word2.doc", typeof(NotSupportedException), "Word for Windows 2.0"),
        (@"apache-tika\testWORDML_tableInsideParagraph.doc", typeof(NotSupportedException), "Word 2003 XML"),
        (@"libreoffice\sw\qa\extras\ooxmlexport\data\Encrypted_MSO2007_abc.docx", typeof(NotSupportedException), "Encrypted"),
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
}
