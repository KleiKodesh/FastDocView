namespace FastDocView.Tests;

/// <summary>
/// Word 6.0 / 95 .doc files (nFib 101-105) read through the DocBinary reader's version branches: real files from test-data
/// (Apache POI, LibreOffice's sw/qa ww6 / ww8export data), their text compared with LibreOffice's text export of the same files
/// (2026-11-17: Word95, 57843, Word6_sections2, Bug50955 word-for-word equal), plus a code-page case built from one of them
/// (Word 16 cannot save Word 6/95 and test-data has no right-to-left Word 6 file).
/// </summary>
public static class Word6DocumentsTest
{
    public static void Run()
    {
        TestReport.Step("Word 6.0 / 95 .doc: FIB, 8-bit and UTF-16 text, code pages, one-byte sprms, styles, fonts, tables");
        string testData = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word6") };

        string Html(string relative)
        {
            string path = Path.Combine(testData, relative);
            if (!File.Exists(path)) return "";
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            return string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        }

        // Word 95 (nFib 101): character sprms 98 (ico), 85 (bold), 104 (iss) as Word 97's sprmCIco / sprmCFBold / sprmCIss
        string word95 = Html(@"apache-poi\Word95.doc");
        TestReport.Check(word95.Contains("Paragraph 3. Has some <span style=\"color:light-dark(#FF0000"), "Word95.doc: RED run is red (sprm 98)");
        TestReport.Check(word95.Contains("font-weight:bold;color:light-dark(#0000FF") && word95.Contains(">BLUE BOLD</span>"), "Word95.doc: BLUE BOLD run is bold and blue (sprms 85, 98)");
        TestReport.Check(word95.Contains("vertical-align:super;font-size:smaller;\">th</span>"), "Word95.doc: 4th is superscript (sprm 104)");
        TestReport.Check(word95.Contains("font-size:12pt;font-family:'Times New Roman'"), "Word95.doc: Normal style size and font from the Word 6 style sheet and font table");

        // wIdent 0xA699 with fExtChar: every piece is UTF-16 (ww8scan.cxx WW8Cp2Fc for nVersion < 8)
        string extended = WordFieldsFidelityTest.Strip(Html(@"apache-poi\Bug51944.doc"));
        TestReport.Check(extended.Contains("Many thanks for"), "Bug51944.doc (0xA699, fExtChar): UTF-16 pieces read whole");

        // 8-bit text in the code page of the fonts' character set: Cyrillic (cp1251) and Czech (cp1250)
        string cyrillic = WordFieldsFidelityTest.Strip(Html(@"apache-poi\Bug60942b.doc"));
        TestReport.Check(cyrillic.Count(c => c >= (char)0x0410 && c <= (char)0x044F) >= 10, "Bug60942b.doc: Cyrillic text decoded as cp1251");
        string czech = Html(@"libreoffice\sw\qa\core\data\ww6\pass\ofz21385-1.doc");
        TestReport.Check(czech.Contains((char)0x017E) && czech.Contains((char)0x0161), "ofz21385-1.doc: Czech text decoded as cp1250");
        TestReport.Check(czech.Contains("<table"), "ofz21385-1.doc: Word 6 tables (sprmTDefTable with 10-byte TCs) render as tables");
        string tables = Html(@"libreoffice\sw\qa\extras\layout\data\forcepoint92.doc");
        TestReport.Check(tables.Split(new[] { "<table" }, StringSplitOptions.None).Length > 10, "forcepoint92.doc: Word 6 tables render");

        // Word 6 numbering (sprmPNLvlAnm 10 then 2, ANLD / OLST): 1. 2. 2.1 2.2 2.3 3. 4. 4.1 4.2 - LibreOffice's test for this
        // file (ww8export4.cxx testTdf150197_anlv2ListFormat) expects 2.1 for the third paragraph and notes it still shows 4.1
        string numbered = WordFieldsFidelityTest.Strip(Html(@"libreoffice\sw\qa\extras\ww8export\data\tdf150197_anlv2ListFormat.doc"));
        string[] expectedLabels = ["1.", "2.", "2.1", "2.2", "2.3", "3.", "4.", "4.1", "4.2"];
        int found = 0, from = 0;
        foreach (string label in expectedLabels)
        {
            int at = numbered.IndexOf(label, from, StringComparison.Ordinal);
            if (at < 0) break;
            found++; from = at + label.Length;
        }
        TestReport.Check(found == expectedLabels.Length && !numbered.Contains("3.1"), $"tdf150197_anlv2ListFormat.doc: Word 6 outline numbers nest under the numbering ({found} of {expectedLabels.Length} in order)");
        // the same labels from a saved sidecar index (a reopen restores the Word 6 counter state from the anchors)
        var sidecar = new WordIndexPolicy { ForcedMode = WordIndexMode.Sidecar, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word6-sidecar") };
        string numberedPath = Path.Combine(testData, @"libreoffice\sw\qa\extras\ww8export\data\tdf150197_anlv2ListFormat.doc");
        string[] opens = new string[2];
        for (int open = 0; open < 2; open++)
        {
            try
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(numberedPath, sidecar);
                if (open == 0) document.ScanAhead(int.MaxValue);
                opens[open] = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
            }
            catch (Exception exception) { opens[open] = "error: " + exception.GetType().Name + " " + exception.Message; }
        }
        TestReport.Check(opens[0] == opens[1] && opens[1].Contains("2.1"), $"tdf150197: sidecar reopen renders the same labels ({(opens[1].StartsWith("error") ? opens[1] : "html " + opens[1].Length)})");

        // A font with character set 177 (cp1255): Word6.doc with its font's chs set to 177 and five letters 0xE0-0xE4 in place
        // of "quick" must show U+05D0..U+05D4 (the fixture is generated here from the test-data file; version suffix v1)
        string source = Path.Combine(testData, @"apache-poi\Word6.doc");
        if (File.Exists(source))
        {
            byte[] bytes = File.ReadAllBytes(source);
            int text = IndexOf(bytes, "quick"u8.ToArray());
            for (int other = IndexOf(bytes, "quick"u8.ToArray(), text + 1); other > 0; other = IndexOf(bytes, "quick"u8.ToArray(), other + 1))
                for (int i = 0; i < 5; i++) bytes[other + i] = (byte)(0xE0 + i);    // the file holds the word twice; patch every copy
            int font = IndexOf(bytes, "Times New Roman\0"u8.ToArray());
            bool patched = text > 0 && font > 6;
            if (patched)
            {
                bytes[font - 2] = 177;                                             // FFN chs, two bytes before the name (ww8struc.hxx WW8_FFN_BASE)
                for (int i = 0; i < 5; i++) bytes[text + i] = (byte)(0xE0 + i);
            }
            string copy = Path.Combine(TestReport.OutputDirectory, "word6-codepage-1255-v1.doc");
            File.WriteAllBytes(copy, bytes);
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(copy, policy);
            string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
            string expected = "The " + new string(Enumerable.Range(0x05D0, 5).Select(code => (char)code).ToArray()) + " brown";
            TestReport.Check(patched && html.Contains(expected), "Word 6 font with character set 177: 8-bit text decoded as cp1255");
        }
        else TestReport.Check(false, "test-data/apache-poi/Word6.doc present");
    }

    /// <summary>
    /// Word for Windows 1.x / 2.0 files (no compound file; LibreOffice's eWW1 / eWW2 branches): text compared with LibreOffice's
    /// text export (2026-11-18: word2.doc, ooo2535-1-WW2.doc, ooo40683-1-WW1.doc word-for-word equal; bold and italic runs as in
    /// its .fodt export), plus the damaged ww5 files LibreOffice keeps as import tests: each opens or fails with a clear error.
    /// </summary>
    public static void RunWord2()
    {
        TestReport.Step("Word for Windows 1.x / 2.0 .doc: FIB with 2-byte lcbs, CHPX runs, old style sheet, Word 2 sprms");
        string testData = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word2") };
        string Html(string relative)
        {
            string path = Path.Combine(testData, relative);
            if (!File.Exists(path)) return "";
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            return string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        }

        // Word 2.0 (wIdent 0xA5DB): bold from the CHPX toggle bits, the Normal style's size from the old-format style sheet
        string word2 = Html(@"apache-poi\word2.doc");
        TestReport.Check(word2.Contains("font-weight:bold;\">Member of 3GPP (ARIB)</span>"), "word2.doc: bold run from a Word 2 CHPX");
        TestReport.Check(word2.Contains("Mr. Benni Alexander\tNokia Japan Co, Ltd") && word2.Contains("font-size:10pt;font-family:'Times New Roman'"),
                         "word2.doc: text and the Normal style (10 pt) from the Word 2 style sheet");
        // a complex Word 2 file (fComplex, 2-byte Pcdt lcb) with sizes and italics from CHPX fsHps / toggles
        string complex = Html(@"libreoffice\sw\qa\core\data\ww5\pass\ooo2535-1-WW2.doc");
        string complexText = WordFieldsFidelityTest.Strip(complex);
        TestReport.Check(complexText.Contains("Orden de Pago al Proveedor") && complexText.Contains("N" + (char)0x00FA + "mero:"), "ooo2535-1-WW2.doc: text and cp1252 letters");
        TestReport.Check(complex.Contains("font-size:12pt") && complex.Contains("font-style:italic"), "ooo2535-1-WW2.doc: run sizes and italics");
        // Word 1.x (wIdent 0xA59B): the shorter FIB (no fcPrDrvr / fcPrEnvLand), a complex file, justified paragraphs
        string word1 = Html(@"libreoffice\sw\qa\core\data\ww5\pass\ooo40683-1-WW1.doc");
        string word1Text = WordFieldsFidelityTest.Strip(word1);
        TestReport.Check(word1Text.Contains("Sonderaktion Winter 1990/1991") && word1Text.Contains("g" + (char)0x00FC + "nstiger"), "ooo40683-1-WW1.doc: Word 1 text");
        TestReport.Check(word1.Contains("text-align:justify"), "ooo40683-1-WW1.doc: paragraph alignment from a Word 1 PAPX");

        foreach (string damaged in (string[])["crash-1.doc", "crash-3.doc", "hang-1.doc", "hang-2.doc", "hang-3.doc"])
        {
            string? error = null;
            try { _ = Html(@"libreoffice\sw\qa\core\data\ww5\pass\" + damaged); }
            catch (InvalidDataException exception) when (exception.Message.StartsWith("Damaged", StringComparison.Ordinal)) { }
            catch (Exception exception) { error = exception.GetType().Name + ": " + exception.Message; }
            TestReport.Check(error is null, $"ww5/pass/{damaged}: opens or fails as damaged{(error is null ? "" : " - " + error)}");
        }
    }

    /// <summary>
    /// Word 6.0 / 95 files saved by the Far East builds (fExtChar: 16-bit units, a unit at or above 0x3000 = a double-byte
    /// character, high byte lead; LibreOffice ww8par.cxx ReadPlainChars): Japanese, Korean, simplified and traditional Chinese
    /// from LibreOffice's sw/qa ww6 data. Each file's text equals LibreOffice's text export (2026-11-18) with white space removed;
    /// the expected text is kept as its first characters and its length.
    /// </summary>
    public static void RunAsian()
    {
        TestReport.Step("Word 6.0 / 95 .doc: Asian double-byte text (Far East builds)");
        string testData = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word6-asian") };
        foreach ((string file, string start, int length) in AsianExpectations)
        {
            string path = Path.Combine(testData, @"libreoffice\sw\qa\core\data\ww6\pass\" + file);
            if (!File.Exists(path)) { TestReport.Check(false, $"test-data {file} present"); continue; }
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
            string text = WordFieldsFidelityTest.Strip(string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html)));
            text = new string(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
            File.WriteAllText(Path.Combine(TestReport.OutputDirectory, "word6-asian-" + file + ".txt"), text);
            TestReport.Check(text.StartsWith(start, StringComparison.Ordinal) && text.Length == length && !text.Contains('�'),
                             $"{file}: double-byte text as LibreOffice reads it ({text.Length} characters, expected {length})");
        }
        // 8-bit text in a font with character set 128 (Shift-JIS): Word6.doc with "quic" replaced by two cp932 pairs (0x82A0,
        // 0x82A2) - each character on its lead byte's CP, the trail's CP shows nothing (fixture generated here, version suffix v1)
        string source = Path.Combine(testData, @"apache-poi\Word6.doc");
        if (File.Exists(source))
        {
            byte[] bytes = File.ReadAllBytes(source);
            byte[] quick = "quick"u8.ToArray();
            int text = IndexOf(bytes, quick);
            for (int at = text; at > 0; at = IndexOf(bytes, quick, at + 1))
            {
                bytes[at] = 0x82; bytes[at + 1] = 0xA0; bytes[at + 2] = 0x82; bytes[at + 3] = 0xA2;
            }
            int font = IndexOf(bytes, "Times New Roman\0"u8.ToArray());
            if (font > 6) bytes[font - 2] = 128;                                   // FFN chs (ww8struc.hxx WW8_FFN_BASE)
            string copy = Path.Combine(TestReport.OutputDirectory, "word6-codepage-932-v1.doc");
            File.WriteAllBytes(copy, bytes);
            using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(copy, policy);
            string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
            TestReport.Check(text > 0 && font > 6 && html.Contains("The あいk brown"), "Word 6 font with character set 128: 8-bit double-byte text decoded as cp932");
            WordDocumentSearchBatch found = document.Search("あいk brown", WordDocumentSearchOptions.Default, 0, document.BlockCount);
            TestReport.Check(found.HitBlocks.Length > 0, "Word 6 cp932 text: search finds a phrase across double-byte characters");
        }
    }

    private static readonly (string file, string start, int length)[] AsianExpectations =
    [
        ("ooo17498-1.doc", "\u7BA1\u7406\u7D44\u5408\u5F79\u54E1\u5404\u4F4D", 178),
        ("ooo17498-2.doc", "\uC0C8\uBB38\uC11C\uB97C\uB9CC\uB4E4\uB824\uBA74", 113),
        ("ooo17498-3.doc", "\u521B\u5EFA\u65B0\u6587\u6863\u5728\u201C\u6587", 87),
        ("ooo17498-4.doc", "\u5EFA\u7ACB\u65B0\u6587\u4EF6\u5728[\u6A94", 98),
        ("ooo8726-1.doc", "\u4E4E\u5C11\u6A01\u7A7F\u5C2C\u6A01\u7A7F\u5C2C", 12),
    ];

    private static int IndexOf(byte[] haystack, byte[] needle, int from = 0)
    {
        for (int i = Math.Max(0, from); i + needle.Length <= haystack.Length; i++)
            if (haystack.AsSpan(i, needle.Length).SequenceEqual(needle)) return i;
        return -1;
    }
}
