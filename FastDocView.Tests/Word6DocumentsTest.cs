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

    private static int IndexOf(byte[] haystack, byte[] needle, int from = 0)
    {
        for (int i = Math.Max(0, from); i + needle.Length <= haystack.Length; i++)
            if (haystack.AsSpan(i, needle.Length).SequenceEqual(needle)) return i;
        return -1;
    }
}
