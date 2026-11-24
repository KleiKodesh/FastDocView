using FastDocView;

namespace FastDocView.Tests;

/// <summary>
/// Word for Windows 2.0 right-to-left runs, line spacing and pictures, and Word 6.0 / 95 ANLD label formatting. No public
/// Hebrew Word 2 file, no Word 2 file with sprm 20 or a picture and no Word 6 file with formatted ANLD labels exist in the
/// LibreOffice / POI corpora, and Word 16 opens neither format here (File Block) nor saves them, so the fixtures are built from
/// test-data files by patching their bytes as the Word 6 binary format document and LibreOffice's readers lay them out
/// (version suffix v1 in the copies' names; bump it when a patch changes).
/// </summary>
public static class OldWordFormattingTest
{
    private static string TestData => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "test-data"));

    public static void Run()
    {
        RunWord2();
        RunWord6Labels();
    }

    // word2.doc (Word 2.0, 512-byte FKP pages): CHP FKP page 13 holds fc 384.., PAP FKP page 21 fc 384..1239
    private const int CharacterPage = 13, ParagraphPage = 21, PageSize = 512;

    private static void RunWord2()
    {
        TestReport.Step("Word 2.0 .doc: right-to-left runs (CHPX fBiDi / fBoldBi / ftcBi / hpsBi / lidBi / icoBi), sprm 50 / 20, pictures");
        string source = Path.Combine(TestData, @"apache-poi\word2.doc");
        if (!File.Exists(source)) { TestReport.Check(false, "test-data/apache-poi/word2.doc present"); return; }
        byte[] bytes = File.ReadAllBytes(source);

        // the font table (fc 14570, 40 bytes up to the next table: Times New Roman, Symbol, Arial) is saved with lcbSttbfFfn 0,
        // which LibreOffice reads as no fonts ("font table is broken", ww8scan.cxx:7222-7224): given its length, and Arial
        // (font 2) as a Hebrew font: FFN chs (Word 2 FFN: cbFfnM1, ff, chs, name) = 177 -> cp1255
        BitConverter.GetBytes((ushort)40).CopyTo(bytes, 0x58 + 15 * 6 + 4);
        int arial = IndexOf(bytes, "Arial\0"u8.ToArray());
        bytes[arial - 1] = 177;

        // the run "Nokia Japan Co, Ltd" (fc 428-447) as 19 Hebrew letters (cp1255 0xE0..0xF2) in a right-to-left run
        for (int i = 0; i < 19; i++) bytes[428 + i] = (byte)(0xE0 + i);
        byte[] rightToLeft = new byte[20];
        rightToLeft[1] = 0x70;                                   // fBoldBi, fItalicBi, fBiDi
        rightToLeft[2] = 0x02 | 0x04 | 0x80;                     // fsFtc, fsHps, fsIcoBi
        rightToLeft[3] = 0x07;                                   // fsFtcBi, fsHpsBi, fsLidBi
        rightToLeft[6] = 0x14;                                   // hps 10 pt, ftc 0
        rightToLeft[11] = 6;                                     // icoBi red
        rightToLeft[12] = 0x09; rightToLeft[13] = 0x04;          // lid 0x0409
        rightToLeft[14] = 2;                                     // ftcBi Arial
        rightToLeft[16] = 0x1C;                                  // hpsBi 14 pt
        rightToLeft[18] = 0x0D; rightToLeft[19] = 0x04;          // lidBi 0x040D Hebrew

        // pictures appended to the file (Word 2: they follow the text in the one stream): a 16 x 16 one-bit bitmap (mm 99)
        // and a Windows metafile (mm 8, a header and the EOF record), each a 0x3A-byte PIC first
        int bitmapAt = bytes.Length, metafileAt = bitmapAt + 0x3A + 32;
        var tail = new List<byte>();
        tail.AddRange(Pic(lcb: 0x3A + 32, mm: 99, dxaGoal: 240, dyaGoal: 240, bitmap: true));
        for (int row = 0; row < 16; row++) tail.AddRange(row % 2 == 0 ? [0xF0, 0x0F] : [0x0F, 0xF0]);
        byte[] metafile = [1, 0, 9, 0, 0, 3, 12, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0];
        tail.AddRange(Pic(lcb: 0x3A + metafile.Length, mm: 8, dxaGoal: 1440, dyaGoal: 720, bitmap: false));
        tail.AddRange(metafile);
        bytes = [.. bytes, .. tail];
        bytes[466] = 0x01; bytes[467] = 0x01;                     // "FI" (fc 466-468) -> two bitmap pictures
        bytes[468] = 0x01;                                       // the tab (fc 468) -> the metafile
        byte[] PictureChpx(int at) { var chpx = new byte[24]; chpx[1] = 0x02; chpx[2] = 0x06; chpx[6] = 0x14; BitConverter.GetBytes(at).CopyTo(chpx, 20); return chpx; }

        if (!RewriteCharacterPage(bytes, CharacterPage * PageSize, new Dictionary<int, byte[]> { [428] = rightToLeft, [466] = PictureChpx(bitmapAt), [468] = PictureChpx(metafileAt) }))
        { TestReport.Check(false, "word2 fixture: CHP page rewritten"); return; }

        // paragraph 1 (fc 384-408): right-to-left (sprm 50) and exactly 24 pt (sprm 20 = -480); paragraph fc 640-664: at least 18 pt
        if (!AddParagraphSprms(bytes, ParagraphPage * PageSize, new Dictionary<int, byte[]> { [384] = [50, 1, 20, 0x20, 0xFE], [640] = [20, 0x68, 0x01] }))
        { TestReport.Check(false, "word2 fixture: PAP page rewritten"); return; }

        string copy = Path.Combine(TestReport.OutputDirectory, "word2-rtl-spacing-pictures-v1.doc");
        File.WriteAllBytes(copy, bytes);
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word2-rtl") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(copy, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        string hebrew = new string(Enumerable.Range(0x05D0, 19).Select(code => (char)code).ToArray());
        int run = html.IndexOf(hebrew, StringComparison.Ordinal);
        string runOpen = run < 0 ? "" : html.Substring(html.LastIndexOf('<', run), run - html.LastIndexOf('<', run));
        TestReport.Check(run > 0, "Word 2 right-to-left run: 8-bit text decoded in its ftcBi font's code page (cp1255)");
        TestReport.Check(runOpen.Contains("dir=\"rtl\"") || runOpen.Contains("direction:rtl"), $"Word 2 fBiDi: the run is right to left ({Ascii(runOpen)})");
        TestReport.Check(runOpen.Contains("font-weight:bold") && runOpen.Contains("font-style:italic"), "Word 2 fBoldBi / fItalicBi: the right-to-left run is bold and italic");
        TestReport.Check(runOpen.Contains("font-size:14pt") && runOpen.Contains("Arial") && runOpen.Contains("#FF0000"), $"Word 2 hpsBi / ftcBi / icoBi: 14 pt, Arial, red ({Ascii(html.Substring(Math.Max(0, run - 700), 760))})");
        string text = WordFieldsFidelityTest.Strip(html);
        WordDocumentSearchBatch found = document.Search(hebrew.Substring(2, 5), WordDocumentSearchOptions.Default, 0, document.BlockCount);
        TestReport.Check(found.HitBlocks.Length == 1 && text.Contains(hebrew), "Word 2 right-to-left text: search finds what is shown");

        string first = document.RenderBlocks(0, 1)[0].Html;
        TestReport.Check(first.Contains("dir=\"rtl\""), "Word 2 sprm 50 (sprmPFBiDi): the paragraph is right to left");
        TestReport.Check(first.Contains("line-height:24pt"), $"Word 2 sprm 20 (dyaLine -480): exactly 24 pt ({Ascii(first.Substring(0, Math.Min(200, first.Length)))})");
        int ccsa = html.IndexOf("Member of 3GPP (CCSA)", StringComparison.Ordinal);
        string ccsaParagraph = ccsa < 0 ? "" : html.Substring(html.LastIndexOf("<p", ccsa, StringComparison.Ordinal), ccsa - html.LastIndexOf("<p", ccsa, StringComparison.Ordinal));
        TestReport.Check(ccsaParagraph.Contains("18pt"), $"Word 2 sprm 20 (dyaLine 360): at least 18 pt ({Ascii(ccsaParagraph)})");

        string bitmapId = bitmapAt.ToString(System.Globalization.CultureInfo.InvariantCulture);
        TestReport.Check(html.Split(["data-word-picture=\"" + bitmapId + "\""], StringSplitOptions.None).Length == 3 && html.Contains("width=\"16\""),
                         "Word 2 picture (mm 99 bitmap): two 12 pt pictures in the line");
        WordDocumentPicture? picture = document.ReadPicture(bitmapId);
        TestReport.Check(picture is { ContentType: "image/bmp" } && picture.Bytes[0] == 0x42 && picture.Bytes[1] == 0x4D && picture.Bytes.Length == 14 + 40 + 8 + 4 * 16,
                         "Word 2 bitmap: served as a .bmp (1 bit, rows padded to 4 bytes)");
        TestReport.Check(html.Contains("word-picture-missing\" style=\"width:72pt;aspect-ratio:72/36;"), "Word 2 metafile: its 72 x 36 pt box");
    }

    private static byte[] Pic(int lcb, int mm, int dxaGoal, int dyaGoal, bool bitmap)
    {
        var pic = new byte[0x3A];
        BitConverter.GetBytes(lcb).CopyTo(pic, 0);
        BitConverter.GetBytes((ushort)0x3A).CopyTo(pic, 4);
        BitConverter.GetBytes((ushort)mm).CopyTo(pic, 6);
        BitConverter.GetBytes((ushort)(bitmap ? 16 : 2540)).CopyTo(pic, 8);
        BitConverter.GetBytes((ushort)(bitmap ? 16 : 1270)).CopyTo(pic, 10);
        if (bitmap)
        {
            BitConverter.GetBytes((ushort)16).CopyTo(pic, 0xE + 2);    // bmWidth
            BitConverter.GetBytes((ushort)16).CopyTo(pic, 0xE + 4);    // bmHeight
            BitConverter.GetBytes((ushort)2).CopyTo(pic, 0xE + 6);     // bmWidthBytes
            pic[0xE + 8] = 1; pic[0xE + 9] = 1;                         // bmPlanes, bmBitsPixel
        }
        BitConverter.GetBytes((ushort)dxaGoal).CopyTo(pic, 0x1C);
        BitConverter.GetBytes((ushort)dyaGoal).CopyTo(pic, 0x1E);
        BitConverter.GetBytes((ushort)1000).CopyTo(pic, 0x20);
        BitConverter.GetBytes((ushort)1000).CopyTo(pic, 0x22);
        return pic;
    }

    /// <summary>
    /// Rewrites a Word 2 CHP FKP (rgfc[crun + 1], one-byte offsets in words, CHPXs = a length byte + the bytes; crun in the last
    /// byte) with each run starting at a key of <paramref name="newProperties"/> given that CHPX. Room comes from merging each
    /// one-character run that keeps the style's formatting (a tab or paragraph mark) into the run after it when that run sets
    /// only ftc 0 / hps 10 pt - the same look - from the end of the page backwards.
    /// </summary>
    private static bool RewriteCharacterPage(byte[] file, int page, Dictionary<int, byte[]> newProperties)
    {
        int count = file[page + PageSize - 1];
        var starts = new List<int>();
        var properties = new List<byte[]?>();
        for (int i = 0; i <= count; i++) starts.Add(BitConverter.ToInt32(file, page + i * 4));
        for (int i = 0; i < count; i++)
        {
            int offset = file[page + (count + 1) * 4 + i] * 2;
            properties.Add(offset == 0 ? null : file.AsSpan(page + offset + 1, file[page + offset]).ToArray());
        }
        byte[] plain = Convert.FromHexString("0000060000001400");
        for (int i = properties.Count - 2; i >= 0 && UsedBytes(properties, newProperties, starts) > PageSize - 1; i--)
        {
            if (properties[i] is null && starts[i + 1] - starts[i] == 1 && properties[i + 1] is { } next && next.AsSpan().SequenceEqual(plain)
                && !newProperties.ContainsKey(starts[i]) && !newProperties.ContainsKey(starts[i + 1]))
            {
                properties[i] = next;
                properties.RemoveAt(i + 1);
                starts.RemoveAt(i + 1);
            }
        }
        foreach ((int start, byte[] chpx) in newProperties)
        {
            int index = starts.IndexOf(start);
            if (index < 0 || index >= properties.Count) return false;
            properties[index] = chpx;
        }
        if (UsedBytes(properties, newProperties, starts) > PageSize - 1) return false;
        var block = new byte[PageSize];
        int run = properties.Count;
        for (int i = 0; i <= run; i++) BitConverter.GetBytes(starts[i]).CopyTo(block, i * 4);
        int end = PageSize - 1;
        var placed = new Dictionary<string, int>();
        for (int i = 0; i < run; i++)
        {
            if (properties[i] is not { } chpx) continue;
            string key = Convert.ToHexString(chpx);
            if (!placed.TryGetValue(key, out int at))
            {
                at = (end - chpx.Length - 1) & ~1;
                block[at] = (byte)chpx.Length;
                chpx.CopyTo(block, at + 1);
                end = at;
                placed[key] = at;
            }
            block[(run + 1) * 4 + i] = (byte)(at / 2);
        }
        block[PageSize - 1] = (byte)run;
        Buffer.BlockCopy(block, 0, file, page, PageSize);
        return true;
    }

    private static int UsedBytes(List<byte[]?> properties, Dictionary<int, byte[]> newProperties, List<int> starts)
    {
        var distinct = new HashSet<string>();
        for (int i = 0; i < properties.Count; i++)
        {
            byte[]? chpx = newProperties.TryGetValue(starts[i], out byte[]? replaced) ? replaced : properties[i];
            if (chpx is not null) distinct.Add(Convert.ToHexString(chpx));
        }
        return (properties.Count + 1) * 4 + properties.Count + distinct.Sum(key => (key.Length / 2 + 2) & ~1) + 2;
    }

    /// <summary>Gives the paragraphs starting at the keys a copy of their PAPX (cw, stc, PHE, grpprl) with sprms appended, placed
    /// in the page's free space (a Word 2 PAP FKP: rgfc, one-byte BX offsets in words, crun in the last byte).</summary>
    private static bool AddParagraphSprms(byte[] file, int page, Dictionary<int, byte[]> sprms)
    {
        int count = file[page + PageSize - 1];
        int freeStart = page + (count + 1) * 4 + count;
        int lowest = PageSize - 1;
        for (int i = 0; i < count; i++) if (file[page + (count + 1) * 4 + i] is var bx and > 0) lowest = Math.Min(lowest, bx * 2);
        int end = page + lowest;
        foreach ((int start, byte[] added) in sprms)
        {
            int index = -1;
            for (int i = 0; i < count; i++) if (BitConverter.ToInt32(file, page + i * 4) == start) index = i;
            if (index < 0) return false;
            int offset = page + file[page + (count + 1) * 4 + index] * 2;
            int words = file[offset];
            byte[] old = file.AsSpan(offset + 1, words * 2).ToArray();
            byte[] papx = [.. old, .. added];                                  // a trailing pad byte reads as sprm 0 (no operand)
            if (papx.Length % 2 != 0) papx = [.. papx, 0];
            int at = (end - papx.Length - 1) & ~1;
            if (at < freeStart) return false;
            file[at] = (byte)(papx.Length / 2);
            papx.CopyTo(file, at + 1);
            file[page + (count + 1) * 4 + index] = (byte)((at - page) / 2);
            end = at;
        }
        return true;
    }

    /// <summary>tdf150197_anlv2ListFormat.doc with its ANLDs (sprmPAnld, 52 bytes) given fSetBold + fBold, fSetItalic (not
    /// italic), red, 14 pt and fPrevSpace; dxaIndent 360 / dxaSpace 120 are the file's own.</summary>
    private static void RunWord6Labels()
    {
        TestReport.Step("Word 6 .doc: ANLD label formatting (fSetBold / fBold, ico, hps, fPrevSpace) and spacing (dxaIndent, dxaSpace, fHang)");
        string source = Path.Combine(TestData, @"libreoffice\sw\qa\extras\ww8export\data\tdf150197_anlv2ListFormat.doc");
        if (!File.Exists(source)) { TestReport.Check(false, "test-data tdf150197_anlv2ListFormat.doc present"); return; }
        byte[] bytes = File.ReadAllBytes(source);
        int patched = 0;
        for (int at = IndexOf(bytes, [12, 52]); at >= 0; at = IndexOf(bytes, [12, 52], at + 1))
        {
            int anld = at + 2;
            if (anld + 52 > bytes.Length || bytes[anld] > 7 || bytes[anld + 1] + bytes[anld + 2] > 32) continue;
            bytes[anld + 3] |= 0x10 | 0x20;                                 // fSetBold, fSetItalic
            bytes[anld + 4] = (byte)((bytes[anld + 4] & ~0x10) | 0x08 | 0x04);   // fBold, not fItalic, fPrevSpace
            bytes[anld + 5] = (byte)((bytes[anld + 5] & 0x07) | (6 << 3));  // ico red
            bytes[anld + 8] = 28; bytes[anld + 9] = 0;                      // hps 14 pt
            patched++;
        }
        string copy = Path.Combine(TestReport.OutputDirectory, "word6-anld-label-format-v1.doc");
        File.WriteAllBytes(copy, bytes);
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-word6-anld") };
        using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(copy, policy);
        string html = string.Concat(document.RenderBlocks(0, document.BlockCount).Select(block => block.Html));
        int marker = html.IndexOf("class=\"list-marker\"", StringComparison.Ordinal);
        string label = marker < 0 ? "" : html.Substring(marker, Math.Min(260, html.Length - marker));
        TestReport.Check(patched >= 1 && label.Contains("font-weight:bold") && !label.Contains("font-style:italic") && label.Contains("font-size:14pt") && label.Contains("#FF0000"),
                         $"Word 6 ANLD: the label is bold, upright, 14 pt, red ({patched} ANLDs; {Ascii(label)})");
        TestReport.Check(label.Contains("min-width:18pt") && label.Contains("padding-inline-end:6pt"), "Word 6 ANLD: label box dxaIndent 360, dxaSpace 120");
        TestReport.Check(label.Contains("> 1.<"), "Word 6 ANLD fPrevSpace: one space before the number");
        string plain = WordFieldsFidelityTest.Strip(html);
        TestReport.Check(plain.Contains(" 1."), "Word 6 ANLD: the label text is part of the block text");
    }

    private static string Ascii(string text) => new(text.Select(c => c is >= ' ' and <= '~' ? c : '?').ToArray());

    private static int IndexOf(byte[] haystack, byte[] needle, int from = 0)
    {
        for (int i = from; i <= haystack.Length - needle.Length; i++)
            if (haystack.AsSpan(i, needle.Length).SequenceEqual(needle)) return i;
        return -1;
    }
}
