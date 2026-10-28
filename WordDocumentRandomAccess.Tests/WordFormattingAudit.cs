using System.IO.Compression;
using System.Xml;
using WordDocumentRandomAccess.DocBinary;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Lists every formatting construct a real document uses that the renderers do NOT handle, with
/// counts, so coverage work follows real files instead of the spec's table of contents.
///   .docx — streams every part that holds text (document, footnotes, endnotes, styles, numbering)
///           with XmlReader and counts element names per parent context that are not in the
///           handled sets below.
///   .doc  — walks every PAPX and CHPX run (and every style's UPX) and counts sprm opcodes that
///           DocBinaryPropertyModifiers does not interpret.
/// The handled sets mirror DocxStyleCascade / DocxBlockHtmlRenderer / DocBinaryPropertyModifiers —
/// update both sides together.
/// </summary>
public static class WordFormattingAudit
{
    // Pagination, proofing and language hints: deliberately ignored by a flow (no pages) renderer.
    private static readonly HashSet<string> IgnoredByDesign =
        ["lang", "noProof", "kern", "widowControl", "keepNext", "keepLines", "pageBreakBefore", "suppressLineNumbers",
         "suppressAutoHyphens", "textAlignment", "snapToGrid", "adjustRightInd", "autoSpaceDE", "autoSpaceDN", "tblLook",
         "rsid", "rsidR", "rsidRPr", "w", "eastAsianLayout", "specVanish", "oMath", "framePr",
         // Word 2010+ typography (w14): OpenType ligatures / number forms / spacing / contextual alternates, text outline
         "ligatures", "numForm", "numSpacing", "cntxtAlts", "textOutline", "effect", "stylisticSets",
         // table style banding and floating-table placement, alt text, conditional-format marks (table styles' base is applied)
         "tblStyleRowBandSize", "tblStyleColBandSize", "tblpPr", "tblOverlap", "tblDescription", "tblCaption", "cnfStyle", "hideMark", "tblPrEx",
         // East-Asian line breaking, mirrored indents, collapsed headings (w15), revision history of properties
         "overflowPunct", "wordWrap", "mirrorIndents", "collapsed", "kinsoku", "topLinePunct", "rPrChange", "pPrChange", "tblPrChange", "tcPrChange", "trPrChange",
         // revision marks on the paragraph mark itself; move ranges (the moved text itself is rendered)
         "ins", "del", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd",
         // no pages, no comments in a reading view
         "pgNum", "commentReference", "annotationRef", "underlineColor", "fitText", "moveFrom", "moveTo"];
    private static readonly HashSet<string> HandledRunProperties =
        ["rStyle", "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vanish", "webHidden", "rtl", "vertAlign", "sz", "szCs", "color",
         "highlight", "shd", "rFonts", "smallCaps", "caps", "spacing", "cs", "position", "outline", "shadow", "emboss", "imprint", "bdr", "textFill"];
    private static readonly HashSet<string> HandledParagraphProperties =
        ["pStyle", "bidi", "jc", "ind", "spacing", "outlineLvl", "numPr", "rPr", "ilvl", "numId", "contextualSpacing", "shd", "tabs", "sectPr", "pBdr"];
    private static readonly HashSet<string> HandledRunChildren =
        ["rPr", "t", "delText", "tab", "ptab", "br", "cr", "noBreakHyphen", "softHyphen", "sym", "footnoteReference", "endnoteReference",
         "footnoteRef", "endnoteRef", "separator", "continuationSeparator", "drawing", "pict", "object", "fldChar", "instrText", "lastRenderedPageBreak",
         "AlternateContent"];
    private static readonly HashSet<string> HandledParagraphChildren =
        ["pPr", "r", "hyperlink", "ins", "smartTag", "sdt", "customXml", "fldSimple", "dir", "bdo", "del",
         "bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd",
         "moveTo", "moveFrom", "AlternateContent"];
    private static readonly HashSet<string> HandledBodyChildren = ["p", "tbl", "sdt", "sectPr", "bookmarkStart", "bookmarkEnd"];
    private static readonly HashSet<string> HandledTableParts =
        ["tblPr", "tblGrid", "tr", "tc", "tcPr", "gridSpan", "vMerge", "shd", "bidiVisual", "gridCol", "trPr", "p", "tbl",
         "tcW", "tcBorders", "vAlign", "tblW", "tblBorders", "tblInd", "tblCellMar", "tblStyle", "jc", "sdt",
         "tcMar", "noWrap", "tblCellSpacing", "tblLayout", "bookmarkStart", "bookmarkEnd"];

    private static readonly HashSet<ushort> HandledParagraphSprms =
        [0x2403, 0x2461, 0x2441, 0x840F, 0x845E, 0x840E, 0x845D, 0x8411, 0x8460, 0xA413, 0xA414, 0x6412, 0x2416, 0x260A, 0x460B, 0x2640, 0x4600,
         0x245B, 0xC64D, 0x442D, 0xC60D, 0xC615, 0x2417, 0x244C, 0xD608,
         // ignored by design: revision ids, pagination / widow / keep, auto-spacing, kinsoku, frames, table layout hints a flow row does not need
         0x6467, 0x2431, 0x2437, 0x2438, 0x2448, 0x2406, 0x2405, 0x2407, 0x240C, 0x4439, 0x2433, 0x2434, 0x2435, 0x2436, 0x245C, 0x245D,
         0x246D, 0x8419, 0x261B, 0x2423, 0x6649, 0x664A, 0x244B, 0x6646, 0x560B, 0xD612, 0xD613, 0xF617, 0x3466, 0x548A, 0x740A, 0x9602, 0x3403, 0x3404, 0x5400, 0x9407, 0xD634, 0xD61A, 0xD61B, 0xD61C, 0xD61D,
         0xF661, 0x7479, 0xD670, 0xF618, 0x646B, 0xD642, 0xD605, 0xD609, 0xF614, 0x9601, 0x3615,
         0x6424, 0x6425, 0x6426, 0x6427, 0xC64E, 0xC64F, 0xC650, 0xC651,
         // ignored by design: frame position (dxaAbs/dyaAbs/width), hyphenation / auto-adjust / kinsoku flags, list revision marks,
         // drop caps (shown as normal text), rsid / revision attribution
         0x8418, 0x8419, 0x842F, 0x940F, 0x9410, 0x941E, 0x2447, 0x242A, 0x2443, 0xC645, 0xC66F, 0x2664, 0x360D, 0x4458, 0x4459, 0x442B, 0x6428, 0x6629, 0xC652];
    private static readonly HashSet<ushort> HandledCharacterSprms =
        [0x0835, 0x085C, 0x0836, 0x085D, 0x0837, 0x083C, 0x085A, 0x2A3E, 0x2A48, 0x4A43, 0x4A61, 0x2A42, 0x6870, 0x2A0C, 0x0855, 0x4A30,
         0xCA71, 0x4866, 0x083A, 0x083B, 0x8840, 0x4A4F, 0x4A5E,
         // ignored by design: revision ids, language ids, proofing, script hint, east-asian/other-script fonts, field data, picture location
         0x6816, 0x6815, 0x6817, 0x286F, 0x486E, 0x4874, 0x485F, 0x486D, 0x4873, 0x0875, 0x4A50, 0x4A51, 0x0811, 0x0802, 0x0806, 0x6A03,
         0x6A09, 0x0800, 0x0801, 0x4804, 0x6805, 0x4807, 0x080A, 0x0882, 0x484B,
         0x4840, 0x0868, 0x2A53, 0x0838, 0x0839, 0x0858, 0x0854, 0x6865, 0xCA72,
         // ignored by design: kerning, character scale, underline color, deletion revision attribution, east-asian layout, script hints
         0x4845, 0x4852, 0x6877, 0x6864, 0x4863, 0x4A95, 0xCA89, 0x2A83, 0x2A86, 0x0856, 0x2859, 0x0016, 0x0146];

    /// <summary>Unhandled constructs: occurrences and the number of files that use each.</summary>
    private sealed class AuditTally
    {
        public readonly Dictionary<string, (long occurrences, int files)> Counts = new(StringComparer.Ordinal);
        private readonly HashSet<string> _seenInCurrentFile = new(StringComparer.Ordinal);

        public void BeginFile() => _seenInCurrentFile.Clear();

        public void Add(string key)
        {
            Counts.TryGetValue(key, out (long occurrences, int files) count);
            Counts[key] = (count.occurrences + 1, count.files + (_seenInCurrentFile.Add(key) ? 1 : 0));
        }
    }

    /// <summary>Per-file report, the samples' mode.</summary>
    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"Formatting audit: {Path.GetFileName(path)}");
            var tally = new AuditTally();
            tally.BeginFile();
            try
            {
                if (WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.Docx) AuditDocx(path, tally, printParts: true);
                else AuditDocBinary(path, tally, printRuns: true);
            }
            catch (Exception exception) { Console.WriteLine($"   audit failed: {exception.GetType().Name}: {exception.Message}"); }
            Print("unhandled constructs", tally, perFile: false);
        }
    }

    /// <summary>
    /// Corpus mode: every file in a list, aggregated — which unhandled constructs appear in how many
    /// files — plus an open-and-render attempt per file through the library to catch crashes.
    /// Below-normal priority, one file at a time, progress every 100 files.
    /// </summary>
    public static void RunCorpus(string listPath, int maxFiles)
    {
        try { System.Diagnostics.Process.GetCurrentProcess().PriorityClass = System.Diagnostics.ProcessPriorityClass.BelowNormal; } catch { }
        string[] paths = File.ReadAllLines(listPath).Where(line => line.Length > 0).Take(maxFiles).ToArray();
        TestReport.Step($"Corpus audit: {paths.Length:N0} files from {listPath}");
        var docxTally = new AuditTally();
        var docTally = new AuditTally();
        var failures = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        var formats = new Dictionary<string, int>(StringComparer.Ordinal);
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory, CacheDirectory = Path.Combine(TestReport.OutputDirectory, "corpus-index-cache") };
        var watch = System.Diagnostics.Stopwatch.StartNew();
        int done = 0;
        foreach (string path in paths)
        {
            done++;
            if (done % 100 == 0) Console.WriteLine($"   [{watch.Elapsed:mm\\:ss}] {done:N0} / {paths.Length:N0} files, {failures.Sum(f => f.Value.Count)} failures");
            if (!File.Exists(path)) { Count(formats, "missing"); continue; }
            WordWindowedDocumentOpener.WordFileFormat format;
            try { format = WordWindowedDocumentOpener.DetectFormat(path); }
            catch (Exception exception) { AddFailure(failures, "detect: " + exception.GetType().Name, path); continue; }
            Count(formats, format.ToString());
            if (format is not (WordWindowedDocumentOpener.WordFileFormat.Docx or WordWindowedDocumentOpener.WordFileFormat.DocBinary)) continue;

            AuditTally tally = format == WordWindowedDocumentOpener.WordFileFormat.Docx ? docxTally : docTally;
            tally.BeginFile();
            try
            {
                if (format == WordWindowedDocumentOpener.WordFileFormat.Docx) AuditDocx(path, tally, printParts: false);
                else AuditDocBinary(path, tally, printRuns: false);
            }
            catch (Exception exception) { AddFailure(failures, $"audit {format}: {exception.GetType().Name}: {Shorten(exception.Message)}", path); }

            try
            {
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                document.RenderBlocks(0, 40);
                if (document.BlockCount > 80) document.RenderBlocks(document.BlockCount / 2, 40);
            }
            catch (Exception exception) { AddFailure(failures, $"open {format}: {exception.GetType().Name}: {Shorten(exception.Message)}", path); }
        }

        Console.WriteLine($"   done in {watch.Elapsed:mm\\:ss}");
        Console.WriteLine("   formats: " + string.Join(", ", formats.OrderByDescending(f => f.Value).Select(f => $"{f.Key} {f.Value:N0}")));
        Print(".docx unhandled constructs (files using it / occurrences)", docxTally, perFile: true);
        Print(".doc unhandled sprms (files using it / occurrences)", docTally, perFile: true);
        Console.WriteLine("   failures:");
        if (failures.Count == 0) Console.WriteLine("      none");
        foreach ((string reason, List<string> files) in failures.OrderByDescending(f => f.Value.Count))
            Console.WriteLine($"      {files.Count,6:N0}  {reason}   e.g. {files[0]}");
    }

    private static void Count(Dictionary<string, int> counts, string key) => counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;

    private static void AddFailure(Dictionary<string, List<string>> failures, string reason, string path)
    {
        if (!failures.TryGetValue(reason, out List<string>? files)) failures[reason] = files = new List<string>();
        files.Add(path);
    }

    private static string Shorten(string message) => message.Length > 90 ? message.Substring(0, 90) + "…" : message;

    // ── .docx ──────────────────────────────────────────────────────────────────────────

    private static void AuditDocx(string path, AuditTally tally, bool printParts)
    {
        using var zip = ZipFile.OpenRead(path);
        foreach (string part in new[] { "word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/styles.xml", "word/numbering.xml" })
        {
            ZipArchiveEntry? entry = zip.GetEntry(part);
            if (entry is null) continue;
            if (printParts) Console.WriteLine($"   part {part}: {entry.Length >> 10} KB");
            using Stream stream = entry.Open();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, DtdProcessing = DtdProcessing.Ignore });
            var parents = new Stack<string>();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement) { if (parents.Count > 0) parents.Pop(); continue; }
                if (reader.NodeType != XmlNodeType.Element) continue;
                string name = reader.LocalName;
                string parent = parents.Count > 0 ? parents.Peek() : "";
                bool isHandled = IgnoredByDesign.Contains(name) || parent switch
                {
                    "rPr" => HandledRunProperties.Contains(name),
                    "pPr" => HandledParagraphProperties.Contains(name),
                    "numPr" => true,
                    "r" => HandledRunChildren.Contains(name),
                    "p" => HandledParagraphChildren.Contains(name),
                    "body" => HandledBodyChildren.Contains(name),
                    "tbl" or "tr" or "tc" or "tcPr" or "tblPr" or "tblGrid" => HandledTableParts.Contains(name),
                    _ => true,       // deeper contexts (drawing internals, style definitions…) are not rendered element by element
                };
                if (!isHandled) tally.Add($"{parent}/{(reader.Prefix.Length > 0 ? reader.Prefix + ":" : "")}{name}");
                if (!reader.IsEmptyElement) parents.Push(name);
            }
        }
    }

    // ── .doc ───────────────────────────────────────────────────────────────────────────

    /// <summary>Reads the FKPs and style sheet directly — no block index is built, so a corpus pass stays cheap.</summary>
    private static void AuditDocBinary(string path, AuditTally tally, bool printRuns)
    {
        using var compoundFile = new CompoundFileReader(path);
        using Stream wordDocument = compoundFile.OpenRootStream("WordDocument") ?? throw new InvalidDataException("No WordDocument stream.");
        var fib = new DocBinaryFileInformationBlock(wordDocument);
        using Stream tableStream = compoundFile.OpenRootStream(fib.TableStreamName) ?? throw new InvalidDataException("No table stream.");
        using Stream? dataStream = compoundFile.OpenRootStream("Data");
        var pages = new DocBinaryFormattedDiskPages(wordDocument, tableStream, fib, dataStream);
        var styles = new DocBinaryStyleSheet(tableStream, fib.StyleSheet);
        int paragraphRuns = 0, characterRuns = 0;
        foreach (DocBinaryFormattedDiskPages.PropertyRun run in pages.AllRuns(isParagraph: true))
        {
            paragraphRuns++;
            if (run.GrpprlLength <= 2) continue;
            (byte[] grpprl, int start, int length) = pages.ExpandHugePapx(run.Page, run.GrpprlStart + 2, run.GrpprlLength - 2);
            CountSprms(grpprl, start, length, HandledParagraphSprms, "PAPX", tally);
        }
        foreach (DocBinaryFormattedDiskPages.PropertyRun run in pages.AllRuns(isParagraph: false))
        {
            characterRuns++;
            if (run.GrpprlLength > 0) CountSprms(run.Page, run.GrpprlStart, run.GrpprlLength, HandledCharacterSprms, "CHPX", tally);
        }
        for (int index = 0; index < 4096; index++)
        {
            if (styles[index] is not { } style) continue;
            if (style.ParagraphGrpprl.Length > 2) CountSprms(style.ParagraphGrpprl, 2, style.ParagraphGrpprl.Length - 2, HandledParagraphSprms, "style PAPX", tally);
            CountSprms(style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, HandledCharacterSprms, "style CHPX", tally);
        }
        if (printRuns) Console.WriteLine($"   {paragraphRuns:N0} PAPX runs, {characterRuns:N0} CHPX runs; footnote text {fib.FootnoteCharacterCount:N0} chars, endnote text {fib.EndnoteCharacterCount:N0} chars");
    }

    private static void CountSprms(byte[] grpprl, int start, int length, HashSet<ushort> handled, string context, AuditTally tally)
    {
        foreach ((ushort opcode, _, _) in DocBinaryPropertyModifiers.Sprms(grpprl, start, length))
            if (!handled.Contains(opcode)) tally.Add($"{context} 0x{opcode:X4} ({SprmName(opcode)})");
    }

    private static void Print(string title, AuditTally tally, bool perFile)
    {
        Console.WriteLine($"   {title}:");
        if (tally.Counts.Count == 0) { Console.WriteLine("      none"); return; }
        IEnumerable<KeyValuePair<string, (long occurrences, int files)>> ordered = perFile
            ? tally.Counts.OrderByDescending(pair => pair.Value.files).ThenByDescending(pair => pair.Value.occurrences)
            : tally.Counts.OrderByDescending(pair => pair.Value.occurrences);
        foreach ((string key, (long occurrences, int files)) in ordered.Take(120))
            Console.WriteLine(perFile ? $"      {files,6:N0} {occurrences,12:N0}  {key}" : $"      {occurrences,10:N0}  {key}");
    }

    /// <summary>Names for the sprms that show up in practice ([MS-DOC] §2.6) — unknown ones print as "?".</summary>
    private static string SprmName(ushort opcode) => opcode switch
    {
        0x4A4F => "CRgFtc0 ascii font", 0x4A50 => "CRgFtc1 east-asian font", 0x4A51 => "CRgFtc2 other font", 0x4A5E => "CFtcBi complex-script font",
        0x4A5F => "CLidBi", 0x486D => "CRgLid0", 0x486E => "CRgLid1", 0x4873 => "CRgLid0_80", 0x4874 => "CRgLid1_80", 0x485F => "CLidBi?",
        0x0800 => "CFRMarkDel", 0x0801 => "CFRMark ins", 0x0802 => "CFFldVanish", 0x6A03 => "CPicLocation", 0x4804 => "CIbstRMark",
        0x6805 => "CDttmRMark", 0x0806 => "CFData", 0x4807 => "CIdslRMark", 0x6A09 => "CSymbol", 0x080A => "CFOle2",
        0x2A0C => "CHighlight", 0x6815 => "CRsidProp", 0x6816 => "CRsidText", 0x6817 => "CRsidRMDel",
        0x0838 => "CFOutline", 0x0839 => "CFShadow", 0x083A => "CFSmallCaps", 0x083B => "CFCaps", 0x0875 => "CFNoProof",
        0x8840 => "CDxaSpace char spacing", 0x4845 => "CHpsKern", 0xCA47 => "CMajority", 0x2A53 => "CFtcDefault?", 0x2859 => "CFtcDefault/sfxtText",
        0x0854 => "CFImprint", 0x0858 => "CFEmboss", 0x2A33 => "CPlain", 0x2A34 => "CKcd", 0x4852 => "CCharScale", 0xCA71 => "CShd",
        0x6877 => "CUnderlineColor", 0x6865 => "CBrc80", 0xCA72 => "CBrc", 0x2A79 => "CCnf?", 0x0882 => "CFComplexScripts",
        0x2407 => "PFKeepFollow?", 0x2405 => "PFKeep", 0x2406 => "PFKeepFollow", 0x2407 + 1 => "PFPageBreakBefore", 0x240C => "PFNoLineNumb",
        0xC60D => "PChgTabsPapx tabs", 0xC615 => "PChgTabs", 0x2423 => "PWr", 0x6424 => "PBrcTop80", 0x6425 => "PBrcLeft80",
        0x6426 => "PBrcBottom80", 0x6427 => "PBrcRight80", 0x6428 => "PBrcBetween80", 0x442D => "PShd80", 0xC64D => "PShd",
        0x2431 => "PFWidowControl", 0x2433 => "PFKinsoku", 0x2434 => "PFWordWrap", 0x2435 => "PFOverflowPunct", 0x2436 => "PFTopLinePunct",
        0x2437 => "PFAutoSpaceDE", 0x2438 => "PFAutoSpaceDN", 0x4439 => "PWAlignFont", 0x2448 => "PISnapBaseLine?", 0x245B => "PFContextualSpacing",
        0x6467 => "PRsid", 0x6470 => "PTableProps?", 0x2462 => "PFDyaBeforeAuto?", 0x4458 => "PDcs drop cap",
        0x245C => "PFDyaBeforeAuto", 0x245D => "PFDyaAfterAuto"[..0] + "PDxaRight", 0x2471 => "PFMirrorIndents?",
        0xD234 => "TDefTableShd?", 0xD606 => "TSetBrc?", 0xD608 => "TDefTable", 0xD609 => "TDefTableShd80", 0x5400 => "TJc80",
        0x9601 => "TDxaLeft", 0x9602 => "TDxaGapHalf", 0x3403 => "TFCantSplit90", 0x3404 => "TTableHeader", 0xD605 => "TTableBorders80",
        0x9407 => "TDyaRowHeight", 0x7629 => "TWidthBefore?", 0xF614 => "TTableWidth", 0x3615 => "TFAutofit?", 0x5664 => "TPc?",
        0x6649 => "PTableProps?", 0x664A => "PTableDepth", 0x244B => "PFInnerTableCell", 0x244C => "PFInnerTtp", 0x6646 => "PHugePapx?",
        0x2416 => "PFInTable", 0x2417 => "PFTtp row end",
        _ => "?",
    };
}
