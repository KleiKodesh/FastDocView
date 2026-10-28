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
         "rsid", "rsidR", "rsidRPr", "w", "eastAsianLayout", "specVanish", "oMath", "framePr"];
    private static readonly HashSet<string> HandledRunProperties =
        ["rStyle", "b", "bCs", "i", "iCs", "u", "strike", "dstrike", "vanish", "webHidden", "rtl", "vertAlign", "sz", "szCs", "color",
         "highlight", "shd", "rFonts", "smallCaps", "caps", "spacing", "cs"];
    private static readonly HashSet<string> HandledParagraphProperties =
        ["pStyle", "bidi", "jc", "ind", "spacing", "outlineLvl", "numPr", "rPr", "ilvl", "numId", "contextualSpacing", "shd", "tabs", "sectPr"];
    private static readonly HashSet<string> HandledRunChildren =
        ["rPr", "t", "delText", "tab", "ptab", "br", "cr", "noBreakHyphen", "softHyphen", "sym", "footnoteReference", "endnoteReference",
         "footnoteRef", "endnoteRef", "separator", "continuationSeparator", "drawing", "pict", "object", "fldChar", "instrText", "lastRenderedPageBreak"];
    private static readonly HashSet<string> HandledParagraphChildren =
        ["pPr", "r", "hyperlink", "ins", "smartTag", "sdt", "customXml", "fldSimple", "dir", "bdo", "del",
         "bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd"];
    private static readonly HashSet<string> HandledBodyChildren = ["p", "tbl", "sdt", "sectPr", "bookmarkStart", "bookmarkEnd"];
    private static readonly HashSet<string> HandledTableParts =
        ["tblPr", "tblGrid", "tr", "tc", "tcPr", "gridSpan", "vMerge", "shd", "bidiVisual", "gridCol", "trPr", "p", "tbl",
         "tcW", "tcBorders", "vAlign", "tblW", "tblBorders", "tblInd", "tblCellMar", "tblStyle", "jc", "sdt"];

    private static readonly HashSet<ushort> HandledParagraphSprms =
        [0x2403, 0x2461, 0x2441, 0x840F, 0x845E, 0x840E, 0x845D, 0x8411, 0x8460, 0xA413, 0xA414, 0x6412, 0x2416, 0x260A, 0x460B, 0x2640, 0x4600,
         0x245B, 0xC64D, 0x442D, 0xC60D, 0xC615, 0x2417, 0x244C, 0xD608,
         // ignored by design: revision ids, pagination / widow / keep, auto-spacing, kinsoku, frames, table layout hints a flow row does not need
         0x6467, 0x2431, 0x2437, 0x2438, 0x2448, 0x2406, 0x2405, 0x2407, 0x240C, 0x4439, 0x2433, 0x2434, 0x2435, 0x2436, 0x245C, 0x245D,
         0x246D, 0x8419, 0x261B, 0x2423, 0x6649, 0x664A, 0x244B, 0x6646, 0x560B, 0xD612, 0xD613, 0xF617, 0x3466, 0x548A, 0x740A, 0x9602, 0x3403, 0x3404, 0x5400, 0x9407, 0xD634, 0xD61A, 0xD61B, 0xD61C, 0xD61D,
         0xF661, 0x7479, 0xD670, 0xF618, 0x646B, 0xD642, 0xD605, 0xD609, 0xF614, 0x9601, 0x3615];
    private static readonly HashSet<ushort> HandledCharacterSprms =
        [0x0835, 0x085C, 0x0836, 0x085D, 0x0837, 0x083C, 0x085A, 0x2A3E, 0x2A48, 0x4A43, 0x4A61, 0x2A42, 0x6870, 0x2A0C, 0x0855, 0x4A30,
         0xCA71, 0x4866, 0x083A, 0x083B, 0x8840, 0x4A4F, 0x4A5E,
         // ignored by design: revision ids, language ids, proofing, script hint, east-asian/other-script fonts, field data, picture location
         0x6816, 0x6815, 0x6817, 0x286F, 0x486E, 0x4874, 0x485F, 0x486D, 0x4873, 0x0875, 0x4A50, 0x4A51, 0x0811, 0x0802, 0x0806, 0x6A03,
         0x6A09, 0x0800, 0x0801, 0x4804, 0x6805, 0x4807, 0x080A, 0x0882, 0x484B];

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"Formatting audit: {Path.GetFileName(path)}");
            try
            {
                if (WordWindowedDocumentOpener.DetectFormat(path) == WordWindowedDocumentOpener.WordFileFormat.Docx) AuditDocx(path);
                else AuditDocBinary(path);
            }
            catch (Exception exception) { Console.WriteLine($"   audit failed: {exception.GetType().Name}: {exception.Message}"); }
        }
    }

    // ── .docx ──────────────────────────────────────────────────────────────────────────

    private static void AuditDocx(string path)
    {
        var unhandled = new Dictionary<string, int>(StringComparer.Ordinal);
        var used = new Dictionary<string, int>(StringComparer.Ordinal);
        using var zip = ZipFile.OpenRead(path);
        foreach (string part in new[] { "word/document.xml", "word/footnotes.xml", "word/endnotes.xml", "word/styles.xml", "word/numbering.xml" })
        {
            ZipArchiveEntry? entry = zip.GetEntry(part);
            if (entry is null) continue;
            Console.WriteLine($"   part {part}: {entry.Length >> 10} KB");
            using Stream stream = entry.Open();
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true, DtdProcessing = DtdProcessing.Ignore });
            var parents = new Stack<string>();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement) { if (parents.Count > 0) parents.Pop(); continue; }
                if (reader.NodeType != XmlNodeType.Element) continue;
                string name = reader.LocalName;
                string parent = parents.Count > 0 ? parents.Peek() : "";
                string key = $"{part.Substring(5)}: {parent}/{(reader.Prefix.Length > 0 ? reader.Prefix + ":" : "")}{name}";
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
                if (parent is "rPr" or "pPr" or "r" or "p" or "body" or "tbl" or "tr" or "tc" or "tcPr" or "tblPr")
                    used[key] = used.TryGetValue(key, out int u) ? u + 1 : 1;
                if (!isHandled) unhandled[key] = unhandled.TryGetValue(key, out int c) ? c + 1 : 1;
                if (!reader.IsEmptyElement) parents.Push(name);
            }
        }
        Print("unhandled .docx constructs", unhandled);
        Console.WriteLine($"   ({used.Count} distinct constructs used in rendering contexts, {unhandled.Count} unhandled)");
    }

    // ── .doc ───────────────────────────────────────────────────────────────────────────

    private static void AuditDocBinary(string path)
    {
        using var document = DocBinaryWindowedDocument.Open(path, new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") });
        var unhandled = new Dictionary<string, int>(StringComparer.Ordinal);
        int paragraphRuns = 0, characterRuns = 0;
        foreach (DocBinaryFormattedDiskPages.PropertyRun run in document.FormattedDiskPages.AllRuns(isParagraph: true))
        {
            paragraphRuns++;
            if (run.GrpprlLength > 2) CountSprms(run.Page, run.GrpprlStart + 2, run.GrpprlLength - 2, HandledParagraphSprms, "PAPX", unhandled);
        }
        foreach (DocBinaryFormattedDiskPages.PropertyRun run in document.FormattedDiskPages.AllRuns(isParagraph: false))
        {
            characterRuns++;
            if (run.GrpprlLength > 0) CountSprms(run.Page, run.GrpprlStart, run.GrpprlLength, HandledCharacterSprms, "CHPX", unhandled);
        }
        for (int index = 0; index < 4096; index++)
        {
            if (document.StyleSheet[index] is not { } style) continue;
            if (style.ParagraphGrpprl.Length > 2) CountSprms(style.ParagraphGrpprl, 2, style.ParagraphGrpprl.Length - 2, HandledParagraphSprms, "style PAPX", unhandled);
            CountSprms(style.CharacterGrpprl, 0, style.CharacterGrpprl.Length, HandledCharacterSprms, "style CHPX", unhandled);
        }
        DocBinaryFileInformationBlock fib = document.FileInformationBlock;
        Console.WriteLine($"   {paragraphRuns:N0} PAPX runs, {characterRuns:N0} CHPX runs; footnote text {fib.FootnoteCharacterCount:N0} chars");
        Print("unhandled .doc sprms", unhandled);
    }

    private static void CountSprms(byte[] grpprl, int start, int length, HashSet<ushort> handled, string context, Dictionary<string, int> unhandled)
    {
        foreach ((ushort opcode, _, _) in DocBinaryPropertyModifiers.Sprms(grpprl, start, length))
        {
            if (handled.Contains(opcode)) continue;
            string key = $"{context} 0x{opcode:X4} ({SprmName(opcode)})";
            unhandled[key] = unhandled.TryGetValue(key, out int count) ? count + 1 : 1;
        }
    }

    private static void Print(string title, Dictionary<string, int> counts)
    {
        Console.WriteLine($"   {title}:");
        if (counts.Count == 0) { Console.WriteLine("      none"); return; }
        foreach ((string key, int count) in counts.OrderByDescending(pair => pair.Value).Take(80))
            Console.WriteLine($"      {count,10:N0}  {key}");
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
