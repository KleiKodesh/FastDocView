using System.Diagnostics;
using FastDocView.Rtf;

namespace FastDocView.Tests;

/// <summary>
/// Which RTF control words the reader does not act on, counted over real files. Words inside skipped destinations
/// (header/footer, info, pictures, bookmarks, \* unknowns, ...) are not counted — skipping them is the design.
/// What is left is either handled, ignored by design (listed below with the reason), or unhandled (printed; keep at 0).
/// The corpus mode also opens and renders every file: 0 crashes.
/// </summary>
public static class RtfFormattingAudit
{
    /// <summary>Control words read but deliberately not rendered, by reason (prefix match when the key ends with '*').</summary>
    public static readonly (string reason, string[] words)[] IgnoredByDesign =
    [
        ("document / section / page layout (flow view has no pages)", [
            "paperw", "paperh", "margl", "margr", "margt", "margb", "gutter", "landscape", "pgwsxn", "pghsxn", "marglsxn", "margrsxn", "margtsxn", "margbsxn",
            "guttersxn", "headery", "footery", "sectd", "sbknone", "sbkpage", "sbkcol", "sbkeven", "sbkodd", "cols", "colsx", "colno", "colw", "colsr", "linebetcol",
            "endnhere", "titlepg", "facingp", "pgnstart", "pgnstarts", "pgnrestart", "pgncont", "pgndec", "pgnx", "pgny", "vertalt", "vertalc", "vertalj", "vertalb",
            "sectdefaultcl", "sectlinegrid", "sectspecifyl", "sectspecifycl", "sectspecifygenN", "sectunlocked", "sectrsid", "ltrsect", "rtlsect", "linex", "linemod",
            "lndscpsxn", "pgbrdrhead", "pgbrdrfoot", "mirrormargins", "twoonone", "gutterprl", "rtlgutter", "bookfold", "bookfoldrev", "bookfoldsheets",
            "srauth", "srdate", "binfsxn", "binsxn", "pgnhn", "pgnhnsn", "pgnhnsh", "pgnhnsp", "pgnhnsc", "pgnhnsm", "pgnid", "stextflow", "horzsect"]),
        ("document settings with no effect on text (compatibility, view, grid, proofing, protection, defaults)", [
            "viewkind", "viewscale", "viewzk", "viewbksp", "viewnobound", "deftab", "deflang", "deflangfe", "adeflang", "adeff", "stshfdbch", "stshfloch",
            "stshfhich", "stshfbi", "themelang", "themelangfe", "themelangcs", "widowctrl", "ftnbj", "aenddoc", "enddoc", "aendnotes", "endnotes", "ftntj", "aftntj",
            "ftnrstcont", "aftnrstcont", "ftnrestart", "aftnrestart", "ftnrstpg", "fet", "noxlattoyen", "expshrtn", "noultrlspc", "dntblnsbdb", "nospaceforul",
            "formshade", "horzdoc", "dgmargin", "dghspace", "dgvspace", "dghorigin", "dgvorigin", "dghshow", "dgvshow", "jcompress", "jexpand", "lytexcttp",
            "lytprtmet", "lytcalctblwd", "lnongrid", "allowfieldendsel", "nobrkwrptbl", "snaptogridincell", "wrppunct", "asianbrkrule", "newtblstyruls",
            "nogrowautofit", "usenormstyforlist", "noindnmbrts", "felnbrelev", "nocxsptable", "indrlsweleven", "noafcnsttbl", "afelev", "utinl", "hwelev",
            "spltpgpar", "notcvasp", "notbrkcnstfrctbl", "notvatxbx", "krnprsnet", "cachedcolbal", "nouicompat", "nofeaturethrottle", "validatexml",
            "showplaceholdtext", "ignoremixedcontent", "saveinvalidxml", "showxmlerrors", "sprstsp", "otblrul", "brkfrm", "sswrap", "truncex", "nolead",
            "msmcap", "lytexcttp", "alntblind", "nolnhtadjtbl", "rsidroot", "mnodirtyhfmrk", "fromtext", "fromhtml", "htmautsp", "htmlrtf", "htmlbase",
            "doctype", "docvar", "readprot", "annotprot", "formprot", "revprot", "readonlyrecommended", "protlevel", "enforceprot", "trackmoves",
            "trackformatting", "revisions", "revprop", "revbar", "makebackup", "linkstyles", "psover", "pubauto", "transmf", "margmirror", "splytwnine",
            "ftnlytwnine", "htmautsp", "useltbaln", "alntblind", "lytcalctblwd", "lyttblrtgr", "lnbrkrule", "bdrrlswsix", "oldas", "ApplyBrkRules",
            "rempersonalinfo", "nolnhtadjtbl", "dntultrlspc", "dntblnsbdb", "dgsnap", "dgmargin", "donotembedsysfont", "donotembedlingdata",
            "grfdocevents", "themelang", "stylesortmethod", "fchars", "lchars", "jclisttab", "hyphhotz", "hyphauto", "hyphcaps", "hyphconsec", "pgbrdropt",
            "ogutter", "deftabN", "doctemp", "rtlgutter", "gutterprl", "pnseclvl", "rtldoc", "ltrdoc", "aftnnar", "nextfile", "template", "fet", "bdbfhdr",
            "remdttm", "lytcalctblwd", "wraptrsp", "prcolbl", "noextrasprl", "nocolbal", "nospaceforul", "dntblnsbdb", "noxlattoyen", "nolead", "noqfpromote",
            "ansi", "uc", "viewbksp", "paperw"]),
        ("font table / style sheet attributes (family, pitch, panose; style links, priorities, quick-format flags)", [
            "fnil", "froman", "fswiss", "fmodern", "fscript", "fdecor", "ftech", "fbidi", "fprq", "fbias", "fttruetype", "ftnil", "snext", "sbasedon", "slink",
            "sautoupd", "shidden", "slocked", "spersonal", "scompose", "sreply", "sqformat", "spriority", "ssemihidden", "sunhideused", "additive", "styrsid",
            "ctrl", "tsrowd", "tscellwidthfts", "tscellwidth", "tsvertalt", "tsvertalc", "tsvertalb", "tsbrdrt", "tsbrdrb", "tsbrdrl", "tsbrdrr", "tsbrdrh", "tsbrdrv",
            "tscellpaddl", "tscellpaddt", "tscellpaddb", "tscellpaddr", "tscellpaddfl", "tscellpaddft", "tscellpaddfb", "tscellpaddfr", "tsnowrap", "tscbandsh",
            "tscbandsv", "tsbghoriz", "tsbgvert", "tsbgfdiag", "tsbgbdiag", "tsbgcross", "tsbgdcross", "tscellcfpat", "tscellcbpat", "tscellpct", "tblind",
            "tblindtype", "tsbrdrdgl", "tsbrdrdgr", "listtemplateid", "listhybrid", "listsimple", "listrestarthdn", "listoverridecount", "listoverrideformat",
            "leveljcn", "levelindent", "levelspace", "levelold", "levelprev", "levelprevspace", "levelpicture", "leveltemplateid",
            "levelpicturenosize", "lvltentative", "liststyleid", "lsdlockeddef", "lsdsemihiddendef", "lsdunhideuseddef", "lsdqformatdef", "lsdprioritydef",
            "lsdstimax", "lsdlocked", "lsdsemihidden", "lsdunhideused", "lsdqformat", "lsdpriority", "csN", "keepn", "colorschememapping"]),
        ("pagination / layout hints (keep, widow, hyphenation, snapping, frames positions)", [
            "keep", "keepn", "widctlpar", "nowidctlpar", "hyphpar", "nohyph", "noline", "sbys", "aspalpha", "aspnum", "faauto", "faroman", "favar", "fafixed",
            "facenter", "adjustright", "nosnaplinegrid", "nooverflow", "noautofit", "contextualspace", "mirrorindents", "absw", "absh", "abslock", "phmrg",
            "phpg", "phcol", "posx", "posxc", "posxi", "posxo", "posxl", "posxr", "posy", "posyc", "posyt", "posyb", "posyin", "posyout", "posyil", "pvmrg", "pvpg",
            "pvpara", "dxfrtext", "dfrmtxtx", "dfrmtxty", "wraparound", "wraptight", "nowrap", "overlay", "dropcapli", "dropcapt", "frmtxlrtb", "frmtxtbrl",
            "frmtxbtlr", "frmtxlrtbv", "frmtxtbrlv", "txbxtwno", "txbxtwalways", "txbxtwfirst", "txbxtwlast", "txbxtwfirstlast", "outlinelevel9", "itap0",
            "sbauto", "saauto", "lisa", "lisb", "pnrnot", "pnhang", "pnrauth", "pnrdate", "pnrstart", "pnrstop", "pnrxst", "pnrrgb", "pnrnfc", "pnrpnbr",
            "yts", "ilfomacatclnup", "nestinglevel"]),
        ("tab stops (CSS has no positioned tabs; the default tab-size is used)", [
            "tx", "tb", "tqr", "tqc", "tqdec", "tldot", "tlhyph", "tlul", "tlth", "tleq", "tlmdot", "jclisttab"]),
        ("borders and shading details (paragraph borders not rendered yet)", [
            "brdrt", "brdrb", "brdrl", "brdrr", "brdrbtw", "brdrbar", "box", "brdrs", "brdrth", "brdrsh", "brdrdb", "brdrdot", "brdrdash", "brdrhair", "brdrw",
            "brdrcf", "brsp", "brdrnone", "brdrnil", "brdrtbl", "brdrinset", "brdroutset", "brdrtriple", "brdrdashsm", "brdrdashd", "brdrdashdd", "brdrthtnsg",
            "brdrtnthsg", "brdrtnthtnsg", "brdrwavy", "brdrwavydb", "brdrengrave", "brdremboss", "brdrframe", "brdrsh", "shading", "cfpat", "cbpat", "bgbdiag",
            "bgcross", "bgdcross", "bgdkbdiag", "bgdkcross", "bgdkdcross", "bgdkfdiag", "bgdkhoriz", "bgdkvert", "bgfdiag", "bghoriz", "bgvert", "chshdng",
            "chcfpat", "chbrdr", "chbgbdiag", "chbgcross", "chbgdcross", "chbgdkbdiag", "chbgdkcross", "chbgdkdcross", "chbgdkfdiag", "chbgdkhoriz", "chbgdkvert",
            "chbgfdiag", "chbghoriz", "chbgvert", "brdrdashdotstr", "brdrthtnlg", "brdrtnthlg", "brdrtnthtnlg", "brdrthtnmg", "brdrtnthmg", "brdrtnthtnmg",
            "brdrdashdot", "brdrdashdotdot", "brdrart"]),
        ("table geometry beyond cell widths (gaps, heights, merges, cell borders, autofit, banding, preferred widths)", [
            "trgaph", "trleft", "trrh", "trqc", "trql", "trqr", "trhdr", "trkeep", "trkeepfollow", "trbrdrt", "trbrdrb", "trbrdrl", "trbrdrr", "trbrdrh", "trbrdrv",
            "trautofit", "trftsWidth", "trwWidth", "trftsWidthB", "trwWidthB", "trftsWidthA", "trwWidthA", "trpaddl", "trpaddr", "trpaddt", "trpaddb",
            "trpaddfl", "trpaddfr", "trpaddft", "trpaddfb", "trspdl", "trspdr", "trspdt", "trspdb", "trspdfl", "trspdfr", "trspdft", "trspdfb", "trcbpat",
            "trcfpat", "trpat", "trshdng", "tblrsid", "tbllkhdrrows", "tbllklastrow", "tbllkhdrcols", "tbllklastcol", "tbllknorowband", "tbllknocolband",
            "tbllkborder", "tbllkshading", "tbllkfont", "tbllkcolor", "tbllkbestfit", "tblind", "tblindtype", "taprtl", "irowband", "irow", "lastrow", "tdfrmtxtLeft",
            "tdfrmtxtRight", "tdfrmtxtTop", "tdfrmtxtBottom", "tposx", "tposxc", "tposxl", "tposxr", "tposxi", "tposxo", "tposy", "tposyc", "tposyt", "tposyb",
            "tposyil", "tposyin", "tposyout", "tposytop", "tposnegx", "tposnegy", "tphmrg", "tphpg", "tphcol", "tpvmrg", "tpvpg", "tpvpara", "tabsnoovrlp",
            "clvertalt", "clvertalc", "clvertalb", "clbrdrt", "clbrdrb", "clbrdrl", "clbrdrr", "cltxlrtb", "cltxtbrl", "cltxbtlr", "cltxlrtbv", "cltxtbrlv",
            "clFitText", "clNoWrap", "clftsWidth", "clwWidth", "clpadl", "clpadr", "clpadt", "clpadb", "clpadfl", "clpadfr", "clpadft", "clpadfb", "clspl",
            "clspr", "clspt", "clspb", "clspfl", "clspfr", "clspft", "clspfb", "clmgf", "clmrg", "clvmgf", "clvmrg", "clcfpat", "clshdng", "clshdngraw",
            "clcbpatraw", "clcfpatraw", "clhidemark", "clbgbdiag", "clbgcross", "clbgdcross", "clbgdkbdiag", "clbgdkcross", "clbgdkdcross", "clbgdkfdiag",
            "clbgdkhor", "clbgdkvert", "clbgfdiag", "clbghoriz", "clbgvert", "cldglu", "cldgll", "rtlrow", "ltrrow", "trowd", "nestrow", "tsrowd", "trflat"]),
        ("character properties without a visible effect here (language, kerning, scaling, script hints, revision ids)", [
            "lang", "langfe", "langnp", "langfenp", "alang", "noproof", "kerning", "charscalex", "fcs", "loch", "hich", "dbch", "cgrid", "insrsid", "charrsid",
            "delrsid", "pararsid", "sectrsid", "tblrsid", "rsid", "styrsid", "oldlinewrap", "ltrmark", "rtlmark", "cchs", "afs", "expnd", "webhidden", "specvanish",
            "spv", "ab", "ai", "acaps", "adn", "aexpnd", "aoutl", "ascaps", "ashad", "astrike", "aul", "auld", "auldb", "aulnone", "aulw", "aup", "aspalpha",
            "itap", "accnone", "accdot", "acccomma", "acccircle", "accunderdot", "dn", "up", "ulc", "deleted", "revised", "revauth", "revauthdel", "revdttm",
            "revdttmdel", "crauth", "crdate", "mhtmltag", "pnhang", "nosectexpand", "g", "gcw", "gridtbl", "chshdng", "faauto", "f", "pnf", "pnfs", "pnindent",
            "pnstart", "pnlvlbody", "pnlvlblt", "pnlvlcont", "pnlvl", "pndec", "pnord", "pnordt", "pncard", "pnlcltr", "pnlcrm", "pnucltr", "pnucrm", "pnb",
            "pni", "pnul", "pnuld", "pnuldb", "pnulnone", "pnulw", "pncf", "pnqc", "pnql", "pnqr", "pnsp", "pnprev", "pnnumonce", "pnacross", "pnaiu",
            "pnaiud", "pnbidia", "pnbidib", "pnchosung", "pncnum", "pndbnum", "pndbnumd", "pndbnumk", "pndbnuml", "pndbnumt", "pndecd", "pnganada", "pngbnum",
            "pngbnumd", "pngbnumk", "pngbnuml", "pniroha", "pnirohad", "pnzodiac", "pnzodiacd", "pnzodiacl", "pnaiueo", "pnaiueod", "pnrestart", "pnscaps",
            "pncaps", "pnstrike"]),
        ("theme references next to explicit values: Word 2007 writes the theme slot AND the resolved font / RGB (flomajor.. in the font table, ctextone ctint cshade.. in the color table)", [
            "flomajor", "flominor", "fhimajor", "fhiminor", "fdbmajor", "fdbminor", "fbimajor", "fbiminor", "ctextone", "ctexttwo", "cbackgroundone",
            "cbackgroundtwo", "caccentone", "caccenttwo", "caccentthree", "caccentfour", "caccentfive", "caccentsix", "chyperlink", "cfollowedhyperlink",
            "cmaindarkone", "cmaindarktwo", "cmainlightone", "cmainlighttwo", "ctint", "cshade"]),
        ("field and revision metadata (the field result and the text are what shows)", [
            "fldedit", "flddirty", "fldlock", "fldpriv", "prauth", "prdate", "crauth", "crdate", "animtext"]),
        ("layout / compatibility flags with no text effect (frame wrapping default, VML, paper size, note placement, RichEdit bidi flag, picture-bullet resolution)", [
            "wrapdefault", "wrapnoaround", "wraparound", "wraptight", "wrapthrough", "relyonvml", "psz", "sftnbj", "sftntj", "saftnnar", "sftnnar", "fbidis",
            "hres", "chhres"]),
        ("not in the RTF spec nor in LibreOffice's keyword table: generator-specific words (qs, ftstart seen in one tool's output)", [
            "qs", "ftstart"]),
        ("text-flow / version metadata outside the document text", [
            "version", "vern", "edmins", "nofpages", "nofwords", "nofchars", "nofcharsws", "id", "yr", "mo", "dy", "hr", "min", "sec", "nextfile", "doccomm",
            "hlinkbase", "operator", "company", "manager", "category", "keywords", "comment", "title", "subject", "author", "creatim", "revtim", "printim",
            "buptim", "dbch"]),
    ];

    private static readonly Dictionary<string, string> IgnoredReasonByWord = BuildIgnored();

    private static Dictionary<string, string> BuildIgnored()
    {
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach ((string reason, string[] words) in IgnoredByDesign)
            foreach (string word in words) map[word] = reason;
        return map;
    }

    private sealed class AuditSink : RtfSink
    {
        public readonly Dictionary<string, int> Unhandled = new(StringComparer.Ordinal);
        public readonly Dictionary<string, int> Ignored = new(StringComparer.Ordinal);
        public readonly Dictionary<string, int> Skipped = new(StringComparer.Ordinal);
        public override bool WantsText => false;
        public override bool WantsUnhandledWords => true;

        public override void UnhandledWord(ReadOnlySpan<byte> name)
        {
            string word = System.Text.Encoding.ASCII.GetString(name.ToArray());
            Increment(IgnoredReasonByWord.ContainsKey(word) ? Ignored : Unhandled, word);
        }

        public override void DestinationSkipped(ReadOnlySpan<byte> name)
            => Increment(Skipped, name.IsEmpty ? "(destination)" : System.Text.Encoding.ASCII.GetString(name.ToArray()));
    }

    private static void Increment(Dictionary<string, int> counts, string key) => counts[key] = counts.TryGetValue(key, out int count) ? count + 1 : 1;

    private static AuditSink AuditFile(string path)
    {
        var sink = new AuditSink();
        var parser = new RtfParser(sink, new RtfDocumentTables(), readsHeaderTables: true);
        using FileStream file = WordDocumentFileStream.OpenShared(path);
        parser.Run(file, 0, file.Length);
        return sink;
    }

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"RTF formatting audit: {Path.GetFileName(path)}");
            AuditSink sink = AuditFile(path);
            Print(sink.Unhandled, sink.Ignored, sink.Skipped, null);
        }
    }

    /// <summary>Every .rtf in the list: audit + open + render a window at the start and the middle; below-normal priority, one file at a time.</summary>
    public static void RunCorpus(string listPath, int maxFiles)
    {
        try { Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.BelowNormal; } catch { }
        string[] paths = File.ReadAllLines(listPath).Select(line => line.Trim().TrimStart('﻿')).Where(line => line.Length > 0).Take(maxFiles).ToArray();
        TestReport.Step($"RTF corpus: {paths.Length:N0} files from {listPath}");
        var unhandled = new Dictionary<string, int>(StringComparer.Ordinal);
        var unhandledFiles = new Dictionary<string, int>(StringComparer.Ordinal);
        var ignored = new Dictionary<string, int>(StringComparer.Ordinal);
        var skipped = new Dictionary<string, int>(StringComparer.Ordinal);
        var failures = new List<string>();
        var suspectEncoding = new List<string>();
        var policy = new WordIndexPolicy { ForcedMode = WordIndexMode.InMemory };
        var watch = Stopwatch.StartNew();
        int done = 0, opened = 0, notRtf = 0, unavailable = 0;
        long blocks = 0;
        (string path, long bytes, double openMs, double windowMs) slowest = ("", 0, 0, 0);
        foreach (string path in paths)
        {
            done++;
            if (done % 100 == 0) Console.WriteLine($"   [{watch.Elapsed:mm\\:ss}] {done:N0} / {paths.Length:N0} files, {failures.Count} failures");
            if (!File.Exists(path)) { failures.Add($"missing: {path}"); continue; }
            try { using FileStream probe = WordDocumentFileStream.OpenShared(path); probe.ReadByte(); }
            catch (IOException) { unavailable++; continue; }                // e.g. a OneDrive placeholder while the cloud provider is off
            try
            {
                if (WordWindowedDocumentOpener.DetectFormat(path) != WordWindowedDocumentOpener.WordFileFormat.RichText) { notRtf++; continue; }
                AuditSink sink = AuditFile(path);
                foreach (KeyValuePair<string, int> pair in sink.Unhandled) { unhandled[pair.Key] = unhandled.GetValueOrDefault(pair.Key) + pair.Value; Increment(unhandledFiles, pair.Key); }
                foreach (KeyValuePair<string, int> pair in sink.Ignored) ignored[pair.Key] = ignored.GetValueOrDefault(pair.Key) + pair.Value;
                foreach (KeyValuePair<string, int> pair in sink.Skipped) skipped[pair.Key] = skipped.GetValueOrDefault(pair.Key) + pair.Value;
            }
            catch (Exception exception) { failures.Add($"audit {exception.GetType().Name}: {exception.Message} - {path}"); }
            try
            {
                var openWatch = Stopwatch.StartNew();
                using IWordWindowedDocument document = WordWindowedDocumentOpener.Open(path, policy);
                double openMs = openWatch.Elapsed.TotalMilliseconds;
                var windowWatch = Stopwatch.StartNew();
                IReadOnlyList<WordDocumentBlockHtml> first = document.RenderBlocks(0, 40);
                if (document.BlockCount > 80) document.RenderBlocks(document.BlockCount / 2, 40);
                if (document.BlockCount > 40) document.RenderBlocks(document.BlockCount - 40, 40);
                double windowMs = windowWatch.Elapsed.TotalMilliseconds;
                if (document.BlockCount > 0 && first.Count == 0) failures.Add($"empty first window: {path}");
                // wrong code page shows as Latin-1 letters where Hebrew was meant (cp1252 decoding of cp1255 bytes)
                int hebrew = 0, latinLetters = 0;
                foreach (WordDocumentBlockHtml block in first)
                    foreach (char c in block.Html)
                    {
                        if (c is >= (char)0x05D0 and <= (char)0x05EA) hebrew++;
                        else if (c is >= (char)0x00C0 and <= (char)0x00FF) latinLetters++;
                    }
                if (latinLetters > 20 && latinLetters > hebrew) suspectEncoding.Add($"{latinLetters} Latin-1 letters vs {hebrew} Hebrew: {path}");
                opened++;
                blocks += document.BlockCount;
                long bytes = new FileInfo(path).Length;
                if (bytes > slowest.bytes) slowest = (path, bytes, openMs, windowMs);
            }
            catch (Exception exception) { failures.Add($"open {exception.GetType().Name}: {exception.Message} - {path}"); }
        }
        Console.WriteLine($"   done in {watch.Elapsed:mm\\:ss}: {opened:N0} opened + rendered, {blocks:N0} blocks, {notRtf} not RTF, {unavailable} not readable (cloud placeholders)");
        Console.WriteLine($"   largest: {slowest.bytes >> 10} KB, open {slowest.openMs:0} ms, 3 windows {slowest.windowMs:0} ms");
        Print(unhandled, ignored, skipped, unhandledFiles);
        Console.WriteLine("   failures:");
        if (failures.Count == 0) Console.WriteLine("      none");
        foreach (string failure in failures.Take(30)) Console.WriteLine("      " + Ascii(failure));
        Console.WriteLine($"   first windows that look mis-decoded: {suspectEncoding.Count}");
        foreach (string suspect in suspectEncoding.Take(20)) Console.WriteLine("      " + Ascii(suspect));
        TestReport.Check(failures.Count == 0, $"RTF corpus: 0 crashes ({failures.Count})");
        TestReport.Check(unhandled.Count == 0, $"RTF corpus: unhandled control words at 0 ({unhandled.Count})");
    }

    private static void Print(Dictionary<string, int> unhandled, Dictionary<string, int> ignored, Dictionary<string, int> skipped, Dictionary<string, int>? files)
    {
        Console.WriteLine($"   unhandled control words ({unhandled.Count}):" + (unhandled.Count == 0 ? " none" : ""));
        foreach (KeyValuePair<string, int> pair in unhandled.OrderByDescending(p => p.Value).Take(80))
            Console.WriteLine($"      \\{pair.Key,-24} {pair.Value,9:N0}" + (files is null ? "" : $"  in {files[pair.Key]} files"));
        Console.WriteLine($"   ignored by design: {ignored.Count} words, {ignored.Values.Sum():N0} occurrences; skipped destinations: "
                          + string.Join(", ", skipped.OrderByDescending(p => p.Value).Take(25).Select(p => $"{p.Key} {p.Value:N0}")));
    }

    private static string Ascii(string text)
    {
        var builder = new System.Text.StringBuilder(text.Length);
        foreach (char c in text) builder.Append(c < 128 ? c : '?');
        return builder.ToString();
    }
}
