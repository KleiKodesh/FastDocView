// FastDocView test harness.
//
//   dotnet run -c Release                          deflate round trip (8 MB) + synthetic docx (50k blocks) + real files passed/found
//   dotnet run -c Release -- --blocks 500000       bigger synthetic document
//   dotnet run -c Release -- --deflate-mb 64       bigger deflate round trip
//   dotnet run -c Release -- --real <path>...      real .docx/.doc files to smoke-test (repeatable)
//   dotnet run -c Release -- --skip-synthetic      only real files
//   dotnet run -c Release -- --page-width [--real <path>...]   page text width per section (synthetic + the given files)
//   dotnet run -c Release -- --huge 1000000        instant reload + flat heap on a ~350 MB body
//   dotnet run -c Release -- --incremental-open [N]  first open with the index built on demand (synthetic N blocks, default 1M, + --real .docx files)
//   dotnet run -c Release -- --notes-streaming 20000   notes read on demand from a 100 MB+ footnotes part
//   dotnet run -c Release -- --audit <path>...     formatting constructs the renderers do not handle
//   dotnet run -c Release -- --odt-incremental [N] first open of a synthetic .odt and .fodt with the index built on demand (N blocks, default 200000)
//   dotnet run -c Release -- --streaming-incremental [N]  on-demand first open of synthetic .rtf / .txt / MHTML (N blocks, default 300000) + cancelled opens
//   dotnet run -c Release -- --odt [N]             synthetic .odt/.fodt (N repeating blocks, default 20000); with LibreOffice installed,
//                                                  every --real .docx is also converted to .odt (TestOutput) and smoke-tested + audited
//   dotnet run -c Release -- --text [N] / --mhtml [N]   synthetic .txt (every encoding / line ending) and MHTML (qp + base64), N filler blocks; --mhtml also runs the .html checks
//   dotnet run -c Release -- --torat-emet <folder> [--torat-emet-sample N | --torat-emet-stats]   Torat Emet books: sample (fixed seed, 300) or statistics pass
//   dotnet run -c Release -- --text-corpus <list> [--text-largest 200 --text-random 300]   .txt corpus sample: encoding mix, 0 crashes
//   dotnet run -c Release -- --mime-scan <word-list>   Word files that start with 'MIME' (web archives): count, open + render each
//   dotnet run -c Release -- --profile-markup <path>...  cold / reopen / window ms of .txt / .mht files
//   dotnet run -c Release -- --rtf [N]             synthetic .rtf (N repeating blocks, default 20000); with --corpus, the RTF files of the list
//                                                  are audited (unhandled control words) and opened + rendered; --rtf-audit <path>... audits files
//   dotnet run -c Release -- --skip-synthetic --corpus ../test-data --corpus-sample 300   quick pass: fixed-seed sample + the fixed-case regressions
//   dotnet run -c Release -- --skip-synthetic --corpus <list.txt>   the same over a whole file list, aggregated + open/render crash check
using FastDocView.Tests;

int blocks = 50_000, deflateMegabytes = 8, hugeBlocks = 0, notesStreamingBlocks = 0;
bool searchOnly = false;
var searchPaths = new List<string>();
bool skipSynthetic = false, incrementalOpen = false, fidelityOnly = false, pageWidthOnly = false;
int incrementalBlocks = 1_000_000;
var realPaths = new List<string>();
var auditPaths = new List<string>();
var profilePaths = new List<string>();
var odtFirstOpenPaths = new List<string>();
var indexProfilePaths = new List<string>();
string? corpusList = null;
string? listLabelsFallback = null;
int corpusMax = int.MaxValue;
int corpusSample = 0;
int odtIncrementalBlocks = 0;
int streamingIncrementalBlocks = 0;
int odtBlocks = 0, rtfBlocks = 0, textBlocks = 0, mhtmlBlocks = 0, textLargest = 200, textRandom = 300;
string? textCorpus = null, mimeScan = null;
string? toratEmetFolder = null; int toratEmetSample = 300; bool toratEmetFull = false;
var markupProfilePaths = new List<string>();
var rtfAuditPaths = new List<string>();
var rtfProfilePaths = new List<string>();
string? goldenDirectory = null;
var compactReportPaths = new List<string>();
var goldenPaths = new List<string>();
var rereadPaths = new List<string>();
var readCountPaths = new List<string>();
var rtfSearchGoldenPaths = new List<string>();
var docReadPaths = new List<string>();
var docSearchGoldenPaths = new List<string>();
for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--blocks": blocks = int.Parse(args[++i]); break;
        case "--huge": hugeBlocks = int.Parse(args[++i]); break;                       // e.g. 1000000 → ~350 MB document.xml
        case "--notes-streaming": notesStreamingBlocks = int.Parse(args[++i]); break;  // e.g. 20000 → 1000 footnotes, ~113 MB footnotes.xml
        case "--deflate-mb": deflateMegabytes = int.Parse(args[++i]); break;
        case "--skip-synthetic": skipSynthetic = true; break;
        case "--search": searchOnly = true; while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) searchPaths.Add(args[++i]); break;   // in-document search checks (+ timings on the given files)
        case "--index-cache": WordDocumentIndexCacheTest.Run(); skipSynthetic = true; break;   // sidecar cache LRU: eviction, caps, stale / orphan cleanup, concurrency (seconds)
        case "--connector-route": OdtConnectorRouteTest.Run(); OdtConnectedShapesTest.Run(); skipSynthetic = true; break;   // only the connector routing checks (seconds)
        case "--mhtml-pictures": WebArchivePicturesTest.Run(); skipSynthetic = true; break;   // web archive pictures (seconds)
        case "--tables": WordTablesFidelityTest.Run(); skipSynthetic = true; break;   // only the table checks (seconds)
        case "--math": OfficeMathFidelityTest.Run(); skipSynthetic = true; break;   // only the equation checks (seconds)   // only the Word field checks (seconds)
        case "--charts": OdtChartsFidelityTest.Run(); skipSynthetic = true; break;   // only the chart checks (seconds); writes TestOutput/charts-*.html
        case "--fields": WordFieldsFidelityTest.Run(); WordFieldsOtherFormatsTest.Run(); skipSynthetic = true; break;   // only the Word field checks (seconds)
        case "--drop-caps": DropCapsAndHyphenationFidelityTest.Run(); skipSynthetic = true; break;
        case "--list-bullets": ListBulletsFidelityTest.Run(); skipSynthetic = true; break;   // only the list bullet / label checks
        case "--body-script-sample": BodyScriptFidelityTest.DumpSample(args[i + 1], int.Parse(args[i + 2]), int.Parse(args[i + 3])); return 0;
        case "--body-script": BodyScriptFidelityTest.Run(); skipSynthetic = true; break;   // Latin / complex-script set per run (Word 16 PDF)
        case "--list-formats": ListNumberFormatsFidelityTest.Run(); skipSynthetic = true; break;   // only the list number-format / bullet checks against Word
        case "--fidelity": fidelityOnly = true; break;                                  // only the small Word-fidelity checks (seconds)
        case "--page-width": pageWidthOnly = true; break;                               // only the page-text-width checks (+ --real files)                                  // only the small Word-fidelity checks (seconds)
        case "--incremental-open":                                                     // optional block count, default 1000000
            incrementalOpen = true;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int parsed)) { incrementalBlocks = parsed; i++; }
            break;
        case "--real": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) realPaths.Add(args[++i]); break;
        case "--audit": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) auditPaths.Add(args[++i]); break;
        case "--odt-first-open": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) odtFirstOpenPaths.Add(args[++i]); break;
        case "--profile-open": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) profilePaths.Add(args[++i]); break;
        case "--profile-index": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) indexProfilePaths.Add(args[++i]); break;
        case "--corpus": corpusList = args[++i]; break;            // a text file with one Word file path per line
        case "--odt-incremental":                                                      // optional block count, default 200000
            odtIncrementalBlocks = 200_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int odtIncrementalCount)) { odtIncrementalBlocks = odtIncrementalCount; i++; }
            break;
        case "--streaming-incremental":                                                // optional block count, default 300000
            streamingIncrementalBlocks = 300_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int streamingCount)) { streamingIncrementalBlocks = streamingCount; i++; }
            break;
        case "--odt":
            odtBlocks = 20_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int odtCount)) { odtBlocks = odtCount; i++; }
            break;
        case "--text":
            textBlocks = 20_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int textCount)) { textBlocks = textCount; i++; }
            break;
        case "--mhtml":
            mhtmlBlocks = 20_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int mhtmlCount)) { mhtmlBlocks = mhtmlCount; i++; }
            break;
        case "--torat-emet": toratEmetFolder = args[++i]; skipSynthetic = true; break;
        case "--torat-emet-sample": toratEmetSample = int.Parse(args[++i]); break;
        case "--torat-emet-stats": toratEmetFull = true; break;
        case "--text-corpus": textCorpus = args[++i]; break;
        case "--text-largest": textLargest = int.Parse(args[++i]); break;
        case "--text-random": textRandom = int.Parse(args[++i]); break;
        case "--mime-scan": mimeScan = args[++i]; break;
        case "--profile-markup": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) markupProfilePaths.Add(args[++i]); break;
        case "--rtf":
            rtfBlocks = 20_000;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int rtfCount)) { rtfBlocks = rtfCount; i++; }
            break;
        case "--rtf-profile": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) rtfProfilePaths.Add(args[++i]); break;
        case "--rtf-audit": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) rtfAuditPaths.Add(args[++i]); break;
        case "--odt-drawing": OdtDrawingFidelityTest.Run(); return TestReport.FailureCount == 0 ? 0 : 1;
        case "--odt-shape-text": OdtShapeTextFidelityTest.Run(); return TestReport.FailureCount == 0 ? 0 : 1;
        case "--dump-html":                                                            // <file>: every block's HTML, one per line (inspection)
        {
            string dumpPath = args[++i];
            using FastDocView.IWordWindowedDocument dumped = FastDocView.WordWindowedDocumentOpener.Open(dumpPath, new FastDocView.WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache") });
            foreach (var block in dumped.RenderBlocks(0, dumped.BlockCount)) Console.WriteLine(block.Html);
            return TestReport.FailureCount == 0 ? 0 : 1;
        }
        case "--golden": goldenDirectory = args[++i]; while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) goldenPaths.Add(args[++i]); break;   // <out-dir> <files...>: observable-output dump per file (index-change proof)
        case "--doc-search-golden": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) docSearchGoldenPaths.Add(args[++i]); break;   // .doc one-pass search = the old block-range path, on these files only
        case "--coexistence-probe": WordCoexistenceTest.Probe(args[++i], int.Parse(args[++i])); return TestReport.FailureCount == 0 ? 0 : 1;   // <file> <rounds>: save-by-replace refusals with / without our document open, and who held the file
        case "--read-counts": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) readCountPaths.Add(args[++i]); break;   // .rtf / .odt / .fodt: file opens, bytes read, stats per window, notes page, search
        case "--rtf-search-golden": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) rtfSearchGoldenPaths.Add(args[++i]); break;   // .rtf one-pass search = the old per-block reparse
        case "--reread": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) rereadPaths.Add(args[++i]); break;   // re-shown windows / notes: checkpoint resumes, bytes inflated
        case "--doc-reads": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) docReadPaths.Add(args[++i]); break;   // .doc reopen / full-search bytes read, read calls, ms
        case "--compact-report": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) compactReportPaths.Add(args[++i]); break;   // sidecar size breakdown + reopen / window / random-block timings
        case "--corpus-max": corpusMax = int.Parse(args[++i]); break;
        case "--corpus-sample": corpusSample = int.Parse(args[++i]); break;   // a fixed-seed random sample of N corpus files (quick regression pass)
        case "--word-compare-page": WordComparisonPage.Run(args[++i], args[++i], args[++i]); return TestReport.FailureCount == 0 ? 0 : 1;   // <document> <word.json> <out.html>: blocks matching Word's sampled paragraphs
        case "--list-labels-fallback": listLabelsFallback = args[++i]; break;   // a path list to scan for a .doc with lists when the --real ones have none
    }
}

if (goldenDirectory is not null) { GoldenDumpTest.Run(goldenDirectory, goldenPaths.Where(File.Exists)); skipSynthetic = true; }
if (docSearchGoldenPaths.Count > 0) { DocBinarySearchOnePassTest.CompareFiles(docSearchGoldenPaths.Where(File.Exists)); skipSynthetic = true; }
if (readCountPaths.Count > 0) { FileReadCountTest.Run(readCountPaths.Where(File.Exists)); skipSynthetic = true; }
if (rtfSearchGoldenPaths.Count > 0) { RtfSearchOnePassTest.Run(rtfSearchGoldenPaths.Where(File.Exists)); skipSynthetic = true; }
if (rereadPaths.Count > 0) { RereadCostTest.Run(rereadPaths.Where(File.Exists)); skipSynthetic = true; }
if (docReadPaths.Count > 0) { DocBinaryReadCountTest.Run(docReadPaths.Where(File.Exists)); skipSynthetic = true; }
if (compactReportPaths.Count > 0) { CompactIndexReport.Run(compactReportPaths.Where(File.Exists)); skipSynthetic = true; }
if (searchOnly) { DocumentSearchTest.Run(searchPaths); skipSynthetic = true; }
if (fidelityOnly) { CompactIndexTest.RunEliasFano(); DocxFidelityTest.Run(); DocxThemeAndUnderlineTest.Run(); DocBinaryBidiTest.Run(); DocBinaryListLayoutTest.Run(); ListLabelFormatsTest.Run(); PagePlacementFidelityTest.Run(); LineSpacingTest.Run(); LinksAllFormatsTest.Run(); DocBinaryLinksTest.Run(); WordFieldsFidelityTest.Run(); WordFieldsOtherFormatsTest.Run(); FormattingMarksFidelityTest.Run(); DocBinaryFloatingPicturesTest.Run(); OdtShapesFidelityTest.Run(); OdtChartsFidelityTest.Run(); OdtDrawingFidelityTest.Run(); RtfShapesFidelityTest.Run(); WrapSidesAndContourFidelityTest.Run(); WordOfficeArtGeometryTest.Run(); DocBinaryShapesFidelityTest.Run(); OdtFirstOpenTiming.RunSynthetic(); OdtConnectorRouteTest.Run(); OdtConnectedShapesTest.Run(); OdtMeasureCaptionTest.Run(); OdtExtrusionTest.Run(); OdtShapeTextFidelityTest.Run(); TextEffectsFidelityTest.Run(); DropCapsAndHyphenationFidelityTest.Run(); ListNumberFormatsFidelityTest.Run(); ListBulletsFidelityTest.Run(); BodyScriptFidelityTest.Run(); WordTablesFidelityTest.Run(); WebArchivePicturesTest.Run(); OfficeMathFidelityTest.Run(); skipSynthetic = true; }
if (pageWidthOnly) { PageWidthTest.Run(realPaths); skipSynthetic = true; realPaths.Clear(); }
if (!skipSynthetic)
{
    DocxNumberingAndPolicyTest.Run();
    WordDocumentIndexCacheTest.Run();
    DocxThemeAndUnderlineTest.Run();
    DocxFidelityTest.Run();
    TextEffectsFidelityTest.Run();
    DropCapsAndHyphenationFidelityTest.Run();
    DocumentSearchTest.Run([]);
    PageWidthTest.Run([]);
    DocBinaryBidiTest.Run();
    DocBinaryListLayoutTest.Run();
    ListLabelFormatsTest.Run();
    LineSpacingTest.Run();
    MalformedXmlRecoveryTest.Run();
    CompactIndexTest.RunEliasFano();
    DocxIndexAndRenderTest.Run(blocks);
}
if (hugeBlocks > 0) HugeDocumentMemoryTest.Run(hugeBlocks);
if (incrementalOpen)
    DocxIncrementalOpenTest.Run(incrementalBlocks, realPaths.Where(path => File.Exists(path) && path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)));
if (odtIncrementalBlocks > 0) OdtIncrementalOpenTest.Run(odtIncrementalBlocks);
if (streamingIncrementalBlocks > 0) StreamingFormatsIncrementalOpenTest.Run(streamingIncrementalBlocks);
if (notesStreamingBlocks > 0) DocxNotesStreamingTest.Run(notesStreamingBlocks, noteFillerCharacters: 60_000);
if (realPaths.Count > 0) RealDocumentSmokeTest.Run(realPaths.Where(File.Exists));
if (realPaths.Count > 0) CompactIndexTest.RunRecovery(realPaths.Where(path => File.Exists(path) && path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)));
if (realPaths.Count > 0) CompactDocIndexTest.Run(realPaths.Where(path => File.Exists(path) && path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)));
if (fidelityOnly || realPaths.Count > 0) NotesWindowTest.Run(realPaths.Where(File.Exists));
if (realPaths.Count > 0) WordCoexistenceTest.Run(realPaths.Where(File.Exists));
if (realPaths.Any(path => path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)))
    DocBinaryListLabelsTest.Run(realPaths.Where(path => File.Exists(path) && path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)), listLabelsFallback ?? corpusList);
if (auditPaths.Count > 0) WordFormattingAudit.Run(auditPaths.Where(File.Exists));
if (odtFirstOpenPaths.Count > 0) OdtFirstOpenTiming.Run(odtFirstOpenPaths.Where(File.Exists));
if (profilePaths.Count > 0) OpenTimingProfile.Run(profilePaths.Where(File.Exists));
if (indexProfilePaths.Count > 0) IndexPassProfile.Run(indexProfilePaths.Where(File.Exists));
if (corpusList is not null && rtfBlocks == 0) WordFormattingAudit.RunCorpus(corpusList, corpusMax, corpusSample);
if (corpusList is not null && Directory.Exists(corpusList) && File.Exists(Path.Combine(corpusList, "apache-poi", "MultipleBodyBug.docx")))
{
    CorpusRegressionTest.Run(corpusList);
    CorpusRegressionTest.RunOpenFailureClasses(corpusList);
    CorpusRegressionTest.RunMalformedBlock(corpusList);
}
if (rtfBlocks > 0) RtfIndexAndRenderTest.Run(rtfBlocks);
foreach (string path in rtfProfilePaths.Where(File.Exists)) RtfIndexAndRenderTest.Profile(path);
if (rtfAuditPaths.Count > 0) RtfFormattingAudit.Run(rtfAuditPaths.Where(File.Exists));
if (rtfBlocks > 0 && corpusList is not null) RtfFormattingAudit.RunCorpus(corpusList, corpusMax);
if (odtBlocks > 0)
{
    OdtIndexAndRenderTest.Run(odtBlocks);
    List<string> converted = LibreOfficeConversion.ConvertToOdt(realPaths.Where(path => File.Exists(path) && path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)));
    if (converted.Count > 0)
    {
        RealDocumentSmokeTest.Run(converted);
        OdtFormattingAudit.Run(converted);
    }
}

if (textBlocks > 0) { MarkupIndexAndRenderTest.RunEdgeCases(); MarkupIndexAndRenderTest.RunDoubleByteSplit(); MarkupIndexAndRenderTest.RunText(textBlocks); ToratEmetTest.Run(); ToratEmetPathRuleTest.Run(); PlainTextNormaliseTest.Run(); }
if (mhtmlBlocks > 0) { MarkupIndexAndRenderTest.RunNestedMultipart(); MarkupIndexAndRenderTest.RunMhtml(mhtmlBlocks); }
if (markupProfilePaths.Count > 0)
{
    TestReport.Step("Markup open timings");
    foreach (string path in markupProfilePaths.Where(File.Exists)) MarkupIndexAndRenderTest.Profile(path);
}
if (toratEmetFolder is not null) ToratEmetTest.RunReal(toratEmetFolder, toratEmetSample, toratEmetFull);
if (textCorpus is not null) MarkupCorpusTest.RunText(textCorpus, textLargest, textRandom);
if (mimeScan is not null) MarkupCorpusTest.RunMimeScan(mimeScan);

Console.WriteLine();
Console.WriteLine(TestReport.FailureCount == 0 ? "ALL PASSED" : $"{TestReport.FailureCount} FAILURE(S)");
return TestReport.FailureCount == 0 ? 0 : 1;
