// WordDocumentRandomAccess test harness.
//
//   dotnet run -c Release                          deflate round trip (8 MB) + synthetic docx (50k blocks) + real files passed/found
//   dotnet run -c Release -- --blocks 500000       bigger synthetic document
//   dotnet run -c Release -- --deflate-mb 64       bigger deflate round trip
//   dotnet run -c Release -- --real <path>...      real .docx/.doc files to smoke-test (repeatable)
//   dotnet run -c Release -- --skip-synthetic      only real files
//   dotnet run -c Release -- --huge 1000000        instant reload + flat heap on a ~350 MB body
//   dotnet run -c Release -- --incremental-open [N]  first open with the index built on demand (synthetic N blocks, default 1M, + --real .docx files)
//   dotnet run -c Release -- --notes-streaming 20000   notes read on demand from a 100 MB+ footnotes part
//   dotnet run -c Release -- --audit <path>...     formatting constructs the renderers do not handle
//   dotnet run -c Release -- --odt [N]             synthetic .odt/.fodt (N repeating blocks, default 20000); with LibreOffice installed,
//                                                  every --real .docx is also converted to .odt (TestOutput) and smoke-tested + audited
//   dotnet run -c Release -- --text [N] / --mhtml [N]   synthetic .txt (every encoding / line ending) and MHTML (qp + base64), N filler blocks
//   dotnet run -c Release -- --text-corpus <list> [--text-largest 200 --text-random 300]   .txt corpus sample: encoding mix, 0 crashes
//   dotnet run -c Release -- --mime-scan <word-list>   Word files that start with 'MIME' (web archives): count, open + render each
//   dotnet run -c Release -- --profile-markup <path>...  cold / reopen / window ms of .txt / .mht files
//   dotnet run -c Release -- --rtf [N]             synthetic .rtf (N repeating blocks, default 20000); with --corpus, the RTF files of the list
//                                                  are audited (unhandled control words) and opened + rendered; --rtf-audit <path>... audits files
//   dotnet run -c Release -- --skip-synthetic --corpus <list.txt>   the same over a whole file list, aggregated + open/render crash check
using WordDocumentRandomAccess.Tests;

int blocks = 50_000, deflateMegabytes = 8, hugeBlocks = 0, notesStreamingBlocks = 0;
bool skipSynthetic = false, incrementalOpen = false;
int incrementalBlocks = 1_000_000;
var realPaths = new List<string>();
var auditPaths = new List<string>();
var profilePaths = new List<string>();
var indexProfilePaths = new List<string>();
string? corpusList = null;
string? listLabelsFallback = null;
int corpusMax = int.MaxValue;
int odtBlocks = 0, rtfBlocks = 0, textBlocks = 0, mhtmlBlocks = 0, textLargest = 200, textRandom = 300;
string? textCorpus = null, mimeScan = null;
var markupProfilePaths = new List<string>();
var rtfAuditPaths = new List<string>();
var rtfProfilePaths = new List<string>();
for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--blocks": blocks = int.Parse(args[++i]); break;
        case "--huge": hugeBlocks = int.Parse(args[++i]); break;                       // e.g. 1000000 → ~350 MB document.xml
        case "--notes-streaming": notesStreamingBlocks = int.Parse(args[++i]); break;  // e.g. 20000 → 1000 footnotes, ~113 MB footnotes.xml
        case "--deflate-mb": deflateMegabytes = int.Parse(args[++i]); break;
        case "--skip-synthetic": skipSynthetic = true; break;
        case "--incremental-open":                                                     // optional block count, default 1000000
            incrementalOpen = true;
            if (i + 1 < args.Length && int.TryParse(args[i + 1], out int parsed)) { incrementalBlocks = parsed; i++; }
            break;
        case "--real": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) realPaths.Add(args[++i]); break;
        case "--audit": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) auditPaths.Add(args[++i]); break;
        case "--profile-open": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) profilePaths.Add(args[++i]); break;
        case "--profile-index": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) indexProfilePaths.Add(args[++i]); break;
        case "--corpus": corpusList = args[++i]; break;            // a text file with one Word file path per line
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
        case "--corpus-max": corpusMax = int.Parse(args[++i]); break;
        case "--list-labels-fallback": listLabelsFallback = args[++i]; break;   // a path list to scan for a .doc with lists when the --real ones have none
    }
}

if (!skipSynthetic)
{
    DocxNumberingAndPolicyTest.Run();
    DocxThemeAndUnderlineTest.Run();
    DocxIndexAndRenderTest.Run(blocks);
}
if (hugeBlocks > 0) HugeDocumentMemoryTest.Run(hugeBlocks);
if (incrementalOpen)
    DocxIncrementalOpenTest.Run(incrementalBlocks, realPaths.Where(path => File.Exists(path) && path.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)));
if (notesStreamingBlocks > 0) DocxNotesStreamingTest.Run(notesStreamingBlocks, noteFillerCharacters: 60_000);
if (realPaths.Count > 0) RealDocumentSmokeTest.Run(realPaths.Where(File.Exists));
if (realPaths.Count > 0) WordCoexistenceTest.Run(realPaths.Where(File.Exists));
if (realPaths.Any(path => path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)))
    DocBinaryListLabelsTest.Run(realPaths.Where(path => File.Exists(path) && path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)), listLabelsFallback ?? corpusList);
if (auditPaths.Count > 0) WordFormattingAudit.Run(auditPaths.Where(File.Exists));
if (profilePaths.Count > 0) OpenTimingProfile.Run(profilePaths.Where(File.Exists));
if (indexProfilePaths.Count > 0) IndexPassProfile.Run(indexProfilePaths.Where(File.Exists));
if (corpusList is not null && rtfBlocks == 0) WordFormattingAudit.RunCorpus(corpusList, corpusMax);
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

if (textBlocks > 0) { MarkupIndexAndRenderTest.RunEdgeCases(); MarkupIndexAndRenderTest.RunText(textBlocks); }
if (mhtmlBlocks > 0) MarkupIndexAndRenderTest.RunMhtml(mhtmlBlocks);
if (markupProfilePaths.Count > 0)
{
    TestReport.Step("Markup open timings");
    foreach (string path in markupProfilePaths.Where(File.Exists)) MarkupIndexAndRenderTest.Profile(path);
}
if (textCorpus is not null) MarkupCorpusTest.RunText(textCorpus, textLargest, textRandom);
if (mimeScan is not null) MarkupCorpusTest.RunMimeScan(mimeScan);

Console.WriteLine();
Console.WriteLine(TestReport.FailureCount == 0 ? "ALL PASSED" : $"{TestReport.FailureCount} FAILURE(S)");
return TestReport.FailureCount == 0 ? 0 : 1;
