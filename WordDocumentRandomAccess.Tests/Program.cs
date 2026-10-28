// WordDocumentRandomAccess test harness.
//
//   dotnet run -c Release                          deflate round trip (8 MB) + synthetic docx (50k blocks) + real files passed/found
//   dotnet run -c Release -- --blocks 500000       bigger synthetic document
//   dotnet run -c Release -- --deflate-mb 64       bigger deflate round trip
//   dotnet run -c Release -- --real <path>...      real .docx/.doc files to smoke-test (repeatable)
//   dotnet run -c Release -- --skip-synthetic      only real files
//   dotnet run -c Release -- --huge 1000000        instant reload + flat heap on a ~350 MB body
//   dotnet run -c Release -- --notes-streaming 20000   notes read on demand from a 100 MB+ footnotes part
//   dotnet run -c Release -- --audit <path>...     formatting constructs the renderers do not handle
//   dotnet run -c Release -- --skip-synthetic --corpus <list.txt>   the same over a whole file list, aggregated + open/render crash check
using WordDocumentRandomAccess.Tests;

int blocks = 50_000, deflateMegabytes = 8, hugeBlocks = 0, notesStreamingBlocks = 0;
bool skipSynthetic = false;
var realPaths = new List<string>();
var auditPaths = new List<string>();
var profilePaths = new List<string>();
string? corpusList = null;
string? listLabelsFallback = null;
int corpusMax = int.MaxValue;
for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--blocks": blocks = int.Parse(args[++i]); break;
        case "--huge": hugeBlocks = int.Parse(args[++i]); break;                       // e.g. 1000000 → ~350 MB document.xml
        case "--notes-streaming": notesStreamingBlocks = int.Parse(args[++i]); break;  // e.g. 20000 → 1000 footnotes, ~113 MB footnotes.xml
        case "--deflate-mb": deflateMegabytes = int.Parse(args[++i]); break;
        case "--skip-synthetic": skipSynthetic = true; break;
        case "--real": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) realPaths.Add(args[++i]); break;
        case "--audit": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) auditPaths.Add(args[++i]); break;
        case "--profile-open": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) profilePaths.Add(args[++i]); break;
        case "--corpus": corpusList = args[++i]; break;            // a text file with one Word file path per line
        case "--corpus-max": corpusMax = int.Parse(args[++i]); break;
        case "--list-labels-fallback": listLabelsFallback = args[++i]; break;   // a path list to scan for a .doc with lists when the --real ones have none
    }
}

if (!skipSynthetic)
{
    DeflateInflaterRoundTripTest.Run(deflateMegabytes);
    DocxNumberingAndPolicyTest.Run();
    DocxIndexAndRenderTest.Run(blocks);
}
if (hugeBlocks > 0) HugeDocumentMemoryTest.Run(hugeBlocks);
if (notesStreamingBlocks > 0) DocxNotesStreamingTest.Run(notesStreamingBlocks, noteFillerCharacters: 60_000);
if (realPaths.Count > 0) RealDocumentSmokeTest.Run(realPaths.Where(File.Exists));
if (realPaths.Any(path => path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)))
    DocBinaryListLabelsTest.Run(realPaths.Where(path => File.Exists(path) && path.EndsWith(".doc", StringComparison.OrdinalIgnoreCase)), listLabelsFallback ?? corpusList);
if (auditPaths.Count > 0) WordFormattingAudit.Run(auditPaths.Where(File.Exists));
if (profilePaths.Count > 0) OpenTimingProfile.Run(profilePaths.Where(File.Exists));
if (corpusList is not null) WordFormattingAudit.RunCorpus(corpusList, corpusMax);

Console.WriteLine();
Console.WriteLine(TestReport.FailureCount == 0 ? "ALL PASSED" : $"{TestReport.FailureCount} FAILURE(S)");
return TestReport.FailureCount == 0 ? 0 : 1;
