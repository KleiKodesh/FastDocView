// WordDocumentRandomAccess test harness.
//
//   dotnet run -c Release                          deflate round trip (8 MB) + synthetic docx (50k blocks) + real files passed/found
//   dotnet run -c Release -- --blocks 500000       bigger synthetic document
//   dotnet run -c Release -- --deflate-mb 64       bigger deflate round trip
//   dotnet run -c Release -- --real <path>...      real .docx/.doc files to smoke-test (repeatable)
//   dotnet run -c Release -- --skip-synthetic      only real files
using WordDocumentRandomAccess.Tests;

int blocks = 50_000, deflateMegabytes = 8, hugeBlocks = 0;
bool skipSynthetic = false;
var realPaths = new List<string>();
for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--blocks": blocks = int.Parse(args[++i]); break;
        case "--huge": hugeBlocks = int.Parse(args[++i]); break;       // e.g. 1000000 → ~350 MB document.xml
        case "--deflate-mb": deflateMegabytes = int.Parse(args[++i]); break;
        case "--skip-synthetic": skipSynthetic = true; break;
        case "--real": while (i + 1 < args.Length && !args[i + 1].StartsWith("--")) realPaths.Add(args[++i]); break;
    }
}

if (!skipSynthetic)
{
    DeflateInflaterRoundTripTest.Run(deflateMegabytes);
    DocxNumberingAndPolicyTest.Run();
    DocxIndexAndRenderTest.Run(blocks);
}
if (hugeBlocks > 0) HugeDocumentMemoryTest.Run(hugeBlocks);
if (realPaths.Count > 0) RealDocumentSmokeTest.Run(realPaths.Where(File.Exists));

Console.WriteLine();
Console.WriteLine(TestReport.FailureCount == 0 ? "ALL PASSED" : $"{TestReport.FailureCount} FAILURE(S)");
return TestReport.FailureCount == 0 ? 0 : 1;
