using System.Diagnostics;
using SeekableDeflate;
using WordDocumentRandomAccess.Docx;
using SeekableDeflate;

namespace WordDocumentRandomAccess.Tests;

/// <summary>
/// Where the first-open index pass of a .docx spends its time: inflating document.xml alone, the
/// same with the body scanner fed, .NET's DeflateStream for comparison, and the full build
/// (body + notes + parts). The numbers the inflater / scanner speed work is judged by.
/// </summary>
public static class IndexPassProfile
{
    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($"Index pass profile: {Path.GetFileName(path)}");
            using var package = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16);
            ZipEntryLocation entry = DocxRandomAccessIndexBuilder.LocateDocumentEntry(package);
            Console.WriteLine($"   document.xml: {entry.CompressedSize >> 20} MB compressed, {entry.UncompressedSize >> 20} MB uncompressed");
            var chunk = new byte[1 << 18];

            double inflateOnly = Time(() =>
            {
                var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
                while (inflater.Read(chunk, 0, chunk.Length) > 0) { }
            });
            Report("our inflater, no scanner", inflateOnly, entry.UncompressedSize);

            double inflateWithScanner = Time(() =>
            {
                var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
                var scanner = new DocxBodyBlockScanner();
                int read;
                while ((read = inflater.Read(chunk, 0, chunk.Length)) > 0) scanner.Feed(chunk, read);
            });
            Report("our inflater + body scanner", inflateWithScanner, entry.UncompressedSize);

            double dotnetInflate = Time(() =>
            {
                package.Position = entry.DataOffset;
                using var deflate = new System.IO.Compression.DeflateStream(package, System.IO.Compression.CompressionMode.Decompress, leaveOpen: true);
                while (deflate.Read(chunk, 0, chunk.Length) > 0) { }
            });
            Report(".NET DeflateStream (zlib), no scanner", dotnetInflate, entry.UncompressedSize);

            // scanner alone, over the inflated bytes (held once, test only): isolates its cost from inflate
            byte[] inflated;
            using (var buffer = new MemoryStream((int)Math.Min(int.MaxValue - 64, entry.UncompressedSize)))
            {
                var inflater = new DeflateCheckpointInflater(package, entry.DataOffset, entry.CompressedSize);
                int read;
                while ((read = inflater.Read(chunk, 0, chunk.Length)) > 0) buffer.Write(chunk, 0, read);
                inflated = buffer.GetBuffer();
            }
            int inflatedLength = (int)entry.UncompressedSize;
            for (int run = 1; run <= 2; run++)
            {
                double scanOnly = Time(() =>
                {
                    var scanner = new DocxBodyBlockScanner();
                    for (int at = 0; at < inflatedLength; at += chunk.Length)
                    {
                        int length = Math.Min(chunk.Length, inflatedLength - at);
                        Buffer.BlockCopy(inflated, at, chunk, 0, length);
                        scanner.Feed(chunk, length);
                    }
                });
                Report($"body scanner alone (run {run})", scanOnly, entry.UncompressedSize);
            }
            inflated = [];

            var info = new FileInfo(path);
            double fullBuild = Time(() =>
            {
                package.Position = 0;
                var parts = new DocxDocumentParts(package);
                DocxRandomAccessIndexBuilder.Build(package, parts, info.Length, info.LastWriteTimeUtc.Ticks);
            });
            Console.WriteLine($"   full index build (parts + body + notes + labels): {fullBuild:0} ms");
        }
    }

    private static double Time(Action action)
    {
        var watch = Stopwatch.StartNew();
        action();
        return watch.Elapsed.TotalMilliseconds;
    }

    private static void Report(string what, double milliseconds, long bytes)
        => Console.WriteLine($"   {what}: {milliseconds:0} ms = {bytes / 1048576.0 / (milliseconds / 1000):0} MB/s");
}
