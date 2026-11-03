using System.Diagnostics;
using FastDocView.DocBinary;

namespace FastDocView.Tests;

/// <summary>
/// How much of a .doc a reopen from the sidecar and a full search read: the file is opened through a counting,
/// unbuffered stream, so the numbers are the bytes and read calls the reader asks the file for (not the OS cache's).
/// </summary>
public static class DocBinaryReadCountTest
{
    internal sealed class CountingStream(Stream inner) : Stream
    {
        public long BytesRead, ReadCalls;
        public void ResetCounts() { BytesRead = 0; ReadCalls = 0; }
        public override int Read(byte[] buffer, int offset, int count)
        {
            int read = inner.Read(buffer, offset, count);
            ReadCalls++;
            BytesRead += read;
            return read;
        }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => inner.Length;
        public override long Position { get => inner.Position; set => inner.Position = value; }
        public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }
    }

    /// <summary>The process's CPU time: the machine may be busy (the harness runs at idle priority), so wall time alone misleads.</summary>
    private static TimeSpan CpuTime() => Process.GetCurrentProcess().TotalProcessorTime;

    public static void Run(IEnumerable<string> paths)
    {
        foreach (string path in paths)
        {
            TestReport.Step($".doc read counts: {Path.GetFileName(path)}");
            var policy = new WordIndexPolicy { CacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-doc-reads"), ForcedMode = WordIndexMode.Sidecar };
            string indexPath = policy.IndexPathFor(path);
            if (File.Exists(indexPath)) File.Delete(indexPath);
            using (DocBinaryWindowedDocument.Open(path, policy)) { }
            var info = new FileInfo(path);
            for (int round = 0; round < 3; round++)
            {
                var watch = Stopwatch.StartNew();
                TimeSpan cpu = CpuTime();
                DocBinaryBlockIndex cached = DocBinaryBlockIndex.TryLoad(indexPath, path) ?? throw new InvalidOperationException("no sidecar");
                double loadMs = watch.Elapsed.TotalMilliseconds;
                var file = new CountingStream(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1));
                using DocBinaryWindowedDocument document = DocBinaryWindowedDocument.OpenFromStream(file, path, policy, cached, info.Length, info.LastWriteTimeUtc.Ticks);
                double reopenMs = watch.Elapsed.TotalMilliseconds, reopenCpuMs = (CpuTime() - cpu).TotalMilliseconds;
                Console.WriteLine($"   reopen #{round + 1}: {reopenMs:0.#} ms (cpu {reopenCpuMs:0} ms; sidecar load {loadMs:0.#} ms), read {file.BytesRead:N0} of {info.Length:N0} bytes in {file.ReadCalls:N0} calls");
                if (round < 2) continue;
                foreach (string query in new[] { "qqzzqq", "\u05E9\u05D1\u05EA" })
                {
                    file.ResetCounts();
                    watch.Restart();
                    cpu = CpuTime();
                    int from = 0, hits = 0, matches = 0;
                    while (true)
                    {
                        WordDocumentSearchBatch batch = document.Search(query, WordDocumentSearchOptions.Default, from, 500);
                        hits += batch.HitBlocks.Length;
                        matches += batch.HitCounts.Sum();
                        if (batch.ReachedEnd) break;
                        from = batch.NextBlock;
                    }
                    Console.WriteLine($"   full search '{query}': {watch.ElapsedMilliseconds} ms (cpu {(CpuTime() - cpu).TotalMilliseconds:0} ms), {hits} hit blocks ({matches} matches), read {file.BytesRead:N0} bytes in {file.ReadCalls:N0} calls");
                }
            }
        }
    }
}
