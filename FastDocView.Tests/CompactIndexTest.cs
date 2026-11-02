using FastDocView.Docx;

namespace FastDocView.Tests;

/// <summary>
/// The compact .docx index: (1) <see cref="EliasFanoSequence"/> round-trips random non-decreasing sequences
/// and answers select / lower bound exactly; (2) block-position recovery — for random blocks of a real
/// .docx the compact table's start offset, kind, text length, style and list label (rescanned from the
/// span anchor) equal a reference pass that scans the whole body once with the same scanner.
/// </summary>
public static class CompactIndexTest
{
    public static void RunEliasFano()
    {
        TestReport.Step("Elias-Fano sequences");
        var random = new Random(3);
        bool allOk = true;
        foreach (int count in new[] { 0, 1, 2, 63, 64, 65, 1000, 50_000 })
        {
            var values = new int[count];
            int running = 0;
            for (int i = 0; i < count; i++) { running += random.Next(4) == 0 ? 0 : random.Next(1, 300); values[i] = running; }   // repeats allowed
            var sequence = new EliasFanoSequence(values);
            byte[] bytes;
            using (var buffer = new MemoryStream()) { using (var writer = new BinaryWriter(buffer, System.Text.Encoding.UTF8, leaveOpen: true)) sequence.Write(writer); bytes = buffer.ToArray(); }
            EliasFanoSequence read;
            using (var reader = new BinaryReader(new MemoryStream(bytes))) read = EliasFanoSequence.Read(reader);
            bool ok = read.Count == count && read.ToArray().SequenceEqual(values);
            for (int i = 0; i < Math.Min(count, 500) && ok; i++)
            {
                int probe = random.Next(count);
                if (read.Get(probe) != values[probe]) ok = false;
                int target = random.Next(values[count - 1] + 2);
                int expected = 0; while (expected < count && values[expected] < target) expected++;
                if (read.LowerBound(target) != expected) ok = false;
            }
            if (!ok) { allOk = false; Console.WriteLine($"   FAIL at count {count}"); }
            if (count == 50_000) Console.WriteLine($"   {count:N0} values up to {values[count - 1]:N0}: {bytes.Length:N0} bytes ({bytes.Length * 8.0 / count:F2} bits/value)");
        }
        TestReport.Check(allOk, "Elias-Fano: round trip, select and lower bound exact for every size");
    }

    /// <summary>Random blocks of each .docx: the compact table's recovered values equal one reference scan of the whole body.</summary>
    public static void RunRecovery(IEnumerable<string> docxPaths)
    {
        foreach (string path in docxPaths)
        {
            TestReport.Step($"Block-position recovery: {Path.GetFileName(path)}");
            string cacheDirectory = Path.Combine(TestReport.OutputDirectory, "index-cache-compact");
            var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = WordIndexMode.Sidecar };
            using var document = DocxWindowedDocument.Open(path, policy);
            DocxRandomAccessIndex index = document.Index;
            Console.WriteLine($"   {index.BlockCount:N0} blocks, {index.Checkpoints.Count} checkpoints, {index.Anchors.Length} anchors, sidecar {new FileInfo(policy.IndexPathFor(path)).Length:N0} bytes");

            // reference: one pass with the same scanner over the whole body
            using FileStream package = WordDocumentFileStream.OpenShared(path);
            var parts = new DocxDocumentParts(package);
            var reference = new DocxBodyBlockScanner(parts.CreateListLabelResolver(), parts.Styles.HeadingLevelOf);
            var cursor = new DocxEntryInflationCursor(package, index.DocumentEntry, long.MaxValue);
            TestReport.Timed("reference full scan", () => { int read; while ((read = cursor.ReadNextChunk()) > 0) reference.Feed(cursor.Chunk, read); reference.Finish(); });
            TestReport.Check(reference.BlockStartOffsets.Count == index.BlockCount, "block count matches the reference scan");

            Stream view = DocxEntryInflation.OpenView(package, index.DocumentEntry, index.Checkpoints);
            var table = new DocxCompactBlockTable(index, view, parts.CreateListLabeller);
            var random = new Random(17);
            int mismatches = 0;
            var probes = new List<int>();
            for (int i = 0; i < 300; i++) probes.Add(random.Next(index.BlockCount));
            foreach (DocxIndexAnchor anchor in index.Anchors) { probes.Add(anchor.FirstBlock); if (anchor.FirstBlock > 0) probes.Add(anchor.FirstBlock - 1); }   // span edges
            probes.Add(index.BlockCount - 1);
            var times = new List<double>();
            foreach (int block in probes)
            {
                var watch = System.Diagnostics.Stopwatch.StartNew();
                long start = table.BlockStartOffset(block);
                times.Add(watch.Elapsed.TotalMilliseconds);
                long end = table.BlockEndOffset(block);
                long referenceEnd = block + 1 < reference.BlockStartOffsets.Count ? reference.BlockStartOffsets[block + 1] : reference.BodyEndOffset;
                short styleIndex = reference.BlockStyleIds[block];
                string? referenceStyle = styleIndex >= 0 ? reference.StyleIdTable[styleIndex] : null;
                int labelIndex = reference.BlockListLabelIds[block];
                DocxListLabel? referenceLabel = labelIndex >= 0 ? reference.ListLabels[labelIndex] : null;
                bool ok = start == reference.BlockStartOffsets[block] && end == referenceEnd
                       && table.BlockKindOf(block) == reference.BlockKinds[block]
                       && table.BlockTextLengthOf(block) == reference.BlockTextLengths[block]
                       && (table.StyleIdOf(block) ?? "") == (referenceStyle ?? "")
                       && Equals(table.ListLabelOf(block), referenceLabel);
                if (!ok && mismatches++ < 5) Console.WriteLine($"   mismatch at block {block}: start {start} vs {reference.BlockStartOffsets[block]}, label {table.ListLabelOf(block)?.Text} vs {referenceLabel?.Text}");
            }
            times.Sort();
            Console.WriteLine($"   {probes.Count} probes: locate median {times[times.Count / 2]:F1} ms, p90 {times[times.Count * 9 / 10]:F1} ms");
            TestReport.Check(mismatches == 0, $"{probes.Count} random + span-edge blocks: position, kind, text length, style and list label equal the reference scan");
            int labelled = Enumerable.Range(0, reference.BlockListLabelIds.Count).Count(b => reference.BlockListLabelIds[b] >= 0);
            Console.WriteLine($"   reference: {labelled:N0} list paragraphs, {reference.ListLabels.Count} distinct labels");
        }
    }
}
