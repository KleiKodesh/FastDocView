// Localhost JSON host for the Vite demo. Endpoints (all JSON, CORS open — dev only):
//   GET  /api/folder?path=…                  policy decision for every Word file under a folder
//   POST /api/documents  { path, forceSidecar? }   opens (indexing if needed) → document summary
//   GET  /api/documents/{id}/weights          estimated text length per block, raw little-endian Int32 (scrollbar weights)
//   GET  /api/documents/{id}/headings         [{ blockIndex, level, text }] — from the index, no block reads
//   GET  /api/documents/{id}/blocks?first&count    rendered blocks [{ index, kind, html, notes: [{ kind, number, label, html }] }];
//        headers X-Known-Block-Count / X-Index-Complete: a .docx first open builds its index
//        on demand (each request scans as far as it needs), so the reader corrects its count from every response.
//        Headings / weights / note-references return what is known so far.
//   DELETE /api/documents/{id}                closes
//   GET  /api/progress                        the last index progress (polled while opening)
//   GET  /api/memory                          managed heap / working set of this process
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.Json.Serialization;
using FastDocView;
using FastDocView.DocBinary;
using FastDocView.Docx;

// FASTDOCVIEW_DEV_URL lets a second rig run beside the one in use (e.g. a worktree's checks on another port).
string Url = Environment.GetEnvironmentVariable("FASTDOCVIEW_DEV_URL") ?? "http://127.0.0.1:5188";

var builder = WebApplication.CreateSlimBuilder(args);
builder.WebHost.UseUrls(Url);
builder.Services.AddCors(options => options.AddDefaultPolicy(policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod().WithExposedHeaders("X-Render-Milliseconds", "X-Known-Block-Count", "X-Index-Complete")));
builder.Services.ConfigureHttpJsonOptions(options => options.SerializerOptions.TypeInfoResolverChain.Insert(0, DevServerJsonContext.Default));
builder.Logging.SetMinimumLevel(LogLevel.Warning);

WebApplication app = builder.Build();
app.UseCors();

// Relative paths ("samples/…") resolve against the repository root, so the demo's sample buttons
// carry no machine-specific folders.
string repositoryRoot = Path.GetFullPath(Path.Combine(app.Environment.ContentRootPath, ".."));
string ResolvePath(string path) => Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(repositoryRoot, path));

var documents = new ConcurrentDictionary<int, IWordWindowedDocument>();
var cacheDirectory = Path.Combine(AppContext.BaseDirectory, "index-cache");
var lastProgress = new WordDocumentIndexProgress(0, 0, 0);
int nextId = 0;

app.MapGet("/api/folder", (string path) =>
{
    path = ResolvePath(path);
    if (!Directory.Exists(path)) return Results.NotFound(new ErrorBody($"Folder not found: {path}"));
    var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory };
    List<FolderEntry> entries = WordWindowedDocumentOpener.DecideFolder(path, policy, recursive: true, maxFiles: 500)
        .Select(d => new FolderEntry(d.FilePath, Path.GetFileName(d.FilePath), new FileInfo(d.FilePath).Length, d.Mode.ToString(), d.HasCachedIndex, d.BodyBytes, d.EstimatedIndexBytes, d.Reason))
        .OrderByDescending(e => e.BodyBytes).ToList();
    return Results.Ok(entries);
});

app.MapPost("/api/documents", (OpenRequest request) =>
{
    request = request with { Path = ResolvePath(request.Path) };
    if (!File.Exists(request.Path)) return Results.NotFound(new ErrorBody($"File not found: {request.Path}"));
    var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = request.ForceSidecar == true ? WordIndexMode.Sidecar : null,
                                 IndexesOnDemand = true };
    if (request.RebuildIndex == true && File.Exists(policy.IndexPathFor(request.Path))) File.Delete(policy.IndexPathFor(request.Path));
    var progress = new Progress<WordDocumentIndexProgress>(p => lastProgress = p);
    var watch = Stopwatch.StartNew();
    IWordWindowedDocument document;
    try { document = WordWindowedDocumentOpener.Open(request.Path, policy, progress); }
    catch (Exception exception) when (exception is IOException or InvalidDataException or NotSupportedException)
    {
        return Results.BadRequest(new ErrorBody(exception.Message));
    }
    double openMilliseconds = watch.Elapsed.TotalMilliseconds;
    int id = Interlocked.Increment(ref nextId);
    documents[id] = document;
    WordIndexDecision decision = document switch
    {
        DocxWindowedDocument docx => docx.Decision,
        DocBinaryWindowedDocument doc => doc.Decision,
        _ => new WordIndexDecision(request.Path, WordIndexMode.InMemory, false, 0, 0, ""),
    };
    (int footnotes, int endnotes) = document.NoteCounts;
    return Results.Ok(new DocumentSummary(id, Path.GetFileName(request.Path), request.Path, document.BlockCount, document.IsRightToLeft, document.Stylesheet,
                                          openMilliseconds, decision.Mode.ToString(), decision.HasCachedIndex, decision.Reason,
                                          document.GetType().Name, document.Headings.Count, footnotes, endnotes,
                                          document.KnownBlockCount, document.IsIndexComplete));
});

app.MapGet("/api/documents/{id:int}/headings", (int id) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    return Results.Ok(document.Headings.Select(h => new HeadingBody(h.BlockIndex, h.Level, h.Text)).ToList());
});

app.MapGet("/api/documents/{id:int}/note-references", (int id) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    return Results.Ok(document.NoteReferences.Select(r => new NoteReferenceBody(r.Kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote", r.Number, r.BlockIndex)).ToList());
});

app.MapGet("/api/documents/{id:int}/weights", (int id) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    // Raw Int32 little-endian: 123k blocks are 0.5 MB and parse in microseconds (JSON took ~200 ms).
    int blockCount = document.BlockCount;                   // one snapshot: an on-demand index may move it between calls
    var bytes = new byte[blockCount * 4];
    for (int i = 0; i < blockCount; i++) BitConverter.TryWriteBytes(bytes.AsSpan(i * 4), document.EstimatedTextLength(i));
    return Results.Bytes(bytes, "application/octet-stream");
});

app.MapGet("/api/documents/{id:int}/blocks", (int id, int first, int count, HttpResponse response) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    var watch = Stopwatch.StartNew();
    IReadOnlyList<WordDocumentBlockHtml> blocks;
    try { blocks = document.RenderBlocks(first, Math.Min(count, 500)); }
    catch (WordDocumentSourceChangedException) { return Results.Conflict(new ErrorBody("source-changed")); }   // the demo reopens
    response.Headers["X-Render-Milliseconds"] = watch.Elapsed.TotalMilliseconds.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture);
    response.Headers["X-Known-Block-Count"] = document.KnownBlockCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
    response.Headers["X-Index-Complete"] = document.IsIndexComplete ? "true" : "false";
    return Results.Ok(blocks.Select(b => new BlockBody(b.BlockIndex, (int)b.Kind, b.Html,
        b.Notes.Select(n => new NoteBody(n.Kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote", n.Number, n.Label, n.Html)).ToList())).ToList());
});

// A first open keeps building its index: the demo calls this until it is complete (each call indexes `blocks` more blocks; renders and sends nothing).
app.MapPost("/api/documents/{id:int}/scan", (int id, int blocks) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    try { document.ScanAhead(Math.Clamp(blocks, 1, 100_000)); }
    catch (WordDocumentSourceChangedException) { return Results.Conflict(new ErrorBody("source-changed")); }
    (int footnotes, int endnotes) = document.NoteCounts;
    return Results.Ok(new ScanBody(document.KnownBlockCount, document.IsIndexComplete, footnotes, endnotes));
});

app.MapDelete("/api/documents/{id:int}", (int id) =>
{
    if (documents.TryRemove(id, out IWordWindowedDocument? document)) document.Dispose();
    return Results.Ok();
});

// Closes everything — a Playwright run that ends by closing the browser never sends its DELETEs.
app.MapDelete("/api/documents", () =>
{
    foreach (int id in documents.Keys.ToList()) if (documents.TryRemove(id, out IWordWindowedDocument? document)) document.Dispose();
    return Results.Ok();
});

app.MapGet("/api/progress", () => Results.Ok(new ProgressBody(lastProgress.BytesDone, lastProgress.BytesTotal, lastProgress.BlocksFound)));

app.MapGet("/api/memory", (bool? collect) =>
{
    if (collect == true) { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); }
    return Results.Ok(new MemoryBody(GC.GetTotalMemory(false) >> 20, Environment.WorkingSet >> 20, documents.Count));
});

Console.WriteLine($"FastDocView dev server on {Url}  (index cache: {cacheDirectory})");
app.Run();

// ── wire shapes ────────────────────────────────────────────────────────────────────

record OpenRequest(string Path, bool? ForceSidecar, bool? RebuildIndex);
record ErrorBody(string Error);
record FolderEntry(string Path, string FileName, long FileBytes, string Mode, bool HasCachedIndex, long BodyBytes, long EstimatedIndexBytes, string Reason);
record HeadingBody(int BlockIndex, int Level, string Text);
record DocumentSummary(int Id, string FileName, string Path, int BlockCount, bool IsRightToLeft, string Stylesheet,
                       double OpenMilliseconds, string Mode, bool HasCachedIndex, string Reason, string ReaderType,
                       int HeadingCount, int FootnoteCount, int EndnoteCount, int KnownBlockCount, bool IsIndexComplete);
record NoteBody(string Kind, int Number, string Label, string Html);
record NoteReferenceBody(string Kind, int Number, int BlockIndex);
record ScanBody(int KnownBlockCount, bool IsIndexComplete, int FootnoteCount, int EndnoteCount);
record BlockBody(int Index, int Kind, string Html, List<NoteBody> Notes);
record ProgressBody(long BytesDone, long BytesTotal, int BlocksFound);
record MemoryBody(long ManagedHeapMegabytes, long WorkingSetMegabytes, int OpenDocuments);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(OpenRequest))]
[JsonSerializable(typeof(ErrorBody))]
[JsonSerializable(typeof(List<FolderEntry>))]
[JsonSerializable(typeof(DocumentSummary))]
[JsonSerializable(typeof(List<HeadingBody>))]
[JsonSerializable(typeof(List<NoteReferenceBody>))]
[JsonSerializable(typeof(List<BlockBody>))]
[JsonSerializable(typeof(ProgressBody))]
[JsonSerializable(typeof(MemoryBody))]
partial class DevServerJsonContext : JsonSerializerContext { }
