// Localhost JSON host for the Vite demo. Endpoints (all JSON, CORS open — dev only):
//   GET  /api/folder?path=…                  policy decision for every Word file under a folder
//   POST /api/documents  { path, forceSidecar? }   opens (indexing if needed) → document summary
//   GET  /api/documents/{id}/weights          estimated text length per block (scrollbar weights)
//   GET  /api/documents/{id}/blocks?first&count    rendered blocks [{ index, kind, html }]
//   DELETE /api/documents/{id}                closes
//   GET  /api/progress                        the last index progress (polled while opening)
//   GET  /api/memory                          managed heap / working set of this process
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text.Json.Serialization;
using WordDocumentRandomAccess;
using WordDocumentRandomAccess.DocBinary;
using WordDocumentRandomAccess.Docx;

const string Url = "http://127.0.0.1:5188";

var builder = WebApplication.CreateSlimBuilder(args);
builder.WebHost.UseUrls(Url);
builder.Services.AddCors(options => options.AddDefaultPolicy(policy => policy.AllowAnyOrigin().AllowAnyHeader().AllowAnyMethod().WithExposedHeaders("X-Render-Milliseconds")));
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
    var policy = new WordIndexPolicy { CacheDirectory = cacheDirectory, ForcedMode = request.ForceSidecar == true ? WordIndexMode.Sidecar : null };
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
    var headingsWatch = Stopwatch.StartNew();
    List<HeadingBody> headings = document.Headings.Take(5000).Select(h => new HeadingBody(h.BlockIndex, h.Level, h.Text)).ToList();
    return Results.Ok(new DocumentSummary(id, Path.GetFileName(request.Path), request.Path, document.BlockCount, document.IsRightToLeft, document.Stylesheet,
                                          headings, openMilliseconds, headingsWatch.Elapsed.TotalMilliseconds, decision.Mode.ToString(), decision.HasCachedIndex, decision.Reason,
                                          document.GetType().Name));
});

app.MapGet("/api/documents/{id:int}/weights", (int id) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    var weights = new int[document.BlockCount];
    for (int i = 0; i < weights.Length; i++) weights[i] = document.EstimatedTextLength(i);
    return Results.Ok(weights);
});

app.MapGet("/api/documents/{id:int}/blocks", (int id, int first, int count, HttpResponse response) =>
{
    if (!documents.TryGetValue(id, out IWordWindowedDocument? document)) return Results.NotFound();
    var watch = Stopwatch.StartNew();
    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(first, Math.Min(count, 500));
    response.Headers["X-Render-Milliseconds"] = watch.Elapsed.TotalMilliseconds.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture);
    return Results.Ok(blocks.Select(b => new BlockBody(b.BlockIndex, (int)b.Kind, b.Html)).ToList());
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

Console.WriteLine($"WordDocumentRandomAccess dev server on {Url}  (index cache: {cacheDirectory})");
app.Run();

// ── wire shapes ────────────────────────────────────────────────────────────────────

record OpenRequest(string Path, bool? ForceSidecar, bool? RebuildIndex);
record ErrorBody(string Error);
record FolderEntry(string Path, string FileName, long FileBytes, string Mode, bool HasCachedIndex, long BodyBytes, long EstimatedIndexBytes, string Reason);
record HeadingBody(int BlockIndex, int Level, string Text);
record DocumentSummary(int Id, string FileName, string Path, int BlockCount, bool IsRightToLeft, string Stylesheet, List<HeadingBody> Headings,
                       double OpenMilliseconds, double HeadingsMilliseconds, string Mode, bool HasCachedIndex, string Reason, string ReaderType);
record BlockBody(int Index, int Kind, string Html);
record ProgressBody(long BytesDone, long BytesTotal, int BlocksFound);
record MemoryBody(long ManagedHeapMegabytes, long WorkingSetMegabytes, int OpenDocuments);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase)]
[JsonSerializable(typeof(OpenRequest))]
[JsonSerializable(typeof(ErrorBody))]
[JsonSerializable(typeof(List<FolderEntry>))]
[JsonSerializable(typeof(DocumentSummary))]
[JsonSerializable(typeof(int[]))]
[JsonSerializable(typeof(List<BlockBody>))]
[JsonSerializable(typeof(ProgressBody))]
[JsonSerializable(typeof(MemoryBody))]
partial class DevServerJsonContext : JsonSerializerContext { }
