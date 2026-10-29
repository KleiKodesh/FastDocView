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

var cacheDirectory = Path.Combine(AppContext.BaseDirectory, "index-cache");
using var documents = new WordWindowedDocumentSet(cacheDirectory);     // the same set and locking the app's service uses
var lastProgress = new WordDocumentIndexProgress(0, 0, 0);

// Runs a call on an open document: 404 for an unknown id, 409 "source-changed" when it must be reopened (the demo reopens).
IResult UseDocument(string id, Func<IWordWindowedDocument, IResult> action)
{
    try { return documents.Use(id, action); }
    catch (WordDocumentClosedException) { return Results.Conflict(new ErrorBody("source-changed")); }
}

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
    string indexPath = documents.Policy.IndexPathFor(request.Path);
    if (request.RebuildIndex == true && File.Exists(indexPath)) File.Delete(indexPath);
    var progress = new Progress<WordDocumentIndexProgress>(p => lastProgress = p);
    var watch = Stopwatch.StartNew();
    string id;
    try { id = documents.Open(request.Path, request.ForceSidecar == true ? WordIndexMode.Sidecar : null, progress); }
    catch (Exception exception) when (exception is IOException or InvalidDataException or NotSupportedException)
    {
        return Results.BadRequest(new ErrorBody(exception.Message));
    }
    double openMilliseconds = watch.Elapsed.TotalMilliseconds;
    return UseDocument(id, document =>
    {
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
});

// Page text width per section (Word / LibreOffice margins), in block order - the app page sizes its column from it.
app.MapGet("/api/documents/{id}/sections", (string id) => UseDocument(id, document =>
    Results.Ok(document.Sections.Select(s => new SectionBody(s.FirstBlockIndex, s.TextWidthTwips)).ToList())));

// In-document search, one batch: hit block indexes from fromBlock on (the app page loops, wraps and highlights).
app.MapGet("/api/documents/{id}/search", (string id, string query, bool matchCase, int fromBlock, int maxBlocks) => UseDocument(id, document =>
{
    WordDocumentSearchBatch batch = document.Search(query, new WordDocumentSearchOptions { MatchCase = matchCase }, fromBlock, Math.Clamp(maxBlocks, 1, 2000));
    return Results.Ok(new SearchBody(batch.HitBlocks, batch.NextBlock, batch.ReachedEnd));
}));

// The block a bookmark marks (in-document links, TOC entries, REF fields): an index lookup.
app.MapGet("/api/documents/{id}/bookmark", (string id, string name) => UseDocument(id, document =>
    Results.Ok(new BookmarkBody(document.FindBookmark(name)))));

app.MapGet("/api/documents/{id}/headings", (string id) => UseDocument(id, document =>
    Results.Ok(document.Headings.Select(h => new HeadingBody(h.BlockIndex, h.Level, h.Text)).ToList())));

app.MapGet("/api/documents/{id}/note-references", (string id) => UseDocument(id, document =>
    Results.Ok(document.NoteReferences.Select(r => new NoteReferenceBody(r.Kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote", r.Number, r.BlockIndex)).ToList())));

app.MapGet("/api/documents/{id}/weights", (string id) => UseDocument(id, document =>
{
    // Raw Int32 little-endian: 123k blocks are 0.5 MB and parse in microseconds (JSON took ~200 ms).
    int[] lengths = document.EstimatedTextLengths();
    var bytes = new byte[lengths.Length * 4];
    Buffer.BlockCopy(lengths, 0, bytes, 0, bytes.Length);
    return Results.Bytes(bytes, "application/octet-stream");
}));

app.MapGet("/api/documents/{id}/blocks", (string id, int first, int count, HttpResponse response) => UseDocument(id, document =>
{
    var watch = Stopwatch.StartNew();
    IReadOnlyList<WordDocumentBlockHtml> blocks = document.RenderBlocks(first, Math.Min(count, 500));
    response.Headers["X-Render-Milliseconds"] = watch.Elapsed.TotalMilliseconds.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture);
    response.Headers["X-Known-Block-Count"] = document.KnownBlockCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
    response.Headers["X-Index-Complete"] = document.IsIndexComplete ? "true" : "false";
    return Results.Ok(blocks.Select(b => new BlockBody(b.BlockIndex, (int)b.Kind, b.Html,
        b.Notes.Select(n => new NoteBody(n.Kind == WordDocumentNoteKind.Footnote ? "footnote" : "endnote", n.Number, n.Label, n.Html)).ToList())).ToList());
}));

// A first open keeps building its index: the demo calls this until it is complete. Each call indexes up to `blocks` more
// blocks within the app service's time budget (renders and sends nothing), so a block fetch never waits long for the lock.
app.MapPost("/api/documents/{id}/scan", (string id, int blocks) => UseDocument(id, document =>
{
    document.ScanFor(Math.Clamp(blocks, 1, 100_000), TimeSpan.FromMilliseconds(20));
    (int footnotes, int endnotes) = document.NoteCounts;
    return Results.Ok(new ScanBody(document.KnownBlockCount, document.IsIndexComplete, footnotes, endnotes));
}));

app.MapDelete("/api/documents/{id}", (string id) =>
{
    documents.Close(id);
    return Results.Ok();
});

// Closes everything — a Playwright run that ends by closing the browser never sends its DELETEs.
app.MapDelete("/api/documents", () =>
{
    documents.CloseAll();
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
record SectionBody(int FirstBlockIndex, int TextWidthTwips);
record BookmarkBody(int? BlockIndex);
record SearchBody(int[] HitBlocks, int NextBlock, bool ReachedEnd);
record DocumentSummary(string Id, string FileName, string Path, int BlockCount, bool IsRightToLeft, string Stylesheet,
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
