using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Web.WebView2.Core;

namespace WordDocumentRandomAccess.Viewer;

/// <summary>
/// JSON messages between viewer.js and the document. The page asks for what it needs:
///   page → host:  { type: "ready" }                                     the page script is up
///                 { type: "requestBlocks", first, count, requestId }    blocks near the viewport
///   host → page:  { type: "document", blockCount, isRightToLeft, stylesheet, estimatedTextLengths, headings, fileName }
///                 { type: "blocks", requestId, blocks: [{ index, kind, html }], renderMilliseconds }
///                 { type: "closed" }
/// The "blocks" shape is deliberately the one the app's lines view consumes (index + html).
/// </summary>
public sealed class ViewerWebMessageBridge
{
    private readonly CoreWebView2 _webView;
    private readonly Action<string> _setStatus;
    private IWordWindowedDocument? _document;
    private bool _isPageReady;
    private bool _hasRenderedFirstWindow;

    public event Action? FirstWindowRendered;

    public ViewerWebMessageBridge(CoreWebView2 webView, Action<string> setStatus)
    {
        _webView = webView;
        _setStatus = setStatus;
        _webView.WebMessageReceived += OnWebMessageReceived;
    }

    public void SetDocument(IWordWindowedDocument? document)
    {
        _document = document;
        _hasRenderedFirstWindow = false;
        if (_isPageReady) SendDocument();
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        JsonNode? message = JsonNode.Parse(e.WebMessageAsJson);
        string? type = (string?)message?["type"];
        switch (type)
        {
            case "ready":
                _isPageReady = true;
                SendDocument();
                break;
            case "requestBlocks":
                SendBlocks((int)(message!["first"] ?? 0), (int)(message["count"] ?? 40), (int)(message["requestId"] ?? 0));
                break;
        }
    }

    private void SendDocument()
    {
        if (_document is null) { Post(new JsonObject { ["type"] = "closed" }); return; }
        IWordWindowedDocument document = _document;
        var estimatedTextLengths = new JsonArray();
        for (int i = 0; i < document.BlockCount; i++) estimatedTextLengths.Add(document.EstimatedTextLength(i));
        var headings = new JsonArray();
        foreach (WordDocumentHeading heading in document.Headings.Take(5000))
            headings.Add(new JsonObject { ["blockIndex"] = heading.BlockIndex, ["level"] = heading.Level, ["text"] = heading.Text });
        Post(new JsonObject
        {
            ["type"] = "document",
            ["fileName"] = Path.GetFileName(document.FilePath),
            ["blockCount"] = document.BlockCount,
            ["isRightToLeft"] = document.IsRightToLeft,
            ["stylesheet"] = document.Stylesheet,
            ["estimatedTextLengths"] = estimatedTextLengths,
            ["headings"] = headings,
        });
    }

    private void SendBlocks(int first, int count, int requestId)
    {
        if (_document is null) return;
        var watch = Stopwatch.StartNew();
        IReadOnlyList<WordDocumentBlockHtml> blocks = _document.RenderBlocks(first, count);
        var array = new JsonArray();
        foreach (WordDocumentBlockHtml block in blocks)
            array.Add(new JsonObject { ["index"] = block.BlockIndex, ["kind"] = (int)block.Kind, ["html"] = block.Html });
        Post(new JsonObject { ["type"] = "blocks", ["requestId"] = requestId, ["first"] = first, ["blocks"] = array, ["renderMilliseconds"] = watch.Elapsed.TotalMilliseconds });
        _setStatus($"Blocks {first:N0}–{first + blocks.Count - 1:N0} of {_document.BlockCount:N0} rendered in {watch.Elapsed.TotalMilliseconds:0.#} ms");
        if (!_hasRenderedFirstWindow) { _hasRenderedFirstWindow = true; FirstWindowRendered?.Invoke(); }
    }

    private void Post(JsonNode message) => _webView.PostWebMessageAsJson(message.ToJsonString(new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }));
}
