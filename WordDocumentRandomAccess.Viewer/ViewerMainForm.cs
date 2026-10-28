using System.Diagnostics;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordDocumentRandomAccess.Viewer;

/// <summary>
/// The viewer window: a menu to open a file, a status strip with index progress and timings, and
/// the WebView2 that hosts WebContent/viewer.html. Document access goes through
/// <see cref="ViewerWebMessageBridge"/>; this form only owns the window and the open/close flow.
/// </summary>
public sealed class ViewerMainForm : Form
{
    private readonly WebView2 _webView = new() { Dock = DockStyle.Fill };
    private readonly ToolStripStatusLabel _status = new() { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
    private readonly string? _screenshotPath;
    private ViewerWebMessageBridge? _bridge;
    private IWordWindowedDocument? _document;
    private string? _pendingFilePath;

    public ViewerMainForm(string? filePath, string? screenshotPath)
    {
        _pendingFilePath = filePath;
        _screenshotPath = screenshotPath;
        Text = "Word random-access viewer";
        Width = 1100;
        Height = 800;
        RightToLeftLayout = false;

        var menu = new MenuStrip();
        var fileMenu = new ToolStripMenuItem("&File");
        fileMenu.DropDownItems.Add("&Open…", null, (_, _) => OnOpenClicked());
        fileMenu.DropDownItems.Add("Open &folder (index decisions)…", null, (_, _) => OnOpenFolderClicked());
        fileMenu.DropDownItems.Add("&Rebuild index and reopen", null, (_, _) => OnRebuildIndexClicked());
        fileMenu.DropDownItems.Add(new ToolStripSeparator());
        fileMenu.DropDownItems.Add("E&xit", null, (_, _) => Close());
        menu.Items.Add(fileMenu);
        MainMenuStrip = menu;

        var statusStrip = new StatusStrip();
        statusStrip.Items.Add(_status);

        Controls.Add(_webView);
        Controls.Add(statusStrip);
        Controls.Add(menu);
        Load += async (_, _) => await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        string userDataFolder = Path.Combine(Path.GetTempPath(), "WordDocumentRandomAccess.Viewer");
        CoreWebView2Environment environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
        await _webView.EnsureCoreWebView2Async(environment);
        string webContent = Path.Combine(AppContext.BaseDirectory, "WebContent");
        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping("viewer.local", webContent, CoreWebView2HostResourceAccessKind.Allow);
        _bridge = new ViewerWebMessageBridge(_webView.CoreWebView2, SetStatus);
        _bridge.FirstWindowRendered += OnFirstWindowRendered;
        _webView.CoreWebView2.Navigate("https://viewer.local/viewer.html");
        if (_pendingFilePath is not null) await OpenDocumentAsync(_pendingFilePath, rebuildIndex: false);
    }

    private void OnOpenClicked()
    {
        using var dialog = new OpenFileDialog { Filter = "Word documents|*.docx;*.docm;*.dotx;*.doc|All files|*.*" };
        if (dialog.ShowDialog(this) == DialogResult.OK) _ = OpenDocumentAsync(dialog.FileName, rebuildIndex: false);
    }

    private void OnOpenFolderClicked()
    {
        using var dialog = new FolderBrowserDialog();
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var folderForm = new ViewerFolderIndexForm(dialog.SelectedPath, WordIndexPolicy.Default);
        folderForm.FileChosen += path => _ = OpenDocumentAsync(path, rebuildIndex: false);
        folderForm.Show(this);
    }

    private void OnRebuildIndexClicked()
    {
        if (_document?.FilePath is { } path) _ = OpenDocumentAsync(path, rebuildIndex: true);
    }

    private async Task OpenDocumentAsync(string filePath, bool rebuildIndex)
    {
        _pendingFilePath = filePath;
        _document?.Dispose();
        _document = null;
        _bridge?.SetDocument(null);
        Text = $"{Path.GetFileName(filePath)} — Word random-access viewer";

        if (rebuildIndex)
        {
            string indexPath = WordDocumentIndexCacheLocation.IndexPathFor(filePath);
            if (File.Exists(indexPath)) File.Delete(indexPath);
        }

        var watch = Stopwatch.StartNew();
        var progress = new Progress<WordDocumentIndexProgress>(p =>
            SetStatus($"Indexing… {p.BytesDone >> 20} / {p.BytesTotal >> 20} MB, {p.BlocksFound:N0} blocks"));
        try
        {
            IWordWindowedDocument document = await Task.Run(() => WordWindowedDocumentOpener.Open(filePath, WordIndexPolicy.Default, progress));
            _document = document;
            string mode = document switch
            {
                Docx.DocxWindowedDocument docx => $"{docx.Decision.Mode} ({docx.Decision.Reason})",
                DocBinary.DocBinaryWindowedDocument doc => $"{doc.Decision.Mode} ({doc.Decision.Reason})",
                _ => "",
            };
            SetStatus($"Opened in {watch.ElapsedMilliseconds} ms — {document.BlockCount:N0} blocks, {(document.IsRightToLeft ? "RTL" : "LTR")}, {mode}");
            _bridge?.SetDocument(document);
        }
        catch (Exception exception)
        {
            SetStatus($"Failed: {exception.Message}");
            MessageBox.Show(this, exception.Message, "Could not open", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private async void OnFirstWindowRendered()
    {
        if (_screenshotPath is null) return;
        await Task.Delay(1500);
        using var file = File.Create(_screenshotPath);
        await _webView.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, file);
        Close();
    }

    private void SetStatus(string text)
    {
        if (InvokeRequired) { BeginInvoke(() => _status.Text = text); return; }
        _status.Text = text;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _document?.Dispose();
        base.OnFormClosed(e);
    }
}
