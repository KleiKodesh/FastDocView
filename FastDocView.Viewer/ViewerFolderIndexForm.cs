using System.Diagnostics;

namespace FastDocView.Viewer;

/// <summary>
/// Demo of the "index intelligently" flow: pick a folder, see every Word file with the policy's
/// decision (in-memory vs sidecar, body size, reason), tick the ones to index now, double-click to
/// open. Plain ListView — a demo UI, not the app's.
/// </summary>
public sealed class ViewerFolderIndexForm : Form
{
    private readonly ListView _list = new() { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, CheckBoxes = true };
    private readonly Label _status = new() { Dock = DockStyle.Bottom, Height = 22, Padding = new Padding(6, 3, 6, 0) };
    private readonly WordIndexPolicy _policy;

    public event Action<string>? FileChosen;

    public ViewerFolderIndexForm(string folderPath, WordIndexPolicy policy)
    {
        _policy = policy;
        Text = $"Word files under {folderPath}";
        Width = 1000;
        Height = 600;
        _list.Columns.Add("File", 420);
        _list.Columns.Add("Mode", 90);
        _list.Columns.Add("Cached", 60);
        _list.Columns.Add("Body", 80, HorizontalAlignment.Right);
        _list.Columns.Add("Index ≈", 80, HorizontalAlignment.Right);
        _list.Columns.Add("Reason", 240);
        _list.DoubleClick += (_, _) => { if (_list.SelectedItems.Count > 0) FileChosen?.Invoke((string)_list.SelectedItems[0].Tag!); };

        var buttons = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, Padding = new Padding(4) };
        var indexButton = new Button { Text = "Build sidecar index for checked", AutoSize = true };
        indexButton.Click += async (_, _) => await BuildCheckedAsync();
        var checkWorthwhileButton = new Button { Text = "Check all the policy recommends", AutoSize = true };
        checkWorthwhileButton.Click += (_, _) => { foreach (ListViewItem item in _list.Items) item.Checked = item.SubItems[1].Text == "Sidecar" && item.SubItems[2].Text != "yes"; };
        buttons.Controls.Add(indexButton);
        buttons.Controls.Add(checkWorthwhileButton);

        Controls.Add(_list);
        Controls.Add(buttons);
        Controls.Add(_status);
        Load += async (_, _) => await ScanAsync(folderPath);
    }

    private async Task ScanAsync(string folderPath)
    {
        _status.Text = "Scanning…";
        var watch = Stopwatch.StartNew();
        List<WordIndexDecision> decisions = await Task.Run(() => WordWindowedDocumentOpener.DecideFolder(folderPath, _policy).ToList());
        _list.BeginUpdate();
        foreach (WordIndexDecision decision in decisions)
        {
            var item = new ListViewItem(Path.GetFileName(decision.FilePath)) { Tag = decision.FilePath, ToolTipText = decision.FilePath };
            item.SubItems.Add(decision.Mode.ToString());
            item.SubItems.Add(decision.HasCachedIndex ? "yes" : "");
            item.SubItems.Add(FormatBytes(decision.BodyBytes));
            item.SubItems.Add(decision.EstimatedIndexBytes > 0 ? FormatBytes(decision.EstimatedIndexBytes) : "");
            item.SubItems.Add(decision.Reason);
            _list.Items.Add(item);
        }
        _list.EndUpdate();
        int sidecar = decisions.Count(d => d.Mode == WordIndexMode.Sidecar);
        _status.Text = $"{decisions.Count} files in {watch.ElapsedMilliseconds} ms — {sidecar} worth a sidecar index, {decisions.Count - sidecar} open by one pass";
    }

    private async Task BuildCheckedAsync()
    {
        var chosen = _list.CheckedItems.Cast<ListViewItem>().Select(item => (string)item.Tag!).ToList();
        if (chosen.Count == 0) return;
        var progress = new Progress<(string file, WordDocumentIndexProgress progress)>(p =>
            _status.Text = $"Indexing {Path.GetFileName(p.file)}: {p.progress.BytesDone >> 20} / {p.progress.BytesTotal >> 20} MB");
        await Task.Run(() => WordWindowedDocumentOpener.BuildSidecarIndexes(
            chosen.Select(path => WordWindowedDocumentOpener.Decide(path, _policy) with { }), _policy, progress));
        foreach (ListViewItem item in _list.CheckedItems) { item.SubItems[2].Text = "yes"; item.Checked = false; }
        _status.Text = $"Indexed {chosen.Count} file(s)";
    }

    private static string FormatBytes(long bytes) => bytes >= 1 << 20 ? $"{bytes >> 20} MB" : $"{Math.Max(1, bytes >> 10)} KB";
}
