using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using OutlookSearch.Services;

namespace OutlookSearch;

public class MainForm : Form
{
    private OutlookService? _svc;
    private readonly AdLookupService _ad = new();
    private readonly ExportService _export = new();
    private AdAutoComplete? _fromAutoComplete;
    private readonly List<EmailResult> _results = [];
    private CancellationTokenSource _cts = new();
    private int _previewSeq;   // guards against a stale body overwriting a newer selection

    // ── Left: mailboxes & folders ─────────────────────────────────
    private readonly Button _refreshBtn = new() { Text = "Refresh", Width = 70 };
    private readonly Button _openMailboxBtn = new() { Text = "Open another mailbox…", Width = 150 };
    private readonly Button _checkAllBtn = new() { Text = "All", Width = 38 };
    private readonly Button _uncheckAllBtn = new() { Text = "None", Width = 44 };
    private readonly TreeView _folderTree = new() { Dock = DockStyle.Fill, CheckBoxes = true, ShowLines = true };

    // ── Right: search criteria ────────────────────────────────────
    private readonly TextBox _kwBox = new() { Dock = DockStyle.Fill, PlaceholderText = "Searches subject and body" };
    private readonly TextBox _subjectBox = new() { Dock = DockStyle.Fill };
    private readonly TextBox _fromBox = new() { Dock = DockStyle.Fill, PlaceholderText = "Type a name — autocompletes from AD" };
    private readonly DateTimePicker _dateFromPicker = new()
    {
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = true,
        Checked = false,
        Value = DateTime.Today.AddDays(-90),
        Width = 125
    };
    private readonly DateTimePicker _dateToPicker = new()
    {
        Format = DateTimePickerFormat.Short,
        ShowCheckBox = true,
        Checked = false,
        Value = DateTime.Today,
        Width = 125
    };
    private readonly Button _clearDatesBtn = new()
    {
        Text = "Clear",
        Width = 55,
        Height = 24,
        Enabled = false,
        Margin = new Padding(8, 1, 0, 0)
    };
    private readonly Button _searchBtn = new() { Text = "Search Mail", Width = 100, Height = 30, Enabled = false };
    private readonly Button _cancelBtn = new() { Text = "Cancel", Width = 70, Height = 30, Enabled = false };
    private readonly Button _clearAllBtn = new() { Text = "Clear All", Width = 75, Height = 30, Enabled = false };
    private readonly Button _exportBtn = new() { Text = "Export…", Width = 75, Height = 30, Enabled = false };
    private readonly Label _countLabel = new() { AutoSize = true, Text = "0 results", ForeColor = Color.Gray };

    // ── Right: results + preview ──────────────────────────────────
    private readonly ListView _resultList = new()
    {
        Dock = DockStyle.Fill,
        View = View.Details,
        FullRowSelect = true,
        GridLines = true,
        MultiSelect = true
    };
    private readonly RichTextBox _bodyPreview = new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        ScrollBars = RichTextBoxScrollBars.Vertical,
        Font = new Font("Segoe UI", 9),
        BackColor = Color.FromArgb(252, 252, 252)
    };

    // ── Status ────────────────────────────────────────────────────
    private readonly StatusStrip _statusBar = new();
    private readonly ToolStripStatusLabel _statusLabel = new()
    {
        Spring = true,
        TextAlign = ContentAlignment.MiddleLeft,
        Text = "Connecting to Outlook…"
    };
    private readonly ToolStripProgressBar _progressBar = new() { Visible = false, Width = 100 };

    public MainForm()
    {
        Text = "Outlook Mail Search";
        Size = new Size(1200, 760);
        MinimumSize = new Size(860, 520);
        Font = new Font("Segoe UI", 9);

        BuildLayout();
        SetupListView();
        WireEvents();

        _fromAutoComplete = new AdAutoComplete(_fromBox, _ad);

        Shown += async (_, _) => await InitOutlookAsync();
    }

    // ── Layout ────────────────────────────────────────────────────

    private void BuildLayout()
    {
        _statusBar.Items.AddRange([_statusLabel, _progressBar]);
        Controls.Add(_statusBar);

        var mainSplit = new SplitContainer
        {
            // Explicit size first so the min-size/splitter constraints are
            // satisfiable before Dock=Fill resizes us to the parent.
            Size = new Size(1000, 650),
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            FixedPanel = FixedPanel.Panel1,
            Panel1MinSize = 240,
            Panel2MinSize = 450,
            SplitterDistance = 300
        };

        BuildLeftPanel(mainSplit.Panel1);
        BuildRightPanel(mainSplit.Panel2);
        Controls.Add(mainSplit);
    }

    private void BuildLeftPanel(SplitterPanel panel)
    {
        var group = new GroupBox
        {
            Text = "Mailboxes & Folders  (check to include in search)",
            Dock = DockStyle.Fill,
            Padding = new Padding(6, 4, 6, 6)
        };

        var topRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 30,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        topRow.Controls.AddRange([_refreshBtn, _openMailboxBtn]);

        var selectRow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 28,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        selectRow.Controls.AddRange([
            new Label { Text = "Check:", AutoSize = true, Padding = new Padding(2, 6, 2, 0) },
            _checkAllBtn, _uncheckAllBtn
        ]);

        var treePanel = new Panel { Dock = DockStyle.Fill };
        treePanel.Controls.Add(_folderTree);

        group.Controls.Add(treePanel);
        group.Controls.Add(selectRow);
        group.Controls.Add(topRow);
        panel.Controls.Add(group);
    }

    private void BuildRightPanel(SplitterPanel panel)
    {
        var criteriaPanel = new Panel
        {
            Dock = DockStyle.Top,
            Height = 178,
            Padding = new Padding(6, 4, 6, 4)
        };

        var tbl = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 4,
            RowCount = 4
        };
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 68));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 68));
        tbl.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));   // keyword / from
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));   // subject / date
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));   // buttons (taller so they aren't clipped)
        tbl.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));   // spare

        tbl.Controls.Add(RightLabel("Keyword:"), 0, 0);
        tbl.Controls.Add(_kwBox, 1, 0);
        tbl.Controls.Add(RightLabel("From:"), 2, 0);
        tbl.Controls.Add(_fromBox, 3, 0);

        tbl.Controls.Add(RightLabel("Subject:"), 0, 1);
        tbl.Controls.Add(_subjectBox, 1, 1);
        tbl.Controls.Add(RightLabel("Date:"), 2, 1);

        var datePanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false
        };
        datePanel.Controls.Add(_dateFromPicker);
        datePanel.Controls.Add(new Label { Text = "to", AutoSize = true, Padding = new Padding(4, 8, 4, 0) });
        datePanel.Controls.Add(_dateToPicker);
        datePanel.Controls.Add(_clearDatesBtn);
        tbl.Controls.Add(datePanel, 3, 1);

        var btnPanel = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            Padding = new Padding(0, 6, 0, 0)
        };
        _countLabel.Padding = new Padding(8, 12, 0, 0);
        btnPanel.Controls.AddRange([_searchBtn, _cancelBtn, _clearAllBtn, _exportBtn, _countLabel]);
        tbl.Controls.Add(btnPanel, 0, 2);
        tbl.SetColumnSpan(btnPanel, 4);

        criteriaPanel.Controls.Add(tbl);

        var resultsSplit = new SplitContainer
        {
            Size = new Size(600, 600),
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            Panel1MinSize = 80,
            Panel2MinSize = 60,
            SplitterDistance = 320
        };
        resultsSplit.Panel1.Controls.Add(_resultList);
        resultsSplit.Panel2.Controls.Add(_bodyPreview);

        panel.Controls.Add(resultsSplit);
        panel.Controls.Add(criteriaPanel);
    }

    private void SetupListView()
    {
        _resultList.Columns.AddRange([
            new ColumnHeader { Text = "Subject",  Width = 300 },
            new ColumnHeader { Text = "From",     Width = 200 },
            new ColumnHeader { Text = "Received", Width = 140 },
            new ColumnHeader { Text = "Folder",   Width = 130 },
            new ColumnHeader { Text = "Att",      Width = 34  },
        ]);
    }

    // ── Events ────────────────────────────────────────────────────

    private void WireEvents()
    {
        _refreshBtn.Click += async (_, _) => await LoadMailboxesAsync();
        _openMailboxBtn.Click += OnOpenMailbox;
        _checkAllBtn.Click += (_, _) => SetAllChecked(_folderTree.Nodes, true);
        _uncheckAllBtn.Click += (_, _) => SetAllChecked(_folderTree.Nodes, false);
        _searchBtn.Click += OnSearch;
        _cancelBtn.Click += (_, _) => { _cts.Cancel(); SetStatus("Cancelling…"); };
        _clearAllBtn.Click += OnClearAll;
        _exportBtn.Click += OnExport;
        _resultList.SelectedIndexChanged += OnResultSelected;

        // A checked date box means the date filter is active; Clear unchecks both.
        _clearDatesBtn.Click += (_, _) =>
        {
            _dateFromPicker.Checked = false;
            _dateToPicker.Checked = false;
            UpdateClearDatesState();
        };
        _dateFromPicker.ValueChanged += (_, _) => UpdateClearDatesState();
        _dateToPicker.ValueChanged += (_, _) => UpdateClearDatesState();

        _folderTree.AfterCheck += (_, e) =>
        {
            if (e.Action != TreeViewAction.Unknown && e.Node != null)
                SetAllChecked(e.Node.Nodes, e.Node.Checked);
        };
    }

    private async Task InitOutlookAsync()
    {
        try
        {
            _svc = new OutlookService();
            await LoadMailboxesAsync();
        }
        catch (Exception ex)
        {
            SetStatus($"Cannot connect to Outlook: {ex.Message}");
        }
    }

    private async Task LoadMailboxesAsync()
    {
        if (_svc is null) return;
        SetBusy("Loading mailboxes and folders…");
        try
        {
            var trees = await _svc.GetMailboxTreesAsync();
            _folderTree.BeginUpdate();
            _folderTree.Nodes.Clear();
            foreach (var t in trees)
            {
                var node = BuildTreeNode(t);
                node.Checked = false;   // don't pre-check whole mailboxes
                _folderTree.Nodes.Add(node);
                node.Expand();
            }
            _folderTree.EndUpdate();
            _searchBtn.Enabled = _folderTree.Nodes.Count > 0;
            SetStatus(_folderTree.Nodes.Count == 0
                ? "No mailboxes found in this Outlook profile."
                : $"Loaded {_folderTree.Nodes.Count} mailbox(es). Check folders, then Search Mail.");
        }
        catch (Exception ex) { SetStatus($"Failed to load folders: {ex.Message}"); }
        finally { SetIdle(); }
    }

    private async void OnOpenMailbox(object? s, EventArgs e)
    {
        if (_svc is null) return;
        var who = ShowInputDialog("Open another mailbox",
            "Enter a name or email address (resolved via the Outlook / org address book):");
        if (string.IsNullOrWhiteSpace(who)) return;

        SetBusy($"Opening mailbox for '{who}'…");
        try
        {
            var tree = await _svc.OpenSharedMailboxAsync(who.Trim());
            if (tree is null)
            {
                SetStatus($"Could not open a mailbox for '{who}'.");
                return;
            }
            var node = BuildTreeNode(tree);
            node.Checked = false;
            _folderTree.Nodes.Add(node);
            node.Expand();
            _searchBtn.Enabled = true;
            SetStatus($"Added mailbox '{node.Text}'. Check folders, then Search Mail.");
        }
        catch (Exception ex)
        {
            SetStatus($"Open mailbox failed: {ex.Message}");
            MessageBox.Show(ex.Message,
                "Could not open mailbox", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally { SetIdle(); }
    }

    private static TreeNode BuildTreeNode(MailFolderNode folder)
    {
        var node = new TreeNode(folder.Name) { Tag = folder };
        foreach (var child in folder.Children)
            node.Nodes.Add(BuildTreeNode(child));
        return node;
    }

    private async void OnSearch(object? s, EventArgs e)
    {
        if (_svc is null) return;

        var checkedFolders = GetCheckedFolders(_folderTree.Nodes).ToList();
        if (checkedFolders.Count == 0)
        {
            MessageBox.Show("Check at least one folder to search.", "No folders selected",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var opts = new SearchOptions(
            NullIfEmpty(_kwBox.Text),
            NullIfEmpty(_subjectBox.Text),
            NullIfEmpty(_fromBox.Text),
            _dateFromPicker.Checked ? _dateFromPicker.Value.Date : null,
            _dateToPicker.Checked ? _dateToPicker.Value.Date : null
        );

        ResetCts();
        var token = _cts.Token;
        SetBusy($"Searching {checkedFolders.Count} folder(s)…", cancelable: true);
        try
        {
            _results.Clear();
            _resultList.Items.Clear();
            _bodyPreview.Clear();

            var outcome = await _svc.SearchAsync(checkedFolders, opts, token);
            _results.AddRange(outcome.Items);

            _resultList.BeginUpdate();
            foreach (var r in _results)
            {
                var item = new ListViewItem(r.Subject);
                item.SubItems.AddRange([
                    r.From,
                    r.ReceivedAt?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    r.FolderName,
                    r.HasAttachments ? "Y" : ""
                ]);
                item.Tag = r;
                _resultList.Items.Add(item);
            }
            _resultList.EndUpdate();

            _countLabel.Text = outcome.Truncated ? $"{_results.Count}+ results" : $"{_results.Count} result(s)";
            _countLabel.ForeColor = _results.Count > 0 ? Color.DarkGreen : Color.Gray;

            var msg = token.IsCancellationRequested
                ? $"Cancelled — {_results.Count} email(s) found so far in {outcome.FoldersSearched} folder(s)."
                : $"Found {_results.Count} email(s) in {outcome.FoldersSearched} folder(s).";
            if (outcome.FoldersFailed > 0)
                msg += $" {outcome.FoldersFailed} folder(s) skipped (no access).";
            if (outcome.Truncated)
                msg += $" Result limit ({_svc.MaxResults}) reached — narrow your search (add a date range or keyword) to see the rest.";
            SetStatus(msg);
        }
        catch (OperationCanceledException) { SetStatus("Search cancelled."); }
        catch (Exception ex) { SetStatus($"Search error: {ex.Message}"); }
        finally { SetIdle(); }
    }

    private async void OnResultSelected(object? s, EventArgs e)
    {
        if (_resultList.SelectedItems.Count == 0 || _svc is null) return;
        var result = (EmailResult)_resultList.SelectedItems[0].Tag!;
        var seq = ++_previewSeq;   // this becomes the newest request

        _bodyPreview.Clear();
        _bodyPreview.SelectionFont = new Font("Segoe UI", 9, FontStyle.Bold);
        _bodyPreview.AppendText($"Subject:  {result.Subject}\n");
        _bodyPreview.SelectionFont = new Font("Segoe UI", 9);
        _bodyPreview.AppendText($"From:     {result.From}\n");
        _bodyPreview.AppendText($"To:       {result.To}\n");
        _bodyPreview.AppendText($"Received: {result.ReceivedAt?.ToString("yyyy-MM-dd HH:mm")}\n");
        _bodyPreview.AppendText($"Folder:   {result.FolderName}\n");
        _bodyPreview.SelectionColor = Color.LightGray;
        _bodyPreview.AppendText(new string('-', 90) + "\n");
        _bodyPreview.SelectionColor = _bodyPreview.ForeColor;

        string body;
        try
        {
            body = await _svc.GetBodyAsync(result.EntryId, result.StoreId);
            if (string.IsNullOrWhiteSpace(body)) body = result.BodyPreview;
            body = Regex.Replace(body, @"\n{3,}", "\n\n").Trim();
        }
        catch
        {
            body = result.BodyPreview;
        }

        // A newer selection came in while we were fetching — discard this stale body
        // so it can't land under the wrong headers.
        if (seq != _previewSeq) return;

        _bodyPreview.AppendText(body);
        _bodyPreview.SelectionStart = 0;
        _bodyPreview.ScrollToCaret();
    }

    // ── Helpers ───────────────────────────────────────────────────

    private static IEnumerable<(string EntryId, string StoreId, string Name)> GetCheckedFolders(TreeNodeCollection nodes)
    {
        foreach (TreeNode node in nodes)
        {
            if (node.Checked && node.Tag is MailFolderNode f && !string.IsNullOrEmpty(f.EntryId))
                yield return (f.EntryId, f.StoreId, f.Name);
            foreach (var child in GetCheckedFolders(node.Nodes))
                yield return child;
        }
    }

    private static void SetAllChecked(TreeNodeCollection nodes, bool value)
    {
        foreach (TreeNode node in nodes)
        {
            node.Checked = value;
            SetAllChecked(node.Nodes, value);
        }
    }

    private static Label RightLabel(string text) => new()
    {
        Text = text,
        Dock = DockStyle.Fill,
        TextAlign = ContentAlignment.MiddleRight,
        Padding = new Padding(0, 0, 4, 0)
    };

    private static string? NullIfEmpty(string s) =>
        string.IsNullOrWhiteSpace(s) ? null : s.Trim();

    private string? ShowInputDialog(string title, string prompt)
    {
        using var dlg = new Form
        {
            Text = title,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            StartPosition = FormStartPosition.CenterParent,
            ClientSize = new Size(420, 130),
            MaximizeBox = false,
            MinimizeBox = false
        };
        var label = new Label { Text = prompt, AutoSize = false, Dock = DockStyle.Top, Height = 48, Padding = new Padding(10, 10, 10, 0) };
        var textBox = new TextBox { Left = 12, Top = 60, Width = 396, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top };
        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Left = 246, Top = 92, Width = 75 };
        var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 330, Top = 92, Width = 75 };
        dlg.Controls.AddRange([label, textBox, ok, cancel]);
        dlg.AcceptButton = ok;
        dlg.CancelButton = cancel;
        return dlg.ShowDialog(this) == DialogResult.OK ? textBox.Text : null;
    }

    private void ResetCts()
    {
        _cts.Cancel();
        _cts.Dispose();
        _cts = new CancellationTokenSource();
    }

    private void SetBusy(string msg, bool cancelable = false)
    {
        _statusLabel.Text = msg;
        _progressBar.Visible = true;
        _progressBar.Style = ProgressBarStyle.Marquee;
        _refreshBtn.Enabled = false;
        _openMailboxBtn.Enabled = false;
        _searchBtn.Enabled = false;
        _clearAllBtn.Enabled = false;
        _exportBtn.Enabled = false;
        _cancelBtn.Enabled = cancelable;
    }

    private void SetIdle()
    {
        _progressBar.Visible = false;
        _cancelBtn.Enabled = false;
        _refreshBtn.Enabled = _svc is not null;
        _openMailboxBtn.Enabled = _svc is not null;
        _searchBtn.Enabled = _folderTree.Nodes.Count > 0;
        _clearAllBtn.Enabled = _svc is not null;
        _exportBtn.Enabled = _results.Count > 0;
    }

    // Resets the search to a clean slate for a fresh query. Folder selections are
    // intentionally left checked so a new search can run on the same folders.
    private void OnClearAll(object? s, EventArgs e)
    {
        _kwBox.Clear();
        _subjectBox.Clear();
        _fromBox.Clear();
        _dateFromPicker.Checked = false;
        _dateToPicker.Checked = false;
        UpdateClearDatesState();

        _results.Clear();
        _resultList.Items.Clear();
        _bodyPreview.Clear();
        _countLabel.Text = "0 results";
        _countLabel.ForeColor = Color.Gray;
        _exportBtn.Enabled = false;
        SetStatus("Cleared. Enter new criteria and Search Mail.");
        _kwBox.Focus();
    }

    // Exports the selected results (or all of them, if none are selected) to a single
    // file in the format chosen in the Save dialog.
    private async void OnExport(object? s, EventArgs e)
    {
        if (_svc is null || _results.Count == 0) return;

        var items = _resultList.SelectedItems.Count > 0
            ? _resultList.SelectedItems.Cast<ListViewItem>().Select(i => (EmailResult)i.Tag!).ToList()
            : _results.ToList();

        string path;
        using (var dlg = new SaveFileDialog
        {
            Title = $"Export {items.Count} email(s)",
            FileName = $"Email export {DateTime.Now:yyyy-MM-dd}",
            Filter = ExportService.FileFilter,
            DefaultExt = "txt",
            OverwritePrompt = true
        })
        {
            if (dlg.ShowDialog(this) != DialogResult.OK) return;
            path = dlg.FileName;
        }

        ResetCts();
        var token = _cts.Token;
        SetBusy($"Preparing {items.Count} email(s) for export…", cancelable: true);
        try
        {
            // Fetch each message body (results carry only metadata).
            var rows = new List<ExportRow>(items.Count);
            int done = 0;
            foreach (var m in items)
            {
                if (token.IsCancellationRequested) break;
                string body = "";
                try { body = await _svc.GetBodyAsync(m.EntryId, m.StoreId); } catch { }
                rows.Add(new ExportRow(
                    m.ReceivedAt?.ToString("yyyy-MM-dd HH:mm") ?? "",
                    m.From, m.To, m.Subject, m.FolderName, body));
                if (++done % 10 == 0) SetStatus($"Preparing… {done}/{items.Count}");
            }

            if (token.IsCancellationRequested) { SetStatus("Export cancelled."); return; }

            SetStatus($"Writing {rows.Count} email(s) to {Path.GetFileName(path)}…");
            await _export.ExportAsync(path, rows, token);

            SetStatus($"Exported {rows.Count} email(s) to {path}");
            if (MessageBox.Show("Export complete. Open the file now?", "Export",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            SetStatus($"Export failed: {ex.Message}");
            MessageBox.Show(ex.Message, "Export failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        finally { SetIdle(); }
    }

    private void UpdateClearDatesState() =>
        _clearDatesBtn.Enabled = _dateFromPicker.Checked || _dateToPicker.Checked;

    private void SetStatus(string msg) => _statusLabel.Text = msg;

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _cts.Cancel();
        _cts.Dispose();
        _fromAutoComplete?.Dispose();
        _export.Dispose();
        _svc?.Dispose();
        base.OnFormClosed(e);
    }
}
