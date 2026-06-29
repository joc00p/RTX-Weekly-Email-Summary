using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RTXReporter;

public class MainForm : Form
{
    private readonly OutlookService _outlook = new();
    private readonly OllamaService _ollama = new();

    private Dictionary<string, List<EmailItem>> _emailsByWeek = new();
    private readonly Dictionary<string, string> _reportCache = new();

    private ListBox _weekList = null!;
    private RichTextBox _reportBox = null!;
    private Button _generateBtn = null!;
    private Button _copyBtn = null!;
    private Button _saveBtn = null!;
    private Label _statusLabel = null!;
    private ProgressBar _progress = null!;
    private CancellationTokenSource? _cts;

    public MainForm()
    {
        BuildUI();
    }

    private void BuildUI()
    {
        Text = "RTX Weekly Reporter";
        Size = new Size(1100, 750);
        MinimumSize = new Size(800, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10f);
        BackColor = Color.FromArgb(240, 240, 240);

        // Top toolbar
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 50 };
        toolbar.BackColor = Color.FromArgb(245, 245, 245);

        _generateBtn = MakeButton("Generate Report", Color.FromArgb(16, 137, 62));
        _generateBtn.Enabled = false;
        _generateBtn.Click += GenerateReport_Click;

        _copyBtn = MakeButton("Copy", Color.FromArgb(100, 100, 100));
        _copyBtn.Enabled = false;
        _copyBtn.Click += (_, _) => Clipboard.SetText(_reportBox.Text);

        _saveBtn = MakeButton("Save As...", Color.FromArgb(100, 100, 100));
        _saveBtn.Enabled = false;
        _saveBtn.Click += SaveReport_Click;

        int x = 10;
        foreach (var btn in new[] { _generateBtn, _copyBtn, _saveBtn })
        {
            btn.Left = x;
            btn.Top = 10;
            toolbar.Controls.Add(btn);
            x += btn.Width + 6;
        }

        // Status bar
        var statusBar = new Panel { Dock = DockStyle.Bottom, Height = 28 };
        statusBar.BackColor = Color.FromArgb(225, 225, 225);

        _statusLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            ForeColor = Color.FromArgb(60, 60, 60),
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8, 0, 0, 0),
            Text = "Ready — click Load Emails to begin."
        };

        _progress = new ProgressBar
        {
            Dock = DockStyle.Right,
            Width = 160,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
            Visible = false
        };

        statusBar.Controls.Add(_statusLabel);
        statusBar.Controls.Add(_progress);

        // Split container
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(240, 240, 240),
        };
        Shown += (_, _) => {
            int dist = Math.Max(split.Panel1MinSize, Math.Min(230, split.Width - split.Panel2MinSize - split.SplitterWidth));
            split.SplitterDistance = dist;
            LoadEmails_Click(null, EventArgs.Empty);
        };

        // Left panel: week list
        var leftHeader = new Label
        {
            Text = "  WEEKS",
            Dock = DockStyle.Top,
            Height = 28,
            BackColor = Color.FromArgb(210, 220, 235),
            ForeColor = Color.FromArgb(30, 60, 100),
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };

        _weekList = new ListBox
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(30, 30, 30),
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 9.5f),
            ItemHeight = 28,
            SelectionMode = SelectionMode.MultiExtended,
        };
        _weekList.SelectedIndexChanged += WeekList_SelectedIndexChanged;

        split.Panel1.Controls.Add(_weekList);
        split.Panel1.Controls.Add(leftHeader);

        // Right panel: report text
        _reportBox = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(30, 30, 30),
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 10f),
            ReadOnly = true,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            WordWrap = true,
            Padding = new Padding(12),
        };

        split.Panel2.Controls.Add(_reportBox);

        Controls.Add(split);
        Controls.Add(toolbar);
        Controls.Add(statusBar);
    }

    private static Button MakeButton(string text, Color backColor) => new()
    {
        Text = text,
        AutoSize = false,
        Width = 135,
        Height = 30,
        BackColor = backColor,
        ForeColor = Color.White,
        FlatStyle = FlatStyle.Flat,
        Font = new Font("Segoe UI", 9.5f),
        Cursor = Cursors.Hand,
    };

    private async void LoadEmails_Click(object? sender, EventArgs e)
    {
        SetBusy(true, "Loading emails from Outlook...");
        _weekList.Items.Clear();
        _reportBox.Clear();
        _generateBtn.Enabled = false;
        _copyBtn.Enabled = false;
        _saveBtn.Enabled = false;
        _reportCache.Clear();


        try
        {
            _emailsByWeek = await Task.Run(() => _outlook.FetchEmailsByWeek());
            foreach (var week in _emailsByWeek.Keys)
                _weekList.Items.Add(week);

            int total = 0;
            foreach (var v in _emailsByWeek.Values) total += v.Count;
            SetStatus($"Loaded {total} email(s) across {_emailsByWeek.Count} week(s). Select a week then click Generate Report.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to load emails:\n\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Error loading emails.");
        }
        finally
        {
            SetBusy(false, "");
        }
    }

    private void WeekList_SelectedIndexChanged(object? sender, EventArgs e)
    {
        int count = _weekList.SelectedItems.Count;
        if (count == 0) return;
        _generateBtn.Enabled = true;

        // If exactly one week selected and cached, show full report
        if (count == 1 && _weekList.SelectedItems[0] is string week && _reportCache.TryGetValue(week, out var cached))
        {
            ShowReport(cached);
            SetStatus($"Cached report — {week}.");
            return;
        }

        // Otherwise show email previews for all selected weeks
        var selectedWeeks = new List<string>();
        foreach (var item in _weekList.SelectedItems)
            if (item is string w) selectedWeeks.Add(w);

        ShowEmailPreviews(selectedWeeks);
    }

    private void ShowEmailPreviews(List<string> weeks)
    {
        _reportBox.Clear();
        _reportBox.ForeColor = Color.FromArgb(30, 30, 30);

        foreach (var week in weeks)
        {
            if (!_emailsByWeek.TryGetValue(week, out var emails)) continue;

            // Week header
            _reportBox.SelectionFont = new Font("Segoe UI", 11f, FontStyle.Bold);
            _reportBox.SelectionColor = Color.FromArgb(0, 84, 166);
            _reportBox.AppendText($"{week}  ({emails.Count} email{(emails.Count != 1 ? "s" : "")})\n");

            foreach (var email in emails)
            {
                // Email subject line + date
                _reportBox.SelectionFont = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                _reportBox.SelectionColor = Color.FromArgb(30, 30, 30);
                _reportBox.AppendText($"  {email.From}  —  {email.Subject}\n");
                _reportBox.SelectionFont = new Font("Segoe UI", 8.5f, FontStyle.Italic);
                _reportBox.SelectionColor = Color.FromArgb(120, 120, 120);
                _reportBox.AppendText($"    Received: {email.Received}\n");

                // First 4 non-empty lines of body
                var lines = email.Body.Split('\n');
                int shown = 0;
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrWhiteSpace(trimmed)) continue;
                    _reportBox.SelectionFont = new Font("Segoe UI", 9f, FontStyle.Regular);
                    _reportBox.SelectionColor = Color.FromArgb(80, 80, 80);
                    _reportBox.AppendText($"    {trimmed}\n");
                    if (++shown >= 4) break;
                }

                _reportBox.AppendText("\n");
            }

            _reportBox.AppendText("\n");
        }

        _reportBox.SelectionStart = 0;
        _reportBox.SelectionLength = 0;

        string label = weeks.Count == 1 ? weeks[0] : $"{weeks.Count} weeks";
        SetStatus($"Previewing {label} — click Generate Report to build the AI summary.");
    }

    private async void GenerateReport_Click(object? sender, EventArgs e)
    {
        var selectedWeeks = new List<string>();
        foreach (var item in _weekList.SelectedItems)
            if (item is string w) selectedWeeks.Add(w);
        if (selectedWeeks.Count == 0) return;

        _cts?.Cancel();
        _cts = new CancellationTokenSource();

        // Combine emails from all selected weeks
        var allEmails = new List<EmailItem>();
        foreach (var w in selectedWeeks)
            if (_emailsByWeek.TryGetValue(w, out var emails)) allEmails.AddRange(emails);

        string weekLabel = selectedWeeks.Count == 1
            ? selectedWeeks[0]
            : $"{selectedWeeks[^1]} through {selectedWeeks[0]}";

        SetBusy(true, $"Generating report for {weekLabel} ({allEmails.Count} email(s))...");
        _generateBtn.Enabled = false;

        try
        {
            var report = await _ollama.SummarizeWeekAsync(weekLabel, allEmails, _cts.Token);
            // Cache under a combined key
            string cacheKey = string.Join("|", selectedWeeks);
            _reportCache[cacheKey] = report;
            ShowReport(report);
            _copyBtn.Enabled = true;
            _saveBtn.Enabled = true;
            SetStatus($"Report ready — {weekLabel}.");
        }
        catch (OperationCanceledException)
        {
            SetStatus("Generation cancelled.");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Ollama error:\n\n{ex.Message}\n\nMake sure Ollama is running on localhost:11434.", "Ollama Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            SetStatus("Error generating report.");
        }
        finally
        {
            SetBusy(false, "");
            _generateBtn.Enabled = true;
        }
    }

    private void ShowReport(string text)
    {
        _reportBox.Clear();
        _reportBox.ForeColor = Color.FromArgb(30, 30, 30);
        _reportBox.Text = text;

        // Highlight ## headers and **bold** lines in blue
        int pos = 0;
        string content = _reportBox.Text;
        while (pos < content.Length)
        {
            int lineEnd = content.IndexOf('\n', pos);
            if (lineEnd < 0) lineEnd = content.Length;
            var line = content[pos..lineEnd].TrimStart();

            if (line.StartsWith("##") || line.StartsWith("**"))
            {
                _reportBox.Select(pos, lineEnd - pos);
                _reportBox.SelectionFont = new Font("Consolas", 10f, FontStyle.Bold);
                _reportBox.SelectionColor = Color.FromArgb(0, 84, 166);
            }

            pos = lineEnd + 1;
        }

        _reportBox.SelectionStart = 0;
        _reportBox.SelectionLength = 0;
    }

    private void SaveReport_Click(object? sender, EventArgs e)
    {
        var week = _weekList.SelectedItems.Count > 0 ? _weekList.SelectedItems[0] as string ?? "report" : "report";
        var safe = string.Join("_", week.Split(System.IO.Path.GetInvalidFileNameChars()));
        using var dlg = new SaveFileDialog
        {
            FileName = $"RTX_{safe}.txt",
            Filter = "Text File|*.txt|All Files|*.*",
            Title = "Save Report"
        };
        if (dlg.ShowDialog() == DialogResult.OK)
            System.IO.File.WriteAllText(dlg.FileName, _reportBox.Text);
    }

    private void SetBusy(bool busy, string msg)
    {
        _progress.Visible = busy;
        if (!string.IsNullOrEmpty(msg)) SetStatus(msg);
    }

    private void SetStatus(string msg) => _statusLabel.Text = msg;
}
