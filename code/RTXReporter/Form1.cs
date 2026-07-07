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
    private readonly OllamaService _ollama;
    private readonly TeamConfig _teamConfig = new();
    private readonly PowerPointService _pptx;

    private Dictionary<string, List<EmailItem>> _emailsByWeek = new();
    private readonly Dictionary<string, string> _reportCache = new();
    private string _lastWeekLabel = "";

    private ListBox _weekList = null!;
    private RichTextBox _reportBox = null!;
    private Button _generateBtn = null!;
    private Button _copyBtn = null!;
    private Button _saveBtn = null!;
    private Button _pptxBtn = null!;
    private ThemeToggle _themeToggle = null!;
    private Label _statusLabel = null!;
    private Label _leftHeader = null!;
    private Panel _toolbar = null!;
    private Panel _statusBar = null!;
    private SplitContainer _split = null!;
    private ProgressBar _progress = null!;
    private CancellationTokenSource? _cts;
    private bool _isDark = false;

    // Light theme
    private static readonly Color LightBg         = Color.FromArgb(240, 240, 240);
    private static readonly Color LightToolbar     = Color.FromArgb(245, 245, 245);
    private static readonly Color LightStatusBar   = Color.FromArgb(225, 225, 225);
    private static readonly Color LightStatusText  = Color.FromArgb(60, 60, 60);
    private static readonly Color LightListBg      = Color.White;
    private static readonly Color LightListFg      = Color.FromArgb(30, 30, 30);
    private static readonly Color LightHeaderBg    = Color.FromArgb(210, 220, 235);
    private static readonly Color LightHeaderFg    = Color.FromArgb(30, 60, 100);
    private static readonly Color LightReportBg    = Color.White;
    private static readonly Color LightReportFg    = Color.FromArgb(30, 30, 30);
    private static readonly Color LightAccent      = Color.FromArgb(0, 84, 166);

    // Dark theme
    private static readonly Color DarkBg           = Color.FromArgb(22, 22, 22);
    private static readonly Color DarkToolbar      = Color.FromArgb(30, 30, 30);
    private static readonly Color DarkStatusBar    = Color.FromArgb(40, 40, 40);
    private static readonly Color DarkStatusText   = Color.Silver;
    private static readonly Color DarkListBg       = Color.FromArgb(28, 28, 28);
    private static readonly Color DarkListFg       = Color.Gainsboro;
    private static readonly Color DarkHeaderBg     = Color.FromArgb(38, 38, 38);
    private static readonly Color DarkHeaderFg     = Color.Silver;
    private static readonly Color DarkReportBg     = Color.FromArgb(18, 18, 18);
    private static readonly Color DarkReportFg     = Color.FromArgb(220, 220, 220);
    private static readonly Color DarkAccent       = Color.FromArgb(100, 180, 255);

    public MainForm()
    {
        _pptx = new PowerPointService(
            Path.Combine(AppContext.BaseDirectory, "RTXReport", "RTX TEMPLATE.pptx"),
            _teamConfig);
        _ollama = new OllamaService();
        _ollama.StatusUpdate += msg => Invoke(() => SetStatus(msg));
        BuildUI();
    }

    private void BuildUI()
    {
        Text = "RTX Weekly Reporter";
        Size = new Size(1100, 750);
        MinimumSize = new Size(800, 600);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 10f);

        _toolbar = new Panel { Dock = DockStyle.Top, Height = 50 };

        var reloadBtn = MakeButton("Reload", Color.FromArgb(70, 130, 180));
        reloadBtn.Click += LoadEmails_Click;

        _generateBtn = MakeButton("Generate Report", Color.FromArgb(16, 137, 62));
        _generateBtn.Enabled = false;
        _generateBtn.Click += GenerateReport_Click;

        _copyBtn = MakeButton("Copy", Color.FromArgb(70, 130, 180));
        _copyBtn.Enabled = false;
        _copyBtn.Click += (_, _) => Clipboard.SetText(_reportBox.Text);

        _saveBtn = MakeButton("Save As...", Color.FromArgb(70, 130, 180));
        _saveBtn.Enabled = false;
        _saveBtn.Click += SaveReport_Click;

        _pptxBtn = MakeButton("Export PPTX", Color.FromArgb(180, 80, 20));
        _pptxBtn.Enabled = false;
        _pptxBtn.Click += ExportPptx_Click;

        var teamsBtn = MakeButton("Manage Teams", Color.FromArgb(100, 70, 160));
        teamsBtn.Click += ManageTeams_Click;

        _themeToggle = new ThemeToggle { Top = 10 };
        _themeToggle.ThemeChanged += ThemeBtn_Click;

        int x = 10;
        foreach (var btn in new[] { reloadBtn, _generateBtn, _copyBtn, _saveBtn, _pptxBtn, teamsBtn })
        {
            btn.Left = x;
            btn.Top = 10;
            _toolbar.Controls.Add(btn);
            x += btn.Width + 6;
        }
        _toolbar.Controls.Add(_themeToggle);
        _toolbar.Resize += (_, _) => _themeToggle.Left = _toolbar.Width - _themeToggle.Width - 10;

        _statusBar = new Panel { Dock = DockStyle.Bottom, Height = 28 };

        _statusLabel = new Label
        {
            AutoSize = false,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding(8, 0, 0, 0),
            Text = "Loading emails..."
        };

        _progress = new ProgressBar
        {
            Dock = DockStyle.Right,
            Width = 160,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30,
            Visible = false
        };

        _statusBar.Controls.Add(_statusLabel);
        _statusBar.Controls.Add(_progress);

        _split = new SplitContainer { Dock = DockStyle.Fill };

        Shown += (_, _) =>
        {
            int dist = Math.Max(_split.Panel1MinSize, Math.Min(230, _split.Width - _split.Panel2MinSize - _split.SplitterWidth));
            _split.SplitterDistance = dist;
            _themeToggle.Left = _toolbar.Width - _themeToggle.Width - 10;
            LoadEmails_Click(null, EventArgs.Empty);
        };

        _leftHeader = new Label
        {
            Text = "  WEEKS",
            Dock = DockStyle.Top,
            Height = 28,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };

        _weekList = new ListBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 9.5f),
            ItemHeight = 28,
            SelectionMode = SelectionMode.MultiExtended,
        };
        _weekList.SelectedIndexChanged += WeekList_SelectedIndexChanged;

        _split.Panel1.Controls.Add(_weekList);
        _split.Panel1.Controls.Add(_leftHeader);

        var reportContextMenu = new ContextMenuStrip();
        var copyMenuItem = new ToolStripMenuItem("Copy");
        copyMenuItem.Click += (_, _) => { if (_reportBox.SelectionLength > 0) Clipboard.SetText(_reportBox.SelectedText); };
        var selectAllMenuItem = new ToolStripMenuItem("Select All");
        selectAllMenuItem.Click += (_, _) => _reportBox.SelectAll();
        reportContextMenu.Items.Add(copyMenuItem);
        reportContextMenu.Items.Add(new ToolStripSeparator());
        reportContextMenu.Items.Add(selectAllMenuItem);
        reportContextMenu.Opening += (_, _) => copyMenuItem.Enabled = _reportBox.SelectionLength > 0;

        _reportBox = new RichTextBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 10f),
            ReadOnly = false,
            ScrollBars = RichTextBoxScrollBars.Vertical,
            WordWrap = true,
            Padding = new Padding(12),
            ContextMenuStrip = reportContextMenu,
        };

        _split.Panel2.Controls.Add(_reportBox);

        Controls.Add(_split);
        Controls.Add(_toolbar);
        Controls.Add(_statusBar);

        ApplyTheme();
    }

    private void ThemeBtn_Click(object? sender, EventArgs e)
    {
        _isDark = _themeToggle.IsDark;
        ApplyTheme();

        // Redraw current content with new accent color
        if (_weekList.SelectedItems.Count > 0)
            WeekList_SelectedIndexChanged(null, EventArgs.Empty);
    }

    private void ApplyTheme()
    {
        var bg         = _isDark ? DarkBg         : LightBg;
        var toolbar    = _isDark ? DarkToolbar     : LightToolbar;
        var statusBar  = _isDark ? DarkStatusBar   : LightStatusBar;
        var statusText = _isDark ? DarkStatusText  : LightStatusText;
        var listBg     = _isDark ? DarkListBg      : LightListBg;
        var listFg     = _isDark ? DarkListFg      : LightListFg;
        var headerBg   = _isDark ? DarkHeaderBg    : LightHeaderBg;
        var headerFg   = _isDark ? DarkHeaderFg    : LightHeaderFg;
        var reportBg   = _isDark ? DarkReportBg    : LightReportBg;
        var reportFg   = _isDark ? DarkReportFg    : LightReportFg;

        BackColor              = bg;
        _toolbar.BackColor     = toolbar;
        _statusBar.BackColor   = statusBar;
        _statusLabel.ForeColor = statusText;
        _split.BackColor       = bg;
        _weekList.BackColor    = listBg;
        _weekList.ForeColor    = listFg;
        _leftHeader.BackColor  = headerBg;
        _leftHeader.ForeColor  = headerFg;
        _reportBox.BackColor   = reportBg;
        _reportBox.ForeColor   = reportFg;

        _themeToggle.IsDark = _isDark;
    }

    private Color Accent => _isDark ? DarkAccent : LightAccent;
    private Color PreviewFg => _isDark ? Color.FromArgb(180, 180, 180) : Color.FromArgb(80, 80, 80);
    private Color SubtleFg => _isDark ? Color.FromArgb(130, 130, 130) : Color.FromArgb(120, 120, 120);

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

        var selectedWeeks = new List<string>();
        foreach (var item in _weekList.SelectedItems)
            if (item is string w) selectedWeeks.Add(w);

        _lastWeekLabel = selectedWeeks.Count == 1
            ? selectedWeeks[0]
            : $"{selectedWeeks[^1]} through {selectedWeeks[0]}";

        string cacheKey = string.Join("|", selectedWeeks);
        if (_reportCache.TryGetValue(cacheKey, out var cached))
        {
            ShowReport(cached);
            SetStatus($"Cached report — {_lastWeekLabel}.");
            return;
        }

        ShowEmailPreviews(selectedWeeks);
    }

    private void ShowEmailPreviews(List<string> weeks)
    {
        _reportBox.Clear();
        _reportBox.ForeColor = _isDark ? DarkReportFg : LightReportFg;

        foreach (var week in weeks)
        {
            if (!_emailsByWeek.TryGetValue(week, out var emails)) continue;

            _reportBox.SelectionFont = new Font("Segoe UI", 11f, FontStyle.Bold);
            _reportBox.SelectionColor = Accent;
            _reportBox.AppendText($"{week}  ({emails.Count} email{(emails.Count != 1 ? "s" : "")})\n");

            foreach (var email in emails)
            {
                _reportBox.SelectionFont = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                _reportBox.SelectionColor = _isDark ? DarkReportFg : LightReportFg;
                _reportBox.AppendText($"  {email.From}  —  {email.Subject}\n");
                _reportBox.SelectionFont = new Font("Segoe UI", 8.5f, FontStyle.Italic);
                _reportBox.SelectionColor = SubtleFg;
                _reportBox.AppendText($"    Received: {email.Received}\n");

                var lines = email.Body.Split('\n');
                int shown = 0;
                foreach (var line in lines)
                {
                    var trimmed = line.Trim();
                    if (string.IsNullOrWhiteSpace(trimmed)) continue;
                    _reportBox.SelectionFont = new Font("Segoe UI", 9f, FontStyle.Regular);
                    _reportBox.SelectionColor = PreviewFg;
                    _reportBox.AppendText($"    {trimmed}\n");
                    if (++shown >= 4) break;
                }
                _reportBox.AppendText("\n");
            }
            _reportBox.AppendText("\n");
        }
        _reportBox.AppendText("\n\n\n");

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
            string cacheKey = string.Join("|", selectedWeeks);
            _reportCache[cacheKey] = report;
            ShowReport(report);
            _lastWeekLabel = weekLabel;
            _copyBtn.Enabled = true;
            _saveBtn.Enabled = true;
            _pptxBtn.Enabled = true;
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
        _reportBox.ForeColor = _isDark ? DarkReportFg : LightReportFg;
        _reportBox.Text = text + "\n\n\n";

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
                _reportBox.SelectionFont = new Font("Segoe UI", 10f, FontStyle.Bold);
                _reportBox.SelectionColor = Accent;
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

    private void ExportPptx_Click(object? sender, EventArgs e)
    {
        if (!_pptx.TemplateExists)
        {
            MessageBox.Show(
                $"Template not found. Place your template file at:\n\n{_pptx.TemplatePath}",
                "Template Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var safe = string.Join("_", _lastWeekLabel.Split(System.IO.Path.GetInvalidFileNameChars()));
        using var dlg = new SaveFileDialog
        {
            FileName = $"RTX_{safe}.pptx",
            Filter = "PowerPoint Presentation|*.pptx|All Files|*.*",
            Title = "Export PowerPoint Report"
        };
        if (dlg.ShowDialog() != DialogResult.OK) return;

        try
        {
            _pptx.Export(_lastWeekLabel, _reportBox.Text, dlg.FileName);
            SetStatus($"PPTX exported: {System.IO.Path.GetFileName(dlg.FileName)}");
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to export PPTX:\n\n{ex.Message}", "Export Error",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ManageTeams_Click(object? sender, EventArgs e)
    {
        var senders = _emailsByWeek.Values
            .SelectMany(emails => emails.Select(em => em.From))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s)
            .ToList();

        using var dlg = new ManageTeamsForm(_teamConfig, senders);
        dlg.ShowDialog(this);
    }

    private void SetBusy(bool busy, string msg)
    {
        _progress.Visible = busy;
        if (!string.IsNullOrEmpty(msg)) SetStatus(msg);
    }

    private void SetStatus(string msg) => _statusLabel.Text = msg;
}
