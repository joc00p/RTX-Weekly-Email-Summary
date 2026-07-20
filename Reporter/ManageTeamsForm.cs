using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reporter;

public class ManageTeamsForm : Form
{
    private readonly TeamConfig _config;
    private readonly List<string> _emailSenders;
    private readonly Func<string, List<string>> _searchAddressBook;

    private ComboBox _towerCombo = null!;
    private ListBox _membersBox = null!;
    private ListBox _usersBox = null!;
    private TextBox _searchBox = null!;
    private Label _searchStatus = null!;
    private Button _addBtn = null!;
    private Button _removeBtn = null!;

    private CancellationTokenSource? _searchCts;
    private System.Windows.Forms.Timer _debounce = null!;

    public ManageTeamsForm(TeamConfig config, IEnumerable<string> emailSenders,
        Func<string, List<string>> searchAddressBook)
    {
        _config = config;
        _emailSenders = emailSenders.OrderBy(u => u).ToList();
        _searchAddressBook = searchAddressBook;
        BuildUI();
        PopulateTower();
    }

    private void BuildUI()
    {
        Text = "Manage Team Towers";
        Size = new Size(720, 530);
        MinimumSize = new Size(620, 460);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);

        _debounce = new System.Windows.Forms.Timer { Interval = 350 };
        _debounce.Tick += (_, _) => { _debounce.Stop(); RunSearch(); };

        // Tower selector
        var towerLabel = new Label { Text = "Tower:", AutoSize = true, Location = new Point(12, 16) };
        _towerCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Location = new Point(65, 12),
            Width = 180,
        };
        foreach (var t in TeamConfig.TowerNames) _towerCombo.Items.Add(t);
        _towerCombo.SelectedIndex = 0;
        _towerCombo.SelectedIndexChanged += (_, _) => PopulateTower();

        // Members panel (left)
        var membersLabel = new Label
        {
            Text = "Tower Members",
            AutoSize = true,
            Location = new Point(12, 52),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };
        _membersBox = new ListBox
        {
            Location = new Point(12, 72),
            Size = new Size(260, 360),
            SelectionMode = SelectionMode.MultiExtended,
            Sorted = true,
        };
        _membersBox.SelectedIndexChanged += (_, _) => UpdateButtons();

        // Arrow buttons (center)
        _addBtn = new Button
        {
            Text = "◀ Add",
            Location = new Point(285, 195),
            Size = new Size(95, 34),
            BackColor = Color.FromArgb(16, 137, 62),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        _addBtn.Click += AddBtn_Click;

        _removeBtn = new Button
        {
            Text = "Remove ▶",
            Location = new Point(285, 239),
            Size = new Size(95, 34),
            BackColor = Color.FromArgb(180, 60, 30),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        _removeBtn.Click += RemoveBtn_Click;

        // Outlook users panel (right)
        var usersLabel = new Label
        {
            Text = "Outlook / GAL Users",
            AutoSize = true,
            Location = new Point(395, 52),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };

        _searchBox = new TextBox
        {
            Location = new Point(395, 72),
            Size = new Size(285, 24),
            PlaceholderText = "Search contacts / GAL…",
        };
        _searchBox.TextChanged += (_, _) => { _debounce.Stop(); _debounce.Start(); };

        _searchStatus = new Label
        {
            Location = new Point(395, 100),
            Size = new Size(285, 18),
            Font = new Font("Segoe UI", 8f, FontStyle.Italic),
            ForeColor = Color.Gray,
            Text = "",
        };

        _usersBox = new ListBox
        {
            Location = new Point(395, 120),
            Size = new Size(285, 312),
            SelectionMode = SelectionMode.MultiExtended,
            Sorted = true,
        };
        _usersBox.SelectedIndexChanged += (_, _) => UpdateButtons();

        // Bottom buttons
        var saveBtn = new Button
        {
            Text = "Save",
            Location = new Point(510, 460),
            Size = new Size(85, 30),
            BackColor = Color.FromArgb(0, 84, 166),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        saveBtn.Click += (_, _) =>
        {
            _config.Save();
            MessageBox.Show("Team configuration saved.", "Saved",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        };

        var closeBtn = new Button
        {
            Text = "Close",
            DialogResult = DialogResult.OK,
            Location = new Point(605, 460),
            Size = new Size(85, 30),
            BackColor = Color.FromArgb(100, 100, 100),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };

        Controls.AddRange(new Control[]
        {
            towerLabel, _towerCombo,
            membersLabel, _membersBox,
            _addBtn, _removeBtn,
            usersLabel, _searchBox, _searchStatus, _usersBox,
            saveBtn, closeBtn,
        });
    }

    private void PopulateTower()
    {
        string tower = _towerCombo.SelectedItem as string ?? "";

        _membersBox.BeginUpdate();
        _membersBox.Items.Clear();
        if (_config.Teams.TryGetValue(tower, out var members))
            foreach (var m in members.OrderBy(x => x))
                _membersBox.Items.Add(m);
        _membersBox.EndUpdate();

        // If search box is empty show email senders; otherwise re-run search
        if (string.IsNullOrWhiteSpace(_searchBox?.Text))
            PopulateUsersFromEmailSenders();
        else
            RunSearch();

        UpdateButtons();
    }

    private void PopulateUsersFromEmailSenders()
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        var memberSet = GetCurrentMemberSet();

        _usersBox.BeginUpdate();
        _usersBox.Items.Clear();
        foreach (var user in _emailSenders)
        {
            if (memberSet.Contains(user)) continue;
            string otherTower = _config.GetTeam(user);
            string display = otherTower == "Other" || otherTower == tower
                ? user : $"{user}  ({otherTower})";
            _usersBox.Items.Add(display);
        }
        _usersBox.EndUpdate();
        _searchStatus.Text = $"{_usersBox.Items.Count} recent senders";
    }

    private void RunSearch()
    {
        string query = _searchBox.Text.Trim();
        if (query.Length < 2)
        {
            PopulateUsersFromEmailSenders();
            return;
        }

        _searchCts?.Cancel();
        _searchCts = new CancellationTokenSource();
        var token = _searchCts.Token;

        _searchStatus.Text = "Searching…";
        _usersBox.Items.Clear();

        Task.Run(() => _searchAddressBook(query), token).ContinueWith(t =>
        {
            if (token.IsCancellationRequested) return;
            if (t.IsFaulted)
            {
                _ = t.Exception; // observe to prevent unobserved task exception
                Invoke(() => _searchStatus.Text = "Search failed — Outlook unavailable");
                return;
            }
            var results = t.Result;
            Invoke(() =>
            {
                if (token.IsCancellationRequested) return;
                string tower = _towerCombo.SelectedItem as string ?? "";
                var memberSet = GetCurrentMemberSet();

                _usersBox.BeginUpdate();
                _usersBox.Items.Clear();
                foreach (var user in results)
                {
                    if (memberSet.Contains(user)) continue;
                    string otherTower = _config.GetTeam(user);
                    string display = otherTower == "Other" || otherTower == tower
                        ? user : $"{user}  ({otherTower})";
                    _usersBox.Items.Add(display);
                }
                _usersBox.EndUpdate();
                _searchStatus.Text = _usersBox.Items.Count == 0
                    ? "No results" : $"{_usersBox.Items.Count} result(s)";
                UpdateButtons();
            });
        }, TaskScheduler.Default);
    }

    private HashSet<string> GetCurrentMemberSet()
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        _config.Teams.TryGetValue(tower, out var members);
        return new HashSet<string>(members ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);
    }

    private void AddBtn_Click(object? sender, EventArgs e)
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        var selected = _usersBox.SelectedItems.Cast<string>().ToList();
        foreach (var item in selected)
        {
            string name = item.Contains("  (") ? item[..item.IndexOf("  (")] : item;
            _config.AddMember(tower, name);
        }
        PopulateTower();
    }

    private void RemoveBtn_Click(object? sender, EventArgs e)
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        var selected = _membersBox.SelectedItems.Cast<string>().ToList();
        foreach (var name in selected)
            _config.RemoveMember(tower, name);
        PopulateTower();
    }

    private void UpdateButtons()
    {
        _addBtn.Enabled = _usersBox.SelectedItems.Count > 0;
        _removeBtn.Enabled = _membersBox.SelectedItems.Count > 0;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _searchCts?.Cancel();
        _debounce.Dispose();
        base.OnFormClosed(e);
    }
}
