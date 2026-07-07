using System.Net.NetworkInformation;
using System.Windows.Forms;

Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.Run(new MainForm());

class MainForm : Form
{
    readonly DataGridView grid = new();
    readonly Label lastRefreshedLabel = new();
    readonly ComboBox intervalCombo = new();
    readonly Button refreshButton = new();
    readonly System.Windows.Forms.Timer pollTimer = new();

    public MainForm()
    {
        Text = "Ethernet Port Monitor";
        Size = new(1100, 440);
        MinimumSize = new(700, 300);
        StartPosition = FormStartPosition.CenterScreen;

        BuildToolbar();
        BuildGrid();
        BuildStatusBar();

        pollTimer.Tick += (_, _) => Refresh();
        SetInterval(5000);
        Refresh();
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        var intervalLabel = new Label
        {
            Text = "Poll interval:",
            AutoSize = true,
            Location = new(8, 14),
        };

        intervalCombo.Items.AddRange(["1 second", "5 seconds", "10 seconds", "30 seconds", "60 seconds"]);
        intervalCombo.SelectedIndex = 1;
        intervalCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        intervalCombo.Location = new(95, 10);
        intervalCombo.Width = 120;
        intervalCombo.SelectedIndexChanged += (_, _) =>
        {
            int[] ms = [1000, 5000, 10000, 30000, 60000];
            SetInterval(ms[intervalCombo.SelectedIndex]);
        };

        refreshButton.Text = "Refresh Now";
        refreshButton.Location = new(230, 9);
        refreshButton.Width = 110;
        refreshButton.Height = 28;
        refreshButton.Click += (_, _) => Refresh();

        toolbar.Controls.AddRange([intervalLabel, intervalCombo, refreshButton]);
        Controls.Add(toolbar);
    }

    void BuildGrid()
    {
        grid.Dock = DockStyle.Fill;
        grid.ReadOnly = true;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.RowHeadersVisible = false;
        grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        grid.BackgroundColor = SystemColors.Window;
        grid.BorderStyle = BorderStyle.None;
        grid.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
        grid.GridColor = Color.FromArgb(220, 220, 220);
        grid.Font = new Font("Segoe UI", 9.5f);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 210, 255);
        grid.DefaultCellStyle.SelectionForeColor = Color.Black;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        grid.ColumnHeadersHeight = 30;

        foreach (var (name, fillWeight) in new (string, int)[]
        {
            ("Name", 25),
            ("Status", 10),
            ("Speed (Mbps)", 12),
            ("MAC Address", 18),
            ("IP Addresses", 35),
        })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = name,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FillWeight = fillWeight,
            });
        }

        Controls.Add(grid);
    }

    void BuildStatusBar()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = SystemColors.ControlLight };
        lastRefreshedLabel.AutoSize = true;
        lastRefreshedLabel.Location = new(8, 5);
        lastRefreshedLabel.ForeColor = SystemColors.ControlDarkDark;
        bar.Controls.Add(lastRefreshedLabel);
        Controls.Add(bar);
    }

    void SetInterval(int ms)
    {
        pollTimer.Stop();
        pollTimer.Interval = ms;
        pollTimer.Start();
    }

    new void Refresh()
    {
        var nics = NetworkInterface.GetAllNetworkInterfaces()
            .Where(n => n.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
            .OrderBy(n => n.Name)
            .ToList();

        grid.Rows.Clear();

        foreach (var nic in nics)
        {
            var props = nic.GetIPProperties();
            var ips = props.UnicastAddresses
                .Select(a => a.Address.ToString())
                .Where(a => !a.StartsWith("fe80", StringComparison.OrdinalIgnoreCase))
                .ToList();

            var mac = nic.GetPhysicalAddress().ToString();
            var macFormatted = mac.Length == 12
                ? string.Join(":", Enumerable.Range(0, 6).Select(i => mac.Substring(i * 2, 2)))
                : mac;

            var speedMbps = nic.Speed > 0 ? (nic.Speed / 1_000_000).ToString() : "—";
            var ipText = ips.Count > 0 ? string.Join("  |  ", ips) : "—";

            int row = grid.Rows.Add(nic.Name, nic.OperationalStatus.ToString(), speedMbps, macFormatted, ipText);

            grid.Rows[row].DefaultCellStyle.BackColor = nic.OperationalStatus == OperationalStatus.Up
                ? Color.FromArgb(220, 245, 220)
                : Color.FromArgb(245, 220, 220);
        }

        lastRefreshedLabel.Text = $"Last refreshed: {DateTime.Now:HH:mm:ss}  |  {nics.Count} interface(s) found";
    }
}
