using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Principal;
using System.Windows.Forms;

Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.Run(new LauncherForm());

// ── Launcher ───────────────────────────────────────────────────────────────

class LauncherForm : Form
{
    public LauncherForm()
    {
        Text = "Network Tools";
        Size = new(620, 340);
        MinimumSize = new(480, 280);
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 246, 248);

        BuildHeader();
        BuildCards();
    }

    void BuildHeader()
    {
        Controls.Add(new Panel
        {
            Dock = DockStyle.Top,
            Height = 48,
            BackColor = Color.FromArgb(40, 40, 60),
            Controls =
            {
                new Label
                {
                    Text = "Network Tools",
                    ForeColor = Color.White,
                    Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                    AutoSize = true,
                    Location = new(16, 12),
                },
            },
        });
    }

    void BuildCards()
    {
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 1,
            Padding = new(16, 14, 16, 14),
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.4f));
        table.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        var entries = new (string title, string sub, string desc, Color accent, Action open)[]
        {
            ("PING",       "a Host",          "Test reachability and measure\nround-trip time to any\nhostname or IP address.",       Color.FromArgb( 55, 135, 255), () => new PingForm().Show()),
            ("TRACEROUTE", "a Host",          "Map each network hop on the\npath to a destination and\nmeasure latency per hop.",     Color.FromArgb( 45, 175,  85), () => new TracerouteForm().Show()),
            ("INTERFACES", "Show Interfaces", "View Ethernet interfaces,\nmonitor live traffic volume,\nand snoop packets.",          Color.FromArgb(155,  75, 220), () => new InterfaceForm().Show()),
        };

        for (int i = 0; i < entries.Length; i++)
        {
            var (title, sub, desc, accent, open) = entries[i];
            var card = MakeCard(title, sub, desc, accent, open);
            card.Margin = new(0, 0, i < entries.Length - 1 ? 8 : 0, 0);
            card.Dock = DockStyle.Fill;
            table.Controls.Add(card, i, 0);
        }

        Controls.Add(table);
    }

    static Panel MakeCard(string title, string sub, string desc, Color accent, Action open)
    {
        var card = new Panel { BackColor = Color.White, Cursor = Cursors.Hand };

        card.Controls.AddRange([
            new Label { Text = title, Font = new Font("Segoe UI", 12f, FontStyle.Bold), ForeColor = accent,                     AutoSize = false, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top, Size = new(10, 26), Location = new(12, 38) },
            new Label { Text = sub,   Font = new Font("Segoe UI",  9f),                 ForeColor = Color.FromArgb(100, 100, 120), AutoSize = false, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top, Size = new(10, 18), Location = new(12, 66) },
            new Label { Text = desc,  Font = new Font("Segoe UI",  8.5f),               ForeColor = Color.FromArgb( 80,  80, 100), AutoSize = false, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top, Size = new(10, 90), Location = new(12, 90) },
        ]);

        card.Resize += (_, _) =>
        {
            foreach (Control c in card.Controls)
                c.Width = card.Width - 24;
        };

        card.Paint += (_, e) =>
        {
            using var br  = new SolidBrush(accent);
            using var pen = new Pen(Color.FromArgb(215, 215, 220));
            e.Graphics.FillRectangle(br,  0, 0, card.Width, 5);
            e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
        };

        void SetBg(Color c) { card.BackColor = c; foreach (Control ctrl in card.Controls) ctrl.BackColor = c; }
        card.MouseEnter += (_, _) => SetBg(Color.FromArgb(245, 248, 255));
        card.MouseLeave += (_, _) => SetBg(Color.White);
        card.Click += (_, _) => open();
        foreach (Control c in card.Controls) c.Click += (_, _) => open();

        return card;
    }
}

// ── Ping ───────────────────────────────────────────────────────────────────

class PingForm : Form
{
    readonly TextBox hostBox = new();
    readonly Button startStopButton = new();
    readonly DataGridView grid = new();
    readonly Label statsLabel = new();

    Thread? pingThread;
    volatile bool pinging;
    int sent, received;
    long totalRtt;

    public PingForm()
    {
        Text = "Ping a Host";
        Size = new(700, 500);
        MinimumSize = new(500, 350);
        StartPosition = FormStartPosition.CenterScreen;

        BuildToolbar();
        BuildGrid();
        BuildStatusBar();
        FormClosing += (_, _) => StopPing();
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        toolbar.Controls.Add(new Label { Text = "Host / IP:", AutoSize = true, Location = new(8, 14) });

        hostBox.Location = new(72, 10);
        hostBox.Width = 240;
        hostBox.Font = new Font("Consolas", 10f);
        hostBox.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) Toggle(); };
        toolbar.Controls.Add(hostBox);

        startStopButton.Text = "Start";
        startStopButton.Location = new(326, 9);
        startStopButton.Width = 80;
        startStopButton.Height = 28;
        startStopButton.BackColor = Color.FromArgb(60, 160, 80);
        startStopButton.ForeColor = Color.White;
        startStopButton.FlatStyle = FlatStyle.Flat;
        startStopButton.Click += (_, _) => Toggle();
        toolbar.Controls.Add(startStopButton);

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
        grid.Font = new Font("Consolas", 9.5f);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        grid.ColumnHeadersHeight = 30;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 210, 255);
        grid.DefaultCellStyle.SelectionForeColor = Color.Black;

        foreach (var (name, weight) in new (string, int)[]
        {
            ("Seq", 8), ("Status", 14), ("RTT (ms)", 12), ("TTL", 8), ("From", 30),
        })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = name,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FillWeight = weight,
            });
        }

        Controls.Add(grid);
    }

    void BuildStatusBar()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = SystemColors.ControlLight };
        statsLabel.AutoSize = true;
        statsLabel.Location = new(8, 5);
        statsLabel.ForeColor = SystemColors.ControlDarkDark;
        bar.Controls.Add(statsLabel);
        Controls.Add(bar);
    }

    void Toggle()
    {
        if (pinging) StopPing();
        else StartPing();
    }

    void StartPing()
    {
        var host = hostBox.Text.Trim();
        if (string.IsNullOrEmpty(host)) return;

        grid.Rows.Clear();
        sent = received = 0;
        totalRtt = 0;
        pinging = true;

        startStopButton.Text = "Stop";
        startStopButton.BackColor = Color.FromArgb(220, 80, 80);
        hostBox.Enabled = false;

        pingThread = new Thread(() => PingLoop(host)) { IsBackground = true };
        pingThread.Start();
    }

    void StopPing()
    {
        pinging = false;
        startStopButton.Text = "Start";
        startStopButton.BackColor = Color.FromArgb(60, 160, 80);
        hostBox.Enabled = true;
    }

    void PingLoop(string host)
    {
        int seq = 1;
        using var pinger = new Ping();
        var buffer = new byte[32];
        var options = new PingOptions { DontFragment = true };

        while (pinging)
        {
            PingReply reply;
            try
            {
                reply = pinger.Send(host, 1000, buffer, options);
            }
            catch (Exception ex)
            {
                int s = seq++;
                BeginInvoke(() => AddRow(s, "Error", -1, 0, ex.Message));
                Thread.Sleep(1000);
                continue;
            }

            int capturedSeq = seq++;
            BeginInvoke(() => AddRow(capturedSeq, reply.Status.ToString(),
                reply.Status == IPStatus.Success ? reply.RoundtripTime : -1,
                reply.Options?.Ttl ?? 0,
                reply.Address?.ToString() ?? ""));

            Thread.Sleep(1000);
        }
    }

    void AddRow(int seq, string status, long rtt, int ttl, string from)
    {
        bool ok = rtt >= 0;
        sent++;
        if (ok) { received++; totalRtt += rtt; }

        if (grid.Rows.Count > 500) grid.Rows.RemoveAt(0);

        int row = grid.Rows.Add(seq, status, ok ? rtt.ToString() : "—", ttl > 0 ? ttl.ToString() : "—", from);
        grid.Rows[row].DefaultCellStyle.BackColor = ok
            ? Color.FromArgb(220, 245, 220)
            : Color.FromArgb(245, 220, 220);
        grid.FirstDisplayedScrollingRowIndex = grid.Rows.Count - 1;

        int lost = sent - received;
        double lostPct = sent > 0 ? (double)lost / sent * 100 : 0;
        long avgRtt = received > 0 ? totalRtt / received : 0;
        statsLabel.Text = $"Sent: {sent}  |  Received: {received}  |  Lost: {lost} ({lostPct:F0}%)  |  Avg RTT: {avgRtt} ms";
    }
}

// ── Traceroute ─────────────────────────────────────────────────────────────

class TracerouteForm : Form
{
    readonly TextBox hostBox = new();
    readonly Button startStopButton = new();
    readonly DataGridView grid = new();
    readonly Label statusLabel = new();

    Thread? traceThread;
    volatile bool tracing;
    int traceSeq;

    public TracerouteForm()
    {
        Text = "Traceroute a Host";
        Size = new(820, 500);
        MinimumSize = new(600, 350);
        StartPosition = FormStartPosition.CenterScreen;

        BuildToolbar();
        BuildGrid();
        BuildStatusBar();
        FormClosing += (_, _) => StopTrace();
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        toolbar.Controls.Add(new Label { Text = "Host / IP:", AutoSize = true, Location = new(8, 14) });

        hostBox.Location = new(72, 10);
        hostBox.Width = 240;
        hostBox.Font = new Font("Consolas", 10f);
        hostBox.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) Toggle(); };
        toolbar.Controls.Add(hostBox);

        startStopButton.Text = "Start";
        startStopButton.Location = new(326, 9);
        startStopButton.Width = 80;
        startStopButton.Height = 28;
        startStopButton.BackColor = Color.FromArgb(60, 160, 80);
        startStopButton.ForeColor = Color.White;
        startStopButton.FlatStyle = FlatStyle.Flat;
        startStopButton.Click += (_, _) => Toggle();
        toolbar.Controls.Add(startStopButton);

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
        grid.Font = new Font("Consolas", 9.5f);
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        grid.ColumnHeadersHeight = 30;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 210, 255);
        grid.DefaultCellStyle.SelectionForeColor = Color.Black;

        foreach (var (name, weight) in new (string, int)[]
        {
            ("Hop", 5), ("IP Address", 18), ("Hostname", 32), ("RTT 1", 10), ("RTT 2", 10), ("RTT 3", 10),
        })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = name,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FillWeight = weight,
            });
        }

        Controls.Add(grid);
    }

    void BuildStatusBar()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = SystemColors.ControlLight };
        statusLabel.AutoSize = true;
        statusLabel.Location = new(8, 5);
        statusLabel.ForeColor = SystemColors.ControlDarkDark;
        statusLabel.Text = "Enter a host and press Start.";
        bar.Controls.Add(statusLabel);
        Controls.Add(bar);
    }

    void Toggle()
    {
        if (tracing) StopTrace();
        else StartTrace();
    }

    void StartTrace()
    {
        var host = hostBox.Text.Trim();
        if (string.IsNullOrEmpty(host)) return;

        grid.Rows.Clear();
        tracing = true;
        hostBox.Enabled = false;
        startStopButton.Text = "Stop";
        startStopButton.BackColor = Color.FromArgb(220, 80, 80);

        int mySeq = ++traceSeq;
        traceThread = new Thread(() => TraceLoop(host, mySeq)) { IsBackground = true };
        traceThread.Start();
    }

    void StopTrace()
    {
        tracing = false;
        startStopButton.Text = "Start";
        startStopButton.BackColor = Color.FromArgb(60, 160, 80);
        hostBox.Enabled = true;
    }

    void TraceLoop(string host, int seq)
    {
        using var pinger = new Ping();
        var buffer = new byte[32];

        for (int ttl = 1; ttl <= 30 && tracing; ttl++)
        {
            var options = new PingOptions { Ttl = ttl, DontFragment = true };
            long[] rtts = [-1, -1, -1];
            string hopIp = "*";
            bool reached = false;

            for (int probe = 0; probe < 3 && tracing; probe++)
            {
                try
                {
                    var reply = pinger.Send(host, 2000, buffer, options);
                    if (reply.Status is IPStatus.TtlExpired or IPStatus.Success)
                    {
                        hopIp = reply.Address.ToString();
                        rtts[probe] = reply.RoundtripTime;
                        if (reply.Status == IPStatus.Success) reached = true;
                    }
                }
                catch { }
            }

            string hostname = hopIp;
            if (hopIp != "*")
            {
                try { hostname = Dns.GetHostEntry(hopIp).HostName; }
                catch { }
            }

            int hop = ttl;
            string ip = hopIp, hn = hostname;
            long[] r = [.. rtts];
            bool done = reached;
            BeginInvoke(() => AddHopRow(hop, ip, hn, r, done));

            if (reached) break;
        }

        BeginInvoke(() =>
        {
            if (traceSeq != seq) return;
            tracing = false;
            startStopButton.Text = "Start";
            startStopButton.BackColor = Color.FromArgb(60, 160, 80);
            hostBox.Enabled = true;
            statusLabel.Text = "Trace complete.";
        });
    }

    void AddHopRow(int hop, string ip, string hostname, long[] rtts, bool reached)
    {
        static string Fmt(long rtt) => rtt >= 0 ? $"{rtt} ms" : "*";

        int row = grid.Rows.Add(hop, ip, hostname != ip ? hostname : "—", Fmt(rtts[0]), Fmt(rtts[1]), Fmt(rtts[2]));
        grid.Rows[row].DefaultCellStyle.BackColor = reached
            ? Color.FromArgb(220, 245, 220)
            : ip == "*" ? Color.FromArgb(250, 250, 250) : Color.FromArgb(245, 246, 248);

        grid.FirstDisplayedScrollingRowIndex = grid.Rows.Count - 1;
        statusLabel.Text = reached ? $"Reached {ip} in {hop} hop(s)." : $"Tracing... hop {hop}";
    }
}

// ── Interfaces ─────────────────────────────────────────────────────────────

class InterfaceForm : Form
{
    readonly DataGridView grid = new();
    readonly Label lastRefreshedLabel = new();
    readonly ComboBox intervalCombo = new();
    readonly System.Windows.Forms.Timer pollTimer = new();
    readonly Dictionary<string, (long sent, long recv, DateTime when)> prevStats = new();
    readonly Label lblHost = new();
    readonly Label lblUp = new();
    readonly Label lblDown = new();
    readonly Label lblRecv = new();
    readonly Label lblSent = new();

    public InterfaceForm()
    {
        Text = "Ethernet Interface Monitor";
        Size = new(1100, 440);
        MinimumSize = new(700, 300);
        StartPosition = FormStartPosition.CenterScreen;

        BuildSummaryBar();
        BuildToolbar();
        BuildGrid();
        BuildStatusBar();

        grid.CellDoubleClick += OnCellDoubleClick;

        pollTimer.Tick += (_, _) => { if (!IsDisposed) Refresh(); };
        FormClosing += (_, _) => pollTimer.Stop();
        SetInterval(1000);
        Refresh();
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        var intervalLabel = new Label { Text = "Poll interval:", AutoSize = true, Location = new(8, 14) };

        intervalCombo.Items.AddRange(["1 second", "5 seconds", "10 seconds", "30 seconds", "60 seconds"]);
        intervalCombo.SelectedIndex = 0;
        intervalCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        intervalCombo.Location = new(95, 10);
        intervalCombo.Width = 120;
        intervalCombo.SelectedIndexChanged += (_, _) =>
        {
            int[] ms = [1000, 5000, 10000, 30000, 60000];
            SetInterval(ms[intervalCombo.SelectedIndex]);
        };

        toolbar.Controls.AddRange([intervalLabel, intervalCombo]);
        Controls.Add(toolbar);
    }

    void BuildSummaryBar()
    {
        var bar = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = Color.FromArgb(245, 246, 248) };

        var boldFont  = new Font("Segoe UI", 9f, FontStyle.Bold);
        var valueFont = new Font("Segoe UI", 9f);

        var left  = new FlowLayoutPanel { Dock = DockStyle.Fill,  FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new(8, 0, 0, 0) };
        var right = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Padding = new(0, 0, 12, 0), AutoSize = true };

        void AddSep(FlowLayoutPanel panel) => panel.Controls.Add(new Label
        {
            Text = "|", Font = valueFont, ForeColor = Color.FromArgb(180, 180, 180),
            AutoSize = true, Margin = new(30, 9, 30, 0),
        });

        void AddPair(FlowLayoutPanel panel, string key, Label valLbl, string initial)
        {
            panel.Controls.Add(new Label { Text = key, Font = boldFont, ForeColor = Color.FromArgb(90, 90, 110), AutoSize = true, Margin = new(0, 9, 5, 0) });
            valLbl.Text = initial; valLbl.Font = valueFont; valLbl.ForeColor = Color.FromArgb(30, 30, 30);
            valLbl.AutoSize = true; valLbl.Margin = new(0, 9, 0, 0);
            panel.Controls.Add(valLbl);
        }

        AddPair(left, "Host:", lblHost, Dns.GetHostName());
        AddSep(left);
        AddPair(left, "Up:",   lblUp,   "—");
        AddSep(left);
        AddPair(left, "Down:", lblDown, "—");

        AddPair(right, "↓ Recv:", lblRecv, "—");
        AddSep(right);
        AddPair(right, "↑ Sent:", lblSent, "—");

        bar.Controls.Add(right);
        bar.Controls.Add(left);
        Controls.Add(bar);
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
        grid.Cursor = Cursors.Default;

        foreach (var (name, fillWeight, align) in new (string, int, DataGridViewContentAlignment)[]
        {
            ("Name",         20, DataGridViewContentAlignment.MiddleLeft),
            ("Status",        8, DataGridViewContentAlignment.MiddleLeft),
            ("Speed (Mbps)", 10, DataGridViewContentAlignment.MiddleLeft),
            ("MAC Address",  15, DataGridViewContentAlignment.MiddleLeft),
            ("IP Addresses", 25, DataGridViewContentAlignment.MiddleLeft),
            ("↓ Recv",       11, DataGridViewContentAlignment.MiddleRight),
            ("↑ Sent",       11, DataGridViewContentAlignment.MiddleRight),
        })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = name,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FillWeight = fillWeight,
                DefaultCellStyle = { Alignment = align },
                HeaderCell = { Style = { Alignment = align } },
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

    void OnCellDoubleClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0) return;
        var row = grid.Rows[e.RowIndex];
        if (row.Cells[1].Value?.ToString() != "Up") return;

        var nicName = row.Cells[0].Value?.ToString() ?? "";
        var nic = NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault(n => n.Name == nicName);
        if (nic == null) return;

        var ip = nic.GetIPProperties().UnicastAddresses
            .Select(a => a.Address)
            .FirstOrDefault(a => a.AddressFamily == AddressFamily.InterNetwork);

        if (ip == null)
        {
            MessageBox.Show("This interface has no IPv4 address — cannot capture traffic.",
                "Cannot Snoop", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!IsRunningAsAdmin())
        {
            MessageBox.Show("Packet capture requires Administrator privileges.\nRestart the app as Administrator.",
                "Admin Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        new SnifferForm(nicName, ip).Show(this);
    }

    static bool IsRunningAsAdmin() =>
        new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);

    new void Refresh()
    {
        var nics = NetworkInterface.GetAllNetworkInterfaces()
            .Where(n => n.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
            .OrderBy(n => n.Name)
            .ToList();

        grid.Rows.Clear();
        double totalRecv = 0, totalSent = 0;

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
            var isUp = nic.OperationalStatus == OperationalStatus.Up;

            string recvStr = "—", sentStr = "—";
            if (isUp)
            {
                try
                {
                    var stats = nic.GetIPStatistics();
                    var now = DateTime.Now;
                    if (prevStats.TryGetValue(nic.Name, out var prev) && (now - prev.when).TotalSeconds > 0)
                    {
                        double elapsed = (now - prev.when).TotalSeconds;
                        double recvRate = (stats.BytesReceived - prev.recv) / elapsed;
                        double sentRate = (stats.BytesSent     - prev.sent) / elapsed;
                        recvStr = FormatRate(recvRate);
                        sentStr = FormatRate(sentRate);
                        totalRecv += recvRate;
                        totalSent += sentRate;
                    }
                    prevStats[nic.Name] = (stats.BytesSent, stats.BytesReceived, now);
                }
                catch { }
            }

            int row = grid.Rows.Add(nic.Name, nic.OperationalStatus.ToString(), speedMbps, macFormatted, ipText, recvStr, sentStr);
            grid.Rows[row].DefaultCellStyle.BackColor = isUp
                ? Color.FromArgb(220, 245, 220)
                : Color.FromArgb(245, 220, 220);

            if (isUp) grid.Rows[row].Cells[0].ToolTipText = "Double-click to snoop traffic";
        }

        int upCount   = nics.Count(n => n.OperationalStatus == OperationalStatus.Up);
        int downCount = nics.Count - upCount;

        lblUp.Text    = upCount.ToString();
        lblUp.ForeColor   = upCount   > 0 ? Color.FromArgb(120, 230, 120) : Color.FromArgb(30, 30, 30);
        lblDown.Text  = downCount.ToString();
        lblDown.ForeColor = downCount > 0 ? Color.FromArgb(255, 120, 120) : Color.FromArgb(30, 30, 30);
        lblRecv.Text  = totalRecv > 0 ? FormatRate(totalRecv) : "—";
        lblSent.Text  = totalSent > 0 ? FormatRate(totalSent) : "—";

        lastRefreshedLabel.Text = $"Last refreshed: {DateTime.Now:HH:mm:ss}  |  {nics.Count} interface(s) found";
    }

    static string FormatRate(double bytesPerSec) => bytesPerSec switch
    {
        < 0          => "—",
        >= 1_048_576 => $"{bytesPerSec / 1_048_576:F1} MB/s",
        >= 1_024     => $"{bytesPerSec / 1_024:F1} KB/s",
        _            => $"{bytesPerSec:F0} B/s",
    };
}

// ── Sniffer ────────────────────────────────────────────────────────────────

class SnifferForm : Form
{
    readonly string nicName;
    readonly IPAddress bindIp;
    readonly DataGridView grid = new();
    readonly Panel trafficPanel = new();
    readonly Label statusLabel = new();
    readonly Button toggleButton = new();
    readonly Button clearButton = new();
    readonly TextBox filterBox = new();
    volatile string filterSnapshot = "";

    readonly Queue<(DateTime time, string protocol, int bytes)> packetHistory = new();
    readonly object historyLock = new();
    readonly System.Windows.Forms.Timer renderTimer = new() { Interval = 500 };

    Socket? rawSocket;
    Thread? captureThread;
    volatile bool capturing;
    int packetCount;

    public SnifferForm(string nicName, IPAddress bindIp)
    {
        this.nicName = nicName;
        this.bindIp = bindIp;

        Text = $"Traffic Sniffer — {nicName} ({bindIp})";
        Size = new(1050, 580);
        MinimumSize = new(700, 420);
        StartPosition = FormStartPosition.CenterParent;

        BuildTrafficBars();
        BuildToolbar();
        BuildGrid();
        BuildStatusBar();

        FormClosing += (_, _) => { StopCapture(); renderTimer.Stop(); };
        StartCapture();
    }

    void BuildTrafficBars()
    {
        trafficPanel.Dock = DockStyle.Top;
        trafficPanel.Height = 70;
        trafficPanel.BackColor = Color.FromArgb(22, 22, 35);
        trafficPanel.Paint += PaintTrafficBars;
        renderTimer.Tick += (_, _) => trafficPanel.Invalidate();
        Controls.Add(trafficPanel);
    }

    void PaintTrafficBars(object? sender, PaintEventArgs e)
    {
        var g = e.Graphics;
        int w = trafficPanel.Width;
        int h = trafficPanel.Height;
        g.Clear(Color.FromArgb(22, 22, 35));

        const int numBuckets = 60;
        var tcp   = new int[numBuckets];
        var udp   = new int[numBuckets];
        var icmp  = new int[numBuckets];
        var other = new int[numBuckets];

        var now = DateTime.Now;
        lock (historyLock)
        {
            foreach (var (time, protocol, bytes) in packetHistory)
            {
                int age = (int)(now - time).TotalSeconds;
                if (age >= numBuckets) continue;
                int b = numBuckets - 1 - age;
                switch (protocol)
                {
                    case "TCP":  tcp[b]   += bytes; break;
                    case "UDP":  udp[b]   += bytes; break;
                    case "ICMP": icmp[b]  += bytes; break;
                    default:     other[b] += bytes; break;
                }
            }
        }

        const int legendH = 16;
        int chartH = h - legendH;
        int maxTotal = Enumerable.Range(0, numBuckets)
            .Select(i => tcp[i] + udp[i] + icmp[i] + other[i])
            .DefaultIfEmpty(1).Max();
        if (maxTotal < 1) maxTotal = 1;

        float barW = (float)w / numBuckets;
        for (int i = 0; i < numBuckets; i++)
        {
            float scale = (float)chartH / maxTotal;
            float x = i * barW;
            float y = legendH + chartH;

            void DrawSeg(int val, Color color)
            {
                if (val <= 0) return;
                float segH = val * scale;
                y -= segH;
                using var br = new SolidBrush(color);
                g.FillRectangle(br, x, y, Math.Max(barW - 1, 1), segH);
            }

            DrawSeg(other[i], Color.FromArgb(110, 110, 130));
            DrawSeg(icmp[i],  Color.FromArgb(230, 175,  50));
            DrawSeg(udp[i],   Color.FromArgb( 55, 185,  90));
            DrawSeg(tcp[i],   Color.FromArgb( 55, 135, 255));
        }

        using var legendFont = new Font("Segoe UI", 7.5f);
        int lx = 6;
        foreach (var (label, color) in new (string, Color)[]
        {
            ("TCP",   Color.FromArgb( 55, 135, 255)),
            ("UDP",   Color.FromArgb( 55, 185,  90)),
            ("ICMP",  Color.FromArgb(230, 175,  50)),
            ("Other", Color.FromArgb(110, 110, 130)),
        })
        {
            using var br = new SolidBrush(color);
            g.FillRectangle(br, lx, 4, 8, 8);
            using var textBr = new SolidBrush(Color.FromArgb(170, 170, 190));
            g.DrawString(label, legendFont, textBr, lx + 11, 2);
            lx += TextRenderer.MeasureText(label, legendFont).Width + 20;
        }
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        toggleButton.Text = "Stop";
        toggleButton.Location = new(8, 9);
        toggleButton.Width = 80;
        toggleButton.Height = 28;
        toggleButton.BackColor = Color.FromArgb(220, 80, 80);
        toggleButton.ForeColor = Color.White;
        toggleButton.FlatStyle = FlatStyle.Flat;
        toggleButton.Click += OnToggleClick;

        clearButton.Text = "Clear";
        clearButton.Location = new(100, 9);
        clearButton.Width = 70;
        clearButton.Height = 28;
        clearButton.Click += (_, _) =>
        {
            grid.Rows.Clear();
            packetCount = 0;
            lock (historyLock) packetHistory.Clear();
            UpdateStatus();
        };

        var filterLabel = new Label { Text = "Filter:", AutoSize = true, Location = new(188, 15) };

        filterBox.Location = new(230, 11);
        filterBox.Width = 220;
        filterBox.Font = new Font("Consolas", 9.5f);
        filterBox.PlaceholderText = "IP or protocol (e.g. 192.168 or TCP)";
        filterBox.TextChanged += (_, _) => filterSnapshot = filterBox.Text.Trim();

        toolbar.Controls.AddRange([toggleButton, clearButton, filterLabel, filterBox]);
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
        grid.Font = new Font("Consolas", 9f);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(180, 210, 255);
        grid.DefaultCellStyle.SelectionForeColor = Color.Black;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        grid.ColumnHeadersHeight = 30;

        foreach (var (name, weight) in new (string, int)[]
        {
            ("Time", 10), ("Protocol", 8), ("Source", 20), ("Destination", 20), ("Len", 6), ("Info", 36),
        })
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = name,
                SortMode = DataGridViewColumnSortMode.NotSortable,
                FillWeight = weight,
            });
        }

        Controls.Add(grid);
    }

    void BuildStatusBar()
    {
        var bar = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = SystemColors.ControlLight };
        statusLabel.AutoSize = true;
        statusLabel.Location = new(8, 5);
        statusLabel.ForeColor = SystemColors.ControlDarkDark;
        bar.Controls.Add(statusLabel);
        Controls.Add(bar);
        UpdateStatus();
    }

    void OnToggleClick(object? sender, EventArgs e)
    {
        if (capturing) StopCapture();
        else StartCapture();
    }

    void StartCapture()
    {
        try
        {
            rawSocket = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
            rawSocket.Bind(new IPEndPoint(bindIp, 0));
            rawSocket.SetSocketOption(SocketOptionLevel.IP, SocketOptionName.HeaderIncluded, true);
            rawSocket.IOControl(IOControlCode.ReceiveAll, [1, 0, 0, 0], [1, 0, 0, 0]);

            capturing = true;
            captureThread = new Thread(CaptureLoop) { IsBackground = true, Name = "PacketCapture" };
            captureThread.Start();
            renderTimer.Start();

            toggleButton.Text = "Stop";
            toggleButton.BackColor = Color.FromArgb(220, 80, 80);
            UpdateStatus();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to start capture:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    void StopCapture()
    {
        capturing = false;
        renderTimer.Stop();
        try { rawSocket?.Close(); } catch { }
        rawSocket = null;

        if (IsHandleCreated)
        {
            BeginInvoke(() =>
            {
                toggleButton.Text = "Start";
                toggleButton.BackColor = Color.FromArgb(60, 160, 80);
                UpdateStatus();
            });
        }
    }

    void CaptureLoop()
    {
        var buffer = new byte[65535];
        while (capturing)
        {
            try
            {
                if (rawSocket == null) break;
                int len = rawSocket.Receive(buffer);
                if (len < 20) continue;

                var packet = PacketParser.Parse(buffer, len);
                if (packet == null) continue;

                lock (historyLock)
                {
                    var cutoff = DateTime.Now.AddSeconds(-61);
                    while (packetHistory.Count > 0 && packetHistory.Peek().time < cutoff)
                        packetHistory.Dequeue();
                    packetHistory.Enqueue((DateTime.Now, packet.Protocol, packet.Length));
                }

                var filter = filterSnapshot;
                if (!string.IsNullOrEmpty(filter) && !PacketMatchesFilter(packet, filter))
                    continue;

                BeginInvoke(() => AddPacketRow(packet));
            }
            catch (SocketException) { break; }
            catch { }
        }
    }

    static bool PacketMatchesFilter(PacketInfo p, string filter) =>
        p.Source.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
        p.Destination.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
        p.Protocol.Contains(filter, StringComparison.OrdinalIgnoreCase) ||
        p.Info.Contains(filter, StringComparison.OrdinalIgnoreCase);

    void AddPacketRow(PacketInfo packet)
    {
        if (grid.Rows.Count > 5000) grid.Rows.RemoveAt(0);

        int row = grid.Rows.Add(packet.Time, packet.Protocol, packet.Source, packet.Destination, packet.Length, packet.Info);
        grid.Rows[row].DefaultCellStyle.BackColor = ProtocolColor(packet.Protocol);
        grid.FirstDisplayedScrollingRowIndex = grid.Rows.Count - 1;

        packetCount++;
        UpdateStatus();
    }

    static Color ProtocolColor(string proto) => proto switch
    {
        "TCP"  => Color.FromArgb(230, 240, 255),
        "UDP"  => Color.FromArgb(230, 255, 235),
        "ICMP" => Color.FromArgb(255, 245, 220),
        _      => Color.FromArgb(245, 245, 245),
    };

    void UpdateStatus()
    {
        var state = capturing ? "Capturing" : "Stopped";
        statusLabel.Text = $"{state}  |  {packetCount} packets";
    }
}

// ── Packet parsing ─────────────────────────────────────────────────────────

record PacketInfo(string Time, string Protocol, string Source, string Destination, int Length, string Info);

static class PacketParser
{
    public static PacketInfo? Parse(byte[] buf, int len)
    {
        if (len < 20) return null;

        int ihl = (buf[0] & 0x0F) * 4;
        int protocol = buf[9];
        var srcIp = new IPAddress(buf[12..16]).ToString();
        var dstIp = new IPAddress(buf[16..20]).ToString();
        var time = DateTime.Now.ToString("HH:mm:ss.fff");

        return protocol switch
        {
            6  => ParseTcp(buf, ihl, len, time, srcIp, dstIp),
            17 => ParseUdp(buf, ihl, len, time, srcIp, dstIp),
            1  => ParseIcmp(buf, ihl, len, time, srcIp, dstIp),
            _  => new(time, $"IP({protocol})", srcIp, dstIp, len, ""),
        };
    }

    static PacketInfo? ParseTcp(byte[] buf, int ihl, int len, string time, string src, string dst)
    {
        if (len < ihl + 20) return null;
        int srcPort = (buf[ihl] << 8) | buf[ihl + 1];
        int dstPort = (buf[ihl + 2] << 8) | buf[ihl + 3];
        var flagStr = TcpFlags(buf[ihl + 13]);
        return new(time, "TCP", $"{src}:{srcPort}", $"{dst}:{dstPort}", len, flagStr);
    }

    static PacketInfo? ParseUdp(byte[] buf, int ihl, int len, string time, string src, string dst)
    {
        if (len < ihl + 8) return null;
        int srcPort = (buf[ihl] << 8) | buf[ihl + 1];
        int dstPort = (buf[ihl + 2] << 8) | buf[ihl + 3];
        int udpLen  = (buf[ihl + 4] << 8) | buf[ihl + 5];
        return new(time, "UDP", $"{src}:{srcPort}", $"{dst}:{dstPort}", len, $"Len={udpLen}");
    }

    static PacketInfo ParseIcmp(byte[] buf, int ihl, int len, string time, string src, string dst)
    {
        if (len < ihl + 4) return new(time, "ICMP", src, dst, len, "");
        var info = (buf[ihl], buf[ihl + 1]) switch
        {
            (0, _)  => "Echo Reply",
            (8, _)  => "Echo Request (Ping)",
            (3, 0)  => "Dest Unreachable – Net",
            (3, 1)  => "Dest Unreachable – Host",
            (3, 3)  => "Dest Unreachable – Port",
            (11, _) => "Time Exceeded (TTL)",
            var (t, c) => $"Type={t} Code={c}",
        };
        return new(time, "ICMP", src, dst, len, info);
    }

    static string TcpFlags(int f)
    {
        var parts = new List<string>();
        if ((f & 0x02) != 0) parts.Add("SYN");
        if ((f & 0x10) != 0) parts.Add("ACK");
        if ((f & 0x01) != 0) parts.Add("FIN");
        if ((f & 0x04) != 0) parts.Add("RST");
        if ((f & 0x08) != 0) parts.Add("PSH");
        if ((f & 0x20) != 0) parts.Add("URG");
        return parts.Count > 0 ? string.Join(" ", parts) : "—";
    }
}
