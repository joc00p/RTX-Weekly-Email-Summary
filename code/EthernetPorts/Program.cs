using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Principal;
using System.Windows.Forms;

Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.Run(new MainForm());

// ── Main window ────────────────────────────────────────────────────────────

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

        grid.CellDoubleClick += OnCellDoubleClick;

        pollTimer.Tick += (_, _) => Refresh();
        SetInterval(5000);
        Refresh();
    }

    void BuildToolbar()
    {
        var toolbar = new Panel { Dock = DockStyle.Top, Height = 48, Padding = new(8, 8, 8, 0) };

        var intervalLabel = new Label { Text = "Poll interval:", AutoSize = true, Location = new(8, 14) };

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
        grid.Cursor = Cursors.Default;

        foreach (var (name, fillWeight) in new (string, int)[]
        {
            ("Name", 25), ("Status", 10), ("Speed (Mbps)", 12), ("MAC Address", 18), ("IP Addresses", 35),
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

            int row = grid.Rows.Add(nic.Name, nic.OperationalStatus.ToString(), speedMbps, macFormatted, ipText);
            grid.Rows[row].DefaultCellStyle.BackColor = isUp
                ? Color.FromArgb(220, 245, 220)
                : Color.FromArgb(245, 220, 220);

            var tip = isUp ? "Double-click to snoop traffic" : "";
            grid.Rows[row].Cells[0].ToolTipText = tip;
        }

        lastRefreshedLabel.Text = $"Last refreshed: {DateTime.Now:HH:mm:ss}  |  {nics.Count} interface(s) found";
    }
}

// ── Sniffer window ─────────────────────────────────────────────────────────

class SnifferForm : Form
{
    readonly string nicName;
    readonly IPAddress bindIp;
    readonly DataGridView grid = new();
    readonly Label statusLabel = new();
    readonly Button toggleButton = new();
    readonly Button clearButton = new();
    readonly TextBox filterBox = new();

    Socket? rawSocket;
    Thread? captureThread;
    volatile bool capturing;
    int packetCount;

    static readonly Color[] ProtocolColors = [];

    public SnifferForm(string nicName, IPAddress bindIp)
    {
        this.nicName = nicName;
        this.bindIp = bindIp;

        Text = $"Traffic Sniffer — {nicName} ({bindIp})";
        Size = new(1050, 520);
        MinimumSize = new(700, 350);
        StartPosition = FormStartPosition.CenterParent;

        BuildToolbar();
        BuildGrid();
        BuildStatusBar();

        FormClosing += (_, _) => StopCapture();
        StartCapture();
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
        clearButton.Click += (_, _) => { grid.Rows.Clear(); packetCount = 0; UpdateStatus(); };

        var filterLabel = new Label { Text = "Filter:", AutoSize = true, Location = new(188, 15) };

        filterBox.Location = new(230, 11);
        filterBox.Width = 220;
        filterBox.Font = new Font("Consolas", 9.5f);
        filterBox.PlaceholderText = "IP or protocol (e.g. 192.168 or TCP)";

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
        if (capturing)
            StopCapture();
        else
            StartCapture();
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

                var filter = filterBox.Text.Trim();
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
        if (grid.Rows.Count > 5000)
            grid.Rows.RemoveAt(0);

        int row = grid.Rows.Add(
            packet.Time,
            packet.Protocol,
            packet.Source,
            packet.Destination,
            packet.Length,
            packet.Info);

        grid.Rows[row].DefaultCellStyle.BackColor = ProtocolColor(packet.Protocol);

        // Auto-scroll to bottom
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
        int flags = buf[ihl + 13];
        var flagStr = TcpFlags(flags);
        return new(time, "TCP", $"{src}:{srcPort}", $"{dst}:{dstPort}", len, flagStr);
    }

    static PacketInfo? ParseUdp(byte[] buf, int ihl, int len, string time, string src, string dst)
    {
        if (len < ihl + 8) return null;
        int srcPort = (buf[ihl] << 8) | buf[ihl + 1];
        int dstPort = (buf[ihl + 2] << 8) | buf[ihl + 3];
        int udpLen = (buf[ihl + 4] << 8) | buf[ihl + 5];
        return new(time, "UDP", $"{src}:{srcPort}", $"{dst}:{dstPort}", len, $"Len={udpLen}");
    }

    static PacketInfo ParseIcmp(byte[] buf, int ihl, int len, string time, string src, string dst)
    {
        if (len < ihl + 4) return new(time, "ICMP", src, dst, len, "");
        int type = buf[ihl];
        int code = buf[ihl + 1];
        var info = (type, code) switch
        {
            (0, _) => "Echo Reply",
            (8, _) => "Echo Request (Ping)",
            (3, 0) => "Dest Unreachable – Net",
            (3, 1) => "Dest Unreachable – Host",
            (3, 3) => "Dest Unreachable – Port",
            (11, _) => "Time Exceeded (TTL)",
            _ => $"Type={type} Code={code}",
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
