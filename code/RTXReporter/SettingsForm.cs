using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace RTXReporter;

public class SettingsForm : Form
{
    private readonly AppSettings _settings;
    private TextBox _templatePathBox = null!;

    public SettingsForm(AppSettings settings)
    {
        _settings = settings;
        BuildUI();
    }

    private void BuildUI()
    {
        Text = "Settings";
        Size = new Size(600, 160);
        MinimumSize = new Size(500, 160);
        MaximumSize = new Size(int.MaxValue, 160);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var label = new Label
        {
            Text = "PPTX Template:",
            Location = new Point(12, 20),
            AutoSize = true,
        };

        _templatePathBox = new TextBox
        {
            Location = new Point(130, 17),
            Width = 360,
            Text = _settings.TemplatePath,
            ReadOnly = true,
            BackColor = SystemColors.Window,
        };

        var browseBtn = new Button
        {
            Text = "Browse...",
            Location = new Point(498, 15),
            Size = new Size(80, 26),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
        };
        browseBtn.Click += Browse_Click;

        var saveBtn = new Button
        {
            Text = "Save",
            Location = new Point(410, 80),
            Size = new Size(80, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(16, 137, 62),
            ForeColor = Color.White,
            DialogResult = DialogResult.OK,
        };
        saveBtn.Click += Save_Click;

        var cancelBtn = new Button
        {
            Text = "Cancel",
            Location = new Point(498, 80),
            Size = new Size(80, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(100, 100, 100),
            ForeColor = Color.White,
            DialogResult = DialogResult.Cancel,
        };

        AcceptButton = saveBtn;
        CancelButton = cancelBtn;

        Controls.AddRange(new Control[] { label, _templatePathBox, browseBtn, saveBtn, cancelBtn });
    }

    private void Browse_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Select PPTX Template",
            Filter = "PowerPoint Template|*.pptx|All Files|*.*",
            FileName = _templatePathBox.Text,
        };

        if (File.Exists(_templatePathBox.Text))
            dlg.InitialDirectory = Path.GetDirectoryName(_templatePathBox.Text);

        if (dlg.ShowDialog() == DialogResult.OK)
            _templatePathBox.Text = dlg.FileName;
    }

    private void Save_Click(object? sender, EventArgs e)
    {
        _settings.TemplatePath = _templatePathBox.Text;
        _settings.Save();
    }
}
