using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Reporter;

// Shown between extraction and final generation: lists the candidate data points per tower with
// checkboxes so the user picks exactly which ones go into the report. Top N are pre-checked.
public class SelectDataPointsForm : Form
{
    private readonly Dictionary<string, List<CheckBox>> _boxes = new();

    public Dictionary<string, List<string>> SelectedByTower { get; } = new();

    public SelectDataPointsForm(Dictionary<string, List<string>> candidates, int defaultChecked)
    {
        Text = "Select Data Points to Include";
        Size = new Size(660, 720);
        MinimumSize = new Size(480, 420);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);

        var flow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(14, 10, 14, 10),
        };

        var intro = new Label
        {
            Text = "Check the points to include in each tower's report. The top 5 per tower are pre-selected.",
            AutoSize = true,
            MaximumSize = new Size(600, 0),
            Margin = new Padding(0, 0, 0, 8),
            ForeColor = Color.FromArgb(80, 80, 80),
        };
        flow.Controls.Add(intro);

        foreach (var tower in candidates.Keys)
        {
            var header = new Label
            {
                Text = tower,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 84, 166),
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 4),
            };
            flow.Controls.Add(header);

            var list = new List<CheckBox>();
            int i = 0;
            foreach (var point in candidates[tower])
            {
                var cb = new CheckBox
                {
                    Text = point,
                    AutoSize = true,
                    MaximumSize = new Size(590, 0),
                    Checked = i < defaultChecked,
                    Margin = new Padding(18, 2, 0, 2),
                };
                flow.Controls.Add(cb);
                list.Add(cb);
                i++;
            }
            _boxes[tower] = list;
        }

        var buttonBar = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var generate = new Button
        {
            Text = "Generate",
            DialogResult = DialogResult.OK,
            Size = new Size(110, 32),
            BackColor = Color.FromArgb(16, 137, 62),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
        };
        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Size = new Size(90, 32),
            BackColor = Color.FromArgb(100, 100, 100),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Anchor = AnchorStyles.Right | AnchorStyles.Top,
        };
        generate.Click += (_, _) =>
        {
            foreach (var (tower, boxes) in _boxes)
            {
                var picked = boxes.Where(b => b.Checked).Select(b => b.Text).ToList();
                if (picked.Count > 0) SelectedByTower[tower] = picked;
            }
        };
        buttonBar.Controls.Add(generate);
        buttonBar.Controls.Add(cancel);
        void PlaceButtons() { generate.Location = new Point(buttonBar.Width - 220, 9); cancel.Location = new Point(buttonBar.Width - 105, 9); }
        buttonBar.Resize += (_, _) => PlaceButtons();
        PlaceButtons();

        AcceptButton = generate;
        CancelButton = cancel;

        Controls.Add(flow);
        Controls.Add(buttonBar);
    }
}
