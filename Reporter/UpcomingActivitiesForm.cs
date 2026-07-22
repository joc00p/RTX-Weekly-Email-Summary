using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Reporter;

// Second curation step: pulls the template's "Upcoming Activities / Actions" entries and lets the
// user check/uncheck, edit inline (double-click / F2), drag to reorder, or add new ones.
public class UpcomingActivitiesForm : Form
{
    private readonly ListView _lv;

    public List<string> SelectedEntries { get; } = new();

    public UpcomingActivitiesForm(List<string> existing)
    {
        Text = "Upcoming Activities / Actions";
        Size = new Size(680, 620);
        MinimumSize = new Size(480, 400);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);

        var intro = new Label
        {
            Dock = DockStyle.Top,
            Height = 44,
            Padding = new Padding(12, 10, 12, 0),
            Text = "Check the activities to include in the report's \"Upcoming Activities / Actions\" box. Double-click to edit, drag to reorder, or click Add Entry for a new one.",
            ForeColor = Color.FromArgb(80, 80, 80),
        };

        _lv = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            CheckBoxes = true,
            LabelEdit = true,
            FullRowSelect = true,
            MultiSelect = false,
            HeaderStyle = ColumnHeaderStyle.None,
            AllowDrop = true,
        };
        _lv.Columns.Add("activity", 640);
        foreach (var e in existing) _lv.Items.Add(new ListViewItem(e) { Checked = true });

        _lv.DoubleClick += (_, _) => { if (_lv.SelectedItems.Count > 0) _lv.SelectedItems[0].BeginEdit(); };
        _lv.ItemDrag += (_, e) => _lv.DoDragDrop(e.Item!, DragDropEffects.Move);
        _lv.DragEnter += (_, e) => e.Effect = DragDropEffects.Move;
        _lv.DragOver += (_, e) =>
        {
            e.Effect = DragDropEffects.Move;
            var p = _lv.PointToClient(new Point(e.X, e.Y));
            var over = _lv.GetItemAt(p.X, p.Y);
            if (over == null) { _lv.InsertionMark.Index = -1; return; }
            var r = over.Bounds;
            _lv.InsertionMark.AppearsAfterItem = p.Y > r.Top + r.Height / 2;
            _lv.InsertionMark.Index = over.Index;
        };
        _lv.DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(typeof(ListViewItem)) is not ListViewItem dragged) return;
            var p = _lv.PointToClient(new Point(e.X, e.Y));
            var over = _lv.GetItemAt(p.X, p.Y);
            bool after = false;
            if (over != null) { var r = over.Bounds; after = p.Y > r.Top + r.Height / 2; }
            _lv.Items.Remove(dragged);
            int insertAt = over == null ? _lv.Items.Count : over.Index + (after ? 1 : 0);
            insertAt = Math.Max(0, Math.Min(insertAt, _lv.Items.Count));
            _lv.Items.Insert(insertAt, dragged);
            _lv.InsertionMark.Index = -1;
        };

        var buttonBar = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var add = new Button
        {
            Text = "Add Entry", Location = new Point(12, 9), Size = new Size(110, 32),
            BackColor = Color.FromArgb(70, 130, 180), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
        };
        var apply = new Button
        {
            Text = "Apply", DialogResult = DialogResult.OK, Size = new Size(100, 32),
            BackColor = Color.FromArgb(16, 137, 62), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
        };
        var cancel = new Button
        {
            Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 32),
            BackColor = Color.FromArgb(100, 100, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
        };
        add.Click += (_, _) =>
        {
            var it = _lv.Items.Add(new ListViewItem("New activity") { Checked = true });
            it.Selected = true;
            _lv.Focus();
            it.BeginEdit();
        };
        apply.Click += (_, _) =>
        {
            foreach (ListViewItem it in _lv.Items)
                if (it.Checked && it.Text.Trim().Length > 0) SelectedEntries.Add(it.Text.Trim());
        };
        buttonBar.Controls.Add(add);
        buttonBar.Controls.Add(apply);
        buttonBar.Controls.Add(cancel);
        void Place() { apply.Location = new Point(buttonBar.Width - 205, 9); cancel.Location = new Point(buttonBar.Width - 100, 9); }
        buttonBar.Resize += (_, _) => Place();
        Place();

        AcceptButton = apply;
        CancelButton = cancel;

        Controls.Add(_lv);
        Controls.Add(buttonBar);
        Controls.Add(intro);
    }
}
