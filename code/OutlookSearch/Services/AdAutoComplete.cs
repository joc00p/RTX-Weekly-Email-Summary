using System.Drawing;
using System.Windows.Forms;

namespace OutlookSearch.Services;

/// <summary>
/// Attaches a live "type-ahead from AD" dropdown to a TextBox. As the user types,
/// it debounces, queries <see cref="AdLookupService"/> in the background, and shows
/// matching people in a floating list. Picking one fills the box with that address.
/// </summary>
public sealed class AdAutoComplete : IDisposable
{
    private const int MaxSuggestions = 8;

    private readonly TextBox _box;
    private readonly AdLookupService _ad;
    private readonly ListBox _list;
    private readonly ToolStripDropDown _popup;
    private readonly System.Windows.Forms.Timer _debounce;
    private CancellationTokenSource? _cts;
    private bool _suppress;   // don't re-query while we set the text programmatically

    public AdAutoComplete(TextBox box, AdLookupService ad)
    {
        _box = box;
        _ad = ad;

        _list = new ListBox
        {
            BorderStyle = BorderStyle.FixedSingle,
            IntegralHeight = false,
            Font = box.Font,
            Dock = DockStyle.Fill
        };
        var host = new ToolStripControlHost(_list)
        {
            AutoSize = false,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        _popup = new ToolStripDropDown
        {
            AutoClose = true,          // framework dismisses it on outside clicks / app deactivate
            AutoSize = false,
            DropShadowEnabled = true,
            Padding = Padding.Empty
        };
        _popup.Items.Add(host);

        _debounce = new System.Windows.Forms.Timer { Interval = 250 };
        _debounce.Tick += async (_, _) => { _debounce.Stop(); await QueryAsync(); };

        _box.TextChanged += OnTextChanged;
        _box.KeyDown += OnKeyDown;
        _list.MouseClick += (_, _) => Commit();
    }

    private void OnTextChanged(object? s, EventArgs e)
    {
        if (_suppress) return;
        _debounce.Stop();
        _debounce.Start();
    }

    private async Task QueryAsync()
    {
        var term = _box.Text.Trim();
        if (term.Length < 2) { Hide(); return; }

        _cts?.Cancel();
        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        List<AdPerson> people;
        try { people = await _ad.SearchAsync(term, MaxSuggestions, token); }
        catch { return; }

        if (token.IsCancellationRequested || !_box.Focused) return;
        if (people.Count == 0) { Hide(); return; }

        _list.BeginUpdate();
        _list.Items.Clear();
        foreach (var p in people) _list.Items.Add(p);
        _list.EndUpdate();
        _list.SelectedIndex = -1;

        int rows = Math.Min(people.Count, MaxSuggestions);
        var size = new Size(Math.Max(_box.Width, 280), _list.ItemHeight * rows + 4);
        _popup.Size = size;

        if (!_popup.Visible)
            _popup.Show(_box, new Point(0, _box.Height));
    }

    private void OnKeyDown(object? s, KeyEventArgs e)
    {
        if (!_popup.Visible) return;
        switch (e.KeyCode)
        {
            case Keys.Down:
                if (_list.Items.Count > 0)
                    _list.SelectedIndex = Math.Min(_list.SelectedIndex + 1, _list.Items.Count - 1);
                e.Handled = e.SuppressKeyPress = true;
                break;
            case Keys.Up:
                if (_list.SelectedIndex > 0) _list.SelectedIndex--;
                e.Handled = e.SuppressKeyPress = true;
                break;
            case Keys.Enter:
                if (_list.SelectedIndex >= 0) { Commit(); e.Handled = e.SuppressKeyPress = true; }
                break;
            case Keys.Escape:
                Hide();
                e.Handled = e.SuppressKeyPress = true;
                break;
        }
    }

    private void Commit()
    {
        if (_list.SelectedItem is not AdPerson p) return;
        _suppress = true;
        _box.Text = p.Email;                       // precise value for the sender filter
        _box.SelectionStart = _box.Text.Length;
        _suppress = false;
        Hide();
        _box.Focus();
    }

    private void Hide()
    {
        if (_popup.Visible) _popup.Close();
    }

    public void Dispose()
    {
        _cts?.Cancel();
        _debounce.Dispose();
        _popup.Dispose();
    }
}
