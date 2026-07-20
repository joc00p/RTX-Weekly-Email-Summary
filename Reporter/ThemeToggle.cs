using System.Drawing;
using System.Windows.Forms;

namespace Reporter;

public class ThemeToggle : Control
{
    private bool _isDark = false;
    [System.ComponentModel.DesignerSerializationVisibility(System.ComponentModel.DesignerSerializationVisibility.Hidden)]
    public bool IsDark
    {
        get => _isDark;
        set { _isDark = value; Invalidate(); }
    }

    public event EventHandler? ThemeChanged;

    public ThemeToggle()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint |
                 ControlStyles.UserPaint |
                 ControlStyles.DoubleBuffer |
                 ControlStyles.ResizeRedraw, true);
        Size = new Size(130, 30);
        Cursor = Cursors.Hand;
        Font = new Font("Segoe UI", 9f);
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        bool clickedDark = e.X > Width / 2;
        if (clickedDark != _isDark)
        {
            _isDark = clickedDark;
            Invalidate();
            ThemeChanged?.Invoke(this, EventArgs.Empty);
        }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

        int w = Width, h = Height, half = w / 2;
        int r = 6;

        // Background track
        var trackColor = Color.FromArgb(200, 200, 200);
        using var trackBrush = new SolidBrush(trackColor);
        FillRoundRect(g, trackBrush, 0, 0, w, h, r);

        // Active half highlight
        var activeColor = _isDark ? Color.FromArgb(50, 50, 70) : Color.FromArgb(0, 120, 215);
        using var activeBrush = new SolidBrush(activeColor);
        if (!_isDark)
            FillRoundRectLeft(g, activeBrush, 0, 0, half, h, r);
        else
            FillRoundRectRight(g, activeBrush, half, 0, half, h, r);

        // Border
        using var borderPen = new Pen(Color.FromArgb(160, 160, 160));
        DrawRoundRect(g, borderPen, 0, 0, w - 1, h - 1, r);

        // Divider
        g.DrawLine(borderPen, half, 2, half, h - 3);

        // Labels
        var lightColor = !_isDark ? Color.White : Color.FromArgb(80, 80, 80);
        var darkColor  = _isDark  ? Color.White : Color.FromArgb(80, 80, 80);

        var fmt = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
        g.DrawString("☀ Light", Font, new SolidBrush(lightColor), new RectangleF(0, 0, half, h), fmt);
        g.DrawString("🌙 Dark",  Font, new SolidBrush(darkColor),  new RectangleF(half, 0, half, h), fmt);
    }

    private static void FillRoundRect(Graphics g, Brush b, int x, int y, int w, int h, int r)
    {
        using var path = RoundRectPath(x, y, w, h, r);
        g.FillPath(b, path);
    }

    private static void FillRoundRectLeft(Graphics g, Brush b, int x, int y, int w, int h, int r)
    {
        // Left half with rounded left corners only
        using var path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(x, y, r * 2, r * 2, 180, 90);
        path.AddLine(x + r, y, x + w, y);
        path.AddLine(x + w, y, x + w, y + h);
        path.AddLine(x + w, y + h, x + r, y + h);
        path.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
        path.CloseFigure();
        g.FillPath(b, path);
    }

    private static void FillRoundRectRight(Graphics g, Brush b, int x, int y, int w, int h, int r)
    {
        // Right half with rounded right corners only
        using var path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddLine(x, y, x + w - r, y);
        path.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
        path.AddLine(x + w, y + r, x + w, y + h - r);
        path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
        path.AddLine(x + w - r, y + h, x, y + h);
        path.CloseFigure();
        g.FillPath(b, path);
    }

    private static void DrawRoundRect(Graphics g, Pen p, int x, int y, int w, int h, int r)
    {
        using var path = RoundRectPath(x, y, w, h, r);
        g.DrawPath(p, path);
    }

    private static System.Drawing.Drawing2D.GraphicsPath RoundRectPath(int x, int y, int w, int h, int r)
    {
        var path = new System.Drawing.Drawing2D.GraphicsPath();
        path.AddArc(x, y, r * 2, r * 2, 180, 90);
        path.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
        path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
        path.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
        path.CloseFigure();
        return path;
    }
}
