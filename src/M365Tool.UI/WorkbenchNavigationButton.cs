using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace Office365CleanupTool;

internal sealed class WorkbenchNavigationButton : Button
{
    private bool _hovered;
    private bool _pressed;
    internal float UiScale { get; set; } = 1F;

    public WorkbenchNavigationButton()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint |
            ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw |
            ControlStyles.SupportsTransparentBackColor, true);
        UseVisualStyleBackColor = false;
        Padding = new Padding(18, 0, 18, 0);
    }

    internal (Rectangle Icon, Rectangle Text) GetContentBounds()
    {
        Size iconSize = Image?.Size ?? Size.Empty;
        if (string.IsNullOrEmpty(Text))
            return (new Rectangle((Width - iconSize.Width) / 2, (Height - iconSize.Height) / 2,
                iconSize.Width, iconSize.Height), Rectangle.Empty);

        int left = Padding.Left;
        int gap = WorkbenchUi.Scale(12, UiScale);
        var icon = new Rectangle(left, (Height - iconSize.Height) / 2, iconSize.Width, iconSize.Height);
        int textLeft = icon.Right + gap;
        return (icon, new Rectangle(textLeft, 0, Math.Max(0, Width - textLeft - Padding.Right), Height));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        // A transparent Button does not clear its buffer when its normal painting is replaced.
        // Repaint the sidebar behind it on every frame, including before selected/hover fills.
        using (var background = new SolidBrush(WorkbenchUi.CanvasColor))
            e.Graphics.FillRectangle(background, ClientRectangle);
        if (Parent != null)
        {
            GraphicsState state = e.Graphics.Save();
            try
            {
                e.Graphics.SetClip(ClientRectangle, CombineMode.Intersect);
                e.Graphics.TranslateTransform(-Left, -Top);
                using var parentPaint = new PaintEventArgs(e.Graphics, new Rectangle(Location, Size));
                InvokePaintBackground(Parent, parentPaint);
                InvokePaint(Parent, parentPaint);
            }
            finally
            {
                e.Graphics.Restore(state);
            }
        }

        Color surface = _pressed ? FlatAppearance.MouseDownBackColor
            : _hovered ? FlatAppearance.MouseOverBackColor : BackColor;
        if (surface.A > 0)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var path = WorkbenchUi.BuildRoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), WorkbenchUi.Scale(7, UiScale));
            using var brush = new SolidBrush(surface);
            e.Graphics.FillPath(brush, path);
        }

        if (BackColor == WorkbenchUi.NavigationSelectedColor)
        {
            using var fill = new SolidBrush(WorkbenchUi.PrimaryColor);
            e.Graphics.FillRectangle(fill, 0, WorkbenchUi.Scale(12, UiScale),
                WorkbenchUi.Scale(3, UiScale), Height - WorkbenchUi.Scale(24, UiScale));
        }

        (Rectangle icon, Rectangle text) = GetContentBounds();
        if (Image != null) e.Graphics.DrawImageUnscaled(Image, icon.Location);
        if (!text.IsEmpty)
            TextRenderer.DrawText(e.Graphics, Text, Font, text, ForeColor,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine |
                TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis |
                TextFormatFlags.PreserveGraphicsClipping | TextFormatFlags.PreserveGraphicsTranslateTransform);
        if (Focused && ShowFocusCues)
            ControlPaint.DrawFocusRectangle(e.Graphics, Rectangle.Inflate(ClientRectangle, -6, -6));
    }

    protected override void OnMouseEnter(EventArgs e)
    {
        _hovered = true;
        base.OnMouseEnter(e);
        Invalidate();
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        _hovered = _pressed = false;
        base.OnMouseLeave(e);
        Invalidate();
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left) _pressed = true;
        base.OnMouseDown(e);
        Invalidate();
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        _pressed = false;
        base.OnMouseUp(e);
        Invalidate();
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Space) _pressed = true;
        base.OnKeyDown(e);
        Invalidate();
    }

    protected override void OnKeyUp(KeyEventArgs e)
    {
        _pressed = false;
        base.OnKeyUp(e);
        Invalidate();
    }

    protected override void OnGotFocus(EventArgs e)
    {
        base.OnGotFocus(e);
        Invalidate();
    }

    protected override void OnLostFocus(EventArgs e)
    {
        _pressed = false;
        base.OnLostFocus(e);
        Invalidate();
    }
}
