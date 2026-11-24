using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

internal sealed class BookmarkButton : ToolStripDropDownButton
{
    internal BookmarkButton()
    {
        AccessibleName = ToolTipText = "Bookmarks";
        AutoToolTip = false;
        AutoSize = false;
        ShowDropDownArrow = false;
        DisplayStyle = ToolStripItemDisplayStyle.None;
        Overflow = ToolStripItemOverflow.Never;
        Margin = Padding.Empty;
    }
}

internal sealed class PageTabButton : ToolStripButton
{
    internal uint Dpi = 96;
    public override Size GetPreferredSize(Size constrainingSize)
    {
        var size = base.GetPreferredSize(constrainingSize);
        size.Width = Math.Max(HostChrome.Px(70, Dpi), Math.Min(size.Width, HostChrome.Px(184, Dpi)));
        return size;
    }
}

internal sealed class TabRenderer : ToolStripProfessionalRenderer
{
    internal readonly Color Surface, Ink, MutedInk, Tab, Active, Hover, ActiveMarker;
    private readonly bool contrast;
    internal readonly uint Dpi;
    internal TabRenderer(Color surface, uint dpi, bool highContrast)
    {
        Dpi = dpi;
        contrast = highContrast;
        Surface = highContrast ? SystemColors.Window : surface;
        bool dark = 299 * surface.R + 587 * surface.G + 114 * surface.B < 138000;
        Ink = highContrast ? SystemColors.WindowText : dark ? Color.White : Color.FromArgb(30, 30, 30);
        Color toward = dark ? Color.White : Color.Black;
        MutedInk = highContrast ? Ink : Blend(Ink, Surface, 13);
        Tab = highContrast ? Surface : Blend(Surface, toward, 9);
        Active = highContrast ? SystemColors.Highlight : Blend(Surface, toward, 22);
        Color accent = dark ? Color.FromArgb(207, 169, 255) : Color.FromArgb(99, 48, 160);
        double a = Luminance(accent), b = Luminance(Active);
        if ((Math.Max(a, b) + .05) / (Math.Min(a, b) + .05) < 3)
            accent = b > .179 ? Color.Black : Color.White;
        ActiveMarker = highContrast ? SystemColors.HighlightText : accent;
        Hover = highContrast ? SystemColors.Highlight : Blend(Surface, toward, 15);
        RoundedEdges = false;
    }
    private static Color Blend(Color a, Color b, int percent)
    {
        int keep = 100 - percent;
        return Color.FromArgb((a.R * keep + b.R * percent) / 100,
            (a.G * keep + b.G * percent) / 100, (a.B * keep + b.B * percent) / 100);
    }
    private static bool Highlighted(ToolStripItem item)
    {
        var button = item as ToolStripButton;
        return item.Selected || item.Pressed || (button != null && button.Checked);
    }
    private Color TextColor(ToolStripItem item)
    {
        if (!item.Enabled) return contrast ? SystemColors.GrayText : Blend(Ink, Surface, 40);
        if (contrast && Highlighted(item)) return SystemColors.HighlightText;
        var button = item as ToolStripButton;
        Color background = item.Pressed || (button != null && button.Checked) ? Active :
            item.Selected ? Hover : item is PageTabButton ? Tab : Surface;
        Color text = Highlighted(item) ? Ink : MutedInk;
        double bg = Luminance(background), fg = Luminance(text);
        if ((Math.Max(bg, fg) + 0.05) / (Math.Min(bg, fg) + 0.05) >= 4.5) return text;
        return bg > 0.179 ? Color.Black : Color.White;
    }
    private static double Luminance(Color c)
    {
        return 0.2126 * Linear(c.R) + 0.7152 * Linear(c.G) + 0.0722 * Linear(c.B);
    }
    private static double Linear(byte component)
    {
        double value = component / 255.0;
        return value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);
    }
    protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { e.Graphics.Clear(Surface); }
    protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
    {
        if (e.ToolStrip is ToolStripDropDown)
            ControlPaint.DrawBorder(e.Graphics, e.ToolStrip.ClientRectangle, contrast ? Ink : Hover, ButtonBorderStyle.Solid);
    }
    protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { }
    protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
    {
        e.ArrowColor = TextColor(e.Item);
        base.OnRenderArrow(e);
    }
    protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
    {
        e.TextColor = TextColor(e.Item);
        var tab = e.Item as PageTabButton;
        if (tab != null)
        {
            if (!tab.Checked) e.TextFont = e.ToolStrip.Font;
            e.TextFormat |= TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
        }
        base.OnRenderItemText(e);
        DrawFocus(e, e.TextColor);
    }
    private void DrawFocus(ToolStripItemRenderEventArgs e, Color color)
    {
        if (e.ToolStrip is PageStrip && e.ToolStrip.ContainsFocus && e.Item.Selected)
        {
            int inset = HostChrome.Px(3, Dpi);
            using (var pen = new Pen(color) { DashStyle = DashStyle.Dot })
                e.Graphics.DrawRectangle(pen, inset, inset, Math.Max(1, e.Item.Width - 2 * inset - 1), Math.Max(1, e.Item.Height - 2 * inset - 1));
        }
    }
    protected override void OnRenderButtonBackground(ToolStripItemRenderEventArgs e)
    {
        var tab = e.Item as PageTabButton;
        if (tab == null || e.ToolStrip is ToolStripDropDown) { PaintCommand(e); return; }
        int w = tab.Width, h = tab.Height;
        int diameter = Math.Min(HostChrome.Px(16, Dpi), Math.Min(w, h));
        if (diameter < 2) return;
        var state = e.Graphics.Save();
        try
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var shape = new GraphicsPath())
            using (var fill = new SolidBrush(tab.Checked ? Active : tab.Selected ? Hover : Tab))
            {
                shape.AddLine(0, 0, w, 0);
                shape.AddArc(w - diameter, h - diameter, diameter, diameter, 0, 90);
                shape.AddArc(0, h - diameter, diameter, diameter, 90, 90);
                shape.CloseFigure();
                e.Graphics.FillPath(fill, shape);
                if (contrast)
                    using (var border = new Pen(Ink) { Alignment = PenAlignment.Inset }) e.Graphics.DrawPath(border, shape);
                if (tab.Checked)
                {
                    int inset = HostChrome.Px(8, Dpi);
                    using (var marker = new SolidBrush(ActiveMarker))
                        e.Graphics.FillRectangle(marker, inset, 0, Math.Max(1, w - 2 * inset), HostChrome.Px(3, Dpi));
                }
            }
        }
        finally { e.Graphics.Restore(state); }
    }
    private void PaintCommand(ToolStripItemRenderEventArgs e)
    {
        if (Highlighted(e.Item))
            using (var brush = new SolidBrush(e.Item.Pressed ? Active : Hover))
                e.Graphics.FillRectangle(brush, 0, 0, e.Item.Width, e.Item.Height);
    }
    protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e) { PaintCommand(e); }
    protected override void OnRenderDropDownButtonBackground(ToolStripItemRenderEventArgs e)
    {
        PaintCommand(e);
        if (!(e.Item is BookmarkButton)) return;
        float cx = e.Item.Width / 2f, cy = e.Item.Height / 2f;
        float radius = Math.Min(HostChrome.Px(8, Dpi), Math.Min(cx, cy) - 1);
        if (radius < 1) return;
        var points = new PointF[10];
        for (int i = 0; i < points.Length; i++)
        {
            double angle = (i * Math.PI / 5) - Math.PI / 2;
            float r = i % 2 == 0 ? radius : radius * 0.46f;
            points[i] = new PointF(cx + r * (float)Math.Cos(angle), cy + r * (float)Math.Sin(angle));
        }
        var saved = e.Graphics.Save();
        try
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(TextColor(e.Item), 1.5f * Dpi / 96f) { LineJoin = LineJoin.Round })
                e.Graphics.DrawPolygon(pen, points);
        }
        finally { e.Graphics.Restore(saved); }
        DrawFocus(e, TextColor(e.Item));
    }
    protected override void OnRenderSplitButtonBackground(ToolStripItemRenderEventArgs e)
    {
        PaintCommand(e);
        var split = (ToolStripSplitButton)e.Item;
        DrawArrow(new ToolStripArrowRenderEventArgs(e.Graphics, split, split.DropDownButtonBounds, TextColor(split), ArrowDirection.Down));
    }
    protected override void OnRenderOverflowButtonBackground(ToolStripItemRenderEventArgs e)
    {
        PaintCommand(e);
        DrawArrow(new ToolStripArrowRenderEventArgs(e.Graphics, e.Item, new Rectangle(Point.Empty, e.Item.Size), TextColor(e.Item), ArrowDirection.Down));
    }
    protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
    {
        using (var pen = new Pen(contrast ? Ink : Active))
        {
            if (e.Vertical)
            {
                int inset = HostChrome.Px(7, Dpi);
                e.Graphics.DrawLine(pen, e.Item.Width / 2, inset, e.Item.Width / 2, e.Item.Height - inset - 1);
            }
            else e.Graphics.DrawLine(pen, 2, e.Item.Height / 2, e.Item.Width - 3, e.Item.Height / 2);
        }
    }
}
