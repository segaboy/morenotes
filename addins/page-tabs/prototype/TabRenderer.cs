using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

internal class StripMenuButton : ToolStripDropDownButton
{
    internal int TabWidth;

    public override Size GetPreferredSize(Size constrainingSize)
    {
        var size = base.GetPreferredSize(constrainingSize);
        size.Width = TabWidth;
        return size;
    }
}

internal sealed class BookmarkButton : StripMenuButton
{
    internal BookmarkButton()
    {
        Name = "Bookmarks";
        AccessibleName = ToolTipText = "Bookmarks";
        AutoToolTip = false;
        ShowDropDownArrow = false;
        DisplayStyle = ToolStripItemDisplayStyle.None;
        Overflow = ToolStripItemOverflow.Never;
        Margin = Padding.Empty;
    }
}

internal enum ToolIcon { Settings, Back, Forward, Home, NewNote, NewTab }

internal sealed class ToolButton : ToolStripButton
{
    internal int TabWidth;
    internal readonly ToolIcon Icon;
    internal ToolButton(string name, ToolIcon icon) : base(name)
    {
        Name = icon.ToString();
        Icon = icon;
        AccessibleName = ToolTipText = name;
        AutoToolTip = false;
        DisplayStyle = ToolStripItemDisplayStyle.None;
        Overflow = ToolStripItemOverflow.Never;
    }
    public override Size GetPreferredSize(Size constrainingSize)
    {
        var size = base.GetPreferredSize(constrainingSize);
        size.Width = TabWidth;
        return size;
    }
}

internal sealed class PageTabButton : ToolStripButton
{
    internal NotebookTabLabel Group;
    internal int WidthLimit;
    internal Color SectionColor { get; set; }
    internal DateTime NextColorCheck { get; set; }
    internal uint Dpi = 96;
    internal int CloseWidth { get { return HostChrome.Px(28, Dpi); } }
    internal Rectangle CloseBounds { get { return Rectangle.FromLTRB(Math.Max(0, Width - CloseWidth), 0, Width, Height); } }
    internal bool CloseHovered { get; private set; }
    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        SetCloseHover(CloseBounds.Contains(e.Location));
    }
    protected override void OnMouseLeave(EventArgs e)
    {
        SetCloseHover(false);
        base.OnMouseLeave(e);
    }
    private void SetCloseHover(bool hover)
    {
        if (CloseHovered == hover) return;
        CloseHovered = hover;
        Invalidate();
    }
    public override Size GetPreferredSize(Size constrainingSize)
    {
        var size = StripSize(constrainingSize);
        if (Overflow == ToolStripItemOverflow.Always && Group != null)
        {
            int width = TextRenderer.MeasureText(OverflowText, Font, Size.Empty,
                TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width + Padding.Horizontal;
            size.Width = Math.Max(size.Width, Math.Min(width, HostChrome.Px(360, Dpi)));
        }
        return size;
    }
    internal string OverflowText { get { return Text + " / " + (Group.Notebook == null ? ((SavedTab)Tag).Page.NotebookForMenu().Name : Group.Text); } }
    internal Size StripSize(Size constrainingSize)
    {
        var size = base.GetPreferredSize(constrainingSize);
        size.Width = Math.Max(HostChrome.Px(70, Dpi), Math.Min(size.Width, HostChrome.Px(184, Dpi)));
        if (WidthLimit > 0) size.Width = Math.Min(WidthLimit, size.Width);
        return size;
    }
}

internal sealed class NotebookTabLabel : Label
{
    internal BookmarkNotebook Notebook;
    internal ToolStripSeparator Divider;
    internal Color GroupFill { get; set; }
    private const TextFormatFlags Flags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
    internal NotebookTabLabel()
    {
        AccessibleRole = AccessibleRole.StaticText;
        TabStop = false;
        UseMnemonic = false;
        SetStyle(ControlStyles.Selectable, false);
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        if (!GroupFill.IsEmpty)
        {
            var saved = e.Graphics.Save();
            using (var shape = TabRenderer.TabOutline(Width, Height, Padding.Left, Parent.Dock == DockStyle.Top))
            using (var brush = new SolidBrush(GroupFill))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.PixelOffsetMode = PixelOffsetMode.Half;
                e.Graphics.FillPath(brush, shape);
            }
            e.Graphics.Restore(saved);
        }
        var title = new Rectangle(Padding.Left, 0, Math.Max(0, Width - Padding.Horizontal), Height);
        TextRenderer.DrawText(e.Graphics, Text, Font, title, ForeColor,
            Flags | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    internal void Align(int left, int right, int height, int inset)
    {
        Padding = new Padding(inset, 0, GroupFill.IsEmpty ? 0 : inset, 0);
        Bounds = Rectangle.FromLTRB(left, 0, right, height);
    }
}

internal sealed class TabRenderer : ToolStripProfessionalRenderer
{
    internal readonly Color Surface, Ink, MutedInk, Tab, Active, Hover, ActiveMarker;
    private readonly bool contrast;
    internal bool HighContrast { get { return contrast; } }
    internal readonly uint Dpi;
    internal readonly bool DarkSections;
    private TabRenderer notebookMenu;
    internal TabRenderer NotebookMenu
    {
        get
        {
            if (contrast) return this;
            if (notebookMenu == null)
            {
                int gray = Luminance(Surface) < .25 ? 40 : 250;
                notebookMenu = new TabRenderer(Color.FromArgb(gray, gray, gray), Dpi, false, DarkSections);
            }
            return notebookMenu;
        }
    }
    internal static int CornerRadius(uint dpi) { return HostChrome.Px(8, dpi); }
    internal static GraphicsPath TabOutline(int width, int height, int radius, bool bottomDock)
    {
        var shape = new GraphicsPath();
        int diameter = Math.Min(2 * radius, Math.Min(width, 2 * height));
        if (diameter < 2) return shape;
        int flat = bottomDock ? height : 0;
        int arcTop = bottomDock ? 0 : height - diameter;
        int turn = bottomDock ? -90 : 90;
        shape.AddLine(0, flat, width, flat);
        shape.AddArc(width - diameter, arcTop, diameter, diameter, 0, turn);
        shape.AddArc(0, arcTop, diameter, diameter, turn, turn);
        shape.CloseFigure();
        return shape;
    }
    internal TabRenderer(Color surface, uint dpi, bool highContrast, bool darkSections)
    {
        DarkSections = darkSections;
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
    private Color Background(ToolStripItem item)
    {
        var page = item as PageTabButton;
        if (page != null && page.Checked) return MutedSection(page.SectionColor);
        var button = item as ToolStripButton;
        if (item.Pressed || (button != null && button.Checked)) return Active;
        if (item.Selected) return Hover;
        return item is PageTabButton || item is StripMenuButton || item is ToolButton ? Tab : Surface;
    }
    internal Color TextColor(ToolStripItem item)
    {
        if (!item.Enabled) return contrast ? SystemColors.GrayText : Blend(Ink, Surface, 40);
        if (contrast && Highlighted(item)) return SystemColors.HighlightText;
        Color background = Background(item);
        Color text = Highlighted(item) ? Ink : MutedInk;
        double bg = Luminance(background), fg = Luminance(text);
        if ((Math.Max(bg, fg) + 0.05) / (Math.Min(bg, fg) + 0.05) >= 4.5) return text;
        return bg > 0.179 ? Color.Black : Color.White;
    }
    private static double Luminance(Color c)
    {
        return 0.2126 * Linear(c.R) + 0.7152 * Linear(c.G) + 0.0722 * Linear(c.B);
    }
    internal Color NotebookFill(Color notebookColor)
    {
        return contrast ? SystemColors.Highlight : notebookColor.IsEmpty ? Active : notebookColor;
    }
    internal Color NotebookInk(Color fill)
    {
        return contrast ? SystemColors.HighlightText : Luminance(fill) > 0.179 ? Color.Black : Color.White;
    }
    internal Color SectionFill(Color stored)
    {
        return contrast ? SystemColors.Highlight : stored.IsEmpty ? ActiveMarker : SectionPalette.Native(stored, DarkSections);
    }
    internal Color MutedSection(Color stored)
    {
        if (contrast || stored.IsEmpty) return Active;
        return SectionFill(stored).ToArgb() == stored.ToArgb() ? Blend(stored, Surface, 25) : stored;
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
            if (e.ToolStrip is ToolStripDropDown && tab.Group != null) e.Text = tab.OverflowText;
            if (!tab.Checked) e.TextFont = e.ToolStrip.Font;
            // A ToolStripButton can supply a text rectangle wider than a capped tab.
            // Truncate against the visible title area, keeping the close target entirely clear.
            int left = tab.Padding.Left;
            e.TextRectangle = new Rectangle(left, 0, Math.Max(0, tab.CloseBounds.Left - left), tab.Height);
            e.TextFormat = (e.TextFormat & TextFormatFlags.RightToLeft) | TextFormatFlags.SingleLine |
                TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter |
                TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix;
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
        var tool = e.Item as ToolButton;
        if (tool != null)
        {
            if (e.ToolStrip is PageStrip) PaintTab(e, null);
            else PaintCommand(e);
            DrawToolIcon(e, tool.Icon);
            return;
        }
        var tab = e.Item as PageTabButton;
        if (tab == null) { PaintCommand(e); return; }
        PaintTab(e, tab);
        PaintClose(e.Graphics, tab);
    }
    private void PaintClose(Graphics g, PageTabButton tab)
    {
        var box = tab.CloseBounds;
        float cx = box.Left + (box.Width - 1) / 2f, cy = (box.Height - 1) / 2f;
        Color ink = TextColor(tab);
        var saved = g.Save();
        try
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (tab.CloseHovered)
            {
                int radius = HostChrome.Px(8, Dpi);
                using (var fill = new SolidBrush(contrast ? SystemColors.Highlight : Blend(Background(tab), ink, 18)))
                    g.FillEllipse(fill, cx - radius, cy - radius, 2 * radius, 2 * radius);
                if (contrast) ink = SystemColors.HighlightText;
            }
            float arm = 3.25f * Dpi / 96f;
            using (var pen = new Pen(ink, 1.4f * Dpi / 96f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                g.DrawLine(pen, cx - arm, cy - arm, cx + arm, cy + arm);
                g.DrawLine(pen, cx - arm, cy + arm, cx + arm, cy - arm);
            }
        }
        finally { g.Restore(saved); }
    }
    internal Color MarkerFor(PageTabButton tab)
    {
        if (contrast) return Highlighted(tab) ? SystemColors.HighlightText : Ink;
        return SectionFill(tab.SectionColor);
    }
    private void PaintTab(ToolStripItemRenderEventArgs e, PageTabButton page)
    {
        int w = e.Item.Width, h = e.Item.Height;
        var state = e.Graphics.Save();
        try
        {
            var strip = e.Item.Owner as PageStrip;
            if (page == null && strip != null && !(e.Item is PagePicker))
            {
                // The tools share an outline while retaining individual hit targets.
                Rectangle both = strip.MenuTabBounds;
                e.Graphics.SetClip(new Rectangle(Point.Empty, e.Item.Size), CombineMode.Intersect);
                e.Graphics.TranslateTransform(both.Left - e.Item.Bounds.Left, 0);
                w = both.Width;
            }
            bool bottom = strip != null && strip.AtBottom;
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var shape = TabOutline(w, h, CornerRadius(Dpi), bottom))
            using (var fill = new SolidBrush(Background(e.Item)))
            {
                if (shape.PointCount == 0) return;
                e.Graphics.FillPath(fill, shape);
                if (contrast)
                    using (var border = new Pen(Ink) { Alignment = PenAlignment.Inset }) e.Graphics.DrawPath(border, shape);
                if (page != null)
                {
                    e.Graphics.SmoothingMode = SmoothingMode.None;
                    int stripe = HostChrome.Px(3, Dpi);
                    using (var marker = new SolidBrush(MarkerFor(page)))
                        e.Graphics.FillRectangle(marker, 0, bottom ? h - stripe : 0, w, stripe);
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
        if (!(e.Item is StripMenuButton)) { PaintCommand(e); return; }
        PaintTab(e, null);
        if (e.Item is PagePicker) { DrawToolIcon(e, ToolIcon.NewTab); return; }
        if (!(e.Item is BookmarkButton)) { DrawToolIcon(e, ToolIcon.Settings); return; }
        float cy = e.Item.Height / 2f;
        float radius = Math.Min(HostChrome.Px(8, Dpi), cy - 1);
        if (radius < 1) return;
        float stroke = 1.5f * Dpi / 96f;
        // The left point is 18 degrees above horizontal. Include the outline's half-width.
        float cx = e.Item.Padding.Left + radius * (float)Math.Cos(Math.PI / 10) + stroke / 2;
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
            e.Graphics.PixelOffsetMode = PixelOffsetMode.Half;
            e.Graphics.FillPolygon(Brushes.Gold, points);
            Color outline = contrast ? TextColor(e.Item) : Luminance(Background(e.Item)) > 0.5 ? Color.DarkGoldenrod : Color.Gold;
            using (var pen = new Pen(outline, stroke) { LineJoin = LineJoin.Round })
                e.Graphics.DrawPolygon(pen, points);
        }
        finally { e.Graphics.Restore(saved); }
        DrawFocus(e, TextColor(e.Item));
    }
    // Original 16-unit outlines based on the silhouettes in OneNote's Quick Access Toolbar and page menu.
    private void DrawToolIcon(ToolStripItemRenderEventArgs e, ToolIcon icon)
    {
        int size = Math.Min(HostChrome.Px(16, Dpi), Math.Min(e.Item.Width, e.Item.Height));
        if (size < 1) return;
        var g = e.Graphics;
        var state = g.Save();
        Color ink = TextColor(e.Item);
        try
        {
            int left = (e.Item.Width - size) / 2;
            var strip = e.ToolStrip as PageStrip;
            if (icon == ToolIcon.Settings && strip != null && e.Item.Bounds.Left == strip.MenuTabBounds.Left)
                left = Math.Max(0, strip.LeadingToolInset - HostChrome.Px(1, Dpi));
            g.TranslateTransform(left, (e.Item.Height - size) / 2);
            g.ScaleTransform(size / 16f, size / 16f);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            using (var pen = new Pen(ink, 1.1f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            {
                switch (icon)
                {
                    case ToolIcon.NewTab:
                        g.DrawLine(pen, 3, 8, 13, 8);
                        g.DrawLine(pen, 8, 3, 8, 13);
                        break;
                    case ToolIcon.Back:
                    case ToolIcon.Forward:
                        if (icon == ToolIcon.Forward) { g.TranslateTransform(16, 0); g.ScaleTransform(-1, 1); }
                        g.DrawEllipse(pen, 1, 1, 14, 14);
                        g.DrawLine(pen, 4.5f, 8, 11.5f, 8);
                        g.DrawLines(pen, new[] { new PointF(7.5f, 5), new PointF(4.5f, 8), new PointF(7.5f, 11) });
                        break;
                    case ToolIcon.Home:
                        g.DrawLines(pen, new[] { new PointF(1, 7.5f), new PointF(8, 1.5f), new PointF(15, 7.5f) });
                        g.DrawLines(pen, new[] { new PointF(3, 7), new PointF(3, 14.5f), new PointF(6.5f, 14.5f),
                            new PointF(6.5f, 10), new PointF(9.5f, 10), new PointF(9.5f, 14.5f), new PointF(13, 14.5f), new PointF(13, 7) });
                        break;
                    case ToolIcon.NewNote:
                        g.DrawLines(pen, new[] { new PointF(6.5f, 2.5f), new PointF(10.5f, 2.5f), new PointF(14, 6),
                            new PointF(14, 14.5f), new PointF(3.5f, 14.5f), new PointF(3.5f, 8.5f) });
                        g.DrawLines(pen, new[] { new PointF(10.5f, 2.5f), new PointF(10.5f, 6), new PointF(14, 6) });
                        if (!contrast && e.Item.Enabled)
                            pen.Color = Luminance(Background(e.Item)) < .3 ? Color.FromArgb(117, 199, 246) : Color.FromArgb(0, 100, 160);
                        g.DrawLine(pen, 1, 4.5f, 6, 4.5f);
                        g.DrawLine(pen, 3.5f, 2, 3.5f, 7);
                        break;
                    case ToolIcon.Settings:
                        g.DrawPolygon(pen, GearOutline);
                        g.DrawEllipse(pen, 5.5f, 5.5f, 5, 5);
                        break;
                }
            }
        }
        finally { g.Restore(state); }
        DrawFocus(e, ink);
    }
    private static readonly PointF[] GearOutline = MakeGearOutline();
    private static PointF[] MakeGearOutline()
    {
        var points = new PointF[32];
        // Two outer and two inner corners per tooth, with a flat tip on each cardinal axis.
        double[] offsets = { -6, 6, 11, 34 };
        for (int i = 0; i < points.Length; i++)
        {
            double angle = (45 * (i / 4) + offsets[i % 4]) * Math.PI / 180;
            double radius = i % 4 < 2 ? 7 : 5.6;
            points[i] = new PointF(8 + (float)(radius * Math.Cos(angle)), 8 + (float)(radius * Math.Sin(angle)));
        }
        return points;
    }
    protected override void OnRenderOverflowButtonBackground(ToolStripItemRenderEventArgs e)
    {
        PaintCommand(e);
        DrawArrow(new ToolStripArrowRenderEventArgs(e.Graphics, e.Item, new Rectangle(Point.Empty, e.Item.Size), TextColor(e.Item), MenuArrow(e.ToolStrip)));
    }
    private static ArrowDirection MenuArrow(ToolStrip strip)
    {
        return strip.DefaultDropDownDirection == ToolStripDropDownDirection.AboveRight ? ArrowDirection.Up : ArrowDirection.Down;
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
