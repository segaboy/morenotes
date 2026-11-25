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
    internal string SectionName
    {
        get
        {
            var tab = Tag as SavedTab;
            return tab == null || tab.Page.Location == null ? null : tab.Page.Location.SectionName;
        }
    }
    internal DateTime NextColorCheck { get; set; }
    internal uint Dpi = 96;
    internal int CloseWidth { get { return HostChrome.Px(28, Dpi) + 2 * HostChrome.Px(1, Dpi); } }
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
        if (Owner is ToolStripDropDown) return base.GetPreferredSize(constrainingSize);
        return StripSize(constrainingSize);
    }
    internal Size StripSize(Size constrainingSize)
    {
        var size = base.GetPreferredSize(constrainingSize);
        var palette = Owner == null ? null : Owner.Renderer as TabRenderer;
        // All Fall tabs budget for the leaf so a change of selection doesn't narrow a title.
        if (palette != null && palette.UsesCustomColors) size.Width += HostChrome.Px(14, Dpi);
        size.Width = Math.Max(HostChrome.Px(70, Dpi), Math.Min(size.Width, HostChrome.Px(184, Dpi)));
        if (WidthLimit > 0) size.Width = Math.Min(WidthLimit, size.Width);
        return size;
    }
}

internal sealed class NotebookPinButton : CheckBox
{
    private Font measuredFont;
    private string measuredCaption;
    private RectangleF captionInk;

    internal NotebookPinButton()
    {
        AutoSize = AutoCheck = TabStop = false;
        SetStyle(ControlStyles.Selectable, false);
        SetStyle(ControlStyles.Opaque | ControlStyles.OptimizedDoubleBuffer, true);
        Cursor = Cursors.Hand;
        AccessibleRole = AccessibleRole.CheckButton;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var label = (NotebookTabLabel)Parent;
        var g = e.Graphics;
        g.Clear(label.BackColor);
        var saved = g.Save();
        g.TranslateTransform(-Left, -Top);
        label.PaintBand(g);
        g.Restore(saved);

        // Measure the letters once per caption/font. The line box also contains space
        // above the capitals and below the baseline that the pin should not fill.
        if (measuredFont != label.Font || measuredCaption != label.Text)
        {
            measuredFont = label.Font;
            measuredCaption = label.Text;
            using (var letters = new GraphicsPath())
            using (var format = StringFormat.GenericTypographic)
            {
                letters.AddString(measuredCaption, measuredFont.FontFamily, (int)measuredFont.Style,
                    measuredFont.Size, PointF.Empty, format);
                if (letters.PointCount == 0)
                    letters.AddString("H", measuredFont.FontFamily, (int)measuredFont.Style, measuredFont.Size, PointF.Empty, format);
                captionInk = letters.GetBounds();
            }
        }
        float scale = Math.Min((Width - 2) / 11f, Math.Min(Height - 2, captionInk.Height) / 13f);
        if (scale <= 0) return;
        saved = g.Save();
        g.SmoothingMode = SmoothingMode.AntiAlias;
        float lineTop = (Height - label.Padding.Vertical - measuredFont.Height) / 2f;
        float centerY = lineTop + captionInk.Top + captionInk.Height / 2f;
        g.TranslateTransform(Width / 2f, centerY + (Checked ? -0.3f : 0.65f) * scale);
        g.ScaleTransform(scale, scale);
        if (!Checked) g.RotateTransform(32);
        using (var pen = new Pen(ForeColor, 1.25f) { LineJoin = LineJoin.Round, EndCap = LineCap.Round })
        using (var fill = new SolidBrush(ForeColor))
        {
            var head = new[] { new PointF(-3.4f,-5.6f), new PointF(3.4f,-5.6f),
                new PointF(2.1f,-3.5f), new PointF(2.1f,-.8f), new PointF(4.3f,1.6f),
                new PointF(-4.3f,1.6f), new PointF(-2.1f,-.8f), new PointF(-2.1f,-3.5f) };
            if (Checked) g.FillPolygon(fill, head);
            g.DrawPolygon(pen, head);
            g.DrawLine(pen, 0, 1.6f, 0, 6.2f);
        }
        g.Restore(saved);
    }
}

internal sealed class NotebookTabLabel : Label
{
    internal BookmarkNotebook Notebook;
    internal bool Pinned;
    internal readonly NotebookPinButton Pin;
    internal ToolStripSeparator Divider;
    internal Color GroupFill { get; set; }
    private int cornerRadius;
    private const TextFormatFlags Flags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
    internal NotebookTabLabel(bool hasNotebookPin = false)
    {
        AccessibleRole = AccessibleRole.StaticText;
        AccessibleDescription = "Click to cycle all group labels through left, center, and right alignment.";
        TextAlign = ContentAlignment.MiddleLeft;
        Cursor = Cursors.Hand;
        TabStop = false;
        UseMnemonic = false;
        SetStyle(ControlStyles.Selectable | ControlStyles.StandardDoubleClick, false);
        if (hasNotebookPin)
        {
            Pin = new NotebookPinButton();
            Controls.Add(Pin);
        }
    }
    internal void PaintBand(Graphics graphics)
    {
        if (GroupFill.IsEmpty) return;
        var saved = graphics.Save();
        using (var shape = TabRenderer.TabOutline(Width, Height, cornerRadius, Parent.Dock == DockStyle.Top))
        using (var brush = new SolidBrush(GroupFill))
        {
            graphics.SmoothingMode = SmoothingMode.AntiAlias;
            graphics.PixelOffsetMode = PixelOffsetMode.Half;
            graphics.FillPath(brush, shape);
        }
        graphics.Restore(saved);
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        PaintBand(e.Graphics);
        var title = new Rectangle(Padding.Left, 0, Math.Max(0, Width - Padding.Horizontal), Height - Padding.Vertical);
        TextFormatFlags horizontal = TextFormatFlags.Left;
        if (TextAlign == ContentAlignment.MiddleCenter)
        {
            horizontal = TextFormatFlags.HorizontalCenter;
            int centeredWidth = Width - 2 * Padding.Left;
            if (Pin != null && centeredWidth >= TextRenderer.MeasureText(Text, Font, Size.Empty, Flags).Width)
                title.Width = centeredWidth;
        }
        if (TextAlign == ContentAlignment.MiddleRight) horizontal = TextFormatFlags.Right;
        TextRenderer.DrawText(e.Graphics, Text, Font, title, ForeColor,
            Flags | horizontal | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    internal void Align(int left, int right, int height, uint dpi, int endPadding = 0)
    {
        cornerRadius = TabRenderer.CornerRadius(dpi);
        int bandLeft = Math.Max(0, left - endPadding);
        // The menu font leaves more space above its visible glyphs than below them.
        int inset = cornerRadius + left - bandLeft;
        if (Pin != null)
        {
            int pinWidth = HostChrome.Px(18, dpi);
            Pin.Bounds = new Rectangle(inset - HostChrome.Px(3, dpi), 0, pinWidth, height);
            inset = Pin.Right + HostChrome.Px(2, dpi);
        }
        Padding = new Padding(inset, 0, cornerRadius + endPadding, 2 * HostChrome.Px(1, dpi));
        Bounds = Rectangle.FromLTRB(bandLeft, 0, right + endPadding, height);
    }
}

internal sealed class TabRenderer : ToolStripProfessionalRenderer
{
    internal readonly Color Surface, Ink, MutedInk, Tab, Active, Hover, ActiveMarker, ToolsBand, PinnedBand;
    private readonly bool contrast;
    private readonly bool dark;
    internal readonly MoreTabsTheme Theme;
    internal bool UsesCustomColors { get { return Theme != MoreTabsTheme.NotebookColors; } }
    internal bool HighContrast { get { return contrast; } }
    internal readonly uint Dpi;
    internal readonly bool DarkSections;
    private TabRenderer notebookMenu;
    internal TabRenderer NotebookMenu
    {
        get
        {
            if (contrast || UsesCustomColors) return this;
            if (notebookMenu == null)
            {
                int gray = dark ? 40 : 250;
                notebookMenu = new TabRenderer(Color.FromArgb(gray, gray, gray), Dpi, false, DarkSections, Theme);
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
    internal TabRenderer(Color surface, uint dpi, bool highContrast, bool darkSections, MoreTabsTheme theme)
    {
        Theme = theme;
        DarkSections = darkSections;
        Dpi = dpi;
        contrast = highContrast;
        Surface = highContrast ? SystemColors.Window : surface;
        dark = 299 * surface.R + 587 * surface.G + 114 * surface.B < 138000;
        Ink = highContrast ? SystemColors.WindowText : dark ? Color.White : Color.FromArgb(30, 30, 30);
        Color toward = dark ? Color.White : Color.Black;
        MutedInk = highContrast ? Ink : Blend(Ink, Surface, 13);
        Tab = highContrast ? Surface : Blend(Surface, toward, 9);
        Active = highContrast ? SystemColors.Highlight : Blend(Surface, toward, 22);
        int gray = (299 * Active.R + 587 * Active.G + 114 * Active.B) / 1000;
        ToolsBand = highContrast ? SystemColors.Control : Color.FromArgb(gray, gray, gray);
        PinnedBand = highContrast ? Color.Empty : dark ? Color.FromArgb(128, 89, 28) : Color.FromArgb(232, 197, 120);
        Color accent = dark ? Color.FromArgb(207, 169, 255) : Color.FromArgb(99, 48, 160);
        double a = Luminance(accent), b = Luminance(Active);
        if ((Math.Max(a, b) + .05) / (Math.Min(a, b) + .05) < 3)
            accent = b > .179 ? Color.Black : Color.White;
        ActiveMarker = highContrast ? SystemColors.HighlightText : accent;
        Hover = highContrast ? SystemColors.Highlight : Blend(Surface, toward, 15);
        if (theme == MoreTabsTheme.Fall && !highContrast)
        {
            Color bark = dark ? Color.FromArgb(152, 99, 65) : Color.FromArgb(171, 100, 49);
            Tab = Blend(Surface, bark, dark ? 28 : 16);
            Hover = Blend(Surface, dark ? Color.SandyBrown : Color.Peru, dark ? 35 : 28);
            Active = Blend(Surface, Color.Peru, 48);
            ToolsBand = dark ? Color.FromArgb(104, 75, 61) : Color.FromArgb(204, 177, 151);
            PinnedBand = dark ? Color.FromArgb(152, 99, 36) : Color.FromArgb(231, 180, 87);
            ActiveMarker = dark ? Color.FromArgb(250, 189, 93) : Color.FromArgb(163, 65, 29);
        }
        RoundedEdges = false;
        if (theme == MoreTabsTheme.RoboCop && !highContrast)
        {
            Tab = Blend(Surface, Color.SteelBlue, dark ? 28 : 17);
            Hover = Blend(Surface, Color.DeepSkyBlue, dark ? 33 : 23);
            Active = Blend(Surface, Color.SteelBlue, 53);
            ToolsBand = dark ? Color.FromArgb(55, 74, 93) : Color.FromArgb(180, 192, 205);
            PinnedBand = dark ? Color.FromArgb(49, 104, 135) : Color.FromArgb(116, 161, 188);
            ActiveMarker = dark ? Color.LightCyan : Color.MidnightBlue;
        }
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
        if (page != null && page.Checked) return MutedSection(page.SectionColor, page.SectionName);
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
    internal Color NotebookCover(Color stored, string key = null)
    {
        if (contrast) return SystemColors.Highlight;
        return UsesCustomColors ? CustomColor(key) : SectionPalette.NotebookIcon(stored, dark);
    }
    private Color CustomColor(string key)
    {
        return Theme == MoreTabsTheme.RoboCop ? RoboCopArt.Steel(key, dark) : FallPalette.Foliage(key, dark);
    }
    internal static Color PickSwatch(string key, int[] colors)
    {
        uint hash = 2166136261;
        foreach (char c in key ?? "") hash = unchecked((hash ^ c) * 16777619);
        return Color.FromArgb(255, Color.FromArgb(colors[hash % colors.Length]));
    }
    internal Color NotebookInk(Color fill)
    {
        return contrast ? SystemColors.HighlightText : Luminance(fill) > 0.179 ? Color.Black : Color.White;
    }
    internal Color SectionFill(Color stored, string name = null)
    {
        if (contrast) return SystemColors.Highlight;
        if (UsesCustomColors) return CustomColor(name ?? "Unknown section");
        return stored.IsEmpty ? ActiveMarker : SectionPalette.Native(stored, DarkSections);
    }
    internal Color MutedSection(Color stored, string name = null)
    {
        if (contrast) return Active;
        if (UsesCustomColors) return Blend(SectionFill(stored, name), Surface, dark ? 60 : 69);
        if (stored.IsEmpty) return Active;
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
            if (!tab.Checked) e.TextFont = e.ToolStrip.Font;
            // A ToolStripButton can supply a text rectangle wider than a capped tab.
            // Truncate against the visible title area, keeping the close target entirely clear.
            int left = tab.Padding.Left;
            if (ShowThemeIcon(tab)) left = HostChrome.Px(24, Dpi);
            e.TextRectangle = new Rectangle(left, 0, Math.Max(0, tab.CloseBounds.Left - left), tab.Height);
            e.TextFormat = (e.TextFormat & TextFormatFlags.RightToLeft) | TextFormatFlags.SingleLine |
                TextFormatFlags.VerticalCenter | (e.ToolStrip is ToolStripDropDown ? TextFormatFlags.Left : TextFormatFlags.HorizontalCenter) |
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
        return SectionFill(tab.SectionColor, tab.SectionName);
    }
    private bool ShowThemeIcon(PageTabButton tab)
    {
        return UsesCustomColors && (Theme != MoreTabsTheme.RoboCop || !contrast) &&
            tab.Checked && tab.CloseBounds.Left >= HostChrome.Px(53, Dpi);
    }
    internal void PaintThemeArt(Graphics g, int left, int right, int height)
    {
        if (Theme == MoreTabsTheme.RoboCop)
        {
            int width = HostChrome.Px(RoboCopArt.Skyline.Width, Dpi), h = HostChrome.Px(RoboCopArt.Skyline.Height, Dpi);
            int gap = HostChrome.Px(3, Dpi);
            if (right - left >= width + 2 * gap && height >= h)
                RoboCopArt.Draw(g, RoboCopArt.Skyline, new Rectangle(right - gap - width, (height - h) / 2, width, h));
            return;
        }
        float s = Dpi / 96f;
        if (right - left < 60 * s) return;
        float mid = height / 2f;
        FallPalette.DrawLeaf(g, right - 47*s, mid + 3*s, .45f*s, -52, Blend(Surface, Color.Peru, 65));
        FallPalette.DrawLeaf(g, right - 29*s, mid - 3*s, .75f*s, 28, Blend(Surface, Color.IndianRed, 82));
        FallPalette.DrawLeaf(g, right - 11*s, mid + 2*s, .57f*s, 70, Blend(Surface, Color.Goldenrod, 78));
    }
    private void PaintTab(ToolStripItemRenderEventArgs e, PageTabButton page)
    {
        int w = e.Item.Width, h = e.Item.Height;
        var state = e.Graphics.Save();
        try
        {
            var strip = e.Item.Owner as PageStrip;
            bool sharedTools = page == null && strip != null && !(e.Item is PagePicker) && e.Item != strip.MoreTabs;
            if (sharedTools)
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
                if ((page != null || sharedTools) && !contrast)
                {
                    int edge = HostChrome.Px(1, Dpi);
                    e.Graphics.PixelOffsetMode = PixelOffsetMode.Half;
                    using (var border = new SolidBrush(Blend(sharedTools ? Tab : fill.Color, Color.Black, 25)))
                        e.Graphics.FillPath(border, shape);
                    var saved = e.Graphics.Save();
                    e.Graphics.TranslateTransform(edge, edge);
                    using (var body = TabOutline(w - 2 * edge, h - 2 * edge, CornerRadius(Dpi) - edge, bottom))
                        e.Graphics.FillPath(fill, body);
                    e.Graphics.Restore(saved);
                }
                else e.Graphics.FillPath(fill, shape);
                if (contrast)
                    using (var border = new Pen(Ink) { Alignment = PenAlignment.Inset }) e.Graphics.DrawPath(border, shape);
                if (page != null)
                {
                    e.Graphics.SmoothingMode = SmoothingMode.None;
                    int stripe = HostChrome.Px(3, Dpi);
                    using (var marker = new SolidBrush(MarkerFor(page)))
                        e.Graphics.FillRectangle(marker, 0, bottom ? h - stripe : 0, w, stripe);
                    if (ShowThemeIcon(page))
                    {
                        if (Theme == MoreTabsTheme.RoboCop)
                        {
                            int visorHeight = HostChrome.Px(RoboCopArt.Visor.Height, Dpi);
                            RoboCopArt.Draw(e.Graphics, RoboCopArt.Visor, new Rectangle(HostChrome.Px(3, Dpi),
                                (page.Height - visorHeight) / 2, HostChrome.Px(RoboCopArt.Visor.Width, Dpi), visorHeight));
                        }
                        else FallPalette.DrawLeaf(e.Graphics, HostChrome.Px(12, Dpi), page.Height / 2f,
                            Dpi / 96f, 0, contrast ? SystemColors.HighlightText : ActiveMarker);
                    }
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
        var strip = e.ToolStrip as PageStrip;
        if (strip != null && e.Item == strip.MoreTabs) return;
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
        TextRenderer.DrawText(e.Graphics, "…", e.ToolStrip.Font, new Rectangle(Point.Empty, e.Item.Size), TextColor(e.Item),
            TextFormatFlags.NoPadding | TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
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
