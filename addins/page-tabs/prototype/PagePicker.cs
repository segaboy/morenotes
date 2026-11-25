using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

// Opening a branch reads just that level of the hierarchy.
internal sealed class PagePicker : StripMenuButton
{
    internal readonly ToolStripMenuItem QuickNote, KeepCurrent;

    internal PagePicker(IPages pages, Action<SavedPage> open, Action quickNote, Action keepCurrent)
    {
        Text = "+";
        Enabled = pages != null;
        AccessibleName = ToolTipText = "Open a page in a new tab";
        AutoToolTip = false;
        DisplayStyle = ToolStripItemDisplayStyle.None;
        ShowDropDownArrow = false;
        Overflow = ToolStripItemOverflow.Never;
        QuickNote = new ToolStripMenuItem("New Quick Note", null, delegate { quickNote(); });
        KeepCurrent = new ToolStripMenuItem("Keep current page in a tab", null, delegate { keepCurrent(); });
        HierarchyPicker.Attach(this, pages, open, null, new ToolStripSeparator(), QuickNote, KeepCurrent);
    }
}

internal enum PickerRowKind { Notebook, SectionGroup, Section, Page }

// Keep the native menu item for submenu navigation, keyboard selection and accessibility.
internal sealed class PickerRow : ToolStripMenuItem
{
    private readonly PickerRowKind kind;
    private readonly Color storedColor;
    internal int PageDepth { get; private set; }
    internal PickerRow(PickerRowKind kind, string name, string id, Color color, ToolStripDropDown menu, int depth = 0)
        : base(name.Replace("&", "&&"))
    {
        this.kind = kind;
        storedColor = color;
        PageDepth = depth;
        Tag = id;
        AccessibleName = ToolTipText = name;
        uint dpi = ((TabRenderer)menu.Renderer).Dpi;
        AutoSize = false;
        Size = new Size(menu.MinimumSize.Width - menu.Padding.Horizontal,
            Math.Max(HostChrome.Px(28, dpi), menu.Font.Height + HostChrome.Px(6, dpi)));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var palette = (TabRenderer)Owner.Renderer;
        bool hot = Selected || Pressed;
        Color background = hot ? palette.Hover : palette.Surface;
        Color ink = hot && palette.HighContrast ? SystemColors.HighlightText : palette.Ink;
        var g = e.Graphics;
        g.Clear(background);

        int left = HostChrome.Px(8, palette.Dpi);
        int textLeft = HostChrome.Px(kind == PickerRowKind.Page ? 12 + 16 * PageDepth : 36, palette.Dpi);
        if (kind != PickerRowKind.Page)
        {
            int height = HostChrome.Px(20, palette.Dpi);
            Color color = kind == PickerRowKind.Notebook ? palette.NotebookCover(storedColor) : palette.SectionFill(storedColor);
            if (palette.HighContrast || storedColor.IsEmpty) color = ink;
            DrawIcon(g, left, height, palette, color, ink, background);
        }
        int right = Math.Min(Width, Owner.ClientSize.Width - Bounds.Left) - left;
        int arrow = HasDropDownItems ? HostChrome.Px(10, palette.Dpi) : 0;
        var title = new Rectangle(textLeft, 0, Math.Max(0, right - arrow - textLeft), Height);
        TextRenderer.DrawText(g, AccessibleName, Font, title, ink, TextFormatFlags.NoPrefix |
            TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        if (hot && arrow > 0)
            palette.DrawArrow(new ToolStripArrowRenderEventArgs(g, this,
                new Rectangle(right - arrow, 0, arrow, Height), ink, ArrowDirection.Right));
    }

    private void DrawIcon(Graphics g, int left, int height, TabRenderer palette, Color color, Color ink, Color background)
    {
        var state = g.Save();
        try
        {
            g.TranslateTransform(left, (Height - height) / 2);
            g.ScaleTransform(height / 20f, height / 20f);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = new GraphicsPath())
            using (var brush = new SolidBrush(color))
            {
                if (kind == PickerRowKind.SectionGroup)
                {
                    // A neutral folder distinguishes containers from selectable sections.
                    using (var pen = new Pen(ink, 1.3f) { LineJoin = LineJoin.Round })
                        g.DrawPolygon(pen, new[] { new Point(1, 5), new Point(7, 5), new Point(9, 7),
                            new Point(17, 7), new Point(17, 16), new Point(1, 16) });
                    return;
                }
                if (kind == PickerRowKind.Section)
                {
                    // OneNote's navigation pane uses a vertical section marker with a V notch.
                    path.AddArc(3, 2, 3, 3, 180, 90);
                    path.AddArc(13, 2, 3, 3, 270, 90);
                    path.AddLine(16, 18, 9.5f, 14);
                    path.AddLine(9.5f, 14, 3, 18);
                }
                else
                {
                    path.AddArc(0, 0, 4, 4, 180, 90);
                    path.AddArc(11, 0, 4, 4, 270, 90);
                    path.AddArc(11, 16, 4, 4, 0, 90);
                    path.AddArc(0, 16, 4, 4, 90, 90);
                }
                path.CloseFigure();
                g.FillPath(brush, path);
                if (Math.Abs(color.GetBrightness() - background.GetBrightness()) < .15f)
                    using (var pen = new Pen(ink)) g.DrawPath(pen, path);
                if (kind == PickerRowKind.Notebook)
                {
                    using (var detail = new SolidBrush(palette.HighContrast ? background : palette.NotebookInk(color)))
                        g.FillRectangle(detail, 4, 4, 7, 2);
                    for (int y = 5; y <= 13; y += 4) g.FillRectangle(brush, 16, y, 2, 3);
                }
            }
        }
        finally { g.Restore(state); }
    }
}

// Shared by the new-tab button and both destination pickers in Settings.
internal sealed class HierarchyPicker
{
    private readonly IPages pages;
    private readonly Action<SavedPage> pickPage;
    private readonly Action<string, string> pickSection;

    private HierarchyPicker(IPages pages, Action<SavedPage> pickPage, Action<string, string> pickSection)
    {
        this.pages = pages;
        this.pickPage = pickPage;
        this.pickSection = pickSection;
    }

    internal static void Attach(ToolStripDropDownItem root, IPages pages, Action<SavedPage> pickPage,
        Action<string, string> pickSection, params ToolStripItem[] footer)
    {
        root.Enabled = pages != null;
        var picker = new HierarchyPicker(pages, pickPage, pickSection);
        picker.AttachBranch(root, "", null, null, "", footer);
    }

    private void AttachBranch(ToolStripDropDownItem parent, string id, BookmarkNotebook notebook,
        BookmarkLocation section, string path, params ToolStripItem[] footer)
    {
        // Give the root its own palette. Child menus inherit it and stay owned by their menu item.
        if (id.Length == 0)
        {
            var rootMenu = new ToolStripDropDownMenu();
            parent.DropDown = rootMenu;
            parent.Disposed += delegate { rootMenu.Dispose(); };
        }
        var menu = (ToolStripDropDownMenu)parent.DropDown;
        menu.ShowImageMargin = false;
        menu.ShowItemToolTips = true;
        menu.Items.AddRange(footer);
        if (footer.Length == 0) menu.Items.Add("Loading...");
        parent.DropDownOpening += delegate {
            var palette = (TabRenderer)parent.Owner.Renderer;
            if (id.Length == 0) menu.Renderer = palette = palette.NotebookMenu;
            int width = HostChrome.Px(220, palette.Dpi);
            menu.Font = parent.Font;
            menu.MinimumSize = menu.MaximumSize = new Size(width, 0);
            Populate(parent, id, notebook, section, path, footer.Length);
        };
    }

    private void Populate(ToolStripDropDownItem parent, string id, BookmarkNotebook notebook, BookmarkLocation section, string path, int footer)
    {
        var menu = parent.DropDown;
        menu.SuspendLayout();
        try
        {
            while (menu.Items.Count > footer)
            {
                int index = menu.Items.Count - footer - 1;
                var old = menu.Items[index];
                menu.Items.RemoveAt(index);
                old.Dispose();
            }
            int count = 0;
            // OneNote returns pages as ordered siblings. pageLevel supplies the parent relationship.
            var ancestors = new string[3];
            string empty = id.Length == 0 ? "No open notebooks" : section == null ? "No sections" : "No pages";
            try
            {
                foreach (var node in pages.PickerChildren(id))
                {
                    var item = MakeItem(node, notebook, section, path, menu, ancestors);
                    if (item != null) menu.Items.Insert(count++, item);
                }
            }
            catch (Exception error)
            {
                if (!(error is COMException || error is UnauthorizedAccessException || error is XmlException)) throw;
                empty = "Could not load this list. Reopen it to retry.";
            }
            if (count == 0) menu.Items.Insert(0, new ToolStripMenuItem(empty) { Enabled = false });
        }
        finally { menu.ResumeLayout(true); }
    }

    private ToolStripMenuItem MakeItem(XElement node, BookmarkNotebook notebook, BookmarkLocation section, string path,
        ToolStripDropDown menu, string[] ancestors)
    {
        string id = (string)node.Attribute("ID"), name = (string)node.Attribute("name");
        if (String.IsNullOrEmpty(id)) return null;
        PickerRowKind kind;
        Color color = Color.Empty;
        switch (node.Name.LocalName)
        {
            case "Notebook":
                notebook = BookmarkNotebook.FromXml(node);
                if (notebook == null) return null;
                name = notebook.Name;
                kind = PickerRowKind.Notebook;
                color = notebook.Color;
                break;
            case "SectionGroup":
                if ((bool?)node.Attribute("isRecycleBin") == true) return null;
                kind = PickerRowKind.SectionGroup;
                break;
            case "Section":
                if ((bool?)node.Attribute("isDeletedPages") == true || notebook == null) return null;
                kind = PickerRowKind.Section;
                color = NotebookPages.ReadSectionColor(node);
                section = new BookmarkLocation(notebook, name, color);
                break;
            case "Page":
                if (section == null || pickPage == null) return null;
                if (String.IsNullOrWhiteSpace(name)) name = "Untitled";
                return MakePageRow(id, name, (string)node.Attribute("pageLevel"), section, menu, ancestors);
            default: return null;
        }
        if (String.IsNullOrWhiteSpace(name)) return null;
        var item = new PickerRow(kind, name, id, color, menu);
        string location = path.Length == 0 ? name : path + " / " + name;
        if (pickSection != null && kind == PickerRowKind.Section)
            item.Click += delegate { pickSection(id, location); };
        else
            AttachBranch(item, id, notebook, section, location);
        return item;
    }

    private PickerRow MakePageRow(string id, string name, string pageLevel, BookmarkLocation section,
        ToolStripDropDown menu, string[] ancestors)
    {
        int level;
        int depth = Int32.TryParse(pageLevel, out level) && level >= 1 && level <= ancestors.Length ? level - 1 : 0;
        while (depth > 0 && ancestors[depth - 1] == null) depth--;
        ancestors[depth] = name;
        Array.Clear(ancestors, depth + 1, ancestors.Length - depth - 1);
        var item = new PickerRow(PickerRowKind.Page, name, id, Color.Empty, menu, depth);
        if (depth > 0)
            item.AccessibleDescription = item.ToolTipText = String.Join(" / ", ancestors, 0, depth + 1);
        item.Click += delegate { pickPage(new SavedPage(id, name + " / " + section.Notebook.Name, pages.LinkFor(id), section)); };
        return item;
    }
}
