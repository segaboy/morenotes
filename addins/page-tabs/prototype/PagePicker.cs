using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

// Opening a branch reads just that level of the hierarchy. Only page items navigate.
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

// Keep the native menu item for submenu navigation, keyboard selection and accessibility.
internal sealed class NotebookPickerItem : ToolStripMenuItem
{
    private readonly Color cover;
    internal NotebookPickerItem(BookmarkNotebook notebook, ToolStripDropDown menu)
    {
        cover = SectionPalette.NotebookIcon(notebook.Color);
        ToolTipText = notebook.Name;
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
        Color color = palette.HighContrast || cover.IsEmpty ? ink : cover;
        var g = e.Graphics;
        g.Clear(background);

        int left = HostChrome.Px(8, palette.Dpi), height = HostChrome.Px(20, palette.Dpi);
        var state = g.Save();
        try
        {
            g.TranslateTransform(left, (Height - height) / 2);
            g.ScaleTransform(height / 20f, height / 20f);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = new GraphicsPath())
            using (var brush = new SolidBrush(color))
            using (var detail = new SolidBrush(palette.HighContrast ? background : palette.NotebookInk(color)))
            {
                path.AddArc(0, 0, 4, 4, 180, 90);
                path.AddArc(11, 0, 4, 4, 270, 90);
                path.AddArc(11, 16, 4, 4, 0, 90);
                path.AddArc(0, 16, 4, 4, 90, 90);
                path.CloseFigure();
                g.FillPath(brush, path);
                if (Math.Abs(color.GetBrightness() - background.GetBrightness()) < .15f)
                    using (var pen = new Pen(ink)) g.DrawPath(pen, path);
                g.FillRectangle(detail, 4, 4, 7, 2);
                for (int y = 5; y <= 13; y += 4) g.FillRectangle(brush, 16, y, 2, 3);
            }
        }
        finally { g.Restore(state); }
        int textLeft = HostChrome.Px(36, palette.Dpi);
        int right = Math.Min(Width, Owner.ClientSize.Width - Bounds.Left) - left;
        int arrow = HostChrome.Px(10, palette.Dpi);
        var title = new Rectangle(textLeft, 0, Math.Max(0, right - arrow - textLeft), Height);
        TextRenderer.DrawText(g, AccessibleName, Font, title, ink, TextFormatFlags.NoPrefix |
            TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        if (hot)
            palette.DrawArrow(new ToolStripArrowRenderEventArgs(g, this,
                new Rectangle(right - arrow, 0, arrow, Height), ink, ArrowDirection.Right));
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
        // Auto-generated dropdowns inherit the strip renderer even after a custom renderer is assigned.
        var menu = new ToolStripDropDownMenu { ShowImageMargin = false, ShowItemToolTips = true };
        root.DropDown = menu;
        root.Disposed += delegate { menu.Dispose(); };
        menu.Items.AddRange(footer);
        if (footer.Length == 0) menu.Items.Add("Loading...");
        var picker = new HierarchyPicker(pages, pickPage, pickSection);
        root.DropDownOpening += delegate {
            var palette = ((TabRenderer)root.Owner.Renderer).NotebookMenu;
            menu.Renderer = palette;
            int width = HostChrome.Px(220, palette.Dpi);
            menu.Font = root.Font;
            menu.MinimumSize = menu.MaximumSize = new Size(width, 0);
            picker.Populate(root, "", null, null, "", footer.Length);
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
            string empty = id.Length == 0 ? "No open notebooks" : section == null ? "No sections" : "No pages";
            try
            {
                foreach (var node in pages.PickerChildren(id))
                {
                    var item = MakeItem(node, notebook, section, path, menu);
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

    private ToolStripMenuItem MakeItem(XElement node, BookmarkNotebook notebook, BookmarkLocation section, string path, ToolStripDropDown menu)
    {
        string id = (string)node.Attribute("ID"), name = (string)node.Attribute("name");
        if (String.IsNullOrEmpty(id)) return null;
        switch (node.Name.LocalName)
        {
            case "Notebook":
                notebook = BookmarkNotebook.FromXml(node);
                if (notebook == null) return null;
                name = notebook.Name;
                break;
            case "SectionGroup":
                if ((bool?)node.Attribute("isRecycleBin") == true) return null;
                break;
            case "Section":
                if ((bool?)node.Attribute("isDeletedPages") == true || notebook == null) return null;
                section = new BookmarkLocation(notebook, name, BookmarkNotebook.ReadColor((string)node.Attribute("color")));
                break;
            case "Page":
                if (section == null || pickPage == null) return null;
                if (String.IsNullOrWhiteSpace(name)) name = "Untitled";
                break;
            default: return null;
        }
        if (String.IsNullOrWhiteSpace(name)) return null;
        var item = node.Name.LocalName == "Notebook" ? new NotebookPickerItem(notebook, menu) : new ToolStripMenuItem();
        item.Text = name.Replace("&", "&&");
        item.AccessibleName = name;
        item.Tag = id;
        string location = path.Length == 0 ? name : path + " / " + name;
        if (node.Name.LocalName == "Page")
            item.Click += delegate { pickPage(new SavedPage(id, name + " / " + notebook.Name, pages.LinkFor(id), section)); };
        else if (pickSection != null && node.Name.LocalName == "Section")
            item.Click += delegate { pickSection(id, location); };
        else
        {
            // WinForms needs a child to expose the submenu before DropDownOpening fires.
            item.DropDownItems.Add(new ToolStripMenuItem("Loading…") { Enabled = false });
            item.DropDownOpening += delegate {
                item.DropDown.Renderer = item.Owner.Renderer;
                Populate(item, id, notebook, section, location, 0);
            };
        }
        return item;
    }
}
