using System;
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
        DropDownItems.AddRange(new ToolStripItem[] { new ToolStripSeparator(), QuickNote, KeepCurrent });
        HierarchyPicker.Attach(this, pages, open, null);
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

    internal static void Attach(ToolStripDropDownItem root, IPages pages, Action<SavedPage> pickPage, Action<string, string> pickSection)
    {
        root.Enabled = pages != null;
        int footer = root.DropDownItems.Count;
        if (footer == 0) root.DropDownItems.Add("Loading...");
        var picker = new HierarchyPicker(pages, pickPage, pickSection);
        root.DropDownOpening += delegate {
            if (root.Owner != null) root.DropDown.Renderer = root.Owner.Renderer;
            picker.Populate(root, "", null, null, "", footer);
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
                    var item = MakeItem(node, notebook, section, path);
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

    private ToolStripMenuItem MakeItem(XElement node, BookmarkNotebook notebook, BookmarkLocation section, string path)
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
        var item = new ToolStripMenuItem(name.Replace("&", "&&")) { AccessibleName = name, Tag = id };
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
