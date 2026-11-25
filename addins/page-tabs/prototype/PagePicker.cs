using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

// Opening a branch reads just that level of the hierarchy. Only page items navigate.
internal sealed class PagePicker : ToolStripDropDownButton
{
    private readonly IPages pages;
    private readonly Action<SavedPage> open;
    internal readonly ToolStripMenuItem QuickNote, KeepCurrent;

    internal PagePicker(IPages pages, Action<SavedPage> open, Action quickNote, Action keepCurrent) : base("+")
    {
        this.pages = pages;
        this.open = open;
        Enabled = pages != null;
        AccessibleName = ToolTipText = "Open a page in a new tab";
        DisplayStyle = ToolStripItemDisplayStyle.Text;
        ShowDropDownArrow = false;
        Overflow = ToolStripItemOverflow.Never;
        QuickNote = new ToolStripMenuItem("New Quick Note", null, delegate { quickNote(); });
        KeepCurrent = new ToolStripMenuItem("Keep current page in a tab", null, delegate { keepCurrent(); });
        DropDownItems.AddRange(new ToolStripItem[] { new ToolStripSeparator(), QuickNote, KeepCurrent });
        DropDownOpening += delegate { Populate(this, "", null, null); };
    }

    private void Populate(ToolStripDropDownItem parent, string id, BookmarkNotebook notebook, BookmarkLocation section)
    {
        var menu = parent.DropDown;
        int footer = parent == this ? 3 : 0;
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
            string empty = parent == this ? "No open notebooks" : section == null ? "No sections" : "No pages";
            try
            {
                foreach (var node in pages.PickerChildren(id))
                {
                    var item = MakeItem(node, notebook, section);
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

    private ToolStripMenuItem MakeItem(XElement node, BookmarkNotebook notebook, BookmarkLocation section)
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
                if (section == null) return null;
                if (String.IsNullOrWhiteSpace(name)) name = "Untitled";
                break;
            default: return null;
        }
        if (String.IsNullOrWhiteSpace(name)) return null;
        var item = new ToolStripMenuItem(name.Replace("&", "&&")) { AccessibleName = name, Tag = id };
        if (node.Name.LocalName == "Page")
            item.Click += delegate { open(new SavedPage(id, name + " / " + notebook.Name, pages.LinkFor(id), section)); };
        else
        {
            // WinForms needs a child to expose the submenu before DropDownOpening fires.
            item.DropDownItems.Add(new ToolStripMenuItem("Loading…") { Enabled = false });
            item.DropDownOpening += delegate {
                item.DropDown.Renderer = DropDown.Renderer;
                Populate(item, id, notebook, section);
            };
        }
        return item;
    }
}
