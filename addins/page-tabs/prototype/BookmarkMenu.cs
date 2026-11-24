using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

internal sealed class BookmarkNotebook
{
    internal readonly string Id, Name;
    internal readonly Color Color;

    internal BookmarkNotebook(string id, string name, Color color) { Id = id; Name = name; Color = color; }
    internal bool SameAs(BookmarkNotebook other)
    {
        return other != null && Id == other.Id && Name == other.Name && Color.ToArgb() == other.Color.ToArgb();
    }
    internal XElement ToXml()
    {
        return new XElement("notebook", new XAttribute("id", Id), new XAttribute("name", Name),
            Color.IsEmpty ? null : new XAttribute("color", "#" + (Color.ToArgb() & 0xffffff).ToString("X6")));
    }
    internal static Color ReadColor(string value)
    {
        int rgb;
        return value != null && value.Length == 7 && value[0] == '#' &&
            Int32.TryParse(value.Substring(1), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out rgb)
            ? Color.FromArgb(255, Color.FromArgb(rgb)) : Color.Empty;
    }
    internal static BookmarkNotebook FromXml(XElement node)
    {
        if (node == null) return null;
        string id = (string)node.Attribute("ID") ?? (string)node.Attribute("id");
        string name = (string)node.Attribute("nickname");
        if (String.IsNullOrWhiteSpace(name)) name = (string)node.Attribute("name");
        return String.IsNullOrEmpty(id) || String.IsNullOrWhiteSpace(name) ? null :
            new BookmarkNotebook(id, name, ReadColor((string)node.Attribute("color")));
    }
    internal static Dictionary<string, BookmarkNotebook> IndexSections(XElement hierarchy)
    {
        var sections = new Dictionary<string, BookmarkNotebook>();
        foreach (var node in hierarchy.DescendantsAndSelf().Where(n => n.Name.LocalName == "Notebook"))
        {
            var notebook = FromXml(node);
            if (notebook == null) continue;
            foreach (var section in node.Descendants().Where(n => n.Name.LocalName == "Section"))
            {
                string sectionId = (string)section.Attribute("ID");
                if (!String.IsNullOrEmpty(sectionId)) sections[sectionId] = notebook;
            }
        }
        return sections;
    }
}

// Menu-item layout gives the heading the full width of the dropdown, including the icon column.
internal sealed class BookmarkHeader : ToolStripMenuItem
{
    internal readonly BookmarkNotebook Notebook;
    internal BookmarkHeader(BookmarkNotebook notebook)
    {
        Notebook = notebook;
        Text = AccessibleName = notebook.Name;
        AccessibleRole = AccessibleRole.StaticText;
        Enabled = false;
    }
    public override bool CanSelect { get { return false; } }
    protected override void OnPaint(PaintEventArgs e)
    {
        var palette = (TabRenderer)Owner.Renderer;
        Color fill = palette.NotebookFill(Notebook.Color);
        Color ink = palette.NotebookInk(fill);
        e.Graphics.Clear(fill);
        int size = HostChrome.Px(16, palette.Dpi), x = HostChrome.Px(7, palette.Dpi), y = (Height - size) / 2;
        using (var pen = new Pen(ink, Math.Max(1, HostChrome.Px(1, palette.Dpi))))
        {
            e.Graphics.DrawRectangle(pen, x, y, size * 3 / 4, size - 1);
            e.Graphics.DrawLine(pen, x + size / 4, y, x + size / 4, y + size - 1);
            for (int i = 1; i <= 3; i++)
                e.Graphics.DrawLine(pen, x + size * 3 / 4, y + i * size / 5, x + size - 1, y + i * size / 5);
        }
        int textLeft = x + size + HostChrome.Px(7, palette.Dpi);
        TextRenderer.DrawText(e.Graphics, Text, Font, new Rectangle(textLeft, 0, Math.Max(0, Width - textLeft - x), Height),
            ink, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
    }
}
