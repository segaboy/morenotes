using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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
}

internal sealed class BookmarkLocation
{
    internal readonly BookmarkNotebook Notebook;
    internal readonly string SectionName;
    internal readonly Color SectionColor;

    internal BookmarkLocation(BookmarkNotebook notebook, string name, Color color)
    {
        Notebook = notebook;
        SectionName = String.IsNullOrWhiteSpace(name) ? null : name;
        SectionColor = color;
    }
    internal bool SameAs(BookmarkLocation other)
    {
        return other != null && Notebook.SameAs(other.Notebook) && SectionName == other.SectionName &&
            SectionColor.ToArgb() == other.SectionColor.ToArgb();
    }
    internal XElement SectionXml()
    {
        return SectionName == null ? null : new XElement("section", new XAttribute("name", SectionName),
            SectionColor.IsEmpty ? null : new XAttribute("color", "#" + (SectionColor.ToArgb() & 0xffffff).ToString("X6")));
    }
    internal static BookmarkLocation FromXml(XElement page)
    {
        var notebook = BookmarkNotebook.FromXml(page.Element("notebook"));
        var section = page.Element("section");
        return notebook == null ? null : new BookmarkLocation(notebook, section == null ? null : (string)section.Attribute("name"),
            section == null ? Color.Empty : BookmarkNotebook.ReadColor((string)section.Attribute("color")));
    }
    internal static Dictionary<string, BookmarkLocation> IndexSections(XElement hierarchy)
    {
        var sections = new Dictionary<string, BookmarkLocation>();
        foreach (var node in hierarchy.DescendantsAndSelf().Where(n => n.Name.LocalName == "Notebook"))
        {
            var notebook = BookmarkNotebook.FromXml(node);
            if (notebook == null) continue;
            foreach (var section in node.Descendants().Where(n => n.Name.LocalName == "Section"))
            {
                string sectionId = (string)section.Attribute("ID");
                if (!String.IsNullOrEmpty(sectionId)) sections[sectionId] = new BookmarkLocation(notebook,
                    (string)section.Attribute("name"), BookmarkNotebook.ReadColor((string)section.Attribute("color")));
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

internal sealed class BookmarkRow : ToolStripMenuItem
{
    private const TextFormatFlags MeasureFlags = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
    private const TextFormatFlags DrawFlags = MeasureFlags | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis;
    internal readonly int SectionWidth;
    internal readonly string SectionName;
    internal Color SectionColor { get; private set; }
    private readonly string title;
    private readonly int padding;
    private readonly int slant;

    internal BookmarkRow(SavedPage page, BookmarkLocation location, int width, uint dpi, EventHandler click)
        : base(page.BookmarkTitle.Replace("&", "&&"), null, click)
    {
        title = page.BookmarkTitle;
        SectionName = NameOfSection(location);
        SectionColor = location == null ? Color.Empty : location.SectionColor;
        SectionWidth = width;
        padding = HostChrome.Px(7, dpi);
        slant = HostChrome.Px(9, dpi);
        AccessibleName = title + ", " + SectionName;
        ToolTipText = SectionName + " — " + title;
    }
    private static string NameOfSection(BookmarkLocation location)
    {
        return location == null || location.SectionName == null ? "Unknown section" : location.SectionName;
    }
    internal static int MeasureColumn(IEnumerable<BookmarkLocation> locations, Font font, uint dpi, int availableWidth)
    {
        int text = locations.Select(NameOfSection).Distinct().Max(name =>
            TextRenderer.MeasureText(name, font, Size.Empty, MeasureFlags).Width);
        return Math.Min(text + 2 * HostChrome.Px(7, dpi) + HostChrome.Px(9, dpi), availableWidth / 2);
    }
    internal int ContentWidth
    {
        get { return SectionWidth + 2 * padding + TextRenderer.MeasureText(title, Font, Size.Empty, MeasureFlags).Width; }
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        var palette = (TabRenderer)Owner.Renderer;
        Color fill = palette.SectionFill(SectionColor);
        Color titleFill = palette.MutedSection(SectionColor);
        Color titleInk = palette.NotebookInk(titleFill);
        if (palette.HighContrast && !Selected)
        {
            titleFill = palette.Surface;
            titleInk = palette.Ink;
        }
        e.Graphics.Clear(titleFill);
        var state = e.Graphics.Save();
        try
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var brush = new SolidBrush(fill))
                e.Graphics.FillPolygon(brush, new[] { Point.Empty, new Point(SectionWidth, 0),
                    new Point(SectionWidth - slant, Height), new Point(0, Height) });
        }
        finally { e.Graphics.Restore(state); }
        TextRenderer.DrawText(e.Graphics, SectionName, Font, new Rectangle(padding, 0, Math.Max(0, SectionWidth - 2 * padding - slant), Height),
            palette.NotebookInk(fill), DrawFlags);
        int left = SectionWidth + padding;
        TextRenderer.DrawText(e.Graphics, title, Font, new Rectangle(left, 0, Math.Max(0, Width - left - padding), Height),
            titleInk, DrawFlags);
        if (Selected) ControlPaint.DrawFocusRectangle(e.Graphics, new Rectangle(left - padding, 1,
            Math.Max(1, Width - left + padding - 2), Math.Max(1, Height - 2)), titleInk, titleFill);
    }
}
