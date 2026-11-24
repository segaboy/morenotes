using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

internal static class AppearanceTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    internal static void Controlled(PageStrip strip, NativeLayout layout, Rectangle slot)
    {
        var labels = strip.Items.OfType<PageTabButton>().Select(tab => tab.Text).ToArray();
        Check(labels.SequenceEqual(new[] { "Dashboard / A", "Schedule", "Dashboard / B" }),
            "Tab labels must distinguish duplicate page titles.");
        Check(strip.Items.OfType<ToolStripSeparator>().Count(s => s.Available) == 1, "An empty pinned group has two dividers.");
        var schedule = strip.Items.OfType<PageTabButton>().ElementAt(1);
        TogglePin(strip, schedule, true);
        var dark = Color.FromArgb(17, 12, 8);
        foreach (uint dpi in new uint[] { 96, 144 })
            foreach (var surface in new[] { dark, Color.FromArgb(247, 247, 247), SystemColors.Window })
            {
                bool contrast = surface.IsSystemColor;
                int inset = HostChrome.Px(22, dpi);
                var expected = new HostChrome { Dpi = dpi, Surface = surface, Inset = inset, HighContrast = contrast };
                using (var row = new Bitmap(HostChrome.Px(48, dpi), HostChrome.Px(43, dpi)))
                using (var g = Graphics.FromImage(row))
                using (var icon = new SolidBrush(contrast ? SystemColors.WindowText : Color.FromArgb(0, 155, 199)))
                {
                    g.Clear(surface);
                    g.FillRectangle(icon, inset, HostChrome.Px(9, dpi), HostChrome.Px(14, dpi), HostChrome.Px(20, dpi));
                    Color measured;
                    int left;
                    Check(HostChrome.FindIcon(row, dpi, out measured, out left) && left == inset && measured.ToArgb() == surface.ToArgb(),
                        "Did not measure the notebook glyph or the row color.");
                    g.Clear(surface);
                    Check(!HostChrome.FindIcon(row, dpi, out measured, out left), "An empty row produced a false icon anchor.");
                }
                strip.Size = new Size(HostChrome.Px(950, dpi), HostChrome.Px(32, dpi));
                strip.ApplyAppearance(expected);
                using (var bitmap = CheckStrip(strip, expected)) { }
            }
        TogglePin(strip, schedule, false);
        Check(labels.SequenceEqual(strip.Items.OfType<PageTabButton>().Select(t => t.Text)),
            "Unpinning changed the relative order of the regular tabs.");
        var active = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        TogglePin(strip, active, true);
        Check(active.Checked && active == strip.Items.OfType<PageTabButton>().First(),
            "Pinning the active tab lost its selection or left it in the regular group.");
        TogglePin(strip, active, false);
        Check(active.Checked && strip.Items.OfType<ToolStripSeparator>().Count(s => s.Available) == 1,
            "Unpinning the active tab lost its selection or left an extra divider.");
        strip.Attach(slot, layout);
        Console.WriteLine("PASS rendered dark, light and high-contrast tabs at 96 and 144 DPI: bookmark star, dividers, pinned group, active tab, lower corners and icon alignment.");
    }

    private static Bitmap CheckStrip(PageStrip strip, HostChrome chrome)
    {
        strip.PerformLayout();
        var first = strip.Items.OfType<PageTabButton>().First(b => b.Placement == ToolStripItemPlacement.Main);
        var star = strip.Items[0] as BookmarkButton;
        var dividers = strip.Items.OfType<ToolStripSeparator>().Where(s => s.Available).ToArray();
        Check(star != null && star.Available && star.AccessibleName == "Bookmarks" && !star.ShowDropDownArrow &&
            star.DisplayStyle == ToolStripItemDisplayStyle.None && dividers.Length == 2,
            "Expected an icon-only bookmark menu and two group dividers.");
        Check(star.Bounds.Right <= dividers[0].Bounds.Left && dividers[0].Bounds.Right <= first.Bounds.Left,
            "Bookmark menu and first divider must precede the tabs.");
        foreach (var tab in strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main))
            Check(((SavedTab)tab.Tag).Pinned ? tab.Bounds.Right <= dividers[1].Bounds.Left : tab.Bounds.Left >= dividers[1].Bounds.Right,
                "A tab appeared in the wrong group.");
        var active = strip.Items.OfType<PageTabButton>().Single(b => b.Checked);
        var palette = (TabRenderer)strip.Renderer;
        Check(first.Bounds.Left == chrome.Inset && first.Bounds.Top == 0 && active.Bounds.Top == 0,
            "Tabs are not flush with the toolbar or notebook icon: tab " + first.Bounds + ", icon X=" + chrome.Inset);
        var bitmap = new Bitmap(strip.Width, strip.Height);
        try
        {
            strip.DrawToBitmap(bitmap, strip.ClientRectangle);
            Rectangle b = active.Bounds;
            int edge = HostChrome.Px(1, chrome.Dpi);
            int ink = 0;
            for (int dx = 0; dx < star.Width; dx++)
                for (int dy = 0; dy < star.Height; dy++)
                    if (bitmap.GetPixel(star.Bounds.Left + dx, star.Bounds.Top + dy).ToArgb() != palette.Surface.ToArgb()) ink++;
            Check(ink >= HostChrome.Px(20, chrome.Dpi), "The bookmark star was not drawn.");
            foreach (var divider in dividers)
            {
                int x = divider.Bounds.Left + divider.Width / 2;
                Check(bitmap.GetPixel(x, divider.Bounds.Top + divider.Height / 2).ToArgb() != palette.Surface.ToArgb() &&
                    bitmap.GetPixel(x, divider.Bounds.Top + edge).ToArgb() == palette.Surface.ToArgb(),
                    "Missing or full-height group divider at " + divider.Bounds);
            }
            foreach (var tab in strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main))
            {
                int pixel = bitmap.GetPixel(tab.Bounds.Left + tab.Width / 2, tab.Bounds.Top + edge).ToArgb();
                var expected = tab.Checked ? palette.ActiveMarker : tab.Selected ? palette.Hover : palette.Tab;
                Check(pixel == expected.ToArgb() && palette.ActiveMarker.ToArgb() != palette.Active.ToArgb(),
                    "The persistent accent must appear on exactly the active tab.");
            }
            Check(bitmap.GetPixel(b.Left + 2 * edge, b.Top + 2 * edge).ToArgb() == palette.Active.ToArgb(),
                "The active tab has a gap or rounding at its upper-left corner.");
            Check(bitmap.GetPixel(b.Left + edge - 1, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(b.Right - edge, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb(),
                "The active tab's lower corners are not rounded.");
            Check(bitmap.GetPixel(b.Left + b.Width / 2, b.Bottom - 3 * edge).ToArgb() == palette.Active.ToArgb(),
                "The active tab did not use its theme color.");
            return bitmap;
        }
        catch { bitmap.Dispose(); throw; }
    }

    internal static void Live(PageStrip strip, NativeLayout layout, string fixtureDirectory)
    {
        var chrome = layout.Appearance(true);
        strip.ApplyAppearance(chrome);
        var schedule = strip.Items.OfType<PageTabButton>().ElementAt(1);
        var active = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        TogglePin(strip, schedule, true);
        using (var bitmap = CheckStrip(strip, chrome))
            bitmap.Save(Path.Combine(fixtureDirectory, "MoreTabs-strip.png"), ImageFormat.Png);
        TogglePin(strip, schedule, false);
        Check(active.Checked && strip.Items.OfType<ToolStripSeparator>().Count(s => s.Available) == 1,
            "Unpinning switched the active tab or left an empty pinned group.");
        Console.WriteLine("PASS bookmark star, subtle vertical dividers and pin/unpin via the tab menu with the active tab preserved.");
        Console.WriteLine("PASS tab rendering in OneNote: surface #{0:X6}, first tab X={1}px ({2}), DPI {3}; square top and rounded lower corners.",
            chrome.Surface.ToArgb() & 0xffffff, chrome.Inset, chrome.Measured ? "measured notebook icon" : "fallback inset", chrome.Dpi);
    }
    internal static void TogglePin(PageStrip strip, PageTabButton button, bool pin)
    {
        var menu = strip.TabMenuFor((SavedTab)button.Tag);
        menu.Items.Cast<ToolStripItem>().Single(item => item.Text == (pin ? "Pin tab" : "Unpin tab")).PerformClick();
        Check(((SavedTab)button.Tag).Pinned == pin, "Tab menu did not change the pin.");
    }
}
