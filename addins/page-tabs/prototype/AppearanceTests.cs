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
        var original = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        TogglePin(strip, schedule, true);
        var dark = Color.FromArgb(17, 12, 8);
        foreach (uint dpi in new uint[] { 96, 120, 144, 192 })
            foreach (var surface in new[] { dark, Color.FromArgb(247, 247, 247), SystemColors.Window })
            {
                bool contrast = surface.IsSystemColor;
                int inset = HostChrome.Px(contrast ? 4 : surface == dark ? 25 : 18, dpi);
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
                foreach (var selected in new[] { schedule, original })
                {
                    selected.PerformClick();
                    using (var bitmap = CheckStrip(strip, expected))
                    {
                        if (contrast) continue;
                        var title = Rectangle.Inflate(selected.Bounds, -HostChrome.Px(10, dpi), -HostChrome.Px(7, dpi));
                        bool lightText = selected == schedule; // Teal needs light text; orange needs dark text.
                        bool foundText = false;
                        for (int y = title.Top; y < title.Bottom && !foundText; y++)
                            for (int x = title.Left; x < title.Right && !foundText; x++)
                            {
                                Color ink = bitmap.GetPixel(x, y);
                                foundText = lightText ? Math.Min(ink.R, Math.Min(ink.G, ink.B)) >= 225 :
                                    Math.Max(ink.R, Math.Max(ink.G, ink.B)) <= 45;
                            }
                        Check(foundText, "The selected tab title did not contrast against its section color.");
                    }
                }
                using (var menuImage = CheckBookmarkMenu(strip, expected)) { }
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
        Console.WriteLine("PASS section stripes on every page tab and a full section-color fill on the selected tab, across themes and 96, 120, 144 and 192 DPI.");
        Console.WriteLine("PASS bookmark groups sort by notebook, retain page order and draw labeled full-width notebook color headers.");
    }

    internal static Bitmap CheckBookmarkMenu(PageStrip strip, HostChrome chrome)
    {
        var pages = new[] { new SavedPage("b/design", "Design notes / B", ""), new SavedPage("a/maps", "Maps / Summer / A", ""),
            new SavedPage("b/parts", "Parts list / B", "") };
        foreach (var page in pages) strip.SaveBookmark(page);
        var star = strip.Items.OfType<BookmarkButton>().Single();
        try
        {
            star.ShowDropDown();
            var menu = star.DropDown;
            var labels = menu.Items.Cast<ToolStripItem>().Select(i => i.Text).ToArray();
            Check(labels.SequenceEqual(new[] { "Bookmark current page", "Atlas", "Maps / Summer", "Workshop", "Design notes", "Parts list", "", "Remove bookmark" }),
                "Bookmark group order, page order or titles were incorrect.");
            var image = new Bitmap(menu.Width, menu.Height);
            try
            {
                menu.DrawToBitmap(image, menu.ClientRectangle);
                foreach (var header in menu.Items.OfType<BookmarkHeader>())
                {
                    var bounds = header.Bounds;
                    Color fill = chrome.HighContrast ? SystemColors.Highlight : header.Notebook.Color;
                    int y = bounds.Top + bounds.Height / 2;
                    Check(!header.CanSelect && header.AccessibleName == header.Text &&
                        image.GetPixel(2, y).ToArgb() == fill.ToArgb() && image.GetPixel(menu.Width - 3, y).ToArgb() == fill.ToArgb(),
                        "The notebook heading must be nonselectable and fill the menu width.");
                    for (int part = 0; part <= 1; part++)
                    {
                        int left = bounds.Left + HostChrome.Px(part == 0 ? 6 : 30, chrome.Dpi);
                        int right = part == 0 ? left + HostChrome.Px(18, chrome.Dpi) : bounds.Right - 3;
                        bool ink = false;
                        for (int x = left; x < right && !ink; x++)
                            for (int py = bounds.Top + 2; py < bounds.Bottom - 2 && !ink; py++)
                                ink = image.GetPixel(x, py).ToArgb() != fill.ToArgb();
                        Check(ink, part == 0 ? "Notebook heading icon missing." : "Notebook heading text missing.");
                    }
                }
                return image;
            }
            catch { image.Dispose(); throw; }
        }
        finally
        {
            star.HideDropDown();
            foreach (var page in pages) strip.RemoveBookmark(page);
        }
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
        Check(first.Bounds.Top == 0 && active.Bounds.Top == 0, "Tabs are not flush with the toolbar.");
        Check(star.Bounds.Top == first.Bounds.Top && star.Height == first.Height,
            "Bookmark and page tabs have different vertical bounds: " + star.Bounds + " and " + first.Bounds);
        var bitmap = new Bitmap(strip.Width, strip.Height);
        try
        {
            strip.DrawToBitmap(bitmap, strip.ClientRectangle);
            Rectangle b = active.Bounds;
            int edge = HostChrome.Px(1, chrome.Dpi);
            Color starBackground = star.Pressed ? palette.Active : star.Selected ? palette.Hover : palette.Tab;
            int starCenter = star.Bounds.Left + star.Width / 2;
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + star.Height / 2).ToArgb() == Color.Gold.ToArgb(),
                "The bookmark star must have a solid yellow center.");
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + 2 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(starCenter, star.Bounds.Bottom - 3 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Left, star.Bounds.Bottom - 1).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Right - 1, star.Bounds.Bottom - 1).ToArgb() == palette.Surface.ToArgb(),
                "The bookmark star must have the tab fill and rounded lower corners.");
            // Read through the two side points, ignoring very faint antialiasing pixels.
            int armY = star.Bounds.Top + star.Height / 2 - HostChrome.Px(2, chrome.Dpi);
            int starLeft = -1;
            for (int x = star.Bounds.Left + 1; x < star.Bounds.Right && starLeft < 0; x++)
                for (int y = armY - edge; y <= armY + edge; y++)
                {
                    Color c = bitmap.GetPixel(x, y);
                    if (Math.Max(Math.Abs(c.R - starBackground.R), Math.Max(Math.Abs(c.G - starBackground.G),
                        Math.Abs(c.B - starBackground.B))) < 38) continue;
                    starLeft = x;
                    break;
                }
            Check(starLeft == chrome.Inset, "Bookmark glyph starts at X=" + starLeft + ", expected notebook icon X=" + chrome.Inset);
            foreach (var divider in dividers)
            {
                int x = divider.Bounds.Left + divider.Width / 2;
                var left = strip.Items[strip.Items.IndexOf(divider) - 1];
                var right = strip.Items[strip.Items.IndexOf(divider) + 1];
                int y = divider.Bounds.Top + divider.Height / 2;
                if (left.Placement == ToolStripItemPlacement.Main && right.Placement == ToolStripItemPlacement.Main)
                {
                    int leftGap = x - left.Bounds.Right, rightGap = right.Bounds.Left - x - 1;
                    Check(leftGap > 0 && leftGap == rightGap, "Divider gaps in pixels: left " + leftGap + ", right " + rightGap);
                    for (int gap = 1; gap <= leftGap; gap++)
                        Check(bitmap.GetPixel(x - gap, y).ToArgb() == palette.Surface.ToArgb() &&
                            bitmap.GetPixel(x + gap, y).ToArgb() == palette.Surface.ToArgb(), "A divider gap contains painted pixels.");
                }
                Check(bitmap.GetPixel(x, y).ToArgb() != palette.Surface.ToArgb() &&
                    bitmap.GetPixel(x, divider.Bounds.Top + edge).ToArgb() == palette.Surface.ToArgb(),
                    "Missing or full-height group divider at " + divider.Bounds);
            }
            foreach (var tab in strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main))
            {
                Color marker = tab.SectionColor.IsEmpty ? palette.ActiveMarker : tab.SectionColor;
                if (chrome.HighContrast)
                    marker = tab.Checked ? SystemColors.Highlight : tab.Selected ? SystemColors.HighlightText : SystemColors.WindowText;
                for (int x = tab.Bounds.Left; x < tab.Bounds.Right; x++)
                    Check(bitmap.GetPixel(x, tab.Bounds.Top + edge).ToArgb() == marker.ToArgb(),
                        "Missing full-width section color on " + (tab.Checked ? "active" : "inactive") + " tab.");
                Color fill = tab.Checked ? marker : tab.Selected ? palette.Hover : palette.Tab;
                Check(bitmap.GetPixel(tab.Bounds.Left + 2 * edge, tab.Bounds.Top + HostChrome.Px(3, chrome.Dpi) + edge).ToArgb() == fill.ToArgb() &&
                    bitmap.GetPixel(tab.Bounds.Left + tab.Width / 2, tab.Bounds.Bottom - 3 * edge).ToArgb() == fill.ToArgb(),
                    "Tab body fill does not match its selection state.");
            }
            Check(bitmap.GetPixel(b.Left + edge - 1, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(b.Right - edge, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb(),
                "The active tab's lower corners are not rounded.");
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
        Console.WriteLine("PASS tab rendering in OneNote: surface #{0:X6}, bookmark star X={1}px ({2}), DPI {3}; square top and rounded lower corners.",
            chrome.Surface.ToArgb() & 0xffffff, chrome.Inset, chrome.Measured ? "measured notebook icon" : "fallback inset", chrome.Dpi);
    }
    internal static void TogglePin(PageStrip strip, PageTabButton button, bool pin)
    {
        var menu = strip.TabMenuFor((SavedTab)button.Tag);
        menu.Items.Cast<ToolStripItem>().Single(item => item.Text == (pin ? "Pin tab" : "Unpin tab")).PerformClick();
        Check(((SavedTab)button.Tag).Pinned == pin, "Tab menu did not change the pin.");
    }
}
