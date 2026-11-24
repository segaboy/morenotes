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
        var custom = Color.FromArgb(255, 0, 0);
        var customDark = SectionPalette.Native(custom, true);
        Check(customDark.R >= 63 && customDark.R <= 64 && customDark.G == 0 && customDark.B == 0 &&
            SectionPalette.Native(custom, false).ToArgb() == custom.ToArgb() &&
            SectionPalette.DarkPalette(Color.FromArgb(142, 228, 234)) == false,
            "Custom section RGB or the light native palette was not recognized.");
        var labels = strip.Items.OfType<PageTabButton>().Select(tab => tab.Text).ToArray();
        Check(labels.SequenceEqual(new[] { "Dashboard", "Schedule", "Dashboard" }) &&
            strip.Items.OfType<PageTabButton>().Select(t => t.Group.Text).SequenceEqual(new[] { "Atlas", "Atlas", "Workshop" }),
            "Interleaved pages must form notebook groups that distinguish duplicate page titles.");
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
                    using (var band = new SolidBrush(Color.FromArgb(0, 97, 104)))
                        g.FillRectangle(band, 0, row.Height - 3, row.Width, 3);
                    Check(SectionPalette.DarkPalette(HostChrome.ReadSectionBand(row, dpi, surface)) == true, "Native dark band was missed.");
                    g.Clear(surface);
                    Check(HostChrome.ReadSectionBand(row, dpi, surface).IsEmpty, "A blank navigation row produced a section color.");
                    Check(!HostChrome.FindIcon(row, dpi, out measured, out left), "An empty row produced a false icon anchor.");
                }
                strip.Host.Size = new Size(HostChrome.Px(950, dpi), HostChrome.StripHeight(dpi));
                strip.ApplyAppearance(expected);
                foreach (var selected in new[] { schedule, original })
                {
                    selected.PerformClick();
                    using (var bitmap = CheckStrip(strip, expected))
                    {
                        if (contrast) continue;
                        int stripe = selected == schedule ? (expected.DarkSections ? 0x006168 : 0x8EE4EA) :
                            (expected.DarkSections ? 0xC44A01 : 0xFFC39F);
                        int xSample = selected.Bounds.Left + HostChrome.Px(6, dpi);
                        Check((bitmap.GetPixel(xSample, FlatEdgeY(selected, strip.AtBottom, 1)).ToArgb() & 0xffffff) == stripe &&
                            bitmap.GetPixel(xSample, FlatEdgeY(selected, strip.AtBottom, HostChrome.Px(4, dpi))).ToArgb() == selected.SectionColor.ToArgb(),
                            "The active tab must pair the native stripe with its original muted fill.");
                        var title = Rectangle.Inflate(selected.Bounds, -HostChrome.Px(10, dpi), -HostChrome.Px(7, dpi));
                        bool foundText = false;
                        for (int y = title.Top; y < title.Bottom && !foundText; y++)
                            for (int x = title.Left; x < title.Right && !foundText; x++)
                            {
                                Color ink = bitmap.GetPixel(x, y);
                                foundText = Math.Max(ink.R, Math.Max(ink.G, ink.B)) <= 45;
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
        Console.WriteLine("PASS equal native-color tab stripes, muted selected fills and angled two-color bookmark rows at 96, 120, 144 and 192 DPI.");
        Console.WriteLine("PASS compact notebook captions span the visible tabs in each group and use the notebook color at all four scales.");
        Console.WriteLine("PASS bookmark groups sort by notebook, retain page order and align section-color labels and page titles across groups.");
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
                var rows = menu.Items.OfType<BookmarkRow>().ToArray();
                Check(rows.Select(r => r.SectionName).SequenceEqual(new[] { "Route plans", "Sketches", "Materials & tools" }) &&
                    rows.Select(r => r.Bounds.Left + r.SectionWidth).Distinct().Count() == 1,
                    "Bookmark page titles are not aligned across notebook groups.");
                var remove = (ToolStripMenuItem)menu.Items[menu.Items.Count - 1];
                Check(remove.DropDownItems.OfType<BookmarkRow>().All(r => r.SectionWidth == rows[0].SectionWidth),
                    "The removal menu does not use the same section column.");
                foreach (var row in rows)
                {
                    int x = row.Bounds.Left + row.SectionWidth, y = row.Bounds.Top + 1;
                    int index = Array.IndexOf(rows, row);
                    int native = chrome.DarkSections ? new[] { 0x006168, 0x653A98, 0x2A7931 }[index] :
                        new[] { 0x8EE4EA, 0xD6BEF2, 0x8FE497 }[index];
                    var fill = chrome.HighContrast ? SystemColors.Highlight : Color.FromArgb(255, Color.FromArgb(native));
                    var titleFill = chrome.HighContrast ? (row.Selected ? SystemColors.Highlight : SystemColors.Window) : row.SectionColor;
                    int cut = x - HostChrome.Px(4, chrome.Dpi);
                    Check(image.GetPixel(row.Bounds.Left + 2, y).ToArgb() == fill.ToArgb() &&
                        image.GetPixel(cut, y).ToArgb() == fill.ToArgb() &&
                        image.GetPixel(cut, row.Bounds.Bottom - 2).ToArgb() == titleFill.ToArgb() &&
                        image.GetPixel(x + 3, y + 1).ToArgb() == titleFill.ToArgb(),
                        "The bookmark must have a native section color, a forward-slanted edge and a muted page fill.");
                    int titleWidth = TextRenderer.MeasureText(row.Text, row.Font, Size.Empty,
                        TextFormatFlags.NoPadding | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width;
                    int visibleWidth = Rectangle.Intersect(menu.ClientRectangle, row.Bounds).Width;
                    Check(visibleWidth >= row.SectionWidth + 2 * HostChrome.Px(7, chrome.Dpi) + titleWidth,
                        "Bookmark at " + chrome.Dpi + " DPI needs " + row.ContentWidth + " pixels. Item " + row.Bounds +
                        ", menu " + menu.ClientSize + ", minimum " + menu.MinimumSize + ".");
                }
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

    private static int FlatEdgeY(ToolStripItem tab, bool bottom, int distance)
    {
        return bottom ? tab.Bounds.Bottom - 1 - distance : tab.Bounds.Top + distance;
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
        bool bottom = strip.AtBottom;
        Check(bottom ? first.Bounds.Bottom == strip.Height && active.Bounds.Bottom == strip.Height :
            first.Bounds.Top == 0 && active.Bounds.Top == 0, "Tabs are not flush with the docked edge.");
        Check(star.Bounds.Top == first.Bounds.Top && star.Height == first.Height,
            "Bookmark and page tabs have different vertical bounds: " + star.Bounds + " and " + first.Bounds);
        var bitmap = new Bitmap(strip.Width, strip.Height);
        try
        {
            strip.DrawToBitmap(bitmap, strip.ClientRectangle);
            Rectangle b = active.Bounds;
            int edge = HostChrome.Px(1, chrome.Dpi);
            var band = strip.Host.Captions;
            Check(band.Left == strip.Left && (bottom ? band.Bottom == strip.Top : band.Top == strip.Bottom),
                "The notebook captions must sit on the side of the tabs facing the page.");
            using (var captions = new Bitmap(band.Width, band.Height))
            {
                band.DrawToBitmap(captions, band.ClientRectangle);
                foreach (var group in strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main).GroupBy(t => t.Group))
                {
                    var label = group.Key;
                    var box = label.Bounds;
                    Check(label.Visible && !label.CanSelect && label.AccessibleName == label.Text && label.Font.Style == FontStyle.Regular &&
                        label.Font.Size < first.Font.Size && box.Left == group.First().Bounds.Left && box.Right == group.Last().Bounds.Right &&
                        box.Height == label.Font.Height, "A notebook caption did not fit its tabs on one compact line.");
                    Color ink = chrome.HighContrast || label.Notebook.Color.IsEmpty ? palette.Ink : label.Notebook.Color;
                    Check(captions.GetPixel(box.Right - 1, box.Top).ToArgb() == palette.Surface.ToArgb(),
                        "Notebook labels should use the strip background.");
                    bool found = false;
                    for (int x = box.Left; x < box.Right && !found; x++)
                        for (int y = box.Top; y < box.Bottom && !found; y++)
                            found = captions.GetPixel(x, y).ToArgb() == ink.ToArgb();
                    Check(found, "The notebook caption did not render in its notebook color.");
                }
            }
            Color starBackground = star.Pressed ? palette.Active : star.Selected ? palette.Hover : palette.Tab;
            int starCenter = star.Bounds.Left + star.Width / 2;
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + star.Height / 2).ToArgb() == Color.Gold.ToArgb(),
                "The bookmark star must have a solid yellow center.");
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + 2 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(starCenter, star.Bounds.Bottom - 3 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Left, FlatEdgeY(star, !bottom, 0)).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Right - 1, FlatEdgeY(star, !bottom, 0)).ToArgb() == palette.Surface.ToArgb(),
                "The bookmark star must have the tab fill and rounded corners facing the editor.");
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
                Color marker = palette.MarkerFor(tab);
                for (int x = tab.Bounds.Left; x < tab.Bounds.Right; x++)
                    for (int y = 0; y < HostChrome.Px(3, chrome.Dpi); y++)
                        Check(bitmap.GetPixel(x, FlatEdgeY(tab, bottom, y)).ToArgb() == marker.ToArgb(),
                            "Missing full-width section color on " + (tab.Checked ? "active" : "inactive") + " tab.");
                Color fill = tab.Checked ? palette.MutedSection(tab.SectionColor) : tab.Selected ? palette.Hover : palette.Tab;
                Check(fill.ToArgb() != marker.ToArgb() && bitmap.GetPixel(tab.Bounds.Left + 2 * edge,
                    FlatEdgeY(tab, bottom, HostChrome.Px(3, chrome.Dpi))).ToArgb() == fill.ToArgb(), "Incorrect section stripe thickness.");
                Check(bitmap.GetPixel(tab.Bounds.Left + 2 * edge, FlatEdgeY(tab, bottom, HostChrome.Px(3, chrome.Dpi) + edge)).ToArgb() == fill.ToArgb() &&
                    bitmap.GetPixel(tab.Bounds.Left + tab.Width / 2, FlatEdgeY(tab, !bottom, 2 * edge)).ToArgb() == fill.ToArgb(),
                    "Tab body fill does not match its selection state.");
            }
            Check(bitmap.GetPixel(b.Left + edge - 1, FlatEdgeY(active, !bottom, 0)).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(b.Right - edge, FlatEdgeY(active, !bottom, 0)).ToArgb() == palette.Surface.ToArgb(),
                "The active tab's corners facing the editor are not rounded.");
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
        if (!chrome.HighContrast && !chrome.SectionBand.IsEmpty && !active.SectionColor.IsEmpty)
            Check(((TabRenderer)strip.Renderer).MarkerFor(active).ToArgb() == chrome.SectionBand.ToArgb(),
                "The selected tab stripe differs from OneNote's visible section band.");
        TogglePin(strip, schedule, true);
        using (CheckStrip(strip, chrome)) { }
        using (var bitmap = new Bitmap(strip.Host.Width, strip.Host.Height))
        {
            strip.Host.DrawToBitmap(bitmap, strip.Host.ClientRectangle);
            bitmap.Save(Path.Combine(fixtureDirectory, "MoreTabs-strip.png"), ImageFormat.Png);
        }
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
