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
        CheckGroupDividers(strip, 2);
        var workshopLabel = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Single(l => l.Text == "Workshop");
        PrototypeTests.ClickNotebookPin(workshopLabel);
        Check(strip.Items.OfType<PageTabButton>().First().Group == workshopLabel, "A notebook pin failed to move its group at this dock edge.");
        PrototypeTests.ClickNotebookPin(workshopLabel);
        var schedule = strip.Items.OfType<PageTabButton>().ElementAt(1);
        var original = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        TogglePin(strip, schedule, true);
        TogglePin(strip, original, true);
        Check(schedule.Group == original.Group && schedule.Group.Text == "Pinned" && schedule.Group.Notebook == null,
            "Pins from different notebooks must share one Pinned caption.");
        TogglePin(strip, original, false);
        CheckCaptionAlignment(strip);
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
                if (!contrast)
                {
                    var ordinary = strip.Items.OfType<PageTabButton>().Select(t => t.Group).Where(g => g.Notebook != null);
                    var atlas = ordinary.Single(g => g.Text == "Atlas");
                    var workshop = ordinary.Single(g => g.Text == "Workshop");
                    Color blue = surface == dark ? Color.FromArgb(0, 136, 240) : Color.FromArgb(0, 79, 139);
                    Check(atlas.GroupFill.ToArgb() == blue.ToArgb() && workshop.GroupFill.ToArgb() == Color.FromArgb(198, 163, 119).ToArgb() &&
                        atlas.ForeColor.ToArgb() == (surface == dark ? Color.Black : Color.White).ToArgb() && workshop.ForeColor.ToArgb() == Color.Black.ToArgb(),
                        "Notebook bands must use the native blue and tan covers and contrast text against the displayed RGB.");
                    var palette = (TabRenderer)strip.Renderer;
                    Color purple = surface == dark ? Color.FromArgb(185, 115, 211) : Color.FromArgb(91, 45, 144);
                    Check(palette.NotebookCover(Color.FromArgb(171, 156, 183)).ToArgb() == purple.ToArgb() &&
                        palette.NotebookCover(Color.Orchid) == Color.Orchid && palette.NotebookCover(Color.Empty).IsEmpty,
                        "Purple Mist must use the cover for this theme; custom and missing notebook RGBs must round-trip.");
                }
                foreach (var selected in new[] { schedule, original })
                {
                    selected.PerformClick();
                    CheckLongTitleSpacing(strip, schedule, original);
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
                        title.Width = selected.Bounds.Right - selected.CloseWidth - title.Left;
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
        Check(active.Checked, "Unpinning the active tab lost its selection.");
        CheckGroupDividers(strip, 2);
        strip.Attach(slot, layout);
        CheckOptions(strip, layout);
        Console.WriteLine("PASS long regular and bold titles leave clear space around the close icon at all four scales.");
        Console.WriteLine("PASS equal native-color tab stripes, muted selected fills and angled two-color bookmark rows at 96, 120, 144 and 192 DPI.");
        Console.WriteLine("PASS rounded notebook backings, contrasting caption text and spacing at all four scales.");
        Console.WriteLine("PASS bookmark groups sort by notebook, retain page order and align section-color labels and page titles across groups.");
    }

    private static void CheckLongTitleSpacing(PageStrip strip, params PageTabButton[] tabs)
    {
        var titles = tabs.Select(tab => tab.Text).ToArray();
        try
        {
            // Both titles exceed the width limit, in regular and bold type.
            foreach (var tab in tabs) tab.Text = "Weekly planning and follow-up for the next review meeting";
            strip.PerformLayout();
            using (var image = new Bitmap(strip.Width, strip.Height))
            {
                strip.DrawToBitmap(image, strip.ClientRectangle);
                var palette = (TabRenderer)strip.Renderer;
                foreach (var tab in tabs)
                {
                    int fill = (tab.Checked ? palette.MutedSection(tab.SectionColor) : tab.Selected ? palette.Hover : palette.Tab).ToArgb();
                    int cy = tab.Bounds.Top + tab.Height / 2, right = tab.Bounds.Right;
                    int cx = right - (tab.CloseBounds.Width + 1) / 2;
                    // Check both sides of the actual painted cross, including the baseline of the ellipsis.
                    for (int y = cy - HostChrome.Px(5, tab.Dpi); y <= cy + HostChrome.Px(5, tab.Dpi); y++)
                        for (int dx = HostChrome.Px(6, tab.Dpi); dx <= HostChrome.Px(12, tab.Dpi); dx++)
                            Check(image.GetPixel(cx - dx, y).ToArgb() == fill && image.GetPixel(cx + dx, y).ToArgb() == fill,
                                String.Format("Close spacing at {0} DPI, bottom={1}, tab {2}, dx {3}, y {4}: {5:X8}/{6:X8}, expected {7:X8}.",
                                    tab.Dpi, strip.AtBottom, tab.Bounds, dx, y, image.GetPixel(cx-dx,y).ToArgb(), image.GetPixel(cx+dx,y).ToArgb(), fill));
                }
            }
        }
        finally { for (int i = 0; i < tabs.Length; i++) tabs[i].Text = titles[i]; }
    }

    private static void CheckOptions(PageStrip strip, NativeLayout layout)
    {
        var gear = (ToolStripDropDownButton)strip.Items[1];
        var star = (BookmarkButton)strip.Items[0];
        Check(gear.AccessibleName == "MoreTabs options" && gear.Bounds.Left == star.Bounds.Right &&
            strip.GetItemAt(star.Bounds.Right - 1, star.Bounds.Top + star.Height / 2) == star &&
            strip.GetItemAt(gear.Bounds.Left, gear.Bounds.Top + gear.Height / 2) == gear,
            "The halves of the menu tab must touch and provide separate click targets.");
        gear.ShowDropDown();
        var group = (ToolStripMenuItem)gear.DropDownItems[0];
        var labels = (ToolStripMenuItem)gear.DropDownItems[1];
        var colors = (ToolStripMenuItem)gear.DropDownItems[2];
        Check(group.Checked && labels.Checked && colors.Checked, "All three options should default to enabled.");
        gear.HideDropDown();
        var selected = strip.ActiveTabKey;
        var grouped = strip.Items.OfType<PageTabButton>().ToArray();
        var tools = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Single(l => l.Text == "Tools");
        colors.PerformClick();
        CheckGroupDividers(strip, 2);
        foreach (var surface in new[] { Color.White, Color.FromArgb(18,18,18) })
        {
            strip.ApplyAppearance(new HostChrome { Dpi = layout.Dpi, Surface = surface });
            var palette = (TabRenderer)strip.Renderer;
            Check(grouped.All(t => t.Group.GroupFill.IsEmpty && t.Group.ForeColor == palette.Ink) &&
                tools.GroupFill == palette.ToolsBand && tools.ForeColor == palette.NotebookInk(palette.ToolsBand),
                "Turning off notebook colors must preserve the neutral Tools band across theme changes.");
        }
        colors.PerformClick();
        CheckGroupDividers(strip, 2);
        using (var image = new Bitmap(strip.Width, strip.Height))
        {
            strip.DrawToBitmap(image, strip.ClientRectangle);
            var firstTab = grouped[0];
            int gap = grouped[1].Bounds.Left - firstTab.Bounds.Right;
            Check(firstTab.Group == grouped[1].Group && gap > 0 &&
                firstTab.Bounds.Left - firstTab.Group.Left == gap && firstTab.Group.Right - grouped[1].Bounds.Right == gap &&
                image.GetPixel(firstTab.Bounds.Right, strip.Height / 2).ToArgb() == Color.FromArgb(0, 136, 240).ToArgb(),
                "The colored band must leave exactly one inter-tab gap at each end of a two-tab group.");
        }
        TogglePin(strip, grouped[2], true);
        int fullDocument = layout.DocumentBounds.Height;
        labels.PerformClick();
        CheckGroupDividers(strip, 2);
        Check(!colors.Enabled && !colors.Checked && group.Checked && labels.Enabled && !labels.Checked &&
            !strip.Host.Captions.Visible && !grouped[2].Group.Visible && strip.Host.Height == HostChrome.StripHeight(layout.Dpi, false) &&
            layout.DocumentBounds.Height == fullDocument + HostChrome.CaptionHeight(layout.Dpi) &&
            layout.StillReserved(strip.Host.Bounds), String.Format(
                "Hiding captions with pins: colors {0}/{1}, grouping {2}, labels {3}/{4}, visible {5}/{6}, strip {7}px, document {8}px (was {9}px), reserved {10}.",
                colors.Enabled, colors.Checked, group.Checked, labels.Enabled, labels.Checked, strip.Host.Captions.Visible,
                grouped[2].Group.Visible, strip.Host.Height, layout.DocumentBounds.Height, fullDocument, layout.StillReserved(strip.Host.Bounds)));
        group.PerformClick();
        CheckGroupDividers(strip, 2);
        Check(!labels.Enabled && !colors.Enabled && strip.Items.OfType<PageTabButton>().First() == grouped[2],
            "Disabling grouping must keep pins first and disable the caption controls.");
        group.PerformClick();
        Check(labels.Enabled && !labels.Checked && !colors.Enabled && !strip.WantsCaptionRow,
            "Regrouping lost the preference to keep labels hidden.");
        labels.PerformClick();
        Check(grouped[2].Group.Visible && colors.Enabled && colors.Checked && layout.DocumentBounds.Height == fullDocument,
            "Showing labels must restore the Pinned caption, the color preference and the original document height.");
        TogglePin(strip, grouped[2], false);
        group.PerformClick();
        Check(strip.Items.OfType<PageTabButton>().SequenceEqual(new[] { grouped[0], grouped[2], grouped[1] }) &&
            !labels.Enabled && !labels.Checked && !colors.Enabled && !colors.Checked &&
            !strip.Host.Captions.Visible && layout.DocumentBounds.Height == fullDocument + HostChrome.CaptionHeight(layout.Dpi),
            "Ungrouping must restore opening order and remove the caption row.");
        CheckGroupDividers(strip, 1);
        group.PerformClick();
        Check(strip.ActiveTabKey == selected && strip.Items.OfType<PageTabButton>().SequenceEqual(grouped) &&
            labels.Checked && colors.Checked && strip.Host.Captions.Visible && layout.DocumentBounds.Height == fullDocument,
            "Regrouping must restore the labels and colors without changing the active page.");
        strip.ApplyAppearance(new HostChrome { Dpi = layout.Dpi, Surface = SystemColors.Window, HighContrast = true });
        Check(!colors.Enabled && !colors.Checked && labels.Enabled && labels.Checked, "High contrast must use system caption colors.");
        strip.Attach(layout.Reserve(), layout);
        Console.WriteLine("PASS shared star and gear tab, settings dependencies, remembered preferences and reclaimed document height with pins.");
    }

    internal static void CheckNotebookPicker(PageStrip strip, string longTitle)
    {
        var picker = strip.Items.OfType<PagePicker>().Single();
        foreach (uint dpi in new uint[] { 96, 144, 192 })
        foreach (var surface in new[] { Color.FromArgb(17, 12, 8), Color.WhiteSmoke, SystemColors.Window })
        {
            var chrome = new HostChrome { Dpi = dpi, Surface = surface, HighContrast = surface.IsSystemColor };
            strip.ApplyAppearance(chrome);
            picker.ShowDropDown();
            try
            {
                var menu = picker.DropDown;
                var rows = menu.Items.OfType<PickerRow>().ToArray();
                var palette = (TabRenderer)menu.Renderer;
                Check(rows.Length == 62 && rows[0].AccessibleName == "Atlas" && rows[1].AccessibleName == "Workshop" &&
                    rows[2].AccessibleName == longTitle && rows[2].ToolTipText == longTitle && menu.ShowItemToolTips &&
                    palette.Surface.ToArgb() == ((TabRenderer)strip.Renderer).NotebookMenu.Surface.ToArgb() &&
                    menu.Width == HostChrome.Px(220, dpi) && rows[0].Height == HostChrome.Px(28, dpi) &&
                    menu.Height < rows.Sum(row => row.Height) && rows.Last().HasDropDownItems,
                    "Notebook order, compact dimensions, long names or the scrollable list changed at " + dpi + " DPI.");
                using (var image = new Bitmap(menu.Width, menu.Height))
                {
                    menu.DrawToBitmap(image, menu.ClientRectangle);
                    Color blue = surface == Color.WhiteSmoke ? Color.FromArgb(0, 79, 139) : Color.FromArgb(0, 136, 240);
                    var expected = new[] { blue, palette.Ink, Color.FromArgb(49, 138, 117) };
                    for (int i = 0; i < expected.Length; i++)
                    {
                        var b = rows[i].Bounds;
                        int y = b.Top + b.Height / 2;
                        Color ink = palette.HighContrast && rows[i].Selected ? SystemColors.HighlightText : palette.Ink;
                        Color bg = rows[i].Selected ? palette.Hover : palette.Surface;
                        Check(image.GetPixel(b.Left + HostChrome.Px(14, dpi), y).ToArgb() ==
                            (palette.HighContrast ? ink : expected[i]).ToArgb() &&
                            image.GetPixel(b.Left + HostChrome.Px(32, dpi), y).ToArgb() == bg.ToArgb(),
                            "Missing notebook cover or neutral gap before the name at " + dpi + " DPI.");
                    }
                }
                rows[0].ShowDropDown();
                var sections = rows[0].DropDown;
                var section = sections.Items.OfType<PickerRow>().First();
                Check(sections.Renderer == palette && sections.Width == menu.Width && section.Height == rows[0].Height,
                    "The section submenu must match the notebook menu's theme and metrics.");
                using (var image = new Bitmap(sections.Width, sections.Height))
                {
                    sections.DrawToBitmap(image, sections.ClientRectangle);
                    var b = section.Bounds;
                    int x = b.Left + HostChrome.Px(17, dpi), y = b.Top + b.Height / 2;
                    Color color = palette.HighContrast ? (section.Selected ? SystemColors.HighlightText : palette.Ink) :
                        palette.DarkSections ? Color.FromArgb(0, 97, 104) : Color.FromArgb(142, 228, 234);
                    Color background = section.Selected ? palette.Hover : palette.Surface;
                    Check(image.GetPixel(x, y).ToArgb() == color.ToArgb() &&
                        image.GetPixel(x, y + HostChrome.Px(7, dpi)).ToArgb() == background.ToArgb(),
                        "The section marker lost its native section color or its notched edge.");
                }
            }
            finally { picker.HideDropDown(); }
        }
        Console.WriteLine("PASS compact notebook picker, native and custom covers, nicknames, long lists and high contrast at 96, 144 and 192 DPI.");
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
    internal static void CheckGroupDividers(PageStrip strip, int? expectedGroups = null)
    {
        var items = strip.Items.Cast<ToolStripItem>().Where(i => i.Available && i.Placement == ToolStripItemPlacement.Main &&
            (i is ToolStripSeparator || i is PageTabButton || i is PagePicker)).ToArray();
        var groups = items.OfType<PageTabButton>().Select(t => t.Group).Distinct().ToArray();
        Check(items.LastOrDefault() is PagePicker && (!expectedGroups.HasValue || groups.Length == expectedGroups.Value),
            "The visible tab groups must precede the page picker.");
        var options = (ToolStripDropDownButton)strip.Items[1];
        if (((ToolStripMenuItem)options.DropDownItems[2]).Checked)
        {
            Check(!strip.Items.OfType<ToolStripSeparator>().Any(s => s.Available), "Colored bands must replace every strip divider.");
            var tools = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Single(l => l.Text == "Tools");
            int right = tools.Right;
            int gap = 2 * HostChrome.Px(1, ((TabRenderer)strip.Renderer).Dpi);
            foreach (var group in groups)
            {
                Check(group.Left - right == gap, "A hidden divider left excess space before " + group.Text + ".");
                right = group.Right;
            }
            var next = strip.MoreTabs.Available ? strip.MoreTabs : items.Last();
            Check(next.Bounds.Left - right >= gap, "The trailing controls must clear the last notebook band.");
            return;
        }
        Check(items[0] is ToolStripSeparator && items.OfType<ToolStripSeparator>().Count() == 1 + groups.Length,
            "Tools, tab groups and the page picker must have dividers when colored labels are off.");
        for (int i = 1; i < items.Length - 1; i++)
        {
            var tab = items[i] as PageTabButton;
            if (tab == null) continue;
            var nextTab = items.Skip(i + 1).OfType<PageTabButton>().FirstOrDefault();
            bool endsGroup = nextTab == null || nextTab.Group != tab.Group;
            Check((items[i + 1] is ToolStripSeparator) == endsGroup, "A divider is missing or splits tabs from the same group.");
        }
    }
    private static void CheckPinAppearance(NotebookTabLabel label, Color fill, Color ink,
        int textTop, int textBottom, uint dpi)
    {
        var pin = label.Pin;
        var control = pin.Bounds;
        int difference = Math.Abs(ink.R - fill.R) + Math.Abs(ink.G - fill.G) + Math.Abs(ink.B - fill.B);
        Check(control.Height == label.Height && control.Width == HostChrome.Px(18, dpi), "The pin lost its full click target.");
        using (var band = new Bitmap(label.Width, label.Height))
        using (var pixels = new Bitmap(pin.Width, pin.Height))
        {
            bool original = pin.Checked;
            try
            {
                pin.Visible = false;
                label.DrawToBitmap(band, label.ClientRectangle);
                pin.Visible = true;
                foreach (bool active in new[] { false, true })
                {
                    pin.Checked = active;
                    pin.DrawToBitmap(pixels, pin.ClientRectangle);
                    int top = pin.Height, bottom = -1;
                    for (int y = 0; y < pin.Height; y++)
                        for (int x = 0; x < pin.Width; x++)
                        {
                            Color pixel = pixels.GetPixel(x, y);
                            if (y <= 1 || y >= pin.Height - 2 || x <= 1 || x >= pin.Width - 2)
                            {
                                Color background = band.GetPixel(control.Left + x, control.Top + y);
                                Check(pixel.ToArgb() == background.ToArgb(), String.Format(
                                    "Pin edge mismatch at {0} DPI on {1}, x={2}, y={3}: actual {4:X8}, band {5:X8}, fill {6:X8}, surface {7:X8}.",
                                    dpi, label.Text, x, y, pixel.ToArgb(), background.ToArgb(), fill.ToArgb(), label.BackColor.ToArgb()));
                            }
                            if (x < HostChrome.Px(3, dpi) || x >= pin.Width - HostChrome.Px(3, dpi)) continue;
                            if (Math.Abs(pixel.R-fill.R) + Math.Abs(pixel.G-fill.G) + Math.Abs(pixel.B-fill.B) < difference / 3) continue;
                            top = Math.Min(top, y);
                            bottom = Math.Max(bottom, y);
                        }
                    // GDI text and the antialiased pin can round the ends to adjacent pixels.
                    Check(bottom >= top && Math.Abs(top - textTop) <= 2 && Math.Abs(bottom - textBottom) <= 2 &&
                        Math.Abs((bottom - top) - (textBottom - textTop)) <= 2,
                        "Pin ink " + top + ".." + bottom + " must match " + label.Text + " ink " + textTop + ".." + textBottom + " at " + dpi + " DPI.");
                }
            }
            finally { pin.Checked = original; pin.Visible = true; }
        }
    }
    private static Bitmap CheckStrip(PageStrip strip, HostChrome chrome)
    {
        strip.PerformLayout();
        var first = strip.Items.OfType<PageTabButton>().First(b => b.Placement == ToolStripItemPlacement.Main);
        var star = strip.Items[0] as BookmarkButton;
        var gear = strip.Items[1];
        var dividers = strip.Items.OfType<ToolStripSeparator>().Where(s => s.Available).ToArray();
        Check(star != null && star.Available && star.AccessibleName == "Bookmarks" && !star.ShowDropDownArrow &&
            star.DisplayStyle == ToolStripItemDisplayStyle.None, "Expected an icon-only bookmark menu.");
        CheckGroupDividers(strip);
        Check(strip.MenuTabBounds.Right <= first.Bounds.Left, "Tools must precede the tabs.");
        var active = strip.Items.OfType<PageTabButton>().Single(b => b.Checked);
        var palette = (TabRenderer)strip.Renderer;
        bool bottom = strip.AtBottom;
        Check(bottom ? first.Bounds.Bottom == strip.Height && active.Bounds.Bottom == strip.Height :
            first.Bounds.Top == 0 && active.Bounds.Top == 0, "Tabs are not flush with the docked edge.");
        Check(star.Bounds.Top == first.Bounds.Top && star.Height == first.Height && star.Bounds.Right == gear.Bounds.Left &&
            star.Bounds.Top == gear.Bounds.Top && star.Height == gear.Height, "Menu tab halves must join across the full height.");
        var bitmap = new Bitmap(strip.Width, strip.Height);
        try
        {
            strip.DrawToBitmap(bitmap, strip.ClientRectangle);
            // The Tools buttons have no text or image fallback. Check their actual painted centers.
            foreach (var command in strip.Items.Cast<ToolStripItem>().Where(i => i == gear || i is ToolButton))
            {
                int size = HostChrome.Px(16, chrome.Dpi);
                int x0 = command.Bounds.Left + (command.Width - size) / 2;
                int y0 = command.Bounds.Top + (command.Height - size) / 2;
                var background = command.Pressed ? palette.Active : command.Selected ? palette.Hover : palette.Tab;
                int marks = 0;
                for (int y = y0; y < y0 + size; y++)
                    for (int x = x0; x < x0 + size; x++)
                    {
                        var pixel = bitmap.GetPixel(x, y);
                        if (Math.Abs(pixel.R - background.R) + Math.Abs(pixel.G - background.G) +
                            Math.Abs(pixel.B - background.B) >= 120) marks++;
                    }
                Check(marks >= size, command.AccessibleName + " icon is missing or too faint at " + chrome.Dpi + " DPI.");
            }
            Rectangle b = active.Bounds;
            int edge = HostChrome.Px(1, chrome.Dpi);
            var picker = strip.Items.OfType<PagePicker>().Single();
            int center = picker.Bounds.Left + picker.Width / 2;
            Color pickerFill = picker.Pressed ? palette.Active : picker.Selected ? palette.Hover : palette.Tab;
            Check(picker.Bounds.Top == first.Bounds.Top && picker.Height == first.Height &&
                bitmap.GetPixel(center, FlatEdgeY(picker, bottom, 2 * edge)).ToArgb() == pickerFill.ToArgb() &&
                bitmap.GetPixel(picker.Bounds.Left, FlatEdgeY(picker, !bottom, 0)).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(picker.Bounds.Right - 1, FlatEdgeY(picker, !bottom, 0)).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(center, picker.Bounds.Top + picker.Height / 2).ToArgb() != pickerFill.ToArgb(),
                "The new-tab button must have a centered plus and the same dock-facing shape as the page tabs.");
            var band = strip.Host.Captions;
            var tools = band.Controls.OfType<NotebookTabLabel>().Single(l => l.Text == "Tools" && l.Notebook == null);
            int toolsPad = 2 * edge;
            Color toolsFill = palette.ToolsBand;
            Color toolsInk = chrome.HighContrast ? SystemColors.ControlText : palette.NotebookInk(toolsFill);
            Check(tools.Visible && tools.ForeColor == toolsInk && tools.Font.Style == FontStyle.Regular && tools.GroupFill == toolsFill &&
                (chrome.HighContrast ? toolsFill == SystemColors.Control : toolsFill.R == toolsFill.G && toolsFill.G == toolsFill.B &&
                    toolsFill.ToArgb() != palette.Surface.ToArgb()) &&
                tools.Left == Math.Max(0, star.Bounds.Left - toolsPad) && tools.Left + tools.Padding.Left == star.Bounds.Left + TabRenderer.CornerRadius(chrome.Dpi) &&
                tools.Right == strip.MenuTabBounds.Right + toolsPad &&
                strip.Items.OfType<ToolButton>().All(t => t.Bounds.Top == star.Bounds.Top && t.Height == star.Height),
                "Tools misalignment at " + chrome.Dpi + " DPI: label " + tools.Bounds + ", group " + strip.MenuTabBounds +
                ", star " + star.Bounds + ", first command " + strip.Items.OfType<ToolButton>().First().Bounds + ".");
            Check(band.Left == strip.Left && (bottom ? band.Bottom == strip.Top : band.Top == strip.Bottom),
                "The notebook captions must sit on the side of the tabs facing the page.");
            using (var captions = new Bitmap(band.Width, band.Height))
            {
                band.DrawToBitmap(captions, band.ClientRectangle);
                int toolsOuter = bottom ? 0 : band.Height - 1, toolsJoin = bottom ? band.Height - 1 : 0;
                Check(captions.GetPixel(tools.Left, toolsOuter).ToArgb() == palette.Surface.ToArgb() &&
                    captions.GetPixel(tools.Right - 1, toolsOuter).ToArgb() == palette.Surface.ToArgb() &&
                    captions.GetPixel(tools.Right - 1, toolsJoin).ToArgb() == toolsFill.ToArgb() &&
                    captions.GetPixel(tools.Left + tools.Width / 2, toolsOuter).ToArgb() == toolsFill.ToArgb() &&
                    bitmap.GetPixel(tools.Right - 1, strip.Height / 2).ToArgb() == toolsFill.ToArgb(),
                    "Tools must have one continuous neutral band, rounded toward the editor.");
                foreach (var group in strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main).GroupBy(t => t.Group))
                {
                    var label = group.Key;
                    var box = label.Bounds;
                    bool pinned = group.All(t => ((SavedTab)t.Tag).Pinned);
                    bool colored = !chrome.HighContrast && (pinned || label.Notebook != null && !label.Notebook.Color.IsEmpty);
                    int pad = colored ? 2 * edge : 0;
                    Check(label.Visible && !label.CanSelect && label.AccessibleName == label.Text && label.Font.Style == FontStyle.Regular &&
                        label.Font.Size < first.Font.Size && box.Left + pad == group.First().Bounds.Left &&
                        (label.Pin == null ? label.Padding.Left == pad + TabRenderer.CornerRadius(chrome.Dpi) :
                            label.Pin.Left == pad + TabRenderer.CornerRadius(chrome.Dpi) - HostChrome.Px(3, chrome.Dpi) &&
                            label.Padding.Left == label.Pin.Right + HostChrome.Px(2, chrome.Dpi)) && box.Right - pad == group.Last().Bounds.Right &&
                        box.Height == label.Font.Height + HostChrome.Px(3, chrome.Dpi) &&
                        (bottom ? group.First().Bounds.Top == 0 : group.First().Bounds.Bottom == strip.Height),
                        "The caption must center in the full space between the tab and band edges, with equal end padding.");
                    Color fill = !colored ? palette.Surface : pinned ? palette.PinnedBand : palette.NotebookCover(label.Notebook.Color);
                    if (pinned && colored)
                        Check(fill.ToArgb() != palette.ToolsBand.ToArgb(), "The Pinned and Tools bands must be distinct.");
                    Color ink = colored ? palette.NotebookInk(fill) : palette.Ink;
                    int outer = bottom ? box.Top : box.Bottom - 1, seam = bottom ? box.Bottom - 1 : box.Top;
                    var inside = new Point(box.Right - label.Padding.Right - 2 * edge, outer + (bottom ? edge : -edge));
                    Check(captions.GetPixel(box.Left, outer).ToArgb() == palette.Surface.ToArgb() &&
                        captions.GetPixel(box.Right - 1, outer).ToArgb() == palette.Surface.ToArgb() &&
                        captions.GetPixel(box.Right - 1, seam).ToArgb() == fill.ToArgb() &&
                        captions.GetPixel(inside.X, inside.Y).ToArgb() == fill.ToArgb() &&
                        bitmap.GetPixel(box.Right - 1, bottom ? 0 : strip.Height - 1).ToArgb() == fill.ToArgb(),
                        String.Format("Notebook backing for {0} at {1} DPI (bottom={2}): corners {3:X8}/{4:X8}, seam {5:X8}, outer {6:X8}, strip {7:X8}; expected surface {8:X8}, fill {9:X8}.",
                            label.Text, chrome.Dpi, bottom, captions.GetPixel(box.Left, outer).ToArgb(), captions.GetPixel(box.Right - 1, outer).ToArgb(),
                            captions.GetPixel(box.Right - 1, seam).ToArgb(), captions.GetPixel(inside.X, inside.Y).ToArgb(),
                            bitmap.GetPixel(box.Right - 1, bottom ? 0 : strip.Height - 1).ToArgb(), palette.Surface.ToArgb(), fill.ToArgb()));
                    if (colored)
                        for (int i = 0; i < pad; i++)
                            Check(bitmap.GetPixel(box.Left + i, strip.Height / 2).ToArgb() == fill.ToArgb() &&
                                bitmap.GetPixel(box.Right - 1 - i, strip.Height / 2).ToArgb() == fill.ToArgb(),
                                "The notebook fill must be visible at both ends of its tab row.");
                    int firstInk = box.Bottom, lastInk = -1;
                    bool solidInk = false;
                    int inkDistance = Math.Abs(ink.R-fill.R) + Math.Abs(ink.G-fill.G) + Math.Abs(ink.B-fill.B);
                    for (int x = box.Left + label.Padding.Left; x < box.Right - label.Padding.Right; x++)
                        for (int y = box.Top; y < box.Bottom; y++)
                        {
                            Color pixel = captions.GetPixel(x, y);
                            if (pixel.ToArgb() == ink.ToArgb()) solidInk = true;
                            int distance = Math.Abs(pixel.R-fill.R) + Math.Abs(pixel.G-fill.G) + Math.Abs(pixel.B-fill.B);
                            if (distance < inkDistance / 3) continue;
                            firstInk = Math.Min(firstInk, y);
                            lastInk = Math.Max(lastInk, y);
                        }
                    Check(solidInk && Math.Abs((firstInk - box.Top) - (box.Bottom - 1 - lastInk)) <= HostChrome.Px(3, chrome.Dpi),
                        "Caption ink falls off center for " + label.Text + " at " + chrome.Dpi + " DPI: " + firstInk + ".." + lastInk + " in " + box.Height + " pixels.");
                    if (label.Pin != null)
                        CheckPinAppearance(label, fill, ink, firstInk - box.Top, lastInk - box.Top, chrome.Dpi);
                }
            }
            Color starBackground = star.Pressed ? palette.Active : star.Selected ? palette.Hover : palette.Tab;
            Color gearBackground = gear.Pressed ? palette.Active : gear.Selected ? palette.Hover : palette.Tab;
            int starCenter = star.Bounds.Left + star.Width / 2;
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + star.Height / 2).ToArgb() == Color.Gold.ToArgb(),
                "The bookmark star must have a solid yellow center.");
            Check(bitmap.GetPixel(starCenter, star.Bounds.Top + 2 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(starCenter, star.Bounds.Bottom - 3 * edge).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Left, FlatEdgeY(star, !bottom, 0)).ToArgb() == toolsFill.ToArgb() &&
                bitmap.GetPixel(strip.MenuTabBounds.Right - 1, FlatEdgeY(star, !bottom, 0)).ToArgb() == toolsFill.ToArgb() &&
                bitmap.GetPixel(star.Bounds.Right - 1, FlatEdgeY(star, !bottom, 2 * edge)).ToArgb() == starBackground.ToArgb() &&
                bitmap.GetPixel(gear.Bounds.Left, FlatEdgeY(star, !bottom, 2 * edge)).ToArgb() == gearBackground.ToArgb(),
                "Only the two outside corners of the shared Tools tab should be rounded.");
            if (!chrome.HighContrast)
            {
                var inactive = strip.Items.OfType<PageTabButton>().First(t => !t.Checked && !t.Selected && t.Placement == ToolStripItemPlacement.Main);
                int middle = star.Bounds.Top + star.Height / 2;
                var border = bitmap.GetPixel(inactive.Bounds.Left, middle).ToArgb();
                Check(bitmap.GetPixel(star.Bounds.Left, middle).ToArgb() == border &&
                    bitmap.GetPixel(strip.MenuTabBounds.Right - 1, middle).ToArgb() == border &&
                    bitmap.GetPixel(gear.Bounds.Left, FlatEdgeY(star, bottom, 0)).ToArgb() == border &&
                    bitmap.GetPixel(gear.Bounds.Left, FlatEdgeY(star, !bottom, 0)).ToArgb() == border,
                    "The shared Tools perimeter must match the inactive page tab outline.");
            }
            // Read through the two side points, ignoring very faint antialiasing pixels.
            int armY = star.Bounds.Top + star.Height / 2 - HostChrome.Px(2, chrome.Dpi);
            int starLeft = -1;
            for (int x = star.Bounds.Left + edge; x < star.Bounds.Right && starLeft < 0; x++)
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
                    var leftTab = left as PageTabButton;
                    var rightTab = right as PageTabButton;
                    int leftGap = x - (leftTab == null ? tools.Right : leftTab.Group.Right);
                    int rightGap = (rightTab == null ? right.Bounds.Left : rightTab.Group.Left) - x - 1;
                    Check(leftGap > 0 && (leftGap == rightGap || (right is PagePicker && rightGap >= leftGap)),
                        "Divider gaps in pixels: left " + leftGap + ", right " + rightGap);
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
                if (!chrome.HighContrast)
                {
                    int midY = tab.Bounds.Top + tab.Height / 2;
                    Color border = bitmap.GetPixel(tab.Bounds.Left, midY);
                    Check(border.GetBrightness() < fill.GetBrightness() &&
                        bitmap.GetPixel(tab.Bounds.Right - 1, midY).ToArgb() == border.ToArgb() &&
                        bitmap.GetPixel(tab.Bounds.Left + tab.Width / 2, FlatEdgeY(tab, !bottom, 0)).ToArgb() == border.ToArgb(),
                        "Each page tab needs a darker outline on both sides and along its rounded edge.");
                }
                var close = tab.CloseBounds;
                int cx = tab.Bounds.Left + close.Left + (close.Width - 1) / 2, cy = tab.Bounds.Top + (close.Height - 1) / 2;
                int arm = HostChrome.Px(2, chrome.Dpi);
                Check(tab.Padding.Right >= close.Width && bitmap.GetPixel(cx-arm, cy-arm).ToArgb() != fill.ToArgb() &&
                    bitmap.GetPixel(cx+arm, cy-arm).ToArgb() != fill.ToArgb() && bitmap.GetPixel(cx, cy).ToArgb() != fill.ToArgb(),
                    "A page tab is missing a visible close glyph or space between the title and the glyph.");
                Check(fill.ToArgb() != marker.ToArgb() && bitmap.GetPixel(tab.Bounds.Left + 2 * edge,
                    FlatEdgeY(tab, bottom, HostChrome.Px(3, chrome.Dpi))).ToArgb() == fill.ToArgb(), "Incorrect section stripe thickness.");
                Check(bitmap.GetPixel(tab.Bounds.Left + 2 * edge, FlatEdgeY(tab, bottom, HostChrome.Px(3, chrome.Dpi) + edge)).ToArgb() == fill.ToArgb() &&
                    bitmap.GetPixel(tab.Bounds.Left + tab.Width / 2, FlatEdgeY(tab, !bottom, 2 * edge)).ToArgb() == fill.ToArgb(),
                    "Tab body fill does not match its selection state.");
            }
            Color behind = active.Group.GroupFill.IsEmpty ? palette.Surface : active.Group.GroupFill;
            Check(bitmap.GetPixel(b.Left + edge - 1, FlatEdgeY(active, !bottom, 0)).ToArgb() == behind.ToArgb() &&
                bitmap.GetPixel(b.Right - edge, FlatEdgeY(active, !bottom, 0)).ToArgb() == behind.ToArgb(),
                "The active tab's corners facing the editor are not rounded.");
            return bitmap;
        }
        catch { bitmap.Dispose(); throw; }
    }

    private static void CheckCaptionAlignment(PageStrip strip)
    {
        var labels = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Where(l => l.Visible).ToArray();
        var tools = labels.Single(l => l.Text == "Tools");
        var pinned = labels.Single(l => l.Pinned);
        var notebook = labels.First(l => l.Notebook != null);
        var bounds = labels.Select(l => l.Bounds).ToArray();
        string active = strip.ActiveTabKey;
        var expected = new[] { ContentAlignment.MiddleLeft, ContentAlignment.MiddleCenter, ContentAlignment.MiddleRight, ContentAlignment.MiddleLeft };
        var sources = new[] { tools, notebook, pinned, tools };
        for (int i = 0; i < expected.Length; i++)
        {
            PrototypeTests.ClickCaption(sources[i], right: i == 0);
            Check(labels.All(l => l.TextAlign == expected[i]) && labels.Select(l => l.Bounds).SequenceEqual(bounds) && active == strip.ActiveTabKey,
                "Clicking " + sources[i].Text + " must give every label " + expected[i] + " alignment without moving or selecting tabs.");
            using (var image = new Bitmap(tools.Width, tools.Height))
            {
                tools.DrawToBitmap(image, tools.ClientRectangle);
                int left = tools.Width, right = -1;
                for (int x = tools.Padding.Left; x < tools.Width - tools.Padding.Right; x++)
                    for (int y = 0; y < tools.Height; y++)
                        if (image.GetPixel(x, y).ToArgb() == tools.ForeColor.ToArgb())
                        {
                            left = Math.Min(left, x);
                            right = Math.Max(right, x);
                        }
                int leftSpace = left - tools.Padding.Left, rightSpace = tools.Width - tools.Padding.Right - right - 1;
                int tolerance = HostChrome.Px(3, ((TabRenderer)strip.Renderer).Dpi);
                Check(right >= left && (expected[i] == ContentAlignment.MiddleCenter ? Math.Abs(leftSpace - rightSpace) <= tolerance :
                    expected[i] == ContentAlignment.MiddleRight ? rightSpace <= tolerance : leftSpace <= tolerance),
                    "Painted Tools text missed its " + expected[i] + " alignment.");
            }
        }
        Console.WriteLine("PASS all group labels cycle together on left click and stay put on right click.");
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
        Check(active.Checked, "Unpinning switched the active tab.");
        CheckGroupDividers(strip);
        Console.WriteLine("PASS bookmark star, group boundaries and pin/unpin via the tab menu with the active tab preserved.");
        Console.WriteLine("PASS tab rendering in OneNote: surface #{0:X6}, bookmark star X={1}px ({2}), DPI {3}; square top and rounded lower corners.",
            chrome.Surface.ToArgb() & 0xffffff, chrome.Inset, !layout.VerticalTabs && chrome.Measured ? "measured notebook icon" : "default inset", chrome.Dpi);
    }
    internal static void TogglePin(PageStrip strip, PageTabButton button, bool pin)
    {
        var menu = strip.TabMenuFor((SavedTab)button.Tag);
        menu.Items.Cast<ToolStripItem>().Single(item => item.Text == (pin ? "Pin tab" : "Unpin tab")).PerformClick();
        Check(((SavedTab)button.Tag).Pinned == pin, "Tab menu did not change the pin.");
    }
}
