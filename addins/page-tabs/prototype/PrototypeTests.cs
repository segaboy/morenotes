using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using One = Microsoft.Office.Interop.OneNote;

internal static class PrototypeTests
{
    private static void Check(bool condition, string failure)
    {
        if (!condition) throw new InvalidOperationException(failure);
    }
    private sealed class TestPages : IPages
    {
        internal int NotesCreated;
        internal bool LocationUnavailable, FailNextNavigation, DelayNextNavigation;
        internal string PendingPage;
        public string Current { get; private set; }
        internal Color BookBColor = Color.FromArgb(246,176,120);
        internal BookmarkNotebook BookB = new BookmarkNotebook("b", "Workshop", Color.Goldenrod);
        public Color SectionColor(string id) { return id.StartsWith("a/") ? Color.FromArgb(145,186,174) : BookBColor; }
        public Dictionary<string, BookmarkLocation> PageLocations(IEnumerable<string> ids, bool refresh = false)
        {
            var a = new BookmarkNotebook("a", "Atlas", Color.SteelBlue);
            var b = BookB;
            var personal = new BookmarkNotebook("personal", "Personal", Color.Coral);
            return ids.Distinct().ToDictionary(id => id, id => id.StartsWith("personal/") ? new BookmarkLocation(personal, "Quick Notes", SectionColor(id)) :
                id.StartsWith("a/") ? new BookmarkLocation(a, "Route plans", SectionColor(id)) :
                id == "b/parts" ? new BookmarkLocation(b, "Materials & tools", Color.FromArgb(183, 201, 151)) :
                new BookmarkLocation(b, "Sketches", id == "b/design" ? Color.FromArgb(171, 156, 183) : SectionColor(id)));
        }
        internal readonly Dictionary<string, string> Renamed = new Dictionary<string, string>();
        public string Caption(string id)
        {
            string caption;
            return Renamed.TryGetValue(id, out caption) ? caption : id == "b/1" ? "Dashboard / B" : id == "a/1" ? "Dashboard / A" :
                id.StartsWith("personal/quick/") ? "Untitled / Personal" : "Schedule / A";
        }
        public string CreateQuickNote()
        {
            if (LocationUnavailable) throw new InvalidOperationException("No configured Quick Notes section.");
            return "personal/quick/" + ++NotesCreated;
        }
        public void Navigate(string id)
        {
            if (FailNextNavigation) { FailNextNavigation = false; throw new System.Runtime.InteropServices.COMException(); }
            if (DelayNextNavigation) { DelayNextNavigation = false; PendingPage = id; return; }
            Current = id;
        }
        public string LinkFor(string id) { return ""; }
        public void NavigateLink(string link) { throw new NotSupportedException(); }
    }
    internal static void ControlledWindow()
    {
        using (var frame = new Form { Text = "MoreTabs layout test", ShowInTaskbar = false, StartPosition = FormStartPosition.Manual,
            Location = new Point(-16000, -16000), ClientSize = new Size(730, 420) })
        using (var ribbon = new Panel())
        using (var document = new Panel())
        using (var navigation = new Panel { Dock = DockStyle.Top, Height = 36 })
        {
            frame.Controls.Add(ribbon);
            frame.Controls.Add(document);
            document.Controls.Add(navigation);
            frame.Show();
            try
            {
                var source = new TestPages();
                Exception failure = null;
                using (var layout = new NativeLayout(frame.Handle, document.Handle, ribbon.Handle, navigation.Handle))
                {
                    Check(NativeLayout.RootOf(navigation.Handle) == frame.Handle && NativeLayout.RootOf(frame.Handle) == frame.Handle,
                        "A descendant handle and its top-level frame must resolve to the same editor.");
                    Check(NativeLayout.RootOf(IntPtr.Zero) == IntPtr.Zero, "A missing handle resolved to a frame.");
                    using (var other = new Form())
                    {
                        var both = new List<IntPtr> { frame.Handle, other.Handle };
                        Check(NotebookPages.SelectFrame(both, NativeLayout.RootOf(document.Handle), IntPtr.Zero) == frame.Handle,
                            "Current child handle did not select its editor.");
                        Check(NotebookPages.SelectFrame(both, frame.Handle, other.Handle) == other.Handle,
                            "Explicit frame did not override the current editor.");
                        both.Remove(frame.Handle);
                        Check(NotebookPages.SelectFrame(both, frame.Handle, IntPtr.Zero) == other.Handle,
                            "Unmatched active window hid the sole eligible editor.");
                        CheckSelectionFails(both, other.Handle, frame.Handle);
                        both.Add(frame.Handle);
                        CheckSelectionFails(both, IntPtr.Zero, IntPtr.Zero);
                        both.Add(frame.Handle);
                        CheckSelectionFails(both, frame.Handle, frame.Handle);
                    }
                    Console.WriteLine("PASS window selection: child handles, alternate editor, explicit choice and ambiguous matches.");
                    Rectangle expected = Rectangle.Empty;
                    using (var strip = new PageStrip(frame.Handle, source, delegate {}, delegate(Exception e) { failure = e; }))
                    {
                        Rectangle slot = Rectangle.Empty;
                        foreach (int height in new int[] { 120, 60, 120 })
                        {
                            frame.ClientSize = new Size(frame.ClientSize.Width + 40, frame.ClientSize.Height + 10);
                            ribbon.Bounds = new Rectangle(0, 0, frame.ClientSize.Width, height);
                            expected = new Rectangle(0, height, frame.ClientSize.Width - 110, frame.ClientSize.Height - height);
                            document.Bounds = expected;
                            slot = layout.Reserve();
                            strip.Attach(slot, layout);
                            layout.CheckHitTargets(strip.Host.Handle, slot);
                            Check(slot == layout.Reserve(), "A repeated tick consumed another row.");
                            Check(slot.Top == height && slot.Bottom == layout.NavigationBounds.Top && slot.Width == expected.Width &&
                                layout.DocumentBounds.Bottom == expected.Bottom, "The test strip overlaps native controls.");
                        }
                        using (var cover = new Panel { Bounds = new Rectangle(0, 0, document.Width / 3, navigation.Height) })
                        {
                            document.Controls.Add(cover);
                            cover.BringToFront();
                            bool detected = false;
                            try { layout.CheckHitTargets(strip.Host.Handle, slot); }
                            catch (InvalidOperationException error) { detected = error.Message.Contains("navigation"); }
                            Check(detected, "The hit test accepted a control covering the left of the navigation row.");
                        }
                        layout.CheckHitTargets(strip.Host.Handle, slot);
                        Console.WriteLine("PASS detected a partially covered navigation row and its recovery.");
                        string[] ids = { "a/1", "a/2", "b/1" };
                        foreach (var id in new[] { "a/1", "b/1", "a/2" }) { source.Navigate(id); strip.ClickKeepCurrent(); }
                        strip.ClickKeepCurrent();
                        strip.ClickTab("b/1");
                        Check(strip.TabCount == 3, "Duplicate add created another tab.");
                        AppearanceTests.Controlled(strip, layout, slot);
                        var previousBook = source.BookB;
                        source.BookB = new BookmarkNotebook("b", "Atlas", Color.Orchid);
                        foreach (var tab in strip.Items.OfType<PageTabButton>()) tab.NextColorCheck = DateTime.MinValue;
                        strip.RefreshActive();
                        var renamed = strip.Items.OfType<PageTabButton>().Select(t => t.Group).Distinct().ToArray();
                        Check(renamed.Length == 2 && renamed.All(label => label.Text == "Atlas") &&
                            renamed[0].Notebook.Id != renamed[1].Notebook.Id && renamed[1].Notebook.Color == Color.Orchid,
                            "Renaming or recoloring a notebook merged distinct notebooks with the same name.");
                        var beforeRename = strip.Items.OfType<PageTabButton>().Select(t => t.Bounds).ToArray();
                        source.BookB = new BookmarkNotebook("b", "Workshop projects and reference material", Color.Goldenrod);
                        foreach (var tab in strip.Items.OfType<PageTabButton>()) tab.NextColorCheck = DateTime.MinValue;
                        strip.RefreshActive();
                        var singlePageCaption = renamed[1];
                        Check(TextRenderer.MeasureText(singlePageCaption.Text, singlePageCaption.Font).Width > singlePageCaption.Width &&
                            singlePageCaption.AccessibleName == source.BookB.Name && singlePageCaption.Visible &&
                            beforeRename.SequenceEqual(strip.Items.OfType<PageTabButton>().Select(t => t.Bounds)),
                            "A long notebook caption moved the tabs or lost its full accessible name.");
                        source.BookB = previousBook;
                        foreach (var tab in strip.Items.OfType<PageTabButton>()) tab.NextColorCheck = DateTime.MinValue;
                        strip.RefreshActive();
                        expected = CheckDocking(frame, document, layout, strip);
                        Check(source.NotesCreated == 0, "Keeping an existing page created a Quick Note.");
                        strip.ClickAdd();
                        Check(strip.HasTabFocus(source.Current), "The first Quick Note tab did not receive keyboard focus.");
                        var note = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                        Check(note.Text == "Untitled", "New Quick Note did not adopt OneNote's initial title.");
                        source.Renamed[source.Current] = "Trip plans / Personal";
                        WaitForLabel(strip, note, "Trip plans");
                        strip.ClickAdd();
                        Check(source.NotesCreated == 2 && strip.TabCount == 5 && source.Current == "personal/quick/2" &&
                            strip.IsSelected(source.Current) && strip.HasTabFocus(source.Current), "Each click should create, open and focus a distinct Quick Note tab.");
                        source.LocationUnavailable = true;
                        strip.ClickAdd();
                        Check(source.NotesCreated == 2 && strip.TabCount == 5 && source.Current == "personal/quick/2",
                            "Missing Quick Notes destination changed tabs or created a page elsewhere.");
                        source.LocationUnavailable = false;
                        source.FailNextNavigation = true;
                        strip.ClickAdd();
                        strip.ClickTab("personal/quick/3");
                        Check(source.NotesCreated == 3 && strip.TabCount == 6 && strip.IsSelected("personal/quick/3"),
                            "Retrying a created Quick Note lost the page or created a duplicate.");
                        source.DelayNextNavigation = true;
                        strip.ClickAdd();
                        Check(strip.IsSelected("personal/quick/3") && !strip.HasTabFocus(source.PendingPage),
                            "Focus moved to the new tab before its page opened.");
                        source.Navigate(source.PendingPage); strip.RefreshActive();
                        Check(strip.IsSelected(source.Current) && strip.HasTabFocus(source.Current), "Delayed navigation lost new-tab focus.");
                        Console.WriteLine("PASS + creates distinct Quick Notes from another notebook and focuses each new tab; missing destination and navigation retry preserve the working set.");
                        Check(NotebookPages.ReadSectionColor(XElement.Parse("<Section color='#138bc4'/>")).ToArgb() == Color.FromArgb(19,139,196).ToArgb() &&
                            NotebookPages.ReadSectionColor(XElement.Parse("<Section color='automatic'/>")).IsEmpty &&
                            NotebookPages.ReadSectionColor(XElement.Parse("<Section/>")).IsEmpty, "Section color parsing failed.");
                        foreach (var id in ids)
                        {
                            strip.ClickTab(id);
                            Check(source.Current == id && strip.IsSelected(id), "Tab click did not navigate.");
                            Check(strip.Items.OfType<PageTabButton>().Single(t => t.Checked).SectionColor == source.SectionColor(id),
                                "The active tab kept the previous section color.");
                        }
                        var selected = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                        var inactive = strip.Items.OfType<PageTabButton>().First(t => !t.Checked && ((SavedTab)t.Tag).Page.Id.StartsWith("personal/"));
                        source.BookBColor = Color.Empty;
                        var colorWait = System.Diagnostics.Stopwatch.StartNew();
                        while ((!selected.SectionColor.IsEmpty || !inactive.SectionColor.IsEmpty) && colorWait.ElapsedMilliseconds < 5000)
                        {
                            Thread.Sleep(50);
                            strip.RefreshActive();
                        }
                        var renderer = (TabRenderer)strip.Renderer;
                        Check(selected.Checked && selected.SectionColor.IsEmpty && inactive.SectionColor.IsEmpty &&
                            renderer.MarkerFor(selected) == (SystemInformation.HighContrast ? SystemColors.HighlightText : renderer.ActiveMarker),
                            "Changing the current section to automatic color did not restore the theme accent.");
                        Console.WriteLine("PASS active and inactive tabs follow section recoloring across notebooks without navigation.");
                        using (var editor = new TextBox { Location = new Point(6, navigation.Bottom + 6) })
                        {
                            document.Controls.Add(editor);
                            editor.Focus();
                            strip.ClickTab(ids[0]);
                            editor.Focus();
                            strip.FocusPageTabs(editor.Handle);
                            Check(strip.Focused, "Keyboard entry did not focus the strip.");
                            PressStripKey(strip, Keys.Right);
                            PressStripKey(strip, Keys.Enter);
                            Check(source.Current == ids[1] && editor.Focused, "Enter did not select the next page and restore editor focus.");
                            strip.FocusPageTabs(editor.Handle);
                            PressStripKey(strip, Keys.Left);
                            PressStripKey(strip, Keys.Escape);
                            Check(source.Current == ids[1] && editor.Focused, "Escape did not return without navigating.");
                            Console.WriteLine("PASS controlled keyboard entry, arrow selection, Enter and Escape return to editor.");
                        }
                        var moving = strip.Items.OfType<PageTabButton>().Single(b => b.Checked);
                        var original = strip.Items.OfType<PageTabButton>().First();
                        string movingKey = ((SavedTab)moving.Tag).Key;
                        source.Navigate("a/1"); strip.RefreshActive();
                        Check(moving.Checked && !original.Checked && moving != original && moving.Text == original.Text &&
                            strip.IsSelected("a/1"), "Native navigation to an already-open page switched tab slots.");
                        original.PerformClick();
                        Check(original.Checked && !moving.Checked, "Clicking a second tab for the same page failed to select that tab.");
                        moving.PerformClick();
                        source.Navigate("other"); strip.RefreshActive();
                        Check(moving.Checked && ((SavedTab)moving.Tag).Key == movingKey && ((SavedTab)moving.Tag).Page.Id == "other" &&
                            moving.Group.Notebook.Id == "b" && original.Group.Notebook.Id == "a" &&
                            ((SavedTab)original.Tag).Page.Id == "a/1", "Native navigation did not move the same tab to its destination notebook.");
                        Console.WriteLine("PASS native navigation regroups the active tab by notebook and preserves its identity and independent selection.");
                        var section = XElement.Parse("<Section xmlns='urn:test'><Page ID='a/1' name='Dashboard' isCurrentlyViewed='true'/>" +
                            "<Page ID='b/1' name='Dashboard' selected='all'/></Section>");
                        string target = PageMenuCommand.SelectedPage(section);
                        Check(target == "b/1", "The context menu chose the displayed page over the selected page.");
                        section.Elements().First().SetAttributeValue("selected", "all");
                        Check(PageMenuCommand.SelectedPage(section) == null && PageMenuCommand.SelectedPage(new XElement("Section")) == null,
                            "An ambiguous or empty selection was accepted.");
                        int count = strip.TabCount, notes = source.NotesCreated;
                        AppearanceTests.TogglePin(strip, moving, true);
                        Check(PageMenuCommand.Send(frame.Handle, target), "The native menu could not reach its editor's strip.");
                        Application.DoEvents();
                        string firstNew = strip.ActiveTabKey;
                        Check(strip.TabCount == count + 1 && firstNew != movingKey && strip.IsSelected(target) &&
                            ((SavedTab)moving.Tag).Pinned && ((SavedTab)moving.Tag).Page.Id == "other", "Opening from the native page list replaced the source tab.");
                        Check(PageMenuCommand.Send(frame.Handle, target), "Could not reopen the selected page in an independent tab.");
                        Application.DoEvents();
                        Check(strip.TabCount == count + 2 && strip.ActiveTabKey != firstNew && strip.HasTabFocus(target) &&
                            !((SavedTab)strip.Items.OfType<PageTabButton>().Single(t => t.Checked).Tag).Pinned && source.NotesCreated == notes,
                            "Opening an existing page should create a focused, unpinned tab each time.");
                        source.FailNextNavigation = true;
                        strip.OpenPageInNewTab(target);
                        Check(strip.TabCount == count + 2, "Failed navigation created an empty tab.");
                        Console.WriteLine("PASS page-list command resolves a selected ID, preserves the pinned source, and opens distinct focused tabs.");
                        CheckCloseIcons(strip, source);
                        if (failure != null) throw failure;
                    }
                    layout.Dispose();
                    Check(layout.DocumentBounds == expected, "Normal close did not restore the document.");
                }
                Console.WriteLine("PASS controlled window: child attachment, 3 layout changes, no cumulative movement, restore, three tabs, duplicate titles, duplicate add and external page changes.");
            }
            finally { frame.Close(); }
        }
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

    private static void ClickCloseTarget(PageTabButton tab, bool releaseOnIcon)
    {
        var parent = tab.GetCurrentParent();
        int x = tab.Bounds.Left + tab.CloseBounds.Left + tab.CloseBounds.Width / 2;
        int y = tab.Bounds.Top + tab.Height / 2;
        Check(parent.GetItemAt(x, y) == tab, "Close target does not hit the intended tab.");
        SendMessage(parent.Handle, 0x0201, (IntPtr)1, (IntPtr)((y << 16) | x));
        if (!releaseOnIcon) x = tab.Bounds.Left + 3;
        SendMessage(parent.Handle, 0x0202, IntPtr.Zero, (IntPtr)((y << 16) | x));
    }

    private static void CheckCloseIcons(PageStrip strip, TestPages source)
    {
        int count = strip.TabCount;
        string selected = strip.ActiveTabKey, page = source.Current;
        var duplicate = strip.Items.OfType<PageTabButton>().First(t => !t.Checked && ((SavedTab)t.Tag).Page.Id == page);
        AppearanceTests.TogglePin(strip, duplicate, true);
        ClickCloseTarget(duplicate, false);
        Check(strip.TabCount == count && strip.ActiveTabKey == selected, "Dragging off the close icon activated or closed its tab.");
        ClickCloseTarget(duplicate, true);
        Check(!strip.Items.Contains(duplicate) && strip.TabCount == count - 1 && strip.ActiveTabKey == selected && source.Current == page,
            "Closing a pinned duplicate must preserve the selected tab and editor page.");

        strip.Host.Width = 245;
        strip.OverflowButton.ShowDropDown();
        var hidden = strip.Items.OfType<PageTabButton>().First(t => t.Placement == ToolStripItemPlacement.Overflow && !t.Checked);
        ClickCloseTarget(hidden, true);
        Check(!strip.Items.Contains(hidden) && strip.TabCount == count - 2 && strip.ActiveTabKey == selected && source.Current == page,
            "Closing an overflow tab switched pages or failed to remove it.");
        strip.Host.Width = 850;
        var active = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        ClickCloseTarget(active, true);
        Check(!strip.Items.Contains(active) && strip.TabCount == count - 3 && source.Current == page, "Closing the active tab changed the editor page.");
        while (strip.TabCount > 0)
        {
            var remaining = strip.Items.OfType<PageTabButton>().First();
            strip.FocusPageTabs(IntPtr.Zero);
            remaining.Select();
            int before = strip.TabCount;
            PressStripKey(strip, Keys.Delete);
            Check(!strip.Items.Contains(remaining) && strip.TabCount == before - 1, "Delete did not close exactly the focused tab.");
        }
        Check(source.Current == page && strip.TabCount == 0, "Closing the last tab must leave the editor page open.");
        Console.WriteLine("PASS close icons on active, inactive, pinned, duplicate and overflow tabs, cancel by dragging off, and keyboard close.");
    }

    private static Rectangle CheckDocking(Form frame, Panel document, NativeLayout layout, PageStrip strip)
    {
        string selected = strip.ActiveTabKey;
        var arrow = strip.Host.MoveButton;
        arrow.PerformClick();
        Check(!strip.AtBottom && layout.StillReserved(strip.Host.Bounds), "A partial-width document allowed a bottom overlay.");
        var full = Rectangle.FromLTRB(0, layout.RibbonBounds.Bottom, frame.ClientSize.Width, frame.ClientSize.Height);
        document.Bounds = full;
        strip.Attach(layout.Reserve(), layout);
        using (var footer = new Panel { Dock = DockStyle.Bottom, Height = 16 })
        using (var sibling = new Panel { Bounds = new Rectangle(50, full.Bottom - 16, 70, 16) })
        {
            document.Controls.Add(footer);
            frame.Controls.Add(sibling);
            sibling.BringToFront();
            arrow.PerformClick();
            Check(!strip.AtBottom && layout.StillReserved(strip.Host.Bounds), "Moving would have covered a frame control.");
            sibling.Hide();
            arrow.PerformClick();
            var bottom = strip.Host.Bounds;
            Check(strip.AtBottom && arrow.PointsUp && strip.DefaultDropDownDirection == ToolStripDropDownDirection.AboveRight &&
                layout.StillReserved(bottom) && bottom == layout.Reserve() && bottom.Bottom == frame.ClientSize.Height &&
                bottom.Width == frame.ClientSize.Width && NativeLayout.Bounds(footer.Handle, frame.Handle).Bottom == bottom.Top,
                "Bottom docking did not reserve a stable full-width slot below the native footer.");
            layout.CheckHitTargets(strip.Host.Handle, bottom);
            AppearanceTests.Controlled(strip, layout, bottom);
            strip.Host.Width = 245;
            Check(strip.OverflowButton.Visible && arrow.Right == strip.Host.ClientSize.Width && strip.Right == arrow.Left,
                "Tab overflow displaced the far-right move button.");
            var visible = strip.Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main).ToArray();
            var headers = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Where(l => l.Visible).ToArray();
            Check(visible.Any(t => t.Checked) && visible.All(t => headers.Contains(t.Group) && t.Group.Left <= t.Bounds.Left + TabRenderer.CornerRadius(t.Dpi) && t.Group.Right >= t.Bounds.Right) &&
                headers.All(h => visible.Any(t => t.Group == h)), "Overflow separated the active page from its notebook label.");
            frame.ClientSize = new Size(frame.ClientSize.Width + 32, frame.ClientSize.Height + 48);
            full = Rectangle.FromLTRB(0, layout.RibbonBounds.Bottom, frame.ClientSize.Width, frame.ClientSize.Height);
            document.Bounds = full;
            // Bring the ribbon to the same width a real host assigns on resize.
            NativeLayout.Move(layout.Ribbon, new Rectangle(0, 0, full.Width, full.Top));
            bottom = layout.Reserve();
            strip.Attach(bottom, layout);
            Check(layout.StillReserved(bottom) && bottom.Bottom == full.Bottom && bottom.Width == full.Width,
                "Resize lost the bottom dock.");
            strip.PauseForDialog();
            layout.ReleaseForDialog();
            Check(layout.DocumentBounds == full && !strip.Host.Visible, "Pausing left a bottom gap.");
            strip.Attach(layout.Reserve(), layout);
            Check(strip.AtBottom && layout.StillReserved(bottom), "Resuming did not restore bottom placement.");
            arrow.PerformClick();
            Check(!strip.AtBottom && !arrow.PointsUp && layout.StillReserved(strip.Host.Bounds) &&
                strip.Host.Top == full.Top && layout.DocumentBounds.Bottom == full.Bottom && strip.ActiveTabKey == selected,
                "The up arrow did not restore top placement and the selected tab.");
            arrow.PerformClick();
            strip.PauseForDialog();
            layout.Dispose();
            Check(layout.DocumentBounds == full, "Closing MoreTabs at the bottom left a gap.");
            strip.Attach(layout.Reserve(false, false), layout);
        }
        Console.WriteLine("PASS down/up arrows, full-width bottom docking, native footer, overflow, resize, modal resume and original selection.");
        return full;
    }

    private static void CheckSelectionFails(List<IntPtr> matches, IntPtr current, IntPtr requested)
    {
        try { NotebookPages.SelectFrame(matches, current, requested); }
        catch (InvalidOperationException) { return; }
        throw new InvalidOperationException("An unmatched or ambiguous editor was accepted.");
    }

    private static void PressStripKey(PageStrip strip, Keys key)
    {
        const int wmKeyDown = 0x0100;
        var message = Message.Create(strip.Handle, wmKeyDown, (IntPtr)key, IntPtr.Zero);
        Check(strip.PreProcessMessage(ref message), "The strip did not handle " + key);
    }

    private static void WaitForPage(NotebookPages source, string id, string action = "tab navigation")
    {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        while (source.Current != id)
        {
            if (timer.ElapsedMilliseconds > 10000) throw new TimeoutException("OneNote did not display the requested page during " + action + ".");
            Application.DoEvents(); Thread.Sleep(40);
        }
    }

    private static string CreatePage(One.Application app, string section, string name)
    {
        string id;
        app.CreateNewPage(section, out id, One.NewPageStyle.npsBlankPageWithTitle);
        RenameFixturePage(app, id, name);
        return id;
    }
    private static void RenameFixturePage(One.Application app, string id, string name)
    {
        string xml;
        app.GetPageContent(id, out xml, One.PageInfo.piBasic, One.XMLSchema.xs2013);
        var page = XElement.Parse(xml);
        XNamespace ns = page.Name.Namespace;
        var title = page.Element(ns + "Title");
        var value = title == null ? null : title.Descendants(ns + "T").FirstOrDefault();
        if (value == null) throw new InvalidOperationException("Missing title in the disposable page.");
        value.ReplaceNodes(new XCData(name));
        app.UpdatePageContent(page.ToString(), DateTime.MinValue, One.XMLSchema.xs2013, false);
    }
    private static void WaitForLabel(PageStrip strip, PageTabButton tab, string expected)
    {
        var clock = System.Diagnostics.Stopwatch.StartNew();
        do
        {
            strip.RefreshActive();
            if (tab.Checked && tab.Text == expected) return;
            Application.DoEvents(); Thread.Sleep(60);
        } while (clock.ElapsedMilliseconds < 10000);
        throw new TimeoutException("The active tab did not adopt the disposable page title.");
    }

    internal static void DisposableNotebooks(NotebookPages source, PageStrip strip, NativeLayout layout, Rectangle slot, bool handsOn, string directory)
    {
        string previous = source.Current;
        if (String.IsNullOrEmpty(previous)) throw new InvalidOperationException("Select a page that can be restored after the test.");
        Directory.CreateDirectory(directory);
        var books = new List<string>();
        string quickSection = null;
        var pages = new List<string>();
        try
        {
            foreach (string bookName in new[] { "MoreTabs A", "MoreTabs B" })
            {
                string book, section;
                source.App.OpenHierarchy(Path.Combine(directory, bookName), "", out book, One.CreateFileType.cftNotebook);
                books.Add(book);
                source.App.OpenHierarchy("Test.one", book, out section, One.CreateFileType.cftSection);
                if (books.Count == 2) quickSection = section;
                pages.Add(CreatePage(source.App, section, "Dashboard"));
                if (books.Count == 1) pages.Add(CreatePage(source.App, section, "Schedule"));
            }
            for (int i = 0; i < pages.Count; i++)
            {
                source.Navigate(pages[i]); WaitForPage(source, pages[i], "opening fixture page " + (i + 1));
                Check(source.Window.CurrentNotebookId == books[i == 2 ? 1 : 0], "Notebook mismatch.");
                Check(layout.StillReserved(slot), "OneNote reclaimed the strip when navigating to a test page.");
                strip.ClickKeepCurrent();
            }
            Console.WriteLine("Kept three fixture pages across two test notebooks.");
            Check(strip.TabCount == 3, "Expected three tracked pages.");
            foreach (int i in new[] { 0, 2, 1, 2, 0 })
            {
                strip.ClickTab(pages[i]); WaitForPage(source, pages[i]); strip.RefreshActive();
                Check(strip.IsSelected(pages[i]) && source.Frame == layout.Frame, "Wrong page or OneNote window.");
                Check(strip.Items.OfType<PageTabButton>().Single(t => t.Checked).SectionColor == source.SectionColor(pages[i]),
                    "The tab has a different color from its disposable section.");
                Check(layout.StillReserved(slot), "OneNote reclaimed the strip after a tab click.");
                layout.CheckHitTargets(strip.Host.Handle, slot);
            }
            if (handsOn) ObserveHandsOn(source, strip, layout, pages);
            else
            {
                AppearanceTests.Live(strip, layout, directory);
                var existing = new HashSet<string>(pages);
                source.Navigate(pages[0]); WaitForPage(source, pages[0], "selecting the starting notebook for Quick Notes");
                for (int i = 0; i < 2; i++)
                {
                    strip.ClickAdd();
                    Check(strip.TabCount == 4 + i, "Quick Note click " + (i + 1) + " did not add a new tab.");
                    string id = strip.Items.OfType<PageTabButton>().Select(t => ((SavedTab)t.Tag).Page.Id).Single(page => !existing.Contains(page));
                    WaitForPage(source, id, "opening Quick Note " + (i + 1)); strip.RefreshActive();
                    string parent;
                    source.App.GetHierarchyParent(id, out parent);
                    Check(existing.Add(id) && parent == quickSection && source.Window.CurrentNotebookId == books[1] &&
                        source.Frame == layout.Frame && strip.IsSelected(id) && strip.HasTabFocus(id) && strip.TabCount == 4 + i,
                        "New Quick Note did not open with its own focused tab in the disposable destination section.");
                }
                Console.WriteLine("PASS two + clicks created two Quick Notes in the disposable destination section and focused each new tab in this editor.");
                var note = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                var samePage = strip.Items.OfType<PageTabButton>().First();
                string tabKey = ((SavedTab)note.Tag).Key;
                string noteId = ((SavedTab)note.Tag).Page.Id;
                RenameFixturePage(source.App, noteId, "Quick Note title check");
                WaitForLabel(strip, note, "Quick Note title check");
                source.Navigate(pages[0]); WaitForPage(source, pages[0]); strip.RefreshActive();
                Check(note.Checked && !samePage.Checked && note.Text == samePage.Text && ((SavedTab)note.Tag).Key == tabKey && note.Group == samePage.Group &&
                    ((SavedTab)note.Tag).Page.Id == pages[0] && strip.TabCount == 5, "Navigation switched away from the active tab.");
                source.Navigate(noteId); WaitForPage(source, noteId);
                WaitForLabel(strip, note, "Quick Note title check");
                Console.WriteLine("PASS Quick Note rename, same-tab navigation to an already-open page and return; one active tab throughout.");
                string configured = source.ResolveQuickNotesSection(null), xml;
                source.App.GetHierarchy(configured, One.HierarchyScope.hsSelf, out xml, One.XMLSchema.xs2013);
                Check(XElement.Parse(xml).DescendantsAndSelf().Any(e => e.Name.LocalName == "Section" && (string)e.Attribute("ID") == configured),
                    "The configured Quick Notes location did not resolve to a section.");
                Console.WriteLine("PASS resolved the user's configured Quick Notes section without creating a personal page.");
                strip.Host.Hide();
                try { VerifySavedTabs(source, pages, layout, slot, Path.Combine(directory, "working-set.xml")); }
                finally { strip.Attach(slot, layout); }
            }
            source.Navigate(previous); WaitForPage(source, previous); strip.RefreshActive();
            foreach (var id in pages) Check(!strip.IsSelected(id), "A stale test tab remained active.");
            Console.WriteLine("PASS OneNote COM and tab clicks: three pages, two notebooks, same editor window, external return to an untracked page.");
        }
        finally
        {
            try { source.Navigate(previous); WaitForPage(source, previous, "restoring the preceding page"); }
            finally { foreach (var book in books) source.App.CloseNotebook(book, false); }
        }
    }

    private static void VerifySavedTabs(NotebookPages source, List<string> pages, NativeLayout layout, Rectangle slot, string path)
    {
        var store = new SessionStore(path);
        string targetBook;
        using (var first = new PageStrip(layout.Frame, source, delegate { }, e => { throw e; }, store))
        {
            first.Attach(slot, layout);
            foreach (var id in pages)
            {
                source.Navigate(id); WaitForPage(source, id); first.ClickKeepCurrent();
            }
            var saved = store.Read().Tabs[2];
            targetBook = source.Window.CurrentNotebookId;
            first.SaveBookmark(saved.Page);
            first.SaveBookmark(store.Read().Tabs[0].Page);
            first.CloseTab(saved);
        }
        source.Navigate(pages[0]); WaitForPage(source, pages[0]);
        using (var restored = new PageStrip(layout.Frame, source, delegate { }, e => { throw e; }, new SessionStore(path)))
        {
            restored.Attach(slot, layout);
            restored.RestoreOrSeed();
            Check(restored.TabCount == 2 && restored.IsSelected(pages[0]), "Reopening MoreTabs did not restore the two remaining tabs.");
            var star = restored.Items.OfType<BookmarkButton>().Single();
            star.ShowDropDown();
            var headings = star.DropDownItems.OfType<BookmarkHeader>().ToArray();
            Check(headings.Length == 2 && StringComparer.CurrentCultureIgnoreCase.Compare(headings[0].Text, headings[1].Text) <= 0,
                "The two disposable notebooks were not grouped alphabetically.");
            var target = headings.Single(h => h.Notebook.Id == targetBook);
            var bookmark = (BookmarkRow)star.DropDownItems[star.DropDownItems.IndexOf(target) + 1];
            Check(bookmark.SectionName == "Test" && bookmark.SectionColor.ToArgb() == source.SectionColor(pages[2]).ToArgb(),
                "The bookmark row does not describe the disposable page's host section.");
            bookmark.PerformClick();
            star.HideDropDown();
            WaitForPage(source, pages[2]); restored.RefreshActive();
            Check(restored.IsSelected(pages[2]) && restored.TabCount == 3 && source.Frame == layout.Frame,
                "Opening a saved bookmark did not restore its tab in the same editor.");
            layout.CheckHitTargets(restored.Host.Handle, slot);
        }
        Console.WriteLine("PASS disk restore in OneNote: two notebook headings in alphabetical order, reopened the bookmarked page under the second notebook in the same editor.");
    }

    private static void ObserveHandsOn(NotebookPages source, PageStrip strip, NativeLayout layout, List<string> pages)
    {
        Console.WriteLine("READY for hands-on input on Dashboard / MoreTabs A. Choose Close prototype when finished.");
        string[] labels = { "Dashboard A", "Schedule A", "Dashboard B" };
        int lastPage = -2;
        Rectangle lastSlot = Rectangle.Empty, lastBody = Rectangle.Empty;
        uint lastDpi = 0;
        Program.RunInteractive(layout, strip, delegate(Rectangle slot)
        {
            Rectangle body = layout.DocumentBounds;
            uint dpi = layout.Dpi;
            if (slot != lastSlot || body != lastBody || dpi != lastDpi)
            {
                layout.CheckHitTargets(strip.Host.Handle, slot);
                Console.WriteLine("Layout sampled: row {0}, document {1}, DPI {2}; all three rows reachable.", slot, body, dpi);
                lastSlot = slot; lastBody = body; lastDpi = dpi;
            }
            int index = pages.IndexOf(source.Current);
            if (index != lastPage)
            {
                Console.WriteLine(index < 0 ? "Page outside the three test pages." :
                    "Displayed " + labels[index] + ", tab highlighted: " + strip.IsSelected(pages[index]));
                lastPage = index;
            }
        });
        // Read only the disposable B page. Report presence of the marker, never the contents.
        string xml;
        source.App.GetPageContent(pages[2], out xml, One.PageInfo.piBasic, One.XMLSchema.xs2013);
        var page = XElement.Parse(xml);
        XNamespace ns = page.Name.Namespace;
        bool typed = page.Elements(ns + "Outline").Descendants(ns + "T").Any(node => node.Value.Contains("MoreTabs focus OK"));
        Console.WriteLine("Typing marker in Dashboard B page body: " + (typed ? "found" : "not observed"));
    }
}
