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
        internal string ListSelection;
        internal bool RejectNextMenuNavigation;
        public string SelectedPage(string sourceId) { return ListSelection; }
        internal readonly List<string> PickerReads = new List<string>();
        internal XElement[] ExtraNotebooks = new XElement[0];
        internal bool PickerUnavailable;
        internal Action DuringNavigation;
        public IEnumerable<XElement> PickerChildren(string parentId)
        {
            PickerReads.Add(parentId);
            if (PickerUnavailable) throw new System.Runtime.InteropServices.COMException("Fixture section unavailable");
            var tree = XElement.Parse("<Notebooks><Notebook ID='a' name='Atlas' color='#8AA8E4'><Section ID='as' name='Projects' color='#91BAAE'>" +
                "<Page ID='a/1' name='Dashboard'/></Section></Notebook><Notebook ID='b' name='B' nickname='Workshop'>" +
                "<SectionGroup ID='bin' name='Recycle Bin' isRecycleBin='true'/><SectionGroup ID='g' name='Plans &amp; specs'>" +
                "<Section ID='bs' name='Projects' color='#F6B078'><Page ID='b/1' name='Dashboard'/>" +
                "<Page ID='b/child' name='Notes' pageLevel='2'/><Page ID='b/grand' name='Details' pageLevel='3'/>" +
                "<Page ID='b/sibling' name='Details' pageLevel='2'/><Page ID='b/next' name='Next' pageLevel='invalid'/>" +
                "<Page ID='b/jump' name='Notes' pageLevel='3'/></Section>" +
                "<Section ID='empty' name='Archive'/></SectionGroup></Notebook></Notebooks>");
            return parentId == "" ? tree.Elements().Concat(ExtraNotebooks) : tree.Descendants().Single(x => (string)x.Attribute("ID") == parentId).Elements();
        }
        internal Color BookBColor = Color.FromArgb(246,176,120);
        internal BookmarkNotebook BookB = new BookmarkNotebook("b", "Workshop", Color.FromArgb(232, 209, 153));
        public Color SectionColor(string id) { return id.StartsWith("a/") ? Color.FromArgb(145,186,174) : BookBColor; }
        public Dictionary<string, BookmarkLocation> PageLocations(IEnumerable<string> ids, bool refresh = false)
        {
            var a = new BookmarkNotebook("a", "Atlas", Color.MidnightBlue);
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
        internal string LastNoteSection;
        public string CreateNote(string sectionId)
        {
            if (LocationUnavailable) throw new InvalidOperationException("No configured Quick Notes section.");
            LastNoteSection = sectionId;
            return (sectionId ?? "personal/quick") + "/" + ++NotesCreated;
        }
        public void Navigate(string id)
        {
            if (RejectNextMenuNavigation)
            {
                RejectNextMenuNavigation = false;
                throw new System.Runtime.InteropServices.COMException("Editor is dismissing the menu", unchecked((int)0x80010001));
            }
            if (FailNextNavigation) { FailNextNavigation = false; throw new System.Runtime.InteropServices.COMException(); }
            if (DelayNextNavigation) { DelayNextNavigation = false; PendingPage = id; return; }
            Current = id;
            if (DuringNavigation != null) DuringNavigation();
        }
        public string LinkFor(string id) { return ""; }
        public void NavigateLink(string link) { throw new NotSupportedException(); }
    }
    internal static void ControlledWindow()
    {
        CheckVerticalTabs();
        CheckPagePicker();
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
                            Check(slot.Top == height && slot.Bottom == layout.LandmarkBounds.Top && slot.Width == expected.Width &&
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
                        source.BookB = new BookmarkNotebook("b", "Workshop projects and reference material", previousBook.Color);
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
                        strip.ClickNewQuickNote();
                        Check(strip.HasTabFocus(source.Current), "The first Quick Note tab did not receive keyboard focus.");
                        var note = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                        Check(note.Text == "Untitled", "New Quick Note did not adopt OneNote's initial title.");
                        source.Renamed[source.Current] = "Trip plans / Personal";
                        WaitForLabel(strip, note, "Trip plans");
                        strip.ClickNewQuickNote();
                        Check(source.NotesCreated == 2 && strip.TabCount == 5 && source.Current == "personal/quick/2" &&
                            strip.IsSelected(source.Current) && strip.HasTabFocus(source.Current), "Each click should create, open and focus a distinct Quick Note tab.");
                        source.LocationUnavailable = true;
                        strip.ClickNewQuickNote();
                        Check(source.NotesCreated == 2 && strip.TabCount == 5 && source.Current == "personal/quick/2",
                            "Missing Quick Notes destination changed tabs or created a page elsewhere.");
                        source.LocationUnavailable = false;
                        source.FailNextNavigation = true;
                        strip.ClickNewQuickNote();
                        strip.ClickTab("personal/quick/3");
                        Check(source.NotesCreated == 3 && strip.TabCount == 6 && strip.IsSelected("personal/quick/3"),
                            "Retrying a created Quick Note lost the page or created a duplicate.");
                        source.DelayNextNavigation = true;
                        strip.ClickNewQuickNote();
                        Check(strip.IsSelected("personal/quick/3") && !strip.HasTabFocus(source.PendingPage),
                            "Focus moved to the new tab before its page opened.");
                        source.Navigate(source.PendingPage); strip.RefreshActive();
                        Check(strip.IsSelected(source.Current) && strip.HasTabFocus(source.Current), "Delayed navigation lost new-tab focus.");
                        Console.WriteLine("PASS New Quick Note creates distinct pages from another notebook and focuses each new tab; missing destination and navigation retry preserve the working set.");
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
                        string target = PageMenuCommand.SelectedPage(section, "a/1");
                        Check(target == "b/1", "The context menu chose the displayed page over the selected page.");
                        section.Elements().First().SetAttributeValue("selected", "partial");
                        Check(PageMenuCommand.SelectedPage(section, "a/1") == target, "A text selection hid the selected page title.");
                        section.Elements().First().SetAttributeValue("selected", "all");
                        Check(PageMenuCommand.SelectedPage(section, "a/1") == null && PageMenuCommand.SelectedPage(new XElement("Section"), "a/1") == null,
                            "An ambiguous or empty selection was accepted.");
                        section.Elements().First().Attribute("selected").Remove();
                        section.Elements().Last().SetAttributeValue("selected", "partial");
                        Check(PageMenuCommand.SelectedPage(section, "a/1") == target, "A single partial page selection was ignored.");
                        section.Elements().Last().Attribute("selected").Remove();
                        Check(PageMenuCommand.SelectedPage(section, "a/1") == "a/1" && PageMenuCommand.SelectedPage(section, "b/1") == null,
                            "The unmarked selection did not match the viewed page in the invoking window.");
                        int count = strip.TabCount, notes = source.NotesCreated;
                        AppearanceTests.TogglePin(strip, moving, true);
                        source.ListSelection = target;
                        source.RejectNextMenuNavigation = true;
                        Check(PageMenuCommand.Request(document.Handle, true), "The page-list command could not reach its editor's strip.");
                        Application.DoEvents();
                        Check(strip.TabCount == count, "OneNote's busy response created a tab before navigation.");
                        source.ListSelection = "a/1";
                        Check(strip.ProcessMenuRequest(), "The page-list command was lost while OneNote was busy.");
                        string firstNew = strip.ActiveTabKey;
                        Check(strip.TabCount == count + 1 && firstNew != movingKey && strip.IsSelected(target) &&
                            ((SavedTab)moving.Tag).Pinned && ((SavedTab)moving.Tag).Page.Id == "other", "Opening from the native page list replaced the source tab.");
                        Check(PageMenuCommand.Request(frame.Handle, false), "Could not open the displayed page from its text menu.");
                        Application.DoEvents();
                        Check(strip.TabCount == count + 2 && strip.ActiveTabKey != firstNew && strip.HasTabFocus(target) &&
                            !((SavedTab)strip.Items.OfType<PageTabButton>().Single(t => t.Checked).Tag).Pinned && source.NotesCreated == notes,
                            "Opening an existing page should create a focused, unpinned tab each time.");
                        source.FailNextNavigation = true;
                        strip.OpenPageInNewTab("a/1");
                        Check(strip.TabCount == count + 2, "Failed navigation created an empty tab.");
                        source.ListSelection = null;
                        Check(PageMenuCommand.Request(frame.Handle, true), "The receiver rejected a page-list request.");
                        Application.DoEvents();
                        Check(strip.TabCount == count + 2, "An ambiguous selection opened a tab.");
                        using (var other = new Form())
                            Check(!PageMenuCommand.Request(other.Handle, true), "A second editor delivered its click to the wrong strip.");
                        Console.WriteLine("PASS page-list command resolves a selected ID, preserves the pinned source, and opens distinct focused tabs.");
                        var originalActive = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                        var sourceTab = (SavedTab)moving.Tag;
                        strip.TabMenuFor(sourceTab).Items.Cast<ToolStripItem>().Single(i => i.Text == "Open in new tab").PerformClick();
                        var opened = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                        Check(strip.TabCount == count + 3 && ((SavedTab)opened.Tag).Key != sourceTab.Key &&
                            !((SavedTab)opened.Tag).Pinned && ((SavedTab)moving.Tag).Pinned && source.Current == sourceTab.Page.Id &&
                            ((SavedTab)originalActive.Tag).Page.Id == target &&
                            strip.HasTabFocus(source.Current) && source.NotesCreated == notes, "Right-clicking an inactive tab opened the wrong page or replaced its source.");
                        strip.CloseTab((SavedTab)opened.Tag);
                        originalActive.PerformClick();
                        Console.WriteLine("PASS MoreTabs right-click opens the clicked page in an independent focused tab.");
                        CheckPageMenuPreselection(strip, source, moving, frame.Handle);
                        originalActive.PerformClick();
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

    private static void CheckPageMenuPreselection(PageStrip strip, TestPages source, PageTabButton original, IntPtr frame)
    {
        int count = strip.TabCount;
        var slot = (SavedTab)original.Tag;
        string target = "b/1";
        foreach (bool lateInput in new[] { false, true })
        {
            original.PerformClick();
            int pressed = Environment.TickCount;
            if (!lateInput) strip.MarkPageMenuGesture(pressed);
            // OneNote has already navigated, and ordinary polling has already changed the source tab.
            source.Navigate(target);
            strip.RefreshActive();
            Check(((SavedTab)original.Tag).Page.Id == target, "Fixture did not reproduce early native navigation.");
            if (lateInput) strip.MarkPageMenuGesture(pressed);
            source.ListSelection = target;
            Check(PageMenuCommand.Request(frame, true), "Could not dispatch the delayed page-menu click.");
            Application.DoEvents();
            var opened = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
            var restored = (SavedTab)original.Tag;
            Check(strip.TabCount == count + 1 && original != opened && restored.Key == slot.Key &&
                restored.Page.Id == slot.Page.Id && restored.Pinned == slot.Pinned && strip.HasTabFocus(target) &&
                !((SavedTab)opened.Tag).Pinned, "The native page selection replaced the source tab before the command arrived.");
            original.PerformClick();
            Check(source.Current == slot.Page.Id, "The original tab did not navigate back to its original page.");
            strip.CloseTab((SavedTab)opened.Tag);
        }
        source.Navigate(target);
        strip.RefreshActive();
        // A completed left-click followed by right-clicking that same page should leave two copies.
        strip.MarkPageMenuGesture(unchecked(Environment.TickCount + 1));
        Check(PageMenuCommand.Request(frame, true), "Could not dispatch a same-page menu click.");
        Application.DoEvents();
        var duplicate = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
        Check(strip.TabCount == count + 1 && duplicate != original && ((SavedTab)original.Tag).Page.Id == target,
            "The menu undid a previous left-click.");
        strip.CloseTab((SavedTab)duplicate.Tag);
        original.PerformClick();
        source.Navigate(slot.Page.Id);
        strip.RefreshActive();
        var section = XElement.Parse("<Section><Page ID='a'/><Page ID='b'/></Section>");
        Check(NotebookPages.MenuSelection(section, "b", "a") == "b" && NotebookPages.MenuSelection(section, "b", "b") == null,
            "An unmarked hierarchy failed to distinguish a right-click navigation from an unknown selection.");
        foreach (var row in section.Elements()) row.SetAttributeValue("selected", "all");
        Check(NotebookPages.MenuSelection(section, "b", "a") == null, "Multiple selected pages were treated as a single click.");
        Console.WriteLine("PASS page-menu preselection, late input delivery, return to source, pin preservation and ordinary left-click navigation.");
    }

    private static PagePicker ClickPicker(PageStrip strip)
    {
        var picker = strip.Items.OfType<PagePicker>().Single();
        LeftClickItem(picker);
        Check(picker.DropDown.Visible, "Left-clicking the + glyph did not display the page picker.");
        return picker;
    }
    private static void LeftClickItem(ToolStripItem item)
    {
        var bounds = item.Bounds;
        var point = (IntPtr)(((bounds.Top + bounds.Height / 2) << 16) | (bounds.Left + bounds.Width / 2));
        var handle = item.GetCurrentParent().Handle;
        SendMessage(handle, 0x0201, (IntPtr)1, point);
        SendMessage(handle, 0x0202, IntPtr.Zero, point);
        Application.DoEvents();
    }

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int GetMessagePos();
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr SetActiveWindow(IntPtr window);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern IntPtr GetActiveWindow();

    private static void CheckActivationClicks(Form frame, PageStrip strip)
    {
        using (var other = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual, Location = frame.Location })
        {
            other.Show();
            foreach (var button in strip.Items.OfType<ToolStripDropDownButton>())
            {
                SetActiveWindow(other.Handle);
                Check(GetActiveWindow() == other.Handle, "The second test window must be active before the click (" +
                    button.AccessibleName + ", got " + GetActiveWindow() + ", expected " + other.Handle + ").");
                var bounds = button.Bounds;
                var center = strip.PointToScreen(new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2));
                var mouse = new Point(GetMessagePos());
                var origin = strip.Host.Location;
                // ToolStrip hit-tests the last queued mouse position. Move its clipped child inside
                // our offscreen frame so the activation probe never needs to move the user's mouse.
                IntPtr result;
                try
                {
                    strip.Host.Location = new Point(origin.X + mouse.X - center.X, origin.Y + mouse.Y - center.Y);
                    result = SendMessage(strip.Handle, 0x0021, frame.Handle, (IntPtr)((0x0201 << 16) | 1));
                }
                finally { strip.Host.Location = origin; }
                Check(result == (IntPtr)1, button.AccessibleName + " must activate and keep the first click (reply " + result + ").");
                frame.Activate();
                LeftClickItem(button);
                Check(button.DropDown.Visible, button.AccessibleName + " did not open on the activation click.");
                button.HideDropDown();
            }
        }
        Console.WriteLine("PASS first click from a second active window opens bookmarks, options and the page picker.");
    }
    private static ToolStripMenuItem PickerItem(ToolStripDropDownItem parent, string id)
    {
        return parent.DropDownItems.OfType<ToolStripMenuItem>().Single(item => (string)item.Tag == id);
    }
    private static ToolStripMenuItem ExpandPickerItem(ToolStripDropDownItem parent, string id)
    {
        var item = PickerItem(parent, id);
        item.ShowDropDown();
        return item;
    }
    private static void CheckPagePicker()
    {
        var source = new TestPages();
        source.Navigate("a/1");
        using (var frame = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual, Location = new Point(-16000,-16000) })
        using (var strip = new PageStrip(frame.Handle, source, delegate {}, e => { throw e; }))
        {
            frame.Show();
            strip.Attach(new Rectangle(0, 0, 700, HostChrome.StripHeight(96)));
            strip.RestoreOrSeed();
            var original = strip.Items.OfType<PageTabButton>().Single();
            var picker = ClickPicker(strip);
            Check(source.PickerReads.SequenceEqual(new[] { "" }) && PickerItem(picker, "b").Text == "Workshop",
                "Opening the picker should list the open notebooks by their displayed names.");
            var book = ExpandPickerItem(picker, "b");
            Check(book.DropDownItems.Count == 1, "The recycle bin appeared in the picker.");
            var group = ExpandPickerItem(book, "g");
            Check(group.AccessibleName == "Plans & specs", "A section group name was altered.");
            var empty = ExpandPickerItem(group, "empty");
            Check(!empty.DropDownItems[0].Enabled && empty.DropDownItems[0].Text == "No pages", "Missing empty-section state.");
            empty.HideDropDown();
            source.PickerUnavailable = true;
            var section = ExpandPickerItem(group, "bs");
            Check(!section.DropDownItems[0].Enabled, "An unavailable section was selectable.");
            source.PickerUnavailable = false;
            section.HideDropDown();
            picker.HideDropDown();
            Check(strip.TabCount == 1 && original.Checked && source.Current == "a/1" && source.NotesCreated == 0,
                "Browsing or canceling the picker changed the working set.");
            source.DuringNavigation = strip.RefreshActive;
            foreach (int expected in new[] { 2, 3 })
            {
                picker = ClickPicker(strip);
                section = ExpandPickerItem(ExpandPickerItem(ExpandPickerItem(picker, "b"), "g"), "bs");
                var rows = section.DropDownItems.OfType<PickerRow>().ToArray();
                Check(rows.Select(row => row.PageDepth).SequenceEqual(new[] { 0, 1, 2, 1, 0, 1 }) &&
                    rows.All(row => !row.HasDropDownItems) && rows[2].AccessibleDescription == "Dashboard / Notes / Details" &&
                    rows[3].ToolTipText == "Dashboard / Details" && rows[5].ToolTipText == "Next / Notes",
                    "Page nesting, accessible ancestry, or selection of a parent with subpages was lost.");
                string targetId = expected == 2 ? "b/1" : "b/grand";
                var target = PickerItem(section, targetId);
                if (expected == 2) LeftClickItem(target);
                else
                {
                    target.Select();
                    var enter = Message.Create(section.DropDown.Handle, 0x0100, (IntPtr)Keys.Enter, IntPtr.Zero);
                    Check(section.DropDown.PreProcessMessage(ref enter), "Enter did not select the highlighted subpage.");
                }
                Check(!picker.DropDown.Visible && strip.TabCount == expected && source.Current == targetId && strip.HasTabFocus(targetId) &&
                    ((SavedTab)original.Tag).Page.Id == "a/1" && source.NotesCreated == 0,
                    "Picking an existing page must open a separate focused tab without changing the source tab.");
            }
            source.DuringNavigation = null;
            original.PerformClick();
            Check(source.Current == "a/1", "The original picker tab did not return to its page.");
            LeftClickItem(ClickPicker(strip).KeepCurrent);
            Check(strip.TabCount == 3 && source.Current == "a/1" && source.NotesCreated == 0,
                "Keeping the current page from the picker changed its existing tab.");
            LeftClickItem(ClickPicker(strip).QuickNote);
            Check(strip.TabCount == 4 && source.NotesCreated == 1 && strip.HasTabFocus(source.Current),
                "Selecting New Quick Note from the picker did not create and focus exactly one new note.");
            Console.WriteLine("PASS picker navigation, both levels of subpages, duplicate titles, selecting parents and opening a subpage with Enter.");
            CheckTools(strip, source);
            CheckToolVisibility(strip);
            CheckActivationClicks(frame, strip);
            string longTitle = "Gallery & workshop planning notes for the autumn exhibition";
            source.ExtraNotebooks = Enumerable.Range(1, 60).Select(i => new XElement("Notebook",
                new XAttribute("ID", "extra" + i), new XAttribute("name", i == 1 ? longTitle : "Notebook " + i),
                new XAttribute("color", "#318A75"))).ToArray();
            AppearanceTests.CheckNotebookPicker(strip, longTitle);
        }
    }

    private static void CheckToolVisibility(PageStrip strip)
    {
        var gear = (StripMenuButton)strip.Items[1];
        var menu = gear.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text == "Visible tool buttons");
        var toggles = menu.DropDownItems.OfType<ToolStripMenuItem>().ToArray();
        Check(toggles.Select(i => i.Text).SequenceEqual(new[] { "Bookmarks", "Back", "Forward", "Home", "New note" }) &&
            toggles.All(i => i.Checked), "Settings must offer exactly the five optional Tools buttons, all initially visible.");
        int fullWidth = strip.MenuTabBounds.Width;
        string key = strip.ActiveTabKey;
        foreach (var toggle in toggles.Reverse())
        {
            LeftClickItem(gear);
            menu.ShowDropDown();
            LeftClickItem(toggle);
            var visible = strip.Items.Cast<ToolStripItem>().TakeWhile(i => !(i is ToolStripSeparator)).Where(i => i.Available).ToArray();
            Check(!toggle.Checked && !strip.Items.Cast<ToolStripItem>().Single(i => i.AccessibleName == toggle.Text).Available &&
                strip.MenuTabBounds == Rectangle.Union(visible.First().Bounds, visible.Last().Bounds) &&
                gear.Placement == ToolStripItemPlacement.Main, "Hiding " + toggle.Text + " left an incorrect Tools outline or checkmark.");
        }
        Check(strip.MenuTabBounds == gear.Bounds && gear.Width < fullWidth && strip.ActiveTabKey == key,
            "Settings alone must keep its hit target and release the space used by hidden buttons.");
        int width = strip.Host.Width;
        strip.Host.Width = 170;
        LeftClickItem(gear);
        menu.ShowDropDown();
        LeftClickItem(toggles.Single(i => i.Text == "Back"));
        strip.OverflowButton.ShowDropDown();
        var overflow = strip.Items.OfType<ToolButton>().Where(t => t.Available && t.GetCurrentParent() == strip.OverflowButton.DropDown).ToArray();
        Check(overflow.Length == 1 && overflow[0].Name == "Back", "Hidden tools appeared in the overflow menu.");
        strip.OverflowButton.HideDropDown();
        strip.Host.Width = width;
        foreach (var toggle in toggles.Where(i => !i.Checked)) toggle.PerformClick();
        Check(toggles.All(i => i.Checked) && strip.MenuTabBounds.Width == fullWidth && strip.ActiveTabKey == key,
            "Restoring the tools lost a button or changed the active tab.");
        Console.WriteLine("PASS individual Tools checkboxes, gear-only layout, overflow and recovery.");
    }

    private static void CheckTools(PageStrip strip, TestPages source)
    {
        var tools = strip.Items.OfType<ToolButton>().ToArray();
        Check(tools.Select(t => t.AccessibleName).SequenceEqual(new[] { "Back", "Forward", "Home", "New note" }),
            "The Tools commands must follow bookmarks and settings.");
        var gear = (StripMenuButton)strip.Items[1];
        var homePicker = gear.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text.StartsWith("Home page"));
        var sectionPicker = gear.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text.StartsWith("New note location"));
        var home = tools[2];
        var note = tools[3];
        string before = source.Current, key = strip.ActiveTabKey;
        int count = strip.TabCount, notes = source.NotesCreated;
        LeftClickItem(home);
        Check(homePicker.DropDown.Visible && source.Current == before, "First Home click must open the settings picker.");
        Check(PickerItem(homePicker, "b") is PickerRow, "Home must use the same notebook list as +.");
        var section = ExpandPickerItem(ExpandPickerItem(ExpandPickerItem(homePicker, "b"), "g"), "bs");
        Check(section is PickerRow && section.DropDown.Renderer == homePicker.DropDown.Renderer &&
            section.DropDown.Width == homePicker.DropDown.Width && ((PickerRow)PickerItem(section, "b/grand")).PageDepth == 2,
            "Home must share the section and subpage layout with +.");
        LeftClickItem(PickerItem(section, "b/grand"));
        Check(source.Current == "b/grand" && strip.ActiveTabKey == key && strip.TabCount == count,
            "Setting Home on first use must open that page in the current tab.");
        LeftClickItem(gear);
        homePicker.ShowDropDown();
        section = ExpandPickerItem(ExpandPickerItem(homePicker, "a"), "as");
        LeftClickItem(PickerItem(section, "a/1"));
        Check(source.Current == "b/grand", "Configuring Home in Settings unexpectedly navigated.");
        LeftClickItem(home);
        Check(source.Current == "a/1" && strip.ActiveTabKey == key, "The Home button did not use its new page.");
        LeftClickItem(note);
        Check(source.LastNoteSection == null && source.NotesCreated == notes + 1 && strip.HasTabFocus(source.Current),
            "New note must default to Quick Notes and focus its new tab. Created " + (source.NotesCreated - notes) +
            ", section " + (source.LastNoteSection ?? "(default)") + ", active " + strip.IsSelected(source.Current) +
            ", focus " + strip.ContainsFocus + ", selected item " + String.Join(", ", strip.Items.Cast<ToolStripItem>().Where(i => i.Selected).Select(i => i.AccessibleName)) + ".");
        LeftClickItem(gear);
        sectionPicker.ShowDropDown();
        Check(PickerItem(sectionPicker, "b") is PickerRow, "New note location must use the same notebook list as +.");
        var group = ExpandPickerItem(ExpandPickerItem(sectionPicker, "b"), "g");
        var target = PickerItem(group, "bs");
        Check(target is PickerRow && !target.HasDropDownItems && group.DropDown.Renderer == sectionPicker.DropDown.Renderer &&
            group.DropDown.Width == sectionPicker.DropDown.Width && target.Height == group.Height,
            "The new note destination picker must show the shared section row and stop at that section.");
        LeftClickItem(target);
        Check(source.NotesCreated == notes + 1 && sectionPicker.Text.Contains("Workshop / Plans && specs / Projects"),
            "Choosing a section must only change the destination setting.");
        LeftClickItem(note);
        Check(source.LastNoteSection == "bs" && source.NotesCreated == notes + 2 && strip.TabCount == count + 2 &&
            strip.HasTabFocus(source.Current), "New note did not use the selected section and focus its own tab.");
        source.LocationUnavailable = true;
        LeftClickItem(note);
        source.LocationUnavailable = false;
        Check(source.NotesCreated == notes + 2 && strip.TabCount == count + 2,
            "An unavailable chosen section must not create a page or tab elsewhere.");
        LeftClickItem(ClickPicker(strip).QuickNote);
        Check(source.LastNoteSection == null && source.NotesCreated == notes + 3, "The explicit Quick Note command lost its destination.");
        LeftClickItem(gear);
        sectionPicker.ShowDropDown();
        LeftClickItem(sectionPicker.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text == "Use default Quick Notes"));
        LeftClickItem(note);
        Check(source.LastNoteSection == null && source.NotesCreated == notes + 4, "Resetting the destination must restore Quick Notes.");
        Console.WriteLine("PASS Tools, first-use Home, configured Home, section-only note picker, creation, failure and reset.");
    }

    private static void CheckVerticalTabs()
    {
        foreach (var surface in new[] { Color.FromArgb(250,250,250), Color.FromArgb(42,38,36), Color.FromArgb(0x52,0x4f,0x4c) })
        {
            int rgb = ColorTranslator.ToWin32(surface);
            Color measured;
            Check(HostChrome.NavigationSurface(new List<int> { rgb, rgb, 0x3254ab, rgb, 0x1146ff }, out measured) &&
                measured.ToArgb() == surface.ToArgb(), "Navigation background lost its Office theme or picked an icon.");
        }
        Color unknown;
        Check(!HostChrome.NavigationSurface(new List<int>(), out unknown) &&
            !HostChrome.NavigationSurface(new List<int> { 1, 2, 3, 4, 5 }, out unknown), "An obscured or ambiguous navigation header produced a theme.");
        using (var frame = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual,
            Location = new Point(-16000, -16000), ClientSize = new Size(800, 500) })
        using (var ribbon = new Panel { Bounds = new Rectangle(0, 0, 800, 80) })
        using (var body = new Panel { Bounds = new Rectangle(0, 80, 800, 420) })
        using (var row = new Panel { Dock = DockStyle.Top, Height = 36 })
        using (var books = new Panel { Bounds = new Rectangle(0, 36, 124, 384) })
        using (var pages = new Panel { Bounds = new Rectangle(124, 12, 120, 408) })
        using (var editor = new Panel { Bounds = new Rectangle(244, 12, 556, 408) })
        {
            frame.Controls.AddRange(new Control[] { ribbon, body });
            body.Controls.AddRange(new Control[] { books, pages, editor, row });
            row.SendToBack();
            frame.Show();
            var original = body.Bounds;
            using (var layout = new NativeLayout(frame.Handle, body.Handle, ribbon.Handle, row.Handle, editor.Handle))
            using (var strip = new PageStrip(frame.Handle, null, delegate {}, e => { throw e; }))
            {
                foreach (bool bottom in new[] { false, true })
                {
                    Rectangle slot = layout.Reserve(false, bottom);
                    strip.Attach(slot, layout);
                    layout.CheckHitTargets(strip.Host.Handle, slot);
                    Check(layout.NavigationHeader == row.Handle, "Vertical tabs lost the notebook header behind the overlapping panes.");
                    Check(layout.VerticalTabs && layout.StillReserved(slot) && slot.Width == original.Width &&
                        body.Height + slot.Height == original.Height && (bottom ? body.Bottom == slot.Top : body.Top == slot.Bottom),
                        "Vertical layout failed to reserve the strip outside the editor and notebook panes.");
                    foreach (var pane in new[] { books, pages, editor })
                        Check(body.GetChildAtPoint(new Point(pane.Left + pane.Width / 2, 70)) == pane,
                            "A vertical notebook pane or the page editor lost its click target.");
                }
                var vertical = editor.Bounds;
                editor.Top = row.Height;
                row.BringToFront();
                strip.Attach(layout.Reserve(), layout);
                Check(!layout.VerticalTabs && layout.AtBottom && layout.NavigationHeader == row.Handle,
                    "Switching to horizontal tabs lost the dock or the native navigation row.");
                // This fixture is off-screen. Seed the last horizontal reading before hiding its header.
                var chrome = layout.Appearance();
                var previousColor = chrome.HighContrast ? SystemColors.Window : Color.FromArgb(0x52, 0x4f, 0x4c);
                chrome.Surface = previousColor;
                chrome.Measured = true;
                row.Hide();
                books.Hide();
                pages.Hide();
                editor.Bounds = new Rectangle(0, 0, body.Width, body.Height);
                strip.Attach(layout.Reserve(false, false), layout);
                Check(layout.VerticalTabs && layout.StillReserved(strip.Host.Bounds), "A collapsed sidebar or hidden row prevented attachment.");
                Check(layout.NavigationHeader == IntPtr.Zero && strip.BackColor.ToArgb() == previousColor.ToArgb(),
                    "With its header hidden, the vertical strip lost the last navigation color.");
                using (var cover = new Panel { Dock = DockStyle.Fill })
                {
                    body.Controls.Add(cover);
                    cover.BringToFront();
                    Check(layout.Reserve(true).IsEmpty, "A covered vertical editor was accepted.");
                }
                strip.PauseForDialog();
                layout.ReleaseForDialog();
                Check(body.Bounds == original, "Pausing vertical tabs left a gap.");
                row.Show();
                row.SendToBack();
                books.Show();
                pages.Show();
                editor.Bounds = vertical;
                strip.Attach(layout.Reserve(), layout);
                Check(layout.VerticalTabs, "The vertical layout did not resume after releasing the strip.");
            }
            Check(body.Bounds == original, "Closing vertical tabs did not restore the document.");
        }
        Console.WriteLine("PASS vertical tabs with overlapping navigation, expanded and collapsed panes, both docks, orientation changes and cleanup.");
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
            var toolsLabel = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Single(l => l.Text == "Tools" && l.Notebook == null);
            Check(toolsLabel.Visible && toolsLabel.Right == strip.MenuTabBounds.Right &&
                strip.Items.OfType<ToolButton>().All(t => t.Placement == ToolStripItemPlacement.Overflow),
                "A narrow window must keep the Tools caption aligned and its commands available in overflow.");
            var headers = strip.Host.Captions.Controls.OfType<NotebookTabLabel>().Where(l => l.Visible && l != toolsLabel).ToArray();
            AppearanceTests.CheckGroupDividers(strip, headers.Length);
            Check(visible.Any(t => t.Checked) && visible.All(t => headers.Contains(t.Group) && t.Group.Left <= t.Bounds.Left + TabRenderer.CornerRadius(t.Dpi) && t.Group.Right >= t.Bounds.Right) &&
                headers.All(h => visible.Any(t => t.Group == h)), "Overflow separated the active page from its notebook label.");
            foreach (var group in visible.GroupBy(t => t.Group))
            {
                int pad = group.Key.GroupFill.IsEmpty ? 0 : 2 * HostChrome.Px(1, layout.Dpi);
                Check(group.First().Bounds.Left - group.Key.Left == pad && group.Key.Right - group.Last().Bounds.Right == pad,
                    "A partially visible group must keep its padding around the tabs still on the strip.");
            }
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
                source.Window.Active = true;
                var activeKey = strip.ActiveTabKey;
                var history = strip.Items.OfType<ToolButton>().ToDictionary(b => b.AccessibleName);
                history["Back"].PerformClick();
                WaitForPage(source, pages[2], "Tools Back"); strip.RefreshActive();
                history["Forward"].PerformClick();
                WaitForPage(source, pages[0], "Tools Forward"); strip.RefreshActive();
                Check(strip.ActiveTabKey == activeKey && strip.TabCount == 3 && source.Frame == layout.Frame,
                    "Native history must navigate inside the existing tab and OneNote window.");
                Console.WriteLine("PASS Tools Back and Forward traversed OneNote history in the same tab.");
                AppearanceTests.Live(strip, layout, directory);
                var original = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                var originalPage = ((SavedTab)original.Tag).Page.Id;
                var picker = ClickPicker(strip);
                var sectionMenu = ExpandPickerItem(ExpandPickerItem(picker, books[1]), quickSection);
                PickerItem(sectionMenu, pages[2]).PerformClick();
                picker.HideDropDown();
                WaitForPage(source, pages[2], "opening an existing page from the picker");
                strip.RefreshActive();
                var picked = strip.Items.OfType<PageTabButton>().Single(t => t.Checked);
                Check(strip.TabCount == 4 && picked != original && strip.HasTabFocus(pages[2]) && source.Frame == layout.Frame,
                    "The picker did not open a separate focused tab in the fixture window.");
                original.PerformClick();
                WaitForPage(source, originalPage, "returning to the source tab after using the picker");
                strip.CloseTab((SavedTab)picked.Tag);
                var existing = new HashSet<string>(pages);
                source.Navigate(pages[0]); WaitForPage(source, pages[0], "selecting the starting notebook for Quick Notes");
                for (int i = 0; i < 2; i++)
                {
                    strip.ClickNewQuickNote();
                    Check(strip.TabCount == 4 + i, "Quick Note click " + (i + 1) + " did not add a new tab.");
                    string id = strip.Items.OfType<PageTabButton>().Select(t => ((SavedTab)t.Tag).Page.Id).Single(page => !existing.Contains(page));
                    WaitForPage(source, id, "opening Quick Note " + (i + 1)); strip.RefreshActive();
                    string parent;
                    source.App.GetHierarchyParent(id, out parent);
                    Check(existing.Add(id) && parent == quickSection && source.Window.CurrentNotebookId == books[1] &&
                        source.Frame == layout.Frame && strip.IsSelected(id) && strip.HasTabFocus(id) && strip.TabCount == 4 + i,
                        "New Quick Note did not open with its own focused tab in the disposable destination section.");
                }
                Console.WriteLine("PASS two New Quick Note selections created two pages in the disposable destination section and focused each new tab in this editor.");
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
