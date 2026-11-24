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
        internal Color BookBColor = Color.DarkOrange;
        public Color SectionColor(string id) { return id.StartsWith("a/") ? Color.Teal : BookBColor; }
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
                            layout.CheckHitTargets(strip.Handle, slot);
                            Check(slot == layout.Reserve(), "A repeated tick consumed another row.");
                            Check(slot.Top == height && slot.Bottom == layout.NavigationBounds.Top && slot.Width == expected.Width &&
                                layout.DocumentBounds.Bottom == expected.Bottom, "The test strip overlaps native controls.");
                        }
                        using (var cover = new Panel { Bounds = new Rectangle(0, 0, document.Width / 3, navigation.Height) })
                        {
                            document.Controls.Add(cover);
                            cover.BringToFront();
                            bool detected = false;
                            try { layout.CheckHitTargets(strip.Handle, slot); }
                            catch (InvalidOperationException error) { detected = error.Message.Contains("navigation"); }
                            Check(detected, "The hit test accepted a control covering the left of the navigation row.");
                        }
                        layout.CheckHitTargets(strip.Handle, slot);
                        Console.WriteLine("PASS detected a partially covered navigation row and its recovery.");
                        string[] ids = { "a/1", "a/2", "b/1" };
                        foreach (var id in ids) { source.Navigate(id); strip.ClickKeepCurrent(); }
                        strip.ClickKeepCurrent();
                        Check(strip.TabCount == 3, "Duplicate add created another tab.");
                        AppearanceTests.Controlled(strip, layout, slot);
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
                        source.BookBColor = Color.Empty;
                        var colorWait = System.Diagnostics.Stopwatch.StartNew();
                        while (!selected.SectionColor.IsEmpty && colorWait.ElapsedMilliseconds < 5000)
                        {
                            Thread.Sleep(50);
                            strip.RefreshActive();
                        }
                        var renderer = (TabRenderer)strip.Renderer;
                        Check(selected.Checked && selected.SectionColor.IsEmpty && renderer.MarkerFor(selected) == renderer.ActiveMarker,
                            "Changing the current section to automatic color did not restore the theme accent.");
                        Console.WriteLine("PASS section marker follows pages across notebooks and refreshes an edited section color without navigation.");
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
                        int position = strip.Items.IndexOf(moving);
                        source.Navigate("a/1"); strip.RefreshActive();
                        Check(moving.Checked && !original.Checked && moving != original && moving.Text == original.Text &&
                            strip.IsSelected("a/1"), "Native navigation to an already-open page switched tab slots.");
                        original.PerformClick();
                        Check(original.Checked && !moving.Checked, "Clicking a second tab for the same page failed to select that tab.");
                        moving.PerformClick();
                        source.Navigate("other"); strip.RefreshActive();
                        Check(moving.Checked && strip.Items.IndexOf(moving) == position && ((SavedTab)moving.Tag).Page.Id == "other" &&
                            ((SavedTab)original.Tag).Page.Id == "a/1", "Native navigation did not update the active tab in place.");
                        Console.WriteLine("PASS Quick Note title edits and native navigation preserve the active tab's identity, position and independent selection.");
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
                layout.CheckHitTargets(strip.Handle, slot);
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
                    string id = ((SavedTab)strip.Items.OfType<PageTabButton>().Last().Tag).Page.Id;
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
                int slotIndex = strip.Items.IndexOf(note);
                string noteId = ((SavedTab)note.Tag).Page.Id;
                RenameFixturePage(source.App, noteId, "Quick Note title check");
                WaitForLabel(strip, note, "Quick Note title check");
                source.Navigate(pages[0]); WaitForPage(source, pages[0]); strip.RefreshActive();
                Check(note.Checked && !samePage.Checked && note.Text == samePage.Text && strip.Items.IndexOf(note) == slotIndex &&
                    ((SavedTab)note.Tag).Page.Id == pages[0] && strip.TabCount == 5, "Navigation switched away from the active tab.");
                source.Navigate(noteId); WaitForPage(source, noteId);
                WaitForLabel(strip, note, "Quick Note title check");
                Console.WriteLine("PASS Quick Note rename, same-tab navigation to an already-open page and return; one active tab throughout.");
                string configured = source.ResolveQuickNotesSection(null), xml;
                source.App.GetHierarchy(configured, One.HierarchyScope.hsSelf, out xml, One.XMLSchema.xs2013);
                Check(XElement.Parse(xml).DescendantsAndSelf().Any(e => e.Name.LocalName == "Section" && (string)e.Attribute("ID") == configured),
                    "The configured Quick Notes location did not resolve to a section.");
                Console.WriteLine("PASS resolved the user's configured Quick Notes section without creating a personal page.");
                strip.Hide();
                try { VerifySavedTabs(source, pages, layout, slot, Path.Combine(directory, "working-set.xml")); }
                finally { strip.Show(); }
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
        using (var first = new PageStrip(layout.Frame, source, delegate { }, e => { throw e; }, store))
        {
            first.Attach(slot, layout);
            foreach (var id in pages)
            {
                source.Navigate(id); WaitForPage(source, id); first.ClickKeepCurrent();
            }
            var saved = store.Read().Tabs[2];
            first.SaveBookmark(saved.Page);
            first.CloseTab(saved);
        }
        source.Navigate(pages[0]); WaitForPage(source, pages[0]);
        using (var restored = new PageStrip(layout.Frame, source, delegate { }, e => { throw e; }, new SessionStore(path)))
        {
            restored.Attach(slot, layout);
            restored.RestoreOrSeed();
            Check(restored.TabCount == 2 && restored.IsSelected(pages[0]), "Reopening MoreTabs did not restore the two remaining tabs.");
            Check(restored.OpenSaved(store.Read().Bookmarks.Single()), "Could not reopen bookmarked page.");
            WaitForPage(source, pages[2]); restored.RefreshActive();
            Check(restored.IsSelected(pages[2]) && restored.TabCount == 3 && source.Frame == layout.Frame,
                "Opening a saved bookmark did not restore its tab in the same editor.");
            layout.CheckHitTargets(restored.Handle, slot);
        }
        Console.WriteLine("PASS disk restore in OneNote: kept two tabs, reopened closed tab from bookmark in the other notebook, same editor.");
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
                layout.CheckHitTargets(strip.Handle, slot);
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
