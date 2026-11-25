using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

internal static class SavedSessionTests
{
    private const string PageLink = "onenote:test.one#Dashboard&page-id=%7B591DBE87-355B-47CD-B495-BB5E51479233%7D";

    private sealed class Pages : IPages
    {
        internal bool reopened;
        public string Current { get; private set; }
        public string SelectedPage(string sourceId) { return Current; }
        public IEnumerable<System.Xml.Linq.XElement> PickerChildren(string parentId) { return Enumerable.Empty<System.Xml.Linq.XElement>(); }
        public Color SectionColor(string id) { return id == "a2" ? Color.Gold : Color.Indigo; }
        public Dictionary<string, BookmarkLocation> PageLocations(IEnumerable<string> ids, bool refresh = false)
        {
            var a = new BookmarkNotebook("a", "A", Color.Coral);
            var b = new BookmarkNotebook("b", "B & C", Color.Teal);
            return ids.Where(id => id != "missing" && !(reopened && id == "b1")).ToDictionary(id => id,
                id => new BookmarkLocation(id.StartsWith("b") ? b : a, "Projects", SectionColor(id)));
        }
        public string Caption(string id) { return id.StartsWith("b") ? "Dashboard / B & C" : "Dashboard / A"; }
        public string LinkFor(string id) { return id.StartsWith("b") ? PageLink : ""; }
        public string CreateNote(string sectionId) { return "quick-" + Guid.NewGuid().ToString("N"); }
        public void Navigate(string id)
        {
            if (id == "missing" || (reopened && id == "b1")) throw new COMException("Unavailable test ID");
            Current = id;
        }
        public void NavigateLink(string link)
        {
            if (link != PageLink) throw new COMException("Unavailable test link");
            Current = "b2";
        }
    }

    private static void Check(bool success, string message)
    {
        if (!success) throw new InvalidOperationException(message);
    }
    private static string DirectoryFor(Guid run)
    {
        return Path.Combine(Path.GetTempPath(), "MoreTabs-selftest-" + run.ToString("N"));
    }

    internal static void Run()
    {
        Guid run = Guid.NewGuid();
        string directory = DirectoryFor(run);
        Directory.CreateDirectory(directory);
        try
        {
            foreach (string phase in new[] { "save", "restore", "empty" })
            {
                var start = new ProcessStartInfo(Application.ExecutablePath, "sessiontest " + run.ToString("D") + " " + phase) {
                    UseShellExecute = false, CreateNoWindow = true
                };
                using (var child = Process.Start(start))
                {
                    if (!child.WaitForExit(30000)) { child.Kill(); child.WaitForExit(); throw new TimeoutException("Session test timed out: " + phase); }
                    Check(child.ExitCode == 0, "Session test failed: " + phase);
                }
            }
            ExerciseEditorLifetime(Path.Combine(directory, "reattach.xml"));
            ExerciseLegacyTabs(Path.Combine(directory, "legacy.xml"));
            string path = Path.Combine(directory, "race.xml");
            Parallel.For(0, 8, i => new SessionStore(path).Change(s =>
                WorkingSet.Keep(s.Bookmarks, new SavedPage(i.ToString(), "Duplicate title", ""))));
            Check(new SessionStore(path).Read().Bookmarks.Count == 8, "Concurrent writers lost bookmarks.");
            File.WriteAllText(path, "<incomplete");
            try { new SessionStore(path).Change(s => true); throw new Exception("Accepted a damaged state file."); }
            catch (InvalidDataException) { Check(File.ReadAllText(path) == "<incomplete", "Damaged state was overwritten."); }
            Console.WriteLine("PASS saved working set: fresh processes, tab order, bookmark reopen/remove, changed ID, missing page, empty session, concurrent writes and damaged-file preservation.");
        }
        finally
        {
            // Every file in this newly created directory was produced by these fixture phases.
            foreach (string file in Directory.GetFiles(directory)) File.Delete(file);
            Directory.Delete(directory);
        }
    }

    private static void ExerciseLegacyTabs(string path)
    {
        var sections = BookmarkLocation.IndexSections(System.Xml.Linq.XElement.Parse(
            "<Notebooks xmlns='urn:test'><Notebook ID='a' name='Shared' nickname='Field book' color='#2468AC'>" +
            "<SectionGroup><SectionGroup><Section ID='s1' name='Drawings' color='#FFFFFF'/></SectionGroup></SectionGroup></Notebook>" +
            "<Notebook ID='b' name='Shared' color='automatic'><Section ID='s2'/></Notebook></Notebooks>"));
        Check(sections["s1"].Notebook.Name == "Field book" && sections["s1"].Notebook.Color.ToArgb() == Color.FromArgb(36,104,172).ToArgb() &&
            sections["s1"].SectionName == "Drawings" && sections["s1"].SectionColor.ToArgb() == Color.White.ToArgb() &&
            sections["s2"].Notebook.Id == "b" && sections["s2"].SectionColor.IsEmpty, "Nested section and notebook metadata must stay distinct.");
        File.WriteAllText(path, "<workingSet version='1' revision='old'><tabs><page id='a' caption='First'/>" +
            "<page id='b' caption='Second'/></tabs><bookmarks><page id='a' caption='First'>" +
            "<notebook id='a' name='Field book' color='#2468AC'/></page></bookmarks></workingSet>");
        var store = new SessionStore(path);
        var legacy = store.Read();
        var first = legacy.Tabs[0];
        var second = legacy.Tabs[1];
        var oldBookmark = legacy.Bookmarks.Single();
        Check(oldBookmark.Notebook.Name == "Field book" && oldBookmark.Location.SectionName == null,
            "A bookmark saved before section labels were introduced lost its notebook.");
        store.Change(s => s.UpdateBookmarkOwners(new Dictionary<string, BookmarkLocation> { { "a", sections["s1"] } }));
        Check(store.Read().Bookmarks.Single().Location.SameAs(sections["s1"]), "Could not backfill a legacy bookmark's section.");
        store.Change(s => s.UpdateTabNotebooks(new Dictionary<string, BookmarkLocation> { { "a", sections["s1"] } }));
        store.Change(s => s.UpdateTab(first, new SavedPage("a", "Renamed / Field book", "")));
        Check(new SessionStore(path).Read().Tabs[0].Page.Notebook.SameAs(sections["s1"].Notebook),
            "A legacy tab lost its saved notebook ID, name or color when renamed and reopened offline.");
        Check(!legacy.Tabs.Any(t => t.Pinned), "A legacy tab unexpectedly migrated as pinned.");
        store.Change(s => s.SetPinned(first.Key, true));
        store.Change(s => s.UpdateTab(first, second.Page));
        var migrated = new SessionStore(path).Read();
        Check(migrated.Tabs.Select(t => t.Key).SequenceEqual(new[] { first.Key, second.Key }) &&
            migrated.Tabs.All(t => t.Page.Id == "b") && migrated.Tabs[0].Pinned && !migrated.Tabs[1].Pinned &&
            migrated.Bookmarks.Single().Id == "a",
            "Migration or restart merged tabs, reordered them, or retargeted the bookmark.");
        store.Change(s => s.UpdateTab(first, new SavedPage("stale", "Stale window", "")));
        Check(store.Read().Tabs[0].Page.Id == "b", "An outdated window overwrote a newer page.");
        store.Change(s => s.Tabs.RemoveAll(t => t.Key == first.Key) > 0);
        Check(store.Read().Tabs.Single().Key == second.Key, "Closing one duplicate closed the other tab.");
        File.WriteAllText(path, "<workingSet version='2' revision='old'><tabs><page tabId='slot' id='x' caption='X'/></tabs><bookmarks/></workingSet>");
        var v2 = store.Read().Tabs.Single();
        Check(v2.Key == "slot" && !v2.Pinned, "Version 2 lost the existing tab identity or invented a pin.");
        Console.WriteLine("PASS migrated saved tabs, independent same-page slots, persistent order, bookmarks and stale-window protection.");
    }

    private static void ExerciseEditorLifetime(string path)
    {
        for (int attachment = 0; attachment < 2; attachment++)
        {
            var source = new Pages();
            source.Navigate(attachment == 0 ? "a1" : "untracked");
            using (var frame = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual,
                Location = new Point(-16000, -16000), ClientSize = new Size(640, 360) })
            using (var ribbon = new Panel { Bounds = new Rectangle(0, 0, 640, 80) })
            using (var body = new Panel { Bounds = new Rectangle(0, 80, 640, 280) })
            using (var row = new Panel { Dock = DockStyle.Top, Height = 32 })
            {
                frame.Controls.Add(ribbon);
                frame.Controls.Add(body);
                body.Controls.Add(row);
                frame.Show();
                using (var layout = new NativeLayout(frame.Handle, body.Handle, ribbon.Handle, row.Handle))
                using (var strip = new PageStrip(frame.Handle, source, Application.ExitThread, e => { throw e; }, new SessionStore(path)))
                using (var closeHost = new Timer { Interval = 600 })
                {
                    strip.Attach(layout.Reserve(), layout);
                    strip.RestoreOrSeed();
                    Check(strip.TabCount == 1 && strip.IsSelected("a1") == (attachment == 0), "New editor lost the prior working set.");
                    int phase = 0;
                    bool hidden = false, modal = false, resumed = false;
                    if (attachment == 0) row.Hide();
                    closeHost.Tick += delegate
                    {
                        if (attachment == 0)
                        {
                            if (phase == 0)
                            {
                                hidden = !strip.Visible && body.Top == ribbon.Bottom;
                                frame.Enabled = false;
                                row.Show();
                                phase++;
                                return;
                            }
                            if (phase == 1)
                            {
                                modal = !strip.Visible;
                                frame.Enabled = true;
                                phase++;
                                return;
                            }
                            resumed = strip.Visible && layout.StillReserved(strip.Host.Bounds) && strip.IsSelected("a1");
                        }
                        closeHost.Stop();
                        frame.Close();
                    };
                    closeHost.Start();
                    Check(Program.RunInteractive(layout, strip, null), "Losing the host did not request reattachment.");
                    if (attachment == 0)
                        Check(hidden && modal && resumed, "The strip did not pause and resume in the same editor after a hidden row and modal dialog.");
                }
            }
        }
        Console.WriteLine("PASS resident loop released the row during a hidden editor and modal Options simulation, resumed the same tab, and reattached after closing the editor.");
    }

    internal static void ProcessPhase(Guid run, string phase)
    {
        var store = new SessionStore(Path.Combine(DirectoryFor(run), "fixture.xml"));
        var source = new Pages { reopened = phase != "save" };
        source.Navigate(phase == "save" ? "a1" : "untracked");
        using (var frame = new Form { ShowInTaskbar = false, StartPosition = FormStartPosition.Manual, Location = new Point(-16000, -16000) })
        using (var strip = new PageStrip(frame.Handle, source, delegate { }, e => { throw e; }, store))
        {
            frame.Show();
            strip.Attach(new Rectangle(0, 0, 650, HostChrome.StripHeight(96)));
            strip.RestoreOrSeed();
            switch (phase)
            {
                case "save":
                    source.Navigate("a2"); strip.ClickKeepCurrent();
                    AppearanceTests.TogglePin(strip, strip.Items.OfType<PageTabButton>().Single(t => t.Checked), true);
                    source.Navigate("b1"); strip.ClickKeepCurrent(); strip.ClickKeepCurrent();
                    var menu = strip.Items.OfType<BookmarkButton>().Single();
                    menu.ShowDropDown();
                    menu.DropDownItems[0].PerformClick();
                    menu.HideDropDown();
                    var b = store.Read().Bookmarks.Single();
                    Check(b.Id == "b1" && b.Link == PageLink, "Bookmark menu did not capture the displayed page.");
                    store.Change(s => { s.HomePage = b; s.NewNoteSectionId = "section-b"; s.NewNoteSectionName = "B & C / Projects"; return true; });
                    strip.CloseTab(store.Read().Tabs.Single(t => t.Page.Matches(b)));
                    Check(strip.TabCount == 2 && source.Current == "b1", "Closing a tab navigated or changed the wrong tab.");
                    ((ToolStripDropDownButton)strip.Items[1]).DropDownItems[0].PerformClick();
                    break;
                case "restore":
                    var stored = store.Read();
                    Check(stored.HomePage.Matches(stored.Bookmarks.Single()) && stored.NewNoteSectionId == "section-b" &&
                        stored.NewNoteSectionName == "B & C / Projects", "A restart lost the Tools destinations.");
                    var settings = ((ToolStripDropDownButton)strip.Items[1]).DropDownItems.Cast<ToolStripItem>().Take(3).Cast<ToolStripMenuItem>().ToArray();
                    Check(stored.ColorNotebookLabels && stored.ShowNotebookLabels && !stored.GroupByNotebook && !strip.WantsCaptionRow &&
                        settings.All(i => !i.Checked) && settings[0].Enabled && settings.Skip(1).All(i => !i.Enabled),
                        "A restart lost the dependent preferences or showed them as active with grouping off.");
                    var pinned = strip.Items.OfType<PageTabButton>().First();
                    Check(((SavedTab)pinned.Tag).Pinned && ((SavedTab)pinned.Tag).Page.Id == "a2" &&
                        strip.Items.OfType<ToolStripSeparator>().Count(s => s.Available) == 2, "Pinned group did not survive restart.");
                    Check(strip.TabCount == 2 && stored.Tabs.Select(t => t.Page.Id).SequenceEqual(new[] { "a1", "a2" }) &&
                        source.Current == "untracked" && !strip.IsSelected("a1"), "Restoration changed the page or lost the tab order.");
                    Check(strip.Items.OfType<PageTabButton>().All(t => !t.Checked && t.SectionColor == source.SectionColor(((SavedTab)t.Tag).Page.Id)),
                        "An inactive tab was restored without its section stripe.");
                    var saved = stored.Bookmarks.Single();
                    Check(saved.Caption == "Dashboard / B & C", "Bookmark caption did not round trip.");
                    var dropdown = strip.Items.OfType<BookmarkButton>().Single();
                    dropdown.ShowDropDown();
                    var offline = dropdown.DropDownItems.OfType<BookmarkHeader>().Single();
                    Check(offline.Text == "B & C" && offline.Notebook.Color.ToArgb() == Color.Teal.ToArgb(),
                        "An unavailable bookmark lost its notebook name or color after a restart.");
                    var offlineRow = dropdown.DropDownItems.OfType<BookmarkRow>().Single();
                    Check(offlineRow.SectionName == "Projects" && offlineRow.SectionColor.ToArgb() == Color.Indigo.ToArgb(),
                        "The closed notebook lost its saved section name or color.");
                    offlineRow.PerformClick();
                    dropdown.HideDropDown();
                    Check(source.Current == "b2" && strip.TabCount == 3 && strip.IsSelected("b1"), "Bookmark link did not restore the renamed ID.");
                    strip.ClickKeepCurrent();
                    Check(strip.TabCount == 3 && strip.IsSelected("b2"), "Changed runtime ID created a duplicate tab.");
                    var gone = new SavedPage("missing", "Unavailable page", "onenote:missing");
                    strip.SaveBookmark(gone);
                    Check(!strip.OpenSaved(gone) && strip.TabCount == 3 && store.Read().Bookmarks.Count == 2, "Failed navigation lost a bookmark.");
                    dropdown.ShowDropDown();
                    var remove = dropdown.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text == "Remove bookmark");
                    Check(remove.DropDownItems.OfType<BookmarkHeader>().Any(i => i.Text == "Unknown notebook"),
                        "An unresolvable bookmark should remain available in the removal menu.");
                    remove.DropDownItems.OfType<ToolStripMenuItem>().Single(i => i.Text == "Unavailable page").PerformClick();
                    dropdown.HideDropDown();
                    strip.RemoveBookmark(saved);
                    Check(strip.TabCount == 3 && store.Read().Bookmarks.Count == 0, "Removing a bookmark closed a tab.");
                    foreach (var tab in store.Read().Tabs) strip.CloseTab(tab);
                    break;
                case "empty":
                    Check(strip.TabCount == 0 && store.Read().Bookmarks.Count == 0, "An intentionally empty session was reseeded.");
                    break;
                default: throw new ArgumentException("Unknown session fixture phase.");
            }
        }
    }
}
