using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using One = Microsoft.Office.Interop.OneNote;

internal interface IPages
{
    string Current { get; }
    string Caption(string id);
    Color SectionColor(string pageId);
    Dictionary<string, BookmarkNotebook> BookmarkOwners(IEnumerable<string> pageIds);
    string CreateQuickNote();
    void Navigate(string id);
    string LinkFor(string id);
    void NavigateLink(string link);
}

internal sealed class NotebookPages : IPages
{
    internal readonly One.Window Window;
    internal readonly One.Application App;
    private readonly string quickNotesPath;
    internal NotebookPages(One.Application app, One.Window window, string quickNotesPath = null)
    {
        App = app;
        Window = window;
        this.quickNotesPath = quickNotesPath;
    }

    private static IntPtr HandleOf(One.Window window)
    {
        return window == null ? IntPtr.Zero : new IntPtr(unchecked((long)window.WindowHandle));
    }
    internal IntPtr Frame { get { return NativeLayout.RootOf(HandleOf(Window)); } }

    internal static NotebookPages Bind(List<NativeLayout> layouts, IntPtr requested, string quickNotesPath = null)
    {
        var app = new One.Application();
        var windows = app.Windows;
        IntPtr current = IntPtr.Zero;
        try { current = HandleOf(windows.CurrentWindow); }
        catch (COMException error) { Console.WriteLine("Current COM window unavailable (0x{0:X8}).", error.HResult); }
        var matches = new List<KeyValuePair<IntPtr, One.Window>>();
        Console.WriteLine("Current COM handle {0}, root {1}", current, NativeLayout.RootOf(current));
        uint count = windows.Count;
        for (uint i = 0; i < count; i++)
        {
            One.Window window = windows[i];
            IntPtr handle = HandleOf(window);
            IntPtr root = NativeLayout.RootOf(handle);
            Console.WriteLine("COM handle {0}, root {1}", handle, root);
            if (layouts.Exists(layout => layout.Frame == root))
                matches.Add(new KeyValuePair<IntPtr, One.Window>(root, window));
        }
        IntPtr frame = SelectFrame(matches.Select(pair => pair.Key).ToList(), NativeLayout.RootOf(current), requested);
        return new NotebookPages(app, matches.Single(pair => pair.Key == frame).Value, quickNotesPath);
    }

    internal static IntPtr SelectFrame(List<IntPtr> matchedFrames, IntPtr currentRoot, IntPtr requested)
    {
        if (matchedFrames.Count == 0)
            throw new InvalidOperationException("No COM window belongs to a measured editor. Compare the COM roots and measured frames above.");
        IntPtr selected = requested;
        if (selected == IntPtr.Zero)
        {
            if (matchedFrames.Contains(currentRoot)) selected = currentRoot;
            else if (matchedFrames.Count == 1) selected = matchedFrames[0];
            else throw new InvalidOperationException("Several COM windows matched and the current window did not identify one. Select a measured frame with -WindowHandle.");
        }
        if (matchedFrames.Count(frame => frame == selected) != 1)
            throw new InvalidOperationException("Frame " + selected + " has no unique COM window. Compare the COM roots and measured frames above.");
        return selected;
    }
    public string Current { get { return Window.CurrentPageId; } }
    public string CreateQuickNote()
    {
        string page;
        App.CreateNewPage(ResolveQuickNotesSection(quickNotesPath), out page, One.NewPageStyle.npsBlankPageWithTitle);
        return page;
    }
    internal string ResolveQuickNotesSection(string path)
    {
        if (path == null) App.GetSpecialLocation(One.SpecialLocation.slUnfiledNotesSection, out path);
        if (String.IsNullOrWhiteSpace(path)) throw new InvalidOperationException("OneNote has no configured Quick Notes location.");
        string xml, section;
        App.GetHierarchy("", One.HierarchyScope.hsSections, out xml, One.XMLSchema.xs2013);
        var nodes = XElement.Parse(xml).DescendantsAndSelf();
        var open = nodes.FirstOrDefault(n => n.Name.LocalName == "Section" && SameLocation((string)n.Attribute("path"), path));
        if (open != null && !String.IsNullOrEmpty((string)open.Attribute("ID"))) return (string)open.Attribute("ID");
        // An already-open section can still lack a backing file while OneNote syncs it.
        // Ask for it relative to its known notebook or section group in that case.
        int split = Math.Max(path.LastIndexOf('/'), path.LastIndexOf('\\'));
        var parent = split <= 0 ? null : nodes.FirstOrDefault(n => (n.Name.LocalName == "Notebook" || n.Name.LocalName == "SectionGroup") &&
            SameLocation((string)n.Attribute("path"), path.Substring(0, split)));
        if (parent != null && !String.IsNullOrEmpty((string)parent.Attribute("ID")))
            App.OpenHierarchy(Uri.UnescapeDataString(path.Substring(split + 1)), (string)parent.Attribute("ID"), out section, One.CreateFileType.cftNone);
        else App.OpenHierarchy(path, "", out section, One.CreateFileType.cftNone);
        if (String.IsNullOrEmpty(section)) throw new InvalidOperationException("OneNote did not resolve the Quick Notes section.");
        return section;
    }
    private static bool SameLocation(string left, string right)
    {
        if (String.IsNullOrEmpty(left) || String.IsNullOrEmpty(right)) return false;
        left = left.TrimEnd('\\', '/');
        right = right.TrimEnd('\\', '/');
        if (String.Equals(left, right, StringComparison.Ordinal)) return true;
        Uri a, b;
        if (!Uri.TryCreate(left, UriKind.Absolute, out a) || !Uri.TryCreate(right, UriKind.Absolute, out b)) return false;
        if (a.IsFile && b.IsFile) return String.Equals(a.LocalPath, b.LocalPath, StringComparison.OrdinalIgnoreCase);
        return (a.Scheme == Uri.UriSchemeHttp || a.Scheme == Uri.UriSchemeHttps) && a.Equals(b) && a.Fragment == b.Fragment;
    }
    public void Navigate(string id) { Window.NavigateTo(id, null); }
    public void NavigateLink(string link) { Window.NavigateToUrl(link); }
    public string LinkFor(string id)
    {
        string link;
        try { App.GetHyperlinkToObject(id, "", out link); return link; }
        catch (COMException) { return ""; }
    }
    private XElement HierarchyObject(string id)
    {
        string xml;
        App.GetHierarchy(id, One.HierarchyScope.hsSelf, out xml, One.XMLSchema.xs2013);
        var root = XDocument.Parse(xml).Root;
        return root.DescendantsAndSelf().FirstOrDefault(node => (string)node.Attribute("ID") == id);
    }
    internal string NameOf(string id)
    {
        var node = HierarchyObject(id);
        return node == null ? "" : (string)node.Attribute("name") ?? "";
    }
    public Color SectionColor(string pageId)
    {
        string section;
        App.GetHierarchyParent(pageId, out section);
        return String.IsNullOrEmpty(section) ? Color.Empty : ReadSectionColor(HierarchyObject(section));
    }
    internal static Color ReadSectionColor(XElement section)
    {
        if (section == null || section.Name.LocalName != "Section") return Color.Empty;
        return BookmarkNotebook.ReadColor((string)section.Attribute("color"));
    }
    public Dictionary<string, BookmarkNotebook> BookmarkOwners(IEnumerable<string> pageIds)
    {
        var result = new Dictionary<string, BookmarkNotebook>();
        string xml;
        try { App.GetHierarchy("", One.HierarchyScope.hsSections, out xml, One.XMLSchema.xs2013); }
        catch (COMException) { return result; }
        var sections = BookmarkNotebook.IndexSections(XElement.Parse(xml));
        foreach (string pageId in pageIds)
        {
            string section;
            try { App.GetHierarchyParent(pageId, out section); }
            catch (COMException) { continue; } // A closed notebook can still be shown using its saved heading.
            BookmarkNotebook owner;
            if (section != null && sections.TryGetValue(section, out owner)) result[pageId] = owner;
        }
        return result;
    }
    public string Caption(string id)
    {
        string title = NameOf(id), book = NameOf(Window.CurrentNotebookId);
        return (String.IsNullOrWhiteSpace(title) ? "Untitled" : title) +
            (String.IsNullOrWhiteSpace(book) ? "" : " / " + book);
    }
}

// An actual child of the OneNote frame. Only show it after the host makes room.
internal sealed class PageStrip : ToolStrip
{
    private readonly IntPtr frame;
    private readonly IPages pages;
    private readonly Dictionary<string, PageTabButton> buttons = new Dictionary<string, PageTabButton>();
    private readonly ToolStripSplitButton add;
    private readonly BookmarkButton bookmarks = new BookmarkButton();
    private readonly ToolStripSeparator bookmarkDivider = new ToolStripSeparator();
    private readonly ToolStripSeparator pinDivider = new ToolStripSeparator { Available = false };
    private readonly ContextMenuStrip tabMenu = new ContextMenuStrip();
    private readonly ToolTip notice = new ToolTip();
    private TabRenderer appearance;
    private readonly HostChrome defaultChrome = new HostChrome();
    private bool highContrast;
    private Font tabFont, selectedTabFont;
    private readonly SessionStore store;
    private WorkingSet state;
    private string displayedRevision;
    private SavedPage displayedPage;
    private string activeKey;
    private SavedTab pendingTab;
    private bool focusPending;
    private DateTime nextRead;
    private DateTime nextCaptionCheck;
    private DateTime navigationDeadline;
    private readonly Action<Exception> onError;
    private const int FocusShortcut = 1;
    private bool shortcutRegistered, shortcutAttempted;
    private IntPtr returnFocus;
    internal PageStrip(IntPtr frame, IPages pages, Action close, Action<Exception> onError, SessionStore store = null)
    {
        this.frame = frame;
        this.pages = pages;
        this.onError = onError;
        this.store = store ?? new SessionStore(null);
        AutoSize = false;
        GripStyle = ToolStripGripStyle.Hidden;
        Dock = DockStyle.None;
        TabStop = true;
        SetStyle(ControlStyles.Selectable, true);
        AccessibleName = "MoreTabs page tabs";
        AccessibleDescription = "Ctrl+Alt+F6 focuses the page tabs. Use the arrow keys, Enter to open, or Escape to return.";
        add = new ToolStripSplitButton("+") { AccessibleName = "New Quick Note", ToolTipText = "Create a new Quick Note", Overflow = ToolStripItemOverflow.Never };
        add.ButtonClick += delegate { Perform(CreateQuickNote); };
        add.DropDownItems.Add("Keep current page in a tab", null, delegate { Perform(AddCurrent); });
        bookmarks.DropDownItems.Add("Bookmark current page");
        bookmarks.DropDownOpening += delegate { Perform(BuildBookmarksMenu); };
        Items.AddRange(new ToolStripItem[] { bookmarks, bookmarkDivider, pinDivider, add });
        var stop = new ToolStripButton("Close prototype") { Alignment = ToolStripItemAlignment.Right };
        stop.Click += delegate { close(); };
        Items.Add(stop);
        defaultChrome.Update(IntPtr.Zero, 96);
        ApplyAppearance(defaultChrome);
        ApplyState(this.store.Read());
    }
    protected override CreateParams CreateParams
    {
        get { var cp = base.CreateParams; cp.Parent = frame; cp.Style |= 0x40000000; return cp; }
    }
    private void Perform(Action action) { try { action(); } catch (Exception e) { onError(e); } }
    protected override void Dispose(bool disposing)
    {
        if (disposing) { tabMenu.Dispose(); notice.Dispose(); }
        base.Dispose(disposing);
        if (disposing)
        {
            if (tabFont != null) tabFont.Dispose();
            if (selectedTabFont != null) selectedTabFont.Dispose();
        }
    }

    internal void UpdateFocusShortcut()
    {
        if (GetForegroundWindow() != frame)
        {
            ReleaseFocusShortcut();
            return;
        }
        if (shortcutAttempted) return;
        shortcutAttempted = true;
        const uint ctrlAltNoRepeat = 0x4003;
        shortcutRegistered = RegisterHotKey(Handle, FocusShortcut, ctrlAltNoRepeat, (uint)Keys.F6);
        Console.WriteLine(shortcutRegistered ? "Ctrl+Alt+F6 ready for this editor." :
            "Ctrl+Alt+F6 unavailable; Windows error " + Marshal.GetLastWin32Error());
    }

    private void ReleaseFocusShortcut()
    {
        if (shortcutRegistered) UnregisterHotKey(Handle, FocusShortcut);
        shortcutRegistered = false;
        shortcutAttempted = false;
    }
    internal void PauseForDialog()
    {
        ReleaseFocusShortcut();
        Hide();
    }
    protected override void OnHandleDestroyed(EventArgs e)
    {
        ReleaseFocusShortcut();
        base.OnHandleDestroyed(e);
    }
    protected override void WndProc(ref Message message)
    {
        const int wmHotkey = 0x0312;
        if (message.Msg == wmHotkey && message.WParam.ToInt32() == FocusShortcut)
        {
            if (GetForegroundWindow() == frame) Perform(delegate { FocusPageTabs(ForegroundFocus()); });
            return;
        }
        base.WndProc(ref message);
    }
    internal void FocusPageTabs(IntPtr previous)
    {
        var active = buttons.Values.FirstOrDefault(button => button.Checked);
        FocusTab((ToolStripItem)active ?? add, previous);
    }
    private bool FocusTab(ToolStripItem tab, IntPtr previous)
    {
        if (!ContainsFocus && previous != IntPtr.Zero && NativeLayout.RootOf(previous) == frame)
            returnFocus = previous;
        if (!Focus()) { Console.WriteLine("Could not focus the tab strip."); return false; }
        tab.Select();
        Console.WriteLine("Keyboard focus entered the tab strip.");
        return ContainsFocus && tab.Selected;
    }
    private void ReturnToEditor()
    {
        IntPtr target = returnFocus;
        returnFocus = IntPtr.Zero;
        // NavigateTo may already have given focus to the new page.
        if (!ContainsFocus && NativeLayout.RootOf(ForegroundFocus()) == frame) return;
        if (target == IntPtr.Zero || NativeLayout.RootOf(target) != frame) target = frame;
        SetFocus(target);
        Console.WriteLine(ContainsFocus ? "Keyboard focus is still on the strip." : "Keyboard focus left the strip.");
    }
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (!ContainsFocus) return base.ProcessCmdKey(ref msg, keyData);
        var choices = Items.Cast<ToolStripItem>().Where(item => (item is ToolStripButton || item == add || item == bookmarks) && item.Available && item.Enabled).ToList();
        int index = choices.FindIndex(item => item.Selected);
        if (choices.Count > 0 && (keyData == Keys.Left || keyData == Keys.Right))
        {
            int delta = keyData == Keys.Left ? -1 : 1;
            choices[(Math.Max(index, 0) + delta + choices.Count) % choices.Count].Select();
            return true;
        }
        if ((keyData == Keys.Enter || keyData == Keys.Space) && index >= 0)
        {
            if (choices[index] == add) add.PerformButtonClick();
            else if (choices[index] == bookmarks) bookmarks.ShowDropDown();
            else { choices[index].PerformClick(); ReturnToEditor(); }
            return true;
        }
        if (keyData == Keys.Escape || keyData == Keys.F6 || keyData == Keys.Tab || keyData == (Keys.Shift | Keys.Tab))
        {
            ReturnToEditor();
            return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [StructLayout(LayoutKind.Sequential)]
    private struct GuiThreadInfo
    {
        public uint Size, Flags;
        public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public int Left, Top, Right, Bottom;
    }
    private static IntPtr ForegroundFocus()
    {
        var info = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
        return GetGUIThreadInfo(0, ref info) ? info.Focus : IntPtr.Zero;
    }
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo info);
    [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
    [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr window, int id);

    internal void Attach(Rectangle slot, NativeLayout layout = null)
    {
        Bounds = slot;
        if (layout != null) ApplyAppearance(layout.Appearance(!Visible));
        if (NativeLayout.GetParent(Handle) != frame || NativeLayout.Bounds(Handle, frame) != slot)
            throw new InvalidOperationException("The strip did not attach within the reserved row.");
        Show();
    }
    internal void ApplyAppearance(HostChrome chrome)
    {
        bool metrics = appearance == null || appearance.Dpi != chrome.Dpi;
        if (metrics || appearance.Surface.ToArgb() != chrome.Surface.ToArgb() || highContrast != chrome.HighContrast)
        {
            highContrast = chrome.HighContrast;
            appearance = new TabRenderer(chrome.Surface, chrome.Dpi, highContrast);
            Renderer = appearance;
            tabMenu.Renderer = bookmarks.DropDown.Renderer = add.DropDown.Renderer = OverflowButton.DropDown.Renderer = appearance;
            BackColor = appearance.Surface;
            ForeColor = appearance.Ink;
        }
        // The notebook inset anchors the visible glyph. Keep some clickable space to its left.
        int iconPadding = Math.Min(HostChrome.Px(8, chrome.Dpi), chrome.Inset);
        var inset = new Padding(chrome.Inset - iconPadding, 0, HostChrome.Px(2, chrome.Dpi), 0);
        if (Padding != inset) Padding = inset;
        if (bookmarks.Padding.Left != iconPadding) bookmarks.Padding = new Padding(iconPadding, 0, 0, 0);
        int bookmarkWidth = Math.Max(HostChrome.Px(24, chrome.Dpi), HostChrome.Px(17, chrome.Dpi) + 2 * iconPadding);
        if (bookmarks.TabWidth != bookmarkWidth) { bookmarks.TabWidth = bookmarkWidth; PerformLayout(); }
        if (metrics)
        {
            var tabMargin = TabMargin(chrome.Dpi);
            bookmarks.Margin = new Padding(0, tabMargin.Top, tabMargin.Right, tabMargin.Bottom);
            Font oldRegular = tabFont, oldSelected = selectedTabFont;
            var menuFont = SystemFonts.MenuFont;
            tabFont = new Font(menuFont.FontFamily, menuFont.SizeInPoints * chrome.Dpi / 72f, FontStyle.Regular, GraphicsUnit.Pixel);
            selectedTabFont = new Font(tabFont, FontStyle.Bold);
            Font = tabMenu.Font = tabFont;
            var commandPadding = new Padding(HostChrome.Px(5, chrome.Dpi), 0, HostChrome.Px(5, chrome.Dpi), 0);
            foreach (ToolStripItem item in Items)
                if (!(item is PageTabButton) && !(item is ToolStripSeparator) && item != bookmarks) item.Padding = commandPadding;
            foreach (var divider in new[] { bookmarkDivider, pinDivider })
            {
                divider.AutoSize = false;
                divider.Margin = Padding.Empty;
                // One painted pixel with the same number of empty pixels on either side.
                divider.Size = new Size(2 * HostChrome.Px(4, chrome.Dpi) + 1, HostChrome.Px(30, chrome.Dpi));
                divider.Overflow = ToolStripItemOverflow.Never;
            }
            foreach (var tab in Items.OfType<PageTabButton>()) SizeTab(tab);
            if (oldRegular != null) oldRegular.Dispose();
            if (oldSelected != null) oldSelected.Dispose();
        }
    }
    private static Padding TabMargin(uint dpi)
    {
        int side = HostChrome.Px(1, dpi);
        return new Padding(side, 0, side, HostChrome.Px(3, dpi));
    }
    private void SizeTab(PageTabButton tab)
    {
        uint dpi = appearance.Dpi;
        tab.Dpi = dpi;
        // Measure both states at the heavier weight so selecting a tab never shifts its neighbors.
        tab.Font = selectedTabFont;
        tab.Margin = TabMargin(dpi);
        tab.Padding = new Padding(HostChrome.Px(10, dpi), 0, HostChrome.Px(10, dpi), 0);
    }
    private static string ShortCaption(SavedPage page)
    {
        const string separator = " / ";
        int split = page.Caption.IndexOf(separator, StringComparison.Ordinal);
        return split > 0 && split == page.Caption.LastIndexOf(separator, StringComparison.Ordinal) ? page.Caption.Substring(0, split) : page.Caption;
    }
    internal void AddCurrent()
    {
        SavedPage page = CaptureCurrent();
        if (page != null) ExpectNavigation(KeepTab(page), false);
        RefreshActive();
    }
    private SavedTab KeepTab(SavedPage page)
    {
        string key = null;
        ApplyState(store.Change(s => s.KeepTab(page, activeKey, out key)));
        return state.Tabs.Find(t => t.Key == key);
    }
    private void ExpectNavigation(SavedTab tab, bool focus)
    {
        pendingTab = tab;
        focusPending = focus;
        navigationDeadline = DateTime.UtcNow.AddSeconds(5);
        nextCaptionCheck = DateTime.MinValue;
    }
    private void CreateQuickNote()
    {
        pendingTab = null;
        focusPending = false;
        string id;
        try { id = pages.CreateQuickNote(); }
        catch (Exception error)
        {
            if (!(error is COMException || error is InvalidOperationException)) throw;
            Console.WriteLine("Quick Note creation failed ({0}, 0x{1:X8}).", error.GetType().Name, error.HResult);
            notice.Show("OneNote could not create a Quick Note. Check the Quick Notes location and write access.", this, 4, Height, 6000);
            return;
        }
        var page = new SavedPage(id, "Quick Note", pages.LinkFor(id));
        SavedTab saved = null;
        try { saved = KeepTab(page); }
        catch (System.IO.IOException) { }
        catch (UnauthorizedAccessException) { }
        try { pages.Navigate(id); }
        catch (COMException error)
        {
            Console.WriteLine("Opening the newly created Quick Note failed (0x{0:X8}).", error.HResult);
            notice.Show(saved != null ? "Quick Note created. Click its tab to try opening it again." :
                "Quick Note created in OneNote. MoreTabs could not save or open its tab.", this, 4, Height, 6000);
            return;
        }
        if (saved != null) ExpectNavigation(saved, true);
        else { activeKey = null; notice.Show("Quick Note opened. MoreTabs could not save its tab.", this, 4, Height, 6000); }
        RefreshActive();
    }
    internal string ActiveTabKey { get { return activeKey; } }
    internal void RestoreOrSeed(string preferredKey = null)
    {
        activeKey = preferredKey;
        if (state.Revision == null)
        {
            SavedPage page = CaptureCurrent();
            if (page != null) ApplyState(store.Change(s => { string key; return s.Revision == null && s.KeepTab(page, null, out key); }));
        }
        RefreshActive();
        Console.WriteLine("Working set attached: {0} tab(s), {1} bookmark(s).", state.Tabs.Count, state.Bookmarks.Count);
    }
    private SavedPage CaptureCurrent()
    {
        string id = pages.Current;
        if (String.IsNullOrEmpty(id)) return null;
        var captured = new SavedPage(id, pages.Caption(id), pages.LinkFor(id));
        return pages.Current == id ? captured : null;
    }
    private void RefreshCaption(SavedTab active, string id)
    {
        if (DateTime.UtcNow >= nextCaptionCheck)
        {
            nextCaptionCheck = DateTime.UtcNow.AddMilliseconds(750);
            try
            {
                var caption = pages.Caption(id);
                if (state.Tabs.Exists(t => t.Page.Matches(active.Page) && t.Page.Caption != caption) && pages.Current == id)
                {
                    ApplyState(store.Change(s => {
                        bool changed = false;
                        foreach (var tab in s.Tabs.FindAll(t => t.Page.Matches(active.Page)))
                            changed |= s.UpdateTab(tab, new SavedPage(tab.Page.Id, caption, tab.Page.Link));
                        return changed;
                    }));
                }
            }
            catch (COMException) { /* Retry on the next tick if OneNote is busy with an edit. */ }
        }
    }
    private void ApplyState(WorkingSet updated)
    {
        state = updated;
        if (displayedRevision == updated.Revision) return;
        foreach (var id in buttons.Keys.Where(id => !state.Tabs.Any(t => t.Key == id)).ToArray())
        {
            if (id == activeKey) activeKey = null;
            var obsolete = buttons[id];
            buttons.Remove(id);
            Items.Remove(obsolete);
            obsolete.Dispose();
        }
        var titles = state.Tabs.ToLookup(t => ShortCaption(t.Page), StringComparer.CurrentCultureIgnoreCase);
        foreach (var slot in state.Tabs)
        {
            SavedPage page = slot.Page;
            PageTabButton button;
            if (!buttons.TryGetValue(slot.Key, out button))
            {
                var pageTab = new PageTabButton { DisplayStyle = ToolStripItemDisplayStyle.Text };
                SizeTab(pageTab);
                button = pageTab;
                var tab = button;
                tab.Click += delegate { Perform(delegate { var target = (SavedTab)tab.Tag; OpenSaved(target.Page, target); }); };
                tab.MouseUp += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button == MouseButtons.Right) TabMenuFor((SavedTab)tab.Tag).Show(Cursor.Position);
                    if (e.Button == MouseButtons.Middle) Perform(delegate { CloseTab((SavedTab)tab.Tag); });
                };
                buttons.Add(slot.Key, button);
            }
            else if (((SavedTab)button.Tag).Page.Id != page.Id)
            {
                button.SectionColor = Color.Empty;
                button.NextColorCheck = DateTime.MinValue;
                button.Invalidate();
            }
            string title = ShortCaption(page);
            button.Text = titles[title].Any(t => t.Page.Caption != page.Caption) ? page.Caption : title;
            button.AccessibleName = page.Caption;
            button.ToolTipText = page.Caption + (slot.Pinned ? "\nPinned tab" : "") + "\nRight-click for tab options.";
            button.Tag = slot;
        }
        int position = 2;
        foreach (var tab in state.Tabs.Where(t => t.Pinned)) PlaceItem(buttons[tab.Key], position++);
        pinDivider.Available = position > 2;
        PlaceItem(pinDivider, position++);
        foreach (var tab in state.Tabs.Where(t => !t.Pinned)) PlaceItem(buttons[tab.Key], position++);
        displayedRevision = updated.Revision;
    }
    private void PlaceItem(ToolStripItem item, int index)
    {
        if (Items.IndexOf(item) == index) return;
        Items.Remove(item);
        Items.Insert(index, item);
    }
    internal ContextMenuStrip TabMenuFor(SavedTab tab)
    {
        ClearMenu(tabMenu.Items);
        tabMenu.Items.Add(tab.Pinned ? "Unpin tab" : "Pin tab", null, delegate { Perform(delegate { SetPinned(tab.Key, !tab.Pinned); }); });
        tabMenu.Items.Add("Bookmark this tab", null, delegate { Perform(delegate { SaveBookmark(tab.Page); }); });
        tabMenu.Items.Add("Close tab", null, delegate { Perform(delegate { CloseTab(tab); }); });
        return tabMenu;
    }
    private void SetPinned(string key, bool pinned)
    {
        ApplyState(store.Change(s => s.SetPinned(key, pinned)));
        RefreshActive();
    }
    private static void ClearMenu(ToolStripItemCollection items)
    {
        while (items.Count > 0) { var item = items[0]; items.RemoveAt(0); item.Dispose(); }
    }
    internal void CloseTab(SavedTab tab)
    {
        if (pendingTab != null && pendingTab.Key == tab.Key) pendingTab = null;
        ApplyState(store.Change(s => s.Tabs.RemoveAll(t => t.Key == tab.Key) > 0));
        RefreshActive();
    }
    internal void SaveBookmark(SavedPage page)
    {
        if (page == null) return;
        BookmarkNotebook owner;
        if (pages.BookmarkOwners(new[] { page.Id }).TryGetValue(page.Id, out owner))
            page = new SavedPage(page.Id, page.Caption, page.Link, owner);
        ApplyState(store.Change(s => WorkingSet.Keep(s.Bookmarks, page)));
    }
    internal void RemoveBookmark(SavedPage page)
    {
        ApplyState(store.Change(s => s.Bookmarks.RemoveAll(page.Matches) > 0));
    }
    private void BuildBookmarksMenu()
    {
        ApplyState(store.Read());
        var menu = bookmarks.DropDownItems;
        ClearMenu(menu);
        menu.Add("Bookmark current page", null, delegate { Perform(delegate { SaveBookmark(CaptureCurrent()); }); });
        if (state.Bookmarks.Count == 0) return;
        var owners = pages.BookmarkOwners(state.Bookmarks.Select(page => page.Id));
        if (state.Bookmarks.Any(p => owners.ContainsKey(p.Id) && !owners[p.Id].SameAs(p.Notebook)))
        {
            // The menu can still use live colors if the optional offline cache cannot be written.
            try { ApplyState(store.Change(s => s.UpdateBookmarkOwners(owners))); }
            catch (System.IO.IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        var remove = new ToolStripMenuItem("Remove bookmark");
        remove.DropDown.Renderer = appearance;
        var booksById = owners.Values.Concat(state.Bookmarks.Select(p => p.Notebook).Where(n => n != null))
            .GroupBy(n => n.Id).ToDictionary(g => g.Key, g => g.First());
        var known = booksById.Values.ToLookup(n => n.Name, StringComparer.CurrentCultureIgnoreCase);
        var groups = state.Bookmarks.Select(page => {
            BookmarkNotebook book;
            if (!owners.TryGetValue(page.Id, out book))
                book = page.Notebook != null ? booksById[page.Notebook.Id] : page.NotebookForMenu();
            if (book.Id.Length == 0 && known[book.Name].Count() == 1) book = known[book.Name].First();
            return new { Page = page, Book = book };
        }).GroupBy(item => item.Book.Id.Length == 0 ? "name:" + item.Book.Name.ToUpperInvariant() : "id:" + item.Book.Id)
          .OrderBy(group => group.First().Book.Name, StringComparer.CurrentCultureIgnoreCase);
        foreach (var group in groups)
        {
            var book = group.First().Book;
            menu.Add(new BookmarkHeader(book));
            remove.DropDownItems.Add(new BookmarkHeader(book));
            foreach (var entry in group)
            {
                var page = entry.Page;
                string label = page.BookmarkTitle.Replace("&", "&&");
                menu.Add(label, null, delegate { Perform(delegate { OpenSaved(page); }); });
                remove.DropDownItems.Add(label, null, delegate { Perform(delegate { RemoveBookmark(page); }); });
            }
        }
        menu.Add(new ToolStripSeparator());
        menu.Add(remove);
    }
    internal bool OpenSaved(SavedPage page, SavedTab target = null)
    {
        pendingTab = null;
        focusPending = false;
        try { pages.Navigate(page.Id); }
        catch (COMException)
        {
            bool opened = false;
            if (page.Link.Length > 0)
            {
                try { pages.NavigateLink(page.Link); opened = true; }
                catch (COMException) { /* Keep the reference so the user can retry or remove it. */ }
            }
            if (!opened)
            {
                notice.Show("OneNote could not open this page. Open its notebook and try again.", this, 4, Height, 4500);
                return false;
            }
        }
        ExpectNavigation(target ?? KeepTab(page), false);
        RefreshActive();
        return true;
    }
    internal void RefreshActive()
    {
        if (DateTime.UtcNow >= nextRead)
        {
            ApplyState(store.Read());
            nextRead = DateTime.UtcNow.AddSeconds(2);
        }
        string id = pages.Current;
        if (String.IsNullOrEmpty(id))
        {
            // OneNote can clear the displayed page briefly. Saved tabs still have their own page IDs.
            RefreshSectionColors(null, false);
            return;
        }
        bool moved = displayedPage == null || displayedPage.Id != id;
        var visible = moved ? new SavedPage(id, "", pages.LinkFor(id)) : displayedPage;
        var activeTab = state.Tabs.Find(t => t.Key == activeKey);
        if (pendingTab != null)
        {
            // A tab click may return before OneNote switches pages. Never copy the outgoing page into the requested tab.
            if (!pendingTab.Page.Matches(visible) && DateTime.UtcNow < navigationDeadline) return;
            activeTab = state.Tabs.Find(t => t.Key == pendingTab.Key && t.Page.Matches(visible));
            if (activeTab == null) focusPending = false;
            pendingTab = null;
        }
        else if (moved && activeTab != null && activeTab.Page.Matches(displayedPage) && !activeTab.Page.Matches(visible))
        {
            focusPending = false;
            try
            {
                var next = new SavedPage(id, pages.Caption(id), visible.Link);
                if (pages.Current != id) return;
                ApplyState(store.Change(s => s.UpdateTab(activeTab, next)));
                activeTab = state.Tabs.Find(t => t.Key == activeKey);
            }
            catch (COMException) { return; } // Retry the same transition once OneNote finishes loading the page.
        }
        if (activeTab == null || !activeTab.Page.Matches(visible))
            activeTab = state.Tabs.Find(t => t.Page.Matches(visible));
        string selectedKey = activeTab == null ? null : activeTab.Key;
        bool selectionChanged = moved || selectedKey != activeKey;
        activeKey = selectedKey;
        displayedPage = visible;
        if (moved) nextCaptionCheck = DateTime.MinValue;
        if (activeTab != null) RefreshCaption(activeTab, id);
        RefreshSectionColors(id, selectionChanged);
        foreach (var entry in buttons)
        {
            bool active = entry.Key == activeKey;
            var overflow = active ? ToolStripItemOverflow.Never : ToolStripItemOverflow.AsNeeded;
            if (entry.Value.Overflow != overflow) entry.Value.Overflow = overflow;
            if (entry.Value.Checked != active)
            {
                entry.Value.Checked = active;
                entry.Value.AccessibleDescription = active ? "Active page tab" : "Page tab";
                entry.Value.Invalidate();
            }
        }
        if (focusPending && (activeKey == null || DateTime.UtcNow > navigationDeadline || FocusTab(buttons[activeKey], ForegroundFocus())))
            focusPending = false;
    }
    private void RefreshSectionColors(string currentId, bool selectionChanged)
    {
        DateTime now = DateTime.UtcNow;
        PageTabButton active = activeKey == null ? null : buttons[activeKey];
        if (active != null && selectionChanged) active.NextColorCheck = DateTime.MinValue;
        // Prioritize the displayed page, then service the oldest entries. Bound COM work per tick on large restores.
        if (!buttons.Values.Any(button => button.NextColorCheck <= now)) return;
        var byPage = buttons.Values.ToLookup(button => ((SavedTab)button.Tag).Page.Id);
        var due = buttons.Values.Where(button => button.NextColorCheck <= now)
            .OrderByDescending(button => button == active).ThenBy(button => button.NextColorCheck)
            .Select(button => ((SavedTab)button.Tag).Page.Id).Distinct().Take(3).ToArray();
        foreach (string pageId in due)
        {
            var tabs = byPage[pageId];
            Color color;
            try { color = pages.SectionColor(currentId != null && tabs.Contains(active) ? currentId : pageId); }
            catch (COMException) { color = Color.Empty; }
            foreach (var tab in tabs)
            {
                tab.NextColorCheck = now.AddSeconds(3);
                if (tab.SectionColor == color) continue;
                tab.SectionColor = color;
                tab.Invalidate();
            }
        }
    }
    internal void ClickAdd() { add.PerformButtonClick(); }
    internal void ClickKeepCurrent() { add.DropDownItems[0].PerformClick(); }
    internal void ClickTab(string id) { buttons[state.Tabs.First(t => t.Page.Id == id).Key].PerformClick(); }
    internal bool IsSelected(string id) { return buttons.Values.Any(b => b.Checked && ((SavedTab)b.Tag).Page.Id == id); }
    internal bool HasTabFocus(string id) { return ContainsFocus && IsSelected(id) && buttons[activeKey].Selected; }
    internal int TabCount { get { return buttons.Count; } }
}
