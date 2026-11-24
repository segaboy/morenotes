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
    void Navigate(string id);
    string LinkFor(string id);
    void NavigateLink(string link);
}

internal sealed class NotebookPages : IPages
{
    internal readonly One.Window Window;
    internal readonly One.Application App;
    internal NotebookPages(One.Application app, One.Window window)
    {
        App = app;
        Window = window;
    }

    private static IntPtr HandleOf(One.Window window)
    {
        return window == null ? IntPtr.Zero : new IntPtr(unchecked((long)window.WindowHandle));
    }
    internal IntPtr Frame { get { return NativeLayout.RootOf(HandleOf(Window)); } }

    internal static NotebookPages Bind(List<NativeLayout> layouts, IntPtr requested)
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
        return new NotebookPages(app, matches.Single(pair => pair.Key == frame).Value);
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
    public void Navigate(string id) { Window.NavigateTo(id, null); }
    public void NavigateLink(string link) { Window.NavigateToUrl(link); }
    public string LinkFor(string id)
    {
        string link;
        try { App.GetHyperlinkToObject(id, "", out link); return link; }
        catch (COMException) { return ""; }
    }
    internal string NameOf(string id)
    {
        string xml;
        App.GetHierarchy(id, One.HierarchyScope.hsSelf, out xml, One.XMLSchema.xs2013);
        var root = XDocument.Parse(xml).Root;
        foreach (var node in root.DescendantsAndSelf())
            if ((string)node.Attribute("ID") == id) return (string)node.Attribute("name") ?? "";
        return "";
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
    private readonly Dictionary<string, ToolStripButton> buttons = new Dictionary<string, ToolStripButton>();
    private readonly ToolStripButton add;
    private readonly ToolStripDropDownButton bookmarks = new ToolStripDropDownButton("Bookmarks");
    private readonly ContextMenuStrip tabMenu = new ContextMenuStrip();
    private readonly ToolTip notice = new ToolTip();
    private readonly SessionStore store;
    private WorkingSet state;
    private string displayedRevision;
    private SavedPage displayedPage;
    private DateTime nextRead;
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
        add = new ToolStripButton("+") { AccessibleName = "Add current page", ToolTipText = "Keep the displayed page in a tab" };
        add.Click += delegate { Perform(AddCurrent); };
        Items.Add(add);
        bookmarks.DropDownItems.Add("Bookmark current page");
        bookmarks.DropDownOpening += delegate { Perform(BuildBookmarksMenu); };
        Items.Add(bookmarks);
        var stop = new ToolStripButton("Close prototype") { Alignment = ToolStripItemAlignment.Right };
        stop.Click += delegate { close(); };
        Items.Add(stop);
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
        if (!ContainsFocus && previous != IntPtr.Zero && NativeLayout.RootOf(previous) == frame)
            returnFocus = previous;
        if (!Focus()) { Console.WriteLine("Could not focus the tab strip."); return; }
        var active = buttons.Values.FirstOrDefault(button => button.Checked);
        (active ?? add).Select();
        Console.WriteLine("Keyboard focus entered the tab strip.");
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
        var choices = Items.OfType<ToolStripButton>().Where(item => item.Available && item.Enabled).ToList();
        int index = choices.FindIndex(item => item.Selected);
        if (choices.Count > 0 && (keyData == Keys.Left || keyData == Keys.Right))
        {
            int delta = keyData == Keys.Left ? -1 : 1;
            choices[(Math.Max(index, 0) + delta + choices.Count) % choices.Count].Select();
            return true;
        }
        if ((keyData == Keys.Enter || keyData == Keys.Space) && index >= 0)
        {
            choices[index].PerformClick();
            ReturnToEditor();
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

    internal void Attach(Rectangle slot)
    {
        Bounds = slot;
        if (NativeLayout.GetParent(Handle) != frame || NativeLayout.Bounds(Handle, frame) != slot)
            throw new InvalidOperationException("The strip did not attach within the reserved row.");
        Show();
    }
    internal void AddCurrent()
    {
        SavedPage page = CaptureCurrent();
        if (page != null) ApplyState(store.Change(s => WorkingSet.Keep(s.Tabs, page)));
        RefreshActive();
    }
    internal void RestoreOrSeed()
    {
        if (state.Revision == null)
        {
            SavedPage page = CaptureCurrent();
            if (page != null) ApplyState(store.Change(s => s.Revision == null && WorkingSet.Keep(s.Tabs, page)));
        }
        RefreshActive();
        Console.WriteLine("Working set attached: {0} tab(s), {1} bookmark(s).", state.Tabs.Count, state.Bookmarks.Count);
    }
    private SavedPage CaptureCurrent()
    {
        string id = pages.Current;
        return String.IsNullOrEmpty(id) ? null : new SavedPage(id, pages.Caption(id), pages.LinkFor(id));
    }
    private void ApplyState(WorkingSet updated)
    {
        state = updated;
        if (displayedRevision == updated.Revision) return;
        displayedPage = null;
        foreach (var id in buttons.Keys.Where(id => !state.Tabs.Any(p => p.Id == id)).ToArray())
        {
            var obsolete = buttons[id];
            buttons.Remove(id);
            Items.Remove(obsolete);
            obsolete.Dispose();
        }
        for (int i = 0; i < state.Tabs.Count; i++)
        {
            SavedPage page = state.Tabs[i];
            ToolStripButton button;
            if (!buttons.TryGetValue(page.Id, out button))
            {
                button = new ToolStripButton { DisplayStyle = ToolStripItemDisplayStyle.Text };
                var tab = button;
                tab.Click += delegate { Perform(delegate { OpenSaved((SavedPage)tab.Tag, false); }); };
                tab.MouseUp += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button == MouseButtons.Right) ShowTabMenu((SavedPage)tab.Tag);
                    if (e.Button == MouseButtons.Middle) Perform(delegate { CloseTab((SavedPage)tab.Tag); });
                };
                buttons.Add(page.Id, button);
            }
            button.Text = page.Caption;
            button.ToolTipText = page.Caption + "\nRight-click to bookmark or close this tab.";
            button.Tag = page;
            if (Items.IndexOf(button) != i) { Items.Remove(button); Items.Insert(i, button); }
        }
        displayedRevision = updated.Revision;
    }
    private void ShowTabMenu(SavedPage page)
    {
        ClearMenu(tabMenu.Items);
        tabMenu.Items.Add("Bookmark this tab", null, delegate { Perform(delegate { SaveBookmark(page); }); });
        tabMenu.Items.Add("Close tab", null, delegate { Perform(delegate { CloseTab(page); }); });
        tabMenu.Show(Cursor.Position);
    }
    private static void ClearMenu(ToolStripItemCollection items)
    {
        while (items.Count > 0) { var item = items[0]; items.RemoveAt(0); item.Dispose(); }
    }
    internal void CloseTab(SavedPage page)
    {
        ApplyState(store.Change(s => s.Tabs.RemoveAll(page.Matches) > 0));
        RefreshActive();
    }
    internal void SaveBookmark(SavedPage page)
    {
        if (page != null) ApplyState(store.Change(s => WorkingSet.Keep(s.Bookmarks, page)));
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
        menu.Add(new ToolStripSeparator());
        var remove = new ToolStripMenuItem("Remove bookmark");
        foreach (var saved in state.Bookmarks)
        {
            var page = saved;
            string label = page.Caption.Replace("&", "&&");
            menu.Add(label, null, delegate { Perform(delegate { OpenSaved(page, true); }); });
            remove.DropDownItems.Add(label, null, delegate { Perform(delegate { RemoveBookmark(page); }); });
        }
        menu.Add(new ToolStripSeparator());
        menu.Add(remove);
    }
    internal bool OpenSaved(SavedPage page, bool addTab)
    {
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
        if (addTab) ApplyState(store.Change(s => WorkingSet.Keep(s.Tabs, page)));
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
        if (displayedPage == null || displayedPage.Id != id)
        {
            string link = !String.IsNullOrEmpty(id) && !buttons.ContainsKey(id) && state.Tabs.Count > 0 ? pages.LinkFor(id) : "";
            displayedPage = new SavedPage(id, "", link);
        }
        foreach (var entry in buttons)
        {
            bool active = ((SavedPage)entry.Value.Tag).Matches(displayedPage);
            if (entry.Value.Checked != active) entry.Value.Checked = active;
        }
    }
    internal void ClickAdd() { add.PerformClick(); }
    internal void ClickTab(string id) { buttons[id].PerformClick(); }
    internal bool IsSelected(string id) { return buttons[id].Checked; }
    internal int TabCount { get { return buttons.Count; } }
}
