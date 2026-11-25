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
    string SelectedPage(string sourceId);
    IEnumerable<XElement> PickerChildren(string parentId);
    string Caption(string id);
    Dictionary<string, BookmarkLocation> PageLocations(IEnumerable<string> pageIds, bool refresh = false);
    string CreateNote(string sectionId);
    void Navigate(string id);
    string LinkFor(string id);
    void NavigateLink(string link);
}

internal sealed class NotebookPages : IPages
{
    internal readonly One.Window Window;
    internal readonly One.Application App;
    private readonly string quickNotesPath;
    private Dictionary<string, BookmarkLocation> sections;
    private DateTime nextHierarchyRead;
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
    public IEnumerable<XElement> PickerChildren(string parentId)
    {
        string xml;
        App.GetHierarchy(parentId, parentId.Length == 0 ? One.HierarchyScope.hsNotebooks : One.HierarchyScope.hsChildren,
            out xml, One.XMLSchema.xs2013);
        if (String.IsNullOrWhiteSpace(xml)) return Enumerable.Empty<XElement>();
        var root = XElement.Parse(xml);
        if (parentId.Length == 0) return root.DescendantsAndSelf(root.Name.Namespace + "Notebook");
        var parent = root.DescendantsAndSelf().FirstOrDefault(n => (string)n.Attribute("ID") == parentId);
        return parent == null ? Enumerable.Empty<XElement>() : parent.Elements();
    }
    public string SelectedPage(string sourceId)
    {
        string section = Window.CurrentSectionId, xml;
        if (String.IsNullOrEmpty(section)) return null;
        App.GetHierarchy(section, One.HierarchyScope.hsPages, out xml, One.XMLSchema.xs2013);
        if (String.IsNullOrWhiteSpace(xml)) return null;
        var root = XElement.Parse(xml);
        var node = root.DescendantsAndSelf(root.Name.Namespace + "Section").FirstOrDefault(s => (string)s.Attribute("ID") == section);
        string current = Current;
        return node != null && Window.CurrentSectionId == section ? MenuSelection(node, current, sourceId) : null;
    }
    internal static string MenuSelection(XElement section, string current, string sourceId)
    {
        string id = PageMenuCommand.SelectedPage(section, current);
        if (id != null || String.IsNullOrEmpty(sourceId) || current == sourceId) return id;
        var rows = section.Elements(section.Name.Namespace + "Page");
        // Some hierarchies omit selection flags after the menu closes. The invoking window
        // must have moved to a page in this section and there must be no conflicting selection.
        return rows.Any(p => (string)p.Attribute("ID") == current) &&
            !rows.Any(p => (string)p.Attribute("selected") == "all" || (string)p.Attribute("selected") == "partial" ||
                (string)p.Attribute("isCurrentlyViewed") == "true") ? current : null;
    }
    public string CreateNote(string sectionId)
    {
        string page;
        App.CreateNewPage(sectionId ?? ResolveQuickNotesSection(quickNotesPath), out page, One.NewPageStyle.npsBlankPageWithTitle);
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
    public Dictionary<string, BookmarkLocation> PageLocations(IEnumerable<string> pageIds, bool refresh = false)
    {
        var result = new Dictionary<string, BookmarkLocation>();
        if (refresh || sections == null || DateTime.UtcNow >= nextHierarchyRead)
        {
            string xml;
            try { App.GetHierarchy("", One.HierarchyScope.hsSections, out xml, One.XMLSchema.xs2013); }
            catch (COMException) { return result; }
            sections = BookmarkLocation.IndexSections(XElement.Parse(xml));
            nextHierarchyRead = DateTime.UtcNow.AddSeconds(3);
        }
        foreach (string pageId in pageIds.Distinct())
        {
            string section;
            try { App.GetHierarchyParent(pageId, out section); }
            catch (COMException) { continue; } // A closed notebook can still be shown using its saved heading.
            BookmarkLocation owner;
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
    private sealed class MenuRequest
    {
        internal bool PageList;
        internal string PageId;
        internal SavedTab Source;
        internal readonly DateTime Deadline = DateTime.UtcNow.AddSeconds(3);
    }
    private readonly Queue<MenuRequest> menuRequests = new Queue<MenuRequest>();
    private bool openingMenuPage;
    private SavedTab menuSource, followedFrom;
    private SavedPage followedTo;
    private int followedAt;
    private readonly Dictionary<string, PageTabButton> buttons = new Dictionary<string, PageTabButton>();
    private readonly Dictionary<string, NotebookTabLabel> labels = new Dictionary<string, NotebookTabLabel>();
    private readonly ToolStripButton stop;
    private readonly PagePicker add;
    private bool openingPickedPage;
    private readonly BookmarkButton bookmarks = new BookmarkButton();
    private readonly StripMenuButton options = new StripMenuButton {
        AccessibleName = "MoreTabs options", ToolTipText = "MoreTabs options", AutoToolTip = false,
        Text = "Settings", ShowDropDownArrow = false, DisplayStyle = ToolStripItemDisplayStyle.None, Overflow = ToolStripItemOverflow.Never
    };
    private readonly ToolButton back = new ToolButton("Back", ToolIcon.Back), forward = new ToolButton("Forward", ToolIcon.Forward);
    private readonly ToolButton home = new ToolButton("Home", ToolIcon.Home), newNote = new ToolButton("New note", ToolIcon.NewNote);
    private readonly ToolStripMenuItem toolVisibility = new ToolStripMenuItem("Visible tool buttons");
    private readonly ToolStripItem[] toolGroup;
    private readonly ToolStripMenuItem chooseHome = new ToolStripMenuItem("Home page");
    private readonly ToolStripMenuItem chooseSection = new ToolStripMenuItem("New note location");
    private readonly ToolStripMenuItem useQuickNotes = new ToolStripMenuItem("Use default Quick Notes");
    private readonly NotebookTabLabel toolsLabel = new NotebookTabLabel { Text = "Tools", AccessibleName = "Tools" };
    private bool openHomeAfterPicking;
    internal int LeadingToolInset { get { return bookmarks.Padding.Left; } }
    internal Rectangle MenuTabBounds
    {
        get
        {
            var bounds = options.Bounds;
            foreach (var tool in toolGroup)
                if (tool.Available && tool.Placement == ToolStripItemPlacement.Main) bounds = Rectangle.Union(bounds, tool.Bounds);
            return bounds;
        }
    }
    private readonly ToolStripMenuItem colorLabels = new ToolStripMenuItem("Use notebook colors for labels");
    private readonly ToolStripMenuItem showLabels = new ToolStripMenuItem("Show notebook labels");
    private readonly ToolStripMenuItem groupTabs = new ToolStripMenuItem("Group tabs by notebook");
    private readonly ToolStripSeparator bookmarkDivider = new ToolStripSeparator();
    private readonly ContextMenuStrip tabMenu = new ContextMenuStrip();
    private readonly ToolTip notice = new ToolTip();
    private readonly TabCloseInput stripClose, overflowClose;
    private TabRenderer appearance;
    private readonly HostChrome defaultChrome = new HostChrome();
    internal readonly StripHost Host;
    private NativeLayout attachedLayout;
    internal bool AtBottom { get; private set; }
    private bool highContrast;
    private Font tabFont, selectedTabFont, captionFont;
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
        if (pages != null) Text = PageMenuCommand.ReceiverName;
        AccessibleDescription = "Ctrl+Alt+F6 focuses the page tabs. Use the arrow keys, Enter to open, Delete to close a tab, or Escape to return.";
        add = new PagePicker(pages, page => Perform(delegate { OpenPickedPage(page); }),
            delegate { Perform(delegate { CreateNote(null); }); }, delegate { Perform(AddCurrent); });
        bookmarks.DropDownItems.Add("Bookmark current page");
        bookmarks.DropDownOpening += delegate { Perform(BuildBookmarksMenu); };
        options.DropDownItems.AddRange(new ToolStripItem[] { groupTabs, showLabels, colorLabels });
        options.DropDownOpening += delegate { openHomeAfterPicking = false; Perform(delegate { ApplyState(this.store.Read()); }); };
        foreach (ToolStripMenuItem item in options.DropDownItems)
        {
            var setting = item;
            setting.Click += delegate { Perform(delegate { ChangeOption(setting); }); };
        }
        toolGroup = new ToolStripItem[] { bookmarks, options, back, forward, home, newNote };
        foreach (var tool in toolGroup.Where(t => t != options))
        {
            var toggle = new ToolStripMenuItem(tool.AccessibleName) { Tag = tool };
            toggle.Click += delegate { Perform(delegate { ChangeOption(toggle); }); };
            toolVisibility.DropDownItems.Add(toggle);
        }
        options.DropDownItems.AddRange(new ToolStripItem[] { new ToolStripSeparator(), toolVisibility, chooseHome, chooseSection });
        useQuickNotes.Click += delegate { Perform(delegate { SaveNoteSection(null, null); }); };
        HierarchyPicker.Attach(chooseSection, pages, null, (id, name) => Perform(delegate { SaveNoteSection(id, name); }),
            new ToolStripSeparator(), useQuickNotes);
        HierarchyPicker.Attach(chooseHome, pages, page => Perform(delegate {
            ApplyState(this.store.Change(s => { s.HomePage = page; return true; }));
            if (openHomeAfterPicking) GoHome();
        }), null);
        BindTool(back, delegate { NavigateHistory(false); });
        BindTool(forward, delegate { NavigateHistory(true); });
        BindTool(home, GoHome);
        BindTool(newNote, delegate { CreateNote(this.store.Read().NewNoteSectionId); });
        Items.AddRange(toolGroup);
        Items.AddRange(new ToolStripItem[] { bookmarkDivider, add });
        stop = new ToolStripButton("Close prototype") { Alignment = ToolStripItemAlignment.Right };
        stop.Click += delegate { close(); };
        Items.Add(stop);
        Action<SavedTab> closePageTab = tab => Perform(delegate { CloseTab(tab); });
        stripClose = new TabCloseInput(this, closePageTab);
        overflowClose = new TabCloseInput(OverflowButton.DropDown, closePageTab);
        Host = new StripHost(frame, this);
        Host.Captions.Controls.Add(toolsLabel);
        Host.OpenPage += delegate(string id) { Perform(delegate { OpenPageInNewTab(id); }); };
        Host.MoveButton.Click += delegate { Perform(MoveStrip); };
        SetDirection(false);
        defaultChrome.Update(IntPtr.Zero, 96);
        ApplyAppearance(defaultChrome);
        ApplyState(this.store.Read());
    }
    private void Perform(Action action) { try { action(); } catch (Exception e) { onError(e); } }
    private void BindTool(ToolButton button, Action action)
    {
        // Finish the mouse-up and dismiss any overflow menu before changing focus or opening a picker.
        button.Click += delegate { BeginInvoke((Action)delegate { Perform(action); }); };
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            stripClose.Dispose();
            overflowClose.Dispose();
            tabMenu.Dispose();
            notice.Dispose();
            Host.Controls.Remove(this);
            Host.Dispose();
        }
        base.Dispose(disposing);
        if (disposing)
        {
            if (tabFont != null) tabFont.Dispose();
            if (selectedTabFont != null) selectedTabFont.Dispose();
            if (captionFont != null) captionFont.Dispose();
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
        Host.Hide();
    }
    protected override void OnHandleDestroyed(EventArgs e)
    {
        ReleaseFocusShortcut();
        base.OnHandleDestroyed(e);
    }
    protected override void WndProc(ref Message message)
    {
        if (message.Msg == PageMenuCommand.OpenFromMenu && pages != null && message.LParam == IntPtr.Zero &&
            (message.WParam == IntPtr.Zero || message.WParam == (IntPtr)1))
        {
            bool pageList = message.WParam == (IntPtr)1;
            Perform(delegate { QueueMenuRequest(pageList, null); });
            return;
        }
        const int wmHotkey = 0x0312;
        if (message.Msg == wmHotkey && message.WParam.ToInt32() == FocusShortcut)
        {
            if (GetForegroundWindow() == frame) Perform(delegate { FocusPageTabs(ForegroundFocus()); });
            return;
        }
        base.WndProc(ref message);
        // ToolStrip normally consumes the click that brings an inactive frame forward.
        const int wmMouseActivate = 0x0021, maActivateAndEat = 2, maActivate = 1;
        if (message.Msg == wmMouseActivate && message.Result == (IntPtr)maActivateAndEat)
            message.Result = (IntPtr)maActivate;
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
        var choices = Items.Cast<ToolStripItem>().Where(item => (item is ToolStripButton || item is ToolStripDropDownItem) && item.Available && item.Enabled).ToList();
        int index = choices.FindIndex(item => item.Selected);
        if (keyData == Keys.Delete && index >= 0 && choices[index] is PageTabButton)
        {
            var tab = (SavedTab)choices[index].Tag;
            Perform(delegate { CloseTab(tab); FocusPageTabs(IntPtr.Zero); });
            return true;
        }
        if (choices.Count > 0 && (keyData == Keys.Left || keyData == Keys.Right))
        {
            int delta = keyData == Keys.Left ? -1 : 1;
            choices[(Math.Max(index, 0) + delta + choices.Count) % choices.Count].Select();
            return true;
        }
        if ((keyData == Keys.Enter || keyData == Keys.Space) && index >= 0)
        {
            if (choices[index] is ToolStripDropDownItem) ((ToolStripDropDownItem)choices[index]).ShowDropDown();
            else { choices[index].PerformClick(); if (!options.DropDown.Visible && choices[index] != newNote) ReturnToEditor(); }
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
    [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int key);
    internal static bool MenuInputActive(IntPtr frame, uint thread)
    {
        var info = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
        return (thread != 0 && GetGUIThreadInfo(thread, ref info) && (info.Flags & 0x1c) != 0) ||
            (GetForegroundWindow() == frame && GetAsyncKeyState(2) < 0);
    }
    [DllImport("user32.dll")] private static extern IntPtr SetFocus(IntPtr window);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);
    [DllImport("user32.dll")] private static extern bool UnregisterHotKey(IntPtr window, int id);

    internal void Attach(Rectangle slot, NativeLayout layout = null)
    {
        attachedLayout = layout;
        Host.Bounds = slot;
        if (layout != null)
        {
            layout.StripWindow = Host.Handle;
            layout.ShowCaptions = WantsCaptionRow;
            if (AtBottom != layout.AtBottom) SetDirection(layout.AtBottom);
            ApplyAppearance(layout.Appearance(!Host.Visible));
        }
        if (NativeLayout.GetParent(Host.Handle) != frame || NativeLayout.Bounds(Host.Handle, frame) != slot)
            throw new InvalidOperationException("The strip did not attach within the reserved row.");
        Host.Show();
    }
    private void MoveStrip()
    {
        if (attachedLayout == null) return;
        bool previous = AtBottom;
        try
        {
            Rectangle slot = attachedLayout.Reserve(false, !previous);
            Attach(slot, attachedLayout);
            attachedLayout.CheckHitTargets(Host.Handle, slot);
        }
        catch (Exception error)
        {
            if (!(error is InvalidOperationException || error is System.ComponentModel.Win32Exception)) throw;
            // A preflight rejection leaves the existing row intact. Roll back if the host changed during the move.
            if (attachedLayout.AtBottom != previous) Attach(attachedLayout.Reserve(false, previous), attachedLayout);
            Console.WriteLine("Could not move the tab strip: " + error.Message);
            notice.Show(error.Message, Host.MoveButton, 0, 0, 4000);
        }
    }
    private void SetDirection(bool bottom)
    {
        AtBottom = bottom;
        Host.Captions.Dock = bottom ? DockStyle.Top : DockStyle.Bottom;
        DefaultDropDownDirection = bottom ? ToolStripDropDownDirection.AboveRight : ToolStripDropDownDirection.BelowRight;
        Host.MoveButton.PointsUp = bottom;
        Host.MoveButton.Invalidate();
        Host.MoveButton.AccessibleName = bottom ? "Move tabs to top" : "Move tabs to bottom";
        notice.SetToolTip(Host.MoveButton, Host.MoveButton.AccessibleName);
        if (appearance != null)
        {
            var margin = TabMargin(appearance.Dpi);
            SetMenuTabMargins(margin);
            foreach (var tab in buttons.Values) tab.Margin = margin;
            Invalidate();
        }
    }
    internal void ExplainTopPlacement()
    {
        const string message = "Tabs moved to the top because OneNote is using the bottom edge.";
        Console.WriteLine(message);
        notice.Show(message, Host.MoveButton, 0, Host.Height, 4000);
    }
    internal void ApplyAppearance(HostChrome chrome)
    {
        bool metrics = appearance == null || appearance.Dpi != chrome.Dpi;
        if (metrics || appearance.Surface.ToArgb() != chrome.Surface.ToArgb() || highContrast != chrome.HighContrast ||
            appearance.DarkSections != chrome.DarkSections)
        {
            highContrast = chrome.HighContrast;
            appearance = new TabRenderer(chrome.Surface, chrome.Dpi, highContrast, chrome.DarkSections);
            Renderer = appearance;
            toolVisibility.DropDown.Renderer = appearance;
            tabMenu.Renderer = bookmarks.DropDown.Renderer = options.DropDown.Renderer = OverflowButton.DropDown.Renderer = appearance;
            add.DropDown.Renderer = appearance.NotebookMenu;
            BackColor = appearance.Surface;
            ForeColor = appearance.Ink;
            Host.Style(appearance);
            foreach (var label in labels.Values) StyleCaption(label);
            toolsLabel.ForeColor = appearance.Ink;
            if (state != null) RefreshOptions();
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
            SetMenuTabMargins(tabMargin);
            add.TabWidth = HostChrome.Px(32, chrome.Dpi);
            Font oldRegular = tabFont, oldSelected = selectedTabFont, oldCaption = captionFont;
            var menuFont = SystemFonts.MenuFont;
            tabFont = new Font(menuFont.FontFamily, menuFont.SizeInPoints * chrome.Dpi / 72f, FontStyle.Regular, GraphicsUnit.Pixel);
            selectedTabFont = new Font(tabFont, FontStyle.Bold);
            captionFont = HostChrome.CaptionFont(chrome.Dpi);
            Font = tabMenu.Font = tabFont;
            options.DropDown.Font = tabFont;
            toolsLabel.Font = captionFont;
            foreach (var tool in Items.OfType<ToolButton>())
            {
                tool.TabWidth = HostChrome.Px(25, chrome.Dpi);
                tool.Enabled = pages != null;
            }
            var commandPadding = new Padding(HostChrome.Px(5, chrome.Dpi), 0, HostChrome.Px(5, chrome.Dpi), 0);
            foreach (ToolStripItem item in Items)
                if (!(item is PageTabButton) && !(item is ToolStripSeparator) && !(item is ToolButton) && item != bookmarks && item != options) item.Padding = commandPadding;
            foreach (var divider in Items.OfType<ToolStripSeparator>()) SizeDivider(divider);
            foreach (var tab in Items.OfType<PageTabButton>()) SizeTab(tab);
            foreach (var label in labels.Values) label.Font = captionFont;
            if (oldRegular != null) oldRegular.Dispose();
            if (oldSelected != null) oldSelected.Dispose();
            if (oldCaption != null) oldCaption.Dispose();
        }
    }
    private Padding TabMargin(uint dpi)
    {
        int side = HostChrome.Px(1, dpi);
        int gap = HostChrome.Px(3, dpi);
        return new Padding(side, AtBottom ? gap : 0, side, AtBottom ? 0 : gap);
    }
    private void SetMenuTabMargins(Padding margin)
    {
        add.Margin = margin;
        foreach (var tool in toolGroup)
            tool.Margin = new Padding(0, margin.Top, 0, margin.Bottom);
    }
    private void SizeDivider(ToolStripSeparator divider)
    {
        uint dpi = appearance.Dpi;
        divider.AutoSize = false;
        // Tools have no side margins. Match the first tab's inset on the Tools side.
        divider.Margin = divider == bookmarkDivider ? new Padding(TabMargin(dpi).Left, 0, 0, 0) : Padding.Empty;
        divider.Size = new Size(2 * HostChrome.Px(4, dpi) + 1, HostChrome.Px(25, dpi));
        divider.Overflow = ToolStripItemOverflow.Never;
    }
    private void SizeTab(PageTabButton tab)
    {
        uint dpi = appearance.Dpi;
        tab.Dpi = dpi;
        // Measure both states at the heavier weight so selecting a tab never shifts its neighbors.
        tab.Font = selectedTabFont;
        tab.Margin = TabMargin(dpi);
        tab.Padding = new Padding(HostChrome.Px(10, dpi), 0, tab.CloseWidth, 0);
    }
    private void StyleCaption(NotebookTabLabel label)
    {
        label.ForeColor = appearance.HighContrast || !state.ColorNotebookLabels || label.Notebook == null || label.Notebook.Color.IsEmpty ? appearance.Ink : label.Notebook.Color;
    }
    internal bool WantsCaptionRow { get { return state != null && state.GroupByNotebook && state.ShowNotebookLabels; } }
    private void RefreshOptions()
    {
        foreach (ToolStripMenuItem toggle in toolVisibility.DropDownItems)
        {
            var tool = (ToolStripItem)toggle.Tag;
            bool show = !state.HiddenTools.Contains(tool.Name);
            toggle.Checked = show;
            if (tool.Available != show) tool.Available = show;
        }
        groupTabs.Checked = state.GroupByNotebook;
        showLabels.Enabled = state.GroupByNotebook;
        showLabels.Checked = WantsCaptionRow;
        colorLabels.Enabled = WantsCaptionRow && !highContrast;
        colorLabels.Checked = colorLabels.Enabled && state.ColorNotebookLabels;
        useQuickNotes.Checked = state.NewNoteSectionId == null;
        chooseSection.Text = "New note location (" + (state.NewNoteSectionName ?? "Quick Notes").Replace("&", "&&") + ")";
        chooseHome.Text = "Home page (" + (state.HomePage == null ? "Choose a page" : state.HomePage.Caption.Replace("&", "&&")) + ")";
    }
    private void SaveNoteSection(string id, string name)
    {
        ApplyState(store.Change(s => { s.NewNoteSectionId = id; s.NewNoteSectionName = name; return true; }));
    }
    private void GoHome()
    {
        var page = store.Read().HomePage;
        if (page == null)
        {
            options.ShowDropDown();
            openHomeAfterPicking = true;
            chooseHome.Select();
            chooseHome.ShowDropDown();
            return;
        }
        if (activeKey == null) { OpenSaved(page); return; }
        pendingTab = null;
        focusPending = false;
        if (NavigateSaved(page)) RefreshActive();
    }
    private void NavigateHistory(bool goForward)
    {
        if (GetForegroundWindow() != frame) return;
        ReturnToEditor();
        var focus = ForegroundFocus();
        if (NativeLayout.RootOf(focus) != frame || ContainsFocus) return;
        // OneNote owns the window history, including visits made through its native navigation.
        pendingTab = null;
        focusPending = false;
        SendKeys.SendWait(goForward ? "%{RIGHT}" : "%{LEFT}");
    }
    private void ChangeOption(ToolStripMenuItem item)
    {
        if (!item.Enabled) return;
        bool enable = !item.Checked;
        try
        {
            ApplyState(store.Change(s => {
                var tool = item.Tag as ToolStripItem;
                if (tool != null) return enable ? s.HiddenTools.Remove(tool.Name) : s.HiddenTools.Add(tool.Name);
                if (item != groupTabs && !s.GroupByNotebook) return false;
                if (item == colorLabels && !s.ShowNotebookLabels) return false;
                if (item == colorLabels) s.ColorNotebookLabels = enable;
                else if (item == showLabels) s.ShowNotebookLabels = enable;
                else s.GroupByNotebook = enable;
                return true;
            }));
        }
        catch (System.IO.IOException) { notice.Show("Could not save MoreTabs options.", this, 4, Height, 4000); }
        catch (UnauthorizedAccessException) { notice.Show("Could not save MoreTabs options.", this, 4, Height, 4000); }
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
        menuSource = followedFrom = null;
        pendingTab = tab;
        focusPending = focus;
        navigationDeadline = DateTime.UtcNow.AddSeconds(5);
        nextCaptionCheck = DateTime.MinValue;
    }
    private SavedPage NavigateForNewTab(string id)
    {
        BookmarkLocation location;
        pages.PageLocations(new[] { id }).TryGetValue(id, out location);
        var page = new SavedPage(id, pages.Caption(id), pages.LinkFor(id), location);
        if (pages.Current != id) pages.Navigate(id);
        return page;
    }
    internal void OpenPageInNewTab(string id)
    {
        QueueMenuRequest(true, id);
    }
    private void QueueMenuRequest(bool pageList, string pageId)
    {
        menuRequests.Enqueue(new MenuRequest { PageList = pageList, Source = pageList ? menuSource : null, PageId = pageId });
        menuSource = null;
        ProcessMenuRequest();
    }
    internal void MarkPageMenuGesture(int inputTime)
    {
        menuSource = null;
        if (activeKey == null || pendingTab != null) return;
        var tab = state.Tabs.Find(t => t.Key == activeKey);
        if (tab == null) return;
        // Input can be queued while the strip is inside a OneNote COM call. Recover the
        // source only when that follow happened after this gesture and into this exact slot.
        menuSource = followedFrom != null && tab.Key == followedFrom.Key && tab.Page.Matches(followedTo) &&
            unchecked((uint)(followedAt - inputTime)) < 60000 ? followedFrom : tab;
    }
    internal bool ProcessMenuRequest()
    {
        if (openingMenuPage) return true;
        if (menuRequests.Count == 0) return false;
        var request = menuRequests.Peek();
        openingMenuPage = true;
        try
        {
            SavedPage page = null;
            try
            {
                if (request.PageId == null) request.PageId = request.PageList ?
                    pages.SelectedPage(request.Source == null ? null : request.Source.Page.Id) : pages.Current;
                if (!String.IsNullOrEmpty(request.PageId)) page = NavigateForNewTab(request.PageId);
            }
            catch (COMException e)
            {
                if (Program.EditorBusy(e) && DateTime.UtcNow < request.Deadline) return true;
                menuRequests.Dequeue();
                Console.WriteLine("OneNote page menu navigation failed (0x{0:X8}).", e.HResult);
                notice.Show("OneNote could not open the selected page.", this, 4, Height, 4000);
                return true;
            }
            menuRequests.Dequeue();
            if (page == null)
            {
                Console.WriteLine("OneNote page menu: no single page could be identified.");
                notice.Show("Select a single page in OneNote's page list and try again.", this, 4, Height, 4000);
            }
            else if (AddOpenedTab(page, request.Source)) Console.WriteLine("OneNote page menu: opened a new tab ({0} tabs).", TabCount);
            return true;
        }
        finally { openingMenuPage = false; }
    }
    private bool AddOpenedTab(SavedPage page, SavedTab source = null)
    {
        var tab = new SavedTab(Guid.NewGuid().ToString("N"), page);
        try
        {
            ApplyState(store.Change(s => {
                if (source != null && source.Key == activeKey && followedFrom != null && followedFrom.Key == source.Key &&
                    followedFrom.Page.Matches(source.Page) && followedTo.Matches(page))
                    s.UpdateTab(new SavedTab(source.Key, followedTo), source.Page);
                s.Tabs.Add(tab);
                return true;
            }));
        }
        catch (System.IO.IOException) { ExplainUnsavedTab(); return false; }
        catch (UnauthorizedAccessException) { ExplainUnsavedTab(); return false; }
        ExpectNavigation(tab, true);
        RefreshActive();
        return true;
    }
    private void ExplainUnsavedTab()
    {
        activeKey = null;
        pendingTab = null;
        focusPending = false;
        notice.Show("The page opened, but MoreTabs could not save a new tab.", this, 4, Height, 4000);
    }
    private void CreateNote(string sectionId)
    {
        pendingTab = null;
        focusPending = false;
        string id;
        try { id = pages.CreateNote(sectionId); }
        catch (Exception error)
        {
            if (!(error is COMException || error is InvalidOperationException || error is UnauthorizedAccessException)) throw;
            Console.WriteLine("Note creation failed ({0}, 0x{1:X8}).", error.GetType().Name, error.HResult);
            notice.Show(sectionId == null ? "OneNote could not create a Quick Note. Check the Quick Notes location and write access." :
                "OneNote could not create the note. Check write access or choose another section in Settings.", this, 4, Height, 6000);
            return;
        }
        var page = new SavedPage(id, "Untitled", pages.LinkFor(id));
        SavedTab saved = null;
        try { saved = KeepTab(page); }
        catch (System.IO.IOException) { }
        catch (UnauthorizedAccessException) { }
        if (saved != null) ExpectNavigation(saved, true);
        try { pages.Navigate(id); }
        catch (COMException error)
        {
            pendingTab = null;
            focusPending = false;
            Console.WriteLine("Opening the newly created note failed (0x{0:X8}).", error.HResult);
            notice.Show(saved != null ? "Note created. Click its tab to try opening it again." :
                "Note created in OneNote. MoreTabs could not save or open its tab.", this, 4, Height, 6000);
            return;
        }
        if (saved == null) { activeKey = null; notice.Show("Note opened. MoreTabs could not save its tab.", this, 4, Height, 6000); }
        RefreshActive();
    }
    internal string ActiveTabKey { get { return activeKey; } }
    internal void ConnectNativePageMenu(One.Application app)
    {
        options.DropDownItems.Add(new ToolStripSeparator());
        options.DropDownItems.Add("Connect OneNote page menu", null, delegate {
            MessageBox.Show(Host, PageMenuConnection.Connect(app, true), "MoreTabs");
        });
        Console.WriteLine(PageMenuConnection.Connect(app, false));
    }
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
        RefreshOptions();
        if (displayedRevision == updated.Revision) return;
        bool resize = attachedLayout != null && attachedLayout.ShowCaptions != WantsCaptionRow;
        SuspendLayout();
        try { UpdateTabItems(); displayedRevision = updated.Revision; }
        finally { ResumeLayout(true); }
        if (resize && Host.Visible && attachedLayout.StillReserved(Host.Bounds))
        {
            Rectangle slot = attachedLayout.Reserve(true);
            if (!slot.IsEmpty) Attach(slot, attachedLayout);
        }
    }
    private void UpdateTabItems()
    {
        foreach (var id in buttons.Keys.Where(id => !state.Tabs.Any(t => t.Key == id)).ToArray())
        {
            if (id == activeKey) activeKey = null;
            var obsolete = buttons[id];
            buttons.Remove(id);
            Items.Remove(obsolete);
            obsolete.Dispose();
        }
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
            button.Text = page.BookmarkTitle;
            button.AccessibleName = page.Caption;
            button.ToolTipText = page.Caption + (slot.Pinned ? "\nPinned tab" : "") + "\nClick × to close this tab. Right-click for tab options.";
            button.Tag = slot;
        }
        int position = Items.IndexOf(bookmarkDivider) + 1;
        var used = new HashSet<string>();
        var known = state.Tabs.Select(t => t.Page.Notebook).Concat(state.Bookmarks.Select(p => p.Notebook))
            .Where(book => book != null).GroupBy(book => book.Id).Select(g => g.First())
            .ToLookup(book => book.Name, StringComparer.CurrentCultureIgnoreCase);
        foreach (bool pinned in new[] { true, false })
        {
            // GroupBy keeps the order of first appearance and the saved order inside each notebook.
            var groups = state.Tabs.Where(tab => tab.Pinned == pinned).Select(tab => {
                var book = tab.Page.NotebookForMenu();
                if (book.Id.Length == 0 && known[book.Name].Count() == 1) book = known[book.Name].First();
                return new { Tab = tab, Book = book };
            }).GroupBy(entry => !pinned && state.GroupByNotebook ? entry.Book.GroupKey : "");
            foreach (var group in groups)
            {
                var key = (pinned ? "p:" : "r:") + group.Key;
                used.Add(key);
                NotebookTabLabel label;
                if (!labels.TryGetValue(key, out label))
                {
                    labels.Add(key, label = new NotebookTabLabel { Font = captionFont, Visible = false, Divider = new ToolStripSeparator() });
                    SizeDivider(label.Divider);
                    Host.Captions.Controls.Add(label);
                }
                label.Notebook = !pinned && state.GroupByNotebook ? group.First().Book : null;
                string caption = pinned ? "Pinned" : label.Notebook == null ? "" : label.Notebook.Name;
                if (label.Text != caption)
                {
                    label.Text = label.AccessibleName = caption;
                    notice.SetToolTip(label, label.Text);
                }
                StyleCaption(label);
                foreach (var entry in group)
                {
                    var button = buttons[entry.Tab.Key];
                    button.Group = label;
                    PlaceItem(button, position++);
                }
                PlaceItem(label.Divider, position++);
            }
        }
        foreach (string key in labels.Keys.Where(key => !used.Contains(key)).ToArray())
        {
            labels[key].Divider.Dispose();
            labels[key].Dispose();
            labels.Remove(key);
        }
        Host.Captions.Visible = WantsCaptionRow;
        if (attachedLayout != null) attachedLayout.ShowCaptions = WantsCaptionRow;
    }
    private void PlaceItem(ToolStripItem item, int index)
    {
        if (Items.IndexOf(item) == index) return;
        Items.Remove(item);
        Items.Insert(index, item);
    }
    protected override void OnLayout(LayoutEventArgs e)
    {
        if (appearance != null && stop != null) FitNotebookGroups();
        base.OnLayout(e);
        if (appearance != null && Host != null) AlignNotebookCaptions();
    }
    private void AlignNotebookCaptions()
    {
        var tools = MenuTabBounds;
        toolsLabel.Bounds = Rectangle.FromLTRB(tools.Left + TabRenderer.CornerRadius(appearance.Dpi), 0, tools.Right, Host.Captions.Height);
        // Use the final main-row item bounds. Overflow pages and notebook names consume no caption width.
        var visible = Items.OfType<PageTabButton>().Where(t => t.Placement == ToolStripItemPlacement.Main).ToLookup(t => t.Group);
        foreach (var label in labels.Values)
        {
            var tabs = visible[label];
            var first = tabs.FirstOrDefault();
            if (first == null || !WantsCaptionRow) { label.Hide(); continue; }
            label.Bounds = Rectangle.FromLTRB(first.Bounds.Left + TabRenderer.CornerRadius(first.Dpi), 0,
                tabs.Last().Bounds.Right, Host.Captions.Height);
            label.Show();
        }
    }
    private static int ItemWidth(ToolStripItem item)
    {
        var tab = item as PageTabButton;
        if (tab != null) return tab.StripSize(Size.Empty).Width + tab.Margin.Horizontal;
        return (item.AutoSize ? item.GetPreferredSize(Size.Empty).Width : item.Width) + item.Margin.Horizontal;
    }
    private void FitNotebookGroups()
    {
        var tabs = Items.OfType<PageTabButton>().Where(t => t.Group != null).ToList();
        foreach (var tab in tabs) tab.WidthLimit = 0;
        var visible = new HashSet<ToolStripItem>();
        var visibleGroups = new HashSet<NotebookTabLabel>();
        // With the star hidden the gear is the first icon and alone may have to span the Tools caption.
        options.TabWidth = bookmarks.Available ? HostChrome.Px(25, appearance.Dpi) : Math.Max(bookmarks.TabWidth, HostChrome.Px(34, appearance.Dpi));
        int space = DisplayRectangle.Width - ItemWidth(options) - ItemWidth(bookmarkDivider) - ItemWidth(add);
        if (bookmarks.Available) space -= ItemWidth(bookmarks);
        int dividerWidth = bookmarkDivider.Width;
        var tools = toolGroup.OfType<ToolButton>().Where(t => t.Available).ToArray();
        int toolsWidth = tools.Sum(tool => ItemWidth(tool));
        int overflowWidth = Math.Max(OverflowButton.Width, HostChrome.Px(16, appearance.Dpi));
        bool collapseTools = tools.Length > 0 && space - toolsWidth < overflowWidth + (tabs.Count == 0 ? 0 : HostChrome.Px(70, appearance.Dpi) + dividerWidth);
        foreach (var tool in tools) tool.Overflow = collapseTools ? ToolStripItemOverflow.Always : ToolStripItemOverflow.Never;
        if (!collapseTools) space -= toolsWidth;
        if (collapseTools || tabs.Sum(tab => ItemWidth(tab)) + labels.Count * dividerWidth + ItemWidth(stop) > space) space -= overflowWidth;
        PageTabButton active;
        if (activeKey != null && buttons.TryGetValue(activeKey, out active) && active.Group != null)
        {
            int content = space - active.Margin.Horizontal - dividerWidth;
            if (content >= HostChrome.Px(48, appearance.Dpi)) active.WidthLimit = content;
        }
        // Reserve the selected page first, then fit other pages in the order they appear on the strip.
        foreach (var tab in tabs.OrderByDescending(t => ((SavedTab)t.Tag).Key == activeKey))
        {
            int cost = ItemWidth(tab) + (visibleGroups.Contains(tab.Group) ? 0 : dividerWidth);
            if (cost > space) continue;
            visible.Add(tab);
            visibleGroups.Add(tab.Group);
            space -= cost;
        }
        foreach (var tab in tabs) tab.Overflow = visible.Contains(tab) ? ToolStripItemOverflow.Never : ToolStripItemOverflow.Always;
        foreach (var label in labels.Values) label.Divider.Available = visibleGroups.Contains(label);
        stop.Overflow = space >= ItemWidth(stop) ? ToolStripItemOverflow.Never : ToolStripItemOverflow.Always;
    }
    internal ContextMenuStrip TabMenuFor(SavedTab tab)
    {
        ClearMenu(tabMenu.Items);
        tabMenu.Items.Add("Open in new tab", null, delegate { Perform(delegate { OpenSaved(tab.Page, newTab: true); }); });
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
        BookmarkLocation owner;
        if (pages.PageLocations(new[] { page.Id }).TryGetValue(page.Id, out owner))
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
        bookmarks.DropDown.MinimumSize = Size.Empty;
        menu.Add("Bookmark current page", null, delegate { Perform(delegate { SaveBookmark(CaptureCurrent()); }); });
        if (state.Bookmarks.Count == 0) return;
        var owners = pages.PageLocations(state.Bookmarks.Select(page => page.Id), true);
        if (state.Bookmarks.Any(p => owners.ContainsKey(p.Id) && !owners[p.Id].SameAs(p.Location)))
        {
            // The menu can still use live colors if the optional offline cache cannot be written.
            try { ApplyState(store.Change(s => s.UpdateBookmarkOwners(owners))); }
            catch (System.IO.IOException) { }
            catch (UnauthorizedAccessException) { }
        }
        if (state.Bookmarks.Count == 0) return;
        var remove = new ToolStripMenuItem("Remove bookmark");
        remove.DropDown.Renderer = appearance;
        bookmarks.DropDown.ShowItemToolTips = remove.DropDown.ShowItemToolTips = true;
        var booksById = owners.Values.Select(o => o.Notebook).Concat(state.Bookmarks.Select(p => p.Notebook).Where(n => n != null))
            .GroupBy(n => n.Id).ToDictionary(g => g.Key, g => g.First());
        var known = booksById.Values.ToLookup(n => n.Name, StringComparer.CurrentCultureIgnoreCase);
        var entries = state.Bookmarks.Select(page => {
            BookmarkLocation location;
            if (!owners.TryGetValue(page.Id, out location)) location = page.Location;
            var book = location != null ? booksById[location.Notebook.Id] : page.NotebookForMenu();
            if (book.Id.Length == 0 && known[book.Name].Count() == 1) book = known[book.Name].First();
            return new { Page = page, Book = book, Location = location };
        }).ToList();
        int screenWidth = Screen.FromHandle(Handle).WorkingArea.Width;
        int sectionWidth = BookmarkRow.MeasureColumn(entries.Select(entry => entry.Location), bookmarks.DropDown.Font,
            appearance.Dpi, screenWidth);
        var groups = entries.GroupBy(item => item.Book.GroupKey)
          .OrderBy(group => group.First().Book.Name, StringComparer.CurrentCultureIgnoreCase);
        foreach (var group in groups)
        {
            var book = group.First().Book;
            menu.Add(new BookmarkHeader(book));
            remove.DropDownItems.Add(new BookmarkHeader(book));
            foreach (var entry in group)
            {
                var page = entry.Page;
                menu.Add(new BookmarkRow(page, entry.Location, sectionWidth, appearance.Dpi, delegate { Perform(delegate { OpenSaved(page); }); }));
                remove.DropDownItems.Add(new BookmarkRow(page, entry.Location, sectionWidth, appearance.Dpi, delegate { Perform(delegate { RemoveBookmark(page); }); }));
            }
        }
        menu.Add(new ToolStripSeparator());
        menu.Add(remove);
        int width = menu.OfType<BookmarkRow>().Max(row => row.ContentWidth);
        SizeBookmarkMenu(bookmarks.DropDown, width, screenWidth);
        SizeBookmarkMenu(remove.DropDown, width, screenWidth);
    }
    private static void SizeBookmarkMenu(ToolStripDropDown menu, int contentWidth, int screenWidth)
    {
        // The native menu measures only standard text and icon columns. Widen both the popup and its hit targets.
        int border = 2 * SystemInformation.BorderSize.Width;
        int width = Math.Min(screenWidth - border, Math.Max(contentWidth, menu.GetPreferredSize(Size.Empty).Width - border));
        foreach (ToolStripItem item in menu.Items)
        {
            int height = item.GetPreferredSize(Size.Empty).Height;
            item.AutoSize = false;
            item.Size = new Size(width - item.Margin.Horizontal, height);
        }
        menu.MinimumSize = new Size(width + border, 0);
    }
    private void OpenPickedPage(SavedPage page)
    {
        openingPickedPage = true;
        try { OpenSaved(page, newTab: true); }
        finally { openingPickedPage = false; }
    }
    internal bool OpenSaved(SavedPage page, SavedTab target = null, bool newTab = false)
    {
        pendingTab = null;
        focusPending = false;
        if (!NavigateSaved(page)) return false;
        if (newTab) return AddOpenedTab(page);
        ExpectNavigation(target ?? KeepTab(page), false);
        RefreshActive();
        return true;
    }
    private bool NavigateSaved(SavedPage page)
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
        return true;
    }
    internal void RefreshActive()
    {
        // NavigateTo can pump this window's timer before the new slot is saved.
        if (openingPickedPage && pendingTab == null) return;
        if (DateTime.UtcNow >= nextRead)
        {
            ApplyState(store.Read());
            nextRead = DateTime.UtcNow.AddSeconds(2);
        }
        string id = pages.Current;
        int observedAt = Environment.TickCount;
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
                if (pages.Current != id || activeKey != activeTab.Key || pendingTab != null) return;
                bool followed = false;
                ApplyState(store.Change(s => followed = s.UpdateTab(activeTab, next)));
                if (followed)
                {
                    followedFrom = activeTab;
                    followedTo = next;
                    followedAt = observedAt;
                }
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
            if (entry.Value.Checked != active)
            {
                entry.Value.Checked = active;
                entry.Value.AccessibleDescription = active ? "Active page tab" : "Page tab";
                entry.Value.Invalidate();
            }
        }
        if (selectionChanged) PerformLayout();
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
        var lookup = due.ToDictionary(pageId => pageId, pageId =>
            currentId != null && byPage[pageId].Contains(active) ? currentId : pageId);
        var locations = pages.PageLocations(lookup.Values.Distinct());
        var owners = new Dictionary<string, BookmarkLocation>();
        foreach (string pageId in due)
        {
            var tabs = byPage[pageId];
            BookmarkLocation owner;
            bool found = locations.TryGetValue(lookup[pageId], out owner);
            if (found && tabs.Any(t => !owner.Notebook.SameAs(((SavedTab)t.Tag).Page.Notebook))) owners.Add(pageId, owner);
            Color color = found ? owner.SectionColor : Color.Empty;
            foreach (var tab in tabs)
            {
                tab.NextColorCheck = now.AddSeconds(3);
                if (tab.SectionColor == color) continue;
                tab.SectionColor = color;
                tab.Invalidate();
            }
        }
        if (owners.Count > 0)
        {
            try { ApplyState(store.Change(s => s.UpdateTabNotebooks(owners))); }
            catch (System.IO.IOException) { ShowUncachedNotebooks(owners); }
            catch (UnauthorizedAccessException) { ShowUncachedNotebooks(owners); }
        }
    }
    private void ShowUncachedNotebooks(Dictionary<string, BookmarkLocation> owners)
    {
        state.UpdateTabNotebooks(owners);
        SuspendLayout();
        try { UpdateTabItems(); }
        finally { ResumeLayout(true); }
    }
    internal void ClickNewQuickNote() { add.QuickNote.PerformClick(); }
    internal void ClickKeepCurrent() { add.KeepCurrent.PerformClick(); }
    internal void ClickTab(string id) { buttons[state.Tabs.First(t => t.Page.Id == id).Key].PerformClick(); }
    internal bool IsSelected(string id) { return buttons.Values.Any(b => b.Checked && ((SavedTab)b.Tag).Page.Id == id); }
    internal bool HasTabFocus(string id) { return ContainsFocus && IsSelected(id) && buttons[activeKey].Selected; }
    internal int TabCount { get { return buttons.Count; } }
}
