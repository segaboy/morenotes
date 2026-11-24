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
    private readonly Action<Exception> onError;
    internal PageStrip(IntPtr frame, IPages pages, Action close, Action<Exception> onError)
    {
        this.frame = frame;
        this.pages = pages;
        this.onError = onError;
        AutoSize = false;
        GripStyle = ToolStripGripStyle.Hidden;
        Dock = DockStyle.None;
        TabStop = true;
        AccessibleName = "MoreTabs page tabs";
        add = new ToolStripButton("+") { AccessibleName = "Add current page", ToolTipText = "Keep the displayed page in a tab" };
        add.Click += delegate { Perform(AddCurrent); };
        Items.Add(add);
        var stop = new ToolStripButton("Close prototype") { Alignment = ToolStripItemAlignment.Right };
        stop.Click += delegate { close(); };
        Items.Add(stop);
    }
    protected override CreateParams CreateParams
    {
        get { var cp = base.CreateParams; cp.Parent = frame; cp.Style |= 0x40000000; return cp; }
    }
    private void Perform(Action action) { try { action(); } catch (Exception e) { onError(e); } }

    internal void Attach(Rectangle slot)
    {
        Bounds = slot;
        if (NativeLayout.GetParent(Handle) != frame || NativeLayout.Bounds(Handle, frame) != slot)
            throw new InvalidOperationException("The strip did not attach within the reserved row.");
        Show();
    }
    internal void AddCurrent()
    {
        string id = pages.Current;
        if (String.IsNullOrEmpty(id)) return;
        if (!buttons.ContainsKey(id))
        {
            var button = new ToolStripButton(pages.Caption(id)) { DisplayStyle = ToolStripItemDisplayStyle.Text };
            button.Click += delegate { Perform(delegate { pages.Navigate(id); RefreshActive(); }); };
            buttons.Add(id, button);
            Items.Insert(Items.IndexOf(add), button);
        }
        RefreshActive();
    }
    internal void RefreshActive()
    {
        string id = pages.Current;
        foreach (var entry in buttons)
            if (entry.Value.Checked != (entry.Key == id)) entry.Value.Checked = entry.Key == id;
    }
    internal void ClickAdd() { add.PerformClick(); }
    internal void ClickTab(string id) { buttons[id].PerformClick(); }
    internal bool IsSelected(string id) { return buttons[id].Checked; }
    internal int TabCount { get { return buttons.Count; } }
}
