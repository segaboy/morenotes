using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;

// Experimental native window classes observed in OneNote build 16.0.20715.20004.
internal sealed class NativeLayout : IDisposable
{
    internal readonly IntPtr Frame;
    internal readonly IntPtr Document;
    internal readonly IntPtr Ribbon;
    private readonly IntPtr navigation;
    private readonly HostChrome chrome = new HostChrome();
    internal HostChrome Appearance(bool force = false)
    {
        chrome.Update(navigation, Dpi, force);
        return chrome;
    }
    private Rectangle original, applied;
    internal bool AtBottom { get; private set; }
    internal IntPtr StripWindow;

    internal NativeLayout(IntPtr frame, IntPtr document, IntPtr ribbon, IntPtr navigationRow)
    {
        Frame = frame;
        Document = document;
        Ribbon = ribbon;
        navigation = navigationRow;
    }

    internal static List<NativeLayout> Discover(bool requireMatch = true)
    {
        var result = new List<NativeLayout>();
        var observed = new List<string>();
        foreach (IntPtr frame in Children(IntPtr.Zero, "Framework::CFrame"))
        {
            if (IsIconic(frame) || !IsWindowEnabled(frame)) { observed.Add("Frame " + frame + " minimized or modal"); continue; }
            MatchDpi(frame);
            observed.Add("Frame " + frame + " at " + GetDpiForWindow(frame) + " DPI");
            foreach (IntPtr doc in Children(frame, "OneNote::MainWindow"))
            {
                Rectangle body = Bounds(doc, frame);
                observed.Add("document " + body);
                foreach (IntPtr dock in Children(frame, "MsoCommandBarDock"))
                {
                    Rectangle top = Bounds(dock, frame);
                    observed.Add("ribbon dock " + top);
                    foreach (IntPtr nav in Children(doc, "OneNote::NavigationUI"))
                    {
                        Rectangle row = Bounds(nav, frame);
                        observed.Add("navigation " + row);
                        if (top.Bottom != body.Top || top.Left > body.Left || top.Right < body.Right ||
                            row.Top != body.Top || row.Left != body.Left || row.Width != body.Width ||
                            row.Height < 12 || row.Height > GetDpiForWindow(frame) ||
                            row.Height >= body.Height) continue;
                        result.Add(new NativeLayout(frame, doc, dock, nav));
                    }
                }
            }
        }
        if (requireMatch && result.Count == 0) throw new InvalidOperationException(observed.Count == 0
            ? "Native discovery found no visible Framework::CFrame windows on this desktop."
            : "Native discovery found no adjacent ribbon, document and full-width navigation row. Observed: " + String.Join("; ", observed));
        return result;
    }

    internal static IntPtr RootOf(IntPtr handle)
    {
        const uint root = 2; // GA_ROOT follows parents, never owners of other top-level windows.
        return GetAncestor(handle, root);
    }

    internal static void MatchDpi(IntPtr frame)
    {
        if (SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(frame)) == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not match the host's DPI context.");
    }

    internal Rectangle DocumentBounds { get { return Bounds(Document, Frame); } }
    internal Rectangle NavigationBounds { get { return Bounds(navigation, Frame); } }
    internal Rectangle RibbonBounds { get { return Bounds(Ribbon, Frame); } }
    private Rectangle ClientBounds
    {
        get
        {
            RECT r;
            if (!GetClientRect(Frame, out r)) throw new Win32Exception(Marshal.GetLastWin32Error());
            return Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
        }
    }
    internal bool StillReserved(Rectangle slot)
    {
        if (!Alive || slot.IsEmpty || DocumentBounds != applied || applied.Top != NavigationBounds.Top) return false;
        if (AtBottom)
        {
            var client = ClientBounds;
            return slot.Left == client.Left && slot.Right == client.Right && slot.Bottom == client.Bottom &&
                slot.Top == applied.Bottom && applied.Top == RibbonBounds.Bottom;
        }
        return slot.Left == applied.Left && slot.Width == applied.Width && slot.Top == RibbonBounds.Bottom && slot.Bottom == applied.Top;
    }
    internal uint Dpi { get { return GetDpiForWindow(Frame); } }
    internal void CheckHitTargets(IntPtr strip, Rectangle slot)
    {
        // Inspect child stacking inside the host, even when another app is in front.
        Rectangle nav = NavigationBounds;
        Rectangle span = strip == IntPtr.Zero ? DocumentBounds : slot;
        int ribbonY = RibbonBounds.Bottom - 1;
        foreach (int x in new[] { span.Left + span.Width / 10, span.Left + span.Width / 2, span.Right - 1 - span.Width / 10 })
        {
            CheckChildAt(Frame, new Point(x, ribbonY), Ribbon, "ribbon");
            if (strip != IntPtr.Zero)
                CheckChildAt(Frame, new Point(x, slot.Top + slot.Height / 2), strip, "tab strip");
            if (AtBottom && strip != IntPtr.Zero)
                CheckChildAt(Frame, new Point(x, slot.Top - 1), Document, "document above bottom strip");
            Point p = new Point(x, nav.Top + nav.Height / 2);
            CheckChildAt(Frame, p, Document, "document at navigation row");
            MapWindowPoints(Frame, Document, ref p, 1);
            CheckChildAt(Document, p, navigation, "navigation");
        }
    }
    private static void CheckChildAt(IntPtr parent, Point p, IntPtr expected, string row)
    {
        const uint skipHiddenDisabledTransparent = 7;
        IntPtr hit = ChildWindowFromPointEx(parent, p, skipHiddenDisabledTransparent);
        if (hit != expected)
            throw new InvalidOperationException(String.Format("Hit test for {0} at ({1},{2}) in parent {3} reached {4}, expected {5}.",
                row, p.X, p.Y, parent, hit, expected));
    }
    internal bool Alive { get { return IsWindow(Frame) && GetParent(Document) == Frame &&
        GetParent(Ribbon) == Frame && GetParent(navigation) == Document; } }

    internal Rectangle Reserve(bool waitForEditor = false, bool? atBottom = null)
    {
        if (!Alive || !IsWindowEnabled(Frame) || !IsWindowVisible(Ribbon) || !IsWindowVisible(navigation))
            return Unavailable(waitForEditor, "OneNote replaced, disabled or hid a required layout window.");
        Rectangle body = DocumentBounds, ribbon = RibbonBounds;
        if (body.Top != ribbon.Bottom && (applied.IsEmpty || body.Top != applied.Top))
            return Unavailable(waitForEditor, "Unrecognized host layout. Cannot reserve a row.");
        Rectangle nav = NavigationBounds;
        if (ribbon.Left > body.Left || ribbon.Right < body.Right || nav.Top != body.Top || nav.Left != body.Left || nav.Width != body.Width)
            return Unavailable(waitForEditor, "Native controls no longer align with the document.");
        var middle = new Point(nav.Left + nav.Width / 2, nav.Top + nav.Height / 2);
        if (ChildWindowFromPointEx(Frame, middle, 7) != Document)
            return Unavailable(waitForEditor, "OneNote is covering the editor navigation.");
        MapWindowPoints(Frame, Document, ref middle, 1);
        if (ChildWindowFromPointEx(Document, middle, 7) != navigation)
            return Unavailable(waitForEditor, "OneNote is covering the navigation row in the document.");
        bool bottom = atBottom ?? AtBottom;
        var client = ClientBounds;
        // When the host has already laid out the full document, use its new edge. When our
        // previous bottom reservation is still in effect, put those pixels back before sizing again.
        int fullBottom = AtBottom && !applied.IsEmpty && body.Bottom == applied.Bottom ?
            FreeBottom(Rectangle.FromLTRB(body.Left, body.Bottom, body.Right, client.Bottom)) : body.Bottom;
        var full = Rectangle.FromLTRB(body.Left, ribbon.Bottom, body.Right, fullBottom);
        int height = HostChrome.StripHeight(Dpi);
        if (height <= 0 || body.Width < 120 || full.Height < height * 3)
            return Unavailable(waitForEditor, "Insufficient usable document space.");
        var slot = new Rectangle(full.Left, bottom ? full.Bottom - height : full.Top, full.Width, height);
        if (bottom && (full.Left != client.Left || full.Right != client.Right || full.Bottom != client.Bottom ||
            FreeBottom(slot) != slot.Bottom))
            return Unavailable(waitForEditor, "Another OneNote layout or control occupies the full-width bottom row.");
        var next = Rectangle.FromLTRB(full.Left, bottom ? full.Top : slot.Bottom, full.Right, bottom ? slot.Top : full.Bottom);
        if (next != body) Move(Document, next);
        original = full;
        applied = next;
        AtBottom = bottom;
        if (DocumentBounds != next || NavigationBounds.Top != next.Top || RibbonBounds != ribbon)
            return Unavailable(waitForEditor, "OneNote did not preserve the reserved row.");
        return slot;
    }
    private int FreeBottom(Rectangle band)
    {
        int limit = band.Bottom;
        if (band.Height <= 0) return limit;
        foreach (IntPtr child in Children(Frame, null))
        {
            if (child == Document || child == StripWindow) continue;
            var other = Bounds(child, Frame);
            if (other.IntersectsWith(band)) limit = Math.Min(limit, Math.Max(band.Top, other.Top));
        }
        return limit;
    }
    private static Rectangle Unavailable(bool waitForEditor, string reason)
    {
        if (waitForEditor) return Rectangle.Empty;
        throw new InvalidOperationException(reason);
    }
    internal void ReleaseForDialog()
    {
        if (applied.IsEmpty || !Alive || DocumentBounds != applied) return;
        var restore = original;
        if (AtBottom) restore.Height = RestoreBottom(applied) - restore.Top;
        Move(Document, restore);
        applied = Rectangle.Empty;
    }
    private int RestoreBottom(Rectangle body)
    {
        return FreeBottom(Rectangle.FromLTRB(body.Left, body.Bottom, body.Right, Math.Min(original.Bottom, ClientBounds.Bottom)));
    }

    public void Dispose()
    {
        if (!applied.IsEmpty && IsWindow(Frame) && GetParent(Document) == Frame)
        {
            Rectangle body = DocumentBounds;
            // Only release a strip we still own. Preserve any intervening host resize.
            if (AtBottom && body.Bottom == applied.Bottom)
                Move(Document, Rectangle.FromLTRB(body.Left, body.Top, body.Right, RestoreBottom(body)));
            else if (!AtBottom && body.Top == applied.Top)
            {
                int top = GetParent(Ribbon) == Frame ? RibbonBounds.Bottom : original.Top;
                if (top < body.Bottom) Move(Document, Rectangle.FromLTRB(body.Left, top, body.Right, body.Bottom));
            }
        }
        applied = Rectangle.Empty;
    }

    internal static Rectangle Bounds(IntPtr handle, IntPtr parent)
    {
        RECT rect;
        if (!GetWindowRect(handle, out rect)) throw new Win32Exception(Marshal.GetLastWin32Error());
        Point p = new Point(rect.Left, rect.Top);
        if (!ScreenToClient(parent, ref p)) throw new Win32Exception(Marshal.GetLastWin32Error());
        return new Rectangle(p.X, p.Y, rect.Right - rect.Left, rect.Bottom - rect.Top);
    }

    internal static void Move(IntPtr handle, Rectangle box)
    {
        if (!SetWindowPos(handle, IntPtr.Zero, box.X, box.Y, box.Width, box.Height, 0x0014))
            throw new Win32Exception(Marshal.GetLastWin32Error());
    }

    private static IEnumerable<IntPtr> Children(IntPtr parent, string className)
    {
        IntPtr child = IntPtr.Zero;
        while ((child = FindWindowEx(parent, child, className, null)) != IntPtr.Zero)
            if (IsWindowVisible(child)) yield return child;
    }

    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string name, string title);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool GetWindowRect(IntPtr window, out RECT bounds);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool GetClientRect(IntPtr window, out RECT bounds);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool ScreenToClient(IntPtr window, ref Point point);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] internal static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern IntPtr ChildWindowFromPointEx(IntPtr parent, Point point, uint flags);
    [DllImport("user32.dll")] private static extern int MapWindowPoints(IntPtr from, IntPtr to, ref Point point, uint count);
    [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr window);
    [DllImport("user32.dll", SetLastError=true)] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
}
