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
    private Rectangle original, applied;

    internal NativeLayout(IntPtr frame, IntPtr document, IntPtr ribbon, IntPtr navigationRow)
    {
        Frame = frame;
        Document = document;
        Ribbon = ribbon;
        navigation = navigationRow;
    }

    internal static List<NativeLayout> Discover(IntPtr preferred)
    {
        var result = new List<NativeLayout>();
        foreach (IntPtr frame in Children(IntPtr.Zero, "Framework::CFrame"))
        {
            if (IsIconic(frame) || (preferred != IntPtr.Zero && frame != preferred)) continue;
            MatchDpi(frame);
            foreach (IntPtr doc in Children(frame, "OneNote::MainWindow"))
            {
                Rectangle body = Bounds(doc, frame);
                foreach (IntPtr dock in Children(frame, "MsoCommandBarDock"))
                {
                    Rectangle top = Bounds(dock, frame);
                    if (top.Bottom != body.Top || top.Left > body.Left || top.Right < body.Right) continue;
                    foreach (IntPtr nav in Children(doc, "OneNote::NavigationUI"))
                    {
                        Rectangle row = Bounds(nav, frame);
                        if (row.Top != body.Top || row.Width != body.Width ||
                            row.Height < 12 || row.Height > GetDpiForWindow(frame) ||
                            row.Height >= body.Height) continue;
                        result.Add(new NativeLayout(frame, doc, dock, nav));
                    }
                }
            }
        }
        if (result.Count == 0) throw new InvalidOperationException("Could not identify the ribbon, document and horizontal navigation row. Open a normal OneNote window, expand the ribbon, and enable horizontal tabs.");
        return result;
    }

    internal static void MatchDpi(IntPtr frame)
    {
        if (SetThreadDpiAwarenessContext(GetWindowDpiAwarenessContext(frame)) == IntPtr.Zero)
            throw new Win32Exception(Marshal.GetLastWin32Error(), "Could not match the host's DPI context.");
    }

    internal Rectangle DocumentBounds { get { return Bounds(Document, Frame); } }
    internal Rectangle NavigationBounds { get { return Bounds(navigation, Frame); } }
    internal Rectangle RibbonBounds { get { return Bounds(Ribbon, Frame); } }
    internal bool StillReserved(Rectangle slot)
    {
        return Alive && slot.Left == DocumentBounds.Left && slot.Width == DocumentBounds.Width &&
            slot.Top == RibbonBounds.Bottom && slot.Bottom == NavigationBounds.Top &&
            slot.Bottom == DocumentBounds.Top && DocumentBounds == applied;
    }
    internal uint Dpi { get { return GetDpiForWindow(Frame); } }
    internal void CheckHitTargets(IntPtr strip, Rectangle slot)
    {
        // Use host-relative hit testing so a terminal on top of OneNote cannot
        // invalidate the measurement. This checks sibling stacking, not user input.
        int x = slot.Left + slot.Width / 2;
        Rectangle nav = NavigationBounds;
        const uint skipHiddenDisabledTransparent = 7;
        if (ChildWindowFromPointEx(Frame, new Point(x, slot.Top - 1), skipHiddenDisabledTransparent) != Ribbon ||
            ChildWindowFromPointEx(Frame, new Point(x, slot.Top + slot.Height / 2), skipHiddenDisabledTransparent) != strip ||
            ChildWindowFromPointEx(Frame, new Point(x, nav.Top + nav.Height / 2), skipHiddenDisabledTransparent) != Document)
            throw new InvalidOperationException("Ribbon, strip or navigation centre is obscured in the host's child-window stack.");
    }
    internal bool Alive { get { return IsWindow(Frame) && GetParent(Document) == Frame &&
        GetParent(Ribbon) == Frame && GetParent(navigation) == Document; } }

    // The host owns the outer edges. Recompute from the ribbon and current bottom each
    // time so repeated ticks cannot shrink the document again.
    internal Rectangle Reserve()
    {
        if (!Alive || !IsWindowVisible(Ribbon) || !IsWindowVisible(navigation))
            throw new InvalidOperationException("OneNote replaced or hid a required layout window.");
        Rectangle body = DocumentBounds, ribbon = RibbonBounds;
        if (body.Top != ribbon.Bottom && (applied.IsEmpty || body.Top != applied.Top))
            throw new InvalidOperationException("Unrecognized host layout. Cannot reserve a row.");
        if (ribbon.Left > body.Left || ribbon.Right < body.Right || NavigationBounds.Top != body.Top)
            throw new InvalidOperationException("Native controls no longer align with the document.");
        int height = (int)Math.Round(32 * Dpi / 96.0);
        if (height <= 0 || body.Width < 120 || body.Bottom - ribbon.Bottom < height * 3)
            throw new InvalidOperationException("Insufficient usable document space.");
        Rectangle full = Rectangle.FromLTRB(body.Left, ribbon.Bottom, body.Right, body.Bottom);
        Rectangle next = Rectangle.FromLTRB(full.Left, full.Top + height, full.Right, full.Bottom);
        if (next != body) Move(Document, next);
        original = full;
        applied = next;
        if (DocumentBounds != next || NavigationBounds.Top != next.Top || RibbonBounds != ribbon)
            throw new InvalidOperationException("OneNote did not preserve the reserved row.");
        return new Rectangle(full.Left, full.Top, full.Width, height);
    }

    public void Dispose()
    {
        if (!applied.IsEmpty && IsWindow(Frame) && GetParent(Document) == Frame)
        {
            Rectangle body = DocumentBounds;
            // Only release a strip we still own. Preserve any intervening host resize.
            if (body.Top == applied.Top)
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
    [DllImport("user32.dll", SetLastError=true)] private static extern bool ScreenToClient(IntPtr window, ref Point point);
    [DllImport("user32.dll", SetLastError=true)] private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] internal static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] internal static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr ChildWindowFromPointEx(IntPtr parent, Point point, uint flags);
    [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr window);
    [DllImport("user32.dll", SetLastError=true)] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
}
