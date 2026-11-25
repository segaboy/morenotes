using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;

// Experimental native window classes observed in OneNote build 16.0.20715.20004.
internal sealed class NativeLayout : IDisposable
{
    internal readonly IntPtr Frame;
    internal readonly IntPtr Document;
    internal readonly IntPtr Ribbon;
    private readonly IntPtr knownRow, knownWorkspace;
    private IntPtr landmark;
    private readonly HostChrome chrome = new HostChrome();
    internal IntPtr NavigationHeader { get; private set; }
    internal bool VerticalTabs { get; private set; }
    internal string WaitReason { get; private set; }
    internal HostChrome Appearance(bool force = false)
    {
        chrome.Update(NavigationHeader, Dpi, force, VerticalTabs);
        return chrome;
    }
    private Rectangle original, applied, editorSlot;
    private IntPtr filePane;
    private Rectangle fileFull, fileReduced;
    internal bool HasReservation { get { return !applied.IsEmpty; } }
    internal bool AtBottom { get; private set; }
    internal Rectangle RowBounds { get; private set; }
    internal IntPtr StripWindow;
    internal bool ShowCaptions = true;

    internal NativeLayout(IntPtr frame, IntPtr document, IntPtr ribbon, IntPtr navigationRow = default(IntPtr), IntPtr workspace = default(IntPtr))
    {
        Frame = frame;
        Document = document;
        Ribbon = ribbon;
        knownRow = landmark = navigationRow;
        knownWorkspace = workspace;
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
                    if (top.Height <= 0 || top.Bottom != body.Top || !OverlapsHorizontally(top, body)) continue;
                    var layout = new NativeLayout(frame, doc, dock);
                    if (layout.SelectLandmark(body)) result.Add(layout);
                }
            }
        }
        if (requireMatch && result.Count == 0) throw new InvalidOperationException(observed.Count == 0
            ? "Native discovery found no visible Framework::CFrame windows on this desktop."
            : "Native discovery found no adjacent ribbon and document with reachable navigation or page editor. Observed: " + String.Join("; ", observed));
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
    internal Rectangle LandmarkBounds { get { return Bounds(landmark, Frame); } }
    internal Rectangle RibbonBounds { get { return Bounds(Ribbon, Frame); } }
    internal Rectangle RibbonFaceBounds { get { return chrome.RibbonFace(Ribbon, RibbonBounds, Dpi); } }
    private static bool OverlapsHorizontally(Rectangle a, Rectangle b) { return Math.Max(a.Left, b.Left) < Math.Min(a.Right, b.Right); }
    internal Rectangle ClientBounds
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
        if (!Alive || slot.IsEmpty || DocumentBounds != applied || !SelectLandmark(applied)) return false;
        if (AtBottom)
        {
            var ribbon = RibbonFaceBounds;
            return slot.Left == ribbon.Left && slot.Right == ribbon.Right && slot.Bottom == ClientBounds.Bottom &&
                slot.Top == applied.Bottom && applied.Top == ribbon.Bottom;
        }
        return slot.Left == applied.Left && slot.Width == applied.Width && slot.Top == RibbonBounds.Bottom && slot.Bottom == applied.Top;
    }
    internal uint Dpi { get { return GetDpiForWindow(Frame); } }
    internal void CheckHitTargets(IntPtr strip, Rectangle slot)
    {
        // Inspect child stacking inside the host, even when another app is in front.
        Rectangle target = LandmarkBounds;
        var ribbon = RibbonBounds;
        CheckRow(ribbon, ribbon.Bottom - 1, Ribbon, "ribbon");
        if (strip != IntPtr.Zero)
        {
            CheckRow(slot, slot.Top + slot.Height / 2, strip, "tab strip");
            if (AtBottom) CheckRow(DocumentBounds, slot.Top - 1, Document, "document above bottom strip");
        }
        foreach (var point in SamplePoints(target, VerticalTabs, Dpi))
        {
            string name = VerticalTabs ? "page editor" : "navigation";
            CheckChildAt(Frame, point, Document, "document at " + name);
            var local = point;
            MapWindowPoints(Frame, Document, ref local, 1);
            CheckChildAt(Document, local, landmark, name);
        }
    }
    private void CheckRow(Rectangle span, int y, IntPtr expected, string row)
    {
        foreach (int x in new[] { span.Left + span.Width / 10, span.Left + span.Width / 2, span.Right - 1 - span.Width / 10 })
            CheckChildAt(Frame, new Point(x, y), expected, row);
    }
    private static void CheckChildAt(IntPtr parent, Point p, IntPtr expected, string row)
    {
        const uint skipHiddenDisabledTransparent = 7;
        IntPtr hit = ChildWindowFromPointEx(parent, p, skipHiddenDisabledTransparent);
        if (hit != expected)
            throw new InvalidOperationException(String.Format("Hit test for {0} at ({1},{2}) in parent {3} reached {4}, expected {5}.",
                row, p.X, p.Y, parent, hit, expected));
    }
    internal bool Alive { get { return IsWindow(Frame) && GetParent(Document) == Frame && GetParent(Ribbon) == Frame; } }
    internal bool FrameEnabled { get { return IsWindowEnabled(Frame); } }

    // Backstage can put a frame-sized child in front of the page. Make physical space
    // at its edge, and give that space back only if our dimensions are still in effect.
    internal void ReleaseFilePane()
    {
        if (filePane != IntPtr.Zero && GetParent(filePane) == Frame && Bounds(filePane, Frame) == fileReduced)
            Move(filePane, fileFull);
        filePane = IntPtr.Zero;
    }

    internal bool HasFilePane { get { return filePane != IntPtr.Zero; } }
    internal bool HoldFileRow(out Rectangle row, out Rectangle slot, bool paneOnly = false)
    {
        row = slot = Rectangle.Empty;
        if (!IsWindow(Frame) || applied.IsEmpty || editorSlot.IsEmpty) return false;
        var client = ClientBounds;
        var editor = Rectangle.Intersect(original, client);
        if (editor.IsEmpty) return false;
        var mid = new Point(editor.Left + editor.Width / 2, editor.Top + editor.Height / 2);
        // Include disabled children: Options owns input while this frame is dimmed.
        var front = ChildWindowFromPointEx(Frame, mid, 5);
        if (filePane != IntPtr.Zero && front != filePane) ReleaseFilePane();
        if (front != Document && FrameEnabled && HasClass(front, "OneNote::MainWindow"))
        {
            ReleaseFilePane();
            return false; // An ordinary replacement editor must go through discovery and COM binding.
        }
        int height = HostChrome.StripHeight(Dpi, ShowCaptions, AtBottom);
        var pane = front != IntPtr.Zero && front != Frame ? Bounds(front, Frame) : Rectangle.Empty;
        var full = front == filePane && pane == fileReduced ? fileFull : pane;
        int inset = HostChrome.Px(12, Dpi);
        bool replacesPage = front != Document || (!IsWindowVisible(Ribbon) && full.Top < original.Top);
        bool coversEditor = front != Frame && front != StripWindow && replacesPage &&
            full.Left <= editor.Left + inset && full.Right >= editor.Right - inset && full.Top <= editor.Top + inset &&
            full.Bottom >= editor.Bottom - inset && full.Height >= height * 3;
        if (coversEditor)
        {
            var reduced = full;
            reduced.Height -= height;
            if (!AtBottom) reduced.Y += height;
            row = Rectangle.Intersect(client, new Rectangle(full.Left, AtBottom ? reduced.Bottom : full.Top, full.Width, height));
            slot = Rectangle.Intersect(row, Rectangle.FromLTRB(editorSlot.Left, row.Top, editorSlot.Right, row.Bottom));
            if (row.Height != height || slot.Width < 120) return false;
            if (pane != reduced) Move(front, reduced);
            filePane = front; fileFull = full; fileReduced = reduced;
            return Bounds(front, Frame) == reduced;
        }
        ReleaseFilePane();
        if (paneOnly) return false;
        // The dialog may leave the original editor visible, or hide its navigation.
        // Keep its reserved row even when OneNote has just laid the full document out again.
        bool ownsDocument = GetParent(Document) == Frame;
        var body = ownsDocument ? DocumentBounds : Rectangle.Empty;
        if (ownsDocument && body != applied && body != original)
        {
            // Windows may resize the owner as the modal dialog opens (for example at a new DPI).
            // Recompute the slot from this frame instead of discarding both old rectangles.
            if (FrameEnabled || !Alive) return false;
            slot = Reserve(true, resizing: true, modal: true);
            if (slot.IsEmpty) return false;
            row = RowBounds;
            return true;
        }
        if (!client.Contains(editorSlot) || FreeBottom(editorSlot) < editorSlot.Bottom) return false;
        if (ownsDocument && body != applied) Move(Document, applied);
        row = RowBounds;
        slot = editorSlot;
        return true;
    }

    private static IEnumerable<Point> SamplePoints(Rectangle area, bool editor, uint dpi)
    {
        int y = area.Top + (editor ? Math.Min(area.Height / 2, HostChrome.Px(24, dpi)) : area.Height / 2);
        yield return new Point(area.Left + area.Width / 10, y);
        yield return new Point(area.Left + area.Width / 2, y);
        yield return new Point(area.Right - 1 - area.Width / 10, y);
    }

    private bool TryUseLandmark(IntPtr child, Rectangle area, bool editor, uint dpi)
    {
        foreach (var point in SamplePoints(area, editor, dpi))
        {
            // During a resize or dock change our existing strip can still occupy the old slot.
            var hit = ChildWindowFromPointEx(Frame, point, 7);
            if (hit != Document && (StripWindow == IntPtr.Zero || hit != StripWindow)) return false;
            var local = point;
            MapWindowPoints(Frame, Document, ref local, 1);
            if (ChildWindowFromPointEx(Document, local, 7) != child) return false;
        }
        VerticalTabs = editor;
        landmark = child;
        return true;
    }

    private bool SelectLandmark(Rectangle body)
    {
        uint dpi = Dpi;
        var rows = knownRow == IntPtr.Zero ? Children(Document, "OneNote::NavigationUI") : new[] { knownRow };
        int rowBottom = body.Top;
        NavigationHeader = IntPtr.Zero;
        foreach (var row in rows)
        {
            if (!IsWindowVisible(row) || GetParent(row) != Document) continue;
            var area = Bounds(row, Frame);
            if (area.Top != body.Top || area.Left != body.Left ||
                area.Height < 12 || area.Height > dpi || area.Height >= body.Height) continue;
            // In vertical view the editor can cover most of this row, leaving its notebook header at the left.
            if (ChildWindowFromPointEx(Document, new Point(HostChrome.Px(6, dpi), HostChrome.Px(6, dpi)), 7) == row)
                NavigationHeader = row;
            if (area.Width != body.Width) continue;
            rowBottom = Math.Max(rowBottom, area.Bottom);
            if (TryUseLandmark(row, area, false, dpi))
            {
                NavigationHeader = row;
                return true;
            }
        }
        // Vertical Tabs can leave a full-width navigation HWND behind the actual panes.
        // Only use the page workspace when it occupies that upper band and receives the hits.
        var workspace = knownWorkspace != IntPtr.Zero ? knownWorkspace : FindWindowEx(Document, IntPtr.Zero, "OneNote::CWorkspace", null);
        if (!IsWindowVisible(workspace) || GetParent(workspace) != Document ||
            (knownWorkspace == IntPtr.Zero && !IsWindowVisible(FindWindowEx(workspace, IntPtr.Zero, "OneNote::DocumentCanvas", null)))) return false;
        var editor = Rectangle.Intersect(body, Bounds(workspace, Frame));
        return editor.Width >= HostChrome.Px(48, dpi) && editor.Height >= HostChrome.Px(48, dpi) &&
            editor.Top - body.Top <= dpi && (rowBottom == body.Top || editor.Top < rowBottom) &&
            TryUseLandmark(workspace, editor, true, dpi);
    }

    internal Rectangle Reserve(bool waitForEditor = false, bool? atBottom = null, bool resizing = false, bool modal = false)
    {
        // A previously verified editor can briefly hide its children while processing WM_SIZE.
        // Only use that reservation for the same native document and an enabled frame.
        bool follow = resizing && HasReservation;
        if (!Alive || (!FrameEnabled && !modal) || (!follow && (!IsWindowVisible(Ribbon) || !IsWindowVisible(Document))))
            return Unavailable(waitForEditor, "OneNote replaced, disabled or hid a required layout window.");
        Rectangle body = DocumentBounds, ribbon = RibbonBounds;
        if (body.Top != ribbon.Bottom && (applied.IsEmpty || body.Top != applied.Top))
            return Unavailable(waitForEditor, "Unrecognized host layout. Cannot reserve a row.");
        // OneNote sizes the frame, document and navigation children in separate passes.
        if (!OverlapsHorizontally(ribbon, body) || (!follow && !SelectLandmark(body)))
            return Unavailable(waitForEditor, "OneNote's navigation row or vertical page editor is hidden, covered or no longer aligned with the ribbon.");
        bool bottom = atBottom ?? AtBottom;
        var client = ClientBounds;
        // When the host has already laid out the full document, use its new edge. When our
        // previous bottom reservation is still in effect, put those pixels back before sizing again.
        int fullBottom = body.Bottom;
        if (follow && bottom) fullBottom = client.Bottom;
        else if (AtBottom && !applied.IsEmpty && body.Bottom == applied.Bottom)
            fullBottom = bottom ? client.Bottom : FreeBottom(Rectangle.FromLTRB(body.Left, body.Bottom, body.Right, client.Bottom));
        var full = Rectangle.FromLTRB(body.Left, ribbon.Bottom, body.Right, fullBottom);
        int height = HostChrome.StripHeight(Dpi, ShowCaptions, bottom);
        if (height <= 0 || body.Width < 120 || full.Height < height * 3)
            return Unavailable(waitForEditor, "Insufficient usable document space.");
        var face = bottom ? chrome.RibbonFace(Ribbon, ribbon, Dpi, !resizing) : ribbon;
        if (follow && bottom) face = Rectangle.Intersect(face, client);
        var slot = bottom ? new Rectangle(face.Left, full.Bottom - height, face.Width, height) :
            new Rectangle(full.Left, full.Top, full.Width, height);
        var row = bottom ? FreeBottomRow(slot, client) : slot;
        if (bottom && (slot.Width < 120 || !client.Contains(slot) || slot.Bottom != client.Bottom || row.IsEmpty))
            return Unavailable(waitForEditor, "OneNote does not have a free bottom row spanning the ribbon.");
        if (follow && !bottom && FreeBottom(slot) != slot.Bottom)
            return Unavailable(waitForEditor, "Another OneNote control occupies the top row during resize.");
        var next = Rectangle.FromLTRB(full.Left, bottom ? full.Top : slot.Bottom, full.Right, bottom ? slot.Top : full.Bottom);
        if (bottom && next.Bottom > body.Bottom && FreeBottom(Rectangle.FromLTRB(body.Left, body.Bottom, body.Right, next.Bottom)) != next.Bottom)
            return Unavailable(waitForEditor, "Another OneNote control occupies the space below the document.");
        if (next != body) Move(Document, next);
        original = full;
        applied = next;
        AtBottom = bottom;
        if (DocumentBounds != next || (!follow && !SelectLandmark(next)) || RibbonBounds != ribbon)
            return Unavailable(waitForEditor, "OneNote did not preserve the reserved row.");
        RowBounds = row;
        editorSlot = slot;
        WaitReason = null;
        return slot;
    }
    // Extend the background into the free gutters. A native pane can own part of either gutter.
    private Rectangle FreeBottomRow(Rectangle slot, Rectangle client)
    {
        var row = Rectangle.FromLTRB(client.Left, slot.Top, client.Right, slot.Bottom);
        int left = row.Left, right = row.Right;
        foreach (var other in OccupiedFrameBounds())
        {
            if (!other.IntersectsWith(row)) continue;
            if (other.Right <= slot.Left) left = Math.Max(left, other.Right);
            else if (other.Left >= slot.Right) right = Math.Min(right, other.Left);
            else return Rectangle.Empty;
        }
        return Rectangle.FromLTRB(left, row.Top, right, row.Bottom);
    }
    internal void KeepStripVisible(Rectangle slot)
    {
        // A host layout pass may have hidden our HWND or raised its document over the strip.
        var hit = ChildWindowFromPointEx(Frame, new Point(slot.Left + slot.Width / 2, slot.Top + slot.Height / 2), 7);
        if (IsWindowVisible(StripWindow) && hit == StripWindow) return;
        if (hit != Document && hit != Frame && hit != StripWindow) return;
        ShowStripWithoutActivating();
    }
    internal bool FileRowIsVisible(Rectangle slot)
    {
        var hit = ChildWindowFromPointEx(Frame, new Point(slot.Left + slot.Width / 2, slot.Top + slot.Height / 2), 5);
        return IsWindowVisible(StripWindow) && hit == StripWindow;
    }
    internal void ShowStripWithoutActivating()
    {
        const uint showWithoutMovingOrActivating = 0x0040 | 0x0001 | 0x0002 | 0x0010;
        if (!SetWindowPos(StripWindow, IntPtr.Zero, 0, 0, 0, 0, showWithoutMovingOrActivating))
            throw new Win32Exception(Marshal.GetLastWin32Error());
    }
    private int FreeBottom(Rectangle band)
    {
        int limit = band.Bottom;
        if (band.Height <= 0) return limit;
        foreach (var other in OccupiedFrameBounds())
            if (other.IntersectsWith(band)) limit = Math.Min(limit, Math.Max(band.Top, other.Top));
        return limit;
    }
    private IEnumerable<Rectangle> OccupiedFrameBounds()
    {
        foreach (IntPtr child in Children(Frame, null))
        {
            if (child == Document || child == StripWindow) continue;
            // OneNote's disabled DropShadow straddles the lower ribbon edge, including
            // the first pixels of our top slot. It is chrome, not an occupied control.
            if (!IsWindowEnabled(child) && HasClass(child, "DropShadow")) continue;
            var other = Bounds(child, Frame);
            if (other.Width > 0 && other.Height > 0) yield return other;
        }
    }
    private static bool HasClass(IntPtr child, string expected)
    {
        var name = new StringBuilder(64);
        return GetClassName(child, name, name.Capacity) > 0 && name.ToString() == expected;
    }
    private Rectangle Unavailable(bool waitForEditor, string reason)
    {
        WaitReason = reason;
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
        ReleaseFilePane();
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
        {
            var error = new Win32Exception(Marshal.GetLastWin32Error());
            throw new Win32Exception(error.NativeErrorCode, "Could not move window " + handle + " to " + box + ". " + error.Message);
        }
    }

    private static IEnumerable<IntPtr> Children(IntPtr parent, string className)
    {
        IntPtr child = IntPtr.Zero;
        while ((child = FindWindowEx(parent, child, className, null)) != IntPtr.Zero)
            if (IsWindowVisible(child)) yield return child;
    }

    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
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
