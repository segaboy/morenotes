using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;

// Shared with the OneNote page-list menu. The receiving window belongs to one editor frame.
internal static class PageMenuCommand
{
    internal const string ProgId = "MoreNotes.MoreTabs.PageMenu";
    internal const string HostName = "MoreTabs page menu target";
    internal const string ReceiverName = "MoreTabs page menu receiver v2";
    internal const int CopyData = 0x004A;
    internal const int OpenFromMenu = 0x8000 + 0x04D5; // WM_APP, handled only by the strip host.
    private const int OpenPage = 0x4D540001;
    private const int MaxChars = 2048;

    internal static string SelectedPage(XElement section, string currentPage)
    {
        if (section.Name.LocalName != "Section") return null;
        var pages = section.Elements(section.Name.Namespace + "Page");
        var selected = pages.Where(p => (string)p.Attribute("selected") == "all").Take(2).ToArray();
        if (selected.Length == 0) selected = pages.Where(p => (string)p.Attribute("selected") == "partial").Take(2).ToArray();
        if (selected.Length != 0) return selected.Length == 1 ? (string)selected[0].Attribute("ID") : null;
        return pages.Any(p => (string)p.Attribute("ID") == currentPage &&
            (string)p.Attribute("isCurrentlyViewed") == "true") ? currentPage : null;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Packet { internal IntPtr Command; internal int Bytes; internal IntPtr Data; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string name);
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);

    internal static IntPtr ForegroundEditor()
    {
        var foreground = GetForegroundWindow();
        var root = GetAncestor(foreground, 2); // GA_ROOT
        if (root != IntPtr.Zero && FindWindowEx(root, IntPtr.Zero, "OneNote::MainWindow", null) != IntPtr.Zero) return root;
        var owner = GetAncestor(foreground, 3); // GA_ROOTOWNER for a native popup.
        return owner != IntPtr.Zero && FindWindowEx(owner, IntPtr.Zero, "OneNote::MainWindow", null) != IntPtr.Zero ? owner : IntPtr.Zero;
    }

    internal static bool Request(IntPtr window, bool pageList)
    {
        var frame = GetAncestor(window, 2);
        if (frame == IntPtr.Zero) return false;
        var host = FindWindowEx(frame, IntPtr.Zero, null, HostName);
        if (host == IntPtr.Zero || !IsWindowVisible(host) || !IsWindowEnabled(host)) return false;
        // Only updated strips expose this receiver. Never wait on a UI thread that may be calling OneNote.
        var target = FindWindowEx(host, IntPtr.Zero, null, ReceiverName);
        return target != IntPtr.Zero && PostMessage(target, OpenFromMenu, pageList ? (IntPtr)1 : IntPtr.Zero, IntPtr.Zero);
    }

    // Older helpers can stay loaded in OneNote until it restarts and still send an explicit page ID.
    internal static string Read(Message message)
    {
        if (message.Msg != CopyData || message.LParam == IntPtr.Zero) return null;
        var packet = (Packet)Marshal.PtrToStructure(message.LParam, typeof(Packet));
        if (packet.Command != new IntPtr(OpenPage) || packet.Data == IntPtr.Zero ||
            packet.Bytes <= 0 || packet.Bytes > MaxChars * 2 || packet.Bytes % 2 != 0) return null;
        var id = Marshal.PtrToStringUni(packet.Data, packet.Bytes / 2);
        return id.Any(Char.IsControl) ? null : id;
    }
}
