using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;

// Shared with the OneNote page-list menu. The receiving window belongs to one editor frame.
internal static class PageMenuCommand
{
    internal const string HostName = "MoreTabs page menu target";
    internal const int CopyData = 0x004A;
    private const int OpenPage = 0x4D540001;
    private const int MaxChars = 2048;

    internal static string SelectedPage(XElement section)
    {
        if (section.Name.LocalName != "Section") return null;
        var pages = section.Elements(section.Name.Namespace + "Page");
        var selected = pages.Where(p => (string)p.Attribute("selected") == "all").Take(2).ToArray();
        if (selected.Length != 1) return null;
        return (string)selected[0].Attribute("ID");
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct Packet { internal IntPtr Command; internal int Bytes; internal IntPtr Data; }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string name);
    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")]
    private static extern IntPtr SendMessageTimeout(IntPtr window, int message, IntPtr sender, ref Packet packet,
        uint flags, uint timeout, out IntPtr result);

    internal static bool Send(IntPtr oneNoteWindow, string pageId)
    {
        if (String.IsNullOrEmpty(pageId) || pageId.Length > MaxChars || oneNoteWindow == IntPtr.Zero) return false;
        var frame = GetAncestor(oneNoteWindow, 2);
        if (frame == IntPtr.Zero) return false;
        var target = FindWindowEx(frame, IntPtr.Zero, null, HostName);
        if (target == IntPtr.Zero) return false;
        var packet = new Packet { Command = new IntPtr(OpenPage), Data = Marshal.StringToHGlobalUni(pageId), Bytes = 2 * pageId.Length };
        try
        {
            IntPtr result;
            return SendMessageTimeout(target, CopyData, IntPtr.Zero, ref packet, 2, 1200, out result) != IntPtr.Zero && result != IntPtr.Zero;
        }
        finally { Marshal.FreeHGlobal(packet.Data); }
    }

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
