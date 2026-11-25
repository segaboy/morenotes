using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;

// Reproduce the disabled native class which extends six pixels below the ribbon on the live host.
internal sealed class RibbonShadowFixture : IDisposable
{
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WindowClass
    {
        public uint Style;
        public IntPtr Procedure;
        public int ClassBytes, WindowBytes;
        public IntPtr Instance, Icon, Cursor, Background;
        public string Menu, Name;
    }
    private readonly IntPtr handle;
    internal RibbonShadowFixture(IntPtr frame, Rectangle bounds)
    {
        var instance = GetModuleHandle(null);
        var wc = new WindowClass { Name = "DropShadow", Instance = instance,
            Procedure = GetProcAddress(GetModuleHandle("user32.dll"), "DefWindowProcW") };
        if (RegisterClass(ref wc) == 0 && Marshal.GetLastWin32Error() != 1410)
            throw new Win32Exception(Marshal.GetLastWin32Error());
        handle = CreateWindowEx(0, wc.Name, null, 0x58000000, bounds.X, bounds.Y, bounds.Width, bounds.Height,
            frame, IntPtr.Zero, instance, IntPtr.Zero);
        if (handle == IntPtr.Zero) throw new Win32Exception(Marshal.GetLastWin32Error());
    }
    public void Dispose() { DestroyWindow(handle); }
    [DllImport("kernel32", CharSet=CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
    [DllImport("kernel32", ExactSpelling=true)] private static extern IntPtr GetProcAddress(IntPtr module, string name);
    [DllImport("user32", CharSet=CharSet.Unicode, SetLastError=true)] private static extern ushort RegisterClass(ref WindowClass value);
    [DllImport("user32", CharSet=CharSet.Unicode, SetLastError=true)] private static extern IntPtr CreateWindowEx(uint ex, string cls,
        string title, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr data);
    [DllImport("user32")] private static extern bool DestroyWindow(IntPtr hwnd);
}
