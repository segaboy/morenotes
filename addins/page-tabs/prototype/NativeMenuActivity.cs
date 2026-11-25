using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// Right-clicking a page can change OneNote's selection before the menu command runs.
// Defer ordinary tab-following until this editor's popup is dismissed.
internal sealed class NativeMenuActivity : NativeWindow, IDisposable
{
    private const uint PopupStart = 6, PopupEnd = 7;
    private readonly HashSet<IntPtr> popups = new HashSet<IntPtr>();
    private readonly WinEvent callback;
    private readonly IntPtr hook, frame, document;
    private readonly Action<int> markGesture;
    private readonly bool mouseRegistered;
    private int? secondaryPress;
    private readonly uint thread;
    private DateTime resume;

    internal NativeMenuActivity(NativeLayout layout, Action<int> markGesture)
    {
        frame = layout.Frame;
        document = layout.Document;
        this.markGesture = markGesture;
        uint process;
        thread = GetWindowThreadProcessId(frame, out process);
        callback = OnPopup;
        if (thread != 0) hook = SetWinEventHook(PopupStart, PopupEnd, IntPtr.Zero, callback, process, thread, 0);
        // Observe the secondary press before OneNote switches pages. Its custom popup can appear later.
        CreateHandle(new CreateParams { Parent = new IntPtr(-3) }); // HWND_MESSAGE
        mouseRegistered = RegisterMouse(0x100, Handle); // RIDEV_INPUTSINK
    }
    private void OnPopup(IntPtr handle, uint kind, IntPtr window, int obj, int child, uint threadId, uint time)
    {
        if (kind == PopupStart)
        {
            if (popups.Count == 0 && GetForegroundWindow() == frame)
            {
                // Pointer input has the earlier timestamp. A keyboard-invoked menu starts a new gesture.
                if (!secondaryPress.HasValue || unchecked(time - (uint)secondaryPress.Value) > SystemInformation.DoubleClickTime)
                    markGesture(unchecked((int)time));
                secondaryPress = null;
            }
            popups.Add(window);
        }
        else
        {
            popups.Remove(window);
            resume = DateTime.UtcNow.AddMilliseconds(350);
        }
    }
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == 0x00ff && GetForegroundWindow() == frame) // WM_INPUT
        {
            MouseInput mouse;
            uint size = MouseSize;
            if (GetRawInputData(m.LParam, 0x10000003, out mouse, ref size, (uint)(8 + 2 * IntPtr.Size)) == MouseSize &&
                mouse.Type == 0 && (mouse.ButtonFlags & 5) != 0 &&
                (mouse.ButtonFlags & (SystemInformation.MouseButtonsSwapped ? 1 : 4)) != 0)
            {
                var hit = WindowFromPoint(new Point(GetMessagePos()));
                if (hit == document || IsChild(document, hit))
                {
                    secondaryPress = GetMessageTime();
                    markGesture(secondaryPress.Value);
                }
            }
        }
        base.WndProc(ref m);
    }
    internal bool Pending
    {
        get
        {
            popups.RemoveWhere(p => !IsWindowVisible(p));
            return popups.Count > 0 || DateTime.UtcNow < resume || PageStrip.MenuInputActive(frame, thread);
        }
    }
    public void Dispose()
    {
        if (mouseRegistered) RegisterMouse(1, IntPtr.Zero); // RIDEV_REMOVE
        if (hook != IntPtr.Zero) UnhookWinEvent(hook);
        DestroyHandle();
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct RawDevice { public ushort Page, Usage; public uint Flags; public IntPtr Target; }
    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput
    {
        public uint Type, Size;
        public IntPtr Device, WParam;
        public ushort Flags, Padding, ButtonFlags, ButtonData;
        public uint RawButtons;
        public int LastX, LastY;
        public uint Extra;
    }
    private static readonly uint MouseSize = (uint)Marshal.SizeOf(typeof(MouseInput));
    private static bool RegisterMouse(uint flags, IntPtr target)
    {
        var mouse = new RawDevice { Page = 1, Usage = 2, Flags = flags, Target = target };
        return RegisterRawInputDevices(ref mouse, 1, (uint)Marshal.SizeOf(typeof(RawDevice)));
    }
    [DllImport("user32.dll")] private static extern bool RegisterRawInputDevices(ref RawDevice device, uint count, uint size);
    [DllImport("user32.dll")] private static extern uint GetRawInputData(IntPtr input, uint command, out MouseInput data, ref uint size, uint headerSize);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern int GetMessagePos();
    [DllImport("user32.dll")] private static extern int GetMessageTime();
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point p);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    private delegate void WinEvent(IntPtr hook, uint kind, IntPtr window, int obj, int child, uint thread, uint time);
    [DllImport("user32.dll")] private static extern IntPtr SetWinEventHook(uint first, uint last, IntPtr module, WinEvent callback, uint process, uint thread, uint flags);
    [DllImport("user32.dll")] private static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
