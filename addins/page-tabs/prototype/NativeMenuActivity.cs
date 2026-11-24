using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

// Right-clicking a page can change OneNote's selection before the menu command runs.
// Defer ordinary tab-following until this editor's popup is dismissed.
internal sealed class NativeMenuActivity : IDisposable
{
    private const uint PopupStart = 6, PopupEnd = 7;
    private readonly HashSet<IntPtr> popups = new HashSet<IntPtr>();
    private readonly WinEvent callback;
    private readonly IntPtr hook, frame;
    private readonly uint thread;
    private DateTime resume;

    internal NativeMenuActivity(IntPtr frame)
    {
        this.frame = frame;
        uint process;
        thread = GetWindowThreadProcessId(frame, out process);
        callback = OnPopup;
        if (thread != 0) hook = SetWinEventHook(PopupStart, PopupEnd, IntPtr.Zero, callback, process, thread, 0);
    }
    private void OnPopup(IntPtr handle, uint kind, IntPtr window, int obj, int child, uint threadId, uint time)
    {
        if (kind == PopupStart) popups.Add(window);
        else { popups.Remove(window); resume = DateTime.UtcNow.AddMilliseconds(350); }
    }
    internal bool Pending
    {
        get
        {
            popups.RemoveWhere(p => !IsWindowVisible(p));
            return popups.Count > 0 || DateTime.UtcNow < resume || PageStrip.MenuInputActive(frame, thread);
        }
    }
    public void Dispose() { if (hook != IntPtr.Zero) UnhookWinEvent(hook); }
    private delegate void WinEvent(IntPtr hook, uint kind, IntPtr window, int obj, int child, uint thread, uint time);
    [DllImport("user32.dll")] private static extern IntPtr SetWinEventHook(uint first, uint last, IntPtr module, WinEvent callback, uint process, uint thread, uint flags);
    [DllImport("user32.dll")] private static extern bool UnhookWinEvent(IntPtr hook);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
