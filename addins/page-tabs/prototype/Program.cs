using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.ComponentModel;
using One = Microsoft.Office.Interop.OneNote;

internal static class Program
{
    private const int WaitForEditor = 10;
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            bool reconnect = false;
            for (;;)
            {
                int result = RunOnce(args, reconnect);
                if (result != WaitForEditor) return result;
                if (!reconnect) Console.WriteLine("Waiting for a compatible OneNote editor. Ctrl+C stops MoreTabs.");
                reconnect = true;
                // Drop interop wrappers from the old attachment before waiting for a new editor.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                while (NativeLayout.Discover(false).Count == 0) System.Threading.Thread.Sleep(750);
                System.Threading.Thread.Sleep(750);
            }
        }
        catch (Exception ex)
        {
            var native = ex as Win32Exception;
            if (native != null) Console.Error.WriteLine("Win32 error {0}: {1}", native.NativeErrorCode, native.Message);
            else Console.Error.WriteLine("{0}: {1} (0x{2:X8})", ex.GetType().Name, ex.Message, ex.HResult);
            return 1;
        }
    }

    private static int RunOnce(string[] args, bool reconnect)
    {
        string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "measure";
        if (mode == "selftest") { PrototypeTests.ControlledWindow(); SavedSessionTests.Run(); return 0; }
        if (mode == "sessiontest" && args.Length == 3) { SavedSessionTests.ProcessPhase(Guid.Parse(args[1]), args[2]); return 0; }
        if (mode == "com")
        {
            var app = new One.Application();
            var window = app.Windows.CurrentWindow;
            Console.WriteLine("OneNote COM window and page readable: " + (window != null && !String.IsNullOrEmpty(window.CurrentPageId)));
            return 0;
        }
        if (mode != "measure" && mode != "placement" && mode != "run" && mode != "verify" && mode != "handson" && mode != "bind")
            throw new ArgumentException("Use Measure, Placement, Com, Bind, Run, Verify, HandsOn or SelfTest.");
        IntPtr frame = !reconnect && args.Length == 2 ? new IntPtr(long.Parse(args[1])) : IntPtr.Zero;
        var layouts = NativeLayout.Discover(mode != "run" || frame != IntPtr.Zero);
        if (layouts.Count == 0) return WaitForEditor;
        foreach (var candidate in layouts)
        {
            NativeLayout.MatchDpi(candidate.Frame);
            Console.WriteLine("Frame {0}, DPI {1}; ribbon {2}; navigation {3}; document {4}", candidate.Frame,
                candidate.Dpi, candidate.RibbonBounds, candidate.NavigationBounds, candidate.DocumentBounds);
            if (mode == "measure")
            {
                var chrome = candidate.Appearance();
                Console.WriteLine("Strip appearance: surface #{0:X6}, bookmark icon inset {1}px ({2}).",
                    chrome.Surface.ToArgb() & 0xffffff, chrome.Inset, chrome.Measured ? "measured notebook icon" : "fallback");
                try
                {
                    candidate.CheckHitTargets(IntPtr.Zero, Rectangle.Empty);
                    Console.WriteLine("Native hit test matched ribbon and navigation at left, center and right.");
                }
                catch (InvalidOperationException error) { Console.WriteLine(error.Message); }
            }
        }
        if (mode == "measure") return 0;
        if (frame != IntPtr.Zero && !layouts.Exists(candidate => candidate.Frame == frame))
            throw new InvalidOperationException("Requested frame " + frame + " was not among the measured frames above.");
        NotebookPages source = null;
        string fixtureDirectory = mode == "verify" || mode == "handson" ?
            Path.Combine(Path.GetTempPath(), "MoreTabs-" + Guid.NewGuid().ToString("N")) : null;
        if (mode != "placement")
        {
            try { source = NotebookPages.Bind(layouts, frame, fixtureDirectory == null ? null : Path.Combine(fixtureDirectory, "MoreTabs B", "Test.one")); }
            catch (System.Runtime.InteropServices.COMException) { if (reconnect) return WaitForEditor; throw; }
            catch (InvalidOperationException) { if (reconnect && layouts.Count == 1) return WaitForEditor; throw; }
            frame = source.Frame;
            if (frame == IntPtr.Zero) throw new InvalidOperationException("The bound COM window closed.");
        }
        if (frame != IntPtr.Zero) layouts = layouts.FindAll(candidate => candidate.Frame == frame);
        if (layouts.Count != 1) throw new InvalidOperationException("Expected one candidate row. Select a frame with exactly one measured row using -WindowHandle.");
        using (var layout = layouts[0])
        {
            NativeLayout.MatchDpi(layout.Frame);
            Console.WriteLine("Selected measured frame {0}.", layout.Frame);
            if (mode == "bind") return 0;
            var initial = layout.DocumentBounds;
            Rectangle slot = layout.Reserve();
            Exception failure = null;
            Action<Exception> stop = delegate(Exception e) { failure = e; Application.ExitThread(); };
            bool closeRequested = false;
            Action close = delegate { closeRequested = true; Application.ExitThread(); };
            var store = mode == "run" ? SessionStore.ForUser() : null;
            using (var strip = new PageStrip(layout.Frame, source, close, stop, store))
            {
                strip.Attach(slot, layout);
                layout.CheckHitTargets(strip.Handle, slot);
                Console.WriteLine("Child strip {0}; navigation now starts at {1}", slot, layout.NavigationBounds.Top);
                if (mode == "placement" || mode == "verify" || mode == "handson")
                {
                    strip.Enabled = false;
                    // Observe without repositioning. Reapplying the reservation here would
                    // hide a OneNote layout pass which took the space back from the strip.
                    for (int i = 0; i < 12; i++)
                    {
                        Application.DoEvents();
                        System.Threading.Thread.Sleep(100);
                        if (!layout.StillReserved(slot) || NativeLayout.Bounds(strip.Handle, layout.Frame) != slot)
                            throw new InvalidOperationException("The host reclaimed or moved the strip during the idle placement test.");
                    }
                    strip.Enabled = true;
                    layout.CheckHitTargets(strip.Handle, slot);
                    Console.WriteLine("PASS reserved row survived the idle placement test with all three rows reachable.");
                }
                if (mode == "verify" || mode == "handson")
                {
                    PrototypeTests.DisposableNotebooks(source, strip, layout, slot, mode == "handson", fixtureDirectory);
                }
                else if (mode == "run")
                {
                    strip.RestoreOrSeed();
                    Console.WriteLine("Tabs restored. + creates a new Quick Note. Its arrow keeps the current page in a tab.");
                    bool detached = RunInteractive(layout, strip, null);
                    if (!closeRequested && failure == null && detached) return WaitForEditor;
                }
                if (failure != null) throw failure;
            }
            layout.Dispose();
            if (mode == "placement" && layout.DocumentBounds != initial)
                throw new InvalidOperationException("The final document bounds differ from the original measurement.");
            Console.WriteLine("Prototype closed and reserved row released.");
        }
        return 0;
    }

    internal static bool RunInteractive(NativeLayout layout, PageStrip strip, Action<Rectangle> observe)
    {
        int cancel = 0;
        bool detached = false;
        Exception failure = null;
        ConsoleCancelEventHandler cancelHandler = delegate(object sender, ConsoleCancelEventArgs e)
        {
            e.Cancel = true;
            System.Threading.Interlocked.Exchange(ref cancel, 1);
        };
        using (var timer = new Timer { Interval = 250 })
        {
            timer.Tick += delegate
            {
                try
                {
                    if (System.Threading.Volatile.Read(ref cancel) != 0) { Application.ExitThread(); return; }
                    if (!layout.Alive)
                    {
                        if (observe != null) throw new InvalidOperationException("OneNote closed or replaced the measured layout during the test.");
                        detached = true;
                        Application.ExitThread(); return;
                    }
                    strip.UpdateFocusShortcut();
                    if (NativeLayout.IsIconic(layout.Frame)) return;
                    Rectangle slot = layout.Reserve();
                    strip.Attach(slot, layout);
                    strip.RefreshActive();
                    if (observe != null) observe(slot);
                }
                catch (Exception e)
                {
                    bool serverClosed = e.HResult == unchecked((int)0x80010108) || e.HResult == unchecked((int)0x800706BA) ||
                        e.HResult == unchecked((int)0x800401FD);
                    if (observe == null && (!layout.Alive || serverClosed)) detached = true;
                    else failure = e;
                    Application.ExitThread();
                }
            };
            strip.UpdateFocusShortcut();
            timer.Start();
            Console.CancelKeyPress += cancelHandler;
            try { Application.Run(); }
            finally { Console.CancelKeyPress -= cancelHandler; }
        }
        if (failure != null) throw failure;
        return cancel == 0 && (detached || !layout.Alive);
    }
}
