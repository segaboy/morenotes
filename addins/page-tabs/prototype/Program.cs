using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            string mode = args.Length > 0 ? args[0].ToLowerInvariant() : "measure";
            if (mode == "selftest") { PrototypeTests.ControlledWindow(); return 0; }
            if (mode == "com") { var pages = new NotebookPages(); Console.WriteLine("OneNote COM window and page readable: " + !String.IsNullOrEmpty(pages.Current)); return 0; }
            if (mode != "measure" && mode != "placement" && mode != "run" && mode != "verify")
                throw new ArgumentException("Use Measure, Placement, Com, Run, Verify or SelfTest.");
            NotebookPages source = mode == "run" || mode == "verify" ? new NotebookPages() : null;
            IntPtr frame = source == null ? IntPtr.Zero : new IntPtr(unchecked((long)source.Window.WindowHandle));
            if (args.Length == 2 && mode == "placement") frame = new IntPtr(long.Parse(args[1]));
            var layouts = NativeLayout.Discover(frame);
            if (mode == "measure")
            {
                foreach (var candidate in layouts)
                {
                    NativeLayout.MatchDpi(candidate.Frame);
                    Console.WriteLine("Frame {0}, DPI {1}; ribbon {2}; navigation {3}; document {4}", candidate.Frame,
                        candidate.Dpi, candidate.RibbonBounds, candidate.NavigationBounds, candidate.DocumentBounds);
                }
                return 0;
            }
            if (layouts.Count != 1) throw new InvalidOperationException("Multiple candidate rows. For Placement, use a frame handle from Measure with -WindowHandle.");
            using (var layout = layouts[0])
            {
                NativeLayout.MatchDpi(layout.Frame);
                Console.WriteLine("DPI {0}; ribbon {1}; navigation {2}; document {3}",
                    layout.Dpi, layout.RibbonBounds, layout.NavigationBounds, layout.DocumentBounds);
                var initial = layout.DocumentBounds;
                Rectangle slot = layout.Reserve();
                Exception failure = null;
                Action<Exception> stop = delegate(Exception e) { failure = e; Application.ExitThread(); };
                using (var strip = new PageStrip(layout.Frame, source, Application.ExitThread, stop))
                {
                    strip.Attach(slot);
                    layout.CheckHitTargets(strip.Handle, slot);
                    Console.WriteLine("Child strip {0}; navigation now starts at {1}", slot, layout.NavigationBounds.Top);
                    if (mode == "placement" || mode == "verify")
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
                        Console.WriteLine("PASS reserved row survived the idle placement test.");
                    }
                    if (mode == "verify")
                    {
                        PrototypeTests.DisposableNotebooks(source, strip, layout, slot);
                    }
                    else if (mode == "run")
                    {
                        strip.AddCurrent();
                        int cancel = 0;
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
                                    if (System.Threading.Volatile.Read(ref cancel) != 0 || !layout.Alive)
                                    { Application.ExitThread(); return; }
                                    if (NativeLayout.IsIconic(layout.Frame)) return;
                                    strip.Attach(layout.Reserve());
                                    strip.RefreshActive();
                                }
                                catch (Exception e) { stop(e); }
                            };
                            Console.WriteLine("Use + to keep the displayed page. Choose Close prototype to restore the layout.");
                            timer.Start();
                            Console.CancelKeyPress += cancelHandler;
                            try { Application.Run(); }
                            finally { Console.CancelKeyPress -= cancelHandler; }
                        }
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
        catch (Exception ex)
        {
            var native = ex as Win32Exception;
            if (native != null) Console.Error.WriteLine("Win32 error {0}: {1}", native.NativeErrorCode, native.Message);
            else Console.Error.WriteLine("{0}: {1} (0x{2:X8})", ex.GetType().Name, ex.Message, ex.HResult);
            return 1;
        }
    }
}
