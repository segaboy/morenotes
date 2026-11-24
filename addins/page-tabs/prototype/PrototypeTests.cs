using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using One = Microsoft.Office.Interop.OneNote;

internal static class PrototypeTests
{
    private static void Check(bool condition, string failure)
    {
        if (!condition) throw new InvalidOperationException(failure);
    }
    private sealed class TestPages : IPages
    {
        public string Current { get; private set; }
        public string Caption(string id) { return id == "b/1" ? "Dashboard / B" : id == "a/1" ? "Dashboard / A" : "Schedule / A"; }
        public void Navigate(string id) { Current = id; }
    }
    internal static void ControlledWindow()
    {
        using (var frame = new Form { Text = "MoreTabs layout test", ShowInTaskbar = false, StartPosition = FormStartPosition.Manual,
            Location = new Point(-16000, -16000), ClientSize = new Size(730, 420) })
        using (var ribbon = new Panel())
        using (var document = new Panel())
        using (var navigation = new Panel { Dock = DockStyle.Top, Height = 36 })
        {
            frame.Controls.Add(ribbon);
            frame.Controls.Add(document);
            document.Controls.Add(navigation);
            frame.Show();
            try
            {
                var source = new TestPages();
                Exception failure = null;
                using (var layout = new NativeLayout(frame.Handle, document.Handle, ribbon.Handle, navigation.Handle))
                {
                    Rectangle expected = Rectangle.Empty;
                    using (var strip = new PageStrip(frame.Handle, source, delegate {}, delegate(Exception e) { failure = e; }))
                    {
                        foreach (int height in new int[] { 120, 60, 120 })
                        {
                            frame.ClientSize = new Size(frame.ClientSize.Width + 40, frame.ClientSize.Height + 10);
                            ribbon.Bounds = new Rectangle(0, 0, frame.ClientSize.Width, height);
                            expected = new Rectangle(0, height, frame.ClientSize.Width - 110, frame.ClientSize.Height - height);
                            document.Bounds = expected;
                            var slot = layout.Reserve();
                            strip.Attach(slot);
                            layout.CheckHitTargets(strip.Handle, slot);
                            Check(slot == layout.Reserve(), "A repeated tick consumed another row.");
                            Check(slot.Top == height && slot.Bottom == layout.NavigationBounds.Top && slot.Width == expected.Width &&
                                layout.DocumentBounds.Bottom == expected.Bottom, "The test strip overlaps native controls.");
                        }
                        string[] ids = { "a/1", "a/2", "b/1" };
                        foreach (var id in ids) { source.Navigate(id); strip.ClickAdd(); }
                        strip.ClickAdd();
                        Check(strip.TabCount == 3, "Duplicate add created another tab.");
                        foreach (var id in ids) { strip.ClickTab(id); Check(source.Current == id && strip.IsSelected(id), "Tab click did not navigate."); }
                        source.Navigate("a/1"); strip.RefreshActive();
                        Check(strip.IsSelected("a/1") && !strip.IsSelected("b/1"), "External navigation left a stale highlight.");
                        source.Navigate("other"); strip.RefreshActive();
                        foreach (var id in ids) Check(!strip.IsSelected(id), "Untracked page left a highlight.");
                        if (failure != null) throw failure;
                    }
                    layout.Dispose();
                    Check(layout.DocumentBounds == expected, "Normal close did not restore the document.");
                }
                Console.WriteLine("PASS controlled window: child attachment, 3 layout changes, no cumulative movement, restore, three tabs, duplicate titles, duplicate add and external page changes.");
            }
            finally { frame.Close(); }
        }
    }

    private static void WaitForPage(NotebookPages source, string id)
    {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        while (source.Current != id)
        {
            if (timer.ElapsedMilliseconds > 5000) throw new TimeoutException("OneNote did not display the requested test page.");
            Application.DoEvents(); Thread.Sleep(40);
        }
    }

    private static string CreatePage(One.Application app, string section, string name)
    {
        string id, xml;
        app.CreateNewPage(section, out id, One.NewPageStyle.npsBlankPageWithTitle);
        app.GetPageContent(id, out xml, One.PageInfo.piBasic, One.XMLSchema.xs2013);
        var page = XElement.Parse(xml);
        XNamespace ns = page.Name.Namespace;
        var title = page.Element(ns + "Title");
        var value = title == null ? null : title.Descendants(ns + "T").FirstOrDefault();
        if (value == null) throw new InvalidOperationException("Missing title in the disposable page.");
        value.ReplaceNodes(new XCData(name));
        app.UpdatePageContent(page.ToString(), DateTime.MinValue, One.XMLSchema.xs2013, false);
        return id;
    }

    internal static void DisposableNotebooks(NotebookPages source, PageStrip strip, NativeLayout layout, Rectangle slot)
    {
        string previous = source.Current;
        if (String.IsNullOrEmpty(previous)) throw new InvalidOperationException("Select a page that can be restored after the test.");
        string directory = Path.Combine(Path.GetTempPath(), "MoreTabs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var books = new List<string>();
        var pages = new List<string>();
        try
        {
            foreach (string bookName in new[] { "MoreTabs A", "MoreTabs B" })
            {
                string book, section;
                source.App.OpenHierarchy(Path.Combine(directory, bookName), "", out book, One.CreateFileType.cftNotebook);
                books.Add(book);
                source.App.OpenHierarchy("Test.one", book, out section, One.CreateFileType.cftSection);
                pages.Add(CreatePage(source.App, section, "Dashboard"));
                if (books.Count == 1) pages.Add(CreatePage(source.App, section, "Schedule"));
            }
            for (int i = 0; i < pages.Count; i++)
            {
                source.Navigate(pages[i]); WaitForPage(source, pages[i]);
                Check(source.Window.CurrentNotebookId == books[i == 2 ? 1 : 0], "Notebook mismatch.");
                Check(layout.StillReserved(slot), "OneNote reclaimed the strip when navigating to a test page.");
                strip.ClickAdd();
            }
            Check(strip.TabCount == 3, "Expected three tracked pages.");
            foreach (int i in new[] { 0, 2, 1, 2, 0 })
            {
                strip.ClickTab(pages[i]); WaitForPage(source, pages[i]); strip.RefreshActive();
                Check(strip.IsSelected(pages[i]) && source.Window.WindowHandle == (ulong)layout.Frame.ToInt64(), "Wrong page or OneNote window.");
                Check(layout.StillReserved(slot), "OneNote reclaimed the strip after a tab click.");
            }
            source.Navigate(previous); WaitForPage(source, previous); strip.RefreshActive();
            foreach (var id in pages) Check(!strip.IsSelected(id), "A stale test tab remained active.");
            Console.WriteLine("PASS OneNote COM and tab clicks: three pages, two notebooks, same editor window, external return to an untracked page.");
        }
        finally
        {
            try { source.Navigate(previous); WaitForPage(source, previous); }
            finally { foreach (var book in books) source.App.CloseNotebook(book, false); }
        }
    }
}
