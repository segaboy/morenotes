using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

internal static class AppearanceTests
{
    private static void Check(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    internal static void Controlled(PageStrip strip, NativeLayout layout, Rectangle slot)
    {
        var labels = strip.Items.OfType<PageTabButton>().Select(tab => tab.Text).ToArray();
        Check(labels.SequenceEqual(new[] { "Dashboard / A", "Schedule", "Dashboard / B" }),
            "Tab labels must distinguish duplicate page titles.");
        var dark = Color.FromArgb(17, 12, 8);
        foreach (uint dpi in new uint[] { 96, 144 })
            foreach (var surface in new[] { dark, Color.FromArgb(247, 247, 247), SystemColors.Window })
            {
                bool contrast = surface.IsSystemColor;
                int inset = HostChrome.Px(22, dpi);
                var expected = new HostChrome { Dpi = dpi, Surface = surface, Inset = inset, HighContrast = contrast };
                using (var row = new Bitmap(HostChrome.Px(48, dpi), HostChrome.Px(43, dpi)))
                using (var g = Graphics.FromImage(row))
                using (var icon = new SolidBrush(contrast ? SystemColors.WindowText : Color.FromArgb(0, 155, 199)))
                {
                    g.Clear(surface);
                    g.FillRectangle(icon, inset, HostChrome.Px(9, dpi), HostChrome.Px(14, dpi), HostChrome.Px(20, dpi));
                    Color measured;
                    int left;
                    Check(HostChrome.FindIcon(row, dpi, out measured, out left) && left == inset && measured.ToArgb() == surface.ToArgb(),
                        "Did not measure the notebook glyph or the row color.");
                    g.Clear(surface);
                    Check(!HostChrome.FindIcon(row, dpi, out measured, out left), "An empty row produced a false icon anchor.");
                }
                strip.Size = new Size(HostChrome.Px(950, dpi), HostChrome.Px(32, dpi));
                strip.ApplyAppearance(expected);
                using (var bitmap = CheckStrip(strip, expected)) { }
            }
        strip.Attach(slot, layout);
        Console.WriteLine("PASS rendered dark, light and high-contrast tabs at 96 and 144 DPI: top seam, lower corners, active tab and icon alignment.");
    }

    private static Bitmap CheckStrip(PageStrip strip, HostChrome chrome)
    {
        strip.PerformLayout();
        var first = strip.Items.OfType<PageTabButton>().First(b => b.Placement == ToolStripItemPlacement.Main);
        var active = strip.Items.OfType<PageTabButton>().Single(b => b.Checked);
        var palette = (TabRenderer)strip.Renderer;
        Check(first.Bounds.Left == chrome.Inset && first.Bounds.Top == 0 && active.Bounds.Top == 0,
            "Tabs are not flush with the toolbar or notebook icon: tab " + first.Bounds + ", icon X=" + chrome.Inset);
        var bitmap = new Bitmap(strip.Width, strip.Height);
        try
        {
            strip.DrawToBitmap(bitmap, strip.ClientRectangle);
            Rectangle b = active.Bounds;
            int edge = HostChrome.Px(1, chrome.Dpi);
            Check(bitmap.GetPixel(b.Left + 2 * edge, b.Top + 2 * edge).ToArgb() == palette.Active.ToArgb(),
                "The active tab has a gap or rounding at its upper-left corner.");
            Check(bitmap.GetPixel(b.Left + edge - 1, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb() &&
                bitmap.GetPixel(b.Right - edge, b.Bottom - 1).ToArgb() == palette.Surface.ToArgb(),
                "The active tab's lower corners are not rounded.");
            Check(bitmap.GetPixel(b.Left + b.Width / 2, b.Bottom - 3 * edge).ToArgb() == palette.Active.ToArgb(),
                "The active tab did not use its theme color.");
            return bitmap;
        }
        catch { bitmap.Dispose(); throw; }
    }

    internal static void Live(PageStrip strip, NativeLayout layout, string fixtureDirectory)
    {
        var chrome = layout.Appearance(true);
        strip.ApplyAppearance(chrome);
        using (var bitmap = CheckStrip(strip, chrome))
            bitmap.Save(Path.Combine(fixtureDirectory, "MoreTabs-strip.png"), ImageFormat.Png);
        Console.WriteLine("PASS tab rendering in OneNote: surface #{0:X6}, first tab X={1}px ({2}), DPI {3}; square top and rounded lower corners.",
            chrome.Surface.ToArgb() & 0xffffff, chrome.Inset, chrome.Measured ? "measured notebook icon" : "fallback inset", chrome.Dpi);
    }
}
