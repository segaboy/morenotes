using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

// Sample OneNote chrome only. Never sample the page editor.
internal sealed class HostChrome
{
    internal uint Dpi;
    internal int Inset;
    internal Color Surface;
    internal Color SectionBand;
    private bool? nativeDark;
    internal bool DarkSections { get { return nativeDark ?? (Surface.GetBrightness() < 0.5f); } }
    internal bool Measured, HighContrast;
    private DateTime nextRead;
    private bool verticalTabs;
    private Size ribbonSize;
    private uint ribbonDpi;
    private Size ribbonInsets;
    private DateTime nextRibbonRead;

    internal static int Px(int logical, uint dpi) { return (int)Math.Round(logical * dpi / 96.0); }

    internal static Font CaptionFont(uint dpi)
    {
        using (var menu = SystemFonts.MenuFont)
            return new Font(menu.FontFamily, menu.SizeInPoints * dpi / 72f * 0.84f, FontStyle.Regular, GraphicsUnit.Pixel);
    }
    private static uint captionDpi;
    private static int captionHeight;
    internal static int CaptionHeight(uint dpi)
    {
        if (captionDpi != dpi)
        {
            using (var font = CaptionFont(dpi)) captionHeight = font.Height;
            captionDpi = dpi;
        }
        return captionHeight;
    }
    internal static int ContentGap(uint dpi, bool atBottom) { return atBottom ? Px(4, dpi) : 0; }
    internal static int StripHeight(uint dpi, bool captions = true, bool atBottom = false)
    {
        return Px(27, dpi) + (captions ? CaptionHeight(dpi) : 0) + ContentGap(dpi, atBottom);
    }

    // Office paints the ribbon card inside its HWND. Its drop shadow is outside the card.
    internal Rectangle RibbonFace(IntPtr ribbon, Rectangle dock, uint dpi, bool measure = true)
    {
        if (dpi != ribbonDpi)
        {
            if (ribbonDpi != 0)
                ribbonInsets = new Size((int)Math.Round(ribbonInsets.Width * (double)dpi / ribbonDpi),
                    (int)Math.Round(ribbonInsets.Height * (double)dpi / ribbonDpi));
            ribbonDpi = dpi;
            nextRibbonRead = DateTime.MinValue;
        }
        if (measure && (dock.Size != ribbonSize || DateTime.UtcNow >= nextRibbonRead))
        {
            ribbonSize = dock.Size;
            nextRibbonRead = DateTime.UtcNow.AddSeconds(2);
            Size measured;
            if (TryReadRibbonInsets(ribbon, dock.Size, dpi, out measured)) ribbonInsets = measured;
        }
        return Rectangle.FromLTRB(dock.Left + ribbonInsets.Width, dock.Top, dock.Right - ribbonInsets.Height, dock.Bottom);
    }

    private static bool TryReadRibbonInsets(IntPtr ribbon, Size size, uint dpi, out Size insets)
    {
        insets = Size.Empty;
        int width = Px(24, dpi), step = Px(6, dpi), startY = size.Height - Px(24, dpi);
        var origin = Point.Empty;
        if (width < 4 || startY < 0 || size.Width < 2 * width || !ClientToScreen(ribbon, ref origin)) return false;
        using (var left = new Bitmap(width, 1))
        using (var right = new Bitmap(width, 1))
        using (var lg = Graphics.FromImage(left))
        using (var rg = Graphics.FromImage(right))
        {
            var votes = new Dictionary<Size, int>();
            for (int row = 0; row < 3; row++)
            {
                int y = origin.Y + startY + row * step;
                int rightX = origin.X + size.Width - width;
                bool visible = true;
                foreach (int x in new[] { origin.X, origin.X + width - 1, rightX, rightX + width - 1 })
                {
                    IntPtr hit = WindowFromPoint(new Point(x, y));
                    if (hit != ribbon && !IsChild(ribbon, hit)) { visible = false; break; }
                }
                if (!visible) continue;
                try
                {
                    lg.CopyFromScreen(origin.X, y, 0, 0, left.Size);
                    rg.CopyFromScreen(rightX, y, 0, 0, right.Size);
                }
                catch (Win32Exception) { return false; }
                int l = RibbonInset(left, false, dpi), r = RibbonInset(right, true, dpi);
                if (left.GetPixel(l, 0) != right.GetPixel(width - 1 - r, 0)) continue;
                var pair = new Size(l, r);
                int count;
                votes.TryGetValue(pair, out count);
                if (count == 1) { insets = pair; return true; }
                votes[pair] = count + 1;
            }
        }
        return false;
    }

    internal static int RibbonInset(Bitmap edge, bool fromRight, uint dpi)
    {
        int direction = fromRight ? -1 : 1, start = fromRight ? edge.Width - 1 : 0;
        int run = Math.Max(3, Px(3, dpi));
        Color previous = edge.GetPixel(start, 0);
        for (int inset = 1; inset <= edge.Width - run; inset++)
        {
            int x = start + inset * direction;
            Color c = edge.GetPixel(x, 0);
            int delta = Math.Max(Math.Abs(c.R - previous.R), Math.Max(Math.Abs(c.G - previous.G), Math.Abs(c.B - previous.B)));
            previous = c;
            if (delta < 6) continue;
            int flat = 1;
            while (flat < run && edge.GetPixel(x + flat * direction, 0) == c) flat++;
            if (flat == run) return inset;
        }
        return 0;
    }

    internal void Update(IntPtr navigation, uint dpi, bool force = false, bool vertical = false)
    {
        bool contrast = SystemInformation.HighContrast;
        if (!force && Dpi == dpi && HighContrast == contrast && verticalTabs == vertical && DateTime.UtcNow < nextRead) return;
        nextRead = DateTime.UtcNow.AddSeconds(2);
        if (Dpi != dpi || HighContrast != contrast)
        {
            // Retain the measured logical inset when moving a covered window between monitors.
            Inset = Dpi == 0 ? Px(25, dpi) : (int)Math.Round(Inset * (double)dpi / Dpi);
            if (HighContrast != contrast) Measured = false;
        }
        if (verticalTabs != vertical) Inset = Px(25, dpi);
        verticalTabs = vertical;
        Dpi = dpi;
        HighContrast = contrast;
        Color background, band;
        int left;
        SectionBand = Color.Empty;
        if (!vertical && navigation != IntPtr.Zero && TryRead(navigation, dpi, out background, out left, out band))
        {
            if (Surface.ToArgb() != background.ToArgb()) nativeDark = null;
            Surface = background;
            Inset = left;
            Measured = true;
            SectionBand = band;
            nativeDark = SectionPalette.DarkPalette(band) ?? nativeDark;
        }
        else if (vertical && navigation != IntPtr.Zero && TryReadNavigationSurface(navigation, dpi, out background))
        {
            Surface = background;
            nativeDark = null;
            Measured = true;
        }
        else if (!Measured) Surface = FallbackSurface();
        if (contrast) Surface = SystemColors.Window;
    }

    private static bool TryReadNavigationSurface(IntPtr header, uint dpi, out Color surface)
    {
        surface = Color.Empty;
        Point origin = Point.Empty;
        RECT client;
        if (!ClientToScreen(header, ref origin) || !GetClientRect(header, out client)) return false;
        var dc = GetDC(IntPtr.Zero);
        if (dc == IntPtr.Zero) return false;
        var colors = new List<int>();
        try
        {
            // Stay in the notebook end of the header. Page-list and editor windows can overlap its bounds.
            for (int y = 4; y < 32 && Px(y, dpi) < client.Bottom; y += 8)
                for (int x = 6; x < 64 && Px(x, dpi) < client.Right; x += 12)
                {
                    var p = new Point(origin.X + Px(x, dpi), origin.Y + Px(y, dpi));
                    var hit = WindowFromPoint(p);
                    if (hit != header && !IsChild(header, hit)) continue;
                    int rgb = GetPixel(dc, p.X, p.Y);
                    if (rgb != -1) colors.Add(rgb);
                }
        }
        finally { ReleaseDC(IntPtr.Zero, dc); }
        return NavigationSurface(colors, out surface);
    }
    internal static bool NavigationSurface(List<int> colors, out Color surface)
    {
        var counts = new Dictionary<int, int>();
        int best = 0, peak = 0;
        foreach (int rgb in colors)
        {
            int count;
            counts.TryGetValue(rgb, out count);
            counts[rgb] = ++count;
            if (count > peak) { best = rgb; peak = count; }
        }
        surface = ColorTranslator.FromWin32(best);
        return peak >= 3 && peak * 2 > colors.Count;
    }

    private static Color FallbackSurface()
    {
        const string office = @"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common";
        const string windows = @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
        try
        {
            object theme = Registry.GetValue(office, "UI Theme", null);
            if (Object.Equals(theme, 4)) return Color.FromArgb(26, 25, 25);
            if (Object.Equals(theme, 3)) return Color.FromArgb(62, 62, 62);
            if (Object.Equals(theme, 0) || Object.Equals(theme, 5)) return Color.FromArgb(244, 244, 244);
            if (Object.Equals(Registry.GetValue(windows, "AppsUseLightTheme", null), 0)) return Color.FromArgb(26, 25, 25);
        }
        catch (System.Security.SecurityException) { }
        catch (UnauthorizedAccessException) { }
        return SystemColors.Control;
    }

    private static bool TryRead(IntPtr navigation, uint dpi, out Color surface, out int left, out Color band)
    {
        surface = Color.Empty;
        band = Color.Empty;
        left = 0;
        Point origin = Point.Empty;
        RECT client;
        if (!ClientToScreen(navigation, ref origin) || !GetClientRect(navigation, out client)) return false;
        var size = new Size(Math.Min(client.Right, Px(48, dpi)), Math.Min(client.Bottom, Px(60, dpi)));
        if (size.Width < Px(30, dpi) || size.Height < Px(18, dpi)) return false;
        // A console, tooltip or another window covering this patch must not set the theme.
        int margin = Px(4, dpi);
        foreach (int x in new[] { margin, size.Width / 2, size.Width - margin - 1 })
            foreach (int y in new[] { margin, size.Height / 2, size.Height - margin - 1 })
            {
                IntPtr hit = WindowFromPoint(new Point(origin.X + x, origin.Y + y));
                if (hit != navigation && !IsChild(navigation, hit)) return false;
            }
        using (var image = new Bitmap(size.Width, size.Height))
        using (var graphics = Graphics.FromImage(image))
        {
            try { graphics.CopyFromScreen(origin, Point.Empty, size); }
            catch (Win32Exception) { return false; }
            if (!FindIcon(image, dpi, out surface, out left)) return false;
            if (size.Height == client.Bottom) band = ReadSectionBand(image, dpi, surface);
            return true;
        }
    }

    internal static Color ReadSectionBand(Bitmap image, uint dpi, Color surface)
    {
        // The native selected section continues as a solid rule below the notebook icon.
        int left = Px(13, dpi), right = image.Width - Px(5, dpi);
        if (left >= right || image.Height < 2) return Color.Empty;
        Color band = image.GetPixel(left, image.Height - 1);
        if (band.ToArgb() == surface.ToArgb()) return Color.Empty;
        for (int x = left; x < right; x++)
            for (int y = image.Height - 2; y < image.Height; y++)
                if (image.GetPixel(x, y).ToArgb() != band.ToArgb()) return Color.Empty;
        return band;
    }

    internal static bool FindIcon(Bitmap image, uint dpi, out Color surface, out int left)
    {
        int margin = Px(4, dpi), bottom = image.Height - Px(6, dpi);
        var histogram = new Dictionary<int, int>();
        int peak = 0, background = 0;
        for (int x = margin; x < image.Width - margin; x++)
            for (int y = margin; y < bottom; y++)
            {
                int pixel = image.GetPixel(x, y).ToArgb(), count;
                histogram.TryGetValue(pixel, out count);
                histogram[pixel] = ++count;
                if (count > peak) { peak = count; background = pixel; }
            }
        surface = Color.FromArgb(background);
        left = 0;
        if (peak < (image.Width - 2 * margin) * (bottom - margin) / 4) return false;
        int start = -1;
        for (int x = margin; x < image.Width - margin; x++)
        {
            int ink = 0;
            for (int y = margin; y < bottom; y++)
            {
                Color c = image.GetPixel(x, y);
                if (Math.Max(Math.Abs(c.R - surface.R), Math.Max(Math.Abs(c.G - surface.G), Math.Abs(c.B - surface.B))) >= 38) ink++;
            }
            if (ink >= Px(7, dpi) && ink <= Px(27, dpi))
            {
                if (start < 0) start = x;
                if (x - start + 1 >= Px(6, dpi)) { left = start; return true; }
            }
            else start = -1;
        }
        return false;
    }
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr h, out RECT rect);
    [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr h, ref Point p);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(Point p);
    [DllImport("user32.dll")] private static extern bool IsChild(IntPtr parent, IntPtr child);
    [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr window);
    [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr window, IntPtr dc);
    [DllImport("gdi32.dll")] private static extern int GetPixel(IntPtr dc, int x, int y);
}
