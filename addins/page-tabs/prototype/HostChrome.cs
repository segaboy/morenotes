using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

// Read only the leading edge of OneNote's native navigation row. The bitmap is
// discarded after measuring the icon and surface. Never sample the page editor.
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
    internal static int StripHeight(uint dpi, bool captions = true) { return Px(27, dpi) + (captions ? CaptionHeight(dpi) : 0); }

    internal void Update(IntPtr navigation, uint dpi, bool force = false)
    {
        bool contrast = SystemInformation.HighContrast;
        if (!force && Dpi == dpi && HighContrast == contrast && DateTime.UtcNow < nextRead) return;
        nextRead = DateTime.UtcNow.AddSeconds(2);
        if (Dpi != dpi || HighContrast != contrast)
        {
            // Retain the measured logical inset when moving a covered window between monitors.
            Inset = Dpi == 0 ? Px(25, dpi) : (int)Math.Round(Inset * (double)dpi / Dpi);
            Measured = false;
        }
        Dpi = dpi;
        HighContrast = contrast;
        Color background, band;
        int left;
        SectionBand = Color.Empty;
        if (navigation != IntPtr.Zero && TryRead(navigation, dpi, out background, out left, out band))
        {
            if (Surface.ToArgb() != background.ToArgb()) nativeDark = null;
            Surface = background;
            Inset = left;
            Measured = true;
            SectionBand = band;
            nativeDark = SectionPalette.DarkPalette(band) ?? nativeDark;
        }
        else if (!Measured) Surface = FallbackSurface();
        if (contrast) Surface = SystemColors.Window;
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
}
