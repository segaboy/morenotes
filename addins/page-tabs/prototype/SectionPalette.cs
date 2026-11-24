using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;

// The desktop hierarchy stores older palette values. Horizontal tabs in the Fluent UI
// translate those values again when painting. RGB swatches from Office 16.0.20715.20004.
internal static class SectionPalette
{
    private struct Swatch
    {
        internal readonly Color Light, Dark;
        internal Swatch(int light, int dark) { Light = Rgb(light); Dark = Rgb(dark); }
    }
    private static readonly Dictionary<int, Swatch> byStoredColor = new Dictionary<int, Swatch>();
    private static readonly Dictionary<int, bool> isDarkSwatch = new Dictionary<int, bool>();
    static SectionPalette()
    {
        // Light and dark tab fills, followed by section colors found in desktop hierarchy palettes.
        Add(0xAED2F5, 0x055098, 0x8AA8E4, 0x004F8B, 0x0088F0, 0x8AB6E2); // Blue
        Add(0xFCD13C, 0x937200, 0xFFD869, 0xFFC114, 0xF1CA5D);           // Yellow
        Add(0x8FE497, 0x2A7931, 0xB7C997, 0x008C3A, 0x009E42, 0x7ACC93); // Green
        Add(0xFFA2A9, 0xA71D28, 0xEE9597, 0xE71224, 0xEC5158, 0xF39B9B); // Red
        Add(0xE7A2FF, 0x810FA9, 0xB49EDE, 0xAB008B, 0xFA00CC, 0xD399CF); // Purple
        Add(0x8EE4EA, 0x006168, 0x91BAAE, 0x00A0D7, 0x87C9D9);           // Cyan
        Add(0xFFC39F, 0xC44A01, 0xF6B078, 0xF6630D, 0xF1B86F);           // Orange
        Add(0xF6A9DC, 0xB82987, 0xD5A4BB, 0xD20078, 0xFF1F9E, 0xEB9FC9); // Magenta
        Add(0xB3C5F0, 0x364465, 0x9BBBD2, 0x0069AF, 0x32A6FF, 0x96ABDE); // Blue Mist
        Add(0xD6BEF2, 0x653A98, 0xAB9CB7, 0x5B2D90, 0xB973D3, 0xAAA5DB); // Purple Mist
        Add(0xE4C7B1, 0x70513A, 0xE8D199, 0xC6A377, 0xCDA485);           // Tan
        Add(0xB9E375, 0x5E7C2C, 0xF5F96F, 0x7EC400, 0xB4D367);           // Lemon Lime
        Add(0x99E0CB, 0x44655B, 0xADE792, 0x00B44B, 0x00C753, 0x9DCD78); // Apple
        Add(0x67E0BC, 0x045840, 0x4DBCCA, 0x1EC3B4, 0x75CDB8);           // Teal
        Add(0xF1A9A9, 0x982D2D, 0xBA7575, 0xC10052, 0xE45383, 0xF0AAAA); // Red Chalk
        Add(0xDADCF1, 0x68686C, 0x9595AA, 0x849398, 0xB9C0C7, 0xFFFFFF); // Silver / None
    }
    private static Color Rgb(int rgb) { return Color.FromArgb(255, Color.FromArgb(rgb)); }
    private static void Add(int light, int dark, params int[] stored)
    {
        var swatch = new Swatch(light, dark);
        foreach (int rgb in stored) byStoredColor.Add(rgb, swatch);
        isDarkSwatch.Add(light, false);
        isDarkSwatch.Add(dark, true);
    }
    internal static bool? DarkPalette(Color band)
    {
        bool dark;
        return !band.IsEmpty && isDarkSwatch.TryGetValue(band.ToArgb() & 0xffffff, out dark) ? dark : (bool?)null;
    }
    internal static Color Native(Color stored, bool dark)
    {
        Swatch swatch;
        if (byStoredColor.TryGetValue(stored.ToArgb() & 0xffffff, out swatch)) return dark ? swatch.Dark : swatch.Light;
        if (!dark || stored.IsEmpty) return stored;
        // For a custom RGB, the dark horizontal tab reduces HLS lightness by 75%.
        uint colorRef = (uint)ColorTranslator.ToWin32(stored);
        ushort h, l, s;
        ColorRGBToHLS(colorRef, out h, out l, out s);
        uint result = ColorHLSToRGB(h, (ushort)(l / 4), s);
        if (result == colorRef) result = ColorHLSToRGB(h, (ushort)(l + (240 - l) / 10), s);
        return ColorTranslator.FromWin32((int)result);
    }
    [DllImport("shlwapi.dll")] private static extern void ColorRGBToHLS(uint rgb, out ushort h, out ushort l, out ushort s);
    [DllImport("shlwapi.dll")] private static extern uint ColorHLSToRGB(ushort h, ushort l, ushort s);
}
