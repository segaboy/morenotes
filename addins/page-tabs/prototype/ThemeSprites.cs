using System.Drawing;
using System.Drawing.Drawing2D;

internal sealed class ThemeSprites
{
    private static readonly ThemeSprites roboCop = new ThemeSprites("robocop-visor.png", "robocop-skyline.png",
        new[] { 0x3B6284, 0x65778A, 0x337688, 0x686B91, 0x596680, 0x456E9A },
        new[] { 0x82C3E4, 0xA0B3C9, 0x75B6C7, 0xA3A9DC, 0x8A9AB6, 0x86A6D0 });
    private static readonly ThemeSprites alf = new ThemeSprites("alf-gordon.png", "alf-melmac.png",
        new[] { 0xA85F3F, 0x488080, 0x956B89, 0xB17D4D, 0x627F6C, 0x8A697D },
        new[] { 0xDE986D, 0x7BC2B4, 0xCEA3BC, 0xF0BF7A, 0xA3C9AB, 0xD18F91 });
    private static readonly ThemeSprites christmas = new ThemeSprites("christmas-tree.png", "christmas-village.png",
        new[] { 0x167650, 0xBD2443, 0xD69D27, 0x427F9D, 0x378D73, 0xA72C49 },
        new[] { 0x55C892, 0xF27786, 0xF6CA6C, 0x88D1EC, 0x77D9C0, 0xEF8799 });
    private readonly int[] light, dark;
    internal readonly Bitmap Icon, Scenery;

    private ThemeSprites(string icon, string scenery, int[] light, int[] dark)
    {
        Icon = Load(icon);
        Scenery = Load(scenery);
        this.light = light;
        this.dark = dark;
    }
    internal static ThemeSprites For(MoreTabsTheme theme)
    {
        switch (theme)
        {
            case MoreTabsTheme.RoboCop: return roboCop;
            case MoreTabsTheme.ALF: return alf;
            case MoreTabsTheme.Christmas: return christmas;
            default: return null;
        }
    }

    private static Bitmap Load(string name)
    {
        using (var stream = typeof(ThemeSprites).Assembly.GetManifestResourceStream(name))
        using (var source = new Bitmap(stream))
            return new Bitmap(source);
    }
    internal Color Swatch(string key, bool darkSurface)
    {
        return TabRenderer.PickSwatch(key, darkSurface ? dark : light);
    }
    internal static void Draw(Graphics g, Bitmap sprite, Rectangle bounds)
    {
        var saved = g.Save();
        try
        {
            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            g.PixelOffsetMode = PixelOffsetMode.Half;
            g.DrawImage(sprite, bounds, 0, 0, sprite.Width, sprite.Height, GraphicsUnit.Pixel);
        }
        finally { g.Restore(saved); }
    }
}
