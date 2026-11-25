using System.Drawing;
using System.Drawing.Drawing2D;

internal static class ThemeSprites
{
    private static readonly int[] light = { 0x3B6284, 0x65778A, 0x337688, 0x686B91, 0x596680, 0x456E9A };
    private static readonly int[] dark = { 0x82C3E4, 0xA0B3C9, 0x75B6C7, 0xA3A9DC, 0x8A9AB6, 0x86A6D0 };
    private static readonly int[] melmacLight = { 0xA85F3F, 0x488080, 0x956B89, 0xB17D4D, 0x627F6C, 0x8A697D };
    private static readonly int[] melmacDark = { 0xDE986D, 0x7BC2B4, 0xCEA3BC, 0xF0BF7A, 0xA3C9AB, 0xD18F91 };
    private static readonly Bitmap visor = Load("robocop-visor.png"), skyline = Load("robocop-skyline.png");
    private static readonly Bitmap gordon = Load("alf-gordon.png"), melmac = Load("alf-melmac.png");
    internal static Bitmap TabIcon(MoreTabsTheme theme) { return theme == MoreTabsTheme.ALF ? gordon : visor; }
    internal static Bitmap Scenery(MoreTabsTheme theme) { return theme == MoreTabsTheme.ALF ? melmac : skyline; }

    private static Bitmap Load(string name)
    {
        using (var stream = typeof(ThemeSprites).Assembly.GetManifestResourceStream(name))
        using (var source = new Bitmap(stream))
            return new Bitmap(source);
    }
    internal static Color Swatch(MoreTabsTheme theme, string key, bool darkSurface)
    {
        var colors = theme == MoreTabsTheme.ALF ? (darkSurface ? melmacDark : melmacLight) : (darkSurface ? dark : light);
        return TabRenderer.PickSwatch(key, colors);
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
