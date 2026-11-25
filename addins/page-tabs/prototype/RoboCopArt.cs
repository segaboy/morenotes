using System.Drawing;
using System.Drawing.Drawing2D;

internal static class RoboCopArt
{
    private static readonly int[] light = { 0x3B6284, 0x65778A, 0x337688, 0x686B91, 0x596680, 0x456E9A };
    private static readonly int[] dark = { 0x82C3E4, 0xA0B3C9, 0x75B6C7, 0xA3A9DC, 0x8A9AB6, 0x86A6D0 };
    internal static readonly Bitmap Visor = Load("robocop-visor.png");
    internal static readonly Bitmap Skyline = Load("robocop-skyline.png");

    private static Bitmap Load(string name)
    {
        using (var stream = typeof(RoboCopArt).Assembly.GetManifestResourceStream(name))
        using (var source = new Bitmap(stream))
            return new Bitmap(source);
    }
    internal static Color Steel(string key, bool darkSurface)
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
