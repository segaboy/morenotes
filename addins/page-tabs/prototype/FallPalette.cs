using System.Drawing;
using System.Drawing.Drawing2D;

internal static class FallPalette
{
    // Maple, amber, oak, cranberry, persimmon and sage. Dark surfaces get brighter foliage.
    private static readonly int[] light = { 0xA23C28, 0xB8760B, 0x626531, 0x8D3D55, 0xC45520, 0x657951 };
    private static readonly int[] dark = { 0xF49164, 0xF2C463, 0xBBBB73, 0xE897AA, 0xF5A54C, 0x9FC18B };

    internal static Color Foliage(string key, bool darkSurface)
    {
        return TabRenderer.PickSwatch(key, darkSurface ? dark : light);
    }

    private static readonly PointF[] maple = {
        new PointF(7,0), new PointF(9,3.5f), new PointF(11,2), new PointF(10,6),
        new PointF(14,4.5f), new PointF(13,7), new PointF(15,7.5f), new PointF(10,10.5f),
        new PointF(10,12), new PointF(7,11), new PointF(3.5f,12), new PointF(4,10),
        new PointF(0,7), new PointF(2,6.5f), new PointF(1,4), new PointF(5,6), new PointF(4,2), new PointF(6,3.5f)
    };
    internal static void DrawLeaf(Graphics g, float x, float y, float scale, float angle, Color color)
    {
        var saved = g.Save();
        try
        {
            g.TranslateTransform(x, y);
            g.RotateTransform(angle);
            g.ScaleTransform(scale, scale);
            g.TranslateTransform(-7, -7);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var fill = new SolidBrush(color))
            using (var stem = new Pen(color, 1.15f))
            {
                g.FillPolygon(fill, maple);
                g.DrawLine(stem, 7, 10, 6, 14);
            }
        }
        finally { g.Restore(saved); }
    }
}
