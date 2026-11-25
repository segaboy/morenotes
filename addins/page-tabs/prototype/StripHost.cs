using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

// The arrow has its own column to the right of the ToolStrip's built-in overflow menu.
internal sealed class StripHost : Panel
{
    private readonly IntPtr frame;
    internal event Action<string> OpenPage;
    internal readonly Panel Captions = new Panel { TabStop = false };
    internal readonly DockArrow MoveButton = new DockArrow {
        Dock = DockStyle.Right, FlatStyle = FlatStyle.Flat, UseVisualStyleBackColor = false, UseMnemonic = false
    };
    internal StripHost(IntPtr frame, ToolStrip strip)
    {
        this.frame = frame;
        Text = PageMenuCommand.HostName;
        Visible = false;
        MoveButton.FlatAppearance.BorderSize = 0;
        strip.Dock = DockStyle.Fill;
        Controls.Add(strip);
        Controls.Add(Captions);
        Controls.Add(MoveButton);
    }
    protected override CreateParams CreateParams
    {
        get { var cp = base.CreateParams; cp.Parent = frame; cp.Style |= 0x40000000; return cp; }
    }
    protected override void WndProc(ref Message message)
    {
        if (message.Msg == PageMenuCommand.CopyData && Visible && Enabled && OpenPage != null)
        {
            string id = PageMenuCommand.Read(message);
            if (id != null)
            {
                // Return to OneNote before making any COM calls back into its editor.
                BeginInvoke(OpenPage, id);
                message.Result = new IntPtr(1);
                return;
            }
        }
        base.WndProc(ref message);
    }
    internal void Style(TabRenderer palette)
    {
        BackColor = MoveButton.BackColor = palette.Surface;
        Captions.BackColor = palette.Surface;
        // Center the names in all of the space between the tab edge and the edge of the band.
        Captions.Height = HostChrome.CaptionHeight(palette.Dpi) + HostChrome.Px(3, palette.Dpi);
        MoveButton.ForeColor = palette.Ink;
        MoveButton.FlatAppearance.MouseOverBackColor = palette.Hover;
        MoveButton.FlatAppearance.MouseDownBackColor = palette.Active;
        MoveButton.Width = HostChrome.Px(32, palette.Dpi);
    }
}

internal sealed class DockArrow : Button
{
    internal bool PointsUp;
    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        float scale = Width / 32f;
        float x = Width / 2f, y = Height / 2f, step = (PointsUp ? -1 : 1) * scale;
        bool highlighted = SystemInformation.HighContrast && ClientRectangle.Contains(PointToClient(Cursor.Position));
        using (var pen = new Pen(highlighted ? SystemColors.HighlightText : ForeColor, 1.6f * scale))
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            pen.StartCap = pen.EndCap = LineCap.Round;
            pen.LineJoin = LineJoin.Round;
            e.Graphics.DrawLine(pen, x, y - 5 * step, x, y + 5 * step);
            e.Graphics.DrawLines(pen, new[] { new PointF(x - 4 * scale, y + step), new PointF(x, y + 5 * step),
                new PointF(x + 4 * scale, y + step) });
        }
    }
}
