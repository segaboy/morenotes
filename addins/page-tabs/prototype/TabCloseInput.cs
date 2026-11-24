using System;
using System.Drawing;
using System.Windows.Forms;

// Handle the close target before ToolStrip turns a mouse release into a page-selection click.
// The overflow popup has its own window and uses the same tab items and hit geometry.
internal sealed class TabCloseInput : NativeWindow, IDisposable
{
    private readonly ToolStrip host;
    private readonly Action<SavedTab> close;
    private PageTabButton pressed;
    private bool tookCapture;

    internal TabCloseInput(ToolStrip host, Action<SavedTab> close)
    {
        this.host = host;
        this.close = close;
        host.HandleCreated += Attach;
        host.HandleDestroyed += Detach;
        if (host.IsHandleCreated) AssignHandle(host.Handle);
    }
    private void Attach(object sender, EventArgs e) { AssignHandle(host.Handle); }
    private void Detach(object sender, EventArgs e)
    {
        pressed = null;
        tookCapture = false;
        ReleaseHandle();
    }
    public void Dispose()
    {
        host.HandleCreated -= Attach;
        host.HandleDestroyed -= Detach;
        EndPress();
        if (Handle != IntPtr.Zero) ReleaseHandle();
    }
    private void EndPress()
    {
        pressed = null;
        if (tookCapture && host.Capture) host.Capture = false;
        tookCapture = false;
    }
    private PageTabButton Hit(IntPtr coordinates)
    {
        var p = new Point(unchecked((int)coordinates.ToInt64()));
        var tab = host.GetItemAt(p) as PageTabButton;
        if (tab == null || !tab.Enabled) return null;
        p.Offset(-tab.Bounds.Left, -tab.Bounds.Top);
        return tab.CloseBounds.Contains(p) ? tab : null;
    }
    protected override void WndProc(ref Message m)
    {
        const int leftDown = 0x0201, leftUp = 0x0202, leftDouble = 0x0203, cancel = 0x001F, captureChanged = 0x0215;
        if (m.Msg == leftDown || m.Msg == leftDouble)
        {
            pressed = Hit(m.LParam);
            if (pressed != null)
            {
                tookCapture = !host.Capture;
                if (tookCapture) host.Capture = true;
                return;
            }
        }
        else if (m.Msg == leftUp && pressed != null)
        {
            var target = Hit(m.LParam) == pressed ? (SavedTab)pressed.Tag : null;
            EndPress();
            if (target != null)
            {
                var menu = host as ToolStripDropDown;
                if (menu != null) menu.Close(ToolStripDropDownCloseReason.ItemClicked);
                close(target);
            }
            return;
        }
        else if (m.Msg == cancel || m.Msg == captureChanged) EndPress();
        base.WndProc(ref m);
    }
}
