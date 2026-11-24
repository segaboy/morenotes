using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using Microsoft.Win32;
using One = Microsoft.Office.Interop.OneNote;

[ComVisible(true), Guid("D5F22F4C-A003-42DF-BCE8-071CD95EBE7D"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPageMenuActions
{
    [DispId(1)] void OpenInNewTab(IRibbonControl control);
}

[ComVisible(true), Guid("32CA472B-C9FF-467B-8836-C98BC8DCB8A8"), ProgId("MoreNotes.MoreTabs.PageMenu"),
 ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IPageMenuActions))]
public sealed class NativePageMenu : IDTExtensibility2, IRibbonExtensibility, IPageMenuActions
{
    private One.Application app;
    internal static int Connections;
    public string GetCustomUI(string ribbonId)
    {
        return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><contextMenus>" +
            "<contextMenu idMso='ContextMenuPage'><button id='MoreTabsOpenPage' label='Open in new tab' " +
            "insertBeforeMso='RenamePageOneNote' onAction='OpenInNewTab'/></contextMenu></contextMenus></customUI>";
    }
    public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
    {
        if (app == null) { app = (One.Application)application; Interlocked.Increment(ref Connections); }
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom)
    {
        if (app != null) { app = null; Interlocked.Decrement(ref Connections); }
    }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public void OpenInNewTab(IRibbonControl control)
    {
        try
        {
            One.Window window = null;
            try { window = control.Context as One.Window; } catch (COMException) { }
            if (window == null && app != null) window = app.Windows.CurrentWindow;
            if (window == null) return;
            string xml, section = window.CurrentSectionId;
            if (String.IsNullOrEmpty(section)) return;
            app.GetHierarchy(section, One.HierarchyScope.hsPages, out xml, One.XMLSchema.xs2013);
            string id = PageMenuCommand.SelectedPage(XElement.Parse(xml));
            if (id == null || window.CurrentSectionId != section)
                MessageBox.Show("Select one page in the OneNote page list and try again.", "MoreTabs");
            else if (!PageMenuCommand.Send(new IntPtr(unchecked((long)window.WindowHandle)), id))
                MessageBox.Show("Start MoreTabs in this OneNote window, then try again.", "MoreTabs");
        }
        catch (COMException) { MessageBox.Show("OneNote could not read the selected page. Try again after the editor is ready.", "MoreTabs"); }
    }
}

internal static class PageMenuServer
{
    private static void SetValues(RegistryKey user, string path, params object[] values)
    {
        using (var key = user.CreateSubKey(path))
            for (int i = 0; i < values.Length; i += 2) key.SetValue((string)values[i], values[i + 1]);
    }
    private static void RegisterMenu(bool remove)
    {
        var type = typeof(NativePageMenu);
        string id = type.GUID.ToString("B"), progId = Marshal.GenerateProgIdForType(type);
        string cls = @"Software\Classes\CLSID\" + id, prog = @"Software\Classes\" + progId;
        string addIn = @"Software\Microsoft\Office\OneNote\AddIns\" + progId;
        using (var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
        {
            if (remove)
            {
                foreach (var path in new[] { addIn, prog, cls }) user.DeleteSubKeyTree(path, false);
                return;
            }
            string exe = type.Assembly.Location;
            SetValues(user, prog + @"\CLSID", "", id);
            SetValues(user, cls + @"\ProgID", "", progId);
            SetValues(user, cls + @"\LocalServer32", "", "\"" + exe + "\"", "ServerExecutable", exe);
            SetValues(user, addIn, "FriendlyName", "MoreTabs page menu", "Description", "Open a OneNote page in a new MoreTabs tab.", "CommandLineSafe", 0);
            using (var key = user.OpenSubKey(addIn, true))
                if (key.GetValue("LoadBehavior") == null) key.SetValue("LoadBehavior", 3);
        }
    }
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length == 1 && (args[0] == "--register" || args[0] == "--unregister"))
        {
            try { RegisterMenu(args[0] == "--unregister"); return 0; }
            catch (UnauthorizedAccessException e) { Console.Error.WriteLine(e.Message); return 1; }
            catch (System.Security.SecurityException e) { Console.Error.WriteLine(e.Message); return 1; }
            catch (System.IO.IOException e) { Console.Error.WriteLine(e.Message); return 1; }
        }
        if (!Array.Exists(args, a => String.Equals(a.TrimStart('-', '/'), "Embedding", StringComparison.OrdinalIgnoreCase))) return 0;
        var registration = new RegistrationServices();
        int cookie = registration.RegisterTypeForComClients(typeof(NativePageMenu), RegistrationClassContext.LocalServer,
            RegistrationConnectionType.MultipleUse);
        try
        {
            DateTime idleDeadline = DateTime.UtcNow.AddMinutes(1);
            using (var timer = new System.Windows.Forms.Timer { Interval = 1000 })
            {
                timer.Tick += delegate {
                    if (Thread.VolatileRead(ref NativePageMenu.Connections) > 0) idleDeadline = DateTime.UtcNow.AddSeconds(5);
                    else if (DateTime.UtcNow >= idleDeadline) Application.ExitThread();
                };
                timer.Start();
                Application.Run();
            }
        }
        finally { registration.UnregisterTypeForComClients(cookie); }
        return 0;
    }
}
