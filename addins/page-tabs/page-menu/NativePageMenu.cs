using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
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
    private EventWaitHandle ribbonLoaded;
    public static string RibbonSignal { get { return @"Local\MoreTabs.Ribbon." + typeof(NativePageMenu).Module.ModuleVersionId; } }
    public string GetCustomUI(string ribbonId)
    {
        if (ribbonLoaded == null)
        {
            try { ribbonLoaded = new EventWaitHandle(true, EventResetMode.ManualReset, RibbonSignal); }
            catch (UnauthorizedAccessException) { } // Status reporting is optional.
        }
        return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><contextMenus>" +
            "<contextMenu idMso='ContextMenuPage'><button id='MoreTabsOpenPage' label='Open in new tab' " +
            "onAction='OpenInNewTab'/></contextMenu>" +
            "<contextMenu idMso='ContextMenuText'><button id='MoreTabsOpenPageText' label='Open in new tab' " +
            "onAction='OpenInNewTab'/></contextMenu></contextMenus></customUI>";
    }
    public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
    {
        app = (One.Application)application;
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom)
    {
        app = null;
        if (ribbonLoaded != null) { ribbonLoaded.Dispose(); ribbonLoaded = null; }
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
            string id;
            if (control.Id == "MoreTabsOpenPageText") id = window.CurrentPageId;
            else
            {
                string xml, section = window.CurrentSectionId;
                if (String.IsNullOrEmpty(section)) return;
                window.Application.GetHierarchy(section, One.HierarchyScope.hsPages, out xml, One.XMLSchema.xs2013);
                id = window.CurrentSectionId == section ? PageMenuCommand.SelectedPage(XElement.Parse(xml)) : null;
            }
            if (String.IsNullOrEmpty(id))
                MessageBox.Show("Select one page in the OneNote page list and try again.", "MoreTabs");
            else if (!PageMenuCommand.Send(new IntPtr(unchecked((long)window.WindowHandle)), id))
                MessageBox.Show("Start MoreTabs in this OneNote window, then try again.", "MoreTabs");
        }
        catch (COMException) { MessageBox.Show("OneNote could not read the selected page. Try again after the editor is ready.", "MoreTabs"); }
    }
}
