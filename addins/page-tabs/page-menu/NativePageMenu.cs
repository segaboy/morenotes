using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using One = Microsoft.Office.Interop.OneNote;

[ComVisible(true), Guid("D5F22F4C-A003-42DF-BCE8-071CD95EBE7D"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPageMenuActions
{
    [DispId(1)] void OpenInNewTab([MarshalAs(UnmanagedType.IDispatch)] object control);
    [DispId(2)] void RibbonLoaded(IRibbonUI ribbon);
    [DispId(3)] void OpenCurrentInNewTab([MarshalAs(UnmanagedType.IDispatch)] object control);
}

[ComVisible(true), Guid("32CA472B-C9FF-467B-8836-C98BC8DCB8A8"), ProgId(PageMenuCommand.ProgId),
 ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(IPageMenuActions))]
public sealed class NativePageMenu : IDTExtensibility2, IRibbonExtensibility, IPageMenuActions
{
    private EventWaitHandle ribbonLoaded;
    public static string RibbonSignal { get { return @"Local\MoreTabs.Ribbon." + typeof(NativePageMenu).Module.ModuleVersionId; } }
    public string GetCustomUI(string ribbonId)
    {
        return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='RibbonLoaded'><contextMenus>" +
            "<contextMenu idMso='ContextMenuPage'><button id='MoreTabsOpenPage' label='Open in new tab' " +
            "insertBeforeMso='RenamePageOneNote' onAction='OpenInNewTab'/></contextMenu>" +
            "<contextMenu idMso='ContextMenuPageArea'><button id='MoreTabsOpenCurrent' label='Open current page in new tab' " +
            "onAction='OpenCurrentInNewTab'/></contextMenu>" +
            "<contextMenu idMso='ContextMenuText'><button id='MoreTabsOpenPageText' label='Open in new tab' " +
            "insertBeforeMso='Cut' onAction='OpenCurrentInNewTab'/></contextMenu></contextMenus></customUI>";
    }
    public void RibbonLoaded(IRibbonUI ribbon)
    {
        if (ribbonLoaded == null)
        {
            try { ribbonLoaded = new EventWaitHandle(true, EventResetMode.ManualReset, RibbonSignal); }
            catch (UnauthorizedAccessException) { } // Status reporting is optional.
        }
    }
    public void OnConnection(object application, ext_ConnectMode mode, object addIn, ref Array custom)
    { }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom)
    {
        if (ribbonLoaded != null) { ribbonLoaded.Dispose(); ribbonLoaded = null; }
    }
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) { }

    public void OpenInNewTab(object control) { RequestTab(control, true); }
    public void OpenCurrentInNewTab(object control) { RequestTab(control, false); }

    private static void RequestTab(object control, bool pageList)
    {
        var editor = PageMenuCommand.ForegroundEditor();
        try
        {
            var ribbon = control as IRibbonControl;
            var window = ribbon == null ? null : ribbon.Context as One.Window;
            if (window != null)
            {
                var handle = new IntPtr(unchecked((long)window.WindowHandle));
                if (handle != IntPtr.Zero) editor = handle;
            }
        }
        catch (COMException) { } // The native popup's owning frame also identifies the editor.
        catch (NotImplementedException) { }
        catch (InvalidCastException) { }
        if (!PageMenuCommand.Request(editor, pageList))
            MessageBox.Show("The page menu could not reach MoreTabs in this window. Start the latest MoreTabs launcher and try again.", "MoreTabs");
    }
}
