using System;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Office.Core;
using One = Microsoft.Office.Interop.OneNote;

internal static class PageMenuConnection
{
    internal static string Connect(One.Application app, bool reconnect)
    {
        string step = "refresh the registered add-ins";
        try
        {
            var entries = app.COMAddIns as COMAddIns;
            if (entries == null) return "OneNote did not expose its COM add-ins collection.";
            entries.Update();
            step = "find the registered MoreTabs page menu";
            object id = PageMenuCommand.ProgId;
            var entry = entries.Item(ref id);
            if (entry == null) return "OneNote could not find MoreTabs page menu in its registered add-ins.";
            if (entry.Connect && !reconnect) return "OneNote reports the MoreTabs page menu is connected.";
            if (!reconnect)
            {
                object startup = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Office\OneNote\AddIns\" + id, "LoadBehavior", null);
                if (!Object.Equals(startup, 3)) return "MoreTabs page menu is disconnected. To enable it, choose the gear > Connect OneNote page menu.";
            }
            step = "connect the MoreTabs page menu";
            if (entry.Connect) entry.Connect = false;
            entry.Connect = true;
            return entry.Connect ? "OneNote connected the MoreTabs page menu. Right-click a page title in OneNote's page list." :
                "OneNote left the MoreTabs page menu disconnected. Check File > Options > Add-ins for disabled items.";
        }
        catch (Exception error)
        {
            if (!(error is COMException || error is InvalidCastException || error is ArgumentException || error is IndexOutOfRangeException ||
                error is UnauthorizedAccessException || error is System.Security.SecurityException)) throw;
            return String.Format("Could not {0} (0x{1:X8}). {2}", step, error.HResult, error.Message);
        }
    }
}
