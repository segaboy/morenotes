using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;

internal static class PageMenuSetup
{
    private static readonly Type Menu = typeof(NativePageMenu);
    private static readonly string Id = Menu.GUID.ToString("B"), ProgId = Marshal.GenerateProgIdForType(Menu);
    private static readonly string ClassKey = @"Software\Classes\CLSID\" + Id, AppKey = @"Software\Classes\AppID\" + Id;
    private static readonly string ProgKey = @"Software\Classes\" + ProgId, OfficeKey = @"Software\Microsoft\Office\OneNote\AddIns\" + ProgId;
    private static readonly string CodeBase = new Uri(Menu.Assembly.Location).AbsoluteUri;

    private static void SetValues(RegistryKey user, string path, params object[] pairs)
    {
        using (var key = user.CreateSubKey(path))
            for (int i = 0; i < pairs.Length; i += 2) key.SetValue((string)pairs[i], pairs[i + 1]);
    }
    private static void Register(RegistryKey user)
    {
        SetValues(user, ProgKey + @"\CLSID", "", Id);
        SetValues(user, AppKey, "DllSurrogate", "");
        SetValues(user, ClassKey, "", Menu.FullName, "AppID", Id);
        SetValues(user, ClassKey + @"\ProgID", "", ProgId);
        SetValues(user, ClassKey + @"\InprocServer32", "", "mscoree.dll", "ThreadingModel", "Both",
            "Class", Menu.FullName, "Assembly", Menu.Assembly.FullName, "RuntimeVersion", Menu.Assembly.ImageRuntimeVersion, "CodeBase", CodeBase);
        using (var office = user.CreateSubKey(OfficeKey))
        {
            office.SetValue("FriendlyName", "MoreTabs page menu");
            office.SetValue("Description", "Open a OneNote page in a new MoreTabs tab.");
            office.SetValue("CommandLineSafe", 0);
            using (var previous = user.OpenSubKey(ClassKey + @"\LocalServer32"))
            {
                object load = office.GetValue("LoadBehavior");
                if (load == null || (previous != null && Object.Equals(load, 2))) office.SetValue("LoadBehavior", 3);
            }
        }
        // Windows gives an EXE server precedence over the DLL surrogate for the same CLSID.
        user.DeleteSubKeyTree(ClassKey + @"\LocalServer32", false);
    }
    private static int Status(RegistryKey user)
    {
        using (var prog = user.OpenSubKey(ProgKey + @"\CLSID"))
        using (var cls = user.OpenSubKey(ClassKey))
        using (var legacy = user.OpenSubKey(ClassKey + @"\LocalServer32"))
        using (var server = user.OpenSubKey(ClassKey + @"\InprocServer32"))
        using (var office = user.OpenSubKey(OfficeKey))
        using (var app = user.OpenSubKey(AppKey))
        {
            if (prog == null || cls == null || server == null || office == null || app == null || legacy != null ||
                !Object.Equals(prog.GetValue(""), Id) || !Object.Equals(cls.GetValue("AppID"), Id) ||
                !Object.Equals(server.GetValue("CodeBase"), CodeBase) ||
                !Object.Equals(app.GetValue("DllSurrogate"), ""))
            {
                Console.WriteLine("The current MoreTabs page-menu DLL is not registered for this user."); return 1;
            }
            Console.WriteLine("OneNote page menu LoadBehavior = {0}.", office.GetValue("LoadBehavior"));
            if (!Object.Equals(office.GetValue("LoadBehavior"), 3))
            {
                Console.WriteLine("Check MoreTabs page menu in OneNote File > Options > Add-ins > Manage COM Add-ins > Go.");
                return 1;
            }
            try
            {
                using (EventWaitHandle.OpenExisting(NativePageMenu.RibbonSignal))
                    Console.WriteLine("OneNote requested the menu from this build.");
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                Console.WriteLine("Registered. Fully close and reopen OneNote to load this build's page menu.");
            }
            return 0;
        }
    }
    private static int Main(string[] args)
    {
        try
        {
            using (var user = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            {
                if (args.Length != 1) return 2;
                if (args[0] == "--unregister")
                {
                    foreach (string path in new[] { OfficeKey, ProgKey, ClassKey, AppKey }) user.DeleteSubKeyTree(path, false);
                    return 0;
                }
                if (args[0] == "--register") Register(user);
                else if (args[0] != "--status") return 2;
                return Status(user);
            }
        }
        catch (UnauthorizedAccessException e) { Console.Error.WriteLine(e.Message); return 1; }
        catch (System.Security.SecurityException e) { Console.Error.WriteLine(e.Message); return 1; }
        catch (System.IO.IOException e) { Console.Error.WriteLine(e.Message); return 1; }
    }
}
