using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

internal sealed class SavedPage
{
    internal readonly string Id, Caption, Link;
    private readonly Guid pageGuid;

    internal SavedPage(string id, string caption, string link)
    {
        Id = id;
        Caption = caption;
        Link = !String.IsNullOrEmpty(link) && link.StartsWith("onenote:", StringComparison.OrdinalIgnoreCase) ? link : "";
        // OneNote desktop hyperlinks carry a page-id in the fragment after the section path.
        foreach (var part in Link.Split('#', '&', '?'))
            if (part.StartsWith("page-id=", StringComparison.OrdinalIgnoreCase))
            {
                Guid.TryParse(Uri.UnescapeDataString(part.Substring(8)), out pageGuid);
                break;
            }
    }
    internal bool Matches(SavedPage other)
    {
        return other != null && (Id == other.Id || (pageGuid != Guid.Empty && pageGuid == other.pageGuid));
    }
    internal XElement ToXml()
    {
        return new XElement("page", new XAttribute("id", Id), new XAttribute("caption", Caption), new XAttribute("link", Link));
    }
}

internal sealed class WorkingSet
{
    internal readonly List<SavedPage> Tabs = new List<SavedPage>(), Bookmarks = new List<SavedPage>();
    internal string Revision;

    internal static bool Keep(List<SavedPage> list, SavedPage page)
    {
        int index = list.FindIndex(page.Matches);
        if (index < 0) { list.Add(page); return true; }
        var old = list[index];
        if (page.Link.Length == 0 && old.Link.Length != 0) page = new SavedPage(page.Id, page.Caption, old.Link);
        if (old.Id == page.Id && old.Caption == page.Caption && old.Link == page.Link) return false;
        list[index] = page;
        return true;
    }
}

// The in-memory form is used by fixture runs. Only normal Run opens the user's file.
internal sealed class SessionStore
{
    private readonly string path;
    private readonly WorkingSet memory = new WorkingSet();
    internal SessionStore(string path) { this.path = path; }

    internal static SessionStore ForUser()
    {
        return new SessionStore(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MoreNotes", "MoreTabs", "working-set.xml"));
    }

    internal WorkingSet Read()
    {
        if (path == null) return memory;
        XElement root;
        try
        {
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var reader = System.Xml.XmlReader.Create(file, new System.Xml.XmlReaderSettings {
                DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null }))
                root = XElement.Load(reader);
        }
        catch (FileNotFoundException) { return new WorkingSet(); }
        catch (DirectoryNotFoundException) { return new WorkingSet(); }
        catch (System.Xml.XmlException ex) { throw new InvalidDataException("Cannot read the saved MoreTabs file: " + path, ex); }
        if (root.Name != "workingSet" || (string)root.Attribute("version") != "1" ||
            String.IsNullOrEmpty((string)root.Attribute("revision")))
            throw new InvalidDataException("Unrecognized MoreTabs file: " + path);
        var state = new WorkingSet { Revision = (string)root.Attribute("revision") };
        ReadList(root.Element("tabs"), state.Tabs);
        ReadList(root.Element("bookmarks"), state.Bookmarks);
        return state;
    }
    private static void ReadList(XElement container, List<SavedPage> list)
    {
        if (container == null) throw new InvalidDataException("Missing MoreTabs page list.");
        foreach (var element in container.Elements("page"))
        {
            string id = (string)element.Attribute("id"), caption = (string)element.Attribute("caption");
            if (String.IsNullOrEmpty(id) || caption == null) throw new InvalidDataException("Invalid saved page reference.");
            WorkingSet.Keep(list, new SavedPage(id, caption, (string)element.Attribute("link")));
        }
    }

    internal WorkingSet Change(Func<WorkingSet, bool> edit)
    {
        if (path == null)
        {
            if (edit(memory)) memory.Revision = Guid.NewGuid().ToString("N");
            return memory;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        using (LockForWrite())
        {
            var state = Read();
            if (!edit(state)) return state;
            state.Revision = Guid.NewGuid().ToString("N");
            var root = new XElement("workingSet", new XAttribute("version", "1"), new XAttribute("revision", state.Revision),
                new XElement("tabs", state.Tabs.Select(page => page.ToXml())),
                new XElement("bookmarks", state.Bookmarks.Select(page => page.ToXml())));
            string temporary = path + "." + state.Revision + ".tmp";
            try
            {
                using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    root.Save(output);
                    output.Flush(true);
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { File.Delete(temporary); }
            return state;
        }
    }
    private FileStream LockForWrite()
    {
        for (int retry = 0; ; retry++)
        {
            try { return new FileStream(path + ".lock", FileMode.OpenOrCreate, FileAccess.Write, FileShare.None); }
            catch (IOException error)
            {
                int code = error.HResult & 0xffff;
                if ((code != 32 && code != 33) || retry >= 30) throw;
                Thread.Sleep(25);
            }
        }
    }
}
