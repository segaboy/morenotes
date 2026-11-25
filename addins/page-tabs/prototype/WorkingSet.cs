using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

internal sealed class SavedPage
{
    internal readonly string Id, Caption, Link;
    internal readonly BookmarkLocation Location;
    internal BookmarkNotebook Notebook { get { return Location == null ? null : Location.Notebook; } }
    private readonly Guid pageGuid;

    internal SavedPage(string id, string caption, string link, BookmarkLocation location = null)
    {
        Id = id;
        Caption = caption;
        Location = location;
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
    internal string BookmarkTitle
    {
        get
        {
            string suffix = Notebook == null ? null : " / " + Notebook.Name;
            if (suffix != null && Caption.EndsWith(suffix, StringComparison.Ordinal) && Caption.Length > suffix.Length)
                return Caption.Substring(0, Caption.Length - suffix.Length);
            int split = Caption.LastIndexOf(" / ", StringComparison.Ordinal);
            return split > 0 ? Caption.Substring(0, split) : Caption;
        }
    }
    internal BookmarkNotebook NotebookForMenu()
    {
        if (Notebook != null) return Notebook;
        int split = Caption.LastIndexOf(" / ", StringComparison.Ordinal);
        return new BookmarkNotebook("", split > 0 ? Caption.Substring(split + 3) : "Unknown notebook", Color.Empty);
    }
    internal XElement ToXml()
    {
        return new XElement("page", new XAttribute("id", Id), new XAttribute("caption", Caption), new XAttribute("link", Link),
            Notebook != null ? Notebook.ToXml() : null, Location == null ? null : Location.SectionXml());
    }
}

internal sealed class SavedTab
{
    internal readonly string Key;
    internal readonly SavedPage Page;
    internal readonly bool Pinned;
    internal SavedTab(string key, SavedPage page, bool pinned = false) { Key = key; Page = page; Pinned = pinned; }
}

internal enum MoreTabsTheme { NotebookColors, Fall, RoboCop }

internal sealed class WorkingSet
{
    internal readonly List<SavedTab> Tabs = new List<SavedTab>();
    internal readonly List<SavedPage> Bookmarks = new List<SavedPage>();
    internal readonly List<string> PinnedNotebookKeys = new List<string>();
    internal string Revision;
    internal readonly HashSet<string> HiddenTools = new HashSet<string>(StringComparer.Ordinal);
    internal SavedPage HomePage;
    internal string NewNoteSectionId, NewNoteSectionName;
    internal bool ThemesEnabled = true, ShowNotebookLabels = true, GroupByNotebook = true;
    internal MoreTabsTheme Theme = MoreTabsTheme.NotebookColors;
    internal MoreTabsTheme EffectiveTheme { get { return ThemesEnabled ? Theme : MoreTabsTheme.NotebookColors; } }
    internal ContentAlignment LabelAlignment = ContentAlignment.MiddleLeft;

    internal bool CycleLabelAlignment()
    {
        switch (LabelAlignment)
        {
            case ContentAlignment.MiddleLeft: LabelAlignment = ContentAlignment.MiddleCenter; break;
            case ContentAlignment.MiddleCenter: LabelAlignment = ContentAlignment.MiddleRight; break;
            default: LabelAlignment = ContentAlignment.MiddleLeft; break;
        }
        return true;
    }

    internal bool KeepTab(SavedPage page, string preferred, out string key)
    {
        var tab = Tabs.Find(t => t.Key == preferred && page.Matches(t.Page)) ?? Tabs.Find(t => page.Matches(t.Page));
        key = tab == null ? Guid.NewGuid().ToString("N") : tab.Key;
        if (tab != null) return UpdateTab(tab, page);
        Tabs.Add(new SavedTab(key, page));
        return true;
    }

    internal bool UpdateTab(SavedTab expected, SavedPage page)
    {
        int index = Tabs.FindIndex(t => t.Key == expected.Key);
        if (index < 0 || !expected.Page.Matches(Tabs[index].Page)) return false;
        var old = Tabs[index].Page;
        if (page.Matches(old))
            page = new SavedPage(page.Id, page.Caption, page.Link.Length == 0 ? old.Link : page.Link, page.Location ?? old.Location);
        if (page.Id == old.Id && page.Caption == old.Caption && page.Link == old.Link &&
            (page.Location == old.Location || (page.Location != null && page.Location.SameAs(old.Location)))) return false;
        Tabs[index] = new SavedTab(expected.Key, page, Tabs[index].Pinned);
        return true;
    }

    internal bool SetPinned(string key, bool pinned)
    {
        int index = Tabs.FindIndex(t => t.Key == key);
        if (index < 0 || Tabs[index].Pinned == pinned) return false;
        Tabs[index] = new SavedTab(key, Tabs[index].Page, pinned);
        return true;
    }

    internal int NotebookPinOrder(string key)
    {
        int position = PinnedNotebookKeys.IndexOf(key);
        return position < 0 ? Int32.MaxValue : position;
    }

    internal bool SetNotebookPinned(string key, bool pinned)
    {
        if (!pinned)
        {
            if (PinnedNotebookKeys.Remove(key)) return true;
            // Another window may have learned the ID since this pin menu opened.
            return key.StartsWith("name:", StringComparison.Ordinal) &&
                PinnedNotebookKeys.Remove(ResolveNamedPin(key, KnownNotebooks()));
        }
        if (PinnedNotebookKeys.Contains(key)) return false;
        PinnedNotebookKeys.Add(key);
        return true;
    }

    internal ILookup<string, BookmarkNotebook> KnownNotebooks()
    {
        return Tabs.Select(t => t.Page.Notebook).Concat(Bookmarks.Select(p => p.Notebook))
            .Where(book => book != null && book.Id.Length != 0).GroupBy(book => book.Id).Select(g => g.First())
            .ToLookup(book => book.Name, StringComparer.OrdinalIgnoreCase);
    }

    internal void ResolveNotebookPins()
    {
        // Older tabs can supply just a notebook name until OneNote identifies the page.
        if (!PinnedNotebookKeys.Any(k => k.StartsWith("name:", StringComparison.Ordinal))) return;
        var known = KnownNotebooks();
        for (int i = 0; i < PinnedNotebookKeys.Count; i++)
        {
            string key = PinnedNotebookKeys[i];
            if (!key.StartsWith("name:", StringComparison.Ordinal)) continue;
            PinnedNotebookKeys[i] = ResolveNamedPin(key, known);
        }
        var seen = new HashSet<string>(StringComparer.Ordinal);
        PinnedNotebookKeys.RemoveAll(key => !seen.Add(key));
    }
    private static string ResolveNamedPin(string key, ILookup<string, BookmarkNotebook> known)
    {
        var matches = known[key.Substring(5)];
        return matches.Count() == 1 ? matches.First().GroupKey : key;
    }

    internal bool UpdateBookmarkOwners(Dictionary<string, BookmarkLocation> owners)
    {
        bool changed = false;
        for (int i = 0; i < Bookmarks.Count; i++)
        {
            var page = Bookmarks[i];
            BookmarkLocation owner;
            if (!owners.TryGetValue(page.Id, out owner) || owner.SameAs(page.Location)) continue;
            Bookmarks[i] = new SavedPage(page.Id, page.Caption, page.Link, owner);
            changed = true;
        }
        return changed;
    }

    internal bool UpdateTabLocations(Dictionary<string, BookmarkLocation> owners)
    {
        bool changed = false;
        for (int i = 0; i < Tabs.Count; i++)
        {
            var tab = Tabs[i];
            var page = tab.Page;
            BookmarkLocation owner;
            if (!owners.TryGetValue(page.Id, out owner) || owner.SameAs(page.Location)) continue;
            Tabs[i] = new SavedTab(tab.Key, new SavedPage(page.Id, page.Caption, page.Link, owner), tab.Pinned);
            changed = true;
        }
        return changed;
    }

    internal static bool Keep(List<SavedPage> list, SavedPage page)
    {
        int index = list.FindIndex(page.Matches);
        if (index < 0) { list.Add(page); return true; }
        var old = list[index];
        page = new SavedPage(page.Id, page.Caption, page.Link.Length == 0 ? old.Link : page.Link, page.Location ?? old.Location);
        if (old.Id == page.Id && old.Caption == page.Caption && old.Link == page.Link &&
            (old.Location == page.Location || (old.Location != null && old.Location.SameAs(page.Location)))) return false;
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
        string version = (string)root.Attribute("version");
        if (root.Name != "workingSet" || (version != "1" && version != "2" && version != "3") ||
            String.IsNullOrEmpty((string)root.Attribute("revision")))
            throw new InvalidDataException("Unrecognized MoreTabs file: " + path);
        var state = new WorkingSet { Revision = (string)root.Attribute("revision") };
        var options = root.Element("options");
        if (options != null)
        {
            foreach (var pin in options.Elements("pinnedNotebook"))
            {
                string key = (string)pin.Attribute("key");
                if (!String.IsNullOrEmpty(key)) state.SetNotebookPinned(key, true);
            }
            foreach (var tool in options.Elements("hiddenTool"))
            {
                string id = (string)tool.Attribute("id");
                if (!String.IsNullOrEmpty(id)) state.HiddenTools.Add(id);
            }
            state.ThemesEnabled = (bool?)options.Attribute("coloredLabels") ?? true;
            switch ((string)options.Attribute("theme"))
            {
                case "Fall": state.Theme = MoreTabsTheme.Fall; break;
                case "RoboCop": state.Theme = MoreTabsTheme.RoboCop; break;
            }
            state.ShowNotebookLabels = (bool?)options.Attribute("showLabels") ?? true;
            state.GroupByNotebook = (bool?)options.Attribute("groupByNotebook") ?? true;
            switch ((string)options.Attribute("labelAlignment"))
            {
                case "MiddleCenter": state.LabelAlignment = ContentAlignment.MiddleCenter; break;
                case "MiddleRight": state.LabelAlignment = ContentAlignment.MiddleRight; break;
            }
            var home = options.Element("page");
            if (home != null) state.HomePage = ReadPage(home);
            var section = options.Element("newNoteSection");
            if (section != null)
            {
                state.NewNoteSectionId = (string)section.Attribute("id");
                state.NewNoteSectionName = (string)section.Attribute("name");
                if (String.IsNullOrEmpty(state.NewNoteSectionId) || String.IsNullOrEmpty(state.NewNoteSectionName))
                    throw new InvalidDataException("Invalid saved new note section.");
            }
        }
        foreach (var element in Pages(root.Element("tabs")))
        {
            var page = ReadPage(element);
            // Two processes reading the same unmigrated file must agree on each slot's identity.
            string key = version == "1" ? "legacy-" + page.Id : (string)element.Attribute("tabId");
            if (String.IsNullOrEmpty(key) || state.Tabs.Exists(t => t.Key == key)) throw new InvalidDataException("Invalid saved tab identity.");
            string pin = version == "3" ? (string)element.Attribute("pinned") : null;
            if (pin != null && pin != "true" && pin != "false") throw new InvalidDataException("Invalid saved tab pin.");
            state.Tabs.Add(new SavedTab(key, page, pin == "true"));
        }
        foreach (var element in Pages(root.Element("bookmarks"))) WorkingSet.Keep(state.Bookmarks, ReadPage(element));
        return state;
    }
    private static IEnumerable<XElement> Pages(XElement container)
    {
        if (container == null) throw new InvalidDataException("Missing MoreTabs page list.");
        return container.Elements("page");
    }
    private static SavedPage ReadPage(XElement element)
    {
        string id = (string)element.Attribute("id"), caption = (string)element.Attribute("caption");
        if (String.IsNullOrEmpty(id) || caption == null) throw new InvalidDataException("Invalid saved page reference.");
        return new SavedPage(id, caption, (string)element.Attribute("link"), BookmarkLocation.FromXml(element));
    }

    internal WorkingSet Change(Func<WorkingSet, bool> edit)
    {
        if (path == null)
        {
            if (edit(memory))
            {
                memory.ResolveNotebookPins();
                memory.Revision = Guid.NewGuid().ToString("N");
            }
            return memory;
        }
        Directory.CreateDirectory(Path.GetDirectoryName(path));
        using (LockForWrite())
        {
            var state = Read();
            if (!edit(state)) return state;
            state.ResolveNotebookPins();
            state.Revision = Guid.NewGuid().ToString("N");
            var root = new XElement("workingSet", new XAttribute("version", "3"), new XAttribute("revision", state.Revision),
                new XElement("options", new XAttribute("coloredLabels", state.ThemesEnabled), new XAttribute("theme", state.Theme),
                    new XAttribute("showLabels", state.ShowNotebookLabels), new XAttribute("groupByNotebook", state.GroupByNotebook),
                    new XAttribute("labelAlignment", state.LabelAlignment),
                    state.PinnedNotebookKeys.Select(key => new XElement("pinnedNotebook", new XAttribute("key", key))),
                    state.HiddenTools.OrderBy(id => id, StringComparer.Ordinal).Select(id => new XElement("hiddenTool", new XAttribute("id", id))),
                    state.HomePage == null ? null : state.HomePage.ToXml(),
                    state.NewNoteSectionId == null ? null : new XElement("newNoteSection",
                        new XAttribute("id", state.NewNoteSectionId), new XAttribute("name", state.NewNoteSectionName))),
                new XElement("tabs", state.Tabs.Select(tab => {
                    var xml = tab.Page.ToXml();
                    xml.SetAttributeValue("tabId", tab.Key);
                    if (tab.Pinned) xml.SetAttributeValue("pinned", true);
                    return xml;
                })),
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
