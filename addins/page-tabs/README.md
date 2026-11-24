# Page tabs for OneNote

**Status:** design and host inspection. The product name for this add-in is open. `page-tabs` is its repository identifier.

A normal workday can involve several pages in different notebooks. Give those pages a persistent row of browser-style tabs in the same OneNote window. OneNote remains the editor. The tab collection survives changes of active notebook. Include a new-tab `+` button and bookmarks so a user can keep a page after closing its tab.

## First milestone

Prove same-window navigation and the requested placement on Windows desktop. Use three disposable pages across two notebooks. Give two pages the same title to verify that selection uses the page reference. Record the Office build, architecture, display scaling and results for each test.

| Test | Acceptance criteria |
| --- | --- |
| Placement | With the Quick Access Toolbar below the ribbon and horizontal section tabs active, the new strip occupies its own row directly between the toolbar and native notebook/section row. Both native rows and the editor remain fully visible and clickable. Capture before and after views of test content. |
| New tab and navigation | Use `+` to add or choose existing pages. Alternate among all three in the original OneNote window. The active page ID matches the selected tab each time, the window handle stays the same, and switching notebooks preserves the set. |
| Native navigation | Use OneNote's native page list, section navigation and Back to visit tracked pages. The active tab follows. Visiting an untracked page clears the previous highlight or indicates the untracked state without navigating the user back. |
| Layout | Resize and maximize/restore the window. Collapse/expand the ribbon. Check at 100% and 150% scaling and across monitors with different scaling when available. Verify alignment, click targets, native controls and page content each time. Record unavailable configurations as untested. |
| Keyboard focus | Navigate and activate tabs with the keyboard, return focus to the page, and type into a test page. Repeat after using native navigation. No stolen keystrokes, unexpected window activation or focus trap. |

Start by identifying whether OneNote can reserve vertical space at the target location. Observe the native layout and accessibility tree while resizing or collapsing the ribbon. An element's visible rectangle identifies a measurement target only. Record how, or whether, the host participates in allocating space for the strip. If a test fails, document the method, reproducible failure and affected configuration. Present alternative layouts and their tradeoffs for a decision before adopting one.

## Technical findings

| Evidence | Implication |
| --- | --- |
| The desktop COM window interface exposes `WindowHandle`, `CurrentPageId`, `CurrentNotebookId` and `NavigateTo`. | Bind a tab collection to the intended window. Read the page ID from that window to verify a navigation. [Microsoft window interfaces](https://learn.microsoft.com/en-us/office/client-developer/onenote/window-interfaces-onenote) |
| The application interface provides hierarchy, linking and navigation methods and describes navigation events. | Evaluate native navigation events against Back and section changes. Check the window identity again when handling an event. [Microsoft application interface](https://learn.microsoft.com/en-us/office/client-developer/onenote/application-interface-onenote) |
| Microsoft's archived desktop sample demonstrates a OneNote COM add-in and setup. The Office.js OneNote overview targets the web. The VSTO application matrix lists the supported hosts. | The COM sample is a starting point. An extension point for this exact row is still unverified. Neither Office.js nor a standard VSTO custom task pane establishes that placement. [COM sample](https://github.com/OneNoteDev/VanillaAddIn) · [Office.js overview](https://learn.microsoft.com/en-us/office/dev/add-ins/onenote/onenote-add-ins-programming-overview) · [VSTO host matrix](https://learn.microsoft.com/en-us/visualstudio/vsto/features-available-by-office-application-and-project-type?view=visualstudio) |

The bootstrap host has 64-bit Microsoft 365 Click-to-Run, OneNote executable version `16.0.20715.20004`, and a registered `OneNote.Application` COM class. Run `tools/Get-OneNoteHost.ps1` from this directory in 64-bit Windows PowerShell to inspect the local installation without opening a notebook. Registration alone does not exercise the API. Choose the UI and packaging approach after the live layout experiment.

## Follow-up requirements

- Decide how closing, reordering and pinning affect the working set and what the new-tab picker shows.
- Store bookmarks and session tabs in add-in-owned state. Test references after renaming and moving pages, notebook closure/reopening, deletion and application restart. Make missing pages identifiable and removable.
- Test page restoration separately from caret, selected text and scroll offset. Record what OneNote restores itself and what the add-in can reliably restore.
- Decide how multiple windows relate to saved sessions, how to organize bookmarks, and how to recover from a missing notebook.
