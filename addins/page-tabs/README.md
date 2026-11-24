# MoreTabs

MoreTabs is the page-tabs add-in in the MoreNotes collection. An experimental Windows prototype tests whether the OneNote document area can give up a row for tabs. Its release identifier is `page-tabs`.

A normal workday can involve several pages in different notebooks. Give those pages a persistent row of browser-style tabs in the same OneNote window. OneNote remains the editor. The tab collection survives changes of active notebook. Include a new-tab `+` button and bookmarks so a user can keep a page after closing its tab.

## Run the prototype

From the repository root in Windows PowerShell, run:

```powershell
& .\addins\page-tabs\tools\Start-MoreTabsPrototype.ps1 -Mode Measure
```

The launcher uses the .NET Framework x64 compiler and the installed OneNote 15 interop assembly. It builds a local executable under `artifacts/`. Open OneNote in normal editing view with horizontal tabs and the Quick Access Toolbar below the ribbon.

| Mode | What it tests |
| --- | --- |
| `Build` | Retains a standalone x64 executable at `artifacts/prototype/MoreTabsPrototype.exe`. Pass any test mode directly to this executable from a desktop terminal. |
| `Measure` | Reports window handles, DPI, control bounds and native hit-test results for each measured window. Samples the left, center and right of the ribbon dock and section row without sending clicks or reading page content. |
| `Placement` | Temporarily reserves 32 logical pixels, attaches a disabled child strip, watches whether OneNote keeps the space for 1.2 seconds without intervention, then restores the layout. Supply `-WindowHandle <number>` from Measure to select among several windows. |
| `Com` | Opens the desktop COM interface and checks if the current window has a readable page ID. |
| `Bind` | Matches the measured native editor to a OneNote COM window. Prints handles and bounds without moving controls or navigating. |
| `Verify` | Runs the idle placement test in the matched editor. Then creates two local disposable notebooks under a unique temporary directory, drives the actual add/tab click handlers for three pages, and checks the window handle, selected page, and layout. Attempts to return to the original page and close the two test notebooks. The disposable notebook files stay in the temporary directory. |
| `HandsOn` | Sets up the same three disposable tabs and keeps the strip open for mouse and keyboard input. On **Close prototype**, checks for a typing marker in the disposable B page, restores the original page and closes the test notebooks. |
| `Run` | Opens a session in the matched editor. Use native navigation to visit a page, then press `+` to keep it in the row. Click a tab to return to its page. Selection follows the current page ID. Visiting an untracked page clears the highlight. Choose **Close prototype** to release the row. Tabs are held in memory during the experiment. |
| `SelfTest` | Runs a controlled WinForms window and in-memory page source to exercise child placement, resize, ribbon-height changes, restoration, and selection. This does not exercise OneNote. |

`Bind`, `Verify`, `HandsOn` and `Run` enumerate the visible layouts, then compare each COM window's top-level ancestor to the measured frames. They prefer the current COM window when it matches. If it cannot identify an editor, a single matching COM window is sufficient. With several candidates, supply `-WindowHandle` with the decimal frame number from the current measurement. A closed or ambiguous target stops the test before it changes OneNote.

### Hands-on test

Run the launcher with `-Mode HandsOn` in a desktop PowerShell window. Wait for the three tabs and the console's READY message.

1. From Dashboard in MoreTabs A, click **Schedule** in OneNote's page list, then **Back**. Pause on each page and check that its tab highlights.
2. Click **Dashboard / MoreTabs B** on the strip. Click the page body and type `MoreTabs focus OK`. Try reaching the strip from the page and returning using only the keyboard. Record any unreachable controls, stolen keystrokes or unexpected focus changes.
3. Resize the window, maximize and restore it, then collapse and expand the ribbon. Pause briefly after each change and check the strip and both adjacent rows. If available, move the window to a display with a different scale setting.
4. Choose **Close prototype**. The output reports displayed test pages, highlighting, sampled layout changes, DPI and whether the B page body contains the marker. The operator supplies the observations about input method and focus. A normal exit alone does not establish that each step was exercised.

### Live desktop verification

On November 24, 2026, prototype `bb88678` passed `Verify` in Windows desktop OneNote `16.0.20715.20004` (x64) at 96 DPI. The COM window handle belonged to a child of the measured editor frame. Matching its top-level ancestor selected the correct editor.

| Check | Observed result |
| --- | --- |
| Reserved row | The 1,345-pixel-wide strip occupied Y=216–247 below the ribbon and Quick Access Toolbar. The native section row moved from Y=216 to Y=248. The strip held its space during the 1.2-second idle test and through page switches. Win32 hit tests reached each row at left, center and right. |
| Page switching | The verifier exercised the add and tab click handlers on three disposable pages from two notebooks, including two pages titled Dashboard. All five tab selections reached the expected page ID in the same window. |
| Return and close | Returning to the original untracked page through COM cleared the active tab highlight. The verifier closed the test notebooks, released the row and exited with code 0. A separate measurement confirmed the original document bounds and section row at Y=216. |

The next live checks cover physical mouse and keyboard input, native page-list and Back navigation, resizing, ribbon collapse, and movement between displays with different scale settings. Before-and-after screenshots of test content are still pending.

### Earlier measurements and controlled tests

The probe initially measured three OneNote frames on the same Office build at 96 DPI. In each, the top ribbon dock ended at client Y=216 and a 48-pixel native navigation window began there. In one frame an open side pane left 985 pixels for the document in a 1,410-pixel client width. The strip therefore follows the document width.

The Work shell could measure the editor. Moving its document returned Win32 error 5, and COM activation returned `0x80070005`. The successful verification above ran under the signed-in Windows desktop account.

The desktop accessibility tree exposed the expanded ribbon, Quick Access Toolbar and Sections tab control. The original verifier required an exact match between the COM handle and the editor frame and could report a missing layout even when these controls were visible.

The controlled-window test passed child attachment, three ribbon/size configurations, preservation of the document's bottom edge, repeated layout passes without cumulative resizing, and restoration on close. Page tests passed with separate IDs for two pages titled Dashboard, repeated add, switching tabs, and navigating to an untracked page. Use `Run` for the remaining manual acceptance checks, including clicks on native controls, Back, keyboard focus, and a move between displays at different scale settings. Capture test content only.

On the 1,345-pixel-wide editor described above, the read-only Windows hit test reached the ribbon dock at Y=215 and the native navigation child at Y=240 at all three sampled positions. A controlled test placed another control over the left third of the navigation row. The original center-only check passed incorrectly. The revised check detected the obstruction and passed once the covering control was removed. The live `Placement` command still stopped at the document move with Win32 error 5, before it could attach a strip.

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

The prototype recognizes a native ribbon dock adjacent to a document window and identifies its horizontal navigation child by class and geometry. It requests space by moving and shortening the document window, then checks the new bounds before adding a child control. These window classes are observations from the tested build. Microsoft does not document them as a MoreTabs extension point. An unrecognized layout stops the experiment. The prototype registers no add-in and does not read or write notebook content in `Measure` or `Placement`.

## Follow-up requirements

- Decide how closing, reordering and pinning affect the working set and what the new-tab picker shows.
- Store bookmarks and session tabs in add-in-owned state. Test references after renaming and moving pages, notebook closure/reopening, deletion and application restart. Make missing pages identifiable and removable.
- Test page restoration separately from caret, selected text and scroll offset. Record what OneNote restores itself and what the add-in can reliably restore.
- Decide how multiple windows relate to saved sessions, how to organize bookmarks, and how to recover from a missing notebook.
