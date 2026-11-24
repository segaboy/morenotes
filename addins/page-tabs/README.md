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
| `Build` | Retains a standalone x64 executable and a double-click launcher, `artifacts/prototype/Run-MoreTabs.cmd`. |
| `Measure` | Reports window handles, DPI, control bounds and native hit-test results for each measured window. Samples the left, center and right of the ribbon dock and section row without sending clicks or reading page content. |
| `Placement` | Temporarily reserves 32 logical pixels, attaches a disabled child strip, watches whether OneNote keeps the space for 1.2 seconds without intervention, then restores the layout. Supply `-WindowHandle <number>` from Measure to select among several windows. |
| `Com` | Opens the desktop COM interface and checks if the current window has a readable page ID. |
| `Bind` | Matches the measured native editor to a OneNote COM window. Prints handles and bounds without moving controls or navigating. |
| `Verify` | Creates two disposable notebooks, keeps three pages, and clicks **+** twice with the Quick Notes destination set to a section in the second test notebook. Checks both new page IDs, their parent section, the editor window and keyboard focus on each new tab. Attempts to restore the original page and close the test notebooks. The disposable notebook files stay in the temporary directory. |
| `HandsOn` | Sets up the same three disposable tabs and keeps the strip open for mouse and keyboard input. On **Close prototype**, checks for a typing marker in the disposable B page, restores the original page and closes the test notebooks. |
| `Run` | Restores the saved tabs in the matched editor and follows its current page. When that editor closes, the running prototype waits for another compatible OneNote editor. **Close prototype** exits the program. |
| `SelfTest` | Exercises a controlled window and starts fresh prototype processes to verify disk persistence, bookmark reopening, and an intentionally empty session. Uses isolated temporary data. |

`Bind`, `Verify`, `HandsOn` and `Run` enumerate the visible layouts, then compare each COM window's top-level ancestor to the measured frames. They prefer the current COM window when it matches. If it cannot identify an editor, a single matching COM window is sufficient. With several candidates, supply `-WindowHandle` with the decimal frame number from the current measurement. The handle selects the initial attachment. A restarted editor receives a new handle.

### Saved tabs and bookmarks

Page tabs hang from the Quick Access Toolbar. Their tops meet the toolbar and their lower corners use an eight-logical-pixel radius. The strip samples the background and notebook icon from the visible leading edge of the native navigation row. It uses that color and aligns the first tab to the icon, refreshing the measurement when OneNote changes appearance or moves between displays. While the icon is covered, the strip keeps the last measurement. At startup it can fall back to the Office or Windows theme and a DPI-scaled inset. Windows high-contrast colors take precedence. The active page has a distinct fill, and keyboard focus adds an outline.

![MoreTabs rendered with sample pages and the measured dark OneNote theme](appearance.png)

Tabs show the page title. If two open pages share a title, their labels also show the notebook. Hovering over a tab shows its full saved caption.

Start `Run-MoreTabs.cmd` from the build output or use `-Mode Run` from desktop PowerShell. On first use, the strip starts with the displayed page. Subsequent runs restore the saved tab order, including an intentionally empty row. OneNote chooses its opening page. A matching tab highlights automatically.

- Click **+** to create a blank page in your configured OneNote Quick Notes section. Each click opens its own tab in the same editor and gives that tab keyboard focus. This follows your personal Quick Notes destination whatever notebook you have open. If you have redirected Quick Notes, MoreTabs uses that setting.
- To keep an existing page, use the arrow beside **+** and choose **Keep current page in a tab**. Right-click a tab and choose **Close tab**, or middle-click it. Closing a tab leaves the underlying OneNote page intact.
- Choose **Bookmarks → Bookmark current page**, or right-click a tab and choose **Bookmark this tab**. Select a saved page in **Bookmarks** to open it in the same editor and return it to the tab row. **Remove bookmark** removes only the bookmark.
- Leave the prototype running when closing OneNote to let it attach when OneNote reopens. After exiting the prototype or restarting Windows, launch it again. The prototype has no installer or sign-in startup entry.

Tabs and bookmarks share one user-local file at `%LOCALAPPDATA%\MoreNotes\MoreTabs\working-set.xml`. Each edit saves the page ID, caption and OneNote desktop hyperlink when available. File updates merge under a short write lock and replace the prior version atomically. Separate attached windows use the same saved working set. The saved file contains notebook metadata and belongs outside public diagnostics.

MoreTabs first navigates using the saved page ID. If OneNote rejects it, MoreTabs tries the saved OneNote link in the same window. An unavailable page stays in the list for retry or removal. Restoring caret position, selection and scroll offset remains unverified.

### Hands-on test

Run the launcher with `-Mode HandsOn` in a desktop PowerShell window. Wait for the three tabs and the console's READY message.

1. From Dashboard in MoreTabs A, click **Schedule** in OneNote's page list, then **Back**. Pause on each page and check that its tab highlights.
2. From **Dashboard / MoreTabs A**, put the caret in the page body and press **Ctrl+Alt+F6** to focus the active page tab. Press **Right** twice and **Enter** to open Dashboard in MoreTabs B. Without clicking, type `MoreTabs focus OK`. Try **Ctrl+Alt+F6** followed by **Escape** to return without navigating. The interactive prototype requests the shortcut while its editor is in the foreground and reports any registration conflict. Left and Right select a strip item, Enter activates it, and Escape exits the strip.
3. Resize the window, maximize and restore it, then collapse and expand the ribbon. Pause briefly after each change and check the strip and both adjacent rows. If available, move the window to a display with a different scale setting.
4. Choose **Close prototype**. The output reports displayed test pages, highlighting, sampled layout changes, DPI and whether the B page body contains the marker. The operator supplies the observations about input method and focus. A normal exit alone does not establish that each step was exercised.

### Live desktop verification

On November 24, 2026, a read-only measurement of OneNote `16.0.20715.20004` (x64) at 96 DPI found the left edge of the notebook icon 25 pixels into the row and its surrounding color was `#110C08`. The preview above shows the prototype rendered with those measurements and two sample pages. The controlled renderer tests checked dark, light and high-contrast palettes at 96 and 144 DPI, including both bottom corners and the position of the tab against the toolbar.

On November 24, 2026, prototype `bb88678` passed `Verify` in Windows desktop OneNote `16.0.20715.20004` (x64) at 96 DPI. The COM window handle belonged to a child of the measured editor frame. Matching its top-level ancestor selected the correct editor.

| Check | Observed result |
| --- | --- |
| Reserved row | The 1,345-pixel-wide strip occupied Y=216–247 below the ribbon and Quick Access Toolbar. The native section row moved from Y=216 to Y=248. The strip held its space during the 1.2-second idle test and through page switches. Win32 hit tests reached each row at left, center and right. |
| Page switching | The verifier exercised the add and tab click handlers on three disposable pages from two notebooks, including two pages titled Dashboard. All five tab selections reached the expected page ID in the same window. |
| Return and close | Returning to the original untracked page through COM cleared the active tab highlight. The verifier closed the test notebooks, released the row and exited with code 0. A separate measurement confirmed the original document bounds and section row at Y=216. |

The saved-session verification also passed in the live editor on the same x64 build at 96 DPI. With the ribbon dock ending at Y=78, the strip moved the navigation row to Y=110. The verifier saved three disposable tabs across two notebooks, bookmarked one and closed its tab. A fresh strip read the file, restored the two remaining tabs, and reopened the bookmarked page in the other notebook in the same editor. The verifier restored the preceding page, released the row and exited with code 0.

The controlled suite separately verified restoration across fresh processes, an empty saved session, bookmark removal, fallback from an obsolete page ID, and concurrent writes from independent stores. It also closed a controlled editor while the interactive loop was running and restored the saved row in a second editor.

On November 24, 2026, prototype `b7c1fd7` passed the Quick Notes test in Windows desktop OneNote `16.0.20715.20004` (x64) at 96 DPI. Two **+** clicks created two distinct pages in the second test notebook’s Quick Notes destination. Each page opened in its own tab in the original editor, and each new tab received keyboard focus. The verifier also resolved the user's configured Quick Notes section. It wrote the test pages only in the disposable notebook and exited with code 0.

The Windows control test also exercised delayed navigation, an unavailable destination, and retrying a created tab after navigation failed.

The next live checks cover physical mouse and keyboard input, native page-list and Back navigation, resizing, ribbon collapse, and movement between displays with different scale settings. Before-and-after screenshots of test content are still pending.

In the first hands-on run, the native page-list and Back sequence displayed Schedule A and Dashboard A with their matching tabs highlighted. Dashboard B also highlighted. During ribbon changes the strip moved from Y=216 through Y=162 and Y=113 to Y=78. Each sample retained a 32-pixel row and passed all three hit tests at 96 DPI. The operator reported typing and mouse interaction worked and identified one failure: neither F6 nor Tab reached the strip from the page. The log ended at Ctrl+C before the marker and cleanup checks. Window-width and mixed-DPI changes were not recorded. A subsequent measurement still found a 32-pixel gap. Resize the host to let OneNote lay it out again after an interrupted run.

The revised prototype provides Ctrl+Alt+F6 as a direct keyboard entry point. A controlled window test starts in an editor field, focuses the strip, selects an adjacent tab with Right and Enter, then verifies Escape returns without navigating. Live keyboard results require the desktop retest.

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
| New tab and navigation | From another notebook, click **+** twice. Each click creates one page in the configured Quick Notes section, opens it in the same editor and focuses its new tab. Keep existing pages via the arrow beside **+**. Alternate among three pages across two notebooks and verify that the active page matches the selected tab. |
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

- Decide how reordering and pinning affect the working set.
- Verify OneNote references after renaming or moving pages, closing and reopening notebooks, and deleting pages. Exercise the resident prototype across a full desktop OneNote restart.
- Test page restoration separately from caret, selected text and scroll offset. Record what OneNote restores itself and what the add-in can reliably restore.
- Evaluate whether users need separate working sets per window and folders for organizing bookmarks.
