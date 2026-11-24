# Working on MoreNotes

MoreNotes is one public repository for independently released add-ins for the Windows desktop edition of Microsoft OneNote.

- Keep an add-in's implementation, requirements and prototype evidence under its directory in `addins/`. Introduce shared code when at least two implementations need it. Test shared changes against their consumers.
- Keep versions and installer identities separate. Use `<add-in-id>/v<version>` tags and link directly to each add-in's release from the root catalog. Test coexistence and removal when introducing a second installable add-in.
- Distinguish documented behavior, observations from a particular OneNote build, and hypotheses. Record Office build, bitness and display scaling with layout findings. Keep real notebook names, page IDs, page links and contents out of public diagnostics and screenshots.
- For the MoreTabs prototype, use OneNote's existing editor in the associated window. The required strip belongs below the ribbon and Quick Access Toolbar and above the native notebook/section navigation row.
- Treat exact placement, durable references and restoration of editing position as open questions. If placement fails, document the tested approach and visible failure, then bring alternatives to the maintainer before changing the design.
- Check each release's changed files and test the affected components before tagging. Report which behaviors were observed on Windows and which still need a live test.
