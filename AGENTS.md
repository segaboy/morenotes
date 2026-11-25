# Working on MoreNotes

MoreNotes is one public repository for independently released add-ins for the Windows desktop edition of Microsoft OneNote.

- Keep an add-in's implementation, requirements and prototype evidence under its directory in `addins/`. Introduce shared code when at least two implementations need it. Test shared changes against their consumers.
- Keep versions and installer identities separate. Use `<add-in-id>/v<version>` tags and link directly to each add-in's release from the root catalog. Test coexistence and removal when introducing a second installable add-in.
- Publish source and release notes on GitHub. Distribute official compiled add-ins and installers through the paid storefront. For each storefront build, provide a link beside the download to its exact tagged source, including build and packaging scripts. Keep compiled output out of Git history, public CI artifacts, and GitHub Release assets.
- Distinguish documented behavior, observations from a particular OneNote build, and hypotheses. Record Office build, bitness and display scaling with layout findings. Keep real notebook names, page IDs, page links and contents out of public diagnostics and screenshots.
- For the MoreTabs prototype, use OneNote's existing editor in the associated window. The default strip follows the document width below the ribbon and Quick Access Toolbar and above the native notebook/section navigation row. The far-right arrow switches to a bottom row aligned with both visible ribbon edges. Reserve space below the editor for the row and its top padding so native controls stay usable.
- Treat exact placement, durable references and restoration of editing position as open questions. If placement fails, document the tested approach and visible failure, then bring alternatives to the maintainer before changing the design.
- Check each release's changed files and test the affected components before tagging. Report which behaviors were observed on Windows and which still need a live test.
