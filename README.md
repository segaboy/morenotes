# MoreNotes

<img src="assets/brand/morenotes-logo.png" alt="MoreNotes" width="420">

*It's OneNote, but More!*

MoreNotes is a collection of independently installable add-ins for Microsoft OneNote on Windows desktop. Each add-in has its own version, downloads, and release notes in this repository.

## Add-ins

| Add-in | What it does | Download |
| --- | --- | --- |
| [MoreTabs](addins/page-tabs/README.md) | Keep pages from different notebooks open as tabs in one OneNote window. Bookmark the ones you come back to. | In development |

The first MoreTabs [prototype milestone](addins/page-tabs/README.md#first-milestone) tests whether a tab strip can occupy the space below the Quick Access Toolbar and above OneNote's native notebook/section row. There is no installer yet.

## Releases

Each add-in lives under `addins/<id>/`. A release tag such as `page-tabs/v0.1.0` identifies one add-in and its version. When publishing, attach that add-in's installer and release notes to the tag's GitHub Release and update its download link in the table above. A user can install, update, and remove one add-in independently. Changes to shared code should be tested against each consumer.

## Artwork

The [collection logo](assets/brand/morenotes-logo.png), [standalone icon](assets/brand/morenotes-icon.png), and [Windows icon](assets/brand/morenotes-icon.ico) are in `assets/brand/`.
