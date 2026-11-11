# SuppRef

A Word add-in that inserts **auto-numbered figure citations** (main, supplementary, extended data, tables…), the way a reference manager inserts references, written with Claude.

- Insert citations such as *Figure 2.B* or *Supplementary Figure 3.A–C*. Each one remembers which figure and which panels it points to.
- Each series (main figures, supplementary figures, and so on) is numbered on its own.
- Press **Refresh** to renumber every citation, based on the order in which figures are first cited.
- Move, add or delete text freely, then Refresh again.
- When the manuscript is final, use **Unlink all** to turn every citation into plain text.

Everything stays inside the `.docx`. The add-in makes no network calls other than loading its own files and Office.js.

## Quick start: use SuppRef on your computer

You don't need GitHub, a code editor or any technical setup. The add-in already runs online; you only tell Word where to find it.

**What you need:** Word for Mac (Microsoft 365, or Office 2021 or later), and the ready-to-use `manifest.xml` file. Download it from this repository. If the file contains the text `{{BASE_URL}}`, it is not ready yet: ask the person who maintains SuppRef for the finished file.

**On a Mac:**

1. Open Finder and press **Cmd+Shift+G**.
2. Paste this path and press Enter:
   ```
   ~/Library/Containers/com.microsoft.Word/Data/Documents/
   ```
3. If there's no folder called `wef` there, create one (all lowercase).
4. Copy `manifest.xml` into the `wef` folder.
5. Quit Word completely (**Cmd+Q**) and open it again.
6. Open any document. Click **Home → Supp. Figures**, at the right end of the Home ribbon. If the button isn't there, click **Home → Add-ins**, then **More Add-ins**. Open the **My Add-ins** tab and choose **SuppRef** under *Developer Add-ins*. On older Word versions, the path is **Insert → Add-ins → My Add-ins**.

   Don't use **Tools → Templates and Add-ins**. That menu is for `.dotm` templates (for example Zotero's), not for this kind of add-in.

The SuppRef panel opens on the right. It works with every document you open, and each document keeps its own figure list.

**In Word on the web:** open a document, go to **Insert → Add-ins → Upload My Add-in**, and choose `manifest.xml`.

**To remove it:** delete `manifest.xml` from the `wef` folder, then restart Word.

**If something goes wrong:**

- *The button or panel doesn't appear.* Make sure Word was fully quit, not just its window closed. Check that the file is named `manifest.xml` and sits directly inside `wef`.
- *The panel is blank or shows an error.* Check that you're online. The panel loads from the internet each time it opens, but your documents are never sent anywhere.
- *Your figure list disappeared after reopening a document.* The list is stored in the document, so it's only kept when you **save** the document.

How to use the panel is described under [Using it](#using-it) below.

---

*The sections below are for the person who hosts and maintains SuppRef.*

## Files

| File | Purpose |
|---|---|
| `manifest.xml` | Sideloaded into Word; points at your GitHub Pages URL |
| `taskpane.html` / `.css` / `.js` | The panel |
| `core.js` | Pure logic: rendering, tags, numbering (no Office.js) |
| `commands.html` | Required stub for the ribbon button |
| `assets/` | Icons |
| `tests/` | Unit tests for `core.js` (`tests/run.sh`; uses Node, or macOS's built-in JavaScriptCore) |

## Install and host (maintainer)

Once these steps are done, commit the finished `manifest.xml`, with your real URL in place of `{{BASE_URL}}`. Users can then follow the Quick start above as-is.

1. Push this repository to GitHub as a public repository. Enable GitHub Pages under **Settings → Pages**, with source `main` / root. Wait until `https://<user>.github.io/<repo>/taskpane.html` loads in a browser.
2. In `manifest.xml`, find and replace every `{{BASE_URL}}` with `https://<user>.github.io/<repo>/`. Keep the trailing slash.
3. Copy `manifest.xml` into the following folder, creating it if needed:
   ```
   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/
   ```
4. Quit and reopen Word. Open a document, then go to **Home → Supp. Figures**. If the button isn't there, use **Insert → Add-ins → My Add-ins → Developer Add-ins → SuppRef**.

For Word on the web, upload the same manifest under **Insert → Add-ins → Upload My Add-in**.

## Update

- Push your changes. GitHub Pages redeploys in about a minute.
- If Word keeps showing an old version, quit Word and empty these folders:
  - `~/Library/Containers/com.microsoft.Word/Data/Library/Caches/`
  - `~/Library/Containers/com.microsoft.Word/Data/Library/Application Support/Microsoft/Office/16.0/Wef/`

  Check these paths against Microsoft's "Clear the Office cache" page; they can change between Office versions.
- Bump `<Version>` in the manifest only when the manifest itself changes, then copy it into `wef/` again.

## Uninstall

Delete `manifest.xml` from the `wef` folder and restart Word.

## Using it

1. **Add figures** in the panel. Choose a series, a short label (e.g. `droso`) and an optional description. Labels must be unique across all series. You can rename a figure or move it to another series at any time with ✎; citations follow automatically.
2. Put the cursor where the citation belongs. Pick the figure and its panels, using the toggle buttons or by typing e.g. `A-C,E`, then press **Insert citation**.
3. Tick **Number only** to insert just `3.A`. This is useful when you write the plural yourself: "Supplementary Figures 3.A and 5.B".
4. To change a citation, click inside it. The form switches to **Update citation**. If the switch doesn't happen automatically, press **Load selection**.
5. At the start of each legend, press **Insert legend label**. Legend labels follow the numbering but never affect it.
6. **Import** converts plain-text placeholders into citations, and creates any figures that are missing. Each series has its own placeholder word (see the table below). For example, `Figure_droso_A` becomes a supplementary citation and `MainFig_overview` a main one.
7. **Settings** controls each series separately: its name, placeholder word, prefix, legend prefix, number and panel styles, separators, ranges and ordering. Pick the series at the top of Settings. **Add** creates a new series and **Delete** removes an empty one. **Save as my defaults** stores all series and options on this computer for new documents.

### Series

New documents start with four series. You can rename, reformat or delete them, or add your own (for example *Figure S1* style, or *Source Data*).

| Series | Citation | Placeholder |
|---|---|---|
| Main figures | Figure 1.A | `MainFig_label_A` |
| Supplementary figures | Supplementary Figure 1.A | `Figure_label_A` |
| Additional figures | Extended Data Figure 1.A | `ExtFig_label_A` |
| Supplementary tables | Supplementary Table 1 | `SuppTable_label` |

Each series is numbered from its own first number, by first citation or in manual order. Documents created before series existed open with their figures under *Supplementary figures*, with their formatting kept.

### Good to know

- **The figure list is stored in the document.** It is written into the file when you **save the document**. If you close without saving, recent figure changes are lost.
- **Track Changes:** Refresh switches Track Changes off while it renumbers, then switches it back on, so renumbering doesn't appear as revisions. This needs a recent Word (WordApi 1.4). Older versions show a warning instead.
- **Accept or reject tracked changes before refreshing.** A citation inside a tracked deletion can still be counted.
- **Co-authors without the add-in** see normal text. Citations can't be typed into, but they can be deleted. Citations those co-authors type themselves are plain text and aren't renumbered.
- Citations in footnotes, endnotes, headers, footers or text boxes aren't supported yet.
- **Unlink all** can't be undone by the add-in, so save a copy of the document first.

## Development

No build step: the repository is deployable as-is. Run the logic tests with:

```sh
tests/run.sh
```

Before a release, test manually in Word for Mac:

- insert, reorder and refresh citations;
- check that a second refresh reports 0 changed;
- change each setting, for more than one series;
- mix main and supplementary citations, and check that each series counts from 1;
- move a figure to another series, then add and delete a series;
- refresh with Track Changes on;
- check orphans and legend-order warnings;
- save and reopen the document;
- import placeholders;
- run Unlink all;
- check that refreshing 200 citations takes under 3 seconds.
