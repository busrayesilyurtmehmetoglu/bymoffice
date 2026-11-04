# BYM Office — API reference

This is the English reference for the public API of the `bymoffice` package, version 0.1.2. The full developer
guide is in Turkish: [kullanim-kilavuzu.md](kullanim-kilavuzu.md). Where this page is short, the guide has the
details; where the two disagree, the code (`src/index.ts` and what it exports; types in `dist/types/`) wins.

Everything listed here works in this version. What does not exist yet is listed in
[Not yet available](#11-not-yet-available) at the end. Please read that table before you build on a feature.

Contents:

1. [Overview and install](#1-overview-and-install)
2. [`mount(container, options)`](#2-mountcontainer-options)
3. [Saving](#3-saving)
4. [Real-time collaboration (`collab`)](#4-real-time-collaboration-collab)
5. [Files without a view: `open`, old formats, passwords, blank files](#5-files-without-a-view)
6. [Documents (docx)](#6-documents-docx)
7. [Spreadsheets (xlsx)](#7-spreadsheets-xlsx)
8. [Low level: models, packages, containers](#8-low-level)
9. [WebView bridge](#9-webview-bridge)
10. [Browser and platform support](#10-browser-and-platform-support)
11. [Not yet available](#11-not-yet-available)

## 1. Overview and install

BYM Office opens, shows, edits and saves Word and Excel files in the browser, written from scratch in TypeScript
with no runtime dependencies. It reads docx, xlsx, doc, xls and RTF, and detects the format from the content (never
from the file name). docx and xlsx are edited and saved losslessly: parts the engine does not understand keep their
original bytes, and only what changed is rewritten; an unchanged file is saved byte for byte as it was opened.
doc, xls and RTF open read-only, or are converted to docx/xlsx for editing (they are never written in their own
formats). The editor draws on a canvas, has toolbars in Turkish and English, and works with mouse, keyboard and
touch. The same code runs in Node and Bun without a screen.

**CDN** (defines a global `BymOffice`; no build step):

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.2/dist/bymoffice.min.js"></script>
<!-- or https://unpkg.com/bymoffice@0.1.2/dist/bymoffice.min.js -->
```

Put the version in the address: 0.x minor versions may contain breaking changes. The bundled font is loaded from
the `fonts/` folder next to the script (on the CDN too).

**npm**:

```sh
npm i bymoffice
```

```js
import { mount, open } from 'bymoffice';
```

The package ships as an **ES module and a `<script>` build only** (no CommonJS build; `require()` works only on
Node versions that can `require` an ES module). There are no runtime dependencies and no source maps.

| File | Use |
|---|---|
| `dist/bymoffice.min.js` | `<script>`; defines the global `BymOffice` |
| `dist/bymoffice.esm.js` | ES module (`import 'bymoffice'`) |
| `dist/types/index.d.ts` | TypeScript types (resolve with `moduleResolution` `bundler`, `node16` and `nodenext`) |
| `dist/bymoffice.webview.html` | Single-file bridge page for WebViews (`bymoffice/webview.html`) |
| `dist/fonts/tex-gyre-bonum/` | The bundled font (TeX Gyre Bonum, four `.otf`) and its licence (`bymoffice/fonts/…`) |
| `server/collab-server.mjs` | Dependency-free example relay for real-time collaboration (Node) |

**Bundlers and ES modules.** An ES module does not know its own address, so the bundled font cannot be found by
itself. Copy `node_modules/bymoffice/dist/fonts/` to your public files and pass the address once:
`mount(el, { file, fontsUrl: '/assets/bymoffice-fonts/' })` or
`fonts.configure({ baseUrl: '/assets/bymoffice-fonts/' })`. The Vite, React and Vue examples
([`examples/`](../examples/)) do the copy in `vite.config.js` and are tested.

**Your own server.** Serve `node_modules/bymoffice/dist/` as it is; `bymoffice.min.js` and `fonts/` must stay side by side.

**Node / Bun** (no view; `open`, the edit operations and `save` work):

```js
import { readFile, writeFile } from 'node:fs/promises';
import { open } from 'bymoffice';

const file = await open(new Uint8Array(await readFile('report.xlsx')));
await writeFile('copy.xlsx', await file.save());
```

## 2. `mount(container, options)`

```ts
mount(container: HTMLElement, options: MountOptions): Promise<MountedView>
type MountedView = MountedSpreadsheet | MountedDocument; // tell them apart by `kind`
```

Opens a file and shows it inside `container`: xlsx and xls in the spreadsheet editor, docx, doc and RTF in the
document editor. The container must have a size (for example `height: 600px`); the view fills it and follows its
size. A file that is not recognised rejects with `UnsupportedFileError` and nothing is added to the container.

```html
<div id="editor" style="height: 600px"></div>
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.2/dist/bymoffice.min.js"></script>
<script>
  (async () => {
    const mounted = await BymOffice.mount(document.getElementById('editor'), {
      file: '/files/report.xlsx',            // URL, Blob/File, ArrayBuffer or Uint8Array
      locale: 'en',
      onSelectionChange: (selection, sheetIndex) => {
        console.log(BymOffice.rangeName(selection.range), selection.active);
      },
    });
    // mounted.kind === 'spreadsheet'; mounted.view, mounted.editor, mounted.workbook
    // mounted.destroy();
  })();
</script>
```

### `MountOptions`

"S" = spreadsheets only, "D" = documents only; no mark = both.

| Option | Type | Default | Meaning |
|---|---|---|---|
| `file` | `Uint8Array \| ArrayBuffer \| Blob \| string` | — (required) | The file, or a URL to `fetch` it from. An HTTP error rejects (`Could not load the file: HTTP 404`) |
| `locale` | `string \| FormatLocale` | browser language | Language of number/date formats and of the UI: `'tr'`, `'en'`… Every language other than Turkish is shown with en-US rules |
| `readOnly` | `boolean` | `false` | No editing (no cell editor, paste, cut, delete, undo, structural commands); copy still works |
| `toolbar` | `boolean` | `true` | Show the toolbars: in a workbook the formatting toolbar; in a document the formatting toolbar and the bar with the zoom and the page number |
| `formulaBar` | `boolean` | `true` | S: show the formula bar and name box |
| `statusBar` | `boolean` | `true` | S: show the status bar (sum/average/count) |
| `calculate` | `boolean` | `true` | S: recalculate formulas after changes |
| `sheet` | `number` | file's active sheet | S: index of the sheet shown first (first visible sheet if the active one is hidden) |
| `author` | `string` | `"Author"` | S: author of notes added in this view (cut to 255 characters). D: author of tracked changes and comments |
| `onNotice` | `(n: { message: string; kind: 'error' \| 'info' }) => void` | — | Messages the view shows the user, in the UI language (also sent to its live region) |
| `onChange` | `(c: { sheets: number[]; remote?: boolean }) => void` | — | S: after every change to the workbook, undo/redo included. `remote: true` when it came from a collaborator. Also called after each finished slice of formula calculation |
| `onDocumentChange` | `() => void` | — | D: after every change made in the editor, undo/redo included |
| `onSelectionChange` | `(selection: GridSelection, sheetIndex: number) => void` | — | S: the selection changed (not called on open or on a sheet change) |
| `onSheetChange` | `(sheetIndex: number) => void` | — | S: the shown sheet changed |
| `onEditingChange` | `(s: { editing: boolean; text: string }) => void` | — | S: cell editing started, its text changed, or it ended |
| `onDrawingSelectionChange` | `(id: number \| undefined) => void` | — | S: another chart, picture or shape (or none) was selected |
| `focusScope` | `HTMLElement` | — | S: an element around the view (your own formula bar, say); focus inside it does not commit an open edit |
| `onPrintProgress` | `(p: { done; total } \| undefined) => void` | — | S: pages prepared by `print()`; `undefined` when the dialog opens or the job ends |
| `onPdfProgress` | `(p: { done; total } \| undefined) => void` | — | S: pages painted by `exportPdf()`; `undefined` when done or given up. For documents use `exportPdf({ onProgress })` |
| `fileName` | `string` | last part of the URL, or the `File` name; else `document.pdf` / `workbook.pdf` | Name the **PDF download** button uses (`report.docx` → `report.pdf`); `&F` in sheet headers/footers |
| `zoom` | `number \| 'fit-width' \| 'fit-page'` | 100%, less if the page is wider than the container | D: zoom (`1` = 100%) |
| `onPageChange` | `(page: number, pageCount: number) => void` | — | D: page in view or number of pages changed (`page` from 1) |
| `onLinkClick` | `(target: string) => boolean \| void` | — | D: a hyperlink was activated; return `false` to stop the view opening it in a new window |
| `trackChanges` | `boolean` | `false` | D: track changes even if the document does not ask for it. A document that asks for tracking is always tracked; `false` does not turn that off |
| `commentsPanel` | `'auto' \| 'open' \| 'closed'` | `'closed'` | D: the comments side panel. `'auto'` opens it when there are comments and the editor is at least 720 px wide |
| `now` | `() => Date` | system clock | D: clock for tracked-change dates (for tests) |
| `convert` | `boolean` | `false` | Open doc/RTF as docx and xls as xlsx so they can be edited; see [Old formats](#old-formats-convert-saveasdocx-saveasxlsx). Ignored with `readOnly: true` |
| `password` | `string` | — | Password of an encrypted file. Missing or wrong: the user is asked (`onPassword` or the editor's own dialog) |
| `onPassword` | `(r: { attempts: number; wrong: boolean }) => string \| null \| undefined \| Promise<…>` | editor's own dialog | Return the password; `null`/`undefined` gives up (`mount()` rejects with `PasswordError` `'cancelled'`). Called again while the password is wrong |
| `requireIntegrity` | `boolean` | `false` | Reject an encrypted docx/xlsx without an integrity check (HMAC) with `PasswordError` `'damaged'`; otherwise it opens and the user is told |
| `fonts` | `FontRegistration[] \| false` | — | Fonts to register before showing the file (see [Fonts](#fonts)); waited for up to 3 s. `false` turns the bundled font off |
| `fontsUrl` | `string \| null` | `fonts/` next to the script | Address of the package's `dist/fonts/` folder. **Required with bundlers and ES modules.** `null` turns the bundled font off. Page-wide setting |
| `saveTo` | `SaveTo \| false` | — | Where the Save button, Ctrl/Cmd+S and `saveNow()` save: see [`saveTo`](#saveto-save-back-to-where-the-file-came-from) |
| `autoSave` | `AutoSave` | off | With `saveTo`: save by itself; see [`autoSave`](#autosave) |
| `onSaveState` | `(s: SaveState) => void` | — | With `saveTo`: called when the save state changes |
| `warnOnLeave` | `boolean` | `true` | With `saveTo`: the browser asks before leaving a page with unsaved work |
| `onSave` | `(bytes: Uint8Array) => void` | — | Without `saveTo`: called with the bytes on Ctrl/Cmd+S (without `onSave` or `saveTo` the shortcut is left to the browser). With `saveTo`: called after every successful save |
| `onError` | `(error: unknown) => void` | — | Saving failed: a `SaveError` with `saveTo`, otherwise what the writer threw |
| `collab` | `string \| CollabOptions` | — | Edit together with others: see [section 4](#4-real-time-collaboration-collab) |

Invalid values of `saveTo`, `autoSave` and `collab` make `mount()` reject with `TypeError` before anything is
fetched or shown.

### The result

| Member | Spreadsheet | Document | Meaning |
|---|---|---|---|
| `kind` | `'spreadsheet'` | `'document'` | Which of the two |
| `file` | ✓ | ✓ | The `OfficeFile` shown (the converted one when `conversion` is set) |
| `workbook` / `document` | `Workbook` | `DocModel` | The model |
| `view` | `WorkbookView` | `DocumentView` | The grid / the pages |
| `editor` | `SpreadsheetEditor` | `DocumentEditor` | The shell around the view (toolbars). `editor.element` is its root element |
| `conversion` | ✓ | ✓ | `{ from: 'doc' \| 'rtf' \| 'xls', to: 'docx' \| 'xlsx', warnings: string[] }` when `convert` converted the file |
| `protection` | — | ✓ | `{ edit: 'readOnly' \| 'comments' \| 'forms' \| 'trackedChanges', exceptions: boolean }` when the document enforces protection |
| `trackChanges` | — | ✓ | `true` when changes made in the editor are tracked |
| `collab` | `MountedCollab` | `MountedDocCollab` | Set when the file is edited together |
| `save(options?)` | ✓ | ✓ | `Promise<Uint8Array>`: the file with all changes; what is being typed is applied first |
| `print(options?)` | ✓ | ✓ | Print through the browser's print dialog; resolves when it closes |
| `exportPdf(options?)` | ✓ | ✓ | `Promise<Uint8Array>`: a PDF file; see [PDF export](#pdf-export) |
| `saveNow(options?)` | ✓ | ✓ | `Promise<boolean>`: save to `saveTo`; never rejects |
| `hasUnsavedChanges` | ✓ | ✓ | Whether there is work no save holds |
| `setAutoSave(value)` | ✓ | ✓ | Turn auto-save on/off or change its times |
| `refreshPivots()` | ✓ | — | Refresh every pivot table that can be: `{ refreshed: number, refused: { name, reason }[] }` |
| `destroy()` | ✓ | ✓ | Remove the view; same as `editor.destroy()` |

You can also build the views yourself from a model: `new BymOffice.SpreadsheetEditor(container, workbook, options)`
(with toolbars) or `new BymOffice.WorkbookView(container, workbook, options)` (the grid only);
`new BymOffice.DocumentView(…)` and `new BymOffice.DocumentEditor(…)` for documents. `editor.setLocale(...)` and
`editor.setReadOnly(...)` change shell and view together.

### `WorkbookView` (spreadsheet `mounted.view`)

The most used members. Rows and columns are 0-based; methods that change the workbook apply what is being typed
first (`commitPending()`).

| Member | Meaning |
|---|---|
| `activeSheet`, `setActiveSheet(index)` | Shown sheet (a hidden sheet can be shown from code) |
| `selection`, `select(range, active?)`, `scrollTo(row, col)` | `{ active: { row, col }, range: { top, left, bottom, right } }` (a copy); select and scroll into view |
| `setLocale(locale)` | Change the format/UI language. Throws `RangeError` while a rule dialog with changes is open |
| `setCellInput(row, col, text)` | Enter text as if the user typed it (one undo step) |
| `startEditing(text?)`, `commitEditing()`, `cancelEditing()`, `isEditing` | Drive the cell editor |
| `commitPending()` | Apply what is being typed (cell, note, sheet name). Returns `false` if it could not, or if a rule dialog with changes is open |
| `clearSelection()` | Clear contents of the selection; formats and notes stay |
| `undo()`, `redo()`, `history.canUndo`, `history.canRedo` | Undo/redo; selection and sheet come back too |
| `copy()`, `cut()`, `paste(text, html?)` | Clipboard core; `copy()` returns `{ text, html, … }`. Pass `html` from `copy()` to paste your own copy with formats and formulas |
| `readOnly`, `setReadOnly(readOnly)` | Switch read-only without rebuilding; history is kept |
| `insertRows(at, count)`, `deleteRows(at, count)`, `insertColumns(at, count)`, `deleteColumns(at, count)` | Throw `RangeError` when refused |
| `format(patch)` | Format the selection: `{ bold, italic, underline, strike, fontName, fontSize, color, fill, horizontal, vertical, wrapText, numFmt, border: { preset } }`; `null` removes a property |
| `selectionFormat` | Format of the active cell (toolbar state) |
| `setColumnWidth(first, last, width)`, `setRowHeight(first, last, height)` | Characters / points; `undefined` resets, `0` hides |
| `sort(descending)`, `toggleFilter()`, `setColumnFilter(col, choice)` | Sort and auto-filter |
| `merge(options?)`, `requestMerge(options?)`, `unmerge()`, `selectionMerged` | Merge cells; see [Merge](#merge-cells-and-freeze-panes) |
| `frozenPanes`, `freezePanes(rows?, cols?)`, `unfreezePanes()` | Freeze panes |
| `findNext`, `replace`, `replaceAll`, `openFind(replace?)`, `closeFind()`, `findOpen` | Find and replace (promises) |
| `cycleReference()` | What F4 does while a formula is typed (`A1` → `$A$1` → `A$1` → `$A1`) |
| `editText`, `setEditText(text)`, `focus()` | For an external editor such as your own formula bar |
| `print`, `printing`, `exportPdf`, `downloadPdf`, `exportingPdf` | Print and PDF |
| `printPreview(options?)`, `closePrintPreview()`, `previewing`, `getPrintPreviewState()` | Print preview |
| Sheets: `insertSheet`, `renameSheet`, `deleteSheet`, `startRename`, `isRenaming`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `hiddenSheets` | See [Sheets](#sheet-operations) |
| `protectSheet`, `unprotectSheet`, `isSheetProtected`, `protectionReason` | See [Sheet protection](#sheet-protection) |
| Notes: `setComment`, `editComment`, `isEditingComment` | See [Notes](#notes) |
| Rules: `setValidation`, `validationAt`, `openValidationDialog`, `conditionalRules`, `addConditionalRule`, `updateConditionalRule`, `removeConditionalRule`, `moveConditionalRule`, `clearConditionalRules`, `openConditionalDialog`, `isDialogOpen`, `ruleStrings` | See [Validation and conditional formats](#data-validation-and-conditional-formats) |
| Charts: `charts`, `addChart`, `updateChart`, `moveChart`, `removeChart`, `selectChart`, `selectedChart`, `openChartDialog`, `copyChart`, `cutChart`, `pasteChart`, `duplicateChart`, `hasChartCopy`, `getChartState` | See [Charts](#charts) |
| Pivots: `pivots`, `refreshPivot`, `refreshAllPivots`, `insertPivot`, `openPivotDialog`, `pivotFields`, `setPivotFields`, `openPivotFields`, `closePivotFields`, `deletePivot` | See [Pivot tables](#pivot-tables) |
| `getViewState()`, `render()` | `{ scrollX, scrollY, width, height, devicePixelRatio }`; redraw now (charts re-read their cells) |
| `destroy()` | Remove DOM and listeners |

The view adds one `<style data-bymoffice>` element; its class names start with `bym-`. No text from a file is ever
inserted as HTML.

## 3. Saving

### `save()` and `onSave`

```js
const mounted = await BymOffice.mount(el, { file, onSave: (bytes) => upload(bytes) }); // Ctrl/Cmd+S
const bytes = await mounted.save();   // any time; Uint8Array
```

`mounted.save()` and `file.save()` write the changes and return the bytes:

- **Only what changed is rewritten.** Untouched sheets, pictures, theme, macros and every part the engine does not
  know keep their original bytes; in a changed sheet, untouched rows are copied byte for byte. BOM and CRLF line
  ends of a part are kept: only the changed element's bytes change.
- **Nothing changed** → the saved file is byte for byte the opened one. A change that was **undone** gives the
  original bytes too, even with saves in between ("insert sheet → save → undo → save"). Exception: an undone *cell*
  edit rewrites that row from the model, so the bytes may differ (same content).
- A changed sheet gets the "recalculate on open" flag (`fullCalcOnLoad`) and loses `calcChain`; computed results are
  written too. New styles are only **added** to the style sheet. Edited text cells are written as `inlineStr`.
- The shown sheet becomes the file's active sheet only together with another change.
- **What is being typed is saved too:** `mounted.save()` (and Ctrl/Cmd+S, the bridge's `save`) calls
  `view.commitPending()` first. If a cell input is refused (validation), `save()` rejects with that error. If a note
  or sheet name cannot be applied, or a rule dialog with changes is open, `save()` rejects and nothing is saved.
  `file.save()` / `file.saveAsXlsx()` do not know the view: call `view.commitPending()` yourself first.
- A UTF-16 part that is rewritten is written as UTF-8.
- Not kept in rewritten rows (kept in untouched rows): the `spans` hint, phonetic (furigana) text, rare rich-text
  font details such as `outline`/`shadow`.
- Saved files are checked against an independent reader on every change; they were not checked automatically in
  Excel itself.

### `saveTo`: save back to where the file came from

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.2/dist/bymoffice.min.js"></script>
<script>
  BymOffice.mount(document.getElementById('editor'), { file: '/files/42', saveTo: true });
</script>
```

With `saveTo` the toolbar gets a **Save** button (hidden while read-only and for files that cannot be written;
none with `toolbar: false`, but the keyboard and API still work), **Ctrl/Cmd+S** saves (listened for on the
container element; the browser's "save page" dialog never opens), and:

- `mounted.saveNow(options?: { overwrite?: boolean }): Promise<boolean>` — `true` when saved or when there was
  nothing to save (no request is sent), `false` when it could not be (why: `onNotice`, `onSaveState`, `onError`);
  always `false` without `saveTo`. Never rejects. `overwrite: true` sends no `If-Match` (also drops an `If-Match` you
  put in `headers`): use it only on the user's explicit decision after a conflict.
- `mounted.hasUnsavedChanges` — changes no save holds (undo and redo are changes too), **or** typed text not yet
  applied (open cell, note, sheet name, rule dialog with changes). `onSaveState`'s `dirty` counts only the first.
- At most one request is in flight; a save requested while one runs waits for it (and saves once more if something
  changed meanwhile). Calls from inside your callbacks do not start a second request.
- `destroy()` cancels a running request; no callback follows.

| `saveTo` value | Meaning |
|---|---|
| `true` | The address the file was fetched from (`file` must be a URL string) |
| `'/files/42'` | That address: `http:` / `https:` or relative to the page |
| `{ url?, method?, headers?, credentials?, timeoutMs?, allowConverted?, onConflict? }` (`SaveToRequest`) | `url`: default the fetch address. `method`: `'PUT'` (default) or `'POST'`. `headers`: added headers (CSRF token, `Authorization`…); one with the name of an engine header replaces it. `credentials`: `'omit'`, `'same-origin'` (default), `'include'`. `timeoutMs`: default `120000`, covers the upload too. `allowConverted`: default `false`. `onConflict`: `'ask'` (default) or `'fail'` |
| `(bytes, info) => void \| Promise<void>` (`SaveToFunction`) | Your own saver. `info: SaveInfo = { format: 'docx' \| 'xlsx', convertedFrom?, name?, overwrite?: true, auto?: true }`. Resolving = saved; rejecting/throwing = failed (`refused`). Times out after 10 minutes. The editor does not ask about conflicts for a function |
| `false` / not given | No saving through `saveTo` |

`mount()` rejects with `TypeError` for an invalid value: `true` (or a request without `url`) for a file given as
bytes/Blob; a non-`http(s):` address (`javascript:`, `data:`, `file:`, `blob:`, an address with user/password, an
empty string); a method other than `PUT`/`POST`; invalid `credentials`; `timeoutMs` that is not a positive number;
non-boolean `allowConverted`; `onConflict` other than `'ask'`/`'fail'`; header names or values with forbidden
characters. Header values must be Latin-1 (up to U+00FF, no line breaks or control characters; tab allowed) — encode
others with `encodeURIComponent`. A header named `__proto__` is refused.

**The request** (the body is the raw file, not a multipart form):

```
PUT /files/42
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
X-Bym-Format: xlsx
X-Bym-Converted-From: xls        (only for a file opened with convert: true)
If-Match: "abc123"               (only if the GET gave a strong ETag and the file is saved to the same address)

<the file's bytes>
```

- Content types: docx `application/vnd.openxmlformats-officedocument.wordprocessingml.document`, xlsx
  `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
- Any 2xx means saved; the response body is not read. A new `ETag` in the response is used for the next save;
  without one, later saves go without `If-Match`.
- Redirects are **not followed** (`redirect: 'manual'`): a 3xx, or a response whose `url` is not the request's, is a
  failure (`redirected`).
- No `If-Match` when the file was fetched through a redirect, for weak ETags (`W/"…"`), or when saving to another
  address. Saving to another origin needs CORS: expose `ETag` and allow `If-Match` and `X-Bym-*` in preflight
  (UNSURE: cross-origin saving was not tested).
- "Same address" compares origin + path + query; `#fragment`, a trailing empty `?`, an explicit default port, `./` and
  host-name case do not count; path case, `%` encoding and a trailing `/` do.

| Result | `SaveError.code` | Notes |
|---|---|---|
| 2xx | — | "Saved" (`onNotice`, `kind: 'info'`) |
| 412 | `conflict` | Nothing is overwritten; the editor then asks "Overwrite / Cancel" unless `onConflict: 'fail'` |
| 3xx, or `response.url` differs | `redirected` | Not followed (an expired session's login redirect must not take the file) |
| converted file to its original address without `allowConverted` | `converted` | No request is sent |
| 401, 403 | `denied` | |
| 423 | `locked` | |
| other non-2xx | `http` | `status` holds the number |
| network error | `network` | |
| timeout | `timeout` | |
| your function rejected/threw; or `fetch` threw at once | `refused` | first 200 characters of the message are shown |
| the file could not be written | `write` | |
| typed text could not be applied | `pending` | the view says why |

`SaveError extends Error` has `code: SaveErrorCode`, `status?: number` and `cause?: unknown`; its message is English
and holds nothing of the file. User-facing notices are in the UI language and shown as text.

`onSaveState(state: SaveState)` is called only when the state changes: `{ dirty, saving, lastSavedAt?, error?, auto?,
autoPaused?, autoWaits? }`. `error` is the reason of the last failure in the UI language and is cleared when a save
starts or succeeds. `autoWaits` is `'room'` or `'offline'` while a collaboration room holds auto-save back on purpose.

**Conflict prompt.** With `onConflict: 'ask'` (default) a 412 calls `onError` (`conflict`) first, then shows an
`alertdialog` over the editor's root element (`mounted.editor.element`): "Overwrite" / "Cancel". "Overwrite" is the
same as `saveNow({ overwrite: true })` (one request, no `If-Match`; other condition headers you set in `headers`, such
as `If-Unmodified-Since`, are still sent). The overwrite permission belongs only to the request it starts. "Cancel" or
Escape sends nothing. Overwrite does not react for the first 400 ms (`aria-disabled="true"`; in tests wait for
`[data-bym-action="overwrite"]:not([aria-disabled])`), arrow keys do nothing, and auto-repeated keys never answer.
No question is asked while the view is read-only. With `onConflict: 'fail'` the editor only reports; ask the user
yourself and call `mounted.saveNow({ overwrite: true })`. The engine never reloads or merges the server's file.

**Converted old formats.** A doc/RTF/xls opened with `convert: true` is saved as docx/xlsx with
`X-Bym-Converted-From`; a function saver gets `info.convertedFrom` and `info.name` with the new extension. It is
**not** saved to the address of the original unless `saveTo: { allowConverted: true }`; otherwise every save fails
with `converted` and no request (when the file was given as bytes, an address ending in `.doc`, `.rtf` or `.xls` also
needs `allowConverted`). Without `convert` these files are read-only: no Save button, `saveNow()` returns `false`.
A document whose protection allows no editing is not saved either.

Server code for Laravel, Django, ASP.NET Core and Node, CSRF headers and request size limits:
[Saving to your server](server.md). UNSURE: not tested behind a proxy (nginx, IIS), across origins, or on real mobile
devices; `beforeunload` may not be shown on iOS Safari and in WebViews.

### `autoSave`

```js
const mounted = await BymOffice.mount(el, { file: '/files/42', saveTo: true, autoSave: true });
mounted.setAutoSave(false);              // off
mounted.setAutoSave({ delayMs: 1000 });  // on again, other times
```

| Value | Meaning |
|---|---|
| `true` | `{ delayMs: 2000, maxWaitMs: 30000 }` |
| `delayMs` | Save this long after the **last** change (each change restarts it). 250–600 000 ms, default 2000 |
| `maxWaitMs` | Save at the latest this long after the first unsaved change. Not below `delayMs`; default 30 000 (or `delayMs` if larger) |
| `unconditional` | `true`: also save when no `If-Match` can be sent. **Can silently overwrite others' changes.** Default `false` |
| `false` / not given | Off |

Invalid values, or `autoSave` without `saveTo`, make `mount()` reject (and `setAutoSave()` throw) with `TypeError`.

Auto-save uses the same path as the Save button (one request, `If-Match`, same callbacks), with these differences:

- **Only with a condition:** a save to an address is sent by itself only with `If-Match` (a strong ETag from the
  fetch/save response, or your own `If-Match` header). Otherwise nothing is sent, `onNotice` says so once and
  `onSaveState` has `autoPaused: true`; the Save button still works, and a save whose answer brings a strong ETag starts
  auto-save again. Most frameworks do not put an ETag on a `PUT` response — your server must send one.
  A function saver is not restricted (check `info.auto`).
- **Never applies or interrupts typing:** nothing is sent while a cell, note, sheet name, rule dialog, IME composition,
  drag or conflict prompt is open (`view.hasPendingInput()`); it looks again one delay later.
- **Conflicts are not asked about:** a 412 is reported and auto-save stops until the user saves.
- **Retries:** `network`, `timeout` and 5xx are retried after 5, 10, 20… s (at most 5 min apart), reported once.
  Every other failure stops auto-save until the user saves. A successful save resets everything.
- Never sent: a converted file without `allowConverted`, a read-only view, a file that cannot be written.
- A successful auto-save is not announced with `onNotice` (only `onSaveState`), except the first after a problem.
- When the page becomes hidden (`visibilitychange`), what waits is sent at once (the browser may still cut it off).
- `destroy()` removes the timer.

UNSURE: not tested with real devices, IME and on-screen keyboards, or mobile tabs sent to the background.

## 4. Real-time collaboration (`collab`)

Several people edit the same **xlsx or docx** together. The relay is "dumb": it only numbers steps and replays the log
to newcomers; merging happens in each browser. xls/doc/RTF, files opened with `convert: true`, and documents that
track changes or contain tracked changes are not edited together: a notice says so, the file is edited alone, and
`mounted.collab` is `undefined`.

```js
const mounted = await BymOffice.mount(el, {
  file: '/files/42.xlsx',
  saveTo: true,
  collab: { url: 'wss://example.com/collab', room: '42', user: { name: 'Ada', color: '#1a73e8' }, token },
});
mounted.collab.state;   // 'connecting' | 'online' | 'offline' | 'detached' | 'closed'
```

`collab` is a `ws://`/`wss://` address or a `CollabOptions` object; anything else makes `mount()` reject with `TypeError`.

| `CollabOptions` field | Meaning |
|---|---|
| `url` | Relay address (`ws://…` / `wss://…`) |
| `transport` | Your own channel instead of `url`: `{ send(text), onMessage(cb), onState?(cb: 'open' \| 'closed'), close() }` (WebSocket service, `BroadcastChannel`…) |
| `room` | A room at the address, when one address serves several files |
| `user` | `{ name?, color?: '#RRGGBB' }` shown to the others (names cut to 64 characters; what someone says of itself, not proof) |
| `token` | Sent to the relay in the first (`join`) message for authorisation; never shown or logged |
| `onState(state)`, `onPeers(peers)`, `onPeople(people)` | State changes; other participants (with selection/presence); everybody in the file, this participant first (at most 200) |
| `saveWatchMs` | With `autoSave`: how long changes wait for the one participant who saves for all before this side saves itself. At least 1000; default 5 auto-save delays, at least 15 s, ±25% |

`mounted.collab` (`MountedCollab`; `MountedDocCollab` for documents, whose peers carry a document position):
`state`, `reason?` (why a `closed` session ended: `'version'` — reload the page, `'refused'`, `'peer'`, `'stale'`),
`rev`, `pending` (local changes waiting for the relay), `peers`, `user`, `people`, settable `onState`/`onPeers`/`onPeople`,
`saved(bytes, etag?)` (for an app that saves with `save()` itself, without `saveTo`), `reload()` (for `detached`: take
the stored file and join; unsaved local changes are lost), `leave()`. When the file is reloaded in place, `file`,
`workbook`/`document`, `view` and `editor` on the same result object are replaced.

**Behaviour that matters:**

- Same cell: the later step in relay order wins. A step whose target is gone is dropped and reported (`onNotice`).
  A merge the engine cannot prove correct is dropped and reported, never guessed.
- Undo/redo affect only your own steps. Not undoable in a room: sheet deletion, rule and chart commands (xlsx);
  list, link and table-structure changes and a deleted comment (docx).
- Offline: editing continues, changes queue (`pending`; one warning above 500) and are sent on reconnect
  (backoff 0.5 s → 30 s). A read-only view joins as a watcher.
- Formulas are calculated on each side; values are not sent.
- Saving: with `autoSave`, one participant saves for the room (`SaveState.autoWaits: 'room'`). A save is "clean" only
  with proof (the stored file is fetched, up to 32 MiB, and compared). The saved xlsx carries the stamp
  `BymOfficeCollab = "<ep>:<rev>"` in `docProps/custom.xml` (only when `collab` is set). In `detached` the editor asks
  how to fork (download my copy / load stored file / overwrite with mine / cancel), not when read-only or with
  `onConflict: 'fail'`.
- Refused in a spreadsheet room (with a notice, often `CollabUnsupportedError` with `what`): copying a sheet; rules
  carried by copy/paste; changing array-formula or data-table cells; formatting a cell with a gradient fill; text with
  characters XML cannot hold; a step over 900 KB; pictures and shapes (`what: 'picture'`); pivot refresh, creation,
  field changes and deletion; page setup; setting/removing protection.
- In a document room these work: typing, plain paste, delete, Enter, character and paragraph formatting (tab stops
  too), lists, page breaks, links, tables, headers/footers/footnotes, text boxes, comments. Refused: inserting/resizing
  pictures, accept/reject and toggling tracked changes, the default tab stop, page breaks or links inside text boxes;
  rich paste does not carry pictures, fields or footnotes. Known limit: a late joiner can write a file that differs in
  bytes (split content control `w:id`, comment part order) — only an extra save results.
- Protocol version 7 (the `pv` of the wire protocol; not exported as a constant): participants with another package version get `closed` / `reason: 'version'`.

**Relay.** `server/collab-server.mjs` is a dependency-free example relay for Node (in memory only; binds 127.0.0.1 by
default): `node node_modules/bymoffice/server/collab-server.mjs --port 8787 [--host …] [--path /prefix] [--origin URL]`.
As a module: `createCollabServer(options).listen(port, host)` / `.attach(httpServer)`, with `authorize(req)`,
`origins` and many limits (1 MiB per message, 64 connections per room, rate limits); it is not exported through the
package's `exports` map, so run or import it by path. In a page, `BymOffice.collab.memoryServer()` gives an in-page
relay (tests, demos) and `BymOffice.collab.broadcastTransport(name)` connects tabs with `BroadcastChannel`;
`socketTransport(url, options)` is the WebSocket transport. Any backend can implement the relay by following the
wire protocol: [collab-sozlesmesi.md](collab-sozlesmesi.md) (Turkish). UNSURE: no production relay, no TLS setup and
no two-machine test were tried.

## 5. Files without a view

### `open(input, options?)` and `OfficeFile`

```ts
open(input: Uint8Array | ArrayBuffer | Blob, options?: OpenOptions): Promise<OfficeFile>
```

`OpenOptions`: `password?`, `passwordOf?: OfficeFile` (an open file whose password opens this one — another version
of the same file), `signal?: { aborted }` (cancels decryption: `PasswordError` `'cancelled'`), `requireIntegrity?`.

| `OfficeFile` member | Meaning |
|---|---|
| `format` | `'docx' \| 'xlsx' \| 'doc' \| 'xls' \| 'rtf'` (from the content) |
| `kind`, `container` | `'document' \| 'spreadsheet'`; `'ooxml' \| 'cfb' \| 'rtf'` |
| `package?`, `cfb?` | `OpcPackage` (docx/xlsx) or `CfbFile` (doc/xls) |
| `encryption?`, `warnings?` | How the file was encrypted; `[{ code: 'no-integrity', message }]` |
| `workbook()` | Spreadsheet model (xlsx, xls); rejects with `UnsupportedFileError` for documents |
| `document?()` | Document model (docx, doc, RTF) |
| `save(options?)` | Bytes. docx/xlsx with model changes written; doc/xls/RTF give their original bytes |
| `saveAsDocx?(options?)`, `saveAsXlsx?(options?)` | The file as a new docx / xlsx (for docx/xlsx the same as `save()`) |
| `passwordProtected`, `setPassword(password \| null)`, `close()` | Password handling (below) |

Errors: `UnsupportedFileError` (not a Word/Excel file, Word 95 and earlier), `PasswordError`, `PartTooLargeError`,
and `Error` for damaged archives ("Corrupt ZIP archive: …").

### Old formats: `convert`, `saveAsDocx`, `saveAsXlsx`

doc, RTF and xls are never written in their own formats. Two ways to get an editable file:

- `mount(el, { file, convert: true })`: doc/RTF become docx and xls becomes xlsx before showing; editing and `save()`
  then work on the new format. `mounted.conversion = { from, to, warnings }` lists (in English, with counts) what the
  new file does not carry, and `onNotice` is told. If conversion fails the file is shown read-only and `onNotice`
  says so (this path was tested only through the shared function and the bridge). Without `convert` these files are
  read-only and `setReadOnly(false)` has no effect.
- `file.saveAsDocx()` / `file.saveAsXlsx()` in code. `buildXlsx(workbook, { report })` and `buildDocx(doc, { report })`
  write a model as a new package and report warnings (`report.warnings`), e.g. xls charts, macro sheets, EMF/WMF
  pictures and external links are not written. Shapes from doc/RTF lose their geometry; theme colours become plain.
  `convertForEditing(file)` and `openConverted(file, open)` are the shared helpers.

### Password-protected files

| Scheme | Read | Write |
|---|---|---|
| docx/xlsx ECMA-376 Agile (AES-128/192/256, SHA-1/256/384/512, HMAC) | ✓ | ✓ — always AES-256, SHA-512, 100 000 spins, HMAC |
| docx/xlsx Standard (AES, SHA-1) | ✓ | written back as Agile |
| xls/doc XOR obfuscation, RC4, RC4 CryptoAPI | ✓ (read-only view) | converted docx/xlsx are written as Agile |
| Extensible, certificate, DES/3DES/RC2, RC4 in Standard | `PasswordError` `'unsupported'` | — |

- Give the password with `password` (`open` and `mount`), `onPassword` or the editor's own dialog (a field of type
  `password`; Escape gives up).
- A file opened with a password is **saved with it again** (new salts and keys each time), by `save()`, `saveNow()`,
  auto-save and the Save button. `file.setPassword(other | null)` changes that; `save({ password })` changes one save
  (`null` = without password). An empty password throws `TypeError`, more than 255 characters (`MAX_PASSWORD_CHARS`)
  `RangeError`. doc/xls/RTF reject `save({ password })`.
- `file.close()` zeroes the kept password (mount does this on `destroy()`); afterwards saving such a file rejects.
- `PasswordError.code`: `'required' | 'wrong' | 'cancelled' | 'unsupported' | 'damaged'`; `format?: 'ooxml' | 'doc' | 'xls'`.
- A file encrypted with Office's default password (`DEFAULT_PASSWORD`) just opens. `encryptionOf(bytes)` tells whether
  and how bytes are encrypted, without a password.
- An Agile file without an integrity check opens with a `no-integrity` warning, unless `requireIntegrity: true`.
- There is no UI to set, change or remove a password (API only). Sheet/workbook/document **protection** passwords are
  not encryption (see their sections).

### `createBlank(kind, options?)`

```ts
createBlank(kind: 'docx' | 'xlsx', options?: { locale?: string; sheetName?: string }): Promise<Uint8Array>
```

A new, empty file (deterministic bytes). Turkish gives the sheet name `Sayfa1`; `en`, `en-US`, `en-CA` use Letter
paper, others A4. An invalid `kind` or `sheetName` makes the promise reject with `RangeError`. Then `mount(el, { file: bytes, saveTo: … })`.

An embedding page that takes `file`, `save`, `mode` (`edit`/`view`), `lang`, `new` (`docx`/`xlsx`), `convert` and
`name` from its address and talks to its parent with `postMessage` is in `examples/html/index.html`.

## 6. Documents (docx)

docx opens in the document editor (`DocumentView` + `DocumentEditor`); doc and RTF are shown read-only (or converted).
Pages are laid out (headers/footers, footnotes, tables, lists, images, floating objects, text boxes, charts) and drawn
on canvas. Editing: caret and selection (mouse, keyboard, touch handles), typing, IME, undo/redo, clipboard,
character/paragraph formatting, styles, lists, tables, links, pictures, page and section breaks, tab stops, find and
replace. `save()` rewrites only the changed blocks.

**`DocumentView` (`mounted.view`), main members:** `pageCount`, `currentPage`, `layoutDone`, `zoom`, `setZoom`,
`zoomIn`, `zoomOut`, `goToPage`, `selectAll`, `getSelectedText`, `getSelection`, `setSelection`, `setLocale`, `focus`,
`readOnly`, `setReadOnly`, `isModified`, `markSaved`, `changeCount`, `canUndo`, `canRedo`, `undo`, `redo`, `insertText`,
`format(patch)`, `toggleFormat`, `formatParagraph`, `toggleList('bullet' | 'number')`, `indent(delta)`,
`insertTable(rows, cols)`, `insertRow`, `deleteRow`, `insertColumn`, `deleteColumn`, `insertPageBreak`,
`insertLink(target)`, `insertPicture(bytes, description?)`, `selectionFormat`, `inTable`, `inTextBox`, `commitEditing`,
`print`, `printing`, `exportPdf`, `downloadPdf`, `exportingPdf`, `getChartState()`, `destroy`. A position is
`{ story: 'body', path: [...], offset }`. Model-level editing: `BymOffice.docOps` (`insertText`, `deleteRange`,
`splitParagraph`, `setRunFormat`, `setParaFormat`, `toggleList`, `insertTable`, `insertImage`, `insertHyperlink`,
`DocHistory`, …), then `file.save()`.

Keys: Ctrl/Cmd + `+` / `-` / `0` and wheel zoom; PageUp/PageDown, Space, Home/End, arrows; Ctrl/Cmd+A, C, X, V, Z;
Ctrl/Cmd+F and Ctrl+H find/replace; Ctrl/Cmd+P print.

Not drawn or not available: in-document link navigation, right-to-left text, hyphenation, EMF/WMF/SVG pictures and
SmartArt (placeholders), fonts embedded in the docx. The clipboard carries rich content only between editors on the
same page.

### Text boxes and tab stops

- Text in text boxes can be typed and edited (`docOps.editInTextBox`, `textBoxProblem`); Escape leaves the box. Refused
  while tracking changes, for boxes locked with `noTextEdit` / `o:lock text`, and for pictures, links, breaks,
  comments and new lists inside a box.
- Tab stops: `view.tabStops`, `view.setTabStops({ set?, clear?, clearAll? } | null)`, `view.defaultTabStop`,
  `view.setDefaultTabStop(twips)` (integer 1–31 680), `view.openTabStopsDialog()`; `TabStop = { pos, align, leader? }`
  with `align` `left | center | right | decimal | bar` (and `start`, `end`, `num` from files), `leader`
  `none | dot | hyphen | underscore | heavy | middleDot`; positions ±31 680 twips, at most 64 per paragraph; invalid
  values throw `RangeError`. Bar stops are kept but not drawn.

### Find and replace in documents

`view.findNext(text, { matchCase, wholeWord, stories: 'all' | 'body', backward? })` → `Promise<DocRange | undefined>`,
`view.replace(…)`, `view.replaceAll(…)` → `{ replaced, skipped, inTextBoxes?, textBoxesKept?, inComments?, … }`,
`openFind(replace?)`, `closeFind()`, `findOpen`. Plain text only (at most 1 000 characters); case folding follows the
locale. Searches body, headers/footers, notes, text boxes and comments. Field results, equations, tracked-deleted
text and locked content are found but not replaced. Replace all is one undo step. Model: `docOps.findInDocument`,
`findAllInDocument`, `replaceText`. No bridge message.

### Track changes

- Tracking is on when the document's setting says so, when a `trackedChanges` protection is enforced, or with
  `mount({ trackChanges: true })`; `false` cannot turn off what the document asks for. `author` names the changes
  (default `"Author"`; without an author nothing counts as "your own" insertion); `now` sets the clock.
- Tracked: typing, deleting, Enter, page breaks, cut, character and paragraph formatting (`w:rPrChange`/`w:pPrChange`),
  styles, lists, table rows and columns, tables, pictures, links, blocks and rich paste. Operations that cannot be
  tracked throw `RangeError` with the message `docOps.TRACK_UNSUPPORTED` ("This change cannot be tracked yet"); the
  view reports it and nothing changes. Section breaks and section properties are refused while tracking.
- Inserted text is underlined, deleted text struck through (colour by author: `revisionColor(author)`).
- Turn on/off: toolbar button, `view.setTrackChanges(on): boolean`, `docOps.setTrackRevisions(doc, on)` (undoable;
  `RangeError(docOps.TRACK_ENFORCED)` when protection or the app forces tracking; `view.trackingEnforced`).
- List and resolve: `docOps.revisionsOf(doc)`, `docOps.listRevisions(doc, { limit? })` (`REVISION_LIST_LIMIT` 2 000) →
  `DocRevision` (`kind`: `insert | delete | moveFrom | moveTo | format | paragraphFormat | tableFormat | sectionFormat | other`,
  `author`, `date`, `text`, `canAccept`, `canReject`, …); `docOps.acceptRevision`, `rejectRevision`, `acceptAll`,
  `rejectAll` (apply the edit with `new docOps.DocHistory(doc).applyEdit(edit)`). View: `revisions()`,
  `goToRevision(n | 'next' | 'previous')`, `acceptRevision(target?)`, `rejectRevision(target?)`, `acceptAllRevisions()`,
  `rejectAllRevisions()`. Errors (`RangeError` messages): `REVISION_UNSUPPORTED` (kinds kept only as XML: `other`;
  "all" skips and counts them), `REVISION_GONE`, `REVISIONS_PROTECTED`. Accept/reject is never tracked; one undo step.
- UNSURE: the produced `w:ins`/`w:del`/property-change structures were checked with independent scripts, not by
  opening every case in Word.

### Comments

- Shown in docx, doc and RTF; edited in docx: add, edit, reply, resolve, delete (toolbar, balloon, Ctrl/Cmd+Alt+M,
  view methods, bridge commands). Each is one undo step and is never tracked. Saved to `word/comments.xml`
  (created when missing) and `word/commentsExtended.xml` (created on the first reply or resolve).
- View: `canComment`, `commentCount`, `comments()` → `DocComment[]` (`{ id, author, initials?, date?, text, parent?,
  done, story?, start?, end?, plain, canChange }`), `addComment(text, at?)`, `editComment(id, text)`,
  `replyToComment(id, text)`, `resolveComment(id, done = true)`, `removeComment(id)`, `goToComment(id | 'next' | 'previous')`,
  `startComment()`, `showComment(id, focus?)`. Model: `docOps.addComment`, `editComment`, `replyToComment`,
  `resolveComment(doc, id, done)` (here `done` is required), `removeComment`, `listComments`.
- Limits and refusals: text at most 32 767 characters (`docOps.MAX_COMMENT_TEXT`); no comments in headers, footers or
  text boxes (`COMMENT_NOT_HERE`). Editing a comment with rich content makes it plain paragraphs (`plain` says so
  beforehand). Comments are not printed.
- Comments panel: `mount({ commentsPanel: 'auto' | 'open' | 'closed' })`, Comments button, Ctrl+Alt+Shift+M,
  `view.setCommentsPanel(open, focus?)`, `toggleCommentsPanel()`, `commentsPanelOpen`. 320 px wide; below 600 px a
  drawer.

### Protection

`document.settings.protection` and `BymOffice.enforcedProtection(doc)` → `'readOnly' | 'comments' | 'forms' |
'trackedChanges' | undefined`. `readOnly`/`forms`: shown read-only, nothing turns editing on. `comments`: text locked,
comments free (`view.canComment`). `trackedChanges`: editable, every change tracked. `mounted.protection =
{ edit, exceptions }` (`exceptions`: the file has `w:permStart` ranges, which are not honoured). The engine never
checks protection passwords and never adds or removes protection. RTF and doc protection rules come from the
specifications (UNSURE: not compared with protected files).

### Print

`mounted.print(options?)` / `view.print({ dpi? })` → `Promise<{ pages, dpi }>` (default 200 dpi; values outside 72–600 are clamped), Print button
or Ctrl/Cmd+P; also read-only and for doc/RTF. Pages go to the browser's print dialog **as images** (no selectable text;
use `exportPdf()` for that). `DocumentPrintError.code`: `'unavailable' | 'too-long' | 'empty' | 'cancelled' | 'failed'`;
at most 2 000 pages. The page's CSP needs `img-src blob: data:`; a sandboxed iframe needs `allow-modals`. Not tried on
real printers, Safari, iOS or Android by hand.

### PDF export

`mounted.exportPdf(options?)` / `view.exportPdf(options?)` → `Promise<Uint8Array>` (PDF 1.4), and
`await view.downloadPdf(options?)` → `{ fileName, size }` (the **PDF download** button). Works for documents and
spreadsheets, also read-only and for unconverted doc/RTF/xls. No PDF library is used.

| `PdfExportOptions` | Default | Meaning |
|---|---|---|
| `dpi` | 150 | Resolution of image pages; values outside 72–300 are clamped (not a number, ≤ 0 or infinite: `RangeError`) |
| `image` | `'auto'` | `'auto' \| 'jpeg' \| 'lossless'` |
| `quality` | 0.85 | JPEG quality, 0 < q ≤ 1 |
| `text` | `'auto'` | `'auto'`: vector page where possible, else image; `'image'`; `'vector'` (fails with `unsupported` if a page cannot be vector) |
| `links`, `outline` | `true` | Write hyperlinks (http, https, mailto only) and an outline (headings; sheet names) |
| `onPage(info)` | — | `{ page, mode: 'vector' \| 'image', reason? }` |
| `onProgress(page, total)` | — | Progress |
| `onLinks(report)` | — | `PdfLinkReport` counts (external, internal, refused, …) |
| `title`, `author`, `date` | — | Document information |
| `signal` | — | Cancel (`cancelled`) |
| `sheets` (spreadsheets) | `'active'` | `'active' \| 'all' \| number[]` |
| `fileName` (spreadsheets) | view's `fileName` | What `&F` stands for |

- A page is **vector** (real text with an embedded font, vector drawing) only if all its text uses fonts whose files
  the engine has: the bundled font and fonts given to `registerFont` / `mount({ fonts })`. Other pages are an image of
  the page with an **invisible text layer**; in both cases text can be searched, selected and copied (Turkish too).
- `PdfExportError` `code`: `'unavailable' | 'empty' | 'too-long' | 'cancelled' | 'unsupported' | 'failed'`; an invalid
  option (not a clamped number) rejects with `RangeError`/`TypeError`. Limits: 2 000 pages; file at most 768 MB (192 MB on touch devices).
- Checked with independent PDF tools; not tried in Acrobat or browser PDF viewers (link clicks untested).

### Fonts

- `BymOffice.registerFont({ family, source, weight?, style?, as?, prefer? })` → `Promise<RegisteredFont>`; `source` is a
  URL, `ArrayBuffer`, `Uint8Array` or `Blob`; `as` makes it stand in for other family names (up to 32); `prefer` uses it
  before a device font of the same name. At most 64 faces; needs `FontFace` (rejects without it).
  `registeredFonts()`, `unregisterFont(family)`; the same three under `BymOffice.fonts`. Or pass them to
  `mount({ fonts: [...] })` (waited for up to 3 s).
- Bundled font: TeX Gyre Bonum (four `.otf`, 576 KB, GUST Font License — copy its licence file along), used for
  **Bookman Old Style** when the device lacks it, loaded lazily from `fonts/` next to the script.
  `BymOffice.fonts.configure({ baseUrl })` (same as `mount({ fontsUrl })`); `baseUrl: null`, `fontsUrl: null` or
  `fonts: false` turn it off; `BymOffice.fonts.embedded()` tells its state (`'off' | 'idle' | 'loading' | 'loaded' | 'failed'`).
- Other families fall back to metric-compatible open fonts if the device has them (Calibri → Carlito, Arial →
  Liberation Sans/Arimo, Times New Roman → Liberation Serif/Tinos, …). Code-page suffixes (`Arial Tur`) are stripped
  for display (`baseFontFamily`). UNSURE: line heights for Bookman Old Style come from Bonum's metrics, not verified
  against the real font. Bundlers were not tested.

### Charts in documents

Charts in docx (inline, floating, in headers/footers, table cells and text boxes) are drawn from the chart part's
cached values (or, for references without a cache, from the embedded workbook, read once and never changed). The same
chart types as in spreadsheets, including surface and `chartex`. Charts cannot be inserted or edited in documents; a
selection containing a chart cannot be cut. Limits: `MAX_DOC_CHARTS` 2 000, `MAX_DOC_CHART_POINTS` 4 M,
`MAX_DOC_CHART_BYTES` 64 MB, embedded workbook at most 16 MB. Not compared with real Word files.

## 7. Spreadsheets (xlsx)

### Formulas

- `BymOffice.FUNCTIONS` (the function registry) has 266 names; the calculator also answers `IF`, `IFS`, `IFERROR`,
  `IFNA`, `SWITCH`, `CHOOSE`, `AND`, `OR`, `XOR`, `NOT`, `ROW(S)`, `COLUMN(S)`, `AREAS`, `OFFSET`, `INDIRECT` (A1 only),
  `INDEX`, `ISREF`, `ISFORMULA`, `FORMULATEXT`, `SHEET(S)`, `ISBLANK`, `N`, `T`, `CELL`, `LET`, `LAMBDA`, `ISOMITTED`,
  `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY` itself. Includes XLOOKUP/XMATCH, FILTER/SORT/UNIQUE,
  VSTACK/HSTACK/TAKE/DROP/TOCOL/TEXTSPLIT…, financial, database and engineering functions.
- **No full recalculation on open:** values stored in the file are trusted. When a cell changes, only formulas that
  depend on it are recalculated, in 250 ms slices in the editor. Formulas the engine cannot calculate (unknown function,
  3-D or external references, structured references, undefined names) keep their stored value and are written back
  unchanged. Circular references evaluate to 0.
- **Dynamic arrays:** spills, `#SPILL!`, `A1#`, `@`. Whether a formula is dynamic is decided from its text. A spill
  over 1 000 000 cells gives `#SPILL!`; a workbook budget of 2 000 000 spilled cells. Known limit: dense clusters of
  colliding spills can settle in more than one stable way.
- **Locale:** files store English names with `,` and `.`. In the Turkish view users type Turkish names, `;` between
  arguments and `,` for decimals (`=TOPLA(A1;2,5)`); a comma that is not a decimal is refused. The formula bar shows
  English names with local separators. Turkish error names (`#YOK`) are accepted when typing. Functions:
  `formulaFromLocal(text, locale: FormatLocale, options?)`, `formulaToLocal(text, locale?)` (get the locale with
  `getLocale('tr')`). UNSURE: the Turkish function names were written from memory and not checked against Turkish Excel.
- Numbers inside text: the file spelling wins (`"1,234"` = 1234); only text that is no number in file spelling is
  read with the Turkish spelling (`"7,5"` = 7.5). `TEXT`, `FIXED`, `DOLLAR` and `&` output is always in file format.
  `UPPER`/`LOWER`/`PROPER` follow the locale (i/İ, ı/I in Turkish).
- Headless: `new BymOffice.CalcEngine(workbook, { locale: () => BymOffice.getLocale('tr') })`, then
  `engine.recalculate(changedCells?)`; before saving, `engine.uncalculated()` must be empty. Other members:
  `evaluateText`, `invalidate()`, `circular`, `unparsed`, `unsupported`, `mismatched`, `spillAt`.
- Typed input (`setCellInput`): numbers, percents, dates, times, currency and booleans are recognised per locale
  (`15.03.2024`, `12,50 ₺`, `DOĞRU`); `'` forces text; long digit strings (IBAN) stay text; Text-format cells keep text.

### Editing basics

Insert/delete rows and columns shift references, merges, links and names (deleted references become `#REF!`).
Cutting an array formula or pushing data off the sheet is refused. Undo keeps at most 200 steps (fewer for heavy
steps). Pastes over 1 000 000 cells are refused. Cut = copy + clear (not a move). Sorting and the auto-filter: toolbar,
`view.sort`, `view.toggleFilter`, `view.setColumnFilter`; ranges with merges, multi-cell array formulas or comment
threads, and sheets with tables, are not sorted. The headless equivalents are `setCellInput`, `clearRange`,
`cutRange` (each returns `Op[]`), `copyRange` (returns a `CopiedRange`, no ops), `pasteAt(workbook, sheet, selection, text, copied?, options?)` (returns a `PasteResult` whose `ops` you apply) and `new History(workbook).apply(ops, before, after)`.

### Sheet operations

- View: `insertSheet(at?, name?)` → index (−1 when read-only), `renameSheet(index, name)`, `deleteSheet(index)` (no
  confirmation; undo brings it back), `startRename(index)`, `moveSheet(from, to)`, `setSheetState(index, 'visible' |
  'hidden' | 'veryHidden')`, `copySheet(index, { at?, name? })`, `setTabColor(index, '#RRGGBB' | undefined)`,
  `hiddenSheets`. Each is one undo step. UI: the **+** button and the tab menu (right-click, long press, Shift+F10).
- Valid names: 1–31 characters, none of `[ ] : * ? / \`, no leading/trailing `'`, not `History`, unique ignoring case.
  Renaming rewrites formulas, names and links; deleting turns references into `#REF!`.
- Model functions (return `Op[]` or throw `RangeError`): `insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`,
  `setSheetState`, `copySheet`, `setTabColor`; checks `canInsertSheet`, `canRenameSheet`, `canDeleteSheet`,
  `canMoveSheet`, `canHideSheet`, `canCopySheet`; `copySheetLosses(wb, index)`.
- Refused: deleting the last visible sheet; delete/rename in a workbook with pivot tables, form controls or charts
  whose references cannot be followed; deleting a sheet with a table; deleting/renaming a chart sheet; every sheet
  operation when two names differ only in case; copying a chart sheet. A copy leaves out pictures, charts, tables,
  pivots and more, and reports each (`copySheetLosses`).

### Sheet protection

Protection in the file is **enforced in the view**: locked cells cannot be written, cleared or pasted over; commands
the protection forbids (formatting, row/column insert/delete, sort, filter, rules, objects, pivots) are refused and their
buttons disabled with the reason; hidden formulas do not show in the formula bar. View: `protectSheet(index?,
{ password?, allow? })` and `unprotectSheet(index?, password?)` (promises; wrong password: `code: 'wrong-password'`),
`isSheetProtected(index?)`, `protectionReason(action)`. A structure-protected workbook (`lockStructure`) refuses sheet
insert/delete/rename/move. Model operations (`applyOp`, `setCellInput`, …) **do not check** protection; use
`protectionProblem(workbook, ops)` yourself. Passwords are written as SHA-512 with salt; legacy hashes are checked.
Not available: `selectLockedCells`/`selectUnlockedCells`, passwords of protected ranges, a command to change workbook
structure protection.

### Page setup and print

- View: `pageSetup(sheet?)` → `PageSetupState`, `setPageSetup(patch, sheet?)` (paper, orientation, margins in inches,
  `scale` 10–400 or `fitTo`, centring, grid lines, headings, `printArea`, `titleRows`/`titleCols`),
  `insertPageBreak(at?)`, `removePageBreak(at?)`, `resetPageBreaks()`, `openPageSetup()`. Model: `setPageSetup(wb, index,
  patch)` and friends. Print area and titles are the names `_xlnm.Print_Area` / `_xlnm.Print_Titles`. Refused on chart
  sheets and while editing together.
- Documents: `view.sectionSetup`, `view.setSectionSetup(patch, 'section' | 'document')`,
  `view.insertSectionBreak('nextPage' | 'continuous' | 'evenPage' | 'oddPage')`; refused while tracking and in a room.
- Print: `print({ dpi?, sheets?: 'active' | 'all' })`, Ctrl/Cmd+P; honours paper, margins, scale/fit, print area,
  repeated titles, manual breaks, headers/footers (`&P`, `&N`, `&D`, `&T`, `&A`, `&F`…), grid lines and headings. A
  row/column larger than the paper is not split. Paper not stated in the file: Letter for `en`/US/Canada locales, A4
  otherwise.
- Print preview (spreadsheets only): `view.printPreview({ sheets?, page? })`, `closePrintPreview()`, `previewing`,
  `getPrintPreviewState()`; toolbar button or Ctrl+F2. Same pagination as print and PDF; only visible pages are drawn.

### Merge cells and freeze panes

- `view.merge({ center? })` merges at once and returns a `MergeReport` (`{ losses, count, notes }`: values that were
  dropped — only the top-left value stays); `view.requestMerge()` asks first if values would be lost (toolbar path);
  `view.unmerge()`. Model: `mergeRange`, `mergeReport`, `mergeProblem`, `unmergeRange`. Refused for whole rows or
  columns, multi-cell array formulas, filter-hidden rows, pivot tables, sheets with tables and chart sheets.
- `view.freezePanes(rows?, cols?)` (no arguments: above and left of the active cell), `unfreezePanes()`, `frozenPanes`;
  counted from A1. A split (not frozen) pane is kept in the file but not shown (`sheet.view.split`).

### Find and replace

`view.findNext(text, { matchCase, entireCell, lookIn: 'formulas' | 'values', within: 'sheet' | 'workbook', range?,
backward? })`, `view.replace(…)`, `view.replaceAll(…)` → `{ replaced, cells, skipped }`, `openFind(replace?)`,
`closeFind()`; Ctrl/Cmd+F, Ctrl+H. Plain text (no regular expressions or wildcards), at most 1 000 characters.
Replaced text is read again like typed input (text cells stay text); array cells and formulas that would not parse
are skipped; Replace all is one undo step, at most 200 000 cells (`MAX_REPLACE_CELLS`), and skips filter-hidden rows.
Model: `findNextCell`, `replaceAllOps`.

### Notes

xlsx cell notes: right-click menu or Shift+F2; `view.setComment(row, col, text | undefined)` (empty or `undefined`
deletes), `view.editComment(row?, col?)`, `isEditingComment`. Author from `mount({ author })` (default `"Author"`).
At most 32 767 characters (`MAX_NOTE_TEXT`). Notes travel with copy/cut/paste and sorting. Threaded comments (Excel
365) are shown and kept byte for byte but cannot be edited: note commands are disabled on such cells. Model:
`setComment(wb, index, row, col, text, { author? })`, `commentAtCell`.

### Data validation and conditional formats

- Validation rules in the file are applied: a "stop" rule blocks the value with its message, "warning"/"information"
  let it in and notify; list rules show a drop-down (Alt+Down). Pasting is not validated. A formula typed into a
  validated cell is checked by its result; when that cannot be calculated it is accepted.
  `view.setValidation(range, rule | undefined)`, `view.validationAt(row, col)`, `view.openValidationDialog()`; model
  `setValidation(wb, index, range, rule)`. Types `none | whole | decimal | list | date | time | textLength | custom`;
  operators `between` (default), `notBetween`, `equal`, `notEqual`, `greaterThan`, `lessThan`, `greaterThanOrEqual`,
  `lessThanOrEqual`; fields `formula1`, `formula2`, `allowBlank`, `hideDropDown`, `showInputMessage`, `promptTitle`,
  `prompt`, `showErrorMessage`, `errorStyle`, `errorTitle`, `error` (title ≤32, message ≤255, list ≤255 characters).
  The API takes file-syntax formulas. A rule kept only in the 2010 extension list (`x14:dataValidation`) is applied but
  cannot be changed; a `setValidation` covering it is refused.
- Conditional formats are drawn (cell rules, formulas, colour scales, data bars, icon sets, top/bottom N, above/below
  average, duplicates/unique, text, blanks, errors, dates) and edited: `view.conditionalRules(range?)`,
  `addConditionalRule(range, settings)`, `updateConditionalRule(index, patch)`, `removeConditionalRule(index)`,
  `moveConditionalRule(index, to)`, `clearConditionalRules(range?)`, `openConditionalDialog('manage' | kind | index)`;
  the same model functions take `(wb, sheetIndex, …)`. `ICON_SET_NAMES` (17; `3Stars`, `3Triangles`, `5Boxes` refused),
  `TIME_PERIODS` (10). New rules go first. Rules kept only in the extension list are read-only.
- The view's own copy/paste carries both kinds of rules (`CopiedRange.rules`, `PasteResult.skippedRules` counts what
  could not be carried). While a rule dialog with changes is open, `commitPending()`/`save()` refuse and `setLocale()`
  throws `RangeError`. Cells can be pointed at on the grid while a rule dialog is open.

### Charts

- **Drawn** (xlsx; xls read-only; docx from cache): column/bar (clustered, stacked, 100%), line, area, pie, doughnut,
  scatter, radar, bubble, stock, combination, secondary axis, data labels, trendlines, error bars, log and date axes,
  manual layout, exploded slices; surface (from above) and `chartex` (waterfall, funnel, histogram/pareto, box &
  whisker, treemap, sunburst; shown only). 3-D charts are drawn flat; region maps show their fallback; pie-of-pie draws
  only the pie. Series read **current cell values**, or the file cache when the reference cannot be followed.
  Untouched chart parts stay byte for byte. Model: `sheet.drawings[i].object` with `type: 'chart'`.
- **Insert and edit** (xlsx): `view.addChart({ type, grouping?, range?, seriesIn?, series?, title?, legend?, anchor? })`
  → id, `updateChart(id, patch)`, `moveChart(id, anchor)`, `removeChart(id)`, `charts()` → `ChartInfo[]`,
  `selectChart(id?)`, `selectedChart`, `openChartDialog(id?)`, `copyChart(id?)`, `cutChart(id?)`, `pasteChart(at?)`,
  `duplicateChart(id?)`, `hasChartCopy`; model `addChart(wb, sheet, spec)`, `updateChart`, `moveChart`, `removeChart`,
  `sheetCharts`. `ChartType`: `'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter'`. `ChartPatch`:
  `type`, `grouping`, `title`, `legend` (`'r' | 'l' | 't' | 'b' | 'tr' | 'none'`), `axisTitles`, `axisBounds`,
  `dataLabels`, `series`, `range`, `seriesIn`. `ChartInfo.editable` is `'all'` (written by the engine: everything),
  `'parts'` (file charts: patched in place, type change refused) or `'none'` (only move, resize, delete). A chart id is
  its index in `sheet.drawings`. Toolbar **Insert chart**, Alt+F1.
- Row/column and sheet operations are allowed when `workbook.chartsFollow` is `true` (all chart references are in the
  model); chart references then shift like formulas. xls charts cannot be changed and are not written by `saveAsXlsx()`.
  No bridge commands. UNSURE: 26 drawing rules were not compared with Excel (listed in the guide).

### Pictures

`view.insertPicture(bytes, { description?, at? })` → id (PNG, JPEG, GIF, BMP — detected from the bytes),
`movePicture`, `removePicture`, `rotatePicture(id, degrees)`, `cropPicture(id, crop | undefined)`,
`startPictureCrop`, `setPictureDescription`, `copyPicture`, `cutPicture`, `pastePicture(at?)` (promise),
`pictures()`; also toolbar **Insert picture**, paste and drag-and-drop. Model: `addPicture(wb, sheet, bytes, mime?,
anchor?)` and friends. Limits: 32 MB per picture (`MAX_PICTURE_BYTES`), 16 384 px per side, 50 M pixels, 256 MB per
workbook. File pictures are moved by changing only their anchor (and `rot`, `srcRect`, `descr` when rotated, cropped,
described). Refused on sheets protected with `objects="1"` and while editing together (`CollabUnsupportedError`,
`what: 'picture'`). No bridge commands.

### Shapes

`view.insertShape(kind, { at?, text? })` with `SHAPE_KINDS` `'rect' | 'roundRect' | 'ellipse' | 'line' | 'arrow' |
'doubleArrow' | 'textBox'`, `moveShape`, `removeShape`, `rotateShape`, `setShapeText`, `editShapeText`,
`formatShape(id, { fill?, lineColor?, lineWidth?, textColor?, textSize?, bold?, italic?, align? })`, `copyShape`,
`cutShape`, `pasteShape`, `shapes()`; toolbar **Insert shape**. Model: `addShape`, `formatShape`, … Shapes the engine
wrote can change in every way; other shapes from the file only move, resize, rotate, delete and get new plain text.
Refused on `objects`-protected sheets and while editing together. No bridge commands.

### Pivot tables

- Pivot cells show as stored. **Refresh**: cell menu, `view.refreshPivot(index)`, `view.refreshAllPivots()`,
  `mounted.refreshPivots()`, bridge `refreshPivots`; re-reads the source range and rewrites the table (labels,
  subtotals, grand totals, 11 functions: `sum`, `count`, `average`, `max`, `min`, `product`, `countNums`, `stdDev`,
  `stdDevp`, `var`, `varp`). `view.pivots()` → `{ index, name, sheet, range, refusal? }`.
- **Create / change fields**: cell menu **Insert pivot table…**, the fields panel, `view.insertPivot({ source, target?,
  rows?, cols?, filters?, values?, name? }, panel = true)`, `view.pivotFields(index)`,
  `view.setPivotFields(index, { rows?, cols?, filters?, values?, rowGrandTotals?, colGrandTotals?, subtotals?, hidden?,
  sort? })`; model `createPivot`, `setPivotFields`. Values: Sum, Count, Average, Max, Min.
- **Delete**: `view.deletePivot(index)` (cells and an unused cache go; one undo step).
- Refusals come with the reason (`PivotRefreshError`, `PivotCreateError`, `PivotFieldsError`, `PivotDeleteError`, all
  `RangeError`): unsupported settings (grouping, calculated fields, show values as, …), caches named by slicers,
  timelines or pivot charts, sources that are not a range or defined name, growth over occupied cells. All pivot
  commands are refused while editing together. Row/column insert/delete works when every pivot part is in the model
  (not through a table or its source header). UNSURE: grand-total places, sort order, "(blank)" and subtotal labels
  were not compared with Excel.


## 8. Low level

### Spreadsheet model: `file.workbook()`

Read on first call, then the same object. Reading does not change the file. Rows and columns are **0-based**
(`A1` → row 0, column 0).

```js
const file = await BymOffice.open(bytes);
const workbook = await file.workbook();
const sheet = workbook.sheets[workbook.activeSheet];
const cell = BymOffice.getCell(sheet, 0, 0);       // A1, or undefined
cell.kind;     // 'blank' | 'number' | 'string' | 'boolean' | 'error' | 'date'
cell.value;    // number | string | boolean | null
cell.formula;  // { text: 'SUM(B2:B3)', type: 'normal' } or undefined
const style = BymOffice.resolveCellStyle(workbook.styles, cell.style);
BymOffice.resolveColor(style.font.color, workbook.styles); // '#RRGGBB' or undefined
BymOffice.cellDisplayText(workbook, cell, BymOffice.getLocale('en')); // text as shown
```

- `Workbook`: `sheets`, `styles`, `sharedStrings`, `definedNames` (`name`, `value`, `localSheet`, `hidden`),
  `date1904`, `activeSheet` (may point at a hidden sheet: check `state`).
- `Sheet`: `name`, `sheetId`, `state` (`'visible' | 'hidden' | 'veryHidden'`), `kind` (`'worksheet' | 'chartsheet' | …`),
  `part`, `rows: Map<number, Row>` (sparse; `Row.cells` sparse by column, `height` in points, `hidden`, `style`,
  `outlineLevel`), `columns` (`min`, `max`, `width` in characters, `hidden`, `style`), `merges`, `hyperlinks`,
  `rowCount`/`colCount` (last used + 1), `dimension`, `defaultRowHeight`, `defaultColWidth`, `view` (`frozenRows`,
  `frozenCols`, `topLeftCell`, `firstVisibleCell`, `split`, `showGridLines`, `zoomScale`, `rightToLeft`, `activeCell`…),
  `tabColor`, `pageSetup` (`SheetPageSetup`: `margins`, `page`, `print`, `headerFooter`, `fitToPage`, `autoPageBreaks`,
  `rowBreaks`, `colBreaks`; a read view — xlsx `save()` writes these elements from the original bytes, change them
  with `setPageSetup`), `drawings`, `comments`, `warnings`.
- Dates are **numbers**; `style.numFmt` tells that they are dates (`isDateFormat(code)`, `serialToDate(serial, date1904)`).
- Formulas are not recalculated on open: `cell.value` is the last saved result.
- Helpers: `getCell`, `columnName`, `cellName`, `parseCellRef('B7')`, `parseRange('A1:C3')`, `rangeName(range)`,
  `resolveCellStyle`, `resolveColor`, `formatCodeOf(styles, numFmtId)`.
- Model operations (`applyOp`, `History`, `setCellInput`, `copyRange`/`cutRange`/`pasteAt`, `formatRange`, `insertSheet`,
  `addChart`, `refreshPivot`, …) are exported for headless editing; `CalcEngine` calculates formulas. Rendering
  helpers (`cellText`, `createSheetMetrics`, `createLayout`, `cellRect`, `hitTest`, `renderGrid`) may change before
  the first stable version.
- A broken sheet rejects with an `Error` whose message starts with `Invalid XML …`, `Invalid xlsx …` or
  `Invalid worksheet …`. Numbers such as `0x1F`, `NaN`, `INF` make a file count as damaged.

**Number formats:** `formatValue(value, code, { locale?, date1904?, builtinId? })` → `{ text, color?, overflow? }`,
`formatCell(cell, code, options)`, `isoToSerial(text, date1904?)`, `cellDisplayText(workbook, cell, locale?)`,
`isDateFormat(code)`, `serialToDate(serial, date1904?)`, `getLocale(tag)`, `LOCALE_TR`, `LOCALE_EN`,
`normalizeLocale(partial)`. Supported: four sections, conditions, colours, `0 # ?`, thousands and scaling commas,
percent, scientific, fractions, literal text, currency and locale codes (`[$₺-41F]`, `[$-409]`), date/time, elapsed
time (`[h]:mm:ss`), AM/PM, 1900 and 1904 date systems. Rounding is half-up on the decimal representation
(`1.005` → `1.01`). A broken format code does not throw: the value is shown as General. Checked against an
independent implementation, not against Excel.

### Document model: `file.document()`

docx, doc and RTF are read into the same `DocModel`. Edit a docx model with the `BymOffice.docOps` operations and
write it with `file.save()`; doc and RTF models are written only as docx (`saveAsDocx()`). Pictures of doc/RTF are in
`model.media`, not in a package: `BymOffice.documentPictures(model, file.package)` gives both through one interface.

### `OpcPackage` (docx/xlsx: `file.package`)

Parts are named with absolute names (`/xl/workbook.xml`, `/word/document.xml`).

| Member | Meaning |
|---|---|
| `partNames(): string[]` | All part names, in file order |
| `hasPart(part): boolean` | Part exists (case-insensitive) |
| `contentTypeOf(part): string \| undefined` | Content type |
| `mainPart(): Promise<string \| undefined>` | Main part |
| `relationships(part = '/'): Promise<Relationship[]>` | Relationships of a part |
| `readPart(part): Promise<Uint8Array>` | Unpacked bytes |
| `readXml(part): Promise<XmlDocument>` | Part as an XML tree |
| `writePart(part, data, contentType?)` | Add or replace a part (`Uint8Array` or `XmlDocument`) |
| `save(): Promise<Uint8Array>` | Write the package; untouched parts stay as they were |

XML: `parseXml(input)` and `serializeXml(doc)` keep prefixes, attribute order, comments, CDATA and white space.
Documents with a DTD are rejected. Helpers: `childElements(element, localName?)`, `localName(name)`.

### `CfbFile` (doc/xls: `file.cfb`) and `writeCfb`

`entries` (`path`, `type`, `size`), `has(path)`, `stream(path)` (e.g. `'Workbook'`, `'WordDocument'`), `head(path, size)`
(only the first `size` bytes). `writeCfb(streams: { path, data }[]): Uint8Array` writes a compound file (version 3,
512-byte sectors, below 4 GB) and `allocateCfb(streams: { path, size }[])` lays it out first. This writer exists for
the container of password-protected docx/xlsx: **it does not write doc or xls**. UNSURE: its output was not opened in
Office directly.

### `ZipArchive`

`ZipArchive.read(bytes)`, `names()`, `has(name)`, `entry(name)`, `read(name)`, `set(name, data)`, `delete(name)`, `write()`.

- ZIP64 archives are read; more than 65 534 entries are written with ZIP64 records. 4 GB archives and larger cannot be written.
- Entry comments, extra fields, attributes, the archive comment and non-UTF-8 names are written back as they were.
- An entry whose content did not change keeps its local header and its data descriptor byte for byte, and its central
  header flags and "version needed". Entries with ZIP64 fields, and entries whose two headers disagree about the
  descriptor, get new (valid, equivalent) headers. A changed entry is written without a descriptor, version 2.0.
- An archive with two entries of the same name is rejected; an entry that unpacks larger than it declares (a
  "zip bomb") is rejected.
- **Unpacked size limits** (fixed in this version) — `read(name)` rejects with `PartTooLargeError`
  (`kind: 'part' | 'archive'`, `part`, `size`, `limit`) before unpacking anything:
  - one part at most `MAX_UNPACKED_PART_BYTES` = 512 MB (sheet bodies, shared strings, pivot cache records, Word
    `document` XML, every non-XML file) or `MAX_UNPACKED_XML_BYTES` = 64 MB (every other `.xml` / `.rels` / `.vml`);
  - all parts read from one archive together at most `MAX_UNPACK_RATIO` = 256 times the archive size, but at least
    `MIN_UNPACK_BUDGET_BYTES` = 64 MB.
  - Entries your code adds with `set()` are not limited.

## 9. WebView bridge

Mobile and desktop apps show the editor in a WebView with one file: `dist/bymoffice.webview.html`. It carries the
engine and the bundled font, needs no network, and can also be used in an `<iframe>`. App and page exchange JSON
strings: the app calls `window.BymOfficeBridge.receive(json)`; the page sends through the first channel that exists
(`window.BymOfficeHost.postMessage`, `window.webkit.messageHandlers.bymoffice`, `flutter_inappwebview.callHandler`,
`window.chrome.webview.postMessage`, `window.ReactNativeWebView.postMessage`, or `window.parent.postMessage` only to
allow-listed origins). `id` on a request comes back on its answer. `BYM_NO_FONTS=1 pnpm build` builds the page
without the font.

| Direction | Messages |
|---|---|
| App → page | `load` (`data` base64, `url` or `blank`; plus `name`, `locale`, `readOnly`, `sheet`, `convert`, `toolbar`, `save`, `author`, `trackChanges`, `commentsPanel`, `fonts`, `password`, `passwordPrompt`, `collab`), `save`, `setReadOnly`, `setLocale`, `undo`, `redo`, `print`, `exportPdf`, `printPreview`, `insertSheet`, `renameSheet`, `deleteSheet`, `moveSheet`, `setSheetState`, `copySheet`, `setTabColor`, `refreshPivots`, `insertPivot`, `setPivotFields`, `deletePivot`, `pivotFields`, `setTabStops`, `setPageSetup`, `insertPageBreak`, `removePageBreak`, `resetPageBreaks`, `insertSectionBreak`, `setDefaultTabStop`, `acceptAllRevisions`, `rejectAllRevisions`, `setTrackChanges`, `addComment`, `editComment`, `replyComment`, `resolveComment`, `removeComment`, `commentsPanel`, `close`, `saveResult`, `password`, `hello` |
| Page → app | `ready` (`version`), `loaded`, `passwordRequired`, `changed`, `saveState`, `saved` (also `by: "user"` / `by: "auto"` with a `saveId` the app answers with `saveResult`), `pdf`, `selection`, `revisions`, `comments`, `sheet`, `sheets`, `done`, `link`, `error`, `collabState`, `collabSaved` |

Error codes include `bad-message`, `not-loaded`, `read-only`, `no-document`, `no-sheet`, `unsupported`, `corrupt`,
`load-failed`, `save-failed`, `password-cancelled`, `pdf-too-large` (above 64 MB), `pdf-failed`, `print-failed`,
`locale-refused`, `pivot-refused`, `sheet-refused`, `protected`, `page-setup-refused`, `comment-refused`.
There are no bridge messages for find and replace, pictures, shapes, charts or notes.

The full contract (fields, ordering, queueing, security, platform snippets for Flutter, Android, iOS/macOS,
Windows WebView2, React Native and `<iframe>`) is in Turkish: [webview-koprusu.md](webview-koprusu.md). The page and
contract are tested in Chromium, Firefox, WebKit and mobile profiles; **not on a real device's WebView**, and the
platform snippets were not compiled.

## 10. Browser and platform support

- Current Chrome, Firefox and Safari; iOS and Android browsers and WebViews. Every change is tested with Playwright in
  Chromium, Firefox, WebKit and two mobile profiles (iPhone, Pixel).
- Without a built-in `DecompressionStream` (or without `deflate-raw`) the engine uses its own inflate/deflate code
  (iOS before 16.4, old Android WebViews); nothing to configure.
- The package is compiled down to ES2019 syntax; neither output uses `import.meta`.
- Node and Bun: everything without a view (`open`, models, operations, `save`, `createBlank`, `buildXlsx`/`buildDocx`).
- Limit: not tried on real old devices (old environments are imitated in tests by removing built-in APIs); screen
  readers were not tried.

## 11. Not yet available

This is a translation of the guide's "Henüz olmayanlar" table. Do not build on anything listed here.

| Feature | Status |
|---|---|
| Freeze and merge: freezing from a scrolled position (from the first visible row, as Excel does); creating and showing a split (not frozen) pane; "Merge across" (row by row); merging on a sheet that holds a table (ListObject) | Not available. Freezing counts from A1; a split pane is kept in the file; merging on a sheet with a table is refused |
| Tab stops: drawing the vertical line of a bar (`bar`) stop; placing stops by dragging on a ruler | Not available. A bar stop is written and kept but not drawn; stops are entered in the dialog and through the API |
| Documents edited together (`collab`, docx): a document with tracked changes; inserting / resizing pictures, accepting / rejecting tracked changes, rich paste (copies carrying pictures, fields, footnotes); undoing a deleted comment, redoing an undone comment insertion | The room is not set up (document with tracked changes) or the operation is not done; both are reported |
| Spreadsheets edited together (`collab`, xlsx): copying a sheet; moving rules with a cell paste; **pasting a chart that was cut or changed after it was copied** (cutting a chart only copies it; a copy of an unchanged chart and duplicating work); undoing a sheet deletion, rule and chart commands; getting notes, links and rules back when undoing a row/column deletion; persistence in the relay; a relay on another origin from the bridge page (its content security policy says `connect-src 'self'`) | Not available. What cannot be done is not done and is reported; for a file that cannot be edited together (not an xlsx/docx, converted, or a document with tracked changes) a notice comes and the file is edited alone |
| Formulas: structured references (`Table[Column]`), references to other files and 3-D references (`Sheet1:Sheet3!A1`), `;` as argument separator in the English view; Turkish names of the `LAMBDA` family (`LET`, `LAMBDA`, `MAP`… are written in English) | Not available. A cell with such a formula keeps its last value from the file, or gives `#REF!`/`#NAME?`; the formula text is not damaged on save |
| Dynamic arrays: showing an ordinary formula from the file with `@` (Excel puts `@` where it intersects; here the text is shown as in the file); inserting/deleting rows or columns through a spill range and sorting that cuts a spill range; a spill into a table (ListObject) giving `#SPILL!` (tables are not in the model); the "ghost" look of spilled cells beyond font/alignment | Refused (insert/delete, sort) or not done |
| Locale output and Turkish format codes in formulas: `TEXT`, `FIXED`, `DOLLAR` output and number-to-text conversion (`&`, `CONCAT`) with local separators (`1.234,50`) and Turkish month/day names; the code Turkish Excel shows inside `TEXT` (`#.##0,00`, `gg.aa.yyyy`) | Not available. In the Turkish view too the output is in file format (`=METNEÇEVİR(1234,5;"#,##0.00")` → `1,234.50`, `=1,5&""` → `1.5`, month name `March`; for Turkish names use `[$-41F]` in the code); format codes are written in English. Reading is local: `"7,5"` is read as a number. A result saved by Turkish Excel is kept as long as its inputs do not change |
| Inserting/deleting rows or columns on sheets with tables or form controls; in workbooks with **pivot parts the model cannot hold** (an unreadable table or one beyond the limit of 1,000, a table/cache part no sheet shows, a cache whose source is not a single range or defined name: external data, OLAP, consolidation); and in workbooks with **a chart whose references cannot be followed** (chart sheet, pivot chart, `chartex`, an unread type, a chart with extra formulas in an extension or data label, one chart part shown by two frames) | Refused (the references of this content are not shifted; refused so that nothing breaks silently); editing and formatting cells is free. Free on sheets with pictures and comments, in workbooks where **all charts' references are in the model** (`workbook.chartsFollow`) and where **all pivot tables are in the model** (an operation through a pivot table, or one that takes the source's header row/column, is still refused) |
| Pointing at cells while typing a formula: colouring all references of the formula (only the last added one is framed); moving/resizing a reference by dragging its frame; pointing at a range on touch (one tap adds a cell, row or column; dragging scrolls); the whole sheet with the corner button; several areas with Ctrl+click; suggesting named ranges or table names; pointing at cells and changing sheets from the formula bar with the keyboard; the in-cell editor being visible while another sheet is shown | Not available. A pointed reference is always relative; add `$` with F4 or by hand. Excel's macOS equivalent of F4, Cmd+T, does **not** exist (it is the browser's new-tab shortcut) |
| Find and replace: regular expressions and wildcards (`*`, `?`, special marks like `^p`), search by format, limiting the search to the selection (in the bar; the API has `range`), a "Find all" list, searching across paragraphs in documents, replacing in comment text while editing together, replacing in text boxes while tracking changes, "Replace all" in text boxes while editing together, replacing a field result; searching column by column | Not available. The searched text is plain text |
| Pointing at cells from rule dialogs: collapsing the dialog to one line or moving it (it only becomes transparent while dragging; a drag cannot start from a cell the dialog covers); scrolling the grid while the dialog is open (it scrolls by itself at the edge while dragging) and pointing with arrow keys; F4 to toggle `$` in a field; a range on touch (one tap points at one cell, row or column; on a phone the dialog covers most of the grid); an "applies to" field in the data validation dialog (the rule applies to the selection the dialog was opened with); colouring all pointed references (only the last one is framed) | Not available. Ranges and sources can also be typed: `=$A$1:$A$5`, `A1:B10; D1` |
| Moving rules with copy/cut and paste: moving rules that are kept as they are in the file | Not available. A rule that stands only in the 2010 extension list (an x14 validation whose source is on another sheet, a conditional rule only in the extension) and a rule whose formula cannot be read to the end do not go into the copy, and stay where they are, as they are, when pasted over or cut (for conditional formats: all rules of the `<conditionalFormatting>` element that holds such a rule). A rule whose extension carries extra information (borders/axis/negative colours of an Excel 2010 data bar, individually chosen icons) goes to a place without it as a new rule in its main-section form. All of this is reported after the paste (`PasteResult.skippedRules`). Text/HTML from outside carries no rules and does not touch the target's rules. The menu's "Paste" pastes plain text when the browser does not give the clipboard's HTML (`navigator.clipboard.read` missing): **cut + paste from the menu does not move rules and notes in that browser** (undo brings the cut source's rules back) |
| Data validation: changing/deleting a rule in the 2010 extension list (`<x14:dataValidation>`; old files whose list is on another sheet); "apply to all other cells with the same settings"; a confirmation dialog for the warning/information type | Not available. A rule in the extension list is shown and applied and kept byte for byte; `setValidation` / the dialog covering those cells is **refused** (so no double rule appears) |
| Conditional formats: **creating** icon sets specific to the 2010 extension (`3Stars`, `3Triangles`, `5Boxes`) and individually chosen icons (those in the file are kept); choosing threshold type and value in icon sets (a new rule gets equal percentages; an existing rule's thresholds are kept); time periods beyond the ten in the list; changing the **content** (colour, threshold, format) of a rule that has a counterpart in the 2010 extension (Excel 2010+ data bars: `extId`) — its range, order and "stop if true" can change and it **can be deleted**; any change to a rule that stands only in the extension list (x14 icon sets etc.) and reordering on a sheet with such a rule; editing border, number format and gradient fill in the format (existing ones are kept); choosing threshold type and value in data bars/colour scales (a new rule gets minimum–maximum, a three-colour midpoint at the 50th percentile; an **existing** rule's thresholds, bar lengths, `showValue` and `stdDev` are kept as they are); ready items for icon sets and time periods in the toolbar menu (choose them from the type list of the rule dialog) | Not available / refused (with a clear message). The rest of the extension list is always kept byte for byte |
| Copying a sheet: the part of conditional formats that stands only in the 2010 extension list (extension-only rules, data bar border/axis/negative colours: `conditionalFormatExtensions`), pictures/shapes/charts, tables, pivot tables, form controls, comment threads, external links, sparklines, custom views, ignored-error rules, watched cells, sort state, phonetic settings, outline settings and levels, the sheet's code name, header/footer pictures (`legacyDrawingHF`), the sheet background picture (`picture`), row attributes **from another namespace** (Excel's `x14ac:dyDescent` hint on every row is not counted; attributes without a namespace — `thickBot`, `thickTop`, `ph` — are written to the copy); copying into another workbook | **Not taken** into the copy (content kept as in the file cannot be written for a sheet that is not in the file). The copy is made, and **each** item left out is reported with its kind and count through `onNotice` (`copySheetLosses`). Sheet protection and protected ranges **are** copied (ranges shifted for rows/columns inserted/deleted so far). Moving tabs has no keyboard shortcut (only the tab menu: browsers keep Ctrl+Shift+PageUp/PageDown); a chart sheet's tab colour cannot be changed; sheet protection **is enforced** and set/removed from the tab menu |
| In a workbook with a pivot table, a form control or **a chart whose references cannot be followed** (row above): deleting and renaming sheets; deleting a sheet that holds a table; deleting and renaming a chart sheet; in a workbook with two sheets whose names differ only in case (`IL` / `İL`): every sheet operation (inserting too) | Refused (this content names sheets and is not rewritten yet); the menu item looks disabled. **Inserting** a sheet is otherwise always free |
| Updating the model for "place in document" links of drawings (shapes, pictures) | Updated in the file, not in the model: `#Sheet!A1` targets in drawing relationships and the sheet's cell links written as `Target="#…"` get the new name/place **in the saved file** on rename, delete and row/column insert/delete (a target to a deleted sheet becomes `#REF!`, the relationship stays — UNSURE: what Excel writes there was not checked; undo returns the part to its original bytes). The open model's `hyperlink.target` keeps the old text until the file is opened again |
| 3-D references (`Sheet2:Sheet4!A1`) in content kept "as in the file", such as conditional formats/validation | UNSURE: if an end sheet is deleted and then rows/columns are inserted next to it, the reference in this content can differ from the same reference in a cell (cells are adjusted operation by operation, this content in one pass from the file state). Excel does not allow 3-D references in these rules; seen only in hand-made files |
| Filtering by colour and icon, filters inside tables, sorting on sheets with comments or tables, a confirmation dialog for warning-type validation | Not available. Filter criteria that cannot be applied are kept in the file |
| Conditional formats: Excel's original icon pictures (simple vector icons are drawn), the font name/size of the rule and pattern/gradient fills | Not available. With `calculate: false` formula rules are not applied |
| Word 95 and earlier | Not available; `UnsupportedFileError` |
| RTF: math, frame positioning, line and fill of Word 95 drawing objects (`\do`; for text boxes only the text is read, others are skipped), old `\pn` lists, table styles, Far East code pages; writing RTF | Not available. RTF opens read-only, `save()` gives the original bytes |
| Editing and saving doc; EMF/WMF pictures, picture bullets, old automatic numbers and comment ranges in doc | Not available. doc opens read-only, `save()` gives the original bytes |
| Page setup: editing headers/footers, page order, first page number, printing of error cells and comments, custom paper size; page break preview; in documents columns, page borders, vertical alignment, line numbers, gutter (in the dialog), mirror margins; changing section properties while tracking (`w:sectPrChange`); page setup while editing together | Not available. These are read from the file, applied when printing (those that are applied) and kept. Paper, orientation, margins, scale/fit, centring, grid lines/headings, print area, print titles, manual page breaks (sheets) and section paper/orientation/margins, section breaks (documents) can be changed |
| Printing sheets: splitting a row/column that does not fit on one page (it gets its own page, what overflows is cut at the margin); printing comments (`cellComments`: at the end or as displayed); printing error cells as blank/dash/`#N/A` (`errors`); black-and-white and draft modes; header/footer pictures (`&G`), theme colours (`&K04+000`), outline and shadow (`&O`, `&H`), file path (`&Z` stays empty); copies and printer resolution settings; printing only the selected range; right-to-left sheets; a PDF with selectable text from printing | Not available. These settings are kept in the file and do not affect the output |
| PDF (`exportPdf`), vector pages: embedding **device fonts** (a page written in Arial, Times New Roman, Calibri… stays an image; Chromium's permission-based `queryLocalFonts` is not used); embedding from **WOFF / WOFF2** and font collections (`.ttc`); CID-keyed CFF fonts; **variable** (`fvar`) and **colour** (`COLR`, `SVG`, `sbix`, `CBDT`) fonts; **text shaping** (from `GSUB` only Latin ligatures are read, `GPOS` is not read — glyph positions come from the browser's measurement; a font with contextual alternates (`calt`) or required features, rules of non-Latin scripts, `locl`/`ccmp`; a page with combining marks, right-to-left or joining/overlapping letters — Arabic, Hebrew, Devanagari, Thai… — stays an image); in browsers that cannot be asked which ligatures they drew (no `fontKerning` on canvas: **Safari and iOS WebView**) a page with letter sequences the font can join (`fi`, `fl`, `ff`…) — so most pages of documents in the bundled font stay images in Safari; a page with characters missing from the font (falling back: emoji, CJK, `₺`); **gradient** fills (gradient data bars); a sheet page with row/column headings in a browser that cannot measure `textBaseline`; **subsetting** OpenType-CFF fonts (embedded whole); embedding original **JPEG bytes** as they are (pictures are re-encoded from decoded pixels); links: `tel:`, `ftp:` and relative addresses, keeping page order once the 50,000-link file limit is passed (external links are counted as they come, internal ones at the end of the file; in a document over the limit, table-of-contents links drop first), opening of host names written with letters of new Unicode versions in every reader (the engine accepts them; a reader with old IDNA tables may not open them; it goes to no other server), `HYPERLINK()` formulas in cells, links on text boxes/shapes/charts/pictures, link order across columns in multi-column sections; outline: reading the document title and author from the file (`docProps/core.xml`), collapsed levels, headings inside text boxes, cell range/defined name entries for sheets; **PDF/A**, **tagged PDF / accessibility**, **form fields**, **encryption / password**, choosing a page range, text inside charts entering the text layer of image pages, the full Unicode bidirectional algorithm (order of runs and cells in a right-to-left line), reading multi-column sections column by column | Not available. A page that cannot be vector is written as an image + invisible text layer (search, select and copy work; it pixelates when zoomed) |
| Printing documents (`print`): selectable/searchable output (pages go to the print dialog as images; use `exportPdf()` for a searchable PDF), page range/copies options (chosen in the browser's dialog), printing by itself in WebViews (in a WebView without `window.print()`, or where it does nothing, the app wires its own printing; see [webview-koprusu.md](webview-koprusu.md)) | Not available |
| Sheet protection: `selectLockedCells` / `selectUnlockedCells` (selection goes to every cell), asking for the password of a protected range with a password (`<protectedRange password>`; such a range counts as closed), user-permission ranges (`securityDescriptor`), `scenarios` and `insertHyperlinks` (neither has a command), a command to protect/unprotect the workbook structure (`workbookProtection` is read and enforced, not changed), setting/removing protection while editing together | Not available. The protection stays byte for byte in the file |
| Editing xls and saving as xls; in xls conditional formats, data validation, auto-filter criteria, printer driver data (PLS), page setup of custom views, header/footer pictures, outline settings, form controls, macros, external workbook paths; Excel 4 and earlier | Not available. xls opens read-only, `save()` gives the original bytes. As these are not read into the model, the xlsx made by `saveAsXlsx()` does not have them either |
| Document editing: carrying footnote/endnote references, comments, equations, content controls, tracked changes, text boxes and shapes with text, charts, diagrams, embedded objects (OLE), VML pictures, `mc:AlternateContent`, fields spanning several paragraphs (table of contents), smart tags and custom XML through the clipboard (a copy takes only their text; a selection holding them **cannot be cut** and stays in place with a notice); carrying rich content to an editor in another tab/page or to another application (the clipboard gives others only text and simple HTML); pasting HTML from outside with its formatting (only its text); pasting over table cells cell by cell (a table is inserted as a nested table); tracked changes and comment editing through the clipboard | Not available. Pictures, links, fields, page/column breaks, tables, paragraph formatting, styles and lists are carried between editors on the same page |
| Typing in text boxes: editing a box while changes are tracked; "Replace all" in boxes while editing together (single Replace works); putting pictures, links, page/section breaks, comments and **new** lists (that add a numbering definition) into a box; rich paste into a box (copies carrying pictures, fields, links); entering a box by clicking its empty area away from its text lines; selecting, moving and resizing a box; growing the box or clipping the text when it does not fit; table buttons other than Tab between cells in a table inside a box were not tried | Not available. These are **refused** with the reason (nothing changes); only text and its formatting are pasted into a box |
| Tracked changes: accepting/rejecting the kinds the model keeps only as XML (**merged cells** `w:cellMerge`, `w:cellIns`/`w:cellDel` marks that do not form a whole column (only in some rows, with a merged cell, two authors or dates), `w:tblGridChange` carrying more than plain columns, list number changes `w:numberingChange`, changes inside text boxes, `mc:AlternateContent`, comments and unknown elements — listed with kind `other`; "all" skips and counts them); accept/reject that would take only part of a field (refused); a right-click / long-press menu on the page; keyboard shortcuts for accept/reject and navigation; a sidebar and balloons; a mouse tooltip over deleted text; a "final only" view; marking inserted/deleted rows and cells themselves in tables (their text shows underlined/struck through, the row or cell is not painted); **while tracking**: inserting/deleting a column that cuts a merged cell or touches columns a row skips, deleting a column only some of whose cells are deleted, a link over a selection holding a picture/footnote reference/equation/field (simple fields too)/content control, blocks other than paragraphs and tables (`insertBlock`), changing section properties (`w:sectPrChange`) | Not available. In a document that tracks changes these operations are **refused** (with a notice; the document does not change), because doing them untracked would go against the document's setting. The editor has **no** command to merge/split cells, change table/row/cell properties, resize/move pictures or remove links (with or without tracking), so it never produces `w:cellMerge`, `w:tblPrChange`, `w:trPrChange`, `w:tcPrChange`; such records from the file are resolved, `w:cellMerge` stays `other`. Typing, deleting, Enter, page breaks, cutting, **character and paragraph formatting, styles, lists, tables, rows, columns, pictures, links, blocks and rich paste** are tracked |
| Document comments: **editing** formatting in comment text (bold, links, tables…) — rich content from the file is kept byte for byte until edited, an edited comment becomes plain paragraphs (`DocComment.plain` tells beforehand); **creating** `commentsIds.xml` / `commentsExtensible.xml` (updated if present) and `people.xml` (author identity, @mentions); adding comments to headers, footers and text boxes (refused; Word does not take comments in headers either); answering/deleting a comment whose marks are in a text box or in XML kept as is (refused; it can be edited and resolved); a mark on the page for a point comment in an empty paragraph (reachable only with Previous/Next); **printing** comments (highlight, mark and balloon are not printed); filtering comments by author; comment **ranges** in doc/RTF (the comment is tied to the point of its reference) and editing comments of doc/RTF without converting | Not available |
| Document protection: checking passwords, removing or adding protection; filling in form fields (`forms` protection opens read-only; under `comments` protection comments are free); editing ranges exempt from protection (`w:permStart`/`w:permEnd`); enforcing `w:writeProtection` (read-only recommendation) and the formatting restriction (`w:formatting`; protection without `w:edit` locks nothing); reading read-only protection in doc | Not available. The protection element is kept byte for byte; exempt ranges are reported by `protection.exceptions`, and the document is still entirely read-only |
| Fonts: a bundled font for families other than Bookman Old Style (only TeX Gyre Bonum ships; other families use the device's and those you pass to `registerFont`); loading fonts embedded in documents (docx `word/fonts/*.odttf`); metric-compatible fallbacks for Tahoma/Verdana/Garamond/Palatino Linotype; Word's line height for Bookman Old Style and other families missing from the table; font registration in browsers without `FontFace` | Not available |
| Documents: in-document links, search (see note), right-to-left text, hyphenation, drawing vector pictures (EMF/WMF/SVG) and diagrams (SmartArt), fonts embedded in the document | Not available. A placeholder appears where an object cannot be drawn. *Note: the guide lists "search" here, but Find and replace does exist in documents (see [Documents](#find-and-replace-in-documents)).* |
| Charts in embedded objects of doc/RTF: Excel 2007+ objects (`Excel.Chart.12` / `Excel.Sheet.12`: an xlsx package inside), workbooks older than BIFF5, linked (not embedded) objects, **floating** OLE objects in doc, the cells of an Excel sheet object; writing the chart to docx **as a chart** (its picture from the file is written, with a warning) | Not available. The object shows as before (in RTF its result picture, in doc nothing); why a chart object could not be read is in `doc.warnings` |
| xls charts (display): surface and pie-of-bar / pie-of-pie (placeholder); **trendlines and error bars** (not counted as series, not drawn; a note is added); series axis; data table; gradient, pattern and picture fills (patterns become solid colour); manually placed title / legend / plot area (legend stays on the right); axis label angles other than quarter turns; Excel 2007+ extra format records (`CrtLayout12`, `ShapePropsStream`, `TextPropsStream`…); drawing objects over a chart sheet; Excel 5/95 (BIFF5) charts (chart sheet empty, no embedded chart); **writing** the chart to xlsx with `saveAsXlsx()` and editing it with `convert: true` | Not available. A chart that cannot be read stays a placeholder and is listed in the sheet's `warnings`; saving to xlsx does not write the chart and counts it in `report.warnings` |
| Charts in documents (docx): refreshing a series that **has** a cache from the embedded workbook (the cache is drawn; the workbook is opened only for references without a cache); binary embedded workbooks (`.xlsb`); password-protected embedded workbooks; linked workbooks (`TargetMode="External"`); reading multi-level categories from the workbook; embedded workbooks larger than 16 MB; **inserting, editing, resizing** charts; carrying a chart through the clipboard (a selection with a chart cannot be cut; a copy does not take the chart); rotation and effects (shadow etc.) of the chart frame; `c:clrMapOvr`. Chart titles and labels do not count as document text (not searched, not copied) | Not available |
| Shapes in documents (docx, doc, RTF): geometries other than rectangle, rounded rectangle, ellipse and straight line (arrows, freeforms, flowchart shapes…); arrow heads; double/triple (compound) lines; custom dashes; gradient, pattern and picture fills; transparency; shadows and other effects; parts of shape groups and drawing canvases; rotating a text box together with its text | Not available. Drawn **in a single-shape drawing** (not a group): preset geometry `rect` (snipped/one-round-corner rectangle kinds and `flowChartProcess` as a plain rectangle too), `roundRect`, `ellipse`, `line`, `straightConnector1`; solid fill; line colour, width (0 = thinnest) and preset dash; rotation of shapes without text, flipping of lines. **Approximated:** arrow lines without heads, compound lines as single, custom dashes as solid; no shadow; partly transparent colours opaque (fully transparent not painted); a *line* with gradient/pattern not drawn at all; shapes with text not rotated. **Placeholders:** a shape without text in another geometry (or in docx with custom geometry or none), or with gradient/pattern/picture fill; groups without text, canvases, charts, diagrams. In those cases a **text box** shows its text: for an unknown geometry without a frame in docx, with a rectangle frame in doc/RTF; with an unshowable fill, without fill. In **groups** only the parts' texts (one below another) and pictures are shown, not their geometry. In doc/RTF also: palette/scheme/system-indexed colours are not resolved (fill not painted, line black), the drawing group's default property table is not read (a shape takes the spec default for what it does not state), arrow heads, compound lines and shadows are not read; a picture shape whose picture is missing is a placeholder. In docx all of this is kept in the file |
| Grid view: rotated text, shrink to fit, justified alignment, right-to-left sheets, zoom, diagonal borders, pattern/gradient fills, hiding headings | Not available; content is drawn plainly or skipped |
| Number formats: languages other than Turkish and English, `*` fill, the real width of `_` spacing, Hijri/Japanese calendars, `[DBNum]` | Not available; `*` writes nothing, `_x` is one space |
| Model | doc and RTF are read with `document()` into the same document model as docx, xls with `workbook()` into the same spreadsheet model. A docx model is edited with `BymOffice.docOps` and written with `file.save()`; a doc/RTF model only as docx with `saveAsDocx()`, an xls model only as xlsx with `saveAsXlsx()` |
| Writing doc/xls | Not available; `save()` returns the original bytes (`saveAsDocx()`/`saveAsXlsx()` convert to the new format) |
| Password-protected files: saving encrypted **as** doc/xls (encrypted doc/xls open read-only; converted to docx/xlsx they are encrypted with the same password); ECMA-376 **Extensible** encryption (4.3; third-party module), certificate encryption, CFB chaining (`ChainingModeCFB`) and non-AES ciphers (DES, 3DES, RC2) in Agile, RC4 in Standard; summary information in the "encrypted stream" of RC4 CryptoAPI (document properties are not read); a "show password" button in the password dialog; setting/changing/removing a password from the UI (API only: `setPassword`, `save({ password })`); checking read-only recommendation / modify passwords (`w:writeProtection`, xlsx `fileSharing`) and sheet/workbook/document **protection** passwords (these are not encryption; kept byte for byte) | Not available. Unreadable encryption is refused with `PasswordError` (`code: 'unsupported'`) |
| Pivot tables: deleting a table from the file in a workbook with **pivot parts the model cannot hold** (unreadable or beyond the 1,000 limit, a table/cache part no sheet shows, a table shown by two sheets); deleting a table whose cache is named by a slicer, timeline, pivot chart or other part (refused with the reason: `PivotDeleteError`); deleting while editing together; creating from an **Excel table** (`ListObject`); sharing an existing cache when creating; showing values **on rows**; a table without row fields or values (empty); filters in several columns (`colPageCount` > 1); **changing fields** of a table with formats on its parts (`formats`/`conditionalFormats`) or an extension naming its fields; drawing the pivot table style (`PivotStyleLight16`…; cells show their own formats); grouping (date/number ranges, manual groups), calculated fields and items, label/value filters (`filters`), show values as (`showDataAs`: ratio, difference, running total…), values on rows (`dataOnRows`), a column field **together with** several values, outline/tabular layout with column fields, a table without row fields, mixed layout, extra subtotal functions per field, sorting by values (`autoSortScope`), item-bound formats (item references inside `formats`/`conditionalFormats`), a table with a pivot chart, repeating item labels, blank rows, classic layout, a cache fed by an Excel table name (`ListObject`), another workbook, OLAP or external data; refreshing while editing together | Refresh is refused with the reason (`PivotRefreshError`); the table and its parts stay as in the file |
| Pictures in Excel: a mirror (flip) command, rotating with the keyboard, moving the picture inside the crop in crop mode, negative crop (showing the picture larger than its box), writing JPEG/GIF/BMP to the system clipboard as an image (only PNG), duplicating a picture (`Ctrl+D`), sharing a pasted picture with a media part of the same bytes in the file, replacing with another picture, moving a picture inside a group on its own, inserting WebP/SVG/EMF/WMF/TIFF, "Insert picture" in the cell menu, bridge commands for pictures, pictures while editing together | Not available; pictures in the file are kept. While editing together a picture operation is refused (`CollabUnsupportedError`, `what: 'picture'`) |
| Shapes in Excel: drawing by dragging on the grid (a shape is placed at the active cell at the default size); inserting geometries other than rectangle, rounded rectangle, oval, line, arrow and text box (those in the file are drawn); flipping a line by dragging one end past the other and a mirror command; gradient/pattern fills, shadows, dash choice, arrow head choice; theme colours; run-by-run text format, vertical alignment and inset commands, bold/italic from the toolbar (only `formatShape`); changing the fill/line/text look of a shape written by Excel (unknown) and copying it; changing the text of a shape whose text has several looks, fields or a cell link (`textlink`); grouping, moving a shape inside a group on its own, attaching connectors to shapes; duplicating (`Ctrl+D`); pasting shapes into another tab/app (the copy is kept inside the page); bridge commands for shapes; shapes while editing together | Not available. Shapes in the file are kept; for an unknown shape only the anchor, rotation and plain text change. While editing together a shape operation is refused (`CollabUnsupportedError`, `what: 'picture'`) |
| Chart editing (xlsx): pasting a chart **from another tab, application or workbook** (a chart is pasted only in the view it was copied in, onto sheets of the same workbook; the system clipboard gets a placeholder word, not the chart); **copying a chart some of whose formulas are not in the model** (pivot chart, a chart with formulas in an extension or data label: refused); shapes drawn **over** the chart (`c:userShapes`) and **external data / embedded workbook** (`c:externalData`) in a copy — the copy is made without them and this is reported; the frame's **macro** and in-drawing link in a copy (not carried; the macro is reported); changing and copying a chart part declaring an encoding other than UTF-8 / UTF-16 / Latin-1 / Latin-5 (refused); pasting a chart with the menu's Paste when the browser does not give the clipboard's HTML (the keyboard shortcut works); pasting an old copy after a row/column operation that shifts its cells, or after the sheet it names was renamed/deleted (refused: copy again); sheet copies (`copySheet`) do not carry charts; **inserting** radar, bubble, stock and combination charts (those in the file can be edited); **changing the type** of a chart written by another application (refused; free for charts the engine wrote); changing the series list of a multi-type (combination) chart; **editing** per-series category ranges, secondary axes, trendlines, error bars, point formats, line widths, markers, fonts and layout (those in the file are kept); axis units and reverse order; editing a chart on a chart sheet; inserting/moving/deleting pictures and shapes (in the chart); dragging a chart that is not selected on touch in one move (tap to select first); chart commands in the WebView bridge | Not available |
| Charts (xlsx), display: `chartex` **region map** (`regionMap`; needs map data — its fallback is shown) and editing, moving, copying, deleting `chartex` charts; style and colour parts of `chartex` (`style`/`colors`: only theme accent colours are used), data labels in box-and-whisker, single-point labels in treemap / sunburst, axis titles and display units, drawing only the first layout (in a chart with several) and only the first series in a layout (except box-and-whisker); the 3-D perspective of surface charts (drawn from above); the **second plot** of pie-of-bar / pie-of-pie (`ofPieChart`; only the pie is drawn, the second plot's points are one slice); 3-D view (drawn flat); **data labels**: leader lines (`showLeaderLines`), legend key (`showLegendKey`), manual label position, label box fill/border and rotation, rich text formatting (the plain text is written); the **trendline** equation and R² label, a set intercept (`intercept`), its legend entry; error bars widening the axis scale; series lines between stacked bars (`serLines`); **data table** (`c:dTable`); display units (`dispUnits`); minor unit (`minorUnit`) and tick marks; log axes in radar charts; rotated / two-level category labels in bar and scatter charts, date axes in bar charts (drawn as category axes), third and further category levels; more than two value axes (series of the third are drawn on the second); `factor` mode of manual layout (in plot area and legend; applied only to the title); gradient, pattern and picture fills (gradient as first colour, pattern as foreground colour), transparency (the colour as seen on white; bubbles opaque), shadows and effects; chart style number (`c:style`), shapes drawn inside a chart (`userShapes`); charts in xls files (BIFF chart records are not read) | Not available. The chart part is kept byte for byte; a chart of an unread type shows as a placeholder, and what is not drawn in a drawn chart is listed in `chart.notes` |
| Cell notes: editing, deleting and answering threaded comments (Excel 365 "comments"); starting a new thread; moving/resizing a note box and changing "always show"; rich text in notes (bold, colour); copying comment threads and sorting them with their rows; a note button on the editor toolbar; a bridge command for notes; writing to xls | Not available. In a cell with a threaded comment the note commands are disabled (with the reason) and the comment stays byte for byte. If a formatted note's text is edited, the note becomes **plain text in one format** (formats of unedited notes are kept). A range with a comment thread is not sorted; a copied/cut cell's comment thread stays where it is. Clearing a cell's content (Delete) does not delete its note |
| Saving to the same place (`saveTo`): a bridge message to turn auto-save on/off while running and a message field saying it is paused; auto-save on a server without ETags (waits unless `unconditional: true`); reloading or merging the server's file after a 412 (on a conflict the editor only asks "Overwrite / Cancel"; it neither shows nor keeps the server's version); refreshing the ETag after a save without an answer; retrying a failed save by itself while `autoSave` is off; upload progress; changing the 10-minute limit of a function saver; a "save as" / download button; a machine-readable code in `onSaveState.error` (the code is in `SaveError.code` of `onError`) | Not available. Auto-save is turned on with `autoSave` (`onSaveState.autoPaused` tells that it waits); for downloading write a function saver (`examples/html`) |
| Writing archives of 4 GB and more | Not available |
