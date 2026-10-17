# BYM Office

**Open, edit and save Word and Excel files in the browser — docx, xlsx, doc, xls and RTF.**
A document and spreadsheet engine written from scratch in TypeScript: one script, no server, no runtime dependencies.

[Live demo](https://busrayesilyurtmehmetoglu.github.io/bymoffice/) · [Türkçe](README.tr.md) · [Developer guide (Turkish)](docs/kullanim-kilavuzu.md) · [WebView bridge](docs/webview-koprusu.md) · [Changelog](CHANGELOG.md)

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>
<div id="editor" style="height: 600px"></div>
<script>
  BymOffice.mount(document.getElementById('editor'), {
    file: '/files/report.xlsx', // a URL, a File/Blob or bytes; the format is detected from the content
    saveTo: true,               // Save button and Ctrl/Cmd+S write the file back to the same URL
  });
</script>
```

## Why

- **Runs entirely in the browser.** Files are read, edited and written on the device; nothing is uploaded unless you save it somewhere yourself (the engine only fetches its bundled font from beside the script, when a document needs it). The same code runs in Node for server-side work.
- **One file, zero dependencies.** Add one `<script>` from a CDN or install from npm. The package (4.7 MB packed) includes a font for documents written in Bookman Old Style.
- **Lossless.** Parts of a file the engine does not understand are kept byte for byte; only what you changed is rewritten. An unchanged docx or xlsx is saved byte for byte as it was opened.
- **A ready editor, not only a parser.** Toolbars, keyboard and touch, undo/redo, clipboard, find and replace, print, PDF export — in Turkish or English.
- **Web pages and apps.** Any web page or framework; mobile and desktop apps through a single-file WebView page.

## What it does

**Excel (xlsx; xls opens read-only or is converted to xlsx)**
- Canvas grid that draws only what is visible; frozen panes, hidden rows/columns, sheet tabs.
- Cell editing, formatting toolbar, number formats (Turkish and English locales), merge, borders, conditional formats, data validation, filter and sort.
- Formulas: 265 functions, incremental recalculation, dynamic arrays (spills, `A1#`), `LET` and `LAMBDA`.
- Charts (draw, insert, edit, copy), pivot tables (refresh, create, change fields, delete), pictures and shapes, notes, sheet protection.
- Print, print preview, page setup, PDF export.

**Word (docx; doc and RTF open read-only or are converted to docx)**
- Paginated layout with headers/footers, footnotes, tables, lists, images, floating objects, text boxes, charts.
- Editing with cursor, selection, IME, formatting, styles, lists, tables, images, links, tab stops, page setup.
- Track changes (tracked typing, deleting and formatting; accept/reject), comments with a side panel, document protection.
- Print and PDF export with selectable text, links and an outline.

**Both**
- Password-protected files: open encrypted docx/xlsx/doc/xls; save docx/xlsx with the same password.
- Real-time collaboration: several people edit the same workbook or document. The relay only orders messages and can be written in any backend; a dependency-free Node example ships in `server/`.
- Save back to where the file came from (`saveTo`, ETag/If-Match, conflict prompt), or auto-save.
- WebView bridge for mobile and desktop apps (`dist/bymoffice.webview.html`, works offline).

What is **not** there yet is listed honestly in the guide's "Henüz olmayanlar" table. The most important gap: doc, xls and RTF cannot be saved in their own formats — they are saved as docx/xlsx.

## Install

**CDN** — no build step, defines a global `BymOffice`:

```html
<script src="https://cdn.jsdelivr.net/npm/bymoffice@0.1.0/dist/bymoffice.min.js"></script>
```

**npm**

```sh
npm i bymoffice
```

```js
import { mount, open } from 'bymoffice';

// Editor in a page
const view = await mount(document.getElementById('editor'), { file: blobOrUrl, locale: 'en' });

// No UI: read, change, write (browser or Node)
const file = await open(bytes);
const saved = await file.save();
```

When you bundle the package, copy `node_modules/bymoffice/dist/fonts/` to your public files and pass its address: `mount(el, { fontsUrl: '/assets/bymoffice-fonts/' })`.

Only the ES module and the `<script>` build are shipped (no CommonJS build). TypeScript types are included.

## Frameworks and apps

| Where | How |
|---|---|
| Plain HTML, Laravel/Blade, Django, Rails, PHP … | `<script>` from the CDN, then `BymOffice.mount(...)` |
| React, Vue, Svelte, Angular | `import { mount } from 'bymoffice'` in an effect / `onMounted`; call `destroy()` on unmount |
| Next.js, Nuxt | Client side only: a `'use client'` component, or a dynamic import with SSR off |
| Flutter, React Native, native iOS/Android, Electron, Tauri | Load `bymoffice.webview.html` in a WebView and talk to it with messages (`load`, `save`, `changed`, `error` …) — see [the bridge contract](docs/webview-koprusu.md) |
| Node | `open()`, the edit operations and `save()` work without a screen |

The table is a recipe, not a test report: so far the engine is used in a Laravel page and in a Flutter web app; Next.js, React Native, Electron, Tauri and real mobile WebViews have not been tried yet. The engine itself is tested on every change in Chromium, Firefox, WebKit and two mobile profiles (iPhone, Pixel). Old WebViews without `DecompressionStream` fall back to the engine's own inflate/deflate.

## Development

```sh
pnpm install
pnpm build      # dist/: bymoffice.min.js (global BymOffice), bymoffice.esm.js, types, fonts, WebView page
pnpm site       # the demo site in site-dist/
pnpm demo       # developer test page: http://127.0.0.1:4318/demo/index.html
pnpm verify     # type check, unit tests, browser tests, package check, round-trip over a file corpus
```

The guide and the code comments explain the rules the engine keeps (lossless saving, what is refused rather than silently lost, limits that keep hostile files from freezing the page).

## Fonts and third-party content

The package includes **TeX Gyre Bonum** (GUST e-foundry), metrically close to Bookman Old Style, for devices that lack that font. The files are shipped unchanged under the GUST Font License and are loaded only when needed. Details: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## License

[MIT](LICENSE)
