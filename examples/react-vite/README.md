# BYM Office with React and Vite

A React app (`<StrictMode>`) with the editor in a component, `src/OfficeEditor.jsx`.

```sh
npm install
npm run dev      # http://localhost:5173/  (?new=xlsx for a workbook, ?file=/some.docx to open a file)
npm run build    # dist/, ready to be served as static files
```

The component mounts the editor in an effect and destroys it in the effect's cleanup. `mount()` is asynchronous, so the cleanup can run before it resolves; an editor that arrives after its cleanup is destroyed at once. That is what keeps one editor in development, where `<StrictMode>` runs every effect twice.

**The one step a bundler needs: the font folder.** The package brings a font (TeX Gyre Bonum, used for documents set in Bookman Old Style) as files next to its scripts, in `node_modules/bymoffice/dist/fonts/`. A bundler does not copy them. `vite.config.js` copies that folder to `public/bymoffice-fonts/` each time Vite starts, so the dev server serves it and `npm run build` puts it into `dist/bymoffice-fonts/`; the editor is told where it is:

```js
mount(element, { file, fontsUrl: import.meta.env.BASE_URL + 'bymoffice-fonts/' });
```

Keep the licence file of the font (`GUST-FONT-LICENSE.txt`) in the copied folder. If you do not want the font, pass `fonts: false` instead and leave the copy out.

`vite build` warns that a chunk is larger than 500 kB: the engine is one module of about 3 MB minified (about 1 MB gzipped). The warning is harmless; raise `build.chunkSizeWarningLimit` in `vite.config.js` if you want it gone.
