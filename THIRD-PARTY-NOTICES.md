# Third-party notices

BYM Office itself is under the MIT licence (see `LICENSE`). The package also
contains the following work of others, under its own licence.

## TeX Gyre Bonum

- **What:** the font family TeX Gyre Bonum, version 2.004: four OpenType files,
  `texgyrebonum-regular.otf`, `texgyrebonum-bold.otf`, `texgyrebonum-italic.otf`
  and `texgyrebonum-bolditalic.otf`.
- **Authors:** Bogusław Jackowski and Janusz M. Nowacki, GUST e-foundry, on the
  basis of the URW Bookman L fonts of URW++.
- **Source:** <https://www.gust.org.pl/projects/e-foundry/tex-gyre/bonum>; on
  CTAN under `fonts/tex-gyre`.
- **Licence:** the GUST Font License (GFL), which is the LaTeX Project Public
  License, version 1.3c or later, with one added request. Its text is shipped
  next to the fonts, in `dist/fonts/tex-gyre-bonum/GUST-FONT-LICENSE.txt`, and
  is at <http://tug.org/fonts/licenses/GUST-FONT-LICENSE.txt>; the LaTeX
  Project Public License is at <https://www.latex-project.org/lppl.txt>.
- **Where in the package:** `dist/fonts/tex-gyre-bonum/`, and as base64 inside
  `dist/bymoffice.webview.html` (the same bytes, with the text of the licence
  in a comment next to them). The two scripts, `dist/bymoffice.min.js` and
  `dist/bymoffice.esm.js`, do not contain the fonts.
- **Complete distribution:** only the OpenType files are distributed here;
  the manifest (`MANIFEST-TeX-Gyre-Bonum.txt`), the documentation and the
  other font formats are in the original distribution, which is available,
  complete and unmodified, from the address above and from CTAN,
  <https://ctan.org/pkg/tex-gyre-bonum>.
- **No endorsement:** the authors of the fonts have no part in BYM Office and
  do not support or endorse it.
- **Changes:** none. The files are distributed exactly as GUST e-foundry
  publishes them, under their own names; they are not subset, converted or
  renamed. The build compares them with their known checksums.
- **Use:** text that a document or a spreadsheet sets in Bookman Old Style is
  drawn with this font on devices that do not have Bookman Old Style. The name
  stored in the file is not changed.

The MIT licence of BYM Office does not apply to these font files.
