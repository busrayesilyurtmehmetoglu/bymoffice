// Where the font files of a page are: what a PDF embeds a font from. The
// registry of the application and the font of the package remember the bytes
// they were given, or the address they loaded a face from; a font of the
// device has neither.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPdfFonts } from '../src/pdf/vector';
import { EMBEDDED_FONT_FAMILY, configureEmbeddedFonts, embeddedFontFiles, embeddedFontFor, embeddedFontsState, resetEmbeddedFonts } from '../src/render/embedded-fonts';
import { registerFont, registeredFontFiles, registeredFonts, unregisterFont } from '../src/render/fonts';
import { bonum, testFont } from './pdf-font-fixtures';

/** The browser's FontFace, as far as the registry uses it: an address with "slow" in it loads on release(), one with "missing" does not load. */
class FakeFace {
  static waiting: (() => void)[] = [];
  constructor(
    readonly family: string,
    readonly source: unknown,
    readonly descriptors: unknown,
  ) {}
  load(): Promise<unknown> {
    const address = typeof this.source === 'string' ? this.source : '';
    if (address.indexOf('missing') >= 0) return Promise.reject(new Error('A network error occurred.'));
    if (address.indexOf('slow') >= 0) return new Promise<void>((resolve) => FakeFace.waiting.push(resolve)).then(() => this);
    return Promise.resolve(this);
  }
}

const globals = globalThis as { FontFace?: unknown; document?: unknown };
const pageFonts = new Set<unknown>();

function reset(): void {
  for (const entry of registeredFonts()) unregisterFont(entry.family);
  FakeFace.waiting.length = 0;
  pageFonts.clear();
  globals.FontFace = FakeFace;
  globals.document = { fonts: { add: (face: unknown) => pageFonts.add(face), delete: (face: unknown) => pageFonts.delete(face) } };
  resetEmbeddedFonts({ scriptUrl: 'https://cdn.example/pkg/dist/bymoffice.min.js', pageUrl: 'https://app.example/' });
}

const settle = (): Promise<void> => new Promise((resolve) => setImmediate(() => setImmediate(resolve)));

test('font files: a registered font remembers the bytes it was given, or the address it was loaded from', async () => {
  reset();
  const bytes = testFont();
  await registerFont({ family: 'Given', source: bytes });
  const buffer = testFont({ name: 'Buffered' }).buffer as ArrayBuffer;
  await registerFont({ family: 'Buffered', source: buffer, weight: 'bold' });
  await registerFont({ family: 'Fetched', source: 'https://fonts.example/a.ttf', style: 'italic' });
  const blobBytes = testFont({ name: 'Blobbed' });
  await registerFont({ family: 'Blobbed', source: new Blob([blobBytes as unknown as BlobPart]) });
  const files = registeredFontFiles();
  assert.deepEqual(
    files.map((file) => [file.family, file.weight, file.style, file.url, file.bytes ? file.bytes.length : undefined]),
    [
      ['Given', 400, 'normal', undefined, bytes.length],
      ['Buffered', 700, 'normal', undefined, buffer.byteLength],
      ['Fetched', 400, 'italic', 'https://fonts.example/a.ttf', undefined],
      ['Blobbed', 400, 'normal', undefined, blobBytes.length],
    ],
  );
  assert.equal(files[0].bytes, bytes, 'the bytes themselves, not a copy');
  assert.equal((files[1].bytes as Uint8Array).buffer, buffer, 'a view of the buffer that was given');
  assert.deepEqual(Array.from(files[3].bytes as Uint8Array), Array.from(blobBytes), 'the bytes the Blob held');
  // A face that is still loading, and one that could not be loaded, have no file to embed.
  const slow = registerFont({ family: 'Slow', source: 'https://fonts.example/slow.ttf' });
  await assert.rejects(registerFont({ family: 'Gone', source: 'https://fonts.example/missing.ttf' }));
  await settle();
  assert.deepEqual(
    registeredFontFiles().map((file) => file.family),
    ['Given', 'Buffered', 'Fetched', 'Blobbed'],
  );
  for (const resolve of FakeFace.waiting.splice(0)) resolve();
  await slow;
  assert.equal(registeredFontFiles().length, 5);
  // A face registered again has the file of the new registration; one that is taken out has none.
  const other = testFont({ name: 'Other' });
  await registerFont({ family: 'Given', source: other });
  assert.equal(registeredFontFiles().filter((file) => file.family === 'Given')[0].bytes, other);
  unregisterFont('Given');
  assert.equal(registeredFontFiles().filter((file) => file.family === 'Given').length, 0);
  reset();
});

test('font files: the font of the package has its four files once it is on the page, as bytes or as addresses', async () => {
  reset();
  assert.deepEqual(embeddedFontFiles(), [], 'nothing before the font was needed');
  // The device has no Bookman Old Style: asking for it fetches the font of the package from next to the script.
  configureEmbeddedFonts({ hasFamily: () => false });
  embeddedFontFor('Bookman Old Style');
  await settle();
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.deepEqual(
    embeddedFontFiles().map((file) => [file.family, file.weight, file.style, file.url]),
    [
      [EMBEDDED_FONT_FAMILY, 400, 'normal', 'https://cdn.example/pkg/dist/fonts/tex-gyre-bonum/texgyrebonum-regular.otf'],
      [EMBEDDED_FONT_FAMILY, 700, 'normal', 'https://cdn.example/pkg/dist/fonts/tex-gyre-bonum/texgyrebonum-bold.otf'],
      [EMBEDDED_FONT_FAMILY, 400, 'italic', 'https://cdn.example/pkg/dist/fonts/tex-gyre-bonum/texgyrebonum-italic.otf'],
      [EMBEDDED_FONT_FAMILY, 700, 'italic', 'https://cdn.example/pkg/dist/fonts/tex-gyre-bonum/texgyrebonum-bolditalic.otf'],
    ],
  );
  // Turned off, the font is no font of the page any more.
  configureEmbeddedFonts({ baseUrl: null });
  assert.deepEqual(embeddedFontFiles(), []);
  // A page that brings the files itself (the page of the WebView bridge) gives their bytes.
  reset();
  const given: string[] = [];
  configureEmbeddedFonts({
    hasFamily: () => false,
    source: (file) => {
      given.push(file);
      return bonum(/bolditalic/.test(file) ? 'bolditalic' : /bold/.test(file) ? 'bold' : /italic/.test(file) ? 'italic' : 'regular');
    },
  });
  embeddedFontFor('Bookman Old Style');
  await settle();
  const files = embeddedFontFiles();
  assert.equal(files.length, 4);
  assert.ok(
    files.every((file) => file.bytes instanceof Uint8Array && file.bytes.length > 100000 && file.url === undefined),
    'the bytes of the four files',
  );
  assert.equal(given.length, 4);
  reset();
});

test('font files: the fonts of the page are what a PDF is written with: the font of the package for the name it stands in for, a registered font under its name', async () => {
  reset();
  configureEmbeddedFonts({ hasFamily: (family) => family === 'Arial', source: (file) => bonum(/bolditalic/.test(file) ? 'bolditalic' : /bold/.test(file) ? 'bold' : /italic/.test(file) ? 'italic' : 'regular') });
  embeddedFontFor('Bookman Old Style');
  await settle();
  await registerFont({ family: 'Boxes', source: testFont(), as: ['Calibri'], prefer: true });
  const fonts = await loadPdfFonts();
  const name = (font: string): string => {
    const resolved = fonts.resolve(font);
    return 'face' in resolved ? resolved.face.name : resolved.reason;
  };
  // The lists are those the painters write: the name of the file first, what stands in for it behind.
  assert.equal(name('16px "Bookman Old Style", "TeX Gyre Bonum", serif'), 'TeXGyreBonum-Regular');
  assert.equal(name('italic bold 16px "Bookman Old Style Tur", "Bookman Old Style", "TeX Gyre Bonum", serif'), 'TeXGyreBonum-BoldItalic');
  assert.equal(name('14.67px "Boxes", "Calibri", "Carlito", "Calibri", Arial, sans-serif'), 'BymTest-Regular');
  // The device has Arial, and nobody has its file.
  assert.match(name('16px "Arial", "TeX Gyre Bonum", serif'), /font of the device \("Arial"\)/);
  assert.match(name('12px "Carlito", "Calibri", Arial, sans-serif'), /font of the device \("Arial"\)/);
  reset();
});
