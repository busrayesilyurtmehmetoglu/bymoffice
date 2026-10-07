import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync, readdirSync } from 'node:fs';
import type { LayoutFont } from '../src/layout/doc/types';
import { CanvasTextMeasurer, type MeasureContext } from '../src/render/doc-measure';
import {
  EMBEDDED_FONT_FACES,
  EMBEDDED_FONT_FAMILY,
  configureEmbeddedFonts,
  embeddedFontFor,
  embeddedFontNames,
  embeddedFontsState,
  onEmbeddedFontNotice,
  resetEmbeddedFonts,
  resolveFontsBase,
} from '../src/render/embedded-fonts';
import { notifyFontsChanged, onFontsChanged, registerFont, registeredFonts, unregisterFont } from '../src/render/fonts';
import { cssFont, FONT_FALLBACK } from '../src/render/text';
import type { Font } from '../src/model/sheet/types';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

// The font the package brings (TeX Gyre Bonum, for Bookman Old Style): where
// its files are looked for, when they are loaded, and what is said when they
// cannot be.

const font = (family: string, size = 16, bold = false, italic = false): LayoutFont => ({ family, size, bold, italic });
const cellFont = (name: string): Font => ({ name, size: 12 }) as Font;

/** The browser's FontFace: an address with "missing" in it does not load, one with "slow" loads on release(). */
class FakeFace {
  static made: FakeFace[] = [];
  static waiting: (() => void)[] = [];
  readonly family: string;
  readonly source: unknown;
  readonly descriptors: { weight: string; style: string };
  constructor(family: string, source: unknown, descriptors: { weight: string; style: string }) {
    this.family = family;
    this.source = source;
    this.descriptors = descriptors;
    FakeFace.made.push(this);
  }
  load(): Promise<unknown> {
    const address = typeof this.source === 'string' ? this.source : '';
    const end = (): Promise<unknown> => (address.indexOf('missing') >= 0 || /bold\.otf.*#nobold|nobold.*bold\.otf/.test(address) ? Promise.reject(new Error('A network error occurred.')) : Promise.resolve(this));
    if (address.indexOf('slow') >= 0) return new Promise<void>((resolve) => FakeFace.waiting.push(resolve)).then(end);
    return end();
  }
}

const pageFonts = new Set<unknown>();
const globals = globalThis as { FontFace?: unknown; document?: unknown };

/** A page with web fonts; `installed` are the families the device has. */
function browser(): void {
  globals.FontFace = FakeFace;
  globals.document = { fonts: { add: (face: unknown) => pageFonts.add(face), delete: (face: unknown) => pageFonts.delete(face) } };
}

function reset(env: { scriptUrl?: string; pageUrl?: string } = {}): void {
  for (const entry of registeredFonts()) unregisterFont(entry.family);
  FakeFace.made.length = 0;
  FakeFace.waiting.length = 0;
  pageFonts.clear();
  browser();
  resetEmbeddedFonts(env);
}

const settle = (): Promise<void> => new Promise((resolve) => setImmediate(() => setImmediate(resolve)));

/** A canvas of a device that has the `installed` families, each with its own width. */
function deviceContext(installed: { [family: string]: number }): MeasureContext {
  return {
    font: '',
    measureText(text: string) {
      const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1] || 10);
      let factor = 0.5;
      const quoted = /"([^"]+)"/g;
      for (let match = quoted.exec(this.font); match; match = quoted.exec(this.font)) {
        if (Object.prototype.hasOwnProperty.call(installed, match[1])) {
          factor = installed[match[1]];
          break;
        }
      }
      return { width: text.length * size * factor, fontBoundingBoxAscent: size * 0.9, fontBoundingBoxDescent: size * 0.3 };
    },
  };
}

const SCRIPT = 'https://cdn.example/npm/bymoffice/dist/bymoffice.min.js';
const PAGE = 'https://app.example/editor/index.html';
const addresses = (): string[] => FakeFace.made.map((face) => String(face.source));

test('embedded font: the four faces are the four files of assets/', () => {
  const files = readdirSync('assets/fonts/tex-gyre-bonum').filter((name) => /\.otf$/.test(name)).sort();
  assert.deepEqual(EMBEDDED_FONT_FACES.map((face) => face.file).sort(), files);
  assert.equal(EMBEDDED_FONT_FAMILY, 'TeX Gyre Bonum');
  assert.deepEqual(
    EMBEDDED_FONT_FACES.map((face) => [face.file, face.weight, face.style]),
    [
      ['texgyrebonum-regular.otf', 400, 'normal'],
      ['texgyrebonum-bold.otf', 700, 'normal'],
      ['texgyrebonum-italic.otf', 400, 'italic'],
      ['texgyrebonum-bolditalic.otf', 700, 'italic'],
    ],
  );
});

test('embedded font: the folder of the fonts is next to the script, or where the application says', () => {
  // A script of a page, from a CDN or from the application's own server.
  assert.equal(resolveFontsBase(undefined, SCRIPT, PAGE), 'https://cdn.example/npm/bymoffice/dist/fonts/');
  assert.equal(resolveFontsBase(undefined, 'https://app.example/vendor/bymoffice.min.js?v=3#x', PAGE), 'https://app.example/vendor/fonts/');
  // A bundle of the application's that is a script of the page: next to that script.
  assert.equal(resolveFontsBase(undefined, 'https://app.example/assets/app.js', PAGE), 'https://app.example/assets/fonts/');
  // A page and a script that are files of the device (a WebView).
  assert.equal(resolveFontsBase(undefined, 'file:///android_asset/www/bymoffice.min.js', 'file:///android_asset/www/index.html'), 'file:///android_asset/www/fonts/');
  assert.equal(resolveFontsBase(undefined, 'app://local/js/bymoffice.min.js', 'app://local/index.html'), 'app://local/js/fonts/');
  // Nothing to go by: an inline script, a module from a blob, a bundle that names the file it was built from.
  assert.equal(resolveFontsBase(undefined, undefined, PAGE), null);
  assert.equal(resolveFontsBase(undefined, '', PAGE), null);
  assert.equal(resolveFontsBase(undefined, 'blob:https://app.example/1b2c', PAGE), null);
  assert.equal(resolveFontsBase(undefined, 'data:text/javascript,1', PAGE), null);
  assert.equal(resolveFontsBase(undefined, 'about:blank', PAGE), null);
  assert.equal(resolveFontsBase(undefined, 'file:///home/someone/project/node_modules/bymoffice/dist/bymoffice.esm.js', PAGE), null);
  assert.equal(resolveFontsBase(undefined, 'file:///home/someone/project/x.js', undefined), null);
  assert.equal(resolveFontsBase(undefined, 'not an address', PAGE), null);
  // What the application sets wins, and is read as an address of the page.
  assert.equal(resolveFontsBase('/static/bym-fonts', SCRIPT, PAGE), 'https://app.example/static/bym-fonts/');
  assert.equal(resolveFontsBase('../fonts/', SCRIPT, PAGE), 'https://app.example/fonts/');
  assert.equal(resolveFontsBase('https://fonts.example/a/', SCRIPT, PAGE), 'https://fonts.example/a/');
  assert.equal(resolveFontsBase('fonts', undefined, undefined), 'fonts/');
  // Turned off.
  assert.equal(resolveFontsBase(null, SCRIPT, PAGE), null);
});

test('embedded font: a document in Bookman Old Style on a device without it loads the four faces, once', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  assert.equal(embeddedFontsState().status, 'idle');
  const measurer = new CanvasTextMeasurer({ context: deviceContext({ Arial: 0.6 }) });
  // Nothing asks for the family yet.
  measurer.cssFamilies('Arial');
  measurer.cssFamilies('Times New Roman');
  measurer.cssFamilies('Bookman');
  assert.deepEqual(addresses(), []);
  assert.equal(embeddedFontsState().status, 'idle');

  measurer.width(font('Bookman Old Style Tur'), 'abc');
  assert.equal(embeddedFontsState().status, 'loading');
  await settle();
  const base = 'https://cdn.example/npm/bymoffice/dist/fonts/tex-gyre-bonum/';
  assert.deepEqual(addresses(), EMBEDDED_FONT_FACES.map((face) => `url("${base}${face.file}")`));
  assert.deepEqual(
    FakeFace.made.map((face) => [face.family, face.descriptors.weight, face.descriptors.style]),
    [
      ['TeX Gyre Bonum', '400', 'normal'],
      ['TeX Gyre Bonum', '700', 'normal'],
      ['TeX Gyre Bonum', '400', 'italic'],
      ['TeX Gyre Bonum', '700', 'italic'],
    ],
  );
  assert.equal(pageFonts.size, 4);
  assert.equal(embeddedFontsState().status, 'loaded');
  // The views are told once for the four faces.
  assert.equal(heard, 1);
  // The registry of the application is not touched.
  assert.deepEqual(registeredFonts(), []);

  // A second view, a spreadsheet and the same view again: nothing more is loaded.
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  measurer.width(font('Bookman Old Style', 20, true), 'abc');
  cssFont(cellFont('Bookman Old Style'));
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(heard, 1);
  off();
});

test('embedded font: nothing is loaded where the device has the family, or a font in front of the embedded one', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  new CanvasTextMeasurer({ context: deviceContext({ 'Bookman Old Style': 0.7 }) }).width(font('Bookman Old Style Tur'), 'abc');
  // The alternative the document names for the font is on the device: that is what is shown.
  const hints = new Map([['Bookman Old Style', { altName: 'Old Bookman' }]]);
  new CanvasTextMeasurer({ context: deviceContext({ 'Old Bookman': 0.7 }), fonts: hints }).width(font('Bookman Old Style'), 'abc');
  // The test of the application decides where it gives one.
  new CanvasTextMeasurer({ context: deviceContext({}), hasFamily: (family) => family === 'Bookman Old Style' }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(addresses(), []);
  assert.equal(embeddedFontsState().status, 'idle');
  // A device that lacks it still gets it afterwards.
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(FakeFace.made.length, 4);
});

test('embedded font: nothing is loaded when the application registered a font for the family', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  await registerFont({ family: 'My Bonum', source: new Uint8Array([0, 1, 0, 0]), as: ['Bookman Old Style'] });
  FakeFace.made.length = 0;
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  assert.equal(measurer.cssFamilies('Bookman Old Style Tur'), '"Bookman Old Style Tur", "Bookman Old Style", "My Bonum", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  measurer.width(font('Bookman Old Style Tur'), 'abc');
  cssFont(cellFont('Bookman Old Style'));
  await settle();
  assert.deepEqual(addresses(), []);
  assert.equal(embeddedFontsState().status, 'idle');

  // One that is still loading counts as well: the application has seen to the family.
  unregisterFont('My Bonum');
  const slow = registerFont({ family: 'Slow Bonum', source: 'slow', as: 'Bookman Old Style' });
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(addresses(), ['url("slow")']);
  for (const resolve of FakeFace.waiting.splice(0)) resolve();
  await slow;

  // And so does a copy of the embedded font the application brought itself, or a font of the very name.
  unregisterFont('Slow Bonum');
  await registerFont({ family: 'TeX Gyre Bonum', source: new Uint8Array([0, 1, 0, 0]) });
  FakeFace.made.length = 0;
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(addresses(), []);
  unregisterFont('TeX Gyre Bonum');
  await registerFont({ family: 'Bookman Old Style', source: new Uint8Array([0, 1, 0, 0]) });
  FakeFace.made.length = 0;
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(addresses(), []);

  // With the application's font gone the embedded one is asked for.
  unregisterFont('Bookman Old Style');
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(FakeFace.made.length, 4);
});

test('embedded font: turned off, nothing is asked for and nothing is offered', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  assert.deepEqual(embeddedFontNames(), ['Bookman Old Style']);
  configureEmbeddedFonts({ baseUrl: null });
  assert.equal(embeddedFontsState().status, 'off');
  assert.equal(embeddedFontsState().baseUrl, null);
  assert.equal(embeddedFontFor('Bookman Old Style'), undefined);
  assert.deepEqual(embeddedFontNames(), []);
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  measurer.width(font('Bookman Old Style'), 'abc');
  // The family list is what it was before there was an embedded font.
  assert.equal(measurer.cssFamilies('Bookman Old Style'), '"Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(cssFont(cellFont('Bookman Old Style')), `16px "Bookman Old Style", ${FONT_FALLBACK}`);
  await settle();
  assert.deepEqual(addresses(), []);
  // Turned on again, with an address.
  configureEmbeddedFonts({ baseUrl: '/static/fonts' });
  assert.equal(embeddedFontsState().baseUrl, 'https://app.example/static/fonts/');
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(addresses()[0], 'url("https://app.example/static/fonts/tex-gyre-bonum/texgyrebonum-regular.otf")');
  assert.equal(FakeFace.made.length, 4);
});

test('embedded font: a spreadsheet cell in the family asks for it, and names it once it is on the page', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  // No canvas here to try the device with: the cell is left as it is.
  assert.equal(cssFont(cellFont('Bookman Old Style Tur')), `16px "Bookman Old Style Tur", "Bookman Old Style", ${FONT_FALLBACK}`);
  assert.deepEqual(addresses(), []);
  assert.equal(cssFont(cellFont('Arial')), `16px "Arial", ${FONT_FALLBACK}`);
  // The test of the device is the application's here.
  configureEmbeddedFonts({ hasFamily: () => false });
  // On its way it is not named yet: the cell is painted anew when it has arrived.
  assert.equal(cssFont(cellFont('Bookman Old Style Tur')), `16px "Bookman Old Style Tur", "Bookman Old Style", ${FONT_FALLBACK}`);
  assert.equal(embeddedFontsState().status, 'loading');
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(cssFont(cellFont('Bookman Old Style Tur')), `16px "Bookman Old Style Tur", "Bookman Old Style", "TeX Gyre Bonum", ${FONT_FALLBACK}`);
  assert.equal(cssFont(cellFont('Bookman Old Style')), `16px "Bookman Old Style", "TeX Gyre Bonum", ${FONT_FALLBACK}`);
  assert.equal(cssFont(cellFont('Arial')), `16px "Arial", ${FONT_FALLBACK}`);

  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ hasFamily: (family) => family === 'Bookman Old Style' });
  cssFont(cellFont('Bookman Old Style Tur'));
  await settle();
  assert.deepEqual(addresses(), []);
});

test('embedded font: files that cannot be loaded are said once, and the text is shown without them', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ baseUrl: 'https://app.example/missing/' });
  const first: { family: string; reason: string }[] = [];
  const second: { family: string; reason: string }[] = [];
  const offFirst = onEmbeddedFontNotice((notice) => first.push(notice));
  const offSecond = onEmbeddedFontNotice((notice) => second.push(notice));
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  measurer.width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(pageFonts.size, 0);
  assert.equal(embeddedFontsState().status, 'failed');
  // One view is told, once.
  assert.equal(first.length, 1);
  assert.equal(second.length, 0);
  assert.equal(first[0].family, 'Bookman Old Style');
  assert.match(first[0].reason, /network error/i);
  assert.equal(heard, 0);
  // It is not tried again for every text, nor said again.
  measurer.width(font('Bookman Old Style', 30), 'abcd');
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style Tur'), 'abc');
  cssFont(cellFont('Bookman Old Style'));
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(first.length, 1);
  assert.deepEqual(embeddedFontNames(), []);
  // The text measures with what the device has.
  assert.equal(measurer.width(font('Bookman Old Style', 10), 'ab'), 10);

  // A new address is a new try.
  configureEmbeddedFonts({ baseUrl: 'https://app.example/fonts/' });
  assert.equal(embeddedFontsState().status, 'idle');
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(FakeFace.made.length, 8);
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.equal(first.length, 1);
  offFirst();
  offSecond();
  off();
});

test('embedded font: an address that cannot be worked out is said once, without a request', async () => {
  reset({ scriptUrl: undefined, pageUrl: PAGE });
  assert.equal(embeddedFontsState().baseUrl, null);
  assert.deepEqual(embeddedFontNames(), []);
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  // The view that asked listens a moment later, as a view does at the end of its constructor.
  const told: { family: string; reason: string }[] = [];
  const off = onEmbeddedFontNotice((notice) => told.push(notice));
  await settle();
  assert.deepEqual(addresses(), []);
  assert.equal(embeddedFontsState().status, 'failed');
  assert.equal(told.length, 1);
  assert.match(told[0].reason, /baseUrl/);
  // A view that comes later is not told again.
  const late: unknown[] = [];
  const offLate = onEmbeddedFontNotice((notice) => late.push(notice));
  await settle();
  assert.deepEqual(late, []);
  off();
  offLate();

  // Nobody listened when it failed: the first view that does is told.
  reset({ scriptUrl: undefined, pageUrl: PAGE });
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  const later: unknown[] = [];
  const offLater = onEmbeddedFontNotice((notice) => later.push(notice));
  await settle();
  assert.equal(later.length, 1);
  offLater();
});

test('embedded font: the application can give the files itself, as bytes', async () => {
  reset({ scriptUrl: undefined, pageUrl: PAGE });
  const asked: string[] = [];
  const data = new Uint8Array([79, 84, 84, 79]);
  configureEmbeddedFonts({
    source: (file) => {
      asked.push(file);
      return data;
    },
  });
  assert.deepEqual(embeddedFontNames(), ['Bookman Old Style']);
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(asked, EMBEDDED_FONT_FACES.map((face) => face.file));
  assert.deepEqual(FakeFace.made.map((face) => face.source), [data, data, data, data]);
  assert.equal(embeddedFontsState().status, 'loaded');

  // A face it does not give is fetched from the folder; one it fails to give is a face that failed.
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  const told: { reason: string }[] = [];
  const off = onEmbeddedFontNotice((notice) => told.push(notice));
  configureEmbeddedFonts({
    source: (file) => {
      if (/bold\.otf$/.test(file)) throw new Error('no such block');
      return /italic/.test(file) ? Promise.resolve(data) : undefined;
    },
  });
  new CanvasTextMeasurer({ context: deviceContext({}) }).width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.deepEqual(FakeFace.made.map((face) => face.source), ['url("https://cdn.example/npm/bymoffice/dist/fonts/tex-gyre-bonum/texgyrebonum-regular.otf")', data, data]);
  // One face is missing: none is used (see "all four faces or none"), and it is said.
  assert.equal(pageFonts.size, 0);
  assert.equal(embeddedFontsState().status, 'failed');
  assert.equal(told.length, 1);
  assert.match(told[0].reason, /no such block/);
  off();
});

test('embedded font: it can stand in for other names, and the names are sanitized', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ as: ['Bookman Old Style', 'Old "Bookman"; x', 'Old Bookman x'] });
  assert.deepEqual(embeddedFontsState().as, ['Bookman Old Style', 'Old Bookman x']);
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  assert.equal(measurer.cssFamilies('Old Bookman x Tur'), '"Old Bookman x Tur", "Old Bookman x", sans-serif');
  assert.equal(embeddedFontsState().status, 'loading');
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(measurer.cssFamilies('Old Bookman x Tur'), '"Old Bookman x Tur", "Old Bookman x", "TeX Gyre Bonum", sans-serif');
  assert.throws(() => configureEmbeddedFonts({ baseUrl: 5 as unknown as string }), TypeError);
  assert.throws(() => configureEmbeddedFonts({ baseUrl: '' }), TypeError);
  assert.throws(() => configureEmbeddedFonts({ source: 'x' as unknown as () => undefined }), TypeError);
  assert.throws(() => configureEmbeddedFonts(null as unknown as object), TypeError);
});

test('embedded font: setting what is already set tells no view, a change does', () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  configureEmbeddedFonts({});
  configureEmbeddedFonts({ baseUrl: undefined });
  assert.equal(heard, 0);
  configureEmbeddedFonts({ baseUrl: '/a/' });
  assert.equal(heard, 1);
  configureEmbeddedFonts({ baseUrl: '/a/' });
  assert.equal(heard, 1);
  configureEmbeddedFonts({ baseUrl: null });
  configureEmbeddedFonts({ baseUrl: null });
  assert.equal(heard, 2);
  off();
});

test('embedded font: outside a browser nothing is tried', () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  delete globals.FontFace;
  delete globals.document;
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  measurer.width(font('Bookman Old Style'), 'abc');
  assert.equal(cssFont(cellFont('Bookman Old Style')), `16px "Bookman Old Style", ${FONT_FALLBACK}`);
  assert.equal(embeddedFontsState().status, 'idle');
  assert.deepEqual(embeddedFontNames(), []);
});

test('embedded font: what a view says when the font cannot be loaded names the family of the document', () => {
  assert.match(STRINGS_EN.embeddedFontNotLoaded('Bookman Old Style'), /"Bookman Old Style"/);
  assert.match(STRINGS_TR.embeddedFontNotLoaded('Bookman Old Style'), /"Bookman Old Style"/);
  assert.notEqual(STRINGS_EN.embeddedFontNotLoaded('x'), STRINGS_TR.embeddedFontNotLoaded('x'));
});

// ------------------------------------------------- review: states and races

/** Lets the faces that wait end, as they end. */
const release = async (): Promise<void> => {
  await settle();
  for (const resolve of FakeFace.waiting.splice(0)) resolve();
  await settle();
};

/** A view: measures text in the family again whenever the fonts change, as an open view does. */
function view(family = 'Bookman Old Style'): { measurer: CanvasTextMeasurer; off(): void } {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  measurer.width(font(family), 'abc');
  const off = onFontsChanged(() => {
    measurer.width(font(family), 'abc');
  });
  return { measurer, off };
}

test('embedded font: an address set while the old one is still being tried is tried when the old one has failed', async () => {
  reset({ scriptUrl: 'https://app.example/slow-missing/x.js', pageUrl: PAGE });
  const told: unknown[] = [];
  const offNotice = onEmbeddedFontNotice((notice) => told.push(notice));
  const open = view();
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(embeddedFontsState().status, 'loading');
  configureEmbeddedFonts({ baseUrl: '/good/' });
  await release();
  // The four requests to the old address failed: that is not the end, and nothing is said of it.
  assert.deepEqual(told, []);
  assert.deepEqual(addresses().slice(4), EMBEDDED_FONT_FACES.map((face) => `url("https://app.example/good/tex-gyre-bonum/${face.file}")`));
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.equal(pageFonts.size, 4);
  open.off();
  offNotice();
});

test('embedded font: turned off and on again while it loads, it is still loaded once', async () => {
  reset({ scriptUrl: 'https://app.example/slow/x.js', pageUrl: PAGE });
  const open = view();
  await settle();
  assert.equal(FakeFace.made.length, 4);
  configureEmbeddedFonts({ baseUrl: null });
  assert.equal(embeddedFontsState().status, 'off');
  configureEmbeddedFonts({ baseUrl: undefined });
  await settle();
  await release();
  assert.equal(FakeFace.made.length, 4, 'one load');
  assert.equal(pageFonts.size, 4, 'four faces on the page, not eight');
  assert.equal(embeddedFontsState().status, 'loaded');
  open.off();
});

test('embedded font: faces that arrive while it is turned off are not put on the page; turned on, they are used without another request', async () => {
  reset({ scriptUrl: 'https://app.example/slow/x.js', pageUrl: PAGE });
  let heard = 0;
  const offHeard = onFontsChanged(() => heard++);
  const open = view();
  await settle();
  configureEmbeddedFonts({ baseUrl: null });
  const before = heard;
  await release();
  assert.equal(pageFonts.size, 0);
  assert.equal(embeddedFontsState().status, 'off');
  assert.equal(heard, before, 'no view is told of faces that are not used');
  assert.equal(open.measurer.cssFamilies('Bookman Old Style'), '"Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  configureEmbeddedFonts({ baseUrl: undefined });
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(pageFonts.size, 4);
  assert.equal(embeddedFontsState().status, 'loaded');
  open.off();
  offHeard();
});

test('embedded font: turned off after it was loaded, it leaves the page, and documents and cells drop it by the same rule', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ as: ['Bookman Old Style', 'Old Bookman x'], hasFamily: () => false });
  const open = view('Old Bookman x');
  const doc = (): string => open.measurer.cssFamilies('Old Bookman x');
  const cell = (): string => cssFont(cellFont('Old Bookman x'));
  // On its way it is in no list yet: a list names it when it can be drawn with.
  assert.equal(doc(), '"Old Bookman x", sans-serif');
  assert.equal(cell(), `16px "Old Bookman x", ${FONT_FALLBACK}`);
  await settle();
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.equal(doc(), '"Old Bookman x", "TeX Gyre Bonum", sans-serif');
  assert.equal(cell(), `16px "Old Bookman x", "TeX Gyre Bonum", ${FONT_FALLBACK}`);
  assert.equal(pageFonts.size, 4);

  let heard = 0;
  const offHeard = onFontsChanged(() => heard++);
  configureEmbeddedFonts({ baseUrl: null });
  assert.equal(heard, 1, 'the views lay their text out again');
  assert.equal(pageFonts.size, 0, 'the faces are taken off the page');
  assert.equal(embeddedFontsState().status, 'off');
  assert.equal(doc(), '"Old Bookman x", sans-serif');
  assert.equal(cell(), `16px "Old Bookman x", ${FONT_FALLBACK}`);

  configureEmbeddedFonts({ baseUrl: undefined });
  await settle();
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.equal(pageFonts.size, 4);
  assert.equal(FakeFace.made.length, 4, 'the files that were read are kept');
  assert.equal(doc(), '"Old Bookman x", "TeX Gyre Bonum", sans-serif');
  assert.equal(cell(), `16px "Old Bookman x", "TeX Gyre Bonum", ${FONT_FALLBACK}`);
  open.off();
  offHeard();
});

test('embedded font: all four faces or none', async () => {
  reset({ scriptUrl: 'https://app.example/dist/x.js#nobold', pageUrl: PAGE });
  configureEmbeddedFonts({ baseUrl: 'https://app.example/fonts/#nobold' });
  const told: { reason: string }[] = [];
  const offNotice = onEmbeddedFontNotice((notice) => told.push(notice));
  let heard = 0;
  const offHeard = onFontsChanged(() => heard++);
  const open = view();
  await settle();
  assert.equal(FakeFace.made.length, 4);
  // Three arrived and one did not: a bold the browser makes up would be set on lines measured for the real one.
  assert.equal(pageFonts.size, 0);
  assert.equal(embeddedFontsState().status, 'failed');
  assert.equal(told.length, 1);
  assert.equal(heard, 0);
  assert.equal(open.measurer.cssFamilies('Bookman Old Style Tur').indexOf('TeX Gyre Bonum') > 0, true, 'the list of the document is the one it always had');
  assert.equal(cssFont(cellFont('Bookman Old Style')), `16px "Bookman Old Style", ${FONT_FALLBACK}`);
  open.off();
  offNotice();
  offHeard();
});

test('embedded font: a new setting after a failure is a new try, and a new failure is said again', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ baseUrl: 'https://app.example/missing/' });
  const told: unknown[] = [];
  const offNotice = onEmbeddedFontNotice((notice) => told.push(notice));
  const open = view();
  await settle();
  assert.equal(told.length, 1);
  configureEmbeddedFonts({ baseUrl: 'https://app.example/missing-too/' });
  await settle();
  assert.equal(FakeFace.made.length, 8);
  assert.equal(embeddedFontsState().status, 'failed');
  assert.equal(told.length, 2);
  open.off();
  offNotice();
});

test('embedded font: it is fetched and comes first although the device has another stand-in of the list', async () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  // A Linux machine: no Bookman Old Style, but URW Bookman.
  const measurer = new CanvasTextMeasurer({ context: deviceContext({ 'URW Bookman': 0.7, 'URW Bookman L': 0.7 }) });
  measurer.width(font('Bookman Old Style'), 'abc');
  await settle();
  assert.equal(FakeFace.made.length, 4);
  assert.equal(embeddedFontsState().status, 'loaded');
  const list = measurer.cssFamilies('Bookman Old Style');
  assert.ok(list.indexOf('"TeX Gyre Bonum"') > 0 && list.indexOf('"TeX Gyre Bonum"') < list.indexOf('"URW Bookman"'), list);
});

/** A canvas whose font box is that of the first family of the list the device has (`installed`, which may change). */
function boxContext(installed: { [family: string]: number }): MeasureContext {
  return {
    font: '',
    measureText(text: string) {
      const size = Number((/(\d+(?:\.\d+)?)px/.exec(this.font) || [])[1] || 10);
      let factor = 0.5;
      const quoted = /"([^"]+)"/g;
      for (let match = quoted.exec(this.font); match; match = quoted.exec(this.font)) {
        if (Object.prototype.hasOwnProperty.call(installed, match[1])) {
          factor = installed[match[1]];
          break;
        }
      }
      return { width: text.length * size * factor, fontBoundingBoxAscent: size * (0.4 + factor), fontBoundingBoxDescent: size * (factor / 2) };
    },
  };
}

test('line metrics: a measurer that is kept follows the fonts as they come and go', () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ baseUrl: null });
  const installed: { [family: string]: number } = {};
  const measurer = new CanvasTextMeasurer({ context: boxContext(installed) });
  const bookman = font('Bookman Old Style', 1000);
  assert.deepEqual(measurer.metrics(bookman), { ascent: 900, descent: 250, lineGap: 0 });
  // The font arrives: the very next line is measured with it.
  installed['TeX Gyre Bonum'] = 0.6;
  notifyFontsChanged();
  assert.deepEqual(measurer.metrics(bookman), { ascent: 1069, descent: 268, lineGap: 0 });
  // And it goes again.
  delete installed['TeX Gyre Bonum'];
  notifyFontsChanged();
  assert.deepEqual(measurer.metrics(bookman), { ascent: 900, descent: 250, lineGap: 0 });
  // clear() forgets as much, without a word from the registry (a font of the page's own style sheet arrived).
  installed['TeX Gyre Bonum'] = 0.6;
  assert.deepEqual(measurer.metrics(bookman), { ascent: 900, descent: 250, lineGap: 0 });
  measurer.clear();
  assert.deepEqual(measurer.metrics(bookman), { ascent: 1069, descent: 268, lineGap: 0 });
});

// ------------------------------------------------------------ line metrics

/** head.unitsPerEm, OS/2 usWinAscent and usWinDescent, and the hhea line of an OpenType file. */
function fileMetrics(path: string): { unitsPerEm: number; winAscent: number; winDescent: number; ascender: number; descender: number; lineGap: number } {
  const bytes = readFileSync(path);
  const tables: { [tag: string]: number } = {};
  const count = bytes.readUInt16BE(4);
  for (let i = 0; i < count; i++) tables[bytes.toString('latin1', 12 + i * 16, 16 + i * 16)] = bytes.readUInt32BE(12 + i * 16 + 8);
  return {
    unitsPerEm: bytes.readUInt16BE(tables['head'] + 18),
    winAscent: bytes.readUInt16BE(tables['OS/2'] + 74),
    winDescent: bytes.readUInt16BE(tables['OS/2'] + 76),
    ascender: bytes.readInt16BE(tables['hhea'] + 4),
    descender: bytes.readInt16BE(tables['hhea'] + 6),
    lineGap: bytes.readInt16BE(tables['hhea'] + 8),
  };
}

test('embedded font: its line metrics in the table are those of the files', () => {
  const measurer = new CanvasTextMeasurer({ context: deviceContext({}) });
  for (const face of EMBEDDED_FONT_FACES) {
    const file = fileMetrics('assets/fonts/tex-gyre-bonum/' + face.file);
    assert.equal(file.unitsPerEm, 1000);
    // The leading Word adds between lines, as GDI works it out.
    const leading = Math.max(0, file.lineGap - (file.winAscent + file.winDescent - (file.ascender - file.descender)));
    const metrics = measurer.metrics(font('TeX Gyre Bonum', 1000, face.weight === 700, face.style === 'italic'));
    assert.deepEqual(metrics, { ascent: file.winAscent, descent: file.winDescent, lineGap: leading }, face.file);
  }
  assert.deepEqual(measurer.metrics(font('TeX Gyre Bonum', 1000)), { ascent: 1069, descent: 268, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('TeX Gyre Bonum', 1000, true)), { ascent: 1153, descent: 306, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('TeX Gyre Bonum', 1000, false, true)), { ascent: 1073, descent: 279, lineGap: 0 });
  assert.deepEqual(measurer.metrics(font('TeX Gyre Bonum', 1000, true, true)), { ascent: 1141, descent: 328, lineGap: 0 });
});

test('line metrics: a family that is not in the table takes the row of the font it is shown with', () => {
  // Bookman Old Style on a device that has only the font that stands in for it: the row of that font, not the box
  // the canvas reports (0.9 and 0.3 of the size here).
  const shown = new CanvasTextMeasurer({ context: deviceContext({ 'TeX Gyre Bonum': 0.6 }) });
  assert.deepEqual(shown.metrics(font('Bookman Old Style', 1000)), { ascent: 1069, descent: 268, lineGap: 0 });
  assert.deepEqual(shown.metrics(font('Bookman Old Style Tur', 1000, true)), { ascent: 1153, descent: 306, lineGap: 0 });
  assert.deepEqual(shown.metrics(font('Bookman Old Style', 500, true, true)), { ascent: 570.5, descent: 164, lineGap: 0 });
  // The device has the family itself, whose metrics are not known: the box of the canvas, as before.
  const real = new CanvasTextMeasurer({ context: deviceContext({ 'Bookman Old Style': 0.7, 'TeX Gyre Bonum': 0.6 }) });
  const box = real.metrics(font('Bookman Old Style', 10));
  assert.ok(Math.abs(box.ascent - 9) < 1e-9 && Math.abs(box.descent - 3) < 1e-9 && box.lineGap === 0, JSON.stringify(box));
  // Nothing of the list is there: the box of the generic font, as before.
  const none = new CanvasTextMeasurer({ context: deviceContext({}) }).metrics(font('Bookman Old Style', 10));
  assert.ok(Math.abs(none.ascent - 9) < 1e-9 && Math.abs(none.descent - 3) < 1e-9 && none.lineGap === 0, JSON.stringify(none));
  // Only the font of the package lends its row. Another family shown with a font of the table (an alternative the
  // document names, a substitute) keeps the box of the canvas, as it always had: those rows say how Word sets a
  // document that asks for the family, not how this one was set.
  const hints = new Map([['House Sans', { altName: 'Arial' }]]);
  const alt = new CanvasTextMeasurer({ context: deviceContext({ Arial: 0.6 }), fonts: hints });
  assert.deepEqual(alt.metrics(font('House Sans', 2000)), { ascent: 1800, descent: 600, lineGap: 0 });
  const math = new CanvasTextMeasurer({ context: deviceContext({ Caladea: 0.6 }) });
  assert.deepEqual(math.metrics(font('Cambria Math', 2000)), { ascent: 1800, descent: 600, lineGap: 0 });
  // The application's test of the device is the one asked.
  const asked = new CanvasTextMeasurer({ context: deviceContext({}), hasFamily: (family) => family === 'TeX Gyre Bonum' });
  assert.deepEqual(asked.metrics(font('Bookman Old Style', 1000)), { ascent: 1069, descent: 268, lineGap: 0 });
});

/** The table as it was before fonts could have a row per face: [ascent, descent, leading] in 1/2048 em. */
const TABLE_BEFORE: { [family: string]: [number, number, number] } = {
  arial: [1854, 434, 67],
  'liberation sans': [1854, 434, 67],
  arimo: [1854, 434, 67],
  helvetica: [1854, 434, 67],
  'times new roman': [1825, 443, 87],
  'liberation serif': [1825, 443, 87],
  tinos: [1825, 443, 87],
  times: [1825, 443, 87],
  'courier new': [1705, 615, 0],
  'liberation mono': [1705, 615, 0],
  cousine: [1705, 615, 0],
  courier: [1705, 615, 0],
  'arial narrow': [1888, 431, 31],
  'arial black': [2254, 634, 0],
  'arial unicode ms': [2189, 555, 0],
  verdana: [2059, 430, 0],
  tahoma: [2049, 423, 0],
  georgia: [1878, 449, 0],
  'trebuchet ms': [1923, 455, 0],
  'comic sans ms': [2257, 597, 0],
  impact: [2066, 432, 0],
  'microsoft sans serif': [1888, 430, 0],
  'andale mono': [1858, 446, 0],
  'arial rounded mt bold': [1938, 432, 0],
  'brush script mt': [1820, 692, 0],
  rockwell: [2099, 600, 0],
  webdings: [1638, 410, 0],
  wingdings: [1841, 432, 0],
  'wingdings 2': [1727, 432, 0],
  'wingdings 3': [1900, 432, 0],
  calibri: [1950, 550, 0],
  'calibri light': [1950, 550, 0],
  carlito: [1950, 550, 0],
  cambria: [1946, 455, 0],
  caladea: [1946, 455, 0],
  'segoe ui': [2210, 514, 0],
  consolas: [1884, 514, 0],
  garamond: [1765, 536, 0],
  'palatino linotype': [1499, 582, 682],
  'century gothic': [2060, 451, 0],
  symbol: [2059, 450, 0],
};

test('line metrics: the families that were in the table measure exactly as before, whatever the device has', () => {
  // Three devices: none of the fonts, all of them, and one whose canvas must not be asked at all.
  const everything: { [family: string]: number } = { 'TeX Gyre Bonum': 0.61 };
  for (const family of Object.keys(TABLE_BEFORE)) everything[family.replace(/(^| )[a-z]/g, (letter) => letter.toUpperCase())] = 0.55;
  const silent: MeasureContext = {
    font: '',
    measureText() {
      throw new Error('the canvas was asked for a family of the table');
    },
  };
  const measurers = [new CanvasTextMeasurer({ context: deviceContext({}) }), new CanvasTextMeasurer({ context: deviceContext(everything) }), new CanvasTextMeasurer({ context: silent })];
  let checked = 0;
  for (const family of Object.keys(TABLE_BEFORE)) {
    const row = TABLE_BEFORE[family];
    const title = family.replace(/(^| )[a-z]/g, (letter) => letter.toUpperCase());
    for (const name of [family, title, title + ' Tur', family.toUpperCase() + ' (Hebrew)']) {
      for (const size of [0.5, 9.33, 12, 14.67, 16, 96, 2048]) {
        for (const [bold, italic] of [[false, false], [true, false], [false, true], [true, true]] as [boolean, boolean][]) {
          // The arithmetic of the measurer before the change, kept here word for word.
          const expected = { ascent: (row[0] * size) / 2048, descent: (row[1] * size) / 2048, lineGap: (row[2] * size) / 2048 };
          for (const measurer of measurers) {
            assert.deepEqual(measurer.metrics(font(name, size, bold, italic)), expected, `${name} ${size} ${bold} ${italic}`);
            checked++;
          }
        }
      }
    }
  }
  assert.equal(checked, 41 * 4 * 7 * 4 * 3);
});

/** Families that are not in the table: with a substitute list of their own, and without. */
const NOT_IN_TABLE = [
  'Book Antiqua', 'Century Schoolbook', 'Cambria Math', 'Century', 'Constantia', 'Candara', 'Corbel', 'Lucida Console', 'Lucida Sans Unicode',
  'Franklin Gothic Medium', 'Gill Sans MT', 'Perpetua', 'Baskerville Old Face', 'Bell MT', 'Bodoni MT', 'Calisto MT', 'Californian FB', 'Centaur',
  'Elephant', 'Eras Medium ITC', 'Footlight MT Light', 'Goudy Old Style', 'Haettenschweiler', 'High Tower Text', 'Maiandra GD', 'Papyrus',
  'Tw Cen MT', 'Segoe Print', 'MS Mincho', 'SimSun', 'House Sans',
];

/** What devices have, each family with a box of its own. */
const DEVICES: { [name: string]: { [family: string]: number } } = {
  none: {},
  windows: { Arial: 0.51, 'Times New Roman': 0.52, 'Courier New': 0.53, Calibri: 0.54, Cambria: 0.55, 'Cambria Math': 0.56, 'Book Antiqua': 0.57, 'Palatino Linotype': 0.58, 'Bookman Old Style': 0.59, 'Century Gothic': 0.61, Georgia: 0.62 },
  mac: { Arial: 0.51, 'Times New Roman': 0.52, 'Courier New': 0.53, Helvetica: 0.63, Palatino: 0.64, Times: 0.65, Courier: 0.66, Georgia: 0.62 },
  linux: { 'Liberation Sans': 0.67, 'Liberation Serif': 0.68, 'Liberation Mono': 0.69, Carlito: 0.71, Caladea: 0.72, 'URW Bookman': 0.73, 'TeX Gyre Pagella': 0.74, 'TeX Gyre Schola': 0.75 },
  bonum: { 'TeX Gyre Bonum': 0.76, Arial: 0.51 },
  everything: { Arial: 0.51, 'Times New Roman': 0.52, 'Courier New': 0.53, Calibri: 0.54, Cambria: 0.55, 'Book Antiqua': 0.57, 'Palatino Linotype': 0.58, 'Bookman Old Style': 0.59, Carlito: 0.71, Caladea: 0.72, 'URW Bookman': 0.73, 'TeX Gyre Pagella': 0.74, 'TeX Gyre Bonum': 0.76, 'TeX Gyre Schola': 0.75, 'Old Friend': 0.77 },
};

test('line metrics: a family that is not in the table measures exactly as before on every device, with and without an alternative; only Bookman Old Style shown with the font of the package differs', () => {
  reset({ scriptUrl: SCRIPT, pageUrl: PAGE });
  configureEmbeddedFonts({ baseUrl: null });
  const alternatives = ['Arial', 'Times New Roman', 'Calibri', 'Old Friend', 'TeX Gyre Bonum'];
  let checked = 0;
  let different = 0;
  for (const device of Object.keys(DEVICES)) {
    for (const withAlternative of [false, true]) {
      const families = NOT_IN_TABLE.concat(['Bookman Old Style', 'Bookman Old Style Tur']);
      const hints = new Map<string, { altName?: string }>();
      if (withAlternative) families.forEach((family, index) => hints.set(family, { altName: alternatives[index % alternatives.length] }));
      const context = boxContext(DEVICES[device]);
      const measurer = new CanvasTextMeasurer({ context, fonts: hints });
      for (const family of families) {
        for (const [bold, italic] of [[false, false], [true, true]] as [boolean, boolean][]) {
          for (const size of [11, 16, 1000]) {
            // The measurer before the change, word for word: the font box of the face at 1000 px, scaled.
            context.font = (italic ? 'italic ' : '') + (bold ? 'bold ' : '') + '1000px ' + measurer.cssFamilies(family);
            const box = context.measureText('Hg');
            const before = { ascent: ((box.fontBoundingBoxAscent as number) / 1000) * size, descent: ((box.fontBoundingBoxDescent as number) / 1000) * size, lineGap: 0 };
            const now = measurer.metrics(font(family, size, bold, italic));
            // Shown with the font of the package: the first family of the list the device has is that font.
            const list = measurer.cssFamilies(family).split(', ').map((name) => name.replace(/"/g, ''));
            const shown = list.filter((name) => Object.prototype.hasOwnProperty.call(DEVICES[device], name))[0];
            checked++;
            if (shown === 'TeX Gyre Bonum') {
              const row = bold ? [1141, 328] : [1069, 268];
              assert.deepEqual(now, { ascent: (row[0] * size) / 1000, descent: (row[1] * size) / 1000, lineGap: 0 }, `${device} ${family}`);
              assert.ok(/^Bookman Old Style/.test(family) || withAlternative, `${device} ${family}: only what is shown with the font of the package`);
              different++;
            } else {
              assert.deepEqual(now, before, `${device} ${withAlternative} ${family} ${size} ${bold}`);
            }
          }
        }
      }
    }
  }
  assert.equal(checked, 6 * 2 * 33 * 2 * 3);
  assert.ok(different > 0);
});

test('embedded font: an ES module does not know where it is: nothing is fetched, it is said once, and the notice says what to do', async () => {
  // In a module there is no script element that runs it (document.currentScript is null), and import.meta is not used.
  reset({ scriptUrl: undefined, pageUrl: PAGE });
  const told: { family: string; reason: string }[] = [];
  const off = onEmbeddedFontNotice((notice) => told.push(notice));
  const open = view();
  await settle();
  assert.deepEqual(addresses(), []);
  assert.equal(embeddedFontsState().status, 'failed');
  assert.equal(told.length, 1);
  assert.match(told[0].reason, /fontsUrl/);
  for (const strings of [STRINGS_EN, STRINGS_TR]) {
    assert.match(strings.embeddedFontNotLoaded('Bookman Old Style'), /fontsUrl/);
    assert.match(strings.embeddedFontNotLoaded('Bookman Old Style'), /fonts: false/);
  }
  // The application says where the folder is, as it would with new URL('./fonts/', import.meta.url).href.
  configureEmbeddedFonts({ baseUrl: 'https://app.example/dist/fonts/' });
  await settle();
  assert.equal(addresses()[0], 'url("https://app.example/dist/fonts/tex-gyre-bonum/texgyrebonum-regular.otf")');
  assert.equal(embeddedFontsState().status, 'loaded');
  assert.equal(told.length, 1);
  open.off();
  off();
});
