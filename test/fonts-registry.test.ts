import { test } from 'node:test';
import assert from 'node:assert/strict';
import type { LayoutFont } from '../src/layout/doc/types';
import { CanvasTextMeasurer, type MeasureContext } from '../src/render/doc-measure';
import { fontFamilyChain, fontsVersion, onFontsChanged, registerFont, registeredFonts, replacementFonts, settleFonts, unregisterFont, type FontRegistration } from '../src/render/fonts';
import { cssFont, FONT_FALLBACK, TextPainter } from '../src/render/text';
import type { Font } from '../src/model/sheet/types';

const font = (family: string, size = 16): LayoutFont => ({ family, size, bold: false, italic: false });
const cellFont = (name: string): Font => ({ name, size: 12 }) as Font;
const bytes = new Uint8Array([0, 1, 0, 0]);

/** The browser's FontFace, as far as the registry uses it. `source` "fail" does not load; "slow" loads when release() is called. */
class FakeFace {
  static made: FakeFace[] = [];
  static waiting: (() => void)[] = [];
  static refuseUnquotedDigits = false;
  /** Names the browser does not take as they are (keywords of CSS, say). */
  static refuseUnquoted: string[] = [];
  readonly family: string;
  readonly source: unknown;
  readonly descriptors: unknown;
  constructor(family: string, source: unknown, descriptors: unknown) {
    if (FakeFace.refuseUnquotedDigits && /(^| )[0-9]/.test(family) && family.charAt(0) !== '"') {
      throw Object.assign(new Error('not a family name'), { name: 'SyntaxError' });
    }
    if (FakeFace.refuseUnquoted.indexOf(family) >= 0) throw Object.assign(new Error('not a family name'), { name: 'SyntaxError' });
    this.family = family;
    this.source = source;
    this.descriptors = descriptors;
    FakeFace.made.push(this);
  }
  load(): Promise<unknown> {
    if (this.source === 'url("fail")') return Promise.reject(new Error('A network error occurred.'));
    if (this.source === 'url("slow")') return new Promise<void>((resolve) => FakeFace.waiting.push(resolve)).then(() => this);
    return Promise.resolve(this);
  }
}

const pageFonts = new Set<unknown>();
const globals = globalThis as { FontFace?: unknown; document?: unknown };

function browser(): void {
  globals.FontFace = FakeFace;
  globals.document = { fonts: { add: (face: unknown) => pageFonts.add(face), delete: (face: unknown) => pageFonts.delete(face) } };
}

function reset(): void {
  for (const entry of registeredFonts()) unregisterFont(entry.family);
  FakeFace.made.length = 0;
  FakeFace.waiting.length = 0;
  FakeFace.refuseUnquotedDigits = false;
  FakeFace.refuseUnquoted = [];
  pageFonts.clear();
}

/** Lets the faces that wait finish loading; the face of a registration is made a moment after the call. */
const release = async (): Promise<void> => {
  await new Promise((resolve) => setImmediate(resolve));
  for (const resolve of FakeFace.waiting.splice(0)) resolve();
  await new Promise((resolve) => setImmediate(resolve));
};

/** A canvas of a device that has the `installed` families (see font-names.test.ts). */
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

test('fonts: outside a browser registering says why it cannot be done', async () => {
  assert.equal(typeof globals.FontFace, 'undefined');
  await assert.rejects(registerFont({ family: 'Test Sans', source: bytes }), /cannot be registered here.*FontFace/);
  assert.deepEqual(registeredFonts(), []);
  assert.equal(unregisterFont('Test Sans'), false);
  // A document without web fonts is no better.
  globals.FontFace = FakeFace;
  globals.document = {};
  await assert.rejects(registerFont({ family: 'Test Sans', source: bytes }), /cannot be registered here/);
  assert.deepEqual(registeredFonts(), []);
  browser();
});

test('fonts: a registered font is loaded, added to the page and listed', async () => {
  reset();
  const before = fontsVersion();
  const loading = registerFont({ family: 'Test Sans', source: bytes });
  assert.deepEqual(registeredFonts(), [{ family: 'Test Sans', weight: 400, style: 'normal', as: [], prefer: false, status: 'loading' }]);
  assert.equal(fontsVersion(), before, 'nothing changes until the font is there');
  assert.deepEqual(await loading, { family: 'Test Sans', weight: 400, style: 'normal', as: [], prefer: false, status: 'loaded' });
  assert.equal(FakeFace.made.length, 1);
  assert.equal(FakeFace.made[0].family, 'Test Sans');
  assert.equal(FakeFace.made[0].source, bytes);
  assert.deepEqual(FakeFace.made[0].descriptors, { weight: '400', style: 'normal' });
  assert.ok(pageFonts.has(FakeFace.made[0]));
  assert.ok(fontsVersion() > before);

  await registerFont({ family: 'Test Sans', source: new ArrayBuffer(4), weight: 'bold', style: 'italic' });
  await registerFont({ family: 'Test Sans', source: 'https://example.com/fonts/a b.woff2', weight: 600 });
  assert.deepEqual(FakeFace.made[1].descriptors, { weight: '700', style: 'italic' });
  assert.equal(FakeFace.made[2].source, 'url("https://example.com/fonts/a b.woff2")');
  assert.deepEqual(FakeFace.made[2].descriptors, { weight: '600', style: 'normal' });
  assert.deepEqual(
    registeredFonts().map((entry) => entry.weight + ' ' + entry.style),
    ['400 normal', '700 italic', '600 normal'],
  );
  assert.equal(pageFonts.size, 3);

  // The same face again replaces the one that was there; "bold" and 700 are one face.
  await registerFont({ family: 'test sans', source: bytes, weight: 700, style: 'italic' });
  assert.equal(registeredFonts().length, 3);
  assert.equal(pageFonts.size, 3);
  assert.ok(!pageFonts.has(FakeFace.made[1]) && pageFonts.has(FakeFace.made[3]));

  assert.equal(unregisterFont('TEST SANS'), true);
  assert.deepEqual(registeredFonts(), []);
  assert.equal(pageFonts.size, 0);
  assert.equal(unregisterFont('Test Sans'), false);
});

test('fonts: a Blob is read before it is handed over', async () => {
  reset();
  await registerFont({ family: 'From Blob', source: new Blob([bytes]) });
  assert.ok(FakeFace.made[0].source instanceof ArrayBuffer);
  assert.deepEqual(Array.from(new Uint8Array(FakeFace.made[0].source as ArrayBuffer)), [0, 1, 0, 0]);
});

test('fonts: an address cannot leave the string it is put in', async () => {
  reset();
  await registerFont({ family: 'Quoted', source: 'https://example.com/a") , url("//evil/b\\\n' });
  assert.equal(FakeFace.made[0].source, 'url("https://example.com/a%22) , url(%22//evil/b%5C")');
});

test('fonts: names are sanitized like the names of a file', async () => {
  reset();
  const made = await registerFont({ family: ' Mine"; } body { x: y ', source: bytes, as: ['Bookman "Old" Style', '";', 'Bookman Old Style', 42 as unknown as string] });
  assert.equal(made.family, 'Mine body x y');
  assert.deepEqual(made.as, ['Bookman Old Style']);
  assert.equal(FakeFace.made[0].family, 'Mine body x y');
  await assert.rejects(registerFont({ family: '";()', source: bytes }), /family name/);
  await assert.rejects(registerFont({ family: 42 as unknown as string, source: bytes }), /family name/);
  await assert.rejects(registerFont(null as unknown as FontRegistration), /must be an object/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, weight: 'heavy' as unknown as number }), /weight/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, weight: 0 }), /weight/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, weight: NaN }), /weight/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, style: 'oblique' as unknown as 'italic' }), /style/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, as: {} as unknown as string }), /"as"/);
  await assert.rejects(registerFont({ family: 'A', source: bytes, as: new Array(33).fill('x') }), /more than 32/);
  await assert.rejects(registerFont({ family: 'A', source: '' }), /empty address/);
  await assert.rejects(registerFont({ family: 'A', source: 7 as unknown as string }), /source of a font must be/);
  assert.deepEqual(registeredFonts().map((entry) => entry.family), ['Mine body x y']);
});

test('fonts: a word of a family that starts with a digit is quoted where the browser asks for it', async () => {
  reset();
  await registerFont({ family: 'Code 39', source: bytes });
  assert.equal(FakeFace.made[0].family, 'Code 39');
  reset();
  FakeFace.refuseUnquotedDigits = true;
  const made = await registerFont({ family: 'Code 39', source: bytes });
  assert.equal(made.family, 'Code 39');
  assert.equal(FakeFace.made[0].family, '"Code 39"');
});

test('fonts: a font that cannot be loaded is not kept', async () => {
  reset();
  const before = fontsVersion();
  await assert.rejects(registerFont({ family: 'Broken', source: 'fail', as: 'Arial' }), /network error/);
  assert.deepEqual(registeredFonts(), []);
  assert.equal(pageFonts.size, 0);
  assert.equal(fontsVersion(), before);
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial']);
});

test('fonts: a font stands in for the families it is registered as, once it is loaded', async () => {
  reset();
  // Made before anything is registered: the lists it has built are built anew.
  const measurer = new CanvasTextMeasurer({ context: deviceContext({ 'Test Bonum': 1 }) });
  const bookman = font('Bookman Old Style', 20);
  const suffixed = font('Bookman Old Style Tur', 20);
  assert.equal(measurer.cssFont(bookman), '20px "Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(measurer.width(bookman, 'abcd'), 40);
  assert.equal(measurer.hasFont('Bookman Old Style'), false);

  const loading = registerFont({ family: 'Test Bonum', source: 'slow', as: ['bookman old style', 'Book Antiqua'] });
  assert.equal(measurer.cssFont(bookman), '20px "Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif', 'not while it loads');
  await release();
  await loading;

  assert.equal(measurer.cssFont(bookman), '20px "Bookman Old Style", "Test Bonum", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(measurer.cssFont(suffixed), '20px "Bookman Old Style Tur", "Bookman Old Style", "Test Bonum", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(measurer.width(bookman, 'abcd'), 80);
  assert.equal(measurer.width(suffixed, 'abcd'), 80);
  assert.equal(measurer.hasFont('Bookman Old Style Tur'), true);
  assert.equal(measurer.cssFamilies('Book Antiqua'), '"Book Antiqua", "Test Bonum", "TeX Gyre Pagella", "Palatino Linotype", "Palatino", "URW Palladio L", serif');
  // Other families are as they were, and the family itself needs no entry.
  assert.equal(measurer.cssFamilies('Arial'), '"Arial", "Liberation Sans", "Arimo", "Helvetica", sans-serif');
  assert.equal(measurer.cssFamilies('Test Bonum'), '"Test Bonum", sans-serif');

  // The device's own font wins where it has one...
  const own = new CanvasTextMeasurer({ context: deviceContext({ 'Bookman Old Style': 0.7, 'Test Bonum': 1 }) });
  assert.equal(own.width(bookman, 'abcd'), 56);
  // ...unless the font was registered to be preferred.
  await registerFont({ family: 'Test Bonum', source: bytes, as: 'Bookman Old Style', prefer: true });
  assert.equal(own.cssFont(suffixed), '20px "Test Bonum", "Bookman Old Style Tur", "Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(own.width(bookman, 'abcd'), 80);

  // Spreadsheets build their font strings from the same list.
  assert.equal(cssFont(cellFont('Bookman Old Style Tur')), `16px "Test Bonum", "Bookman Old Style Tur", "Bookman Old Style", ${FONT_FALLBACK}`);
  assert.deepEqual(fontFamilyChain('Arial Tur'), ['Arial Tur', 'Arial']);

  unregisterFont('Test Bonum');
  assert.equal(own.width(bookman, 'abcd'), 56);
  assert.equal(measurer.cssFont(bookman), '20px "Bookman Old Style", "TeX Gyre Bonum", "URW Bookman", "URW Bookman L", "Bookman", serif');
  assert.equal(cssFont(cellFont('Bookman Old Style')), `16px "Bookman Old Style", ${FONT_FALLBACK}`);
});

test('fonts: several faces of a family name it once, and prefer needs something to be preferred to', async () => {
  reset();
  await registerFont({ family: 'Two Faces', source: bytes, as: 'Garamond' });
  await registerFont({ family: 'Two Faces', source: bytes, weight: 'bold', as: 'Garamond' });
  assert.deepEqual(fontFamilyChain('Garamond'), ['Garamond', 'Two Faces']);
  const alone = await registerFont({ family: 'Alone', source: bytes, prefer: true });
  assert.equal(alone.prefer, false);
});

test('fonts: listeners hear of fonts that arrive and of fonts that go', async () => {
  reset();
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  const failing = onFontsChanged(() => {
    throw new Error('a view that fails');
  });
  const loading = registerFont({ family: 'Heard', source: 'slow' });
  assert.equal(heard, 0);
  await release();
  await loading;
  assert.equal(heard, 1);
  await assert.rejects(registerFont({ family: 'Unheard', source: 'fail' }));
  assert.equal(heard, 1);
  unregisterFont('Heard');
  assert.equal(heard, 2);
  unregisterFont('Heard');
  assert.equal(heard, 2);
  off();
  failing();
  await registerFont({ family: 'Later', source: bytes });
  assert.equal(heard, 2);
});

test('fonts: a font that is unregistered while it loads never arrives', async () => {
  reset();
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  const loading = registerFont({ family: 'Gone', source: 'slow', as: 'Arial' });
  assert.equal(unregisterFont('Gone'), true);
  const outcome = assert.rejects(loading, /unregistered or registered again/);
  await release();
  await outcome;
  assert.equal(pageFonts.size, 0);
  assert.equal(heard, 0);
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial']);
  off();
});

test('fonts: the number of registered faces is bounded', async () => {
  reset();
  for (let i = 0; i < 64; i++) await registerFont({ family: 'Face ' + i, source: bytes });
  await assert.rejects(registerFont({ family: 'One too many', source: bytes }), /No more than 64/);
  assert.equal(registeredFonts().length, 64);
  // Replacing one of them is still possible.
  await registerFont({ family: 'Face 3', source: bytes });
  assert.equal(registeredFonts().length, 64);
  reset();
});

test('fonts: a list of fonts is waited for, but not for ever', async () => {
  reset();
  assert.deepEqual(await settleFonts([], 1000), []);
  assert.deepEqual(await settleFonts([{ family: 'Good', source: bytes }], 1000), []);
  const failures = await settleFonts(
    [
      { family: 'Good Two', source: bytes },
      { family: 'Bad', source: 'fail' },
      { family: '";', source: bytes },
      { family: 'Slow', source: 'slow', as: 'Arial' },
    ],
    20,
  );
  assert.deepEqual(failures, [
    { family: 'Bad', reason: 'A network error occurred.', pending: false },
    { family: '', reason: 'A font needs a family name of letters, digits, spaces or hyphens', pending: false },
    { family: 'Slow', reason: 'The font did not load in time', pending: true },
  ]);
  // The slow one is used once it is there.
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial']);
  await release();
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial', 'Slow']);
  // A register function that throws is a failure like any other.
  const thrown = await settleFonts([{ family: 'X', source: bytes }], 20, () => {
    throw new Error('no fonts here');
  });
  assert.deepEqual(thrown, [{ family: 'X', reason: 'no fonts here', pending: false }]);
  reset();
});

test('fonts: the same face registered again takes over when it has loaded, with one notice; until then the old one works', async () => {
  reset();
  await registerFont({ family: 'Swap', source: bytes, weight: 'bold', as: 'Arial' });
  const old = FakeFace.made[0];
  let heard = 0;
  const off = onFontsChanged(() => heard++);
  const before = fontsVersion();
  const loading = registerFont({ family: 'SWAP', source: 'slow', weight: 700, as: 'Garamond' });
  await new Promise((resolve) => setImmediate(resolve));
  // Nothing changed yet: the font that works is still there, for the names it was registered as.
  assert.equal(heard, 0);
  assert.equal(fontsVersion(), before);
  assert.ok(pageFonts.has(old));
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial', 'Swap']);
  assert.deepEqual(fontFamilyChain('Garamond'), ['Garamond']);
  assert.deepEqual(
    registeredFonts().map((entry) => entry.family + ' ' + entry.status),
    ['Swap loaded', 'SWAP loading'],
  );
  await release();
  await loading;
  assert.equal(heard, 1, 'one notice for the swap');
  assert.ok(!pageFonts.has(old) && pageFonts.has(FakeFace.made[1]));
  assert.equal(pageFonts.size, 1);
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial']);
  assert.deepEqual(fontFamilyChain('Garamond'), ['Garamond', 'SWAP']);
  assert.deepEqual(registeredFonts(), [{ family: 'SWAP', weight: 700, style: 'normal', as: ['Garamond'], prefer: false, status: 'loaded' }]);

  // A replacement that cannot be loaded leaves the font that works alone.
  await assert.rejects(registerFont({ family: 'Swap', source: 'fail', weight: 'bold' }), /network error/);
  assert.equal(heard, 1);
  assert.ok(pageFonts.has(FakeFace.made[1]));
  assert.deepEqual(fontFamilyChain('Garamond'), ['Garamond', 'SWAP']);
  assert.equal(registeredFonts().length, 1);

  // Of two replacements under way the later one counts.
  const first = registerFont({ family: 'Swap', source: 'slow', weight: 'bold', as: 'Tahoma' });
  const firstOutcome = assert.rejects(first, /unregistered or registered again/);
  const second = registerFont({ family: 'Swap', source: 'slow', weight: 'bold', as: 'Verdana' });
  await release();
  await firstOutcome;
  await second;
  assert.equal(heard, 2);
  assert.equal(pageFonts.size, 1);
  assert.deepEqual(fontFamilyChain('Tahoma'), ['Tahoma']);
  assert.deepEqual(fontFamilyChain('Verdana'), ['Verdana', 'Swap']);

  // Unregistering takes the font and the replacement that is on its way.
  const late = registerFont({ family: 'Swap', source: 'slow', weight: 'bold' });
  const lateOutcome = assert.rejects(late, /unregistered or registered again/);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(unregisterFont('swap'), true);
  await release();
  await lateOutcome;
  assert.equal(heard, 3);
  assert.equal(pageFonts.size, 0);
  assert.deepEqual(registeredFonts(), []);
  off();
});

test('fonts: a font unregistered while its file is read is never made; one unregistered while it loads is never added', async () => {
  reset();
  // The Blob is still being read: no face is made at all.
  const early = registerFont({ family: 'Early', source: new Blob([bytes]) });
  const earlyOutcome = assert.rejects(early, /unregistered or registered again/);
  assert.equal(unregisterFont('Early'), true);
  await earlyOutcome;
  await release();
  assert.equal(FakeFace.made.length, 0);
  assert.equal(pageFonts.size, 0);

  // The face is made and loading: when it has loaded it is not added to the page.
  const late = registerFont({ family: 'Late', source: 'slow', as: 'Arial' });
  const lateOutcome = assert.rejects(late, /unregistered or registered again/);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(FakeFace.made.length, 1);
  assert.equal(FakeFace.waiting.length, 1);
  assert.equal(unregisterFont('Late'), true);
  await release();
  await lateOutcome;
  assert.equal(pageFonts.size, 0);
  assert.deepEqual(registeredFonts(), []);
  assert.deepEqual(fontFamilyChain('Arial'), ['Arial']);
});

test('fonts: a name the browser does not take as it is, is given in quotes', async () => {
  reset();
  FakeFace.refuseUnquoted = ['serif', 'inherit', 'Çok Güzel'];
  for (const family of ['serif', 'inherit', 'Çok Güzel']) {
    const made = await registerFont({ family, source: bytes });
    assert.equal(made.family, family);
    assert.equal(FakeFace.made[FakeFace.made.length - 1].family, '"' + family + '"');
  }
  // Another error than that of the name is not tried again.
  const count = FakeFace.made.length;
  await assert.rejects(registerFont({ family: 'Broken', source: 'fail' }), /network error/);
  assert.equal(FakeFace.made.length, count + 1);
});

test('fonts: the faces of a family stand in for a name once', async () => {
  reset();
  await registerFont({ family: 'Two Faces', source: bytes, as: 'Garamond' });
  await registerFont({ family: 'Two Faces', source: bytes, weight: 'bold', as: 'Garamond', prefer: true });
  await registerFont({ family: 'Two Faces', source: bytes, style: 'italic', as: 'Garamond', prefer: true });
  await registerFont({ family: 'Two Faces', source: bytes, weight: 300, as: 'Garamond' });
  assert.deepEqual(replacementFonts('Garamond', 'Garamond'), { first: ['Two Faces'], then: ['Two Faces'] });
  assert.deepEqual(replacementFonts('Garamond Tur', 'Garamond'), { first: ['Two Faces'], then: ['Two Faces'] });
  assert.equal(replacementFonts('Arial', 'Arial'), undefined);
});

/** A canvas whose widths follow `installed`, which a test changes as fonts arrive. */
function changingContext(installed: { [family: string]: number }): MeasureContext & { measured: number } {
  const inner = deviceContext(installed);
  const context = {
    font: '',
    measured: 0,
    measureText(text: string) {
      context.measured++;
      inner.font = context.font;
      return inner.measureText(text);
    },
  };
  return context;
}

test('fonts: text measured before a font of its own name arrived is measured again', async () => {
  reset();
  const installed: { [family: string]: number } = {};
  const context = changingContext(installed);
  const measurer = new CanvasTextMeasurer({ context });
  const mine = font('Mine Alone', 20);
  assert.equal(measurer.width(mine, 'abcd'), 40);
  assert.equal(measurer.width(mine, 'abcd'), 40);
  assert.deepEqual(measurer.metrics(mine), { ascent: 18, descent: 6, lineGap: 0 });
  const css = measurer.cssFont(mine);
  // The font arrives under its own name: the font string is the same, what it measures is not.
  installed['Mine Alone'] = 1;
  await registerFont({ family: 'Mine Alone', source: bytes });
  assert.equal(measurer.cssFont(mine), css);
  assert.equal(measurer.width(mine, 'abcd'), 80);
  unregisterFont('Mine Alone');
  delete installed['Mine Alone'];
  assert.equal(measurer.width(mine, 'abcd'), 40);
});

test('fonts: hasFont asks again once a font was registered', async () => {
  reset();
  const installed: { [family: string]: number } = {};
  const measurer = new CanvasTextMeasurer({ context: changingContext(installed) });
  assert.equal(measurer.hasFont('Mine Alone'), false);
  assert.equal(measurer.hasFont('Garamond'), false);
  installed['Mine Alone'] = 1;
  await registerFont({ family: 'Mine Alone', source: bytes, as: 'Garamond' });
  // Nothing else of the measurer is used in between.
  assert.equal(measurer.hasFont('Garamond'), true);
  assert.equal(measurer.hasFont('Mine Alone'), true);
  unregisterFont('Mine Alone');
  delete installed['Mine Alone'];
  assert.equal(measurer.hasFont('Garamond'), false);
});

test('grid text: widths measured before a font arrived are measured again, with no view to say so', async () => {
  reset();
  const installed: { [family: string]: number } = {};
  const context = changingContext(installed);
  const painter = new TextPainter(context as unknown as CanvasRenderingContext2D);
  const css = cssFont(cellFont('Grid Alone'));
  assert.equal(painter.measure(css, 'abcd'), 32);
  assert.equal(painter.measure(css, 'abcd'), 32);
  assert.equal(context.measured, 1, 'cached');
  installed['Grid Alone'] = 1;
  await registerFont({ family: 'Grid Alone', source: bytes });
  assert.equal(cssFont(cellFont('Grid Alone')), css);
  assert.equal(painter.measure(css, 'abcd'), 64);
  // Another painter, as another view has, shares the cache and is as fresh.
  const other = new TextPainter(context as unknown as CanvasRenderingContext2D);
  delete installed['Grid Alone'];
  unregisterFont('Grid Alone');
  assert.equal(other.measure(css, 'abcd'), 32);
});

test('fonts: the wait for a list of fonts ends its timer when the fonts are there first', async () => {
  reset();
  const timers: { handle: number; ms: number }[] = [];
  const cancelled: unknown[] = [];
  const timer = (_callback: () => void, ms: number) => {
    timers.push({ handle: timers.length + 100, ms });
    return timers.length + 99;
  };
  assert.deepEqual(await settleFonts([{ family: 'Quick', source: bytes }], 3000, registerFont, timer, (handle) => cancelled.push(handle)), []);
  assert.deepEqual(timers, [{ handle: 100, ms: 3000 }]);
  assert.deepEqual(cancelled, [100]);
  // When the timer ends the wait there is nothing to cancel.
  let fire: () => void = () => undefined;
  const pending = settleFonts(
    [{ family: 'Slow Two', source: 'slow' }],
    3000,
    registerFont,
    (callback) => {
      fire = callback;
      return 7;
    },
    (handle) => cancelled.push(handle),
  );
  await new Promise((resolve) => setImmediate(resolve));
  fire();
  assert.equal((await pending).length, 1);
  assert.deepEqual(cancelled, [100]);
  await release();
  reset();
});
