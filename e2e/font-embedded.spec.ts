import { readFileSync } from 'node:fs';
import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { docSection, fontStyles } from './fixtures/docx';
import { testFont } from './fixtures/font';
import { expect, test } from './helpers';

// The font the package brings: TeX Gyre Bonum, fetched from next to the
// script the first time a file draws text in Bookman Old Style on a device
// without that font.
//
// A machine that has Bookman Old Style installed would never fetch it, so
// most tests have the font stand in for a family no machine has (`as`); the
// tests that use the real name are skipped where the device has the font.

/** A family no device has; the embedded font is told to stand in for it. */
const ABSENT = 'Bym Absent Bookman';
const AS = ['Bookman Old Style', ABSENT];
/** 15 pt, in half points: 20 px. */
const SIZE = 30;
const PX = 20;
const TURKISH = 'ğşıİçöü';
const FILES = ['texgyrebonum-regular.otf', 'texgyrebonum-bold.otf', 'texgyrebonum-italic.otf', 'texgyrebonum-bolditalic.otf'];

const inFont = (text: string, family: string) => r(text, `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>`);

/** Ten letters in the family, ten in the family with a code page suffix, the Turkish letters, ten letters in Arial. */
function docIn(family: string): Promise<Uint8Array> {
  const body = p(inFont('nnnnnnnnnn', family)) + p(inFont('nnnnnnnnnn', family + ' Tur')) + p(inFont(TURKISH, family)) + p(inFont('cccccccccc', 'Arial'));
  return buildDocx({ styles: fontStyles(family, SIZE), body: body + docSection() });
}

function arialDocx(): Promise<Uint8Array> {
  return buildDocx({ styles: fontStyles('Arial', SIZE), body: p('Nothing here is set in the family.') + p(inFont('Nor this.', 'Times New Roman')) + docSection() });
}

/** The advance width of a character of a font file, in 1/1000 em: read from the file, not asked of a browser. */
function advance(file: string, char: string): number {
  const bytes = readFileSync('assets/fonts/tex-gyre-bonum/' + file);
  const tables: { [tag: string]: number } = {};
  for (let i = 0; i < bytes.readUInt16BE(4); i++) tables[bytes.toString('latin1', 12 + i * 16, 16 + i * 16)] = bytes.readUInt32BE(12 + i * 16 + 8);
  const cmap = tables['cmap'];
  let glyph = -1;
  for (let i = 0; i < bytes.readUInt16BE(cmap + 2) && glyph < 0; i++) {
    const record = cmap + 4 + i * 8;
    const table = cmap + bytes.readUInt32BE(record + 4);
    if (bytes.readUInt16BE(record) !== 3 || bytes.readUInt16BE(record + 2) !== 1 || bytes.readUInt16BE(table) !== 4) continue;
    const segments = bytes.readUInt16BE(table + 6) / 2;
    const ends = table + 14;
    const starts = ends + segments * 2 + 2;
    const deltas = starts + segments * 2;
    const offsets = deltas + segments * 2;
    const code = char.charCodeAt(0);
    for (let s = 0; s < segments; s++) {
      if (code > bytes.readUInt16BE(ends + s * 2) || code < bytes.readUInt16BE(starts + s * 2)) continue;
      const offset = bytes.readUInt16BE(offsets + s * 2);
      const delta = bytes.readUInt16BE(deltas + s * 2);
      glyph = offset === 0 ? (code + delta) & 0xffff : bytes.readUInt16BE(offsets + s * 2 + offset + (code - bytes.readUInt16BE(starts + s * 2)) * 2);
      break;
    }
  }
  if (glyph <= 0) throw new Error(`${file} has no glyph for "${char}"`);
  const metrics = bytes.readUInt16BE(tables['hhea'] + 34);
  return bytes.readUInt16BE(tables['hmtx'] + Math.min(glyph, metrics - 1) * 4);
}

/** The requests for the font files of the package a page makes, as addresses. */
function fontRequests(page: Page): string[] {
  const seen: string[] = [];
  page.on('request', (request) => {
    if (/tex-gyre-bonum|\.otf(\?|$)/.test(request.url())) seen.push(request.url());
  });
  return seen;
}

async function openDocPage(page: Page, as: string[] | null = AS): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  if (as) await page.evaluate((names) => (window as any).BymOffice.fonts.configure({ as: names }), as);
}

const mountDoc = (page: Page, bytes: Uint8Array) => page.evaluate((input) => (window as any).__mountDoc(input, {}), Array.from(bytes));
const laidOut = (page: Page) => page.waitForFunction(() => (window as any).__docView.layoutDone === true, undefined, { timeout: 60_000 });
const status = (page: Page, api = 'BymOffice') => page.evaluate((name) => (window as any)[name].fonts.embedded().status as string, api);
const loaded = (page: Page, api = 'BymOffice') => expect.poll(() => status(page, api), { timeout: 20_000 }).toBe('loaded');
/** Time for a request that must not be made to be made after all. */
const quiet = (page: Page) => page.evaluate(() => new Promise((resolve) => setTimeout(resolve, 400)));

/** Whether the device shows a family with a font of its own (not with the generic one behind it). */
function deviceHas(page: Page, family: string): Promise<boolean> {
  return page.evaluate((name) => {
    const context = document.createElement('canvas').getContext('2d')!;
    for (const generic of ['monospace', 'serif']) {
      context.font = `72px ${generic}`;
      const plain = context.measureText('mmmmmmmmmmlliWW@#01').width;
      context.font = `72px "${name}", ${generic}`;
      if (context.measureText('mmmmmmmmmmlliWW@#01').width !== plain) return true;
    }
    return false;
  }, family);
}

/**
 * The document of the view laid out anew with a measurer of its own: the
 * widths of its paragraphs by their text, and widths the canvas gives for
 * the same texts in the embedded font named directly.
 */
function measured(page: Page): Promise<{ letters: number[]; turkish: number; arial: number; direct: number; turkishDirect: number; turkishOverMono: number; turkishOverSerif: number; serif: number; check: boolean; line: { ascent: number; descent: number; lineGap: number }; bold: { ascent: number; descent: number; lineGap: number } }> {
  return page.evaluate(
    async ([turkish, px]) => {
      const api = (window as any).BymOffice;
      const doc = (window as any).__docView.document;
      const measurer = api.createCanvasMeasurer({ fonts: doc.fonts });
      const layout = new api.DocLayout(doc, measurer);
      layout.layout();
      const letters: number[] = [];
      let turkishWidth = 0;
      let arial = 0;
      let family = '';
      for (const item of layout.pages[0].items) {
        if (item.type !== 'text') continue;
        if (item.text === 'nnnnnnnnnn') {
          letters.push(item.width);
          family = item.font.family;
        } else if (item.text === 'cccccccccc') arial = item.width;
        else if (turkish.indexOf(item.text) >= 0) turkishWidth += item.width;
      }
      const context = document.createElement('canvas').getContext('2d')!;
      const width = (css: string, text: string) => {
        context.font = css;
        return context.measureText(text).width;
      };
      return {
        letters,
        turkish: turkishWidth,
        arial,
        direct: width(`${px}px "TeX Gyre Bonum"`, 'nnnnnnnnnn'),
        turkishDirect: width(`${px}px "TeX Gyre Bonum"`, turkish),
        // A letter the font lacks would come from the generic family, and the two generic families differ.
        turkishOverMono: width(`${px}px "TeX Gyre Bonum", monospace`, turkish),
        turkishOverSerif: width(`${px}px "TeX Gyre Bonum", serif`, turkish),
        serif: width(`${px}px serif`, 'nnnnnnnnnn'),
        check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean,
        line: measurer.metrics({ family, size: 1000, bold: false, italic: false }),
        bold: measurer.metrics({ family, size: 1000, bold: true, italic: false }),
      };
    },
    [TURKISH, PX] as [string, number],
  );
}

test.describe('the font of the package', () => {
  test('a document in a family the device lacks fetches the four faces once, and is drawn and measured with them', async ({ page }) => {
    const requests = fontRequests(page);
    await openDocPage(page);
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.embedded())).toEqual({
      family: 'TeX Gyre Bonum',
      as: AS,
      baseUrl: new URL('/dist/fonts/', page.url()).href,
      status: 'idle',
    });
    expect(requests).toEqual([]);

    await mountDoc(page, await docIn(ABSENT));
    await loaded(page);
    const origin = new URL(page.url()).origin;
    expect(requests.slice().sort()).toEqual(FILES.map((file) => `${origin}/dist/fonts/tex-gyre-bonum/${file}`).sort());
    await laidOut(page);

    const now = await measured(page);
    expect(now.check).toBe(true);
    // The view that was open laid its pages out again when the font arrived: nothing here asked it to.
    await expect
      .poll(() =>
        page.evaluate(() => {
          const item = (window as any).__docView.pages[0].items.find((entry: any) => entry.type === 'text' && entry.text === 'nnnnnnnnnn');
          return item ? (item.width as number) : -1;
        }),
      )
      .toBeCloseTo(now.direct, 3);
    // The two names, with and without the code page suffix, are the same font: the embedded one.
    expect(now.letters.length).toBe(2);
    expect(now.letters[0]).toBeCloseTo(now.direct, 3);
    expect(now.letters[1]).toBeCloseTo(now.direct, 3);
    // Ten letters as wide as the font file says they are, which is not what the generic font gives.
    const expected = (10 * advance(FILES[0], 'n') * PX) / 1000;
    expect(Math.abs(now.letters[0] - expected)).toBeLessThan(0.5);
    expect(Math.abs(now.serif - expected)).toBeGreaterThan(2);
    // The Turkish letters come from the same font: none falls back.
    expect(now.turkish).toBeCloseTo(now.turkishDirect, 3);
    expect(now.turkishOverMono).toBeCloseTo(now.turkishOverSerif, 3);
    let turkish = 0;
    for (const char of TURKISH) turkish += advance(FILES[0], char);
    expect(Math.abs(now.turkish - (turkish * PX) / 1000)).toBeLessThan(0.5);
    // The lines are as high as the font files say, the same in every browser.
    expect(now.line).toEqual({ ascent: 1069, descent: 268, lineGap: 0 });
    expect(now.bold).toEqual({ ascent: 1153, descent: 306, lineGap: 0 });
    // The registry of the application does not list it.
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.registeredFonts())).toEqual([]);
    expect(await page.evaluate(() => (window as any).__notices)).toEqual([]);

    // A second view of the same page fetches nothing.
    await mountDoc(page, await docIn(ABSENT));
    await laidOut(page);
    await quiet(page);
    expect(requests.length).toBe(4);
    expect((await measured(page)).letters[0]).toBeCloseTo(now.direct, 3);
  });

  test('Bookman Old Style and Bookman Old Style Tur are the names it stands in for by default', async ({ page }) => {
    const requests = fontRequests(page);
    await openDocPage(page, null);
    test.skip(await deviceHas(page, 'Bookman Old Style'), 'this device has Bookman Old Style: nothing is fetched for it');
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.embedded().as)).toEqual(['Bookman Old Style']);
    await mountDoc(page, await docIn('Bookman Old Style'));
    await loaded(page);
    await laidOut(page);
    expect(requests.length).toBe(4);
    const now = await measured(page);
    expect(now.letters.length).toBe(2);
    expect(now.letters[0]).toBeCloseTo(now.direct, 3);
    expect(now.letters[1]).toBeCloseTo(now.direct, 3);
    expect(now.turkish).toBeCloseTo(now.turkishDirect, 3);
    expect(now.line).toEqual({ ascent: 1069, descent: 268, lineGap: 0 });
  });

  test('a document that does not use the family fetches nothing', async ({ page }) => {
    const requests = fontRequests(page);
    await openDocPage(page);
    await mountDoc(page, await arialDocx());
    await laidOut(page);
    await quiet(page);
    expect(requests).toEqual([]);
    expect(await status(page)).toBe('idle');
  });

  test('mount with fonts: false or fontsUrl: null turns it off: nothing is fetched', async ({ page }) => {
    const requests = fontRequests(page);
    const bytes = Array.from(await docIn(ABSENT));
    for (const off of [{ fonts: false }, { fontsUrl: null }]) {
      await page.goto('/e2e/pages/mount.html');
      const result = await page.evaluate(
        async ([doc, names, options]) => {
          const api = (window as any).BymOffice;
          api.fonts.configure({ as: names });
          const notices: string[] = [];
          const mounted = await api.mount(document.getElementById('editor'), { file: new Uint8Array(doc), onNotice: (notice: { message: string }) => notices.push(notice.message), ...options });
          await new Promise<void>((resolve) => (function wait() { mounted.view.layoutDone ? resolve() : setTimeout(wait, 20); })());
          await new Promise((resolve) => setTimeout(resolve, 400));
          return { state: api.fonts.embedded(), notices, check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean };
        },
        [bytes, AS, off] as [number[], string[], object],
      );
      expect(result.state.status).toBe('off');
      expect(result.state.baseUrl).toBe(null);
      expect(result.notices).toEqual([]);
    }
    expect(requests).toEqual([]);
  });

  test('a wrong fontsUrl: the document opens in another font, and it is said once', async ({ page, pageProblems }) => {
    const requests = fontRequests(page);
    await page.goto('/e2e/pages/mount.html');
    const result = await page.evaluate(
      async ([doc, names]) => {
        const api = (window as any).BymOffice;
        api.fonts.configure({ as: names });
        const notices: { message: string; kind: string }[] = [];
        const open = async (host: HTMLElement) => {
          const mounted = await api.mount(host, { file: new Uint8Array(doc), locale: 'en', fontsUrl: '/no-such-folder/fonts', onNotice: (notice: { message: string; kind: string }) => notices.push({ message: notice.message, kind: notice.kind }) });
          await new Promise<void>((resolve) => (function wait() { mounted.view.layoutDone ? resolve() : setTimeout(wait, 20); })());
          return mounted;
        };
        const first = await open(document.getElementById('editor')!);
        await new Promise<void>((resolve) => (function wait() { api.fonts.embedded().status === 'failed' ? resolve() : setTimeout(wait, 20); })());
        await new Promise((resolve) => setTimeout(resolve, 200));
        // A second view of the page: nothing is tried or said again.
        const other = document.createElement('div');
        other.style.height = '300px';
        document.body.appendChild(other);
        const second = await open(other);
        await new Promise((resolve) => setTimeout(resolve, 400));
        const measurer = api.createCanvasMeasurer({ fonts: first.document.fonts });
        return {
          notices,
          state: api.fonts.embedded(),
          pages: [first.view.pageCount, second.view.pageCount],
          width: measurer.width({ family: names[1], size: 20, bold: false, italic: false }, 'nnnnnnnnnn') as number,
          check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean,
          canvases: document.querySelectorAll('#editor canvas').length,
        };
      },
      [Array.from(await docIn(ABSENT)), AS] as [number[], string[]],
    );
    expect(result.notices.length).toBe(1);
    expect(result.notices[0].kind).toBe('info');
    expect(result.notices[0].message).toContain(`"${ABSENT}"`);
    expect(result.state.status).toBe('failed');
    expect(result.state.baseUrl).toBe(new URL('/no-such-folder/fonts/', page.url()).href);
    expect(result.pages).toEqual([1, 1]);
    expect(result.width).toBeGreaterThan(0);
    expect(result.canvases).toBeGreaterThan(0);
    // Four files were asked for, at the address that was given, and never again.
    expect(requests.slice().sort()).toEqual(FILES.map((file) => new URL('/no-such-folder/fonts/tex-gyre-bonum/' + file, page.url()).href).sort());
    // The browsers report the files they could not get; nothing else may be on the console.
    for (let i = pageProblems.length - 1; i >= 0; i--) {
      if (/no-such-folder/.test(pageProblems[i]) || (/downloadable font/.test(pageProblems[i]) && /TeX Gyre Bonum/.test(pageProblems[i]))) pageProblems.splice(i, 1);
    }
  });

  test('an ES module is told where the fonts are, in one line of the application', async ({ page }) => {
    const requests = fontRequests(page);
    await page.goto('/e2e/pages/esm-fonts.html');
    await page.waitForFunction(() => (window as any).__bym !== undefined);
    expect(await page.evaluate(() => (window as any).__bym.fonts.embedded().baseUrl)).toBe(new URL('/dist/fonts/', page.url()).href);
    const result = await page.evaluate(
      async ([doc, names]) => {
        const api = (window as any).__bym;
        api.fonts.configure({ as: names });
        const host = document.createElement('div');
        host.style.height = '400px';
        document.body.appendChild(host);
        const mounted = await api.mount(host, { file: new Uint8Array(doc) });
        await new Promise<void>((resolve) => (function wait() { api.fonts.embedded().status === 'loaded' ? resolve() : setTimeout(wait, 20); })());
        const context = document.createElement('canvas').getContext('2d')!;
        context.font = '20px "TeX Gyre Bonum"';
        const direct = context.measureText('nnnnnnnnnn').width;
        const measurer = api.createCanvasMeasurer({ fonts: mounted.document.fonts });
        return { direct, width: measurer.width({ family: names[1], size: 20, bold: false, italic: false }, 'nnnnnnnnnn') as number, check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean };
      },
      [Array.from(await docIn(ABSENT)), AS] as [number[], string[]],
    );
    expect(result.check).toBe(true);
    expect(result.width).toBeCloseTo(result.direct, 3);
    expect(Math.abs(result.width - (10 * advance(FILES[0], 'n') * PX) / 1000)).toBeLessThan(0.5);
    const origin = new URL(page.url()).origin;
    expect(requests.slice().sort()).toEqual(FILES.map((file) => `${origin}/dist/fonts/tex-gyre-bonum/${file}`).sort());
  });

  test('an ES module that is not told fetches nothing, opens the document and says once what to do', async ({ page }) => {
    const requests = fontRequests(page);
    await page.goto('/e2e/pages/esm.html');
    await page.waitForFunction(() => (window as any).__bym !== undefined);
    expect(await page.evaluate(() => (window as any).__bym.fonts.embedded())).toEqual({ family: 'TeX Gyre Bonum', as: ['Bookman Old Style'], baseUrl: null, status: 'idle' });
    const result = await page.evaluate(
      async ([doc, names]) => {
        const api = (window as any).__bym;
        api.fonts.configure({ as: names });
        const notices: { message: string; kind: string }[] = [];
        const open = async () => {
          const host = document.createElement('div');
          host.style.height = '300px';
          document.body.appendChild(host);
          const mounted = await api.mount(host, { file: new Uint8Array(doc), locale: 'en', onNotice: (notice: { message: string; kind: string }) => notices.push({ message: notice.message, kind: notice.kind }) });
          await new Promise<void>((resolve) => (function wait() { mounted.view.layoutDone ? resolve() : setTimeout(wait, 20); })());
          return mounted;
        };
        const first = await open();
        await new Promise<void>((resolve) => (function wait() { api.fonts.embedded().status === 'failed' ? resolve() : setTimeout(wait, 20); })());
        const second = await open();
        await new Promise((resolve) => setTimeout(resolve, 400));
        return { notices, pages: [first.view.pageCount, second.view.pageCount], state: api.fonts.embedded(), toolbar: Array.from(document.querySelectorAll('.bym-tb-font option')).map((option) => option.textContent) };
      },
      [Array.from(await docIn(ABSENT)), AS] as [number[], string[]],
    );
    expect(requests).toEqual([]);
    expect(result.pages).toEqual([1, 1]);
    expect(result.state.status).toBe('failed');
    expect(result.notices.length).toBe(1);
    expect(result.notices[0].kind).toBe('info');
    expect(result.notices[0].message).toContain(`"${ABSENT}"`);
    expect(result.notices[0].message).toContain('fontsUrl');
    expect(result.notices[0].message).toContain('fonts: false');
    // A name that cannot be shown is not offered.
    expect(result.toolbar).not.toContain('Bookman Old Style');
  });

  test('turned off after it was loaded it leaves the page and the document is laid out without it; turned on it is back without a request', async ({ page }) => {
    const requests = fontRequests(page);
    await openDocPage(page);
    await mountDoc(page, await docIn(ABSENT));
    await loaded(page);
    await laidOut(page);
    const viewWidth = () =>
      page.evaluate(() => {
        const item = (window as any).__docView.pages[0].items.find((entry: any) => entry.type === 'text' && entry.text === 'nnnnnnnnnn');
        return item ? (item.width as number) : -1;
      });
    const direct = (await measured(page)).direct;
    await expect.poll(viewWidth).toBeCloseTo(direct, 3);
    const lists = () =>
      page.evaluate((family) => {
        const api = (window as any).BymOffice;
        const doc = api.createCanvasMeasurer().cssFamilies(family) as string;
        return { doc: doc.indexOf('TeX Gyre Bonum') >= 0, check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean, faces: Array.from((document as any).fonts).filter((face: any) => /TeX Gyre Bonum/.test(face.family)).length };
      }, ABSENT);
    expect(await lists()).toEqual({ doc: true, check: true, faces: 4 });

    await page.evaluate(() => (window as any).BymOffice.fonts.configure({ baseUrl: null }));
    expect(await status(page)).toBe('off');
    expect((await lists()).doc).toBe(false);
    expect((await lists()).faces).toBe(0);
    // The open view laid its text out again, in the font the device has.
    const serif = (await measured(page)).serif;
    await expect.poll(viewWidth).not.toBeCloseTo(direct, 1);
    expect(Math.abs((await viewWidth()) - direct)).toBeGreaterThan(1);
    expect(serif).toBeGreaterThan(0);

    await page.evaluate(() => (window as any).BymOffice.fonts.configure({ baseUrl: undefined }));
    await loaded(page);
    await expect.poll(viewWidth).toBeCloseTo(direct, 3);
    expect(await lists()).toEqual({ doc: true, check: true, faces: 4 });
    expect(requests.length).toBe(4);
  });

  test('a script loaded with async or defer, added by a script or given a query still knows where it came from', async ({ page }) => {
    for (const address of ['/e2e/pages/script-async.html', '/e2e/pages/script-defer.html']) {
      await page.goto(address);
      await page.waitForFunction(() => (window as any).BymOffice !== undefined);
      expect(await page.evaluate(() => (window as any).BymOffice.fonts.embedded().baseUrl), address).toBe(new URL('/dist/fonts/', page.url()).href);
    }
    for (const src of ['/dist/bymoffice.min.js', '../../dist/bymoffice.min.js?v=1.2.3#top']) {
      await page.goto('/e2e/pages/blank.html');
      const base = await page.evaluate(async (address) => {
        await new Promise<void>((resolve, reject) => {
          const script = document.createElement('script');
          script.src = address;
          script.onload = () => resolve();
          script.onerror = () => reject(new Error('the script did not load'));
          document.head.appendChild(script);
        });
        return (window as any).BymOffice.fonts.embedded().baseUrl as string;
      }, src);
      expect(base, src).toBe(new URL('/dist/fonts/', page.url()).href);
    }
  });

  test('the demo page and the example page find the fonts next to the script they load', async ({ page }) => {
    for (const address of ['/demo/index.html', '/examples/html/index.html']) {
      await page.goto(address);
      await page.waitForFunction(() => (window as any).BymOffice !== undefined);
      const found = await page.evaluate(async () => {
        const base = (window as any).BymOffice.fonts.embedded().baseUrl as string;
        const response = await fetch(base + 'tex-gyre-bonum/texgyrebonum-regular.otf');
        return { base, ok: response.ok, bytes: (await response.arrayBuffer()).byteLength };
      });
      expect(found.base, address).toBe(new URL('/dist/fonts/', page.url()).href);
      expect(found.ok, address).toBe(true);
      expect(found.bytes, address).toBe(readFileSync('assets/fonts/tex-gyre-bonum/texgyrebonum-regular.otf').length);
    }
    // The test page shows the state of the font and turns it off and on.
    await page.goto('/demo/index.html');
    await expect(page.locator('#font-embedded-state')).toContainText('TeX Gyre Bonum: henüz gerekmedi');
    await expect(page.locator('#font-embedded-state')).toContainText('/dist/fonts/');
    await page.locator('#font-embedded').uncheck();
    await expect(page.locator('#font-embedded-state')).toHaveText('TeX Gyre Bonum: kapalı');
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.embedded().status)).toBe('off');
    await page.locator('#font-embedded').check();
    await expect(page.locator('#font-embedded-state')).toContainText('henüz gerekmedi');
  });

  test('a font the application registered for the family keeps the embedded one from being fetched', async ({ page }) => {
    const requests = fontRequests(page);
    await openDocPage(page);
    await page.evaluate(async ([font, family]) => {
      await (window as any).BymOffice.fonts.registerFont({ family: 'Bym Own Bookman', source: new Uint8Array(font as number[]), as: [family] });
    }, [Array.from(testFont()), ABSENT] as [number[], string]);
    await mountDoc(page, await docIn(ABSENT));
    await laidOut(page);
    await quiet(page);
    expect(requests).toEqual([]);
    expect(await status(page)).toBe('idle');
    // The application's font is what is shown: ten characters of 20 px each.
    expect((await measured(page)).letters[0]).toBeCloseTo(10 * PX, 3);
  });

  test('a spreadsheet cell in the family fetches it, and the cell is painted with it', async ({ page }) => {
    const requests = fontRequests(page);
    await page.goto('/e2e/pages/mount.html');
    const result = await page.evaluate(
      async ([names]) => {
        const api = (window as any).BymOffice;
        api.fonts.configure({ as: names });
        const host = document.getElementById('editor')!;
        const mounted = await api.mount(host, { file: await api.createBlank('xlsx'), toolbar: false });
        const view = mounted.view;
        const frames = () => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
        const dark = () => {
          let count = 0;
          host.querySelectorAll('canvas').forEach((canvas) => {
            const data = (canvas as HTMLCanvasElement).getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
            for (let i = 0; i < data.length; i += 4) if (data[i + 3] > 200 && data[i] < 110 && data[i + 1] < 110 && data[i + 2] < 110) count++;
          });
          return count;
        };
        view.setCellInput(1, 1, 'HHHHHHHHHH');
        view.select({ top: 8, left: 0, bottom: 8, right: 0 });
        await frames();
        await new Promise((resolve) => setTimeout(resolve, 300));
        const idle = api.fonts.embedded().status;
        view.select({ top: 1, left: 1, bottom: 1, right: 1 });
        view.format({ fontName: names[1] + ' Tur', fontSize: 24, bold: true });
        view.select({ top: 8, left: 0, bottom: 8, right: 0 });
        await frames();
        const before = dark();
        await new Promise<void>((resolve) => (function wait() { api.fonts.embedded().status === 'loaded' ? resolve() : setTimeout(wait, 20); })());
        await frames();
        await frames();
        const after = dark();
        return { idle, before, after, check: (document as any).fonts.check('bold 16px "TeX Gyre Bonum"') as boolean };
      },
      [AS] as [string[]],
    );
    expect(result.idle).toBe('idle');
    expect(result.check).toBe(true);
    expect(requests.length).toBe(4);
    // The cell was painted anew when the font arrived: other glyphs, other pixels.
    expect(result.after).not.toBe(result.before);
  });

  test('the font list of the document toolbar offers Bookman Old Style, and not when the embedded font is off', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    const bytes = Array.from(await arialDocx());
    await page.evaluate(async (doc) => {
      (window as any).__mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
    }, bytes);
    const options = () => page.locator('#editor .bym-tb-font option').allTextContents();
    expect(await options()).toContain('Bookman Old Style');
    expect(await options()).not.toContain('TeX Gyre Bonum');
    await page.evaluate(() => (window as any).BymOffice.fonts.configure({ baseUrl: null }));
    await expect.poll(options).not.toContain('Bookman Old Style');
    await page.evaluate(() => (window as any).BymOffice.fonts.configure({ baseUrl: undefined }));
    await expect.poll(options).toContain('Bookman Old Style');
  });

  test('choosing Bookman Old Style in the toolbar writes that name into the document and draws it with the embedded font', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    test.skip(await deviceHas(page, 'Bookman Old Style'), 'this device has Bookman Old Style: nothing is fetched for it');
    const requests = fontRequests(page);
    const bytes = Array.from(await arialDocx());
    await page.evaluate(async (doc) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
      (window as any).__mounted = mounted;
      mounted.view.setSelection({ anchor: { story: 'body', path: [0], offset: 0 }, focus: { story: 'body', path: [0], offset: 7 } });
    }, bytes);
    await page.locator('#editor .bym-tb-font').selectOption('Bookman Old Style');
    await loaded(page);
    expect(requests.length).toBe(4);
    const saved = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const file = await api.open(await (window as any).__mounted.save());
      const xml = new TextDecoder().decode(await file.package.readPart('/word/document.xml'));
      return { bookman: xml.indexOf('Bookman Old Style') >= 0, bonum: xml.indexOf('TeX Gyre Bonum') >= 0 };
    });
    // The file names the font Word has; the stand-in is a matter of this device.
    expect(saved).toEqual({ bookman: true, bonum: false });
  });

  test('the WebView bridge page carries the font inside it: no request, under its own policy', async ({ page }) => {
    const all: string[] = [];
    page.on('request', (request) => {
      if (!/^(data|blob):/.test(request.url())) all.push(request.url());
    });
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await page.goto('/dist/bymoffice.webview.html');
    await page.evaluate((names) => (window as any).BymOffice.fonts.configure({ as: names }), AS);
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.embedded())).toEqual({ family: 'TeX Gyre Bonum', as: AS, baseUrl: null, status: 'idle' });
    const send = (message: unknown) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
    const answered = async (id: string) => {
      let found: any;
      await expect.poll(async () => (found = await page.evaluate((wanted) => (window as any).__messages.find((message: any) => message.id === wanted), id)) !== undefined).toBe(true);
      return found;
    };

    // A file that does not use the family: the font stays where it is.
    await send({ type: 'load', id: 'plain', data: Buffer.from(await arialDocx()).toString('base64') });
    expect((await answered('plain')).type).toBe('loaded');
    await quiet(page);
    expect(await status(page)).toBe('idle');

    await send({ type: 'load', id: 'bookman', data: Buffer.from(await docIn(ABSENT)).toString('base64') });
    expect((await answered('bookman')).type).toBe('loaded');
    await loaded(page);
    const result = await page.evaluate(
      ([family]) => {
        const api = (window as any).BymOffice;
        const context = document.createElement('canvas').getContext('2d')!;
        const width = (css: string, text: string) => {
          context.font = css;
          return context.measureText(text).width;
        };
        const measurer = api.createCanvasMeasurer();
        const face = (bold: boolean, italic: boolean) => measurer.width({ family, size: 20, bold, italic }, 'nnnnnnnnnn') as number;
        return {
          check: (document as any).fonts.check('16px "TeX Gyre Bonum"') as boolean,
          widths: [face(false, false), face(true, false), face(false, true), face(true, true)],
          direct: [width('20px "TeX Gyre Bonum"', 'nnnnnnnnnn'), width('bold 20px "TeX Gyre Bonum"', 'nnnnnnnnnn'), width('italic 20px "TeX Gyre Bonum"', 'nnnnnnnnnn'), width('italic bold 20px "TeX Gyre Bonum"', 'nnnnnnnnnn')],
          notices: (window as any).__messages.filter((message: any) => message.type === 'notice' || message.type === 'error'),
        };
      },
      [ABSENT] as [string],
    );
    expect(result.check).toBe(true);
    // Each of the four faces is the file of that face: as wide as that file says.
    for (let i = 0; i < 4; i++) {
      expect(result.widths[i]).toBeCloseTo(result.direct[i], 3);
      expect(Math.abs(result.widths[i] - (10 * advance(FILES[i], 'n') * PX) / 1000), FILES[i]).toBeLessThan(0.5);
    }
    expect(result.notices).toEqual([]);
    // The page itself is all that was ever requested.
    expect(all).toEqual([page.url()]);
  });
});
