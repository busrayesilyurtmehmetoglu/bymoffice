import type { Page } from '@playwright/test';
import { p, r } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { docSection, fontStyles } from './fixtures/docx';
import { testFont } from './fixtures/font';
import { expect, test } from './helpers';

// Fonts an application registers, and font names with a code page suffix.
//
// The test font is made in code (fixtures/font.ts): every character is as
// wide as the font size, so a text of n characters at 20 px is n * 20 px wide.

const FONT = Array.from(testFont());
/** 15 pt, in half points: 20 px. */
const SIZE = 30;
const PX = 20;

const inFont = (text: string, family: string) => r(text, `<w:rFonts w:ascii="${family}" w:hAnsi="${family}" w:cs="${family}"/>`);

/**
 * A document in Bookman Old Style, which no test browser has, on pages of
 * 600 x 400 px with 520 x 320 px of text: 40 paragraphs of 30 letters each.
 * In the test font a paragraph is 600 px wide and takes two lines.
 */
function bookmanDocx(): Promise<Uint8Array> {
  let body = p(inFont('nnnnnnnnnn', 'Bookman Old Style')) + p(inFont('nnnnnnnnnn', 'Bookman Old Style Tur')) + p(inFont('cccccccccc', 'Arial'));
  for (let i = 0; i < 40; i++) body += p('x'.repeat(30));
  return buildDocx({ styles: fontStyles('Bookman Old Style', SIZE), body: body + docSection() });
}

/**
 * These tests are about the fonts an application registers. The font the
 * package brings for Bookman Old Style would arrive on its own in the middle
 * of them: it is turned off (it has its own tests in font-embedded.spec.ts).
 */
const withoutEmbeddedFont = (page: Page) => page.evaluate(() => (window as any).BymOffice.fonts.configure({ baseUrl: null }));

async function mountDoc(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await withoutEmbeddedFont(page);
  await page.evaluate(([input, opts]) => (window as any).__mountDoc(input, opts), [Array.from(bytes), {}] as [number[], object]);
}

const laidOut = (page: Page) => page.waitForFunction(() => (window as any).__docView.layoutDone === true, undefined, { timeout: 60_000 });
const pageCount = (page: Page) => page.evaluate(() => (window as any).__docView.pageCount as number);

/**
 * The document of the view laid out anew with a measurer of its own: the
 * widths of the three short paragraphs (Bookman Old Style, the same with a
 * code page suffix, Arial) and the number of pages.
 */
function measured(page: Page): Promise<{ widths: number[]; pages: number }> {
  return page.evaluate(async () => {
    const api = (window as any).BymOffice;
    const doc = (window as any).__docView.document;
    const layout = new api.DocLayout(doc, api.createCanvasMeasurer({ fonts: doc.fonts }));
    layout.layout();
    const widths: number[] = [];
    for (const item of layout.pages[0].items) if (item.type === 'text' && /^(n{10}|c{10})$/.test(item.text)) widths.push(item.width);
    return { widths, pages: layout.pages.length as number };
  });
}

/** Firefox itself reports a font file it refuses on the console; that report is expected, nothing else is. */
function forgetRefusedFont(problems: string[], family: string): void {
  for (let i = problems.length - 1; i >= 0; i--) {
    if (problems[i].indexOf('downloadable font: ') >= 0 && problems[i].indexOf(`"${family}"`) >= 0) problems.splice(i, 1);
  }
}

test.describe('fonts', () => {
  test('a registered font stands in for a family the device lacks, and the open document is laid out again', async ({ page }) => {
    await mountDoc(page, await bookmanDocx());
    await laidOut(page);
    const before = await measured(page);
    const pagesBefore = await pageCount(page);
    expect(pagesBefore).toBe(before.pages);
    // Without the font both names fall back to the same font: the suffixed name is its family.
    expect(before.widths.length).toBe(3);
    expect(before.widths[0]).toBeGreaterThan(0);
    expect(before.widths[1]).toBeCloseTo(before.widths[0], 3);
    expect(before.widths[0]).not.toBeCloseTo(10 * PX, 0);

    const registered = await page.evaluate(async (bytes) => {
      const api = (window as any).BymOffice;
      const made = await api.fonts.registerFont({ family: 'Bym Test Bonum', source: new Uint8Array(bytes), as: ['Bookman Old Style'] });
      return { made, same: api.registerFont === api.fonts.registerFont, list: api.fonts.registeredFonts() };
    }, FONT);
    expect(registered.made).toEqual({ family: 'Bym Test Bonum', weight: 400, style: 'normal', as: ['Bookman Old Style'], prefer: false, status: 'loaded' });
    expect(registered.same).toBe(true);
    expect(registered.list).toEqual([registered.made]);

    // The view heard of the font: nothing here asks it to lay out again.
    await expect.poll(() => pageCount(page)).not.toBe(pagesBefore);
    await laidOut(page);
    const after = await measured(page);
    // Ten characters of 20 px each, whatever the browser.
    expect(after.widths[0]).toBeCloseTo(10 * PX, 3);
    expect(after.widths[1]).toBeCloseTo(10 * PX, 3);
    // Arial is not touched.
    expect(after.widths[2]).toBeCloseTo(before.widths[2], 3);
    // Every paragraph of 30 letters takes two lines now, of 20 px each: 3 + 80 lines, 16 to a page.
    expect(after.pages).toBe(6);
    expect(await pageCount(page)).toBe(6);
    expect(after.pages).toBeGreaterThan(before.pages);

    // Taking the font out again brings the first layout back.
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.unregisterFont('Bym Test Bonum'))).toBe(true);
    await expect.poll(() => pageCount(page)).toBe(pagesBefore);
    await laidOut(page);
    expect((await measured(page)).widths[0]).toBeCloseTo(before.widths[0], 3);
    expect(await page.evaluate(() => (window as any).BymOffice.fonts.registeredFonts())).toEqual([]);
  });

  test('an editor keeps its caret when a font arrives', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const bytes = Array.from(await bookmanDocx());
    const result = await page.evaluate(
      async ([doc, font]) => {
        const api = (window as any).BymOffice;
        const mounted = await api.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
        const view = mounted.view;
        const done = () => new Promise<void>((resolve) => (function wait() { view.layoutDone ? resolve() : setTimeout(wait, 20); })());
        await done();
        const pages = view.pageCount;
        view.setSelection({ story: 'body', path: [1], offset: 4 });
        // Counts the layouts from the start that the fonts cause.
        let relayouts = 0;
        const relayout = view.relayout.bind(view);
        view.relayout = () => {
          relayouts++;
          relayout();
        };
        await api.fonts.registerFont({ family: 'Bym Test Bonum', source: new Uint8Array(font), as: 'Bookman Old Style' });
        // The time a view that is gone is given below to do nothing in: a view that is there lays out within it.
        await new Promise((resolve) => setTimeout(resolve, 120));
        const whileOpen = relayouts;
        await new Promise<void>((resolve) => (function wait() { view.pageCount !== pages ? resolve() : setTimeout(wait, 20); })());
        await done();
        view.insertText('Z');
        const text = api.paragraphText(mounted.document.body.blocks[1]);
        const out = { pages, after: view.pageCount, caret: view.getSelection().focus, text, whileOpen, afterDestroy: 0 };
        mounted.destroy();
        // A view that is gone hears nothing: taking the font out must not throw or lay anything out.
        api.fonts.unregisterFont('Bym Test Bonum');
        await new Promise((resolve) => setTimeout(resolve, 120));
        out.afterDestroy = relayouts - whileOpen;
        return out;
      },
      [bytes, FONT] as [number[], number[]],
    );
    expect(result.whileOpen).toBe(1);
    expect(result.afterDestroy).toBe(0);
    expect(result.after).toBe(6);
    expect(result.after).toBeGreaterThan(result.pages);
    // The caret was at the fifth letter and still is: the letter typed lands there.
    expect(result.text).toBe('nnnnZnnnnnn');
    expect(result.caret).toMatchObject({ story: 'body', path: [1], offset: 5 });
  });

  test('a font that arrives while an input method composes is laid out when the composition has ended', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const bytes = Array.from(await bookmanDocx());
    const result = await page.evaluate(
      async ([doc, font]) => {
        const api = (window as any).BymOffice;
        const mounted = await api.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
        const view = mounted.view;
        const until = (test: () => boolean) => new Promise<void>((resolve) => (function wait() { test() ? resolve() : setTimeout(wait, 20); })());
        const pause = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
        await until(() => view.layoutDone);
        let relayouts = 0;
        const relayout = view.relayout.bind(view);
        view.relayout = () => {
          relayouts++;
          relayout();
        };
        const pages = view.pageCount;
        view.setSelection({ story: 'body', path: [1], offset: 4 });
        view.focus();
        await pause(50);
        const field = document.querySelector('#editor .bym-doc-input') as HTMLTextAreaElement;
        const fire = (type: string, data: string) => field.dispatchEvent(new InputEvent(type, { inputType: 'insertCompositionText', data, isComposing: true, bubbles: true, cancelable: false }));
        field.dispatchEvent(new CompositionEvent('compositionstart', { data: '', bubbles: true }));
        const start = field.selectionStart;
        let length = field.selectionEnd - field.selectionStart;
        const put = (text: string) => {
          fire('beforeinput', text);
          field.dispatchEvent(new CompositionEvent('compositionupdate', { data: text, bubbles: true }));
          field.setRangeText(text, start, start + length, 'end');
          length = text.length;
          fire('input', text);
        };
        const text = () => api.paragraphText(mounted.document.body.blocks[1]) as string;
        put('に');
        await api.registerFont({ family: 'Bym Test Bonum', source: new Uint8Array(font), as: 'Bookman Old Style' });
        // Far longer than a view takes to lay out when nothing is being composed.
        await pause(200);
        const during = { relayouts, pages: view.pageCount, text: text() };
        // The composition goes on as if nothing had happened.
        put('にほ');
        const composed = text();
        field.dispatchEvent(new CompositionEvent('compositionend', { data: 'にほ', bubbles: true }));
        await until(() => relayouts > 0);
        await until(() => view.layoutDone);
        const after = { relayouts, pages: view.pageCount, text: text(), caret: view.getSelection().focus };
        view.insertText('Z');
        const typed = text();
        mounted.destroy();
        return { pages, during, composed, after, typed };
      },
      [bytes, FONT] as [number[], number[]],
    );
    expect(result.during).toEqual({ relayouts: 0, pages: result.pages, text: 'nnnnにnnnnnn' });
    expect(result.composed).toBe('nnnnにほnnnnnn');
    expect(result.after.relayouts).toBe(1);
    expect(result.after.pages).toBe(6);
    expect(result.after.text).toBe('nnnnにほnnnnnn');
    expect(result.after.caret).toMatchObject({ story: 'body', path: [1], offset: 6 });
    expect(result.typed).toBe('nnnnにほZnnnnnn');
  });

  test('a font that arrives while a selection is dragged is laid out when the pointer is released', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const bytes = Array.from(await bookmanDocx());
    await page.evaluate(async (doc) => {
      const api = (window as any).BymOffice;
      const mounted = await api.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
      const view = mounted.view;
      await new Promise<void>((resolve) => (function wait() { view.layoutDone ? resolve() : setTimeout(wait, 20); })());
      (window as any).__mounted = mounted;
      (window as any).__relayouts = 0;
      const relayout = view.relayout.bind(view);
      view.relayout = () => {
        (window as any).__relayouts++;
        relayout();
      };
    }, bytes);
    const sheet = (await page.locator('#editor .bym-doc-page[data-page="1"]').boundingBox())!;
    const pages = await page.evaluate(() => (window as any).__mounted.view.pageCount as number);
    // Pressed in the text of the first lines and held.
    await page.mouse.move(sheet.x + sheet.width * 0.12, sheet.y + sheet.height * 0.14);
    await page.mouse.down();
    await page.mouse.move(sheet.x + sheet.width * 0.3, sheet.y + sheet.height * 0.22, { steps: 3 });
    await page.evaluate(async (font) => {
      await (window as any).BymOffice.registerFont({ family: 'Bym Test Bonum', source: new Uint8Array(font), as: 'Bookman Old Style' });
      await new Promise((resolve) => setTimeout(resolve, 200));
    }, FONT);
    const during = await page.evaluate(() => ({ relayouts: (window as any).__relayouts as number, pages: (window as any).__mounted.view.pageCount as number, selected: (window as any).__mounted.view.getSelectedText() as string }));
    expect(during.relayouts).toBe(0);
    expect(during.pages).toBe(pages);
    expect(during.selected.length).toBeGreaterThan(0);
    await page.mouse.up();
    await expect.poll(() => page.evaluate(() => (window as any).__relayouts as number)).toBe(1);
    await page.waitForFunction(() => (window as any).__mounted.view.layoutDone === true);
    expect(await page.evaluate(() => (window as any).__mounted.view.pageCount as number)).toBe(6);
    // What was selected still is.
    expect(await page.evaluate(() => (window as any).__mounted.view.getSelectedText() as string)).toBe(during.selected);
  });

  test('mount registers its fonts before the document is laid out, and says which did not load', async ({ page, pageProblems }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const bytes = Array.from(await bookmanDocx());
    const result = await page.evaluate(
      async ([doc, font]) => {
        const api = (window as any).BymOffice;
        const notices: { message: string; kind: string }[] = [];
        const mounted = await api.mount(document.getElementById('editor'), {
          file: new Uint8Array(doc),
          fonts: [
            { family: 'Bym Test Bonum', source: new Uint8Array(font), as: ['Bookman Old Style'] },
            // Not a font file.
            { family: 'Bym Broken', source: new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]), as: 'Arial' },
          ],
          onNotice: (notice: { message: string; kind: string }) => notices.push({ message: notice.message, kind: notice.kind }),
        });
        const view = mounted.view;
        // Laid out with the font from the start: the page count never was another.
        const first = view.pageCount;
        await new Promise<void>((resolve) => (function wait() { view.layoutDone ? resolve() : setTimeout(wait, 20); })());
        const out = { kind: mounted.kind, first, pages: view.pageCount, notices, list: api.fonts.registeredFonts().map((entry: { family: string }) => entry.family) };
        mounted.destroy();
        return out;
      },
      [bytes, FONT] as [number[], number[]],
    );
    expect(result.kind).toBe('document');
    expect(result.pages).toBe(6);
    expect(result.first).toBeLessThanOrEqual(6);
    expect(result.list).toEqual(['Bym Test Bonum']);
    expect(result.notices.length).toBe(1);
    expect(result.notices[0].kind).toBe('error');
    expect(result.notices[0].message).toBe('The font "Bym Broken" could not be loaded; the file is shown without it');
    forgetRefusedFont(pageProblems, 'Bym Broken');
  });

  test('a family name with a word that starts with a digit can be registered', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const width = await page.evaluate(async (font) => {
      const api = (window as any).BymOffice;
      await api.registerFont({ family: 'Bym 39 Test-2', source: new Uint8Array(font).buffer });
      const measurer = api.createCanvasMeasurer();
      return measurer.width({ family: 'Bym 39 Test-2', size: 20, bold: false, italic: false }, 'Hello') as number;
    }, FONT);
    expect(width).toBeCloseTo(5 * PX, 3);
  });

  test('a font from an address and a Blob; the bold face is the one registered as bold', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const wide = Array.from(testFont({ advance: 2000 }));
    const result = await page.evaluate(
      async ([regular, bold]) => {
        const api = (window as any).BymOffice;
        const url = URL.createObjectURL(new Blob([new Uint8Array(regular)], { type: 'font/ttf' }));
        await api.registerFont({ family: 'Bym Faces', source: url });
        await api.registerFont({ family: 'Bym Faces', source: new Blob([new Uint8Array(bold)]), weight: 'bold' });
        URL.revokeObjectURL(url);
        const measurer = api.createCanvasMeasurer();
        return [
          measurer.width({ family: 'Bym Faces', size: 20, bold: false, italic: false }, 'abc') as number,
          measurer.width({ family: 'Bym Faces', size: 20, bold: true, italic: false }, 'abc') as number,
          measurer.hasFont('Bym Faces') as boolean,
          measurer.hasFont('Bym No Such Face') as boolean,
        ];
      },
      [FONT, wide] as [number[], number[]],
    );
    expect(result[0]).toBeCloseTo(3 * PX, 3);
    expect(result[1]).toBeCloseTo(6 * PX, 3);
    expect(result[2]).toBe(true);
    expect(result[3]).toBe(false);
  });

  test('a spreadsheet paints its cells anew with a font that arrives', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const result = await page.evaluate(async (font) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const mounted = await api.mount(host, { file: await api.createBlank('xlsx'), toolbar: false });
      const view = mounted.view;
      view.setCellInput(1, 1, 'HHHHHHHHHH');
      view.select({ top: 1, left: 1, bottom: 1, right: 1 });
      view.format({ fontName: 'Bookman Old Style Tur', fontSize: 24 });
      view.select({ top: 8, left: 0, bottom: 8, right: 0 });
      const frames = () => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve())));
      // Dark pixels of the grid: the glyphs of the test font are solid boxes, far darker than letters.
      const dark = () => {
        let count = 0;
        host.querySelectorAll('canvas').forEach((canvas) => {
          const data = (canvas as HTMLCanvasElement).getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
          for (let i = 0; i < data.length; i += 4) if (data[i + 3] > 200 && data[i] < 60 && data[i + 1] < 60 && data[i + 2] < 60) count++;
        });
        return count;
      };
      await frames();
      const before = dark();
      await api.registerFont({ family: 'Bym Test Bonum', source: new Uint8Array(font), as: ['Bookman Old Style'] });
      await frames();
      const after = dark();
      api.unregisterFont('Bym Test Bonum');
      await frames();
      const back = dark();
      mounted.destroy();
      return { before, after, back, ratio: window.devicePixelRatio };
    }, FONT);
    // Ten boxes of 0.8 x 0.7 em at 32 px: about 5700 px at one device pixel per px, less what the column cuts off.
    expect(result.after).toBeGreaterThan(result.before + 800 * result.ratio * result.ratio);
    expect(Math.abs(result.back - result.before)).toBeLessThan(40 * result.ratio * result.ratio);
  });

  test('the font list of the document toolbar offers the registered families', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await withoutEmbeddedFont(page);
    const bytes = Array.from(await bookmanDocx());
    await page.evaluate(async (doc) => {
      (window as any).__mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(doc) });
    }, bytes);
    const options = () => page.locator('#editor .bym-tb-font option').allTextContents();
    expect(await options()).toContain('Arial');
    expect(await options()).not.toContain('Bym Listed');
    await page.evaluate(async (font) => {
      await (window as any).BymOffice.registerFont({ family: 'Bym Listed', source: new Uint8Array(font) });
    }, FONT);
    await expect.poll(options).toContain('Bym Listed');
    await page.evaluate(() => (window as any).BymOffice.unregisterFont('Bym Listed'));
    await expect.poll(options).not.toContain('Bym Listed');
  });
  test('the WebView bridge registers the fonts of a load, from data and from an address, under its own policy', async ({ page, pageProblems }) => {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await page.goto('/dist/bymoffice.webview.html');
    await withoutEmbeddedFont(page);
    const base64 = (bytes: Uint8Array | number[]) => Buffer.from(Uint8Array.from(bytes)).toString('base64');
    const loaded = async (id: string) => {
      let found: any;
      await expect.poll(async () => (found = await page.evaluate((wanted) => (window as any).__messages.find((message: any) => message.id === wanted), id)) !== undefined).toBe(true);
      return found;
    };
    const send = (message: unknown) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
    const file = base64(await bookmanDocx());

    // Without a font the document has the pages the fallback font gives it.
    await send({ type: 'load', id: 'plain', data: file });
    expect((await loaded('plain')).type).toBe('loaded');
    const measure = () =>
      page.evaluate(() => {
        const measurer = (window as any).BymOffice.createCanvasMeasurer();
        return [
          measurer.width({ family: 'Bookman Old Style Tur', size: 20, bold: false, italic: false }, 'nnnnnnnnnn') as number,
          measurer.width({ family: 'Book Antiqua', size: 20, bold: false, italic: false }, 'nnnnn') as number,
        ];
      });
    const before = await measure();
    expect(before[0]).not.toBeCloseTo(200, 0);

    await send({
      type: 'load',
      id: 'fonts',
      data: file,
      fonts: [
        { family: 'Bym Bridge Bonum', data: base64(FONT), as: ['Bookman Old Style'] },
        { family: 'Bym Bridge Pagella', url: 'data:font/ttf;base64,' + base64(testFont({ advance: 2000 })), as: 'Book Antiqua' },
        { family: 'Bym Bridge Broken', data: base64([1, 2, 3, 4, 5, 6, 7, 8]) },
      ],
    });
    const answer = await loaded('fonts');
    expect(answer.type).toBe('loaded');
    expect(answer.warnings.length).toBe(1);
    expect(answer.warnings[0]).toMatch(/^The font "Bym Bridge Broken" could not be loaded: /);
    const after = await measure();
    expect(after[0]).toBeCloseTo(200, 3);
    expect(after[1]).toBeCloseTo(200, 3);
    expect(await page.evaluate(() => (window as any).BymOffice.registeredFonts().map((entry: { family: string }) => entry.family))).toEqual(['Bym Bridge Bonum', 'Bym Bridge Pagella']);

    // Fonts that do not follow the contract refuse the load, and the open file stays.
    await send({ type: 'load', id: 'bad', data: file, fonts: [{ family: 'X', url: 'https://elsewhere.example/a.ttf' }] });
    const refused = await loaded('bad');
    expect([refused.type, refused.code]).toEqual(['error', 'bad-message']);
    await expect(page.locator('#editor canvas').first()).toBeVisible();
    forgetRefusedFont(pageProblems, 'Bym Bridge Broken');
  });
});
