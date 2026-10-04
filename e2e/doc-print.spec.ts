// Printing a document: the pages are prepared as pictures in an element that
// only shows in print, the paper is that of the document, and the window's
// print() is called. No test opens a real print dialog: print() is replaced,
// and what would be printed is looked at in the document, under the print
// medium and, in Chromium, in the PDF the browser makes of the page.

import { readFile } from 'node:fs/promises';
import type { Frame, Page } from '@playwright/test';
import { p } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { DOC_STYLES, sampleDocx } from './fixtures/docx';
import { expect, expectResponsive, test } from './helpers';

const MARGINS = '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="300" w:footer="300" w:gutter="0"/>';
const A4 = `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${MARGINS}</w:sectPr>`;
const A4_LANDSCAPE = `<w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>${MARGINS}</w:sectPr>`;
const LETTER = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>${MARGINS}</w:sectPr>`;

// Paper sizes in points, as the style sheet states them.
const A4_PT = [595.3, 841.9];
const LETTER_PT = [612, 792];
// A4 at 200 dots per inch.
const A4_PIXELS = [1654, 2339];

const SENTENCE = 'The quick brown fox jumps over the lazy dog and keeps running along the river bank. ';

/** `count` pages of one section, each with a heading and a few lines of text. */
function pagesDocx(count: number, section: string): Promise<Uint8Array> {
  let body = '';
  for (let i = 1; i <= count; i++) body += p(`Sheet ${i}`, i > 1 ? '<w:pageBreakBefore/>' : '') + p(SENTENCE + SENTENCE + SENTENCE);
  return buildDocx({ styles: DOC_STYLES, body: body + section });
}

/** Two pages of A4, one of A4 landscape, one of US Letter. */
function mixedDocx(): Promise<Uint8Array> {
  const body =
    p('Portrait one') +
    p('Portrait two', '<w:pageBreakBefore/>') +
    p('End of the first section', A4) +
    p('Landscape') +
    p('End of the second section', A4_LANDSCAPE) +
    p('Letter') +
    LETTER;
  return buildDocx({ styles: DOC_STYLES, body });
}

type Target = Page | Frame;

interface ViewOptions {
  shadow?: 'open' | 'closed';
  toolbar?: boolean;
  locale?: string;
}

async function mountDoc(page: Page, bytes: Uint8Array, options: ViewOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/doc.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountDoc(input, opts), [Array.from(bytes), options] as [number[], ViewOptions]);
}

async function mountEdit(page: Page, bytes: Uint8Array, options: { readOnly?: boolean; locale?: string } = {}): Promise<void> {
  await page.goto('/e2e/pages/doc-edit.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), options] as [number[], object]);
  await page.evaluate(() => ((window as any).__docView = (window as any).__view));
}

/**
 * Replaces print() of the window: the call is recorded with what the
 * document holds at that moment. With `close` the dialog "closes" at once
 * (the afterprint event); without, it stays open until closeDialog().
 * Blob addresses made and given back are recorded as well.
 */
function fakePrint(target: Target, close: boolean): Promise<void> {
  return target.evaluate((closes) => {
    const w = window as any;
    w.__prints = [];
    w.__urls = { made: [] as string[], revoked: [] as string[], bytes: 0 };
    const make = URL.createObjectURL;
    const revoke = URL.revokeObjectURL;
    URL.createObjectURL = function (this: unknown, blob: Blob) {
      const url = make.call(URL, blob);
      if (blob.type === 'image/png') {
        w.__urls.made.push(url);
        w.__urls.bytes += blob.size;
      }
      return url;
    } as typeof URL.createObjectURL;
    URL.revokeObjectURL = function (this: unknown, url: string) {
      w.__urls.revoked.push(url);
      return revoke.call(URL, url);
    };
    window.print = () => {
      const element = document.querySelector('body > .bym-print');
      w.__prints.push({ sheets: element ? element.children.length : -1 });
      if (closes) window.dispatchEvent(new Event('afterprint'));
    };
  }, close);
}

/** Runs `action` (a method of the view, without arguments unless `argument` is given) at the moment the print job has `count` pages ready. */
function whenPagesReady(target: Target, count: number, action: string, argument?: string): Promise<void> {
  return target.evaluate(
    ([wanted, method, value]) => {
      const observer = new MutationObserver(() => {
        if (document.querySelectorAll('.bym-print-sheet').length < (wanted as number)) return;
        observer.disconnect();
        const view = (window as any).__docView;
        if (value === null) view[method as string]();
        else view[method as string](value);
      });
      observer.observe(document.body, { childList: true, subtree: true });
    },
    [count, action, argument === undefined ? null : argument] as [number, string, string | null],
  );
}

const closeDialog = (target: Target) => target.evaluate(() => void window.dispatchEvent(new Event('afterprint')));

/** Starts print() of the view and keeps its outcome in window.__printed. */
function startPrint(target: Target, options?: { dpi?: number }): Promise<void> {
  return target.evaluate((opts) => {
    const w = window as any;
    w.__printed = undefined;
    w.__docView.print(opts).then(
      (result: unknown) => (w.__printed = { result }),
      (error: any) => (w.__printed = { error: { name: error && error.name, code: error && error.code, message: String(error && error.message) } }),
    );
  }, options);
}

const printCalls = (target: Target) => target.evaluate(() => ((window as any).__prints as { sheets: number }[]).slice());
const printed = (target: Target) => target.evaluate(() => (window as any).__printed);
const waitForDialog = (target: Target, calls = 1) =>
  expect.poll(async () => (await printCalls(target)).length, { timeout: 60_000, message: 'print() of the window is called' }).toBe(calls);
const waitForOutcome = async (target: Target) => {
  await expect.poll(() => printed(target), { timeout: 60_000, message: 'print() of the view settles' }).toBeTruthy();
  return printed(target);
};

interface Frame_ {
  present: boolean;
  /** The element is a child of the body of the document, not of the view or a shadow root. */
  atBody: boolean;
  text: string;
  styles: number;
  css: string;
  sheets: { className: string; style: string }[];
  images: { width: number; height: number; complete: boolean; alt: string; scheme: string }[];
}

/** What the print job has put into the document. */
function frameState(target: Target): Promise<Frame_> {
  return target.evaluate(() => {
    const element = document.querySelector('.bym-print') as HTMLElement | null;
    const styles = document.querySelectorAll('style[data-bymoffice-print]');
    const sheets = element ? (Array.from(element.children) as HTMLElement[]) : [];
    const images = element ? (Array.from(element.querySelectorAll('img')) as HTMLImageElement[]) : [];
    return {
      present: !!element,
      atBody: !!element && element.parentNode === document.body,
      text: element ? element.textContent || '' : '',
      styles: styles.length,
      css: styles.length ? styles[0].textContent || '' : '',
      // As it was written: browsers put spaces and a last semicolon into what they give back.
      sheets: sheets.map((sheet) => ({ className: sheet.className, style: (sheet.getAttribute('style') || '').replace(/\s+/g, '').replace(/;$/, '') })),
      images: images.map((image) => ({
        width: image.naturalWidth,
        height: image.naturalHeight,
        complete: image.complete,
        alt: image.alt,
        scheme: image.src.slice(0, 5),
      })),
    };
  });
}

const urls = (target: Target) => target.evaluate(() => (window as any).__urls as { made: string[]; revoked: string[] });

/** Nothing of a print job is left in the document, and every picture it made was given back. */
async function expectCleanedUp(target: Target, pictures?: number): Promise<void> {
  const left = await frameState(target);
  expect(left.present).toBe(false);
  expect(left.styles).toBe(0);
  const all = await urls(target);
  if (pictures !== undefined) expect(all.made.length).toBe(pictures);
  expect(all.made.filter((url) => all.revoked.indexOf(url) < 0)).toEqual([]);
}

/** Counts the pixels of a printed page by colour. */
function pictureStats(target: Target, index: number): Promise<{ dark: number; red: number; blue: number; white: number; total: number }> {
  return target.evaluate((n) => {
    const image = document.querySelectorAll('.bym-print img')[n] as HTMLImageElement;
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d')!;
    context.drawImage(image, 0, 0);
    const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
    const out = { dark: 0, red: 0, blue: 0, white: 0, total: data.length / 4 };
    for (let i = 0; i < data.length; i += 4) {
      const [red, green, blue] = [data[i], data[i + 1], data[i + 2]];
      if (red >= 250 && green >= 250 && blue >= 250) out.white++;
      else if (red < 80 && green < 80 && blue < 80) out.dark++;
      else if (red > 180 && green < 60 && blue < 60) out.red++;
      else if (blue > 180 && red < 60 && green < 60) out.blue++;
    }
    return out;
  }, index);
}

/** The pages of a PDF: how many it has and the size of each in points. */
function pdfPages(pdf: Buffer): { count: number; sizes: number[][] } {
  const text = pdf.toString('latin1');
  const count = (text.match(/\/Type\s*\/Page(?![A-Za-z])/g) || []).length;
  const sizes: number[][] = [];
  const box = /\/MediaBox\s*\[\s*([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s*\]/g;
  for (let match = box.exec(text); match; match = box.exec(text)) sizes.push([Number(match[3]) - Number(match[1]), Number(match[4]) - Number(match[2])]);
  return { count, sizes };
}

const near = (actual: number[], wanted: number[], tolerance = 1) => Math.abs(actual[0] - wanted[0]) <= tolerance && Math.abs(actual[1] - wanted[1]) <= tolerance;

test.describe('printing a document', () => {
  test('print() prepares one picture per page on paper of the document’s size, calls print() of the window and cleans up after the dialog', async ({ page }) => {
    await mountDoc(page, await pagesDocx(3, A4));
    await fakePrint(page, false);
    // On screen nothing of a print job exists before it is asked for.
    expect((await frameState(page)).present).toBe(false);
    await startPrint(page);
    await waitForDialog(page);
    // print() was called once, with every page ready.
    expect(await printCalls(page)).toEqual([{ sheets: 3 }]);
    const frame = await frameState(page);
    expect(frame.atBody).toBe(true);
    expect(frame.styles).toBe(1);
    expect(frame.css).toContain(`@page{size:${A4_PT[0]}pt ${A4_PT[1]}pt;margin:0}`);
    expect(frame.sheets.map((sheet) => sheet.style)).toEqual(new Array(3).fill(`width:${A4_PT[0]}pt;height:${A4_PT[1]}pt`));
    expect(frame.images.map((image) => [image.width, image.height, image.complete, image.scheme])).toEqual(new Array(3).fill([...A4_PIXELS, true, 'blob:']));
    // The only words in the element are the view's own; nothing of the document is there as text.
    expect(frame.images.map((image) => image.alt)).toEqual(['Page 1', 'Page 2', 'Page 3']);
    expect(frame.text).toBe('');
    // The text of the pages is in the pictures.
    for (let i = 0; i < 3; i++) {
      const stats = await pictureStats(page, i);
      expect(stats.dark).toBeGreaterThan(2000);
      expect(stats.white).toBeGreaterThan(stats.total * 0.9);
    }
    // On screen the element is not shown and the view is as it was.
    await expect(page.locator('.bym-print')).toBeHidden();
    await expect(page.locator('.bym-doc-scroller')).toBeVisible();
    expect(await page.evaluate(() => (window as any).__docView.printing)).toBe(true);
    // The promise is still open while the dialog is.
    expect(await printed(page)).toBeUndefined();

    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 3, dpi: 200 } });
    await expectCleanedUp(page, 3);
    expect(await page.evaluate(() => (window as any).__docView.printing)).toBe(false);
    // The start was announced, and the end takes its place: "preparing" must not stay on the screen.
    expect(await page.evaluate(() => (window as any).__notices)).toEqual([
      { message: 'Preparing to print…', kind: 'info' },
      { message: 'The print dialog was closed', kind: 'info' },
    ]);
  });

  test('a wrong option is refused even while an earlier call is preparing its pages', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2, LETTER));
    await fakePrint(page, false);
    const second = await page.evaluate(() => {
      const view = (window as any).__docView;
      const first = view.print();
      first.catch(() => undefined);
      (window as any).__first = first;
      return view.print({ dpi: -1 }).then(
        () => 'resolved',
        (error: Error) => error.name,
      );
    });
    expect(second).toBe('RangeError');
    // The earlier call goes on untouched.
    await waitForDialog(page);
    await closeDialog(page);
    expect(await page.evaluate(() => (window as any).__first)).toEqual({ pages: 2, dpi: 200 });
    expect((await printCalls(page)).length).toBe(1);
  });

  test('the resolution option decides the size of the pictures; a wrong one is refused before anything is prepared', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2, LETTER));
    await fakePrint(page, false);
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    const frame = await frameState(page);
    expect(frame.images.map((image) => [image.width, image.height])).toEqual([
      [816, 1056],
      [816, 1056],
    ]);
    expect(frame.css).toContain(`@page{size:${LETTER_PT[0]}pt ${LETTER_PT[1]}pt;margin:0}`);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 2, dpi: 96 } });

    await startPrint(page, { dpi: -1 });
    const outcome = await waitForOutcome(page);
    expect(outcome.error.name).toBe('RangeError');
    expect((await printCalls(page)).length).toBe(1);
    await expectCleanedUp(page, 2);
    // A mistake of the caller is not something to tell the user.
    expect(await page.evaluate(() => (window as any).__notices.filter((notice: any) => notice.kind === 'error'))).toEqual([]);
  });

  test('under the print medium only the pages show: the bar, the scrolling pages and the host page are hidden', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2, A4));
    await fakePrint(page, false);
    await startPrint(page);
    await waitForDialog(page);
    await page.emulateMedia({ media: 'print' });
    await expect(page.locator('.bym-print')).toBeVisible();
    await expect(page.locator('.bym-print img').first()).toBeVisible();
    await expect(page.locator('.bym-doc-bar')).toBeHidden();
    await expect(page.locator('.bym-doc-scroller')).toBeHidden();
    await expect(page.locator('#host')).toBeHidden();
    await expect(page.locator('#before')).toBeHidden();
    await expect(page.locator('#state')).toBeHidden();
    // Every sheet is as large as its paper and starts at the left edge: 595.3 pt are 793.73 px.
    const boxes = await page.evaluate(() =>
      (Array.from(document.querySelectorAll('.bym-print-sheet')) as HTMLElement[]).map((sheet) => {
        const rect = sheet.getBoundingClientRect();
        const image = (sheet.firstElementChild as HTMLElement).getBoundingClientRect();
        return [rect.left, rect.width, rect.height, image.width - rect.width, image.height - rect.height, image.left - rect.left, image.top - rect.top];
      }),
    );
    for (const box of boxes) {
      expect(box[0]).toBe(0);
      expect(Math.abs(box[1] - 793.73)).toBeLessThan(0.5);
      expect(Math.abs(box[2] - 1122.53)).toBeLessThan(0.5);
      for (let i = 3; i < 7; i++) expect(Math.abs(box[i])).toBeLessThan(0.01);
    }
    // The second sheet follows the first without a gap.
    expect(await page.evaluate(() => getComputedStyle(document.body).margin)).toBe('0px');
    await page.emulateMedia({ media: 'screen' });
    await expect(page.locator('.bym-print')).toBeHidden();
    await expect(page.locator('.bym-doc-scroller')).toBeVisible();
    await closeDialog(page);
    await waitForOutcome(page);
    await expectCleanedUp(page, 2);
  });

  test('the PDF the browser makes has one page per page of the document, of the document’s paper size', async ({ page, browserName }) => {
    test.skip(browserName !== 'chromium' || !!test.info().project.use.isMobile, 'only Chromium makes a PDF of a page, and only on the desktop');
    for (const [section, size] of [
      [A4, A4_PT],
      [LETTER, LETTER_PT],
    ] as [string, number[]][]) {
      await mountDoc(page, await pagesDocx(5, section));
      await fakePrint(page, false);
      await startPrint(page);
      await waitForDialog(page);
      const pdf = pdfPages(await page.pdf({ preferCSSPageSize: true }));
      expect(pdf.count).toBe(5);
      expect(pdf.sizes.length).toBe(5);
      for (const actual of pdf.sizes) expect(near(actual, size), `${actual} is ${size}`).toBe(true);
      // Making the PDF may have closed the "dialog" by itself (the browser sends afterprint); either way the job ends.
      await closeDialog(page);
      await waitForOutcome(page);
      await expectCleanedUp(page);
    }
  });

  test('landscape pages are printed on landscape paper', async ({ page, browserName }) => {
    await mountDoc(page, await pagesDocx(2, A4_LANDSCAPE));
    await fakePrint(page, false);
    await startPrint(page);
    await waitForDialog(page);
    const frame = await frameState(page);
    expect(frame.css).toContain(`@page{size:${A4_PT[1]}pt ${A4_PT[0]}pt;margin:0}`);
    expect(frame.images.map((image) => [image.width, image.height])).toEqual(new Array(2).fill([A4_PIXELS[1], A4_PIXELS[0]]));
    if (browserName === 'chromium' && !test.info().project.use.isMobile) {
      const pdf = pdfPages(await page.pdf({ preferCSSPageSize: true }));
      expect(pdf.count).toBe(2);
      for (const actual of pdf.sizes) expect(near(actual, [A4_PT[1], A4_PT[0]]), String(actual)).toBe(true);
    }
    await closeDialog(page);
    await waitForOutcome(page);
    await expectCleanedUp(page, 2);
  });

  test('sections of different paper sizes: every page gets its own paper where the browser can name pages, the first size elsewhere', async ({ page, browserName }) => {
    await mountDoc(page, await mixedDocx());
    await fakePrint(page, false);
    await startPrint(page);
    await waitForDialog(page);
    const frame = await frameState(page);
    expect(frame.images.length).toBe(4);
    expect(frame.images.map((image) => [image.width, image.height])).toEqual([A4_PIXELS, A4_PIXELS, [A4_PIXELS[1], A4_PIXELS[0]], [1700, 2200]]);
    expect(frame.css).toContain(`@page{size:${A4_PT[0]}pt ${A4_PT[1]}pt;margin:0}`);
    const named = await page.evaluate(() => CSS.supports('page', 'bym-print-0'));
    expect(frame.sheets.map((sheet) => sheet.className)).toEqual([
      'bym-print-sheet bym-print-s0',
      'bym-print-sheet bym-print-s0',
      'bym-print-sheet bym-print-s1',
      'bym-print-sheet bym-print-s2',
    ]);
    const portrait = `width:${A4_PT[0]}pt;height:${A4_PT[1]}pt`;
    if (named) {
      expect(frame.css).toContain(`@page bym-print-1{size:${A4_PT[1]}pt ${A4_PT[0]}pt;margin:0}`);
      expect(frame.css).toContain(`@page bym-print-2{size:${LETTER_PT[0]}pt ${LETTER_PT[1]}pt;margin:0}`);
      expect(frame.sheets.map((sheet) => sheet.style)).toEqual([portrait, portrait, `width:${A4_PT[1]}pt;height:${A4_PT[0]}pt`, `width:${LETTER_PT[0]}pt;height:${LETTER_PT[1]}pt`]);
    } else {
      // Every sheet is of the first size; the other pages are fitted into it, whole.
      expect(frame.css).not.toContain('@page bym-print-');
      expect(frame.sheets.map((sheet) => sheet.style)).toEqual(new Array(4).fill(portrait));
      await page.emulateMedia({ media: 'print' });
      const inside = await page.evaluate(() =>
        (Array.from(document.querySelectorAll('.bym-print-sheet')) as HTMLElement[]).map((sheet) => {
          const rect = sheet.getBoundingClientRect();
          const image = (sheet.firstElementChild as HTMLElement).getBoundingClientRect();
          return image.left >= rect.left - 0.01 && image.right <= rect.right + 0.01 && image.top >= rect.top - 0.01 && image.bottom <= rect.bottom + 0.01;
        }),
      );
      expect(inside).toEqual([true, true, true, true]);
      await page.emulateMedia({ media: 'screen' });
    }
    if (browserName === 'chromium' && !test.info().project.use.isMobile) {
      expect(named).toBe(true);
      const pdf = pdfPages(await page.pdf({ preferCSSPageSize: true }));
      expect(pdf.count).toBe(4);
      const wanted = [A4_PT, A4_PT, [A4_PT[1], A4_PT[0]], LETTER_PT];
      pdf.sizes.forEach((actual, index) => expect(near(actual, wanted[index]), `page ${index + 1}: ${actual}`).toBe(true));
    }
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 4, dpi: 200 } });
    await expectCleanedUp(page, 4);
  });

  test('the pictures of the document are printed, not their placeholders', async ({ page }) => {
    await mountDoc(page, await sampleDocx());
    await fakePrint(page, false);
    await startPrint(page);
    await waitForDialog(page);
    // The picture of the sample (120 x 60 px, half red and half blue) at 200 dpi on the first page.
    const stats = await pictureStats(page, 0);
    const half = (120 * 60 * (200 / 96) * (200 / 96)) / 2;
    expect(stats.red).toBeGreaterThan(half * 0.9);
    expect(stats.red).toBeLessThan(half * 1.1);
    // The links of the sample are blue as well.
    expect(stats.blue).toBeGreaterThan(half * 0.9);
    expect(stats.blue).toBeLessThan(half * 1.5);
    await closeDialog(page);
    await waitForOutcome(page);
    await expectCleanedUp(page);
  });

  test('a long document is prepared within its time, without freezing the page and without keeping the pages as canvases', async ({ page, browserName }) => {
    test.setTimeout(180_000);
    const count = 60;
    await mountDoc(page, await pagesDocx(count, A4));
    await fakePrint(page, false);
    // The longest time the page could not run a timer while the pages were prepared.
    await page.evaluate(() => {
      const w = window as any;
      w.__gap = 0;
      w.__started = performance.now();
      let last = performance.now();
      w.__ticker = setInterval(() => {
        const now = performance.now();
        if (now - last > w.__gap) w.__gap = now - last;
        last = now;
      }, 10);
    });
    await startPrint(page);
    await expectResponsive(page);
    // While the pages are prepared the bar says how far they are, and the button says that it is at work.
    await expect(page.locator('.bym-doc-pages')).toContainText('Preparing to print');
    await expect(page.locator('.bym-doc-print')).toHaveAttribute('aria-busy', 'true');
    await waitForDialog(page);
    const measured = await page.evaluate(() => {
      const w = window as any;
      clearInterval(w.__ticker);
      return { ms: performance.now() - w.__started, gap: w.__gap, bytes: w.__urls.bytes, mounted: w.__docView.getViewState().mounted.length, canvases: document.querySelectorAll('canvas').length };
    });
    console.log(`print of ${count} A4 pages (${browserName}, ${test.info().project.name}): ${Math.round(measured.ms)} ms, longest pause ${Math.round(measured.gap)} ms, ${Math.round(measured.bytes / 1024)} KB of pictures`);
    // The pictures of sixty pages of text are small next to the limit on their bytes.
    expect(measured.bytes).toBeLessThan(64 * 1024 * 1024);
    expect(await printCalls(page)).toEqual([{ sheets: count }]);
    expect(measured.ms).toBeLessThan(90_000);
    expect(measured.gap).toBeLessThan(2_000);
    // Only the pages in sight have a canvas; the one the pages were painted on is gone.
    expect(measured.canvases).toBe(measured.mounted);
    expect(measured.mounted).toBeLessThan(8);
    const frame = await frameState(page);
    expect(frame.images.length).toBe(count);
    // A phone gets fewer pixels for all pages together: sixty pages of A4 are painted at a little less than 200 dpi there.
    const coarse = await page.evaluate(() => matchMedia('(pointer:coarse)').matches);
    const dpi = coarse ? Math.sqrt(200_000_000 / (count * (11906 / 1440) * (16838 / 1440))) : 200;
    if (coarse) expect(dpi).toBeLessThan(200);
    const pixels = [Math.round((11906 / 15) * (dpi / 96)), Math.round((16838 / 15) * (dpi / 96))];
    expect(frame.images.every((image) => image.complete && image.width === pixels[0] && image.height === pixels[1])).toBe(true);
    await expect(page.locator('.bym-doc-print')).toHaveAttribute('aria-busy', 'false');
    await expect(page.locator('.bym-doc-pages')).toContainText(`/ ${count}`);
    await closeDialog(page);
    const outcome = await waitForOutcome(page);
    expect(outcome.result.pages).toBe(count);
    expect(Math.abs(outcome.result.dpi - dpi)).toBeLessThan(0.001);
    await expectCleanedUp(page, count);
  });

  test('print() waits for pages that are not laid out yet', async ({ page }) => {
    test.setTimeout(120_000);
    await mountDoc(page, await pagesDocx(40, A4));
    await fakePrint(page, true);
    // Asked for right away: the layout is still at work on a document of this length in at least some runs; either way all pages are printed.
    await startPrint(page, { dpi: 72 });
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 40, dpi: 72 } });
    expect(await printCalls(page)).toEqual([{ sheets: 40 }]);
    expect(await page.evaluate(() => (window as any).__docView.layoutDone)).toBe(true);
    await expectCleanedUp(page, 40);
  });

  test('a second call while the pages are prepared joins the first; one after a dialog that never said it closed starts anew', async ({ page }) => {
    await mountDoc(page, await pagesDocx(6, A4));
    await fakePrint(page, false);
    const same = await page.evaluate(() => {
      const view = (window as any).__docView;
      const first = view.print();
      const second = view.print({ dpi: 96 });
      (window as any).__first = first;
      return first === second;
    });
    expect(same).toBe(true);
    await waitForDialog(page);
    // One dialog, with the pages of the first call.
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
    expect((await frameState(page)).images[0].width).toBe(A4_PIXELS[0]);
    // The dialog "is open" (print() came back, no afterprint): a new call ends the old job and prepares its own pages.
    await page.evaluate(() => {
      const w = window as any;
      w.__first.then((result: unknown) => (w.__firstResult = result));
      w.__second = w.__docView.print({ dpi: 96 });
      w.__second.then((result: unknown) => (w.__secondResult = result));
    });
    await waitForDialog(page, 2);
    expect(await page.evaluate(() => (window as any).__firstResult)).toEqual({ pages: 6, dpi: 200 });
    const frame = await frameState(page);
    expect(frame.styles).toBe(1);
    expect(frame.images.map((image) => image.width)).toEqual(new Array(6).fill(794));
    await closeDialog(page);
    await expect.poll(() => page.evaluate(() => (window as any).__secondResult)).toEqual({ pages: 6, dpi: 96 });
    await expectCleanedUp(page, 12);
  });

  test('a press in the page ends the wait for a dialog that never said it closed', async ({ page }) => {
    await mountDoc(page, await pagesDocx(1, A4));
    await fakePrint(page, false);
    await startPrint(page);
    await waitForDialog(page);
    expect(await printed(page)).toBeUndefined();
    await page.locator('#before').click();
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 1, dpi: 200 } });
    await expectCleanedUp(page, 1);
  });

  test('destroy() while the pages are prepared, and while the dialog is open, leaves nothing behind', async ({ page }) => {
    await mountDoc(page, await pagesDocx(12, A4));
    await fakePrint(page, false);
    // In the middle of the preparation.
    await whenPagesReady(page, 3, 'destroy');
    await startPrint(page);
    const outcome = await waitForOutcome(page);
    expect(outcome.error).toMatchObject({ name: 'DocumentPrintError', code: 'cancelled' });
    expect(await printCalls(page)).toEqual([]);
    await expectCleanedUp(page);
    await expect(page.locator('.bym-doc')).toHaveCount(0);
    expect(await page.evaluate(() => document.querySelectorAll('canvas').length)).toBe(0);
    // Nothing is said to the user about a view that is gone.
    expect(await page.evaluate(() => (window as any).__notices.filter((notice: any) => notice.kind === 'error'))).toEqual([]);
    // A view that is gone cannot print.
    await startPrint(page);
    expect((await waitForOutcome(page)).error.code).toBe('cancelled');

    // While the dialog is open: the job ends as if the dialog had been closed.
    await page.evaluate(([input]) => (window as any).__mountDoc(input, {}), [Array.from(await pagesDocx(2, A4))]);
    await startPrint(page);
    await waitForDialog(page);
    expect((await frameState(page)).images.length).toBe(2);
    await page.evaluate(() => (window as any).__docView.destroy());
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 2, dpi: 200 } });
    await expectCleanedUp(page);
    await expectResponsive(page);
    // No listener of the wait for the dialog is left on the window: a press and an afterprint do nothing.
    await page.locator('#before').click();
    await closeDialog(page);
  });

  test('a window that cannot print: the call is refused with a reason and the user is told', async ({ page }) => {
    await mountDoc(page, await pagesDocx(1, A4), { locale: 'tr' });
    await page.evaluate(() => {
      (window as any).print = undefined;
    });
    await startPrint(page);
    const outcome = await waitForOutcome(page);
    expect(outcome.error).toMatchObject({ name: 'DocumentPrintError', code: 'unavailable' });
    expect(await page.evaluate(() => (window as any).__notices)).toEqual([{ message: 'Burada yazdırma kullanılamıyor', kind: 'error' }]);
    expect((await frameState(page)).present).toBe(false);
    // A print() that throws.
    await page.evaluate(() => {
      (window as any).__urls = { made: [], revoked: [] };
      window.print = () => {
        throw new Error('not allowed');
      };
    });
    await startPrint(page);
    const thrown = await waitForOutcome(page);
    expect(thrown.error).toMatchObject({ name: 'DocumentPrintError', code: 'failed' });
    expect(thrown.error.message).toContain('not allowed');
    expect((await frameState(page)).present).toBe(false);
    expect(await page.evaluate(() => (window as any).__notices.map((notice: any) => notice.message))).toEqual([
      'Burada yazdırma kullanılamıyor',
      'Yazdırmaya hazırlanıyor…',
      'Belge yazdırılamadı',
    ]);
  });

  test('the Print button of the bar and Ctrl+P print the document; the browser’s own printing of the host page is prevented', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2, A4), { locale: 'tr' });
    await fakePrint(page, true);
    const button = page.locator('.bym-doc-print');
    await expect(button).toBeVisible();
    await expect(button).toHaveAttribute('aria-label', 'Yazdır');
    await expect(button).toHaveAttribute('title', 'Yazdır');
    // Inside the bar, next to the other buttons.
    const boxes = await page.evaluate(() => {
      const bar = document.querySelector('.bym-doc-bar')!.getBoundingClientRect();
      const print = document.querySelector('.bym-doc-print')!.getBoundingClientRect();
      const fit = document.querySelector('.bym-doc-fit')!.getBoundingClientRect();
      return { inside: print.left >= bar.left && print.right <= bar.right && print.top >= bar.top && print.bottom <= bar.bottom, after: print.left >= fit.right, size: [print.width, print.height] };
    });
    expect(boxes.inside).toBe(true);
    expect(boxes.after).toBe(true);
    expect(boxes.size[0]).toBeGreaterThanOrEqual(32);
    await button.click();
    await waitForDialog(page);
    await expect.poll(() => page.evaluate(() => (window as any).__docView.printing)).toBe(false);
    await expectCleanedUp(page, 2);

    // The key, with the focus in the pages: the view's print job and not the browser's default.
    const prevented = await page.evaluate(() => {
      const view = (window as any).__docView;
      view.focus();
      const event = new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true });
      document.activeElement!.dispatchEvent(event);
      return event.defaultPrevented;
    });
    expect(prevented).toBe(true);
    await waitForDialog(page, 2);
    await expect.poll(() => page.evaluate(() => (window as any).__docView.printing)).toBe(false);
    // With the focus on a button of the bar as well; other key combinations with P are not taken.
    const others = await page.evaluate(() => {
      const press = (init: KeyboardEventInit) => {
        const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, ...init });
        document.querySelector('.bym-doc-zoom-in')!.dispatchEvent(event);
        return event.defaultPrevented;
      };
      return [press({ key: 'p' }), press({ key: 'P', ctrlKey: true, shiftKey: true }), press({ key: 'p', ctrlKey: true, altKey: true }), press({ key: 'P', metaKey: true })];
    });
    expect(others).toEqual([false, false, false, true]);
    await waitForDialog(page, 3);
    await expect.poll(() => page.evaluate(() => (window as any).__docView.printing)).toBe(false);
    await expectCleanedUp(page, 6);
  });

  test('a view without its bar prints through the API', async ({ page }) => {
    await mountDoc(page, await pagesDocx(2, A4), { toolbar: false });
    await expect(page.locator('.bym-doc-print')).toHaveCount(0);
    await fakePrint(page, true);
    await startPrint(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 2, dpi: 200 } });
    expect(await printCalls(page)).toEqual([{ sheets: 2 }]);
    await expectCleanedUp(page, 2);
  });

  for (const mode of ['open', 'closed'] as const) {
    test(`a view inside a ${mode} shadow root prints: its pages go to the document, outside the shadow root`, async ({ page }) => {
      await mountDoc(page, await pagesDocx(2, A4), { shadow: mode });
      await fakePrint(page, false);
      await startPrint(page);
      await waitForDialog(page);
      const frame = await frameState(page);
      expect(frame.atBody).toBe(true);
      expect(frame.styles).toBe(1);
      expect(frame.images.map((image) => [image.width, image.height, image.complete])).toEqual(new Array(2).fill([...A4_PIXELS, true]));
      // Nothing of the job is inside the shadow root.
      expect(await page.evaluate(() => (window as any).__shadow.querySelectorAll('.bym-print, style[data-bymoffice-print]').length)).toBe(0);
      await page.emulateMedia({ media: 'print' });
      await expect(page.locator('body > .bym-print')).toBeVisible();
      // The view in the shadow root is hidden with the element that hosts it.
      expect(await page.evaluate(() => ((window as any).__shadow.querySelector('.bym-doc') as HTMLElement).getClientRects().length)).toBe(0);
      await page.emulateMedia({ media: 'screen' });
      await closeDialog(page);
      await waitForOutcome(page);
      await expectCleanedUp(page, 2);
    });
  }

  test('a view inside an iframe prints the window of the frame, not the page around it', async ({ page }) => {
    await page.goto('/e2e/pages/blank.html');
    await page.evaluate(() => {
      const frame = document.createElement('iframe');
      frame.id = 'inner';
      frame.name = 'inner';
      frame.style.cssText = 'display:block;width:100%;height:560px;border:0';
      frame.src = '/e2e/pages/doc.html';
      document.body.appendChild(frame);
      (window as any).__prints = [];
      window.print = () => (window as any).__prints.push('outer');
    });
    const inner = (await (await page.waitForSelector('#inner')).contentFrame())!;
    await expect(inner.locator('#state')).toHaveText('ready');
    await inner.evaluate(([input]) => (window as any).__mountDoc(input, {}), [Array.from(await pagesDocx(2, A4))]);
    await fakePrint(inner, false);
    await startPrint(inner);
    await waitForDialog(inner);
    expect(await printCalls(inner)).toEqual([{ sheets: 2 }]);
    // The page around the frame was neither printed nor changed.
    expect(await page.evaluate(() => (window as any).__prints)).toEqual([]);
    expect((await frameState(page)).present).toBe(false);
    expect((await frameState(page)).styles).toBe(0);
    const frame = await frameState(inner);
    expect(frame.atBody).toBe(true);
    expect(frame.images.length).toBe(2);
    await closeDialog(inner);
    expect(await waitForOutcome(inner)).toEqual({ result: { pages: 2, dpi: 200 } });
    await expectCleanedUp(inner, 2);
  });

  test('the editor prints, read-only or not: mounted.print(), editor.print(), the bar under the toolbar, Ctrl+P on the toolbar', async ({ page }) => {
    for (const readOnly of [false, true]) {
      await mountEdit(page, await pagesDocx(2, A4), { readOnly });
      await fakePrint(page, true);
      await expect(page.locator('.bym-doc-print')).toBeVisible();
      const results = await page.evaluate(async () => {
        const mounted = (window as any).__mounted;
        return [await mounted.print({ dpi: 96 }), await mounted.editor.print({ dpi: 120 })];
      });
      expect(results).toEqual([
        { pages: 2, dpi: 96 },
        { pages: 2, dpi: 120 },
      ]);
      await page.locator('.bym-doc-print').click();
      await waitForDialog(page, 3);
      await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
      if (!readOnly) {
        // The formatting toolbar is outside the pages: the key works there too.
        const prevented = await page.evaluate(() => {
          const event = new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true });
          document.querySelector('.bym-toolbar button')!.dispatchEvent(event);
          return event.defaultPrevented;
        });
        expect(prevented).toBe(true);
        await waitForDialog(page, 4);
        await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
      }
      await expectCleanedUp(page);
      // Printing changes nothing in the document.
      expect(await page.evaluate(() => [(window as any).__changes, (window as any).__view.isModified])).toEqual([0, false]);
    }
  });

  test('a document that is edited while its pages are prepared is printed as it is after the edit', async ({ page }) => {
    await mountEdit(page, await pagesDocx(8, A4));
    await page.waitForFunction(() => (window as any).__view.layoutDone === true);
    await fakePrint(page, false);
    // Typed when three pages are ready: every page prepared so far is out of date.
    await page.evaluate(() => (window as any).__view.focus());
    await whenPagesReady(page, 3, 'insertText', 'Typed while printing');
    await startPrint(page);
    await waitForDialog(page);
    expect(await page.evaluate(() => (window as any).BymOffice.documentText((window as any).__view.document).split('\n')[0])).toBe('Typed while printingSheet 1');
    expect(await printCalls(page)).toEqual([{ sheets: 8 }]);
    const all = await urls(page);
    // More pictures were made than there are pages, and the ones of the first round were given back.
    expect(all.made.length).toBeGreaterThan(8);
    expect(all.made.filter((url) => all.revoked.indexOf(url) < 0).length).toBe(8);
    expect((await frameState(page)).styles).toBe(1);
    // The first page is printed with what was typed: it has more ink than the page of the first round had.
    expect(await page.evaluate(() => (window as any).__changes)).toBe(1);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 8, dpi: 200 } });
    await expectCleanedUp(page);
  });
});

test.describe('printing through the WebView bridge', () => {
  const base64 = (bytes: Uint8Array) => Buffer.from(bytes).toString('base64');

  async function startBridge(page: Page): Promise<void> {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: string) => (window as any).__messages.push(JSON.parse(json)) };
    });
    await page.goto('/dist/bymoffice.webview.html');
  }
  const send = (page: Page, message: object) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
  const answer = async (page: Page, id: string) => {
    let found: any;
    await expect
      .poll(async () => (found = await page.evaluate((wanted) => (window as any).__messages.find((message: any) => message.id === wanted), id)), { timeout: 60_000, message: `answer to ${id}` })
      .toBeTruthy();
    return found;
  };

  test('the print command prints the document under the page’s content security policy and answers with done', async ({ page }) => {
    // The page as it is built, with its policy: an image or a style it refuses would show as a console error.
    expect(await readFile(new URL('../dist/bymoffice.webview.html', import.meta.url), 'utf8')).toContain("default-src 'none'");
    await startBridge(page);
    await send(page, { type: 'load', id: 'l', data: base64(await pagesDocx(3, A4)) });
    expect((await answer(page, 'l')).type).toBe('loaded');
    await fakePrint(page, false);
    await send(page, { type: 'print', id: 'p', dpi: 150 });
    await waitForDialog(page);
    const frame = await frameState(page);
    expect(frame.atBody).toBe(true);
    // The pictures were loaded and the style applied: the policy allows blob: images and inline styles.
    expect(frame.images.map((image) => [image.width, image.height, image.complete, image.scheme])).toEqual(new Array(3).fill([1240, 1754, true, 'blob:']));
    await expect(page.locator('.bym-print')).toBeHidden();
    await page.emulateMedia({ media: 'print' });
    await expect(page.locator('.bym-print img').first()).toBeVisible();
    await expect(page.locator('#editor')).toBeHidden();
    // The page's own dark background does not reach the paper.
    expect(await page.evaluate(() => getComputedStyle(document.body).backgroundColor)).toBe('rgb(255, 255, 255)');
    await page.emulateMedia({ media: 'screen' });
    // The dialog is open: no answer yet, and other commands are still handled.
    expect(await page.evaluate(() => (window as any).__messages.some((message: any) => message.id === 'p'))).toBe(false);
    await send(page, { type: 'setLocale', id: 's', locale: 'tr' });
    expect(await answer(page, 's')).toEqual({ type: 'done', id: 's', command: 'setLocale' });
    await closeDialog(page);
    expect(await answer(page, 'p')).toEqual({ type: 'done', id: 'p', command: 'print' });
    await expectCleanedUp(page, 3);
  });

  test('print is answered with an error where it cannot be done: no print() in the window, a workbook, a closed file', async ({ page }) => {
    const { buildXlsx, worksheetXml } = await import('../test/fixtures');
    await startBridge(page);
    await send(page, { type: 'print', id: 'none' });
    expect(await answer(page, 'none')).toMatchObject({ type: 'error', code: 'not-loaded' });
    await send(page, { type: 'load', id: 'l', data: base64(await pagesDocx(4, A4)) });
    await answer(page, 'l');
    // A WebView without window.print.
    await page.evaluate(() => {
      (window as any).print = undefined;
    });
    await send(page, { type: 'print', id: 'p1' });
    expect(await answer(page, 'p1')).toMatchObject({ type: 'error', id: 'p1', code: 'print-unavailable' });
    // Closing the file while its pages are prepared.
    await fakePrint(page, false);
    await send(page, { type: 'print', id: 'p2' });
    await send(page, { type: 'close', id: 'c' });
    expect(await answer(page, 'p2')).toMatchObject({ type: 'error', id: 'p2', code: 'print-failed' });
    expect(await printCalls(page)).toEqual([]);
    await expectCleanedUp(page);
    // A workbook.
    const workbook = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>') }] });
    await send(page, { type: 'load', id: 'x', data: base64(workbook) });
    await answer(page, 'x');
    await send(page, { type: 'print', id: 'p3' });
    expect(await answer(page, 'p3')).toMatchObject({ type: 'error', id: 'p3', code: 'print-unavailable' });
    await send(page, { type: 'print', id: 'p4', dpi: 'high' });
    expect(await answer(page, 'p4')).toMatchObject({ type: 'error', id: 'p4', code: 'bad-message' });
  });
});
