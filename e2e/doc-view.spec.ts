import type { Page } from '@playwright/test';
import { p, r, tbl, tc, tr } from '../test/doc-layout-fixtures';
import { buildDocx } from '../test/docx-fixtures';
import { DOC_STYLES, docSection, docx, field, hugePicturesDocx, linesDocx, link, longDocx, longWordDocx, manyPagesDocx, picture, png, sampleDocx, textBoxRun } from './fixtures/docx';
import { expect, expectResponsive, gotoHarness, test, type Flavour } from './helpers';

// The sample pages are 600 x 400 px with margins of 40 px.
const PAGE_WIDTH = 600;
const PAGE_HEIGHT = 400;
/** Space the view leaves on each side of the page. */
const PAD = 12;

interface MountOptions {
  zoom?: number | 'fit-width' | 'fit-page';
  locale?: string;
  toolbar?: boolean;
  shadow?: 'open' | 'closed';
  transform?: string;
  veto?: boolean;
  noPackage?: boolean;
}

async function mountDoc(page: Page, bytes: Uint8Array, options: MountOptions = {}, query = ''): Promise<void> {
  await page.goto(`/e2e/pages/doc.html${query}`);
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountDoc(input, opts), [Array.from(bytes), options] as [number[], MountOptions]);
}

const laidOut = (page: Page) => page.waitForFunction(() => (window as any).__docView.layoutDone === true, undefined, { timeout: 60_000 });
const state = (page: Page) => page.evaluate(() => (window as any).__docView.getViewState());

/** Lets the view bring the screen up to date: its updates run in animation frames. */
const settled = (page: Page) => page.evaluate(() => new Promise<void>((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve()))));

interface CanvasStats {
  width: number;
  height: number;
  cssWidth: number;
  cssHeight: number;
  dpr: number;
  white: number;
  dark: number;
  red: number;
  blue: number;
  green: number;
  grey: number;
  /** Mean horizontal position of the red and of the blue pixels, as shares of the width. */
  redX: number;
  blueX: number;
  corner: number[];
}

/** Counts the pixels of the canvas of a page (counted from 1) by colour. */
function canvasStats(page: Page, number: number): Promise<CanvasStats> {
  return page.evaluate((n) => {
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    const canvas = root.querySelector(`.bym-doc-page[data-page="${n}"] canvas`) as HTMLCanvasElement | null;
    if (!canvas) throw new Error(`page ${n} has no canvas`);
    const data = canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height).data;
    const out = { white: 0, dark: 0, red: 0, blue: 0, green: 0, grey: 0, redX: 0, blueX: 0 };
    for (let i = 0; i < data.length; i += 4) {
      const red = data[i];
      const green = data[i + 1];
      const blue = data[i + 2];
      const x = (i / 4) % canvas.width;
      if (red >= 250 && green >= 250 && blue >= 250) out.white++;
      else if (red < 80 && green < 80 && blue < 80) out.dark++;
      else if (red > 180 && green < 60 && blue < 60) {
        out.red++;
        out.redX += x;
      } else if (blue > 180 && red < 60 && green < 60) {
        out.blue++;
        out.blueX += x;
      } else if (green > 140 && red < 60 && blue < 60) out.green++;
      else if (red === 241 && green === 243 && blue === 244) out.grey++;
    }
    const rect = canvas.getBoundingClientRect();
    return {
      ...out,
      redX: out.red ? out.redX / out.red / canvas.width : -1,
      blueX: out.blue ? out.blueX / out.blue / canvas.width : -1,
      width: canvas.width,
      height: canvas.height,
      cssWidth: rect.width,
      cssHeight: rect.height,
      dpr: window.devicePixelRatio,
      corner: Array.from(data.slice(0, 4)),
    };
  }, number);
}

/** Text the view offers assistive technology for a page (counted from 1). */
const pageText = (page: Page, number: number) =>
  page.evaluate((n) => {
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    const element = root.querySelector(`.bym-doc-page[data-page="${n}"] .bym-doc-text`);
    return element ? element.textContent : null;
  }, number);

const canvasCount = (page: Page) => page.evaluate(() => ((window as any).__shadow || document).querySelectorAll('canvas').length);

/**
 * Where a piece of text is on screen: the document is laid out again in the
 * page with the engine's own measurer, and the text item that starts with
 * `start` is mapped through the rectangle of its page.
 */
function textBox(page: Page, start: string): Promise<{ x: number; y: number; width: number; height: number; page: number }> {
  return page.evaluate(async (prefix) => {
    const api = (window as any).BymOffice;
    const doc = await (window as any).__docFile.document();
    const measurer = api.createCanvasMeasurer({ fonts: doc.fonts });
    const layout = new api.DocLayout(doc, measurer);
    layout.layout();
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    for (let index = 0; index < layout.pages.length; index++) {
      const laid = layout.pages[index];
      for (const item of laid.items) {
        if (item.type !== 'text' || item.text.indexOf(prefix) !== 0) continue;
        const element = root.querySelector(`.bym-doc-page[data-page="${index + 1}"]`);
        if (!element) continue;
        const rect = element.getBoundingClientRect();
        const sx = rect.width / laid.width;
        const sy = rect.height / laid.height;
        const metrics = measurer.metrics(item.font);
        return {
          x: rect.left + item.x * sx,
          y: rect.top + (item.y - metrics.ascent) * sy,
          width: measurer.width(item.font, prefix) * sx,
          height: (metrics.ascent + metrics.descent) * sy,
          page: index + 1,
        };
      }
    }
    throw new Error(`no text on a mounted page starts with "${prefix}"`);
  }, start);
}

/** Fires a copy event at the view's input the way the browser does for Ctrl+C, and returns what ended up on the clipboard. */
function copyEvent(page: Page): Promise<{ prevented: boolean; text: string }> {
  return page.evaluate(() => {
    const data = new DataTransfer();
    let event: Event | undefined;
    try {
      event = new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true });
    } catch {
      event = undefined;
    }
    // Some browsers give a scripted clipboard event a clipboard of their own instead of the one passed in.
    if (!event || (event as ClipboardEvent).clipboardData !== data) {
      event = new Event('copy', { bubbles: true, cancelable: true });
      Object.defineProperty(event, 'clipboardData', { value: data });
    }
    const root: ParentNode = (window as any).__shadow || document.getElementById('host')!;
    root.querySelector('.bym-doc-input')!.dispatchEvent(event);
    return { prevented: event.defaultPrevented, text: data.getData('text/plain') };
  });
}

const SAMPLE_TEXT_START = 'Document title\nSee the example site or do not open.\nCell A\tCell B\nCell C\tCell D\n\nParagraph 1 of the sample document.\nParagraph 2 of the sample document.';

test.describe('document view', () => {
  test('the first page is painted: paper, text, table borders and the picture', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    // The picture is decoded after the first paint.
    await expect.poll(async () => (await canvasStats(page, 1)).red).toBeGreaterThan(1000);
    const stats = await canvasStats(page, 1);
    expect(stats.corner).toEqual([255, 255, 255, 255]);
    expect(stats.white).toBeGreaterThan(stats.width * stats.height * 0.7);
    expect(stats.dark, 'text pixels').toBeGreaterThan(500);
    expect(stats.green, 'table borders').toBeGreaterThan(500);
    expect(stats.blue, 'right half of the picture and the links').toBeGreaterThan(1000);
    // The picture is not mirrored: red left of blue.
    expect(stats.redX).toBeLessThan(stats.blueX);
    // 120 x 60 px on the page, half of it red.
    const scale = stats.width / PAGE_WIDTH;
    expect(stats.red).toBeGreaterThan(60 * 60 * scale * scale * 0.9);
    expect(stats.red).toBeLessThan(60 * 60 * scale * scale * 1.1);

    const text = await pageText(page, 1);
    expect(text).toContain('Sample header');
    expect(text).toContain('Document title');
    expect(text).toContain('Cell A');
    // The footer shows the number of pages once the whole document is laid out.
    expect(text).toContain('Page 1 of 3');
    expect((await state(page)).pageCount).toBe(3);
  });

  test('the canvas has one pixel per device pixel at every zoom', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    for (const zoom of [1, 1.5, 0.75]) {
      await page.evaluate((z) => (window as any).__docView.setZoom(z), zoom);
      await settled(page);
      const stats = await canvasStats(page, 1);
      expect((await state(page)).zoom).toBe(zoom);
      expect(stats.cssWidth).toBeCloseTo(PAGE_WIDTH * zoom, 1);
      expect(stats.cssHeight).toBeCloseTo(PAGE_HEIGHT * zoom, 1);
      // No blur: the backing store is the size on screen times the pixel ratio.
      expect(stats.width).toBe(Math.round(PAGE_WIDTH * zoom * stats.dpr));
      expect(stats.height).toBe(Math.round(PAGE_HEIGHT * zoom * stats.dpr));
      expect(stats.dark).toBeGreaterThan(300);
    }
  });

  test('a very large zoom is drawn at a lower resolution instead of using unbounded memory', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 5 });
    await laidOut(page);
    const stats = await canvasStats(page, 1);
    expect(stats.cssWidth).toBeCloseTo(PAGE_WIDTH * 5, 0);
    expect(stats.width * stats.height).toBeLessThanOrEqual(12_100_000);
    expect(stats.dark).toBeGreaterThan(1000);
    // Out of range values are brought into range.
    await page.evaluate(() => (window as any).__docView.setZoom(1000));
    expect((await state(page)).zoom).toBe(5);
    await page.evaluate(() => (window as any).__docView.setZoom(0.0001));
    expect((await state(page)).zoom).toBe(0.1);
    const errors = await page.evaluate(() => {
      const names: string[] = [];
      for (const value of [0, -1, NaN, 'wide', undefined]) {
        try {
          (window as any).__docView.setZoom(value);
          names.push('ok');
        } catch (error) {
          names.push((error as Error).name);
        }
      }
      return names;
    });
    expect(errors).toEqual(['RangeError', 'RangeError', 'RangeError', 'RangeError', 'RangeError']);
  });

  test('a document of 500 pages keeps a handful of canvases, wherever it is scrolled', async ({ page }) => {
    test.setTimeout(90_000);
    await mountDoc(page, await manyPagesDocx(500), { zoom: 1 });
    await laidOut(page);
    expect((await state(page)).pageCount).toBe(500);
    expect(await canvasCount(page)).toBeLessThanOrEqual(6);

    const pixels = () =>
      page.evaluate(() => {
        let sum = 0;
        document.querySelectorAll('#host canvas').forEach((canvas) => (sum += (canvas as HTMLCanvasElement).width * (canvas as HTMLCanvasElement).height));
        return sum;
      });
    const dpr = await page.evaluate(() => window.devicePixelRatio);
    const onePage = PAGE_WIDTH * PAGE_HEIGHT * dpr * dpr;

    for (const target of [250, 500, 37, 1]) {
      await page.evaluate((n) => (window as any).__docView.goToPage(n), target);
      await settled(page);
      const now = await state(page);
      expect(now.mounted).toContain(target - 1);
      expect(now.mounted.length).toBeLessThanOrEqual(6);
      expect(await canvasCount(page)).toBe(now.mounted.length);
      expect(await pixels(), 'memory of the canvases').toBeLessThanOrEqual(onePage * 6 * 1.01);
      expect(await pageText(page, target)).toBe(`Page ${target}`);
      expect((await canvasStats(page, target)).dark).toBeGreaterThan(20);
      await expect(page.locator('#host .bym-doc-pages')).toHaveText(`Page ${now.page} / 500`);
    }
    expect((await state(page)).page).toBe(1);

    // Scrolling by hand brings the next pages in and lets the earlier ones go.
    await page.evaluate(() => {
      document.querySelector('#host .bym-doc-scroller')!.scrollTop = 20 * 412;
    });
    await expect.poll(async () => (await state(page)).mounted).toContain(20);
    const scrolled = await state(page);
    expect(scrolled.mounted).not.toContain(0);
    expect(scrolled.mounted.length).toBeLessThanOrEqual(6);
    await expect.poll(async () => (await canvasStats(page, 21)).dark).toBeGreaterThan(20);
  });

  test('zooming keeps the page in view; at the smallest zoom the number of canvases stays bounded', async ({ page }) => {
    await mountDoc(page, await manyPagesDocx(120), { zoom: 1 });
    await laidOut(page);
    await page.evaluate(() => (window as any).__docView.goToPage(57));
    for (const zoom of [0.5, 2, 'fit-width', 1.25, 'fit-page', 1] as const) {
      await page.evaluate((z) => (window as any).__docView.setZoom(z), zoom);
      await settled(page);
      expect((await state(page)).mounted, String(zoom)).toContain(56);
      expect(await pageText(page, 57)).toBe('Page 57');
    }
    expect((await state(page)).page).toBe(57);

    await page.evaluate(() => (window as any).__docView.setZoom(0.1));
    await settled(page);
    const small = await state(page);
    // Pages are 40 px high now: many are in sight, and every one of them is small.
    expect(small.mounted.length).toBeGreaterThanOrEqual(8);
    expect(small.mounted.length).toBeLessThanOrEqual(24);
    expect(await canvasCount(page)).toBe(small.mounted.length);
    const stats = await canvasStats(page, small.mounted[1] + 1);
    expect(stats.cssWidth).toBeCloseTo(PAGE_WIDTH * 0.1, 0);
    expect(stats.width).toBe(Math.round(PAGE_WIDTH * 0.1 * stats.dpr));
  });

  test('a document taller than an element can be is still scrolled from its first page to its last', async ({ page }) => {
    test.setTimeout(120_000);
    // 4,100 pages of 2,012 px at 500%: more than 8 million px.
    await mountDoc(page, await manyPagesDocx(4100), { zoom: 5 });
    await laidOut(page);
    expect((await state(page)).pageCount).toBe(4100);
    const strip = await page.evaluate(() => {
      const scroller = document.querySelector('#host .bym-doc-scroller')!;
      return { height: scroller.scrollHeight, strip: (scroller.firstElementChild as HTMLElement).offsetHeight };
    });
    expect(strip.strip).toBe(8_000_000);
    expect(strip.height).toBeLessThanOrEqual(8_000_100);

    for (const target of [4100, 2050, 3, 1]) {
      await page.evaluate((n) => (window as any).__docView.goToPage(n), target);
      await settled(page);
      const now = await state(page);
      expect(now.page).toBe(target);
      expect(now.mounted).toContain(target - 1);
      expect(now.mounted.length).toBeLessThanOrEqual(4);
      expect(await pageText(page, target)).toBe(`Page ${target}`);
      // The page is where the view says it is: its top is at the top of the visible area.
      const box = (await page.locator(`#host .bym-doc-page[data-page="${target}"]`).boundingBox())!;
      const view = (await page.locator('#host .bym-doc-scroller').boundingBox())!;
      expect(box.y - view.y).toBeLessThanOrEqual(12.5);
      expect(box.y - view.y).toBeGreaterThanOrEqual(0);
      await expect(page.locator('#host .bym-doc-pages')).toHaveText(`Page ${target} / 4100`);
    }
    // The scroll bar reaches the end too.
    await page.evaluate(() => {
      const scroller = document.querySelector('#host .bym-doc-scroller')!;
      scroller.scrollTop = scroller.scrollHeight;
    });
    await expect.poll(async () => (await state(page)).page).toBe(4100);
    const last = (await page.locator('#host .bym-doc-page[data-page="4100"]').boundingBox())!;
    const area = (await page.locator('#host .bym-doc-scroller').boundingBox())!;
    // The last page ends where the visible area ends, with the space below it.
    expect(Math.abs(last.y + last.height + 12 - (area.y + area.height))).toBeLessThan(2);
    await page.evaluate(() => (window as any).__docView.focus());
    await page.keyboard.press('Home');
    expect((await state(page)).scrollY).toBe(0);
    await page.keyboard.press('PageDown');
    const paged = await state(page);
    expect(paged.scrollY).toBeGreaterThan(paged.height - 45);
    expect(paged.scrollY).toBeLessThan(paged.height - 35);
  });

  test('fonts that arrive later, and relayout(), lay the document out again and keep the page in view', async ({ page }) => {
    await mountDoc(page, await manyPagesDocx(60), { zoom: 1 });
    await laidOut(page);
    await page.evaluate(() => (window as any).__docView.goToPage(31));
    await settled(page);
    const mark = () =>
      page.evaluate(() => {
        const element = document.querySelector('#host .bym-doc-page[data-page="31"]') as HTMLElement & { __seen?: boolean };
        const seen = element.__seen === true;
        element.__seen = true;
        return seen;
      });
    expect(await mark()).toBe(false);
    expect(await mark()).toBe(true);

    await page.evaluate(() => (window as any).__docView.relayout());
    await laidOut(page);
    await settled(page);
    // The pages were built anew, and the same page is in view.
    expect(await mark()).toBe(false);
    const again = await state(page);
    expect([again.page, again.pageCount]).toEqual([31, 60]);
    expect((await canvasStats(page, 31)).dark).toBeGreaterThan(20);

    // What the browser reports when web fonts finished loading.
    await page.evaluate(() => {
      document.fonts.dispatchEvent(new Event('loadingdone'));
      document.fonts.dispatchEvent(new Event('loadingdone'));
    });
    await expect.poll(mark).toBe(false);
    await laidOut(page);
    await settled(page);
    expect((await state(page)).page).toBe(31);
    expect(await pageText(page, 31)).toBe('Page 31');

    // A page that keeps reporting fonts cannot make the view lay out forever.
    const runs = await page.evaluate(async () => {
      const view = (window as any).__docView;
      let count = 0;
      const relayout = view.relayout.bind(view);
      view.relayout = () => {
        count++;
        relayout();
      };
      for (let i = 0; i < 30; i++) {
        document.fonts.dispatchEvent(new Event('loadingdone'));
        await new Promise((resolve) => setTimeout(resolve, 70));
      }
      return count;
    });
    expect(runs).toBeLessThanOrEqual(12);
    expect(runs).toBeGreaterThanOrEqual(1);
    await laidOut(page);
    expect((await state(page)).pageCount).toBe(60);
  });

  test('layout is progressive: the first pages show at once, the rest follows and the page keeps responding', async ({ page }) => {
    test.setTimeout(120_000);
    const footer = { f1: p(r('Page ') + field(' PAGE ', '1') + r(' of ') + field(' NUMPAGES ', '1')) };
    let body = '';
    for (let i = 1; i <= 3000; i++) body += p(`${i}. The quick brown fox jumps over the lazy dog and keeps running along the river bank until the evening.`);
    const bytes = await buildDocx({ styles: DOC_STYLES, body: body + docSection('<w:footerReference w:type="default" r:id="f1"/>'), footers: footer });
    await page.goto('/e2e/pages/doc.html');
    const first = await page.evaluate(async (input) => {
      await (window as any).__mountDoc(input, { zoom: 1 });
      const view = (window as any).__docView;
      const host = document.getElementById('host')!;
      return {
        done: view.layoutDone,
        count: view.pageCount,
        label: host.querySelector('.bym-doc-pages')!.textContent,
        canvases: host.querySelectorAll('canvas').length,
        text: host.querySelector('.bym-doc-page[data-page="1"] .bym-doc-text')!.textContent,
      };
    }, Array.from(bytes));
    expect(first.done).toBe(false);
    expect(first.count).toBeGreaterThanOrEqual(1);
    expect(first.canvases).toBeGreaterThanOrEqual(1);
    // More pages are on their way.
    expect(first.label).toBe(`Page 1 / ${first.count}+`);
    expect((await canvasStats(page, 1)).dark).toBeGreaterThan(200);
    await expectResponsive(page);

    await laidOut(page);
    const end = await state(page);
    expect(end.pageCount).toBeGreaterThan(first.count);
    expect(end.pageCount).toBeGreaterThan(300);
    await expect(page.locator('#host .bym-doc-pages')).toHaveText(`Page 1 / ${end.pageCount}`);
    // The first page was handed out before the total was known; it has been brought up to date.
    expect(await pageText(page, 1)).toContain(`Page 1 of ${end.pageCount}`);
    const counts = await page.evaluate(() => (window as any).__pages.map((entry: number[]) => entry[1]));
    expect(counts[counts.length - 1]).toBe(end.pageCount);
    expect(counts.length).toBeGreaterThanOrEqual(2);
    expect(await canvasCount(page)).toBeLessThanOrEqual(6);

    // A page asked for before it exists is shown as soon as it does.
    await page.evaluate((input) => (window as any).__mountDoc(input, { zoom: 1 }).then(() => (window as any).__docView.goToPage(200)), Array.from(bytes));
    await expect.poll(async () => (await state(page)).page, { timeout: 60_000 }).toBe(200);
  });

  test('performance: 5,000 paragraphs show their first page quickly', async ({ page }) => {
    test.setTimeout(180_000);
    const bytes = await longDocx(5000);
    await page.goto('/e2e/pages/doc.html');
    const result = await page.evaluate(async (input) => {
      await (window as any).__mountDoc(input, { zoom: 1 });
      const painted = (window as any).__viewMs as number;
      const started = performance.now();
      // The longest time the page could not run a timer while the rest was laid out.
      let longest = 0;
      let last = performance.now();
      await new Promise<void>((resolve) => {
        const tick = () => {
          const now = performance.now();
          longest = Math.max(longest, now - last);
          last = now;
          if ((window as any).__docView.layoutDone) resolve();
          else setTimeout(tick, 0);
        };
        tick();
      });
      return { read: (window as any).__readMs as number, painted, rest: performance.now() - started, longest, pages: (window as any).__docView.pageCount };
    }, Array.from(bytes));
    console.log(
      `document of 5,000 paragraphs [${test.info().project.name}]: model ${result.read.toFixed(0)} ms, first paint ${result.painted.toFixed(0)} ms, ` +
        `all ${result.pages} pages after ${result.rest.toFixed(0)} ms more, longest pause ${result.longest.toFixed(0)} ms`,
    );
    expect(result.pages).toBeGreaterThan(500);
    // Generous: the first slice of layout is 30 ms, then one paint.
    expect(result.painted).toBeLessThan(1500);
    expect(result.longest).toBeLessThan(1000);
    expect(result.rest).toBeLessThan(120_000);
    expect(await canvasCount(page)).toBeLessThanOrEqual(6);
  });

  test('the page indicator and onPageChange follow the page in view, in the language of the view', async ({ page }) => {
    await mountDoc(page, await manyPagesDocx(12), { zoom: 1, locale: 'tr' });
    await laidOut(page);
    const host = page.locator('#host');
    await expect(host.locator('.bym-doc-pages')).toHaveText('Sayfa 1 / 12');
    await expect(host.locator('.bym-doc')).toHaveAttribute('aria-label', 'Belge');
    await expect(host.locator('.bym-doc-page[data-page="1"]')).toHaveAttribute('aria-label', 'Sayfa 1');
    await expect(host.locator('.bym-doc-zoom-in')).toHaveAttribute('aria-label', 'Yakınlaştır');
    await expect(host.locator('.bym-doc-zoom-out')).toHaveAttribute('aria-label', 'Uzaklaştır');
    await expect(host.locator('.bym-doc-fit')).toHaveAttribute('aria-label', 'Genişliğe sığdır');

    await page.evaluate(() => (window as any).__docView.goToPage(3));
    await expect(host.locator('.bym-doc-pages')).toHaveText('Sayfa 3 / 12');
    await expect(host.locator('[aria-live]')).toHaveText('Sayfa 3 / 12');
    expect(await page.evaluate(() => (window as any).__docView.currentPage)).toBe(3);
    const reported = await page.evaluate(() => (window as any).__pages);
    expect(reported[reported.length - 1]).toEqual([3, 12]);

    await page.evaluate(() => (window as any).__docView.setLocale('en'));
    await expect(host.locator('.bym-doc-pages')).toHaveText('Page 3 / 12');
    await expect(host.locator('.bym-doc')).toHaveAttribute('aria-label', 'Document');
    await expect(host.locator('.bym-doc-page[data-page="3"]')).toHaveAttribute('aria-label', 'Page 3');
    await expect(host.locator('.bym-doc-zoom-in')).toHaveAttribute('aria-label', 'Zoom in');

    // Beyond the document: the last page. Not a page number: an error.
    await page.evaluate(() => (window as any).__docView.goToPage(999));
    await expect(host.locator('.bym-doc-pages')).toHaveText('Page 12 / 12');
    const errors = await page.evaluate(() => {
      const names: string[] = [];
      for (const value of [0, -3, NaN, '2']) {
        try {
          (window as any).__docView.goToPage(value);
          names.push('ok');
        } catch (error) {
          names.push((error as Error).name);
        }
      }
      return names;
    });
    expect(errors).toEqual(['RangeError', 'RangeError', 'RangeError', 'RangeError']);
  });

  test('the bar zooms in steps, resets to 100% and fits the width', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    const host = page.locator('#host');
    const level = host.locator('.bym-doc-zoom-level');
    await expect(level).toHaveText('100%');
    await host.locator('.bym-doc-zoom-in').click();
    await expect(level).toHaveText('110%');
    await host.locator('.bym-doc-zoom-in').click();
    await expect(level).toHaveText('125%');
    expect((await state(page)).zoomMode).toBe('manual');
    await host.locator('.bym-doc-zoom-out').click();
    await host.locator('.bym-doc-zoom-out').click();
    await host.locator('.bym-doc-zoom-out').click();
    await expect(level).toHaveText('90%');
    await level.click();
    await expect(level).toHaveText('100%');

    await host.locator('.bym-doc-fit').click();
    await expect(host.locator('.bym-doc-fit')).toHaveAttribute('aria-pressed', 'true');
    const fitted = await state(page);
    expect(fitted.zoomMode).toBe('fit-width');
    expect(fitted.zoom).toBeCloseTo((fitted.width - 2 * PAD) / PAGE_WIDTH, 5);
    await host.locator('.bym-doc-zoom-in').click();
    await expect(host.locator('.bym-doc-fit')).toHaveAttribute('aria-pressed', 'false');

    // At the ends of the range the buttons are off.
    await page.evaluate(() => (window as any).__docView.setZoom(5));
    await expect(host.locator('.bym-doc-zoom-in')).toBeDisabled();
    await page.evaluate(() => (window as any).__docView.setZoom(0.1));
    await expect(host.locator('.bym-doc-zoom-out')).toBeDisabled();
    await expect(host.locator('.bym-doc-zoom-in')).toBeEnabled();
  });

  test('toolbar: false leaves only the pages', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { toolbar: false });
    await laidOut(page);
    await expect(page.locator('#host .bym-doc-bar')).toHaveCount(0);
    await expect(page.locator('#host button')).toHaveCount(0);
    expect((await canvasStats(page, 1)).dark).toBeGreaterThan(300);
    // The view still zooms from the keyboard and the API.
    await page.evaluate(() => (window as any).__docView.zoomIn());
    expect((await state(page)).zoomMode).toBe('manual');
  });

  test('keys zoom with Ctrl or Cmd and move through the document without it', async ({ page }) => {
    await mountDoc(page, await manyPagesDocx(8), { zoom: 1 });
    await laidOut(page);
    await page.evaluate(() => (window as any).__docView.focus());
    await expect(page.locator('#host .bym-doc-input')).toBeFocused();

    await page.keyboard.press('Control+=');
    expect((await state(page)).zoom).toBe(1.1);
    await page.keyboard.press('Meta+-');
    await page.keyboard.press('Control+-');
    expect((await state(page)).zoom).toBe(0.9);
    await page.keyboard.press('Control+0');
    expect((await state(page)).zoom).toBe(1);

    const start = await state(page);
    await page.keyboard.press('PageDown');
    const paged = await state(page);
    expect(paged.scrollY).toBe(start.height - 40);
    await page.keyboard.press('ArrowDown');
    expect((await state(page)).scrollY).toBe(paged.scrollY + 40);
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('PageUp');
    expect((await state(page)).scrollY).toBe(0);
    await page.keyboard.press('Space');
    expect((await state(page)).scrollY).toBe(start.height - 40);
    await page.keyboard.press('Shift+Space');
    expect((await state(page)).scrollY).toBe(0);
    await page.keyboard.press('End');
    expect((await state(page)).page).toBe(8);
    await expect(page.locator('#host .bym-doc-pages')).toHaveText('Page 8 / 8');
    await page.keyboard.press('Home');
    expect((await state(page)).page).toBe(1);
    expect((await state(page)).scrollY).toBe(0);
  });

  test('Tab passes through the view: there is no keyboard trap', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    await page.locator('#before').focus();
    const focused = () =>
      page.evaluate(() => {
        const element = document.activeElement as HTMLElement | null;
        if (!element) return 'none';
        return element.id || element.className.replace('bym-doc-btn ', '');
      });
    // Safari leaves buttons out of the Tab order unless the user asked for them; the order of the rest is what counts.
    const seen: string[] = [];
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press('Tab');
      const now = await focused();
      seen.push(now);
      if (now === 'after') break;
    }
    expect(seen[seen.length - 1]).toBe('after');
    expect(seen.filter((name) => name === 'bym-doc-input' || name === 'bym-doc-link')).toEqual(['bym-doc-input', 'bym-doc-link']);
    expect(seen.length).toBeLessThanOrEqual(7);

    // And back out the other way.
    const back: string[] = [];
    for (let i = 0; i < 12; i++) {
      await page.keyboard.press('Shift+Tab');
      const now = await focused();
      back.push(now);
      if (now === 'before') break;
    }
    expect(back[back.length - 1]).toBe('before');
  });

  test('fit-width follows the width of the container; fit-page shows the whole page', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 'fit-width' });
    await laidOut(page);
    const wide = await state(page);
    expect(wide.zoomMode).toBe('fit-width');
    expect(wide.zoom).toBeCloseTo((wide.width - 2 * PAD) / PAGE_WIDTH, 5);
    const before = await canvasStats(page, 1);
    expect(before.cssWidth).toBeCloseTo(wide.width - 2 * PAD, 0);
    expect(before.width).toBe(Math.round(PAGE_WIDTH * wide.zoom * before.dpr));

    await page.evaluate(() => {
      document.getElementById('host')!.style.width = '300px';
    });
    await expect.poll(async () => (await state(page)).width).toBeLessThanOrEqual(300);
    await settled(page);
    const narrow = await state(page);
    expect(narrow.zoom).toBeCloseTo((narrow.width - 2 * PAD) / PAGE_WIDTH, 5);
    expect(narrow.zoom).toBeLessThan(wide.zoom);
    const after = await canvasStats(page, 1);
    expect(after.cssWidth).toBeCloseTo(narrow.width - 2 * PAD, 0);
    expect(after.width).toBe(Math.round(PAGE_WIDTH * narrow.zoom * after.dpr));
    expect(after.dark).toBeGreaterThan(50);
    // Nothing to scroll sideways.
    expect(await page.evaluate(() => { const s = document.querySelector('#host .bym-doc-scroller')!; return s.scrollWidth <= s.clientWidth; })).toBe(true);

    await page.evaluate(() => (window as any).__docView.setZoom('fit-page'));
    await settled(page);
    const whole = await state(page);
    expect(whole.zoomMode).toBe('fit-page');
    expect(whole.zoom).toBeCloseTo(Math.min((whole.width - 2 * PAD) / PAGE_WIDTH, (whole.height - 2 * PAD) / PAGE_HEIGHT), 5);
    const box = await page.locator('#host .bym-doc-page[data-page="1"]').boundingBox();
    const view = await page.locator('#host .bym-doc-scroller').boundingBox();
    expect(box!.y).toBeGreaterThanOrEqual(view!.y);
    expect(box!.y + box!.height).toBeLessThanOrEqual(view!.y + view!.height + 0.5);
  });

  test('without a zoom the page is shown at 100% where it fits and smaller where it does not', async ({ page }) => {
    await mountDoc(page, await sampleDocx());
    await laidOut(page);
    const now = await state(page);
    expect(now.zoomMode).toBe('auto');
    expect(now.zoom).toBeCloseTo(Math.min(1, (now.width - 2 * PAD) / PAGE_WIDTH), 5);
    expect(await page.evaluate(() => { const s = document.querySelector('#host .bym-doc-scroller')!; return s.scrollWidth <= s.clientWidth; })).toBe(true);
    await page.evaluate(() => {
      document.getElementById('host')!.style.width = '260px';
    });
    await expect.poll(async () => (await state(page)).zoom).toBeLessThan(0.4);
    // Wider again: back up, but never above 100%.
    await page.evaluate(() => {
      document.getElementById('host')!.style.width = '2000px';
      document.getElementById('host')!.style.maxWidth = 'none';
    });
    await expect.poll(async () => (await state(page)).zoom).toBe(1);
  });

  test('Ctrl + wheel zooms around the pointer', async ({ page, isMobile }) => {
    test.skip(isMobile, 'mouse wheel events are not available on the mobile profiles');
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    const target = await textBox(page, 'Paragraph 5 of');
    await page.mouse.move(target.x + 5, target.y + 5);
    await page.keyboard.down('Control');
    await page.mouse.wheel(0, -120);
    await expect.poll(async () => (await state(page)).zoom).toBeGreaterThan(1.2);
    await page.keyboard.up('Control');
    await settled(page);
    expect((await state(page)).zoomMode).toBe('manual');
    // The line under the pointer still is. (Sideways the page stays centred while it is narrower than the view.)
    const moved = await textBox(page, 'Paragraph 5 of');
    expect(Math.abs(moved.y + (5 * moved.height) / target.height - (target.y + 5))).toBeLessThan(4);
    // Once the page is wider than the view, the point under the pointer keeps its place sideways too.
    await page.evaluate(() => {
      (window as any).__docView.setZoom(2);
      (window as any).__docView.goToPage(1);
    });
    await settled(page);
    const wide = await textBox(page, 'Cell B');
    await page.mouse.move(wide.x + 30, wide.y + 10);
    await page.keyboard.down('Control');
    await page.mouse.wheel(0, -60);
    await expect.poll(async () => (await state(page)).zoom).toBeGreaterThan(2.1);
    await page.keyboard.up('Control');
    await settled(page);
    const ratio = (await state(page)).zoom / 2;
    const wider = await textBox(page, 'Cell B');
    expect(Math.abs(wider.x + 30 * ratio - (wide.x + 30))).toBeLessThan(4);
    expect(Math.abs(wider.y + 10 * ratio - (wide.y + 10))).toBeLessThan(4);
    // Without the key the wheel scrolls and the zoom stays.
    const zoom = (await state(page)).zoom;
    await page.mouse.wheel(0, 150);
    await expect.poll(async () => (await state(page)).scrollY).toBeGreaterThan(100);
    expect((await state(page)).zoom).toBe(zoom);
  });

  test('two fingers zoom', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    // Not every browser lets a script construct touch events; the view only reads the touch points.
    const prevented = await page.evaluate(() => {
      const scroller = document.querySelector('#host .bym-doc-scroller')!;
      const rect = scroller.getBoundingClientRect();
      const y = rect.top + 100;
      const centre = rect.left + rect.width / 2;
      const fire = (type: string, xs: number[]) => {
        const event = new Event(type, { bubbles: true, cancelable: true });
        Object.defineProperty(event, 'touches', { value: xs.map((x) => ({ clientX: x, clientY: y })) });
        scroller.dispatchEvent(event);
        return event.defaultPrevented;
      };
      fire('touchstart', [centre - 40, centre + 40]);
      const moved = fire('touchmove', [centre - 60, centre + 60]);
      fire('touchend', []);
      // One finger afterwards scrolls: it is not a pinch.
      fire('touchstart', [centre]);
      const one = fire('touchmove', [centre + 50]);
      return [moved, one];
    });
    // The browser must not zoom the whole page with the same gesture.
    expect(prevented).toEqual([true, false]);
    const now = await state(page);
    expect(now.zoom).toBeCloseTo(1.5, 5);
    expect(now.zoomMode).toBe('manual');
    await settled(page);
    expect((await canvasStats(page, 1)).cssWidth).toBeCloseTo(PAGE_WIDTH * 1.5, 0);
    expect(await page.evaluate(() => getComputedStyle(document.querySelector('#host .bym-doc-scroller')!).touchAction)).toBe('pan-x pan-y');
  });

  test('a hyperlink calls onLinkClick and opens safe targets in a new window; unsafe ones are not links', async ({ page, isMobile }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    const links = page.locator('#host .bym-doc-link');
    // The second hyperlink of the document points to javascript: and is plain text here.
    await expect(links).toHaveCount(1);
    await expect(links).toHaveAttribute('role', 'link');
    await expect(links).toHaveAttribute('aria-label', 'the example site');
    expect(await links.evaluate((element) => getComputedStyle(element).cursor)).toBe('pointer');
    // The element lies over the words of the link.
    const words = await textBox(page, 'the example site');
    const box = (await links.boundingBox())!;
    expect(Math.abs(box.x - words.x)).toBeLessThan(1.5);
    expect(Math.abs(box.y - words.y)).toBeLessThan(1.5);
    expect(Math.abs(box.width - words.width)).toBeLessThan(2);

    if (isMobile) await links.tap();
    else await links.click();
    expect(await page.evaluate(() => (window as any).__links)).toEqual(['https://example.com/docs?a=1']);
    expect(await page.evaluate(() => (window as any).__opened)).toEqual([{ url: 'https://example.com/docs?a=1', target: '_blank', features: 'noopener,noreferrer' }]);

    // From the keyboard.
    await links.focus();
    await page.keyboard.press('Enter');
    expect(await page.evaluate(() => (window as any).__links.length)).toBe(2);
    expect(await page.evaluate(() => (window as any).__opened.length)).toBe(2);

    // The words of the unsafe link do nothing.
    const unsafe = await textBox(page, 'do not open');
    if (isMobile) await page.touchscreen.tap(unsafe.x + 10, unsafe.y + unsafe.height / 2);
    else await page.mouse.click(unsafe.x + 10, unsafe.y + unsafe.height / 2);
    expect(await page.evaluate(() => (window as any).__links.length)).toBe(2);
    expect(await page.evaluate(() => (window as any).__opened.length)).toBe(2);
    expect(await page.evaluate(() => (window as any).__hacked)).toBeUndefined();
  });

  test('onLinkClick returning false keeps the link from being opened', async ({ page, isMobile }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1, veto: true });
    await laidOut(page);
    const links = page.locator('#host .bym-doc-link');
    if (isMobile) await links.tap();
    else await links.click();
    expect(await page.evaluate(() => (window as any).__links)).toEqual(['https://example.com/docs?a=1']);
    expect(await page.evaluate(() => (window as any).__opened)).toEqual([]);
  });

  test('select all and copy put the text of the whole document on the clipboard', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    await page.evaluate(() => (window as any).__docView.focus());
    // Nothing selected: the copy leaves the clipboard alone.
    expect(await copyEvent(page)).toEqual({ prevented: true, text: '' });

    const plain = await canvasStats(page, 1);
    await page.keyboard.press('Control+a');
    await settled(page);
    const copied = await copyEvent(page);
    expect(copied.prevented).toBe(true);
    expect(copied.text.startsWith(SAMPLE_TEXT_START)).toBe(true);
    expect(copied.text.endsWith('Paragraph 40 of the sample document.')).toBe(true);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe(copied.text);
    // The selection shows: the paper behind the text is no longer white.
    const selected = await canvasStats(page, 1);
    expect(selected.white).toBeLessThan(plain.white - 2000);

    await page.keyboard.press('Escape');
    await settled(page);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('');
    expect((await canvasStats(page, 1)).white).toBe(plain.white);
  });

  test('Ctrl+A and Ctrl+C reach the real clipboard', async ({ page, context, browserName, isMobile }) => {
    test.skip(browserName !== 'chromium' || isMobile, 'only Chromium lets a test read the system clipboard');
    await context.grantPermissions(['clipboard-read', 'clipboard-write']);
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    const target = await textBox(page, 'Paragraph 3 of');
    await page.mouse.click(target.x + 300, target.y + 5);
    await expect(page.locator('#host .bym-doc-input')).toBeFocused();
    await page.keyboard.press('ControlOrMeta+a');
    await page.keyboard.press('ControlOrMeta+c');
    const text = await page.evaluate(() => navigator.clipboard.readText());
    expect(text.startsWith(SAMPLE_TEXT_START)).toBe(true);
  });

  test('dragging the mouse selects text; a double click selects a word', async ({ page, isMobile }) => {
    test.skip(isMobile, 'a finger scrolls; selecting by dragging is for the mouse');
    await mountDoc(page, await sampleDocx(), { zoom: 1 });
    await laidOut(page);
    const first = await textBox(page, 'Paragraph 2 of');
    const second = await textBox(page, 'Paragraph 4 of');
    const plain = await canvasStats(page, 1);
    // From before the "P" of paragraph 2 to behind the "4" of paragraph 4.
    await page.mouse.move(first.x + 1, first.y + first.height / 2);
    await page.mouse.down();
    await page.mouse.move(first.x + 60, first.y + first.height / 2 + 4, { steps: 3 });
    await page.mouse.move(second.x + second.width - 20, second.y + second.height / 2, { steps: 4 });
    await page.mouse.up();
    await settled(page);
    const text = await page.evaluate(() => (window as any).__docView.getSelectedText());
    expect(text).toBe('Paragraph 2 of the sample document.\nParagraph 3 of the sample document.\nParagraph 4');
    expect((await copyEvent(page)).text).toBe(text);
    expect((await canvasStats(page, 1)).white).toBeLessThan(plain.white - 500);
    await expect(page.locator('#host .bym-doc-input')).toBeFocused();

    // Shift + click moves the end of the selection.
    await page.keyboard.down('Shift');
    await page.mouse.click(first.x + first.width, first.y + first.height / 2);
    await page.keyboard.up('Shift');
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('Paragraph 2 of');

    await page.mouse.dblclick(second.x + 30, second.y + second.height / 2);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('Paragraph');

    // A plain click drops the selection.
    await page.mouse.click(first.x + 400, first.y + first.height / 2);
    await settled(page);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('');
    expect((await canvasStats(page, 1)).white).toBe(plain.white);
    expect(await page.evaluate(() => (window as any).__links.length)).toBe(0);
  });

  test('hit testing is right when the host is scaled with a CSS transform', async ({ page, isMobile }) => {
    test.skip(isMobile, 'selecting by pointer is for the mouse');
    await mountDoc(page, await sampleDocx(), { zoom: 1, transform: 'scale(0.6)' });
    await laidOut(page);
    const word = await textBox(page, 'Paragraph 7 of');
    // On screen the text is 60% of its size.
    const box = (await page.locator('#host .bym-doc-page[data-page="1"]').boundingBox())!;
    expect(box.width).toBeCloseTo(PAGE_WIDTH * 0.6, 0);
    await page.mouse.dblclick(word.x + word.width - 3, word.y + word.height / 2);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('of');
    await page.locator('#host .bym-doc-link').click();
    expect(await page.evaluate(() => (window as any).__links)).toEqual(['https://example.com/docs?a=1']);
    // The transform does not change the zoom: the layout width of the container is what is fitted.
    expect((await state(page)).zoom).toBe(1);
  });

  for (const mode of ['open', 'closed'] as const) {
    test(`works inside a ${mode} shadow root`, async ({ page, isMobile }) => {
      await mountDoc(page, await sampleDocx(), { zoom: 1, shadow: mode });
      await laidOut(page);
      await expect.poll(async () => (await canvasStats(page, 1)).red).toBeGreaterThan(1000);
      expect((await canvasStats(page, 1)).dark).toBeGreaterThan(500);
      // The view's style element lives in the shadow root, not in the page.
      expect(await page.evaluate(() => [(window as any).__shadow.querySelectorAll('style[data-bymoffice-doc]').length, document.querySelectorAll('style[data-bymoffice-doc]').length])).toEqual([1, 0]);
      expect(await page.evaluate(() => getComputedStyle((window as any).__shadow.querySelector('.bym-doc-page')).backgroundColor)).toBe('rgb(255, 255, 255)');
      await page.evaluate(() => (window as any).__docView.focus());
      await page.keyboard.press('Control+a');
      expect((await copyEvent(page)).text.startsWith(SAMPLE_TEXT_START)).toBe(true);
      if (!isMobile) {
        const word = await textBox(page, 'Paragraph 2 of');
        await page.mouse.dblclick(word.x + 20, word.y + word.height / 2);
        expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('Paragraph');
      }
      await page.evaluate(() => (window as any).__docView.destroy());
      expect(await page.evaluate(() => (window as any).__shadow.querySelectorAll('style, .bym-doc, canvas').length)).toBe(0);
    });
  }

  test('rules of the host page for bare elements do not reach into the view', async ({ page }) => {
    await mountDoc(page, await sampleDocx(), { zoom: 1 }, '?hostile');
    await laidOut(page);
    const measured = await page.evaluate(() => {
      const host = document.getElementById('host')!;
      const canvas = host.querySelector('canvas')!;
      const slot = host.querySelector('.bym-doc-page')!;
      const input = host.querySelector('.bym-doc-input')!;
      const button = host.querySelector('.bym-doc-zoom-in')!;
      const style = getComputedStyle(canvas);
      return {
        canvas: [canvas.getBoundingClientRect().width, style.borderTopWidth, style.marginLeft],
        slot: slot.getBoundingClientRect().width,
        input: [input.getBoundingClientRect().width, getComputedStyle(input).opacity],
        button: [button.getBoundingClientRect().width <= 44, button.getBoundingClientRect().height <= 40, getComputedStyle(button).textTransform],
      };
    });
    expect(measured.canvas).toEqual([PAGE_WIDTH, '0px', '0px']);
    expect(measured.slot).toBe(PAGE_WIDTH);
    expect(measured.input).toEqual([1, '0']);
    expect(measured.button).toEqual([true, true, 'none']);
    expect((await canvasStats(page, 1)).dark).toBeGreaterThan(500);
  });

  test('text from the file never becomes markup', async ({ page }) => {
    const hostile = '<img src=x onerror="window.__hacked=1"><script>window.__hacked=2</script>';
    const bytes = await docx(p(r(hostile)) + p(`<w:hyperlink r:id="rIdLink">${r(hostile)}</w:hyperlink>`) + p(picture(100, 40)), {
      rels: [
        ['rIdLink', 'hyperlink', 'https://example.com/&quot;&gt;&lt;img src=x onerror=window.__hacked=3&gt;', true],
        ['rIdImg', 'image', 'media/missing.png'],
      ],
    });
    await mountDoc(page, bytes, { zoom: 1 });
    await laidOut(page);
    await settled(page);
    expect(await pageText(page, 1)).toContain(hostile);
    // The link may be broken over two lines; its pieces together carry its words, as an attribute value.
    expect((await page.locator('#host .bym-doc-link').evaluateAll((elements) => elements.map((element) => element.getAttribute('aria-label')))).join('')).toBe(hostile);
    expect(await page.evaluate(() => document.querySelectorAll('#host img, #host script, #host a, #host iframe').length)).toBe(0);
    expect(await page.evaluate(() => (window as any).__hacked)).toBeUndefined();
  });

  test('pictures: no package or a missing part gives a placeholder; vector pictures are never handed to the browser', async ({ page }) => {
    const requests: string[] = [];
    page.on('request', (request) => requests.push(request.url()));
    await mountDoc(page, await sampleDocx(), { zoom: 1, noPackage: true });
    await laidOut(page);
    await settled(page);
    const without = await canvasStats(page, 1);
    expect(without.red).toBe(0);
    expect(without.grey, 'placeholder').toBeGreaterThan(120 * 60 * 0.8);

    // An SVG that would fetch from the network if it were rendered, under a name that claims to be a PNG.
    const origin = new URL(page.url()).origin;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20"><image href="${origin}/e2e/pages/tracked.png" width="40" height="20"/><rect width="40" height="20" fill="#dc0000"/></svg>`;
    const bytes = await docx(p('Before') + p(picture(120, 60)) + p(picture(120, 60, 'rIdGone')) + p('After'), {
      rels: [
        ['rIdImg', 'image', 'media/image1.png'],
        ['rIdGone', 'image', 'media/gone.png'],
      ],
      parts: [['word/media/image1.png', svg]],
    });
    await page.evaluate((input) => (window as any).__mountDoc(input, { zoom: 1 }), Array.from(bytes));
    await laidOut(page);
    await page.waitForTimeout(300);
    const vector = await canvasStats(page, 1);
    expect(vector.red).toBe(0);
    expect(vector.grey).toBeGreaterThan(2 * 120 * 60 * 0.8);
    expect(requests.filter((url) => url.includes('tracked'))).toEqual([]);
    // Only the page and the engine were ever requested.
    expect(requests.every((url) => url.startsWith(origin) && (url.includes('/e2e/pages/doc.html') || url.includes('/dist/bymoffice.min.js')))).toBe(true);
    // The placeholders say what they stand for to those who cannot see them.
    expect(await pageText(page, 1)).toContain('Before');
  });

  test('a rotated and mirrored picture is drawn turned', async ({ page }) => {
    // 120 x 60 on the page, red left and blue right; mirrored, blue comes first.
    const bytes = await docx(p(picture(120, 60, 'rIdImg', ' flipH="1"')), { rels: [['rIdImg', 'image', 'media/image1.png']], parts: [['word/media/image1.png', png(40, 20, [220, 0, 0], [0, 0, 220])]] });
    await mountDoc(page, bytes, { zoom: 1 });
    await laidOut(page);
    await expect.poll(async () => (await canvasStats(page, 1)).red).toBeGreaterThan(1000);
    const stats = await canvasStats(page, 1);
    expect(stats.blueX).toBeLessThan(stats.redX);
  });

  test('destroy leaves nothing behind and can be called twice', async ({ page }) => {
    await mountDoc(page, await manyPagesDocx(300), { zoom: 1 });
    // While the layout is still running.
    const left = await page.evaluate(async () => {
      const view = (window as any).__docView;
      const count = view.pageCount;
      view.destroy();
      view.destroy();
      await new Promise((resolve) => setTimeout(resolve, 150));
      // Calls on a destroyed view do nothing.
      view.setZoom(2);
      view.goToPage(2);
      view.selectAll();
      view.relayout();
      return {
        children: document.getElementById('host')!.childNodes.length,
        styles: document.querySelectorAll('style[data-bymoffice-doc]').length,
        canvases: document.querySelectorAll('canvas').length,
        grew: view.pageCount > count + 1,
        reports: (window as any).__pages.length,
      };
    });
    expect(left.children).toBe(0);
    expect(left.styles).toBe(0);
    expect(left.canvases).toBe(0);
    expect(left.grew, 'layout stopped').toBe(false);
    await page.waitForTimeout(100);
    expect(await page.evaluate(() => (window as any).__pages.length)).toBe(left.reports);
    await expectResponsive(page);

    // Two views side by side share one style element; it goes with the last of them.
    const shared = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('host')!;
      const make = async () => {
        const holder = document.createElement('div');
        holder.style.cssText = 'height:200px';
        host.appendChild(holder);
        const file = await api.open(new Uint8Array(input));
        return new api.DocumentView(holder, await file.document(), { package: file.package });
      };
      const a = await make();
      const b = await make();
      const counts = [document.querySelectorAll('style[data-bymoffice-doc]').length];
      a.destroy();
      counts.push(document.querySelectorAll('style[data-bymoffice-doc]').length);
      b.destroy();
      counts.push(document.querySelectorAll('style[data-bymoffice-doc]').length);
      return counts;
    }, Array.from(await sampleDocx()));
    expect(shared).toEqual([1, 1, 0]);
  });

  test('a device pixel ratio of 2 doubles the backing store', async ({ browser, browserName, isMobile }) => {
    test.skip(isMobile, 'the mobile profiles already run at a high pixel ratio');
    const context = await browser.newContext({ deviceScaleFactor: 2, viewport: { width: 900, height: 700 }, baseURL: test.info().project.use.baseURL });
    const page = await context.newPage();
    const problems: string[] = [];
    page.on('pageerror', (error) => problems.push(error.message));
    try {
      await mountDoc(page, await sampleDocx(), { zoom: 1 });
      await laidOut(page);
      const stats = await canvasStats(page, 1);
      expect(stats.dpr).toBe(2);
      expect([stats.width, stats.height]).toEqual([PAGE_WIDTH * 2, PAGE_HEIGHT * 2]);
      expect([stats.cssWidth, stats.cssHeight]).toEqual([PAGE_WIDTH, PAGE_HEIGHT]);
      expect(stats.dark).toBeGreaterThan(2000);
      expect(problems, browserName).toEqual([]);
    } finally {
      await context.close();
    }
  });
});

test.describe('BymOffice.mount with a document', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
  });

  test('shows a docx from bytes, read-only, and saves it as it was', async ({ page }) => {
    const input = await sampleDocx();
    const result = await page.evaluate(async (bytes) => {
      const api = (window as any).BymOffice;
      const pages: number[][] = [];
      const saves: number[][] = [];
      const mounted = await api.mount(document.getElementById('editor'), {
        file: new Uint8Array(bytes),
        locale: 'tr',
        onPageChange: (page: number, count: number) => pages.push([page, count]),
        onSave: (saved: Uint8Array) => saves.push(Array.from(saved)),
      });
      (window as any).__mounted = mounted;
      (window as any).__saves = saves;
      const saved: Uint8Array = await mounted.save();
      return {
        kind: mounted.kind,
        format: mounted.file.format,
        keys: Object.keys(mounted).sort(),
        blocks: mounted.document.body.blocks.length > 40,
        isView: mounted.view instanceof api.DocumentView,
        saved: Array.from(saved),
        pages: pages.length,
      };
    }, Array.from(input));
    expect(result.kind).toBe('document');
    expect(result.format).toBe('docx');
    expect(result.keys).toEqual(['destroy', 'document', 'file', 'kind', 'save', 'view']);
    expect(result.blocks).toBe(true);
    expect(result.isView).toBe(true);
    expect(result.saved).toEqual(Array.from(input));
    expect(result.pages).toBeGreaterThanOrEqual(1);
    await page.waitForFunction(() => (window as any).__mounted.view.layoutDone);
    await expect(page.locator('#editor .bym-doc-pages')).toHaveText('Sayfa 1 / 3');
    await expect(page.locator('#editor canvas').first()).toBeVisible();

    // Ctrl+S hands the bytes to onSave.
    await page.evaluate(() => (window as any).__mounted.view.focus());
    await page.keyboard.press('Control+s');
    await expect.poll(() => page.evaluate(() => (window as any).__saves.length)).toBe(1);
    expect(await page.evaluate(() => (window as any).__saves[0])).toEqual(Array.from(input));

    await page.evaluate(() => (window as any).__mounted.destroy());
    await expect(page.locator('#editor')).toBeEmpty();
    expect(await page.evaluate(() => document.querySelectorAll('style[data-bymoffice-doc]').length)).toBe(0);
  });

  test('the zoom and toolbar options reach the document view; a workbook still mounts as a spreadsheet', async ({ page }) => {
    const zoom = await page.evaluate(async (bytes) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(bytes), zoom: 0.5, toolbar: false });
      const out = { zoom: mounted.view.zoom, bars: document.querySelectorAll('#editor .bym-doc-bar').length };
      mounted.destroy();
      return out;
    }, Array.from(await sampleDocx()));
    expect(zoom).toEqual({ zoom: 0.5, bars: 0 });

    const { sampleWorkbook } = await import('../test/fixtures');
    const kind = await page.evaluate(async (bytes) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(bytes) });
      const out = [mounted.kind, typeof mounted.workbook, typeof mounted.editor, mounted.document];
      mounted.destroy();
      return out;
    }, Array.from(await sampleWorkbook()));
    expect(kind).toEqual(['spreadsheet', 'object', 'object', undefined]);
  });

  test('a damaged docx is rejected with an error, the container stays empty and the page keeps working', async ({ page }) => {
    const good = await sampleDocx();
    const truncated = good.slice(0, good.length >> 1);
    const broken = await buildDocx({ document: '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p>' });
    const dtd = await buildDocx({
      document:
        '<?xml version="1.0"?><!DOCTYPE w:document [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;">]>' +
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>&b;</w:t></w:r></w:p></w:body></w:document>',
    });
    const noBody = await buildDocx({ document: '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>' });
    const outcome = await page.evaluate(async (inputs) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor')!;
      const results: string[] = [];
      for (const input of inputs) {
        const started = performance.now();
        try {
          const mounted = await api.mount(host, { file: new Uint8Array(input) });
          results.push(`mounted ${mounted.kind}`);
          mounted.destroy();
        } catch (error) {
          results.push(error instanceof Error && error.message.length > 0 ? 'error' : 'bad');
        }
        if (performance.now() - started > 5000) results.push('slow');
      }
      return { results, children: host.childNodes.length };
    }, [Array.from(truncated), Array.from(broken), Array.from(dtd), Array.from(noBody)]);
    // A document without a body is an empty document, not an error.
    expect(outcome.results.slice(0, 3)).toEqual(['error', 'error', 'error']);
    expect(outcome.results.length).toBe(4);
    expect(outcome.children).toBe(0);
    await expectResponsive(page);
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('a damaged xls and plain text are refused', async ({ page }) => {
    const { minimalCfb } = await import('../test/fixtures');
    // A compound file with a stream called "Workbook" is an xls; neither it nor plain text can be shown.
    const names = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const out: string[] = [];
      for (const file of [new Uint8Array(input), new TextEncoder().encode('plain text')]) {
        await api.mount(document.getElementById('editor'), { file }).catch((error: Error) => out.push(error.name));
      }
      return out;
    }, Array.from(minimalCfb(new Uint8Array(9 * 512), new Uint8Array(10))));
    expect(names).toEqual(['Error', 'UnsupportedFileError']);
    await expect(page.locator('#editor')).toBeEmpty();
  });
});

for (const flavour of ['script', 'esm'] as Flavour[]) {
  test(`document API is exported by the ${flavour} build`, async ({ page }) => {
    await gotoHarness(page, flavour);
    const names = ['DocumentView', 'DocLayout', 'renderPage', 'CanvasTextMeasurer', 'createCanvasMeasurer', 'docCssFont', 'documentText', 'mount'];
    const result = await page.evaluate(
      async ([wanted, bytes]) => {
        const api = (window as any).__bym;
        const file = await api.open(new Uint8Array(bytes as number[]));
        const doc = await file.document();
        // The pieces work on their own: lay out with the canvas measurer, paint a page on a canvas.
        const measurer = api.createCanvasMeasurer({ fonts: doc.fonts });
        const layout = new api.DocLayout(doc, measurer);
        layout.layout();
        const canvas = document.createElement('canvas');
        canvas.width = layout.pages[0].width;
        canvas.height = layout.pages[0].height;
        const context = canvas.getContext('2d')!;
        api.renderPage(context, layout.pages[0], { cssFont: measurer.cssFont });
        const data = context.getImageData(0, 0, canvas.width, canvas.height).data;
        let dark = 0;
        for (let i = 0; i < data.length; i += 4) if (data[i] < 80 && data[i + 1] < 80 && data[i + 2] < 80) dark++;
        // The measurer tells fonts the device has from fonts it lacks.
        const fonts = [measurer.hasFont('Definitely Missing Font 9f3a'), ['Arial', 'Helvetica', 'Liberation Sans', 'DejaVu Sans', 'Roboto', 'Noto Sans'].some((name) => measurer.hasFont(name))];
        return {
          fonts,
          exported: (wanted as string[]).filter((name) => typeof api[name] === 'function'),
          pages: layout.pages.length,
          dark: dark > 300,
          text: api.documentText(doc).slice(0, 14),
        };
      },
      [names, Array.from(await sampleDocx())] as [string[], number[]],
    );
    expect(result).toEqual({ fonts: [false, true], exported: names, pages: 3, dark: true, text: 'Document title' });
  });
}

// Used by the fixtures above; kept here so that an unused import is noticed.
void link;

// Defects found by the independent review of the layout engine and the view; every test here failed before its fix.
test.describe('document view: review fixes', () => {
  test('a single word of 300,000 characters is laid out in under three seconds and the page keeps responding', async ({ page }) => {
    test.setTimeout(90_000);
    const bytes = await longWordDocx(300_000);
    await page.goto('/e2e/pages/doc.html');
    await expect(page.locator('#state')).toHaveText('ready');
    const result = await page.evaluate(async (input) => {
      // What measuring costs: the characters handed to the canvas.
      let measured = 0;
      const original = CanvasRenderingContext2D.prototype.measureText;
      CanvasRenderingContext2D.prototype.measureText = function (text: string) {
        measured += text.length;
        return original.call(this, text);
      };
      const started = performance.now();
      await (window as any).__mountDoc(input, { zoom: 1 });
      const view = (window as any).__docView;
      await new Promise<void>((resolve) => {
        const tick = () => (view.layoutDone ? resolve() : setTimeout(tick, 0));
        tick();
      });
      const ms = performance.now() - started;
      CanvasRenderingContext2D.prototype.measureText = original;
      return { ms, measured, pages: view.pageCount as number };
    }, Array.from(bytes));
    console.log(
      `one word of 300,000 characters [${test.info().project.name}]: ${result.pages} pages after ${result.ms.toFixed(0)} ms, ${result.measured} characters measured`,
    );
    expect(result.ms).toBeLessThan(3000);
    // Linear in the length of the word: once as a whole, once more line by line. (It was 1.4 thousand million.)
    expect(result.measured).toBeLessThan(4 * 300_000);
    // About 70 letters a line and 18 lines a page.
    expect(result.pages).toBeGreaterThan(150);
    await expectResponsive(page);
    expect((await canvasStats(page, 1)).dark).toBeGreaterThan(500);
    // The end of the word is there.
    await page.evaluate((count) => (window as any).__docView.goToPage(count), result.pages);
    await settled(page);
    expect(((await pageText(page, result.pages)) || '').length).toBeGreaterThan(0);
    const all = await page.evaluate(() => (window as any).__docView.selectAll() || (window as any).__docView.getSelectedText().length);
    expect(all).toBe(300_000);
  });

  test('pictures far larger than they are shown never take more decoded memory than the budget', async ({ page }) => {
    test.setTimeout(240_000);
    const count = 14;
    const bytes = await hugePicturesDocx(count);
    expect(bytes.length).toBeLessThan(400_000);
    await mountDoc(page, bytes, { zoom: 1 });
    await laidOut(page);
    expect((await state(page)).pageCount).toBe(count);
    // Shown at 400 x 300 px: the red half and the blue half are there, decoded from 42 megapixels.
    await expect.poll(async () => (await canvasStats(page, 1)).red, { timeout: 60_000 }).toBeGreaterThan(10_000);
    const first = await state(page);
    expect(first.imageBudget).toBeGreaterThan(0);
    // 169 MB each as the file has them; held at about the size they are shown at, and no picture at more than an
    // eighth of the budget. Only the pictures of the pages that have a canvas are held at all.
    expect(first.imageBytes).toBeGreaterThan(0);
    expect(first.imageBytes).toBeLessThanOrEqual((first.mounted.length * first.imageBudget) / 8);
    expect(first.mounted.length).toBeLessThanOrEqual(4);

    // At the largest zoom every picture is wanted at more pixels than its share of the budget: each is held at that
    // share, and going through all the pages never takes the total over the budget.
    await page.evaluate(() => (window as any).__docView.setZoom(5));
    let most = 0;
    for (let n = 1; n <= count; n++) {
      await page.evaluate((number) => (window as any).__docView.goToPage(number), n);
      await expect
        .poll(
          async () => {
            const now = await state(page);
            most = Math.max(most, now.imageBytes);
            expect(now.imageBytes).toBeLessThanOrEqual(now.imageBudget);
            return (await canvasStats(page, n)).red;
          },
          { timeout: 60_000 },
        )
        .toBeGreaterThan(10_000);
    }
    const end = await state(page);
    console.log(`14 pictures of 6500 x 6500 px [${test.info().project.name}]: at most ${(most / 1048576).toFixed(1)} MB of ${(end.imageBudget / 1048576).toFixed(0)} MB held`);
    expect(most).toBeLessThanOrEqual(end.imageBudget);
    // More than half of the budget was in use: the limit was reached and pictures were given up, not just few held.
    expect(most).toBeGreaterThan(end.imageBudget / 2);
    await expectResponsive(page);
    // Destroying the view gives everything back.
    await page.evaluate(() => (window as any).__docView.destroy());
  });

  for (const [family, size, pitch] of [
    ['Times New Roman', 24, 18.4],
    ['Arial', 22, 16.87],
  ] as const) {
    test(`lines of ${family} ${size / 2} pt are as high as in Word, in every browser`, async ({ page }) => {
      await mountDoc(page, await linesDocx(family, size, 200), { zoom: 1 });
      await laidOut(page);
      const measured = await page.evaluate(async () => {
        const api = (window as any).BymOffice;
        const doc = await (window as any).__docFile.document();
        const layout = new api.DocLayout(doc, api.createCanvasMeasurer({ fonts: doc.fonts }));
        layout.layout();
        const ys: number[] = layout.pages[0].items.filter((item: any) => item.type === 'text').map((item: any) => item.y);
        return { pages: layout.pages.length as number, first: ys[0], pitch: ys[1] - ys[0], lines: ys.length, last: ys[ys.length - 1] - ys[ys.length - 2] };
      });
      // Word: 13.8 pt for Times New Roman 12 pt, 12.65 pt for Arial 11 pt.
      expect(Math.abs(measured.pitch - pitch)).toBeLessThanOrEqual(0.05);
      expect(Math.abs(measured.last - pitch)).toBeLessThanOrEqual(0.05);
      // 320 px of text area: the same number of lines on a page, and so of pages, whatever the browser.
      const perPage = Math.floor(320 / pitch);
      expect(measured.lines).toBe(perPage);
      expect(measured.pages).toBe(Math.ceil(200 / perPage));
      expect((await state(page)).pageCount).toBe(measured.pages);
    });
  }

  test('a selection dragged from one page to the next holds the text between, without footer or header', async ({ page, isMobile }) => {
    test.skip(isMobile, 'a finger scrolls; selecting by dragging is for the mouse');
    // Half size: two pages are in sight.
    await mountDoc(page, await sampleDocx(), { zoom: 0.5 });
    await laidOut(page);
    // The first paragraph of the second page.
    let k = 0;
    for (let n = 2; n <= 40 && !k; n++) if ((await textBox(page, `Paragraph ${n} of`)).page === 2) k = n;
    expect(k).toBeGreaterThan(3);
    const from = await textBox(page, `Paragraph ${k - 2} of`);
    const to = await textBox(page, `Paragraph ${k + 1} of`);
    expect([from.page, to.page]).toEqual([1, 2]);
    const sentence = (n: number) => `Paragraph ${n} of the sample document.`;
    const expected = [sentence(k - 2), sentence(k - 1), sentence(k), `Paragraph ${k + 1} of`].join('\n');

    await page.mouse.move(from.x + 0.5, from.y + from.height / 2);
    await page.mouse.down();
    await page.mouse.move(from.x + 40, from.y + from.height / 2 + 30, { steps: 3 });
    await page.mouse.move(to.x + to.width, to.y + to.height / 2, { steps: 5 });
    await page.mouse.up();
    await settled(page);
    const forward = await page.evaluate(() => (window as any).__docView.getSelectedText());
    expect(forward).toBe(expected);
    expect((await copyEvent(page)).text).toBe(expected);

    // The other way round: the same text.
    await page.mouse.click(from.x + 250, from.y + from.height / 2);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe('');
    await page.mouse.move(to.x + to.width, to.y + to.height / 2);
    await page.mouse.down();
    await page.mouse.move(to.x + 20, to.y - 20, { steps: 3 });
    await page.mouse.move(from.x + 0.5, from.y + from.height / 2, { steps: 5 });
    await page.mouse.up();
    await settled(page);
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe(expected);

    // A selection that starts in the footer stays there: the body it is dragged over is not taken.
    const footer = await textBox(page, 'Page ');
    await page.mouse.move(footer.x + 0.5, footer.y + footer.height / 2);
    await page.mouse.down();
    await page.mouse.move(from.x + 100, from.y, { steps: 4 });
    await page.mouse.up();
    const inFooter = await page.evaluate(() => (window as any).__docView.getSelectedText());
    expect(inFooter).not.toContain('Paragraph');
    expect('Page 1 of 1').toContain(inFooter);
  });

  test('a selection across the cells of a table takes whole cells of the rows and columns between', async ({ page, isMobile }) => {
    test.skip(isMobile, 'selecting by dragging is for the mouse');
    let rows = '';
    for (let r = 1; r <= 3; r++) rows += tr(tc(`A${r} one`) + tc(`B${r} two`) + tc(`C${r} three`));
    await mountDoc(page, await docx(p('Before the table') + tbl(rows, [2000, 2000, 2000]) + p('After the table')), { zoom: 1 });
    await laidOut(page);
    const drag = async (start: string, end: string) => {
      const a = await textBox(page, start);
      const b = await textBox(page, end);
      await page.mouse.move(a.x + 2, a.y + a.height / 2);
      await page.mouse.down();
      await page.mouse.move((a.x + b.x) / 2, (a.y + b.y) / 2 + 3, { steps: 3 });
      // The nearest character boundary is the end of the prefix.
      await page.mouse.move(b.x + b.width - 0.5, b.y + b.height / 2, { steps: 3 });
      await page.mouse.up();
      await settled(page);
      return page.evaluate(() => (window as any).__docView.getSelectedText());
    };
    // Down a column: only that column.
    expect(await drag('B1', 'B3')).toBe('B1 two\nB2 two\nB3 two');
    expect(await drag('B3', 'B1')).toBe('B1 two\nB2 two\nB3 two');
    // A block of cells: tabs between the cells of a row.
    expect(await drag('A1', 'B2')).toBe('A1 one\tB1 two\nA2 one\tB2 two');
    expect(await drag('C2', 'B1')).toBe('B1 two\tC1 three\nB2 two\tC2 three');
    // Inside one cell the selection is of its text.
    expect(await drag('B2', 'B2')).toBe('B2');
    // From the text before the table into it and beyond: in document order, the table as rows of cells.
    expect(await drag('Before', 'After')).toBe('Before the table\nA1 one\tB1 two\tC1 three\nA2 one\tB2 two\tC2 three\nA3 one\tB3 two\tC3 three\nAfter');
    // The selection is painted over the selected cells only: A1 keeps its white paper when B1 to B3 are selected.
    await drag('B1', 'B3');
    const shade = await page.evaluate(async () => {
      const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
      const scale = canvas.width / 600;
      const context = canvas.getContext('2d')!;
      // The first column is x = 40..173, the second 173..307; the rows start at y = 58.
      const count = (x: number, y: number, w: number, h: number) => {
        const data = context.getImageData(Math.round(x * scale), Math.round(y * scale), Math.round(w * scale), Math.round(h * scale)).data;
        let tinted = 0;
        for (let i = 0; i < data.length; i += 4) if (data[i] < 240 && data[i + 2] > 200 && data[i + 2] > data[i] + 20) tinted++;
        return tinted;
      };
      return { first: count(45, 60, 120, 50), second: count(180, 60, 120, 50) };
    });
    expect(shade.first).toBe(0);
    expect(shade.second).toBeGreaterThan(300);
  });

  test('a selection that starts in the body leaves a text box out; one that starts in the box stays in it', async ({ page, isMobile }) => {
    test.skip(isMobile, 'selecting by dragging is for the mouse');
    const box = textBoxRun(300, 0, 180, 60, p('Boxed words here') + p('Second box line'));
    const body = p('First paragraph of the body') + p('<w:r><w:t>Second paragraph, with a box</w:t></w:r>' + box) + p('Third paragraph of the body') + p('Fourth paragraph of the body');
    await mountDoc(page, await docx(body), { zoom: 1 });
    await laidOut(page);
    // The box is filled green: fill and frame of text boxes are painted.
    expect((await canvasStats(page, 1)).green).toBeGreaterThan(180 * 60 * 0.6);
    const first = await textBox(page, 'First paragraph');
    const fourth = await textBox(page, 'Fourth paragraph');
    const boxed = await textBox(page, 'Boxed words');
    const second = await textBox(page, 'Second box');
    await page.mouse.move(first.x + 1, first.y + first.height / 2);
    await page.mouse.down();
    // Over the box on the way.
    await page.mouse.move(boxed.x + 20, boxed.y + boxed.height / 2, { steps: 3 });
    await page.mouse.move(fourth.x + fourth.width, fourth.y + fourth.height / 2, { steps: 3 });
    await page.mouse.up();
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe(
      'First paragraph of the body\nSecond paragraph, with a box\nThird paragraph of the body\nFourth paragraph',
    );
    // From the box outwards: only text of the box.
    await page.mouse.move(boxed.x + 1, boxed.y + boxed.height / 2);
    await page.mouse.down();
    await page.mouse.move(second.x + second.width, second.y + second.height / 2, { steps: 3 });
    await page.mouse.move(fourth.x + 5, fourth.y + fourth.height / 2, { steps: 3 });
    await page.mouse.up();
    const inBox = await page.evaluate(() => (window as any).__docView.getSelectedText());
    expect(inBox.startsWith('Boxed words here')).toBe(true);
    expect(inBox).not.toContain('paragraph');
    // Select all is still the text of the body.
    await page.evaluate(() => (window as any).__docView.selectAll());
    expect(await page.evaluate(() => (window as any).__docView.getSelectedText())).toBe(
      'First paragraph of the body\nSecond paragraph, with a box\nThird paragraph of the body\nFourth paragraph of the body',
    );
  });

  test('the view mounts where elements have no getRootNode', async ({ page }) => {
    await page.goto('/e2e/pages/doc.html');
    await expect(page.locator('#state')).toHaveText('ready');
    const bytes = await sampleDocx();
    const result = await page.evaluate(async (input) => {
      const original = Node.prototype.getRootNode;
      // An old WebView: the method is not there at all.
      delete (Node.prototype as any).getRootNode;
      try {
        await (window as any).__mountDoc(input, { zoom: 1 });
        const styles = document.querySelectorAll('style[data-bymoffice-doc]').length;
        const canvases = document.querySelectorAll('#host canvas').length;
        (window as any).__docView.destroy();
        return { missing: typeof document.body.getRootNode, styles, canvases, left: document.querySelectorAll('style[data-bymoffice-doc]').length };
      } finally {
        Node.prototype.getRootNode = original;
      }
    }, Array.from(bytes));
    expect(result).toEqual({ missing: 'undefined', styles: 1, canvases: result.canvases, left: 0 });
    expect(result.canvases).toBeGreaterThanOrEqual(1);
  });

  test('text in a cell with a vertical text direction is painted turned', async ({ page }) => {
    const vertical = tc('Vertical words', '<w:textDirection w:val="btLr"/>');
    const table = tbl(tr(vertical + tc('across'), '<w:trHeight w:val="3000" w:hRule="atLeast"/>'), [1500, 3000]);
    await mountDoc(page, await docx(table), { zoom: 1 });
    await laidOut(page);
    // The extent of the dark pixels in the first cell (x = 40..140 of the page) and in the second.
    const ink = await page.evaluate(() => {
      const canvas = document.querySelector('#host .bym-doc-page[data-page="1"] canvas') as HTMLCanvasElement;
      const scale = canvas.width / 600;
      const context = canvas.getContext('2d')!;
      const extent = (x0: number, x1: number) => {
        const left = Math.round(x0 * scale);
        const width = Math.round((x1 - x0) * scale);
        const data = context.getImageData(left, 0, width, canvas.height).data;
        let minX = Infinity;
        let maxX = -1;
        let minY = Infinity;
        let maxY = -1;
        for (let i = 0; i < data.length; i += 4) {
          if (data[i] > 90 || data[i + 1] > 90 || data[i + 2] > 90) continue;
          const x = (i / 4) % width;
          const y = Math.floor(i / 4 / width);
          if (x < minX) minX = x;
          if (x > maxX) maxX = x;
          if (y < minY) minY = y;
          if (y > maxY) maxY = y;
        }
        return { width: (maxX - minX) / scale, height: (maxY - minY) / scale };
      };
      return { first: extent(42, 138), second: extent(142, 330) };
    });
    // Vertical: one line of 16 px set upwards, about 100 px long. Horizontal: the other way round.
    expect(ink.first.height).toBeGreaterThan(60);
    expect(ink.first.width).toBeLessThan(25);
    expect(ink.second.width).toBeGreaterThan(30);
    expect(ink.second.height).toBeLessThan(25);
    // The text is still there for copying and for assistive technology.
    expect(await pageText(page, 1)).toContain('Vertical words');
  });
});
