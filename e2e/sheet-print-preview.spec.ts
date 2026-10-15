// The print preview of a workbook: a modal layer over the editor with the
// pages print() prints, laid out and painted by the same code. Only the
// pages in view have an element and a canvas. No test opens a real print
// dialog: print() of the window is replaced, as in sheet-print.spec.ts.
//
// Positions are written out by hand: columns are 64 px wide and rows 20 px
// high; with the default margins cell A1 starts at (67.2, 72) on the paper.

import type { Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { expect, expectResponsive, test } from './helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const LEFT = 67.2;
const TOP = 72;
const A4_WIDTH = 794;

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const column = (index: number) => String.fromCharCode(65 + index);
const A4_SETUP = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9"/>';

function numbers(rows: number, cols: number): string {
  const parts: string[] = ['<sheetData>'];
  for (let r = 1; r <= rows; r++) {
    let row = `<row r="${r}">`;
    for (let c = 0; c < cols; c++) row += num(`${column(c)}${r}`, r * 100 + c);
    parts.push(row + '</row>');
  }
  parts.push('</sheetData>');
  return parts.join('');
}

function once<T>(make: () => Promise<T>): () => Promise<T> {
  let value: Promise<T> | undefined;
  return () => (value ??= make());
}

/** 100 rows by 12 columns on A4: three pages down and two across. */
const sixPages = once(() => buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(numbers(100, 12) + A4_SETUP) }] }));

/** Two sheets: six pages and one page. */
const twoSheets = once(() =>
  buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(numbers(100, 12) + A4_SETUP) },
      { name: 'Small', xml: worksheetXml(numbers(3, 2) + A4_SETUP) },
    ],
  }),
);

/** B2 has a red fill, nothing else. */
const redCell = once(() =>
  buildXlsx({
    styles:
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">` +
      '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
      '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></fills>' +
      '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs></styleSheet>',
    sheets: [{ name: 'Red', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row><row r="2"><c r="B2" s="1"/></row></sheetData>` + A4_SETUP) }],
  }),
);

/** About 1,980 pages of A4: one column of 95,000 numbers. */
const nearlyTwoThousand = once(() => buildXlsx({ sheets: [{ name: 'Long', xml: worksheetXml(numbers(95000, 1) + A4_SETUP) }] }));

interface MountOptions {
  locale?: string;
  readOnly?: boolean;
}

async function mountEditor(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(([input, opts]) => (window as any).__mountEditor(input, opts), [Array.from(bytes), { locale: 'en-US', ...options }] as [number[], MountOptions]);
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
  await page.evaluate(() => ((window as any).__view = (window as any).__editor.view));
}

/** print() of the window is recorded with how many pages the job put into the document, and "closes" at once. */
function fakePrint(page: Page): Promise<void> {
  return page.evaluate(() => {
    const w = window as any;
    w.__prints = [];
    window.print = () => {
      const element = document.querySelector('body > .bym-print');
      w.__prints.push({ sheets: element ? element.children.length : -1 });
      window.dispatchEvent(new Event('afterprint'));
    };
  });
}

const printCalls = (page: Page) => page.evaluate(() => ((window as any).__prints as { sheets: number }[]).slice());
const preview = (page: Page) => page.locator('#host .bym-preview');
const button = (page: Page, key: string) => page.locator(`#host .bym-preview [data-bym-preview="${key}"]`);
const counter = (page: Page) => page.locator('#host .bym-preview-counter');
const state = (page: Page) => page.evaluate(() => (window as any).__view.getPrintPreviewState());
const focusedKey = (page: Page) =>
  page.evaluate(() => {
    const active = document.activeElement as HTMLElement | null;
    return active ? active.getAttribute('data-bym-preview') || active.getAttribute('data-command') || active.className : '';
  });

/** Waits until the pages are laid out and the current one is painted. */
async function painted(page: Page, which = 0): Promise<void> {
  await expect.poll(async () => ((await state(page)) || { painted: [] }).painted.includes(which), { timeout: 30_000, message: `page ${which + 1} is painted` }).toBe(true);
}

test.describe('print preview', () => {
  test('the toolbar button opens a modal preview of the pages printing has; the buttons and the keys go through them', async ({ page }) => {
    await mountEditor(page, await sixPages());
    await fakePrint(page);
    const opener = page.locator('#host [data-command="printPreview"]');
    await expect(opener).toHaveAttribute('aria-haspopup', 'dialog');
    await expect(opener).toHaveAttribute('aria-label', /^Print preview \((Ctrl\+|⌃)F2\)$/);
    await opener.click();
    const dialog = preview(page);
    await expect(dialog).toBeVisible();
    await expect(dialog).toHaveAttribute('role', 'dialog');
    await expect(dialog).toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('#host .bym-preview-title')).toHaveText('Print preview');
    await expect(page.locator('#host .bym-preview-pages')).toHaveAttribute('aria-label', 'Pages');
    // The editor behind it takes nothing.
    expect(await page.evaluate(() => document.querySelector('#host .bym-toolbar')!.hasAttribute('inert'))).toBe(true);
    await expect(counter(page)).toHaveText('1 / 6');
    await painted(page, 0);
    await expect(page.locator('#host .bym-preview-page[data-page="0"]')).toHaveAttribute('aria-label', 'Page 1 of 6');
    await expect(page.locator('#host .bym-preview-page[data-page="0"]')).toHaveAttribute('role', 'img');
    expect((await state(page)).pages).toBe(6);

    await button(page, 'next').click();
    await expect(counter(page)).toHaveText('2 / 6');
    await painted(page, 1);
    await button(page, 'previous').click();
    await expect(counter(page)).toHaveText('1 / 6');
    await expect(button(page, 'previous')).toBeDisabled();
    await page.locator('#host .bym-preview-pages').focus();
    await page.keyboard.press('End');
    await expect(counter(page)).toHaveText('6 / 6');
    await painted(page, 5);
    await expect(button(page, 'next')).toBeDisabled();
    await page.keyboard.press('PageUp');
    await expect(counter(page)).toHaveText('5 / 6');
    await page.keyboard.press('Home');
    await expect(counter(page)).toHaveText('1 / 6');
    await page.keyboard.press('PageDown');
    await expect(counter(page)).toHaveText('2 / 6');

    // Print from the preview: the job has as many pages as the preview shows.
    await button(page, 'print').click();
    await expect.poll(async () => (await printCalls(page)).length, { timeout: 60_000 }).toBe(1);
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
    await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
    await expect(dialog).toBeVisible();
    await expectResponsive(page);
  });

  test('Escape closes it and the focus goes back to the button; Tab stays inside it', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, 'Tab and Escape are keys of hardware keyboards; the desktop projects cover them');
    await mountEditor(page, await sixPages());
    const opener = page.locator('#host [data-command="printPreview"]');
    await opener.focus();
    await page.keyboard.press('Enter');
    await expect(preview(page)).toBeVisible();
    await painted(page, 0);
    for (let i = 0; i < 20; i++) {
      await page.keyboard.press('Tab');
      expect(await page.evaluate(() => !!document.activeElement && !!document.activeElement.closest('.bym-preview'))).toBe(true);
    }
    for (let i = 0; i < 5; i++) {
      await page.keyboard.press('Shift+Tab');
      expect(await page.evaluate(() => !!document.activeElement && !!document.activeElement.closest('.bym-preview'))).toBe(true);
    }
    await page.keyboard.press('Escape');
    await expect(preview(page)).toHaveCount(0);
    expect(await focusedKey(page)).toBe('printPreview');
    expect(await page.evaluate(() => document.querySelector('#host .bym-toolbar')!.hasAttribute('inert'))).toBe(false);
    expect(await page.evaluate(() => (window as any).__view.previewing)).toBe(false);
    expect(await state(page)).toBeUndefined();
  });

  test('Ctrl+F2 opens it from the grid; Ctrl+P in it prints', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, 'a shortcut of hardware keyboards; the desktop projects cover it');
    await mountEditor(page, await sixPages());
    await fakePrint(page);
    await page.evaluate(() => (window as any).__view.focus());
    await page.keyboard.press('Control+F2');
    await expect(preview(page)).toBeVisible();
    await painted(page, 0);
    await page.keyboard.press('ControlOrMeta+p');
    await expect.poll(async () => (await printCalls(page)).length, { timeout: 60_000 }).toBe(1);
    await page.keyboard.press('Escape');
    await expect(preview(page)).toHaveCount(0);
    // Back in the grid, where it was opened from.
    expect(await page.evaluate(() => document.activeElement!.className)).toBe('bym-input');
  });

  test('what it paints is what is printed: a red cell at its place on the paper, at 100%', async ({ page }) => {
    await mountEditor(page, await redCell());
    await page.evaluate(() => (window as any).__view.printPreview());
    await expect(preview(page)).toBeVisible();
    await button(page, 'actualSize').click();
    await expect.poll(async () => (await state(page)).zoom).toBe(1);
    await painted(page, 0);
    const pixels = await page.evaluate(
      ([x, y, a4]) => {
        const canvas = document.querySelector('#host .bym-preview-page[data-page="0"] canvas') as HTMLCanvasElement;
        const k = canvas.width / a4;
        const context = canvas.getContext('2d')!;
        const at = (px: number, py: number) => Array.from(context.getImageData(Math.round(px * k), Math.round(py * k), 1, 1).data);
        return { red: at(x, y), paper: at(5, 5), css: canvas.getBoundingClientRect().width };
      },
      [LEFT + 64 + 32, TOP + 20 + 10, A4_WIDTH] as [number, number, number],
    );
    expect(pixels.red).toEqual([255, 0, 0, 255]);
    expect(pixels.paper).toEqual([255, 255, 255, 255]);
    // 100%: a px of the paper is a px of the page.
    expect(Math.round(pixels.css)).toBe(A4_WIDTH);
  });

  test('zoom: the buttons and + and - step it, "fit" shows a whole page', async ({ page }) => {
    await mountEditor(page, await sixPages());
    await page.evaluate(() => (window as any).__view.printPreview());
    await painted(page, 0);
    const fit = await state(page);
    expect(fit.fit).toBe(true);
    await expect(button(page, 'fit')).toBeDisabled();
    await button(page, 'actualSize').click();
    await expect.poll(async () => (await state(page)).zoom).toBe(1);
    await expect(button(page, 'actualSize')).toHaveText('100%');
    await button(page, 'zoomIn').click();
    await expect.poll(async () => (await state(page)).zoom).toBe(1.1);
    await button(page, 'zoomOut').click();
    await button(page, 'zoomOut').click();
    await expect.poll(async () => (await state(page)).zoom).toBe(0.9);
    await page.locator('#host .bym-preview-pages').focus();
    await page.keyboard.press('+');
    await expect.poll(async () => (await state(page)).zoom).toBe(1);
    await page.keyboard.press('-');
    await expect.poll(async () => (await state(page)).zoom).toBe(0.9);
    await button(page, 'fit').click();
    await expect.poll(async () => (await state(page)).fit).toBe(true);
    await expect.poll(async () => (await state(page)).zoom).toBeCloseTo(fit.zoom, 5);
    // A whole page in view: its element is no larger than the area.
    const boxes = await page.evaluate(() => ({
      page: document.querySelector('#host .bym-preview-page[data-page="0"]')!.getBoundingClientRect().height,
      area: (document.querySelector('#host .bym-preview-pages') as HTMLElement).clientHeight,
    }));
    expect(boxes.page).toBeLessThanOrEqual(boxes.area);
  });

  test('the whole workbook: as many pages as print() has for all sheets', async ({ page }) => {
    await mountEditor(page, await twoSheets());
    await fakePrint(page);
    await page.evaluate(() => (window as any).__view.printPreview());
    await painted(page, 0);
    expect((await state(page)).pages).toBe(6);
    await page.locator('#host .bym-preview-sheets').selectOption('all');
    await expect(counter(page)).toHaveText('1 / 7');
    expect((await state(page)).sheets).toBe('all');
    await page.evaluate(() => (window as any).__view.print({ sheets: 'all' }));
    expect(await printCalls(page)).toEqual([{ sheets: 7 }]);
    // From outside: the shown sheet again, at a page.
    await page.evaluate(() => (window as any).__view.printPreview({ sheets: 'active', page: 4 }));
    await expect(counter(page)).toHaveText('4 / 6');
    await expect(page.locator('#host .bym-preview-sheets')).toHaveValue('active');
  });

  test('the page setup from the preview lays the pages out anew; the focus comes back to its button', async ({ page }) => {
    await mountEditor(page, await sixPages());
    await page.evaluate(() => (window as any).__view.printPreview());
    await painted(page, 0);
    await button(page, 'pageSetup').click();
    const dialog = page.locator('#host .bym-page-setup-dialog');
    await expect(dialog).toBeVisible();
    // The preview is behind the dialog and takes nothing meanwhile.
    expect(await page.evaluate(() => document.querySelector('#host .bym-preview')!.hasAttribute('inert'))).toBe(true);
    await page.locator('#host .bym-page-setup-orientation').selectOption('landscape');
    await page.locator('#host .bym-page-setup-ok').click();
    await expect(dialog).toHaveCount(0);
    // Landscape: the twelve columns fit across, so 100 rows take four pages.
    await expect(counter(page)).toHaveText('1 / 4');
    await painted(page, 0);
    const box = await page.locator('#host .bym-preview-page[data-page="0"]').boundingBox();
    expect(box!.width).toBeGreaterThan(box!.height);
    expect(await focusedKey(page)).toBe('pageSetup');
    // Undone: back to six pages.
    await page.keyboard.press('Escape');
    await page.evaluate(() => (window as any).__view.undo());
    await page.evaluate(() => (window as any).__view.printPreview());
    await expect(counter(page)).toHaveText('1 / 6');
  });

  test('read-only, in Turkish: it opens without the page setup, in the words of the language', async ({ page }) => {
    await mountEditor(page, await sixPages(), { readOnly: true, locale: 'tr' });
    const opener = page.locator('#host [data-command="printPreview"]');
    await expect(opener).toBeEnabled();
    await opener.click();
    await painted(page, 0);
    await expect(page.locator('#host .bym-preview-title')).toHaveText('Baskı önizleme');
    await expect(page.locator('#host .bym-preview-page[data-page="0"]')).toHaveAttribute('aria-label', 'Sayfa 1 / 6');
    await expect(button(page, 'pageSetup')).toBeDisabled();
    await expect(button(page, 'print')).toBeEnabled();
    await expect(button(page, 'pdf')).toBeEnabled();
    await expect(button(page, 'close')).toHaveAttribute('aria-label', 'Baskı önizlemeyi kapat');
    await button(page, 'close').click();
    await expect(preview(page)).toHaveCount(0);
  });

  test('a sheet with nothing on it says so', async ({ page }) => {
    await mountEditor(page, await buildXlsx({ sheets: [{ name: 'Empty', xml: worksheetXml('<sheetData/>') }] }));
    await page.evaluate(() => (window as any).__view.printPreview());
    await expect(page.locator('#host .bym-preview-message')).toHaveText('There is nothing to print');
    expect((await state(page)).state).toBe('empty');
    await expect(button(page, 'print')).toBeDisabled();
    await expect(button(page, 'pdf')).toBeDisabled();
  });

  test('closePrintPreview() and destroy() take it away', async ({ page }) => {
    await mountEditor(page, await sixPages());
    expect(await page.evaluate(() => (window as any).__view.closePrintPreview())).toBe(false);
    expect(await page.evaluate(() => (window as any).__view.printPreview())).toBe(true);
    await painted(page, 0);
    expect(await page.evaluate(() => (window as any).__view.closePrintPreview())).toBe(true);
    await expect(preview(page)).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.printPreview())).toBe(true);
    await painted(page, 0);
    await page.evaluate(() => (window as any).__editor.destroy());
    await expect(page.locator('.bym-preview')).toHaveCount(0);
    expect(await page.evaluate(() => { try { (window as any).__view.printPreview({ sheets: 'selection' }); return 'none'; } catch (error: any) { return error.name; } })).toBe('RangeError');
  });

  test('about 2,000 pages: only those in view have a canvas; the first is shown quickly and the last one is reached', async ({ page }) => {
    test.setTimeout(240_000);
    await mountEditor(page, await nearlyTwoThousand());
    const opened = await page.evaluate(() => {
      const w = window as any;
      const started = performance.now();
      w.__view.printPreview();
      return started;
    });
    await painted(page, 0);
    const firstShown = await page.evaluate((started) => performance.now() - started, opened);
    const first = await state(page);
    expect(first.pages).toBeGreaterThan(1900);
    expect(first.pages).toBeLessThanOrEqual(2000);
    expect(first.shown.length).toBeLessThanOrEqual(64);
    expect(first.canvases).toBeLessThanOrEqual(64);
    // A guard against a hang, not a measure of speed.
    expect(firstShown).toBeLessThan(60_000);
    // Thumbnails: more pages in view, still a few of the 2,000.
    await page.locator('#host .bym-preview-pages').focus();
    for (let i = 0; i < 8; i++) await page.keyboard.press('-');
    await expect.poll(async () => (await state(page)).zoom).toBe(0.25);
    await expect.poll(async () => (await state(page)).painted.length, { timeout: 60_000 }).toBeGreaterThan(1);
    const thumbs = await state(page);
    expect(thumbs.canvases).toBeLessThanOrEqual(64);
    expect(thumbs.shown.length).toBeLessThan(200);
    await page.locator('#host .bym-preview-pages').focus();
    await page.keyboard.press('End');
    await expect(counter(page)).toHaveText(`${first.pages} / ${first.pages}`);
    await painted(page, first.pages - 1);
    const last = await state(page);
    // The pages at the start gave their canvases back.
    expect(last.shown.includes(0)).toBe(false);
    expect(last.canvases).toBeLessThanOrEqual(64);
    // 4 bytes a pixel: within the budget of the device.
    expect(last.pixels).toBeLessThanOrEqual(32 * 1024 * 1024);
    test.info().annotations.push({ type: 'measure', description: `${first.pages} pages: first page painted after ${Math.round(firstShown)} ms; ${last.canvases} canvases, ${(last.pixels / 1e6).toFixed(1)} M pixels at the end` });
    await expectResponsive(page);
  });

  test('review: Print and PDF take the sheets the preview shows, also after the page setup of another sheet was cancelled', async ({ page }) => {
    await mountEditor(page, await twoSheets());
    await fakePrint(page);
    await page.evaluate(() => (window as any).__view.printPreview());
    await painted(page, 0);
    await page.locator('#host .bym-preview-sheets').selectOption('all');
    await expect(counter(page)).toHaveText('1 / 7');
    await page.evaluate(() => (window as any).__view.printPreview({ page: 7 }));
    await expect(counter(page)).toHaveText('7 / 7');
    // The page setup of the page shown is that of the second sheet: it is shown behind the preview now.
    await button(page, 'pageSetup').click();
    await expect(page.locator('#host .bym-page-setup-dialog')).toBeVisible();
    await page.locator('#host .bym-page-setup-cancel').click();
    await expect(page.locator('#host .bym-page-setup-dialog')).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.activeSheet)).toBe(1);
    await page.locator('#host .bym-preview-sheets').selectOption('active');
    await expect(counter(page)).toHaveText('1 / 6');
    await button(page, 'print').click();
    await expect.poll(async () => (await printCalls(page)).length, { timeout: 60_000 }).toBe(1);
    // What was previewed is what was printed: the six pages of the first sheet.
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
  });

  test('review: a press on the bar outside its buttons leaves the focus in the preview: Escape closes it, Tab stays in it', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, 'Tab and Escape are keys of hardware keyboards; the desktop projects cover them');
    await mountEditor(page, await sixPages());
    await page.locator('#host [data-command="printPreview"]').click();
    await painted(page, 0);
    const inside = () => page.evaluate(() => !!document.activeElement && !!document.activeElement.closest('.bym-preview'));
    await page.locator('#host .bym-preview-title').click();
    expect(await inside()).toBe(true);
    await page.locator('#host .bym-preview-counter').click();
    expect(await inside()).toBe(true);
    await page.keyboard.press('Tab');
    expect(await inside()).toBe(true);
    await page.locator('#host .bym-preview-bar').click({ position: { x: 2, y: 2 } });
    await page.keyboard.press('Escape');
    await expect(preview(page)).toHaveCount(0);
  });

  test('review: the page asked for when it opens is kept when the workbook changes before the pages are laid out', async ({ page }) => {
    await mountEditor(page, await sixPages());
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.printPreview({ page: 4 });
      // A change in the same turn: the layout waits for it, and still opens at page 4.
      view.setLocale('en-US');
    });
    await expect(counter(page)).toHaveText('4 / 6');
  });
});
