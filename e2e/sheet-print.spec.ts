// Printing a workbook: the cells are laid out on pages by the page setup of
// their sheet, each page is prepared as a picture in an element that only
// shows in print, and the window's print() is called. No test opens a real
// print dialog: print() is replaced, and what would be printed is looked at
// in the document. The pictures are read back pixel by pixel.
//
// Positions are written out by hand: columns are 64 px wide and rows 20 px
// high; at 96 dots per inch a pixel of the picture is a px of the paper, and
// with the default margins cell A1 starts at (67.2, 72).

import type { Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { comment, commentsXml, drawingWorkbook, solid, sp, twoCell, vmlNote, vmlXml } from '../test/sheet-drawing-fixtures';
import { expect, expectResponsive, test } from './helpers';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const LEFT = 67.2;
const TOP = 72;
const A4_PT = [595.28, 841.89];
/** A4 at 200 and at 96 dots per inch. */
const A4_PIXELS = [1654, 2339];
const A4_SCREEN = [794, 1123];

const num = (ref: string, value: number, style?: number) => `<c r="${ref}"${style === undefined ? '' : ` s="${style}"`}><v>${value}</v></c>`;
const str = (ref: string, text: string, style?: number) => `<c r="${ref}"${style === undefined ? '' : ` s="${style}"`} t="inlineStr"><is><t>${text}</t></is></c>`;
const column = (index: number) => String.fromCharCode(65 + index);

/** Cell formats: 0 plain, 1 red fill, 2 blue fill. One differential format: green fill. */
const STYLES =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="${NS}">` +
  '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>' +
  '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>' +
  '<fill><patternFill patternType="solid"><fgColor rgb="FF0000FF"/></patternFill></fill></fills>' +
  '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
  '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
  '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
  '<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>' +
  '<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/></cellXfs>' +
  '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FF00B050"/></patternFill></fill></dxf></dxfs></styleSheet>';

/** Rows of numbers: `rows` x `cols` cells from A1. */
function numbers(rows: number, cols: number): string {
  let data = '<sheetData>';
  for (let r = 1; r <= rows; r++) {
    data += `<row r="${r}">`;
    for (let c = 0; c < cols; c++) data += num(`${column(c)}${r}`, r * 100 + c);
    data += '</row>';
  }
  return data + '</sheetData>';
}

const A4_SETUP = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9"/>';

function once<T>(make: () => Promise<T>): () => Promise<T> {
  let value: Promise<T> | undefined;
  return () => (value ??= make());
}

/** 100 rows by 12 columns on A4: three pages down (48 + 48 + 4 rows) and two across (10 + 2 columns). */
const sixPages = once(() => buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(numbers(100, 12) + A4_SETUP) }] }));

/**
 * A1 text, B2 red, C3 blue, A5:B6 merged and red, D1:D3 numbers of which those above 4 turn green.
 * Printed with grid lines, with a header and a footer.
 */
const formatted = (extra = '<printOptions gridLines="1"/>') =>
  buildXlsx({
    styles: STYLES,
    sheets: [
      {
        name: 'Styled',
        xml: worksheetXml(
          '<sheetData>' +
            `<row r="1">${str('A1', 'Hello')}${num('D1', 1)}</row>` +
            `<row r="2"><c r="B2" s="1"/>${num('D2', 5)}</row>` +
            `<row r="3"><c r="C3" s="2"/>${num('D3', 9)}</row>` +
            `<row r="5"><c r="A5" s="1"/><c r="B5" s="1"/></row>` +
            `<row r="6"><c r="A6" s="1"/><c r="B6" s="1"/></row>` +
            '</sheetData>' +
            '<mergeCells count="1"><mergeCell ref="A5:B6"/></mergeCells>' +
            '<conditionalFormatting sqref="D1:D3"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>4</formula></cfRule></conditionalFormatting>' +
            extra +
            A4_SETUP +
            '<headerFooter><oddHeader>&amp;L&amp;"Arial,Bold"&amp;20HEADER&amp;RPage &amp;P of &amp;N</oddHeader><oddFooter>&amp;C&amp;20&amp;A</oddFooter></headerFooter>',
        ),
      },
    ],
  });

interface ViewOptions {
  locale?: string;
  readOnly?: boolean;
  sheet?: number;
}

async function mountView(page: Page, bytes: Uint8Array, options: ViewOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/view.html');
  await expect(page.locator('#state')).toHaveText('ready');
  await page.evaluate(([input, opts]) => (window as any).__mountView(input, opts), [Array.from(bytes), { locale: 'en-US', ...options }] as [number[], ViewOptions]);
  await expect(page.locator('.bym-canvas')).toBeVisible();
}

async function mountEditor(page: Page, bytes: Uint8Array, options: ViewOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(([input, opts]) => (window as any).__mountEditor(input, opts), [Array.from(bytes), { locale: 'en-US', ...options }] as [number[], ViewOptions]);
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
  await page.evaluate(() => ((window as any).__view = (window as any).__editor.view));
}

/**
 * Replaces print() of the window: the call is recorded with what the
 * document holds at that moment. With `close` the dialog "closes" at once
 * (the afterprint event); without, it stays open until closeDialog().
 * Blob addresses made and given back are recorded as well.
 */
function fakePrint(page: Page, close: boolean): Promise<void> {
  return page.evaluate((closes) => {
    const w = window as any;
    w.__prints = [];
    w.__urls = { made: [] as string[], revoked: [] as string[] };
    const make = URL.createObjectURL;
    const revoke = URL.revokeObjectURL;
    URL.createObjectURL = function (this: unknown, blob: Blob) {
      const url = make.call(URL, blob);
      if (blob.type === 'image/png') w.__urls.made.push(url);
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

const closeDialog = (page: Page) => page.evaluate(() => void window.dispatchEvent(new Event('afterprint')));

/** Starts print() of the view and keeps its outcome in window.__printed. */
function startPrint(page: Page, options?: Record<string, unknown>): Promise<void> {
  return page.evaluate((opts) => {
    const w = window as any;
    w.__printed = undefined;
    w.__view.print(opts).then(
      (result: unknown) => (w.__printed = { result }),
      (error: any) => (w.__printed = { error: { name: error && error.name, code: error && error.code, message: String(error && error.message) } }),
    );
  }, options);
}

/** Runs a method of the view at the moment the print job has `count` pages ready. */
function whenPagesReady(page: Page, count: number, action: string, args: unknown[] = []): Promise<void> {
  return page.evaluate(
    ([wanted, method, values]) => {
      const observer = new MutationObserver(() => {
        if (document.querySelectorAll('.bym-print-sheet').length < (wanted as number)) return;
        observer.disconnect();
        const view = (window as any).__view;
        view[method as string](...(values as unknown[]));
      });
      observer.observe(document.body, { childList: true, subtree: true });
    },
    [count, action, args] as [number, string, unknown[]],
  );
}

const printCalls = (page: Page) => page.evaluate(() => ((window as any).__prints as { sheets: number }[]).slice());
const printed = (page: Page) => page.evaluate(() => (window as any).__printed);
const notices = (page: Page) => page.evaluate(() => (window as any).__notices as { message: string; kind: string }[]);
const waitForDialog = (page: Page, calls = 1) =>
  expect.poll(async () => (await printCalls(page)).length, { timeout: 60_000, message: 'print() of the window is called' }).toBe(calls);
const waitForOutcome = async (page: Page) => {
  await expect.poll(() => printed(page), { timeout: 60_000, message: 'print() of the view settles' }).toBeTruthy();
  return printed(page);
};

interface FrameState {
  present: boolean;
  atBody: boolean;
  text: string;
  styles: number;
  css: string;
  sheets: string[];
  images: { width: number; height: number; complete: boolean; alt: string; scheme: string }[];
}

/** What the print job has put into the document. */
function frameState(page: Page): Promise<FrameState> {
  return page.evaluate(() => {
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
      sheets: sheets.map((sheet) => (sheet.getAttribute('style') || '').replace(/\s+/g, '').replace(/;$/, '')),
      images: images.map((image) => ({ width: image.naturalWidth, height: image.naturalHeight, complete: image.complete, alt: image.alt, scheme: image.src.slice(0, 5) })),
    };
  });
}

/** Nothing of a print job is left in the document, and every picture it made was given back. */
async function expectCleanedUp(page: Page, pictures?: number): Promise<void> {
  const left = await frameState(page);
  expect(left.present).toBe(false);
  expect(left.styles).toBe(0);
  const all = await page.evaluate(() => (window as any).__urls as { made: string[]; revoked: string[] });
  if (pictures !== undefined) expect(all.made.length).toBe(pictures);
  expect(all.made.filter((url) => all.revoked.indexOf(url) < 0)).toEqual([]);
}

/** Reads the picture of a printed page into window.__picture, so that its pixels can be asked for. */
async function loadPicture(page: Page, index: number): Promise<void> {
  await page.evaluate((n) => {
    const image = document.querySelectorAll('.bym-print img')[n] as HTMLImageElement;
    const canvas = document.createElement('canvas');
    canvas.width = image.naturalWidth;
    canvas.height = image.naturalHeight;
    const context = canvas.getContext('2d', { willReadFrequently: true })!;
    context.drawImage(image, 0, 0);
    (window as any).__picture = context.getImageData(0, 0, canvas.width, canvas.height);
  }, index);
}

/** The colour of the pixel at (x, y) of the loaded picture. */
const pixel = (page: Page, x: number, y: number): Promise<number[]> =>
  page.evaluate(
    ([px, py]) => {
      const picture = (window as any).__picture as ImageData;
      const at = (Math.floor(py) * picture.width + Math.floor(px)) * 4;
      return Array.from(picture.data.slice(at, at + 3));
    },
    [x, y],
  );

/** Counts in a box of the loaded picture: pixels darker than mid grey, pixels of each plain colour, and pixels that are not white. */
const count = (page: Page, x: number, y: number, width: number, height: number): Promise<{ dark: number; red: number; blue: number; green: number; selection: number; ink: number }> =>
  page.evaluate(
    ([x0, y0, w, h]) => {
      const picture = (window as any).__picture as ImageData;
      const out = { dark: 0, red: 0, blue: 0, green: 0, selection: 0, ink: 0 };
      const x1 = Math.min(Math.floor(x0 + w), picture.width);
      const y1 = Math.min(Math.floor(y0 + h), picture.height);
      for (let py = Math.max(Math.floor(y0), 0); py < y1; py++) {
        for (let px = Math.max(Math.floor(x0), 0); px < x1; px++) {
          const at = (py * picture.width + px) * 4;
          const r = picture.data[at];
          const g = picture.data[at + 1];
          const b = picture.data[at + 2];
          if (r < 250 || g < 250 || b < 250) out.ink++;
          if (r < 140 && g < 140 && b < 140) out.dark++;
          if (r > 200 && g < 60 && b < 60) out.red++;
          if (b > 200 && r < 60 && g < 60) out.blue++;
          if (g > 150 && r < 60 && b < 120) out.green++;
          // The border of the selection on screen: #1A7F45.
          if (Math.abs(r - 26) < 12 && Math.abs(g - 127) < 12 && Math.abs(b - 69) < 12) out.selection++;
        }
      }
      return out;
    },
    [x, y, width, height],
  );

const RED = [255, 0, 0];
const BLUE = [0, 0, 255];
const WHITE = [255, 255, 255];

test.describe('printing a workbook', () => {
  test('print() lays the sheet out on pages of its paper, prepares one picture per page, calls print() of the window and cleans up after the dialog', async ({ page }) => {
    await mountView(page, await sixPages());
    await fakePrint(page, false);
    expect((await frameState(page)).present).toBe(false);
    await startPrint(page);
    await waitForDialog(page);
    // print() was called once, with every page ready.
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
    const frame = await frameState(page);
    expect(frame.atBody).toBe(true);
    expect(frame.styles).toBe(1);
    expect(frame.css).toContain(`@page{size:${A4_PT[0]}pt ${A4_PT[1]}pt;margin:0}`);
    expect(frame.sheets).toEqual(new Array(6).fill(`width:${A4_PT[0]}pt;height:${A4_PT[1]}pt`));
    expect(frame.images.map((image) => [image.width, image.height, image.complete, image.scheme])).toEqual(new Array(6).fill([...A4_PIXELS, true, 'blob:']));
    // The only words in the element are the view's own; nothing of the workbook is there as text.
    expect(frame.images.map((image) => image.alt)).toEqual(['Page 1', 'Page 2', 'Page 3', 'Page 4', 'Page 5', 'Page 6']);
    expect(frame.text).toBe('');
    // Down, then over: page 1 is full of numbers, page 3 has four rows, page 6 four rows of two columns.
    const scale = 200 / 96;
    const inked: number[] = [];
    for (const index of [0, 2, 5]) {
      await loadPicture(page, index);
      // The lower half of the room for cells: rows 25 to 48.
      inked.push((await count(page, LEFT * scale, (TOP + 480) * scale, 640 * scale, 480 * scale)).dark);
      // Nothing is printed in the margins.
      expect((await count(page, 0, 0, A4_PIXELS[0], (TOP - 2) * scale)).ink).toBe(0);
      expect((await count(page, 0, 0, (LEFT - 2) * scale, A4_PIXELS[1])).ink).toBe(0);
    }
    expect(inked[0]).toBeGreaterThan(3000);
    expect(inked[1]).toBe(0);
    expect(inked[2]).toBe(0);
    // The last page: ink in its two columns, none right of them.
    expect((await count(page, LEFT * scale, TOP * scale, 128 * scale, 80 * scale)).dark).toBeGreaterThan(100);
    expect((await count(page, (LEFT + 130) * scale, TOP * scale, 500 * scale, 80 * scale)).ink).toBe(0);
    // On screen the element is not shown and the view is as it was.
    await expect(page.locator('.bym-print')).toBeHidden();
    await expect(page.locator('.bym-canvas')).toBeVisible();
    expect(await page.evaluate(() => (window as any).__view.printing)).toBe(true);
    expect(await printed(page)).toBeUndefined();

    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 6, dpi: 200 } });
    await expectCleanedUp(page, 6);
    expect(await page.evaluate(() => (window as any).__view.printing)).toBe(false);
    expect((await notices(page)).filter((notice) => /print/i.test(notice.message))).toEqual([
      { message: 'Preparing to print…', kind: 'info' },
      { message: 'The print dialog was closed', kind: 'info' },
    ]);
  });

  test('fills, merged cells, conditional formats, grid lines, header and footer are where the page setup puts them; nothing of the screen is printed', async ({ page }) => {
    await mountView(page, await formatted());
    await fakePrint(page, false);
    // A selection across the cells that are looked at: its frame and its tint belong to the screen.
    await page.evaluate(() => (window as any).__view.select({ top: 0, left: 0, bottom: 3, right: 3 }));
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    const frame = await frameState(page);
    expect(frame.images.map((image) => [image.width, image.height])).toEqual([A4_SCREEN]);
    await loadPicture(page, 0);
    const cell = (col: number, row: number, dx = 0.5, dy = 0.5) => pixel(page, LEFT + (col + dx) * 64, TOP + (row + dy) * 20);
    // Cell fills.
    expect(await cell(1, 1)).toEqual(RED);
    expect(await cell(2, 2)).toEqual(BLUE);
    // The merged area A5:B6 is one red block: no grid line crosses it.
    expect((await count(page, LEFT + 2, TOP + 82, 124, 36)).red).toBe(124 * 36);
    // The conditional format: D2 and D3 are green, D1 is not.
    expect((await count(page, LEFT + 3 * 64 + 2, TOP + 22, 30, 16)).green).toBeGreaterThan(300);
    expect((await count(page, LEFT + 3 * 64 + 2, TOP + 42, 30, 16)).green).toBeGreaterThan(300);
    expect((await count(page, LEFT + 3 * 64 + 2, TOP + 2, 60, 16)).green).toBe(0);
    // An empty cell inside the selection is white: no tint, and no frame of the selection anywhere on the page.
    expect(await cell(2, 0)).toEqual(WHITE);
    expect((await count(page, 0, 0, A4_SCREEN[0], A4_SCREEN[1])).selection).toBe(0);
    // Grid lines: a line between A1 and B1 and one under row 1, none beyond the used range (D6).
    const darkest = async (x: number, y: number, w: number, h: number) => (await count(page, x, y, w, h)).ink;
    expect(await darkest(LEFT + 62, TOP + 105, 4, 6)).toBeGreaterThan(0);
    expect(await darkest(LEFT + 70, TOP + 18, 40, 4)).toBeGreaterThan(0);
    expect(await darkest(LEFT + 4 * 64 + 4, TOP, 300, 120)).toBe(0);
    expect(await darkest(LEFT, TOP + 6 * 20 + 4, 600, 300)).toBe(0);
    // The text of A1.
    expect((await count(page, LEFT + 2, TOP + 2, 50, 16)).dark).toBeGreaterThan(20);
    // Header: text at the left margin and at the right margin, 0.3 inch from the top, none in the middle.
    expect((await count(page, LEFT, 28, 200, 34)).dark).toBeGreaterThan(150);
    expect((await count(page, A4_SCREEN[0] - LEFT - 120, 28, 120, 34)).dark).toBeGreaterThan(30);
    expect((await count(page, 330, 20, 130, 45)).ink).toBe(0);
    // Footer: the name of the sheet in the middle, above the bottom margin of 0.3 inch.
    expect((await count(page, 340, A4_SCREEN[1] - 70, 114, 42)).dark).toBeGreaterThan(60);
    expect((await count(page, LEFT, A4_SCREEN[1] - 70, 200, 42)).ink).toBe(0);
    expect((await count(page, 0, A4_SCREEN[1] - 27, A4_SCREEN[0], 27)).ink).toBe(0);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 1, dpi: 96 } });
    await expectCleanedUp(page, 1);
  });

  test('without the print option no grid line is printed; headings, centring, landscape and the scale are applied', async ({ page }) => {
    await mountView(
      page,
      await buildXlsx({
        styles: STYLES,
        sheets: [
          {
            name: 'Wide',
            xml: worksheetXml(
              `<sheetData><row r="1"><c r="A1" s="1"/><c r="B1"/></row><row r="2"><c r="B2" s="2"/></row></sheetData>` +
                '<printOptions horizontalCentered="1" headings="1"/>' +
                '<pageMargins left="1" right="1" top="1" bottom="1" header="0.3" footer="0.3"/><pageSetup paperSize="1" scale="200" orientation="landscape"/>',
            ),
          },
        ],
      }),
    );
    await fakePrint(page, false);
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    const frame = await frameState(page);
    // US Letter on its side.
    expect(frame.css).toContain('@page{size:792pt 612pt;margin:0}');
    expect(frame.images.map((image) => [image.width, image.height])).toEqual([[1056, 816]]);
    await loadPicture(page, 0);
    // At 200% the block is (40 + 2 x 64) x 2 = 336 px wide, centred between the margins: it starts at (1056 - 336) / 2 = 360.
    // The headings take 80 x 40 px of it; A1 is 128 x 40 px at (440, 136), red; B2 is blue.
    expect(await pixel(page, 440 + 64, 136 + 20)).toEqual(RED);
    expect(await pixel(page, 440 + 128 + 64, 136 + 40 + 20)).toEqual(BLUE);
    expect(await pixel(page, 440 + 128 + 64, 136 + 20)).toEqual(WHITE);
    // The headings: "A" above the first column, "1" left of the first row.
    expect((await count(page, 440 + 40, 96 + 4, 48, 32)).dark).toBeGreaterThan(10);
    expect((await count(page, 360 + 20, 136 + 4, 40, 32)).dark).toBeGreaterThan(5);
    // No grid line: the boundary between the empty cells B1 and (nothing) A2/B2 carries no ink inside B1.
    expect((await count(page, 440 + 128 + 4, 136 + 2, 120, 36)).ink).toBe(0);
    // Nothing left of the centred block and nothing right of it.
    expect((await count(page, 0, 0, 356, 816)).ink).toBe(0);
    expect((await count(page, 700, 0, 356, 816)).ink).toBe(0);
    await closeDialog(page);
    await waitForOutcome(page);
    await expectCleanedUp(page, 1);
  });

  test('shapes are printed; comment marks and the boxes of comments are not', async ({ page }) => {
    const bytes = await drawingWorkbook({
      sheet: `<sheetData><row r="1">${str('A1', 'a')}${str('C1', 'c')}</row></sheetData>` + A4_SETUP,
      // A green rectangle over B3:C6, below everything else, so that the used range ends at C6.
      drawing: twoCell([1, 2], [3, 6], sp({ props: solid('00B050') })),
      comments: commentsXml(['Ada'], comment('C1', 0, '<r><t>Always shown</t></r>') + comment('A1', 0, '<r><t>Hidden</t></r>')),
      vml: vmlXml(vmlNote(0, 2, { visible: true }) + vmlNote(0, 0)),
    });
    await mountView(page, bytes);
    // On screen the cell with a comment has its red corner and the box of the visible comment is drawn.
    expect(await page.evaluate(() => (window as any).__view.workbook.sheets[0].comments.length)).toBe(2);
    await fakePrint(page, false);
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    await loadPicture(page, 0);
    // The rectangle: 128 x 80 px at B3.
    expect((await count(page, LEFT + 64 + 4, TOP + 40 + 4, 120, 72)).green).toBe(120 * 72);
    // No red corner in A1 or C1, and no yellow box right of C1: only the two letters are on the first rows.
    expect((await count(page, LEFT, TOP, 600, 38)).red).toBe(0);
    expect((await count(page, LEFT + 3 * 64 + 2, 0, 400, 200)).ink).toBe(0);
    expect((await count(page, LEFT + 50, TOP, 14, 8)).ink).toBe(0);
    // The picture is one page: the box of the comment (it would reach to column F and beyond) did not widen the used range.
    expect((await frameState(page)).images.length).toBe(1);
    await closeDialog(page);
    await waitForOutcome(page);
    await expectCleanedUp(page, 1);
  });

  test("sheets: 'all' prints every visible worksheet on its own paper; the active sheet alone is the default", async ({ page }) => {
    const bytes = await buildXlsx({
      sheets: [
        { name: 'One', xml: worksheetXml(numbers(60, 2) + A4_SETUP) },
        { name: 'Hidden', xml: worksheetXml(numbers(2, 2) + A4_SETUP), state: 'hidden' },
        { name: 'Empty', xml: worksheetXml('<sheetData/>') },
        { name: 'Letter', xml: worksheetXml(numbers(2, 2) + '<pageSetup paperSize="1" orientation="landscape"/>') },
      ],
    });
    await mountView(page, bytes, { sheet: 3 });
    await fakePrint(page, true);
    await startPrint(page, { dpi: 96 });
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 1, dpi: 96 } });
    await expectCleanedUp(page, 1);

    await page.evaluate(() => {
      (window as any).__prints.length = 0;
      window.print = () => void (window as any).__prints.push({ sheets: document.querySelectorAll('.bym-print-sheet').length });
    });
    await startPrint(page, { dpi: 96, sheets: 'all' });
    await waitForDialog(page);
    const frame = await frameState(page);
    // Two pages of the first sheet and one of the last; the hidden and the empty sheet have none.
    expect(frame.images.map((image) => [image.width, image.height])).toEqual([A4_SCREEN, A4_SCREEN, [1056, 816]]);
    expect(frame.css).toContain(`@page{size:${A4_PT[0]}pt ${A4_PT[1]}pt;margin:0}`);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 3, dpi: 96 } });
    await expectCleanedUp(page, 4);
    // The view still shows the sheet it showed.
    expect(await page.evaluate(() => (window as any).__view.activeSheet)).toBe(3);
  });

  test('print area and print titles of the defined names decide what is printed', async ({ page }) => {
    const bytes = await buildXlsx({
      styles: STYLES,
      workbookTail:
        '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$60</definedName>' +
        '<definedName name="_xlnm.Print_Titles" localSheetId="0">Data!$1:$1</definedName></definedNames>',
      sheets: [
        {
          name: 'Data',
          xml: worksheetXml(
            `<sheetData><row r="1"><c r="A1" s="1"/><c r="B1" s="1"/><c r="C1" s="2"/></row><row r="49"><c r="A49" s="2"/></row><row r="60">${num('A60', 1)}</row><row r="90">${num('D90', 5)}</row></sheetData>` +
              A4_SETUP,
          ),
        },
      ],
    });
    await mountView(page, bytes);
    await fakePrint(page, false);
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    // 60 rows of the area: two pages, not the 90 rows and four columns in use.
    expect((await frameState(page)).images.length).toBe(2);
    await loadPicture(page, 0);
    // Column C is outside the area: its blue cell is not printed.
    expect((await count(page, 0, 0, A4_SCREEN[0], A4_SCREEN[1])).blue).toBe(0);
    expect(await pixel(page, LEFT + 32, TOP + 10)).toEqual(RED);
    // The second page starts with the title row (red), and row 49 (blue) follows it.
    await loadPicture(page, 1);
    expect(await pixel(page, LEFT + 32, TOP + 10)).toEqual(RED);
    expect(await pixel(page, LEFT + 96, TOP + 10)).toEqual(RED);
    expect(await pixel(page, LEFT + 32, TOP + 30)).toEqual(BLUE);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 2, dpi: 96 } });
    await expectCleanedUp(page, 2);
  });

  test('nothing to print, a sheet beyond the limit, wrong options and a window without print() are refused, each in its way', async ({ page }) => {
    const far = '<sheetData><row r="1048576"><c r="A1048576"><v>1</v></c></row></sheetData>';
    const bytes = await buildXlsx({ sheets: [{ name: 'Empty', xml: worksheetXml('<sheetData/>') }, { name: 'Far', xml: worksheetXml(far) }, { name: 'Data', xml: worksheetXml(numbers(2, 2)) }] });
    await mountView(page, bytes);
    await fakePrint(page, true);
    await startPrint(page);
    let outcome = await waitForOutcome(page);
    expect([outcome.error.name, outcome.error.code]).toEqual(['DocumentPrintError', 'empty']);
    expect((await notices(page)).pop()).toEqual({ message: 'There is nothing to print', kind: 'error' });

    await page.evaluate(() => (window as any).__view.setActiveSheet(1));
    const started = Date.now();
    await startPrint(page);
    outcome = await waitForOutcome(page);
    expect([outcome.error.name, outcome.error.code]).toEqual(['DocumentPrintError', 'too-long']);
    expect(Date.now() - started).toBeLessThan(5000);
    expect((await notices(page)).pop()).toEqual({ message: 'The sheet is too long to print at once', kind: 'error' });
    await expectResponsive(page);
    // All sheets together are too long as well.
    await startPrint(page, { sheets: 'all' });
    expect((await waitForOutcome(page)).error.code).toBe('too-long');

    // A mistake of the caller is not something to tell the user.
    await page.evaluate(() => (window as any).__view.setActiveSheet(2));
    const before = (await notices(page)).length;
    for (const options of [{ dpi: -1 }, { dpi: 'high' }, { sheets: 'visible' }, { sheets: 3 }]) {
      await startPrint(page, options);
      expect((await waitForOutcome(page)).error.name, JSON.stringify(options)).toBe('RangeError');
    }
    expect((await notices(page)).length).toBe(before);
    expect(await printCalls(page)).toEqual([]);

    await page.evaluate(() => ((window as any).print = undefined));
    await startPrint(page);
    outcome = await waitForOutcome(page);
    expect([outcome.error.name, outcome.error.code]).toEqual(['DocumentPrintError', 'unavailable']);
    expect((await notices(page)).pop()).toEqual({ message: 'Printing is not available here', kind: 'error' });
    await expectCleanedUp(page, 0);
    expect(await page.evaluate(() => (window as any).__view.printing)).toBe(false);
  });

  test('a second call while the pages are prepared joins the first; a call after the dialog opened ends that dialog and prints again', async ({ page }) => {
    await mountView(page, await sixPages());
    await fakePrint(page, false);
    const same = await page.evaluate(() => {
      const view = (window as any).__view;
      const first = view.print({ dpi: 96 });
      const second = view.print({ dpi: 96 });
      (window as any).__first = first;
      first.catch(() => undefined);
      return first === second;
    });
    expect(same).toBe(true);
    await waitForDialog(page);
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
    // The dialog never said that it was closed: a new call says so, and prepares its own pages.
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page, 2);
    expect(await page.evaluate(() => (window as any).__first)).toEqual({ pages: 6, dpi: 96 });
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 6, dpi: 96 } });
    await expectCleanedUp(page, 12);
  });

  test('destroy() while the pages are prepared cancels the job and leaves nothing behind', async ({ page }) => {
    await mountView(page, await sixPages());
    await fakePrint(page, false);
    await whenPagesReady(page, 2, 'destroy');
    await startPrint(page, { dpi: 96 });
    const outcome = await waitForOutcome(page);
    expect([outcome.error.name, outcome.error.code]).toEqual(['DocumentPrintError', 'cancelled']);
    expect(await printCalls(page)).toEqual([]);
    await expectCleanedUp(page);
    // The end of a view is nothing to tell the user.
    expect((await notices(page)).filter((notice) => notice.kind === 'error')).toEqual([]);
    // A view that is gone prints nothing.
    await startPrint(page);
    expect((await waitForOutcome(page)).error.code).toBe('cancelled');
    await page.evaluate(() => ((window as any).__view = undefined));
  });

  test('a change of the workbook while the pages are prepared starts them anew: what is printed is the workbook as it is', async ({ page }) => {
    await mountView(page, await formatted(''));
    await fakePrint(page, false);
    // Once the first (and only) page is ready, a row of text far below the cells is added: the sheet then has two pages.
    await whenPagesReady(page, 1, 'setCellInput', [60, 0, 'late']);
    await startPrint(page, { dpi: 96 });
    await waitForDialog(page);
    expect(await printCalls(page)).toEqual([{ sheets: 2 }]);
    await loadPicture(page, 1);
    // Row 61 is the thirteenth row of the second page.
    expect((await count(page, LEFT + 2, TOP + 12 * 20 + 2, 50, 16)).dark).toBeGreaterThan(10);
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 2, dpi: 96 } });
    // The page of the first try was given back as well.
    await expectCleanedUp(page, 3);
  });

  test('a cell that is being edited is written before the sheet is printed', async ({ page }) => {
    await mountView(page, await buildXlsx({ sheets: [{ name: 'Data', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 1)}</row></sheetData>` + A4_SETUP) }] }));
    await fakePrint(page, true);
    await page.evaluate(() => {
      const view = (window as any).__view;
      view.select({ top: 3, left: 1, bottom: 3, right: 1 });
      view.startEditing('WWWW');
    });
    expect(await page.evaluate(() => (window as any).__view.isEditing)).toBe(true);
    await page.evaluate(() => {
      window.print = () => {
        const image = document.querySelector('.bym-print img') as HTMLImageElement;
        const canvas = document.createElement('canvas');
        canvas.width = image.naturalWidth;
        canvas.height = image.naturalHeight;
        const context = canvas.getContext('2d')!;
        context.drawImage(image, 0, 0);
        (window as any).__picture = context.getImageData(0, 0, canvas.width, canvas.height);
        window.dispatchEvent(new Event('afterprint'));
      };
    });
    await startPrint(page, { dpi: 96 });
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 1, dpi: 96 } });
    expect(await page.evaluate(() => (window as any).__view.isEditing)).toBe(false);
    // B4 holds the text now, on paper too.
    expect((await count(page, LEFT + 64 + 2, TOP + 60 + 2, 60, 16)).dark).toBeGreaterThan(20);
  });
});

test.describe('printing from the editor', () => {
  test('the Print button of the toolbar prints, is busy while the pages are prepared and shows the progress in the status bar', async ({ page }) => {
    await mountEditor(page, await sixPages());
    await fakePrint(page, false);
    const button = page.locator('#host [data-command="print"]');
    await expect(button).toBeVisible();
    await expect(button).toBeEnabled();
    await expect(button).toHaveAttribute('aria-label', 'Print');
    await expect(button).toHaveAttribute('title', 'Print');
    await expect(button).toHaveAttribute('aria-busy', 'false');
    // It stands after Undo and Redo, before the first separator.
    expect(await page.evaluate(() => Array.from(document.querySelectorAll('#host .bym-toolbar > *')).slice(0, 4).map((element) => element.getAttribute('data-command') || element.className))).toEqual([
      'undo',
      'redo',
      'print',
      'bym-tb-sep',
    ]);
    // What the button and the status bar show while the second page is being added.
    await page.evaluate(() => {
      const observer = new MutationObserver(() => {
        if (document.querySelectorAll('.bym-print-sheet').length < 2) return;
        observer.disconnect();
        (window as any).__during = {
          busy: document.querySelector('#host [data-command="print"]')!.getAttribute('aria-busy'),
          status: document.querySelector('#host .bym-status-message')!.textContent,
        };
      });
      observer.observe(document.body, { childList: true, subtree: true });
    });
    await button.click();
    await waitForDialog(page);
    expect(await printCalls(page)).toEqual([{ sheets: 6 }]);
    const during = await page.evaluate(() => (window as any).__during);
    expect(during.busy).toBe('true');
    expect(during.status).toMatch(/^Preparing to print… [12] \/ 6$/);
    // The dialog is open: the button is no longer busy, and the progress has left the status bar.
    await expect(button).toHaveAttribute('aria-busy', 'false');
    await expect(page.locator('#host .bym-status-message')).not.toContainText('/ 6');
    await closeDialog(page);
    await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
    await expect(page.locator('#host .bym-status-message')).toHaveText('The print dialog was closed');
    await expectCleanedUp(page, 6);
    // The host page heard both notices.
    const heard = await page.evaluate(() => (window as any).__events.filter((event: any) => event.type === 'notice').map((event: any) => event.message));
    expect(heard).toEqual(['Preparing to print…', 'The print dialog was closed']);
  });

  test('a read-only editor prints as well; the button is the one control that is not disabled, in Turkish too', async ({ page }) => {
    await mountEditor(page, await formatted(), { readOnly: true, locale: 'tr' });
    await fakePrint(page, true);
    const button = page.locator('#host [data-command="print"]');
    await expect(button).toBeEnabled();
    await expect(button).toHaveAttribute('aria-label', 'Yazdır');
    await expect(page.locator('#host [data-command="bold"]')).toBeDisabled();
    await button.click();
    await expect.poll(async () => (await printCalls(page)).length).toBe(1);
    await expect(page.locator('#host .bym-status-message')).toHaveText('Yazdırma penceresi kapandı');
    await expectCleanedUp(page, 1);
  });

  test('Ctrl+P prints from the grid and from the toolbar instead of letting the browser print the host page', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, 'a shortcut of hardware keyboards; the desktop projects cover it');
    await mountEditor(page, await formatted());
    await fakePrint(page, true);
    await page.evaluate(() => (window as any).__view.focus());
    await page.keyboard.press('ControlOrMeta+p');
    await expect.poll(async () => (await printCalls(page)).length).toBe(1);
    await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
    await page.locator('#host [data-command="print"]').focus();
    await page.keyboard.press('ControlOrMeta+p');
    await expect.poll(async () => (await printCalls(page)).length).toBe(2);
    await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
    await page.locator('#host .bym-formula-input').focus();
    await page.keyboard.press('ControlOrMeta+p');
    await expect.poll(async () => (await printCalls(page)).length).toBe(3);
    await expect.poll(() => page.evaluate(() => (window as any).__view.printing)).toBe(false);
    await expectCleanedUp(page, 3);
  });

  test('mount() gives print() for a workbook; under the print medium only the pages show', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    await page.waitForFunction(() => typeof (window as any).BymOffice === 'object');
    const bytes = await formatted();
    const keys = await page.evaluate(async (input) => {
      const host = document.getElementById('host') || document.body.appendChild(Object.assign(document.createElement('div'), { id: 'host' }));
      if (!host.style.height) host.style.cssText = 'width:600px;max-width:100%;height:400px';
      const mounted = await (window as any).BymOffice.mount(host, { file: new Uint8Array(input), locale: 'en-US' });
      (window as any).__mounted = mounted;
      (window as any).__view = mounted.view;
      return Object.keys(mounted).sort();
    }, Array.from(bytes));
    expect(keys).toEqual(['destroy', 'editor', 'file', 'hasUnsavedChanges', 'kind', 'print', 'save', 'saveNow', 'setAutoSave', 'view', 'workbook']);
    await fakePrint(page, false);
    await page.evaluate(() => {
      const w = window as any;
      w.__printed = undefined;
      w.__mounted.print({ dpi: 96, fileName: 'book.xlsx' }).then((result: unknown) => (w.__printed = { result }));
    });
    await waitForDialog(page);
    await page.emulateMedia({ media: 'print' });
    await expect(page.locator('.bym-print img')).toBeVisible();
    await expect(page.locator('.bym-editor')).toBeHidden();
    const box = await page.evaluate(() => {
      const rect = document.querySelector('.bym-print-sheet')!.getBoundingClientRect();
      return [Math.round(rect.left), Math.round(rect.top), Math.round(rect.width)];
    });
    expect(box).toEqual([0, 0, 794]);
    await page.emulateMedia({ media: 'screen' });
    await closeDialog(page);
    expect(await waitForOutcome(page)).toEqual({ result: { pages: 1, dpi: 96 } });
    await expectCleanedUp(page, 1);
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('the bridge page prints a workbook on a print message and answers when the dialog was closed', async ({ page }) => {
    await page.addInitScript(() => {
      (window as any).__messages = [];
      (window as any).BymOfficeHost = { postMessage: (json: unknown) => (window as any).__messages.push(json) };
    });
    await page.goto('/dist/bymoffice.webview.html');
    const send = (message: unknown) => page.evaluate((json) => (window as any).BymOfficeBridge.receive(json), JSON.stringify(message));
    const messages = () => page.evaluate(() => ((window as any).__messages as string[]).map((entry) => JSON.parse(entry)));
    const answer = async (id: string) => {
      let found: any;
      await expect.poll(async () => (found = (await messages()).filter((message) => message.id === id && message.type !== 'loaded')[0]), { timeout: 60_000 }).toBeTruthy();
      return found;
    };
    const bytes = await buildXlsx({ sheets: [{ name: 'One', xml: worksheetXml(numbers(60, 2) + A4_SETUP) }, { name: 'Two', xml: worksheetXml(numbers(2, 2) + A4_SETUP) }] });
    await send({ type: 'load', id: 'l', data: Buffer.from(bytes).toString('base64'), name: 'book.xlsx', readOnly: true });
    await expect.poll(async () => (await messages()).some((message) => message.type === 'loaded')).toBe(true);
    await page.evaluate(() => {
      (window as any).__prints = [];
      window.print = () => {
        (window as any).__prints.push({ sheets: document.querySelectorAll('body > .bym-print .bym-print-sheet').length });
        window.dispatchEvent(new Event('afterprint'));
      };
    });
    await send({ type: 'print', id: 'p1', dpi: 96 });
    expect(await answer('p1')).toEqual({ type: 'done', id: 'p1', command: 'print' });
    await send({ type: 'print', id: 'p2', dpi: 96, sheets: 'all' });
    expect(await answer('p2')).toEqual({ type: 'done', id: 'p2', command: 'print' });
    expect(await printCalls(page)).toEqual([{ sheets: 2 }, { sheets: 3 }]);
    await send({ type: 'print', id: 'p3', sheets: 'some' });
    expect((await answer('p3')).code).toBe('bad-message');
    await expect(page.locator('.bym-print')).toHaveCount(0);
  });
});
