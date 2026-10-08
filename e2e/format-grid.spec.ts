// Formatting, merging, resizing columns and rows, the cell menu and the
// editing events of the workbook view.
//
// Pointer positions are written out by hand from the default sizes of a sheet
// (header 72 x 22, columns 64 wide, rows 20 high). Every workbook is generated
// by code.
import type { Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { expect, test } from './helpers';

const MAX_ROWS = 1048576;
const MAX_COLS = 16384;
const HEADER_WIDTH = 72;
const HEADER_HEIGHT = 22;
const COLUMN_WIDTH = 64;
const ROW_HEIGHT = 20;
const MOUSE_ONLY = 'mouse input is covered by the desktop projects';
const TOUCH_ONLY = 'touch input is covered by the mobile projects';
const LONG_TEXT = 'a long piece of text that does not fit its column';

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;
const str = (ref: string, text: string) => `<c r="${ref}" t="inlineStr"><is><t>${text}</t></is></c>`;

function once<T>(make: () => Promise<T>): () => Promise<T> {
  let value: Promise<T> | undefined;
  return () => (value ??= make());
}

/** A1 "Name", B1 2, C1 3; A2 a long text; B2 5, C3 7, B4 "x". Default sizes everywhere. */
const formatWorkbook = once(() =>
  buildXlsx({
    sheets: [
      {
        name: 'Data',
        xml: worksheetXml(
          '<sheetData>' +
            `<row r="1">${str('A1', 'Name')}${num('B1', 2)}${num('C1', 3)}</row>` +
            `<row r="2">${str('A2', LONG_TEXT)}${num('B2', 5)}</row>` +
            `<row r="3">${num('C3', 7)}</row>` +
            `<row r="4">${str('B4', 'x')}</row>` +
            '</sheetData>',
        ),
      },
    ],
  }),
);

interface MountOptions {
  locale?: string;
  readOnly?: boolean;
  focusScope?: string;
}

async function open(page: Page): Promise<void> {
  await page.goto('/e2e/pages/view.html');
  await page.waitForFunction(() => typeof (window as any).__mountView === 'function');
}

async function show(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.evaluate(
    async ({ base64, options }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountView(data, options);
    },
    { base64: Buffer.from(bytes).toString('base64'), options: { locale: 'en-US', ...options } },
  );
  // Inside a closed shadow root no locator reaches the canvas.
  if ((options as { shadow?: string }).shadow) await page.waitForFunction(() => (window as any).__shadow.querySelector('.bym-canvas') !== null);
  else await expect(page.locator('.bym-canvas')).toBeVisible();
}

async function mount(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await open(page);
  await show(page, bytes, options);
}

function selection(page: Page): Promise<{ active: string; range: string }> {
  return page.evaluate(() => {
    const api = (window as any).BymOffice;
    const s = (window as any).__view.selection;
    return { active: api.cellName(s.active.row, s.active.col), range: api.rangeName(s.range) };
  });
}

const rawRange = (page: Page): Promise<{ top: number; left: number; bottom: number; right: number }> =>
  page.evaluate(() => (window as any).__view.selection.range);

const select = (page: Page, ref: string) =>
  page.evaluate((ref) => {
    const api = (window as any).BymOffice;
    (window as any).__view.select(api.parseRange(ref));
  }, ref);

const selectColumns = (page: Page, left: number, right: number) =>
  page.evaluate(([left, right, rows]) => (window as any).__view.select({ top: 0, bottom: rows - 1, left, right }), [left, right, MAX_ROWS]);

const selectRows = (page: Page, top: number, bottom: number) =>
  page.evaluate(([top, bottom, cols]) => (window as any).__view.select({ top, bottom, left: 0, right: cols - 1 }), [top, bottom, MAX_COLS]);

interface CellInfo {
  value?: number | string | boolean | null;
  formula?: string;
  shown: string;
  bold?: boolean;
  italic?: boolean;
  underline?: string;
  fill?: string;
  horizontal?: string;
  numFmt?: string;
}

/** A cell of the model with the parts of its format the tests look at. */
function cell(page: Page, ref: string): Promise<CellInfo> {
  return page.evaluate((ref) => {
    const api = (window as any).BymOffice;
    const view = (window as any).__view;
    const at = api.parseCellRef(ref);
    const found = api.getCell(view.workbook.sheets[view.activeSheet], at.row, at.col);
    if (!found) return { shown: '' };
    const style = api.resolveCellStyle(view.workbook.styles, found.style);
    return {
      value: found.value,
      formula: found.formula ? found.formula.text : undefined,
      shown: api.cellDisplayText(view.workbook, found, api.getLocale('en-US')),
      bold: style.font.bold,
      italic: style.font.italic,
      underline: style.font.underline,
      fill: style.fill.type === 'pattern' && style.fill.pattern === 'solid' && style.fill.fg ? style.fill.fg.rgb : undefined,
      horizontal: style.alignment.horizontal,
      numFmt: style.numFmt,
    };
  }, ref);
}

/** Pixel sizes of columns and rows as the layout has them now. */
function sizes(page: Page, cols: number[], rows: number[] = []): Promise<{ cols: number[]; rows: number[] }> {
  return page.evaluate(
    ([cols, rows]) => {
      const api = (window as any).BymOffice;
      const view = (window as any).__view;
      const metrics = api.createSheetMetrics(view.workbook.sheets[view.activeSheet]);
      return { cols: cols.map((c: number) => metrics.cols.size(c)), rows: rows.map((r: number) => metrics.rows.size(r)) };
    },
    [cols, rows],
  );
}

const events = (page: Page, type?: string): Promise<any[]> =>
  page.evaluate((type) => (window as any).__events.filter((event: any) => !type || event.type === type), type);

const view = <T>(page: Page, run: string): Promise<T> => page.evaluate(`(() => { const view = window.__view; return ${run}; })()`) as Promise<T>;

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
}

const canvasBox = async (page: Page): Promise<Box> => (await page.locator('.bym-canvas').boundingBox())!;

/** A point inside a cell of an unscrolled sheet with default sizes. */
function pointIn(box: Box, row: number, col: number, fx = 0.5, fy = 0.5): { x: number; y: number } {
  return { x: box.x + HEADER_WIDTH + COLUMN_WIDTH * (col + fx), y: box.y + HEADER_HEIGHT + ROW_HEIGHT * (row + fy) };
}

/** The right border of a column header and the bottom border of a row header, with default sizes before them. */
const columnBorder = (box: Box, col: number) => ({ x: box.x + HEADER_WIDTH + COLUMN_WIDTH * (col + 1), y: box.y + HEADER_HEIGHT / 2 });
const rowBorder = (box: Box, row: number) => ({ x: box.x + HEADER_WIDTH / 2, y: box.y + HEADER_HEIGHT + ROW_HEIGHT * (row + 1) });

const input = (page: Page) => page.locator('.bym-input');
const live = (page: Page) => page.locator('.bym-view [aria-live]');
const menu = (page: Page) => page.locator('.bym-view .bym-menu');

async function focusGrid(page: Page): Promise<void> {
  await input(page).focus();
  await expect(input(page)).toBeFocused();
}

/** A right click as the browser reports it: a contextmenu event at page coordinates. */
function contextMenuAt(page: Page, x: number, y: number): Promise<boolean> {
  return page.evaluate(
    ([x, y]) => {
      const event = new MouseEvent('contextmenu', { clientX: x, clientY: y, button: 2, bubbles: true, cancelable: true });
      document.querySelector('.bym-canvas')!.dispatchEvent(event);
      return event.defaultPrevented;
    },
    [x, y],
  );
}

/** Colour of one canvas pixel (CSS coordinates relative to the canvas) after a forced redraw. */
function pixel(page: Page, x: number, y: number): Promise<number[]> {
  return page.evaluate(
    ([x, y]) => {
      const view = (window as any).__view;
      view.render();
      const canvas = document.querySelector('.bym-canvas') as HTMLCanvasElement;
      const dpr = view.getViewState().devicePixelRatio;
      return Array.from(canvas.getContext('2d')!.getImageData(Math.round(x * dpr), Math.round(y * dpr), 1, 1).data);
    },
    [x, y],
  );
}

async function drag(page: Page, from: { x: number; y: number }, to: { x: number; y: number }): Promise<void> {
  await page.mouse.move(from.x, from.y);
  await page.mouse.down();
  await page.mouse.move(to.x, to.y, { steps: 5 });
  await page.mouse.up();
}

test.describe('formatting', () => {
  test('format() changes the cells of the selection, the drawing and what a toolbar reads back', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await select(page, 'B2:C3');
    expect(await view(page, 'view.selectionFormat.font.bold')).toBe(false);

    await view(page, 'view.format({ bold: true })');
    expect([(await cell(page, 'B2')).bold, (await cell(page, 'C3')).bold, (await cell(page, 'C2')).bold, (await cell(page, 'B1')).bold]).toEqual([true, true, true, false]);
    // Values are untouched; the empty cells of the range now exist to carry the format.
    expect([(await cell(page, 'B2')).value, (await cell(page, 'C2')).value]).toEqual([5, null]);
    expect(await view(page, 'view.selectionFormat.font.bold')).toBe(true);
    expect(await events(page, 'change')).toEqual([{ type: 'change', sheets: [0] }]);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });

    await view(page, "view.format({ fill: { rgb: 'FFFF0000' } })");
    expect((await cell(page, 'C2')).fill).toBe('FFFF0000');
    expect(await view(page, 'view.selectionFormat.fill.fg.rgb')).toBe('FFFF0000');
    // The empty cell C2 is drawn red; D2 beside it is not.
    const inside = await pixel(page, HEADER_WIDTH + 2.5 * COLUMN_WIDTH, HEADER_HEIGHT + 1.5 * ROW_HEIGHT);
    expect(inside[0]).toBeGreaterThan(200);
    expect(inside[1]).toBeLessThan(80);
    expect(inside[2]).toBeLessThan(80);
    const outside = await pixel(page, HEADER_WIDTH + 3.5 * COLUMN_WIDTH, HEADER_HEIGHT + 1.5 * ROW_HEIGHT);
    expect(outside[1]).toBeGreaterThan(200);

    await view(page, "view.format({ horizontal: 'center', italic: true })");
    expect(await cell(page, 'B2')).toMatchObject({ horizontal: 'center', italic: true, bold: true, fill: 'FFFF0000' });
    expect(await view(page, 'view.selectionFormat.alignment.horizontal')).toBe('center');

    await select(page, 'B1');
    await view(page, "view.format({ numFmt: '0.00' })");
    expect(await cell(page, 'B1')).toMatchObject({ value: 2, shown: '2.00', numFmt: '0.00' });
    await expect(live(page)).toHaveText('B1 2.00');
    expect(await view(page, 'view.selectionFormat.numFmt')).toBe('0.00');

    // Each call is one step back.
    expect(await view(page, 'view.undo()')).toBe(true);
    expect((await cell(page, 'B1')).shown).toBe('2');
    await view(page, 'view.undo()');
    expect(await cell(page, 'B2')).toMatchObject({ bold: true, fill: 'FFFF0000', italic: false });
    await view(page, 'view.undo()');
    await view(page, 'view.undo()');
    expect((await cell(page, 'B2')).bold).toBe(false);
    expect((await cell(page, 'C2')).shown).toBe('');
    expect(await view(page, 'view.history.canUndo')).toBe(false);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });
    expect(await events(page, 'change')).toHaveLength(8);
  });

  test('format() writes an open edit first and does nothing in a read-only view', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await select(page, 'E5');
    await page.keyboard.type('typed');
    await view(page, 'view.format({ bold: true })');
    expect(await view(page, 'view.isEditing')).toBe(false);
    expect(await cell(page, 'E5')).toMatchObject({ value: 'typed', bold: true });

    await view(page, 'view.setReadOnly(true)');
    await view(page, 'view.format({ italic: true })');
    await view(page, 'view.merge()');
    await view(page, 'view.setColumnWidth(0, 0, 30)');
    await view(page, 'view.setRowHeight(0, 0, 40)');
    await view(page, "view.setEditText('no')");
    expect((await cell(page, 'E5')).italic).toBe(false);
    expect(await view(page, 'view.isEditing')).toBe(false);
    expect(await sizes(page, [0], [0])).toEqual({ cols: [COLUMN_WIDTH], rows: [ROW_HEIGHT] });
    // What a toolbar reads is still there.
    expect(await view(page, '[view.selectionFormat.font.bold, view.selectionMerged, view.editText]')).toEqual([true, false, 'typed']);
  });

  test('a format too large to apply is an error for code and an announcement for keys', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await page.evaluate(() => (window as any).__view.select({ top: 0, left: 0, bottom: 99_999, right: 9 }));
    const thrown = await page.evaluate(() => {
      try {
        (window as any).__view.format({ bold: true });
        return 'nothing';
      } catch (error) {
        return (error as Error).name;
      }
    });
    expect(thrown).toBe('RangeError');
    await page.keyboard.press('Control+b');
    await expect(live(page)).toHaveText(/Too many cells/);
    expect(await events(page, 'change')).toEqual([]);
    expect((await cell(page, 'A1')).bold).toBe(false);
  });

  test('Ctrl+B, Ctrl+I and Ctrl+U switch on and off by what the active cell shows', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await select(page, 'B1:C2');
    await page.keyboard.press('Control+b');
    expect([(await cell(page, 'B1')).bold, (await cell(page, 'C2')).bold]).toEqual([true, true]);
    await page.keyboard.press('Control+i');
    await page.keyboard.press('Control+u');
    expect(await cell(page, 'B1')).toMatchObject({ bold: true, italic: true, underline: 'single' });
    await page.keyboard.press('Control+b');
    await page.keyboard.press('Control+u');
    expect(await cell(page, 'C2')).toMatchObject({ bold: false, italic: true, underline: 'none' });
    await page.keyboard.press('Control+i');
    expect((await cell(page, 'B1')).italic).toBe(false);
    expect(await events(page, 'change')).toHaveLength(6);
    await expect(input(page)).toBeFocused();
    expect(await view(page, 'view.isEditing')).toBe(false);

    // While a cell is edited the keys are the editor's.
    await select(page, 'B2');
    await page.keyboard.press('F2');
    const taken = await page.evaluate(() => {
      const element = document.querySelector('.bym-input')!;
      return ['b', 'i', 'u'].map((key) => {
        const event = new KeyboardEvent('keydown', { key, ctrlKey: true, bubbles: true, cancelable: true });
        element.dispatchEvent(event);
        return event.defaultPrevented;
      });
    });
    expect(taken).toEqual([false, false, false]);
    expect(await view(page, 'view.isEditing')).toBe(true);
    await page.keyboard.press('Escape');
    expect((await cell(page, 'B2')).bold).toBe(false);
    expect(await events(page, 'change')).toHaveLength(6);

    // Nor are they a read-only view's.
    await view(page, 'view.setReadOnly(true)');
    await page.keyboard.press('Control+b');
    expect((await cell(page, 'B2')).bold).toBe(false);
  });
});

test.describe('merging', () => {
  test('merge() and unmerge() change how the cells are selected and walked', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await select(page, 'B2:C3');
    expect(await view(page, 'view.selectionMerged')).toBe(false);
    await view(page, 'view.merge()');
    expect(await view(page, 'view.workbook.sheets[0].merges')).toEqual([{ top: 1, left: 1, bottom: 2, right: 2 }]);
    expect(await view(page, 'view.selectionMerged')).toBe(true);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });
    // As in Excel only the first cell keeps its content.
    expect([(await cell(page, 'B2')).value, (await cell(page, 'C3')).value]).toEqual([5, undefined]);
    await expect(live(page)).toHaveText('B2 5');

    // The keyboard meets the merged cell as one.
    await select(page, 'A3');
    await page.keyboard.press('ArrowRight');
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });
    await page.keyboard.press('ArrowRight');
    expect(await selection(page)).toEqual({ active: 'D2', range: 'D2' });
    await select(page, 'C3');
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });

    await view(page, 'view.unmerge()');
    expect(await view(page, '[view.workbook.sheets[0].merges.length, view.selectionMerged]')).toEqual([0, false]);
    await select(page, 'C3');
    expect(await selection(page)).toEqual({ active: 'C3', range: 'C3' });
    // Splitting where nothing is merged is no change.
    const changes = (await events(page, 'change')).length;
    await view(page, 'view.unmerge()');
    expect(await events(page, 'change')).toHaveLength(changes);

    // Undo brings the merged cell back, then the cells it took.
    await view(page, 'view.undo()');
    await select(page, 'C3');
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });
    await view(page, 'view.undo()');
    expect((await cell(page, 'C3')).value).toBe(7);
    await select(page, 'C3');
    expect(await selection(page)).toEqual({ active: 'C3', range: 'C3' });
    await view(page, 'view.redo()');
    await select(page, 'C3');
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });

    const thrown = await page.evaluate((cols) => {
      const view = (window as any).__view;
      view.select({ top: 5, bottom: 5, left: 0, right: cols - 1 });
      try {
        view.merge();
        return 'nothing';
      } catch (error) {
        return (error as Error).name;
      }
    }, MAX_COLS);
    expect(thrown).toBe('RangeError');
  });

  test('a click anywhere in a newly merged cell selects all of it', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    await select(page, 'B2:C3');
    await view(page, 'view.merge()');
    const c3 = pointIn(box, 2, 2);
    const a1 = pointIn(box, 0, 0);
    await page.mouse.click(a1.x, a1.y);
    await page.mouse.click(c3.x, c3.y);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C3' });
    await view(page, 'view.unmerge()');
    await page.mouse.click(a1.x, a1.y);
    await page.mouse.click(c3.x, c3.y);
    expect(await selection(page)).toEqual({ active: 'C3', range: 'C3' });
  });
});

test.describe('sizes from code', () => {
  test('setColumnWidth and setRowHeight move what is under the pointer; undo moves it back', async ({ page, hasTouch }) => {
    await mount(page, await formatWorkbook());
    // 20 characters are 140 pixels, 30 points are 40 pixels.
    await view(page, 'view.setColumnWidth(1, 1, 20)');
    await view(page, 'view.setRowHeight(1, 1, 30)');
    expect(await sizes(page, [0, 1, 2], [0, 1, 2])).toEqual({ cols: [64, 140, 64], rows: [20, 40, 20] });
    expect((await events(page, 'change')).map((event) => event.sheets)).toEqual([[0], [0]]);

    const box = await canvasBox(page);
    const click = async (dx: number, dy: number) => {
      await page.mouse.click(box.x + HEADER_WIDTH + dx, box.y + HEADER_HEIGHT + dy);
      return (await selection(page)).active;
    };
    if (!hasTouch) {
      // 130 pixels into the cells is still column B; 10 pixels past its 140 is column C.
      expect(await click(64 + 130, 10)).toBe('B1');
      expect(await click(64 + 150, 10)).toBe('C1');
      // 35 pixels into row 2 is still row 2; 45 is row 3.
      expect(await click(10, 20 + 35)).toBe('A2');
      expect(await click(10, 20 + 45)).toBe('A3');
    }

    await view(page, 'view.undo()');
    await view(page, 'view.undo()');
    expect(await sizes(page, [1], [1])).toEqual({ cols: [64], rows: [20] });
    if (!hasTouch) {
      expect(await click(64 + 130, 10)).toBe('D1');
      expect(await click(10, 20 + 35)).toBe('A3');
    }
    await view(page, 'view.redo()');
    expect(await sizes(page, [1], [1])).toEqual({ cols: [140], rows: [20] });
    if (!hasTouch) expect(await click(64 + 130, 10)).toBe('B1');

    // Several at once, in either order; undefined is the default again.
    await view(page, 'view.setColumnWidth(3, 1, 5)');
    expect((await sizes(page, [1, 2, 3, 4])).cols).toEqual([35, 35, 35, 64]);
    await view(page, 'view.setColumnWidth(1, 3, undefined)');
    await view(page, 'view.setRowHeight(0, 2, 15)');
    await view(page, 'view.setRowHeight(0, 2, undefined)');
    expect(await sizes(page, [1, 2, 3], [0, 1, 2])).toEqual({ cols: [64, 64, 64], rows: [20, 20, 20] });
    // Width 0 hides.
    await view(page, 'view.setColumnWidth(2, 2, 0)');
    expect((await sizes(page, [2])).cols).toEqual([0]);
    const invalid = await page.evaluate(() => {
      try {
        (window as any).__view.setColumnWidth(0, 0, Number.NaN);
        return 'nothing';
      } catch (error) {
        return (error as Error).name;
      }
    });
    expect(invalid).toBe('RangeError');
  });
});

test.describe('resizing with the mouse', () => {
  test.beforeEach(({ hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
  });

  test('dragging the border of a column header resizes the column and selects nothing', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = columnBorder(box, 1);
    const cursor = () => page.evaluate(() => (document.querySelector('.bym-canvas') as HTMLElement).style.cursor);

    await page.mouse.move(border.x - 20, border.y);
    expect(await cursor()).toBe('');
    await page.mouse.move(border.x - 3, border.y);
    expect(await cursor()).toBe('col-resize');
    await page.mouse.move(border.x + 3, border.y);
    expect(await cursor()).toBe('col-resize');
    await page.mouse.move(border.x + 3, border.y + HEADER_HEIGHT);
    expect(await cursor(), 'over the cells, not the headers').toBe('');

    // From one pixel left of the border to 50 pixels right of it: the column ends where the pointer is.
    await page.mouse.move(border.x - 1, border.y);
    await page.mouse.down();
    await page.mouse.move(border.x + 20, border.y + 30, { steps: 4 });
    const guide = page.locator('.bym-view .bym-resize-guide');
    await expect(guide).toHaveCount(1);
    const line = (await guide.boundingBox())!;
    expect(Math.round(line.x - box.x)).toBe(HEADER_WIDTH + 2 * COLUMN_WIDTH + 20);
    expect(line.height).toBeGreaterThan(box.height - 2);
    expect((await sizes(page, [1])).cols, 'nothing is applied while dragging').toEqual([64]);
    await page.mouse.move(border.x + 50, border.y + 30, { steps: 3 });
    await page.mouse.up();
    await expect(guide).toHaveCount(0);
    expect((await sizes(page, [0, 1, 2])).cols).toEqual([64, 114, 64]);
    expect(await events(page, 'change')).toHaveLength(1);
    // Not a click on the header and not a selection drag.
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });
    expect(await events(page, 'selection')).toEqual([]);

    // Narrower, from the new border.
    await drag(page, { x: border.x + 50, y: border.y }, { x: border.x - 14, y: border.y });
    expect((await sizes(page, [1])).cols).toEqual([50]);
    await view(page, 'view.undo()');
    await view(page, 'view.undo()');
    expect((await sizes(page, [1])).cols).toEqual([64]);
    expect(await view(page, 'view.history.canUndo')).toBe(false);
  });

  test('dragging the border of a row header resizes the row', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = rowBorder(box, 1);
    await page.mouse.move(border.x, border.y + 2);
    expect(await page.evaluate(() => (document.querySelector('.bym-canvas') as HTMLElement).style.cursor)).toBe('row-resize');
    await page.mouse.down();
    await page.mouse.move(border.x + 40, border.y + 15, { steps: 4 });
    const line = (await page.locator('.bym-resize-guide').boundingBox())!;
    expect(Math.round(line.y - box.y)).toBe(HEADER_HEIGHT + 2 * ROW_HEIGHT + 15);
    expect(line.width).toBeGreaterThan(box.width - 2);
    await page.mouse.up();
    expect((await sizes(page, [], [0, 1, 2])).rows).toEqual([20, 35, 20]);
    expect(await view(page, 'view.workbook.sheets[0].rows.get(1).height')).toBe(26.25);
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });

    // The row below starts 15 pixels lower now.
    await page.mouse.click(pointIn(box, 0, 1).x, box.y + HEADER_HEIGHT + 20 + 30);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2' });
  });

  test('a press on a border without a drag changes nothing; beside the border it is a header click', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = columnBorder(box, 1);
    await page.mouse.click(border.x - 2, border.y);
    await page.mouse.move(border.x + 2, border.y);
    await page.mouse.down();
    await page.mouse.move(border.x + 4, border.y);
    await page.mouse.up();
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });
    expect(await events(page)).toEqual([]);
    expect((await sizes(page, [1, 2])).cols).toEqual([64, 64]);

    await page.mouse.click(border.x - 8, border.y);
    expect(await rawRange(page)).toEqual({ top: 0, bottom: MAX_ROWS - 1, left: 1, right: 1 });
  });

  test('selected columns and rows are resized together', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    await selectColumns(page, 1, 3);
    const border = columnBorder(box, 2);
    await drag(page, border, { x: border.x + 36, y: border.y });
    expect((await sizes(page, [0, 1, 2, 3, 4])).cols).toEqual([64, 100, 100, 100, 64]);
    expect(await events(page, 'change')).toHaveLength(1);
    expect(await rawRange(page)).toEqual({ top: 0, bottom: MAX_ROWS - 1, left: 1, right: 3 });

    // A column outside the selection is resized alone.
    const first = columnBorder(box, 0);
    await drag(page, first, { x: first.x + 10, y: first.y });
    expect((await sizes(page, [0, 1, 2, 3])).cols).toEqual([74, 100, 100, 100]);

    await selectRows(page, 1, 2);
    const row = rowBorder(box, 1);
    await drag(page, row, { x: row.x, y: row.y + 10 });
    expect((await sizes(page, [], [0, 1, 2, 3])).rows).toEqual([20, 30, 30, 20]);
    await view(page, 'view.undo()');
    expect((await sizes(page, [], [1, 2])).rows).toEqual([20, 20]);
  });

  test('dragging to nothing hides a column or row; dragging from where it was shows it again', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    const box = await canvasBox(page);
    const border = columnBorder(box, 1);
    await drag(page, border, { x: border.x - 100, y: border.y });
    expect((await sizes(page, [0, 1, 2])).cols).toEqual([64, 0, 64]);
    // Column C sits where B was, and the arrow keys pass over B.
    await page.mouse.click(pointIn(box, 0, 1).x, pointIn(box, 0, 1).y);
    expect(await selection(page)).toEqual({ active: 'C1', range: 'C1' });
    await page.keyboard.press('ArrowLeft');
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });

    // Just right of the border between A and C lies the hidden B.
    const hidden = columnBorder(box, 0);
    await drag(page, { x: hidden.x + 2, y: hidden.y }, { x: hidden.x + 62, y: hidden.y });
    expect((await sizes(page, [0, 1, 2])).cols).toEqual([64, 62, 64]);
    // Just left of it is column A itself.
    await view(page, 'view.undo()');
    await drag(page, { x: hidden.x - 2, y: hidden.y }, { x: hidden.x + 20, y: hidden.y });
    expect((await sizes(page, [0, 1, 2])).cols).toEqual([84, 0, 64]);

    const row = rowBorder(box, 2);
    await drag(page, row, { x: row.x, y: row.y - 60 });
    expect((await sizes(page, [], [1, 2, 3])).rows).toEqual([20, 0, 20]);
    const where = rowBorder(box, 1);
    await drag(page, { x: where.x, y: where.y + 2 }, { x: where.x, y: where.y + 26 });
    expect((await sizes(page, [], [1, 2, 3])).rows).toEqual([20, 26, 20]);
  });

  test('a double click on a border fits the column to its content, or the row to the default', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const expected = await page.evaluate((text) => {
      const context = document.createElement('canvas').getContext('2d')!;
      context.font = (window as any).BymOffice.cssFont ? (window as any).BymOffice.cssFont(undefined) : '14.67px "Calibri", "Carlito", Arial, sans-serif';
      return context.measureText(text).width;
    }, LONG_TEXT);

    const first = columnBorder(box, 0);
    await page.mouse.dblclick(first.x, first.y);
    const wide = (await sizes(page, [0])).cols[0];
    expect(wide).toBeGreaterThan(200);
    // The text and a little room around it; the font of the measurement here may differ slightly.
    expect(Math.abs(wide - expected)).toBeLessThan(40);
    expect(await view(page, 'view.isEditing')).toBe(false);
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });
    expect(await events(page, 'change')).toHaveLength(1);

    // Column B holds "2", "5" and "x": made wide, it comes back narrow.
    await view(page, 'view.setColumnWidth(1, 1, 40)');
    const border = { x: box.x + HEADER_WIDTH + wide + 280, y: first.y };
    expect((await sizes(page, [1])).cols).toEqual([280]);
    await page.mouse.dblclick(border.x, border.y);
    const narrow = (await sizes(page, [1])).cols[0];
    expect(narrow).toBeGreaterThan(8);
    expect(narrow).toBeLessThan(40);

    // An empty column goes back to the default.
    await view(page, 'view.setColumnWidth(5, 5, 3)');
    await view(page, 'view.setColumnWidth(0, 1, undefined)');
    const empty = { x: box.x + HEADER_WIDTH + 5 * COLUMN_WIDTH + 21, y: first.y };
    expect((await sizes(page, [5])).cols).toEqual([21]);
    await page.mouse.dblclick(empty.x, empty.y);
    expect((await sizes(page, [5])).cols).toEqual([64]);

    await view(page, 'view.setRowHeight(1, 1, 45)');
    await page.mouse.dblclick(box.x + HEADER_WIDTH / 2, box.y + HEADER_HEIGHT + ROW_HEIGHT + 60);
    expect((await sizes(page, [], [1])).rows).toEqual([20]);
  });

  test('a read-only view does not resize', async ({ page }) => {
    await mount(page, await formatWorkbook(), { readOnly: true });
    const box = await canvasBox(page);
    const border = columnBorder(box, 1);
    await page.mouse.move(border.x - 2, border.y);
    expect(await page.evaluate(() => (document.querySelector('.bym-canvas') as HTMLElement).style.cursor)).toBe('');
    await drag(page, { x: border.x - 2, y: border.y }, { x: border.x + 40, y: border.y });
    await page.mouse.dblclick(border.x - 2, border.y);
    expect((await sizes(page, [1])).cols).toEqual([64]);
    expect(await events(page, 'change')).toEqual([]);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
  });
});

test.describe('resizing with a finger', () => {
  const finger = (page: Page, type: string, x: number, y: number) =>
    page.evaluate(
      ([type, x, y]) => {
        document
          .querySelector('.bym-canvas')!
          .dispatchEvent(new PointerEvent(type as string, { pointerId: 71, pointerType: 'touch', isPrimary: true, clientX: x as number, clientY: y as number, bubbles: true, cancelable: true }));
      },
      [type, x, y] as [string, number, number],
    );

  test('a finger within reach of a border resizes; further away it does not', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = columnBorder(box, 0);
    // Eight pixels off the border: out of reach for a mouse, within reach of a finger.
    await finger(page, 'pointerdown', border.x - 8, border.y);
    await finger(page, 'pointermove', border.x + 20, border.y);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(1);
    await finger(page, 'pointermove', border.x + 30, border.y);
    await finger(page, 'pointerup', border.x + 30, border.y);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
    expect((await sizes(page, [0, 1])).cols).toEqual([94, 64]);
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });
    expect(await events(page, 'change')).toHaveLength(1);

    // In the middle of a header a moving finger is a scroll and changes nothing.
    const middle = { x: box.x + HEADER_WIDTH + 94 + 32, y: border.y };
    await finger(page, 'pointerdown', middle.x, middle.y);
    await finger(page, 'pointermove', middle.x + 40, middle.y);
    await finger(page, 'pointerup', middle.x + 40, middle.y);
    expect((await sizes(page, [0, 1])).cols).toEqual([94, 64]);
    expect(await events(page, 'change')).toHaveLength(1);
  });

  test('the browser is kept from scrolling with a finger that is on a border', async ({ page, hasTouch }) => {
    test.skip(!hasTouch, TOUCH_ONLY);
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = columnBorder(box, 0);
    const prevented = await page.evaluate(
      ([x, y, elsewhere]) => {
        const canvas = document.querySelector('.bym-canvas')!;
        return [x, elsewhere].map((at) => {
          // Not every browser lets a script construct touches; the handler only reads the list of them.
          const event = new Event('touchstart', { bubbles: true, cancelable: true });
          Object.defineProperty(event, 'touches', { value: [{ identifier: 1, target: canvas, clientX: at, clientY: y }] });
          canvas.dispatchEvent(event);
          return event.defaultPrevented;
        });
      },
      [border.x - 6, border.y, border.x + 30],
    );
    expect(prevented).toEqual([true, false]);
  });
});

test.describe('cell menu', () => {
  const labels = (page: Page) => menu(page).locator('button');

  test('a right click on a cell selects it and offers the clipboard, clearing, rows and columns', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    const box = await canvasBox(page);
    const c3 = pointIn(box, 2, 2);
    expect(await contextMenuAt(page, c3.x, c3.y)).toBe(true);
    await expect(menu(page)).toBeVisible();
    expect(await selection(page)).toEqual({ active: 'C3', range: 'C3' });
    await expect(labels(page)).toHaveText(['Cut', 'Copy', 'Paste', 'Clear contents', 'Insert row above', 'Insert column left', 'Delete row', 'Delete column', 'Add note', 'Data validation\u2026']);
    await expect(menu(page).locator('[role="separator"]')).toHaveCount(3);
    await expect(menu(page).locator('[role="menuitem"]')).toHaveCount(10);
    await expect(labels(page).nth(0)).toBeFocused();

    // The arrow keys pass over the separator and go round.
    for (let i = 0; i < 4; i++) await page.keyboard.press('ArrowDown');
    await expect(labels(page).nth(4)).toBeFocused();
    await page.keyboard.press('ArrowUp');
    await expect(labels(page).nth(3)).toBeFocused();
    await page.keyboard.press('Enter');
    await expect(menu(page)).toHaveCount(0);
    await expect(input(page)).toBeFocused();
    expect((await cell(page, 'C3')).value).toBeUndefined();
    await view(page, 'view.undo()');
    expect((await cell(page, 'C3')).value).toBe(7);

    // Inside a selection the selection stays, and the menu counts its rows and columns.
    await select(page, 'B2:C4');
    await contextMenuAt(page, c3.x, c3.y);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2:C4' });
    await expect(labels(page)).toHaveText(['Cut', 'Copy', 'Paste', 'Clear contents', 'Insert 3 rows above', 'Insert 2 columns left', 'Delete 3 rows', 'Delete 2 columns', 'Add note', 'Data validation\u2026']);
    await labels(page).nth(4).click();
    await expect(menu(page)).toHaveCount(0);
    expect([(await cell(page, 'B2')).value, (await cell(page, 'B5')).value, (await cell(page, 'C6')).value]).toEqual([undefined, 5, 7]);
    await view(page, 'view.undo()');

    await select(page, 'B2:C4');
    await contextMenuAt(page, c3.x, c3.y);
    await labels(page).nth(7).click();
    expect([(await cell(page, 'A1')).value, (await cell(page, 'B1')).value, (await cell(page, 'B2')).value]).toEqual(['Name', undefined, undefined]);
    await view(page, 'view.undo()');
    expect((await cell(page, 'B2')).value).toBe(5);

    // Whole columns have no rows to insert, whole rows no columns.
    await selectColumns(page, 1, 1);
    await contextMenuAt(page, pointIn(box, 2, 1).x, c3.y);
    await expect(labels(page)).toHaveText(['Cut', 'Copy', 'Paste', 'Clear contents', 'Insert column left', 'Delete column', 'Add note', 'Data validation\u2026']);
    await page.keyboard.press('Escape');
    await expect(menu(page)).toHaveCount(0);
    await expect(input(page)).toBeFocused();
  });

  test('Copy and Cut from the menu go through the same clipboard events as the keyboard', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await page.evaluate(() => {
      (window as any).__clips = [];
      for (const type of ['copy', 'cut']) {
        document.addEventListener(type, (event) => {
          const data = (event as ClipboardEvent).clipboardData;
          (window as any).__clips.push({ type, text: data ? data.getData('text/plain') : null, html: data ? data.getData('text/html') : null, prevented: event.defaultPrevented });
        });
      }
    });
    const box = await canvasBox(page);
    await select(page, 'B1:C2');
    const b1 = pointIn(box, 0, 1);
    await contextMenuAt(page, b1.x, b1.y);
    await labels(page).nth(1).click();
    await expect(menu(page)).toHaveCount(0);
    const clips = () => page.evaluate(() => (window as any).__clips);
    await expect.poll(async () => (await clips()).length).toBe(1);
    expect((await clips())[0]).toMatchObject({ type: 'copy', text: '2\t3\n5\t', prevented: true });
    expect((await clips())[0].html).toContain('data-bym-copy="');
    expect((await cell(page, 'B1')).value).toBe(2);
    await expect(input(page)).toBeFocused();

    await contextMenuAt(page, b1.x, b1.y);
    await labels(page).nth(0).click();
    await expect.poll(async () => (await clips()).length).toBe(2);
    expect((await clips())[1]).toMatchObject({ type: 'cut', text: '2\t3\n5\t' });
    expect([(await cell(page, 'B1')).value, (await cell(page, 'C1')).value, (await cell(page, 'B2')).value]).toEqual([undefined, undefined, undefined]);
    // One step.
    await view(page, 'view.undo()');
    expect([(await cell(page, 'B1')).value, (await cell(page, 'B2')).value]).toEqual([2, 5]);
    expect(await view(page, 'view.history.canUndo')).toBe(false);
  });

  test('Paste from the menu reads the clipboard as text, or says how to paste', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    const box = await canvasBox(page);
    const e5 = pointIn(box, 4, 4);
    const stub = (run: string) =>
      page.evaluate(`Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: () => ${run}, writeText: () => Promise.resolve() } })`);

    // The copied formula comes back as the text it showed: the menu cannot see the mark of the copy.
    await view(page, "view.setCellInput(6, 0, '=B1+C1')");
    await select(page, 'A7');
    const copied = await view<string>(page, 'view.copy().text');
    expect(copied).toBe('5');
    await stub("Promise.resolve('9\\t8\\n=B1+C1')");
    await contextMenuAt(page, e5.x, e5.y);
    await labels(page).nth(2).click();
    await expect.poll(async () => (await cell(page, 'E5')).value).toBe(9);
    expect([(await cell(page, 'F5')).value, (await cell(page, 'E6')).formula]).toEqual([8, 'B1+C1']);
    expect(await selection(page)).toEqual({ active: 'E5', range: 'E5:F6' });

    // Refused by the browser or the user.
    const before = (await events(page, 'change')).length;
    await stub("Promise.reject(new Error('denied'))");
    await contextMenuAt(page, e5.x, e5.y);
    await labels(page).nth(2).click();
    await expect(live(page)).toHaveText('The clipboard cannot be read here. Press Ctrl+V to paste.');
    // Not there at all.
    await select(page, 'A1');
    await page.evaluate(() => Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined }));
    await contextMenuAt(page, e5.x, e5.y);
    await labels(page).nth(2).click();
    await expect(live(page)).toHaveText('The clipboard cannot be read here. Press Ctrl+V to paste.');
    expect(await events(page, 'change')).toHaveLength(before);
  });

  test('where nothing can write the clipboard, Cut and Copy are shown unavailable and passed over', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await page.evaluate(() => {
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined });
      Object.defineProperty(document, 'execCommand', { configurable: true, value: undefined });
    });
    const box = await canvasBox(page);
    const b1 = pointIn(box, 0, 1);
    await contextMenuAt(page, b1.x, b1.y);
    await expect(labels(page).nth(0)).toBeDisabled();
    await expect(labels(page).nth(1)).toBeDisabled();
    await expect(labels(page).nth(1)).toHaveAttribute('aria-disabled', 'true');
    await expect(labels(page).nth(2)).toBeFocused();
    await page.keyboard.press('ArrowUp');
    await expect(labels(page).nth(9)).toBeFocused();
    await page.keyboard.press('Home');
    await expect(labels(page).nth(2)).toBeFocused();
    // A press on an unavailable item does nothing and leaves the menu open.
    await labels(page).nth(0).click({ force: true });
    await expect(menu(page)).toBeVisible();
    expect((await cell(page, 'B1')).value).toBe(2);
    await page.keyboard.press('Escape');
  });

  test('a read-only view offers Copy alone; the labels follow the language', async ({ page }) => {
    await mount(page, await formatWorkbook(), { readOnly: true, locale: 'tr' });
    const box = await canvasBox(page);
    const b1 = pointIn(box, 0, 1);
    await contextMenuAt(page, b1.x, b1.y);
    await expect(labels(page)).toHaveText(['Kopyala']);
    await expect(menu(page).locator('[role="separator"]')).toHaveCount(0);
    await page.keyboard.press('Escape');
    // No menu on a header of a read-only view.
    expect(await contextMenuAt(page, box.x + HEADER_WIDTH / 2, b1.y)).toBe(false);
    await expect(menu(page)).toHaveCount(0);

    await view(page, 'view.setReadOnly(false)');
    await select(page, 'B1:B2');
    await contextMenuAt(page, b1.x, b1.y);
    await expect(labels(page)).toHaveText(['Kes', 'Kopyala', 'Yapıştır', 'İçeriği temizle', 'Üste 2 satır ekle', 'Sola sütun ekle', '2 satırı sil', 'Sütunu sil', 'Not ekle', 'Veri doğrulama\u2026']);
    await view(page, "view.setLocale('en-US')");
    await contextMenuAt(page, b1.x, b1.y);
    await expect(labels(page).nth(3)).toHaveText('Clear contents');
    await page.keyboard.press('Escape');
  });

  test('a right click with the mouse opens the menu at the pointer; a press elsewhere closes it', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const c3 = pointIn(box, 2, 2);
    await page.mouse.click(c3.x, c3.y, { button: 'right' });
    await expect(menu(page)).toBeVisible();
    const at = (await menu(page).boundingBox())!;
    expect(Math.round(at.x)).toBe(Math.round(c3.x));
    expect(Math.round(at.y)).toBe(Math.round(c3.y));
    expect(await selection(page)).toEqual({ active: 'C3', range: 'C3' });
    const a8 = pointIn(box, 7, 0);
    await page.mouse.click(a8.x, a8.y);
    await expect(menu(page)).toHaveCount(0);
    expect(await selection(page)).toEqual({ active: 'A8', range: 'A8' });
    expect(await events(page, 'change')).toEqual([]);

    // Opening it writes an open edit.
    await page.keyboard.type('before');
    await page.mouse.click(c3.x, c3.y, { button: 'right' });
    await expect(menu(page)).toBeVisible();
    expect((await cell(page, 'A8')).value).toBe('before');
    await page.keyboard.press('Escape');
  });
});

test.describe('editing events and the formula bar', () => {
  const editing = (page: Page): Promise<{ editing: boolean; text: string }[]> => page.evaluate(() => (window as any).__editing.slice());
  const clear = (page: Page) => page.evaluate(() => ((window as any).__editing.length = 0));

  test('onEditingChange reports the start, every change and the end', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await select(page, 'B4');
    expect(await editing(page)).toEqual([]);
    expect(await view(page, 'view.editText')).toBe('x');

    await page.keyboard.type('ab');
    expect(await editing(page)).toEqual([
      { editing: true, text: 'a' },
      { editing: true, text: 'ab' },
    ]);
    expect(await view(page, 'view.editText')).toBe('ab');
    await page.keyboard.press('Alt+Enter');
    await page.keyboard.press('Enter');
    // The end names the cell that was edited, with what it holds now.
    expect((await editing(page)).slice(2)).toEqual([
      { editing: true, text: 'ab\n' },
      { editing: false, text: 'ab\n' },
    ]);
    expect((await cell(page, 'B4')).value).toBe('ab\n');

    await clear(page);
    await select(page, 'B1');
    await page.keyboard.press('F2');
    await page.keyboard.press('Escape');
    expect(await editing(page)).toEqual([
      { editing: true, text: '2' },
      { editing: false, text: '2' },
    ]);

    await clear(page);
    await view(page, "view.startEditing('=1')");
    await view(page, 'view.commitEditing()');
    expect(await editing(page)).toEqual([
      { editing: true, text: '=1' },
      { editing: false, text: '=1' },
    ]);
    // Moving the selection is not editing.
    await clear(page);
    await page.keyboard.press('ArrowDown');
    await view(page, "view.setCellInput(9, 9, 'q')");
    expect(await editing(page)).toEqual([]);
  });

  test('setEditText edits from outside the grid without taking the focus', async ({ page }) => {
    await mount(page, await formatWorkbook(), { focusScope: '#shell' });
    const bar = page.locator('#bar');
    await select(page, 'B1');
    await bar.focus();
    await view(page, "view.setEditText('=B2')");
    await expect(bar).toBeFocused();
    expect(await view(page, '[view.isEditing, view.editText]')).toEqual([true, '=B2']);
    // The cell editor shows the same text over the cell.
    const editor = await page.evaluate(() => {
      const element = document.querySelector('.bym-input') as HTMLTextAreaElement;
      const rect = element.getBoundingClientRect();
      const canvas = document.querySelector('.bym-canvas')!.getBoundingClientRect();
      return { value: element.value, opacity: getComputedStyle(element).opacity, x: Math.round(rect.left - canvas.left), y: Math.round(rect.top - canvas.top) };
    });
    expect(editor).toEqual({ value: '=B2', opacity: '1', x: HEADER_WIDTH + COLUMN_WIDTH - 1, y: HEADER_HEIGHT - 1 });
    await view(page, "view.setEditText('=B2*2')");
    expect(await page.evaluate(() => (window as any).__editing.slice())).toEqual([
      { editing: true, text: '=B2' },
      { editing: true, text: '=B2*2' },
    ]);
    expect((await cell(page, 'B1')).formula).toBeUndefined();

    expect(await view(page, 'view.commitEditing()')).toBe(true);
    expect((await cell(page, 'B1')).formula).toBe('B2*2');
    await expect(bar).toBeFocused();
    expect(await view(page, 'view.editText')).toBe('=B2*2');

    // Escape in the bar is cancelEditing().
    await view(page, "view.setEditText('discard')");
    await view(page, 'view.cancelEditing()');
    expect(await view(page, '[view.isEditing, view.editText]')).toEqual([false, '=B2*2']);

    // focus() hands the keyboard to the grid.
    await view(page, 'view.focus()');
    await expect(input(page)).toBeFocused();
    await page.keyboard.press('ArrowDown');
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2' });
  });

  test('an edit goes on while the focus is inside the focus scope and is written when it leaves', async ({ page }) => {
    await mount(page, await formatWorkbook(), { focusScope: '#shell' });
    await focusGrid(page);
    const bar = page.locator('#bar');
    await select(page, 'E5');
    await page.keyboard.type('one');
    await bar.focus();
    await expect(bar).toBeFocused();
    expect(await view(page, '[view.isEditing, view.editText]')).toEqual([true, 'one']);
    expect((await cell(page, 'E5')).value).toBeUndefined();
    // Back into the cell editor and out again: still the same edit.
    await input(page).focus();
    await page.keyboard.type('!');
    await bar.focus();
    expect(await view(page, 'view.editText')).toBe('one!');

    // Out of the scope altogether.
    await page.locator('#after').focus();
    await expect.poll(() => view(page, 'view.isEditing')).toBe(false);
    expect((await cell(page, 'E5')).value).toBe('one!');
  });

  test('without a focus scope, focus anywhere outside the view writes the edit', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    await select(page, 'E5');
    await page.keyboard.type('two');
    await page.locator('#bar').focus();
    await expect.poll(() => view(page, 'view.isEditing')).toBe(false);
    expect((await cell(page, 'E5')).value).toBe('two');
  });
});

test.describe('teardown', () => {
  test('destroy during a resize, with a focus scope, leaves no listener and no guide', async ({ page }) => {
    await open(page);
    await page.evaluate(() => {
      const ledger = { listeners: 0 };
      (window as any).__ledger = ledger;
      const live = new Map<EventTarget, Map<string, Set<unknown>>>();
      const add = EventTarget.prototype.addEventListener;
      const remove = EventTarget.prototype.removeEventListener;
      EventTarget.prototype.addEventListener = function (this: EventTarget, type: string, listener: any, options?: any) {
        if (listener) {
          let types = live.get(this);
          if (!types) live.set(this, (types = new Map()));
          let set = types.get(type);
          if (!set) types.set(type, (set = new Set()));
          if (!set.has(listener)) {
            set.add(listener);
            ledger.listeners++;
          }
        }
        return add.call(this, type, listener, options);
      };
      EventTarget.prototype.removeEventListener = function (this: EventTarget, type: string, listener: any, options?: any) {
        const set = live.get(this)?.get(type);
        if (set && set.delete(listener)) ledger.listeners--;
        return remove.call(this, type, listener, options);
      };
    });
    await show(page, await formatWorkbook(), { focusScope: '#shell' });
    const result = await page.evaluate(() => {
      const ledger = (window as any).__ledger;
      const view = (window as any).__view;
      const canvas = document.querySelector('.bym-canvas') as HTMLCanvasElement;
      const box = canvas.getBoundingClientRect();
      const mounted = ledger.listeners;
      const fire = (type: string, x: number) =>
        canvas.dispatchEvent(new PointerEvent(type, { pointerId: 1, pointerType: 'mouse', isPrimary: true, button: 0, buttons: 1, clientX: x, clientY: box.top + 11, bubbles: true, cancelable: true }));
      // The border between columns A and B, dragged 40 pixels.
      fire('pointerdown', box.left + 72 + 64);
      fire('pointermove', box.left + 72 + 64 + 40);
      const guides = document.querySelectorAll('.bym-resize-guide').length;
      view.destroy();
      fire('pointermove', box.left + 72 + 64 + 60);
      fire('pointerup', box.left + 72 + 64 + 60);
      return { mounted, guides, left: ledger.listeners, guidesLeft: document.querySelectorAll('.bym-resize-guide').length, changes: (window as any).__events.length };
    });
    expect(result.mounted).toBeGreaterThan(10);
    expect(result).toMatchObject({ guides: 1, left: 0, guidesLeft: 0, changes: 0 });
  });
});

// ---------------------------------------------------------------------------
// Second round: defects found in review, each with the test that shows it.

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

/** Runs an expression with `root` (the document or the shadow root the view is in) and `view`. */
const inRoot = <T>(page: Page, run: string): Promise<T> =>
  page.evaluate(`(() => { const root = window.__shadow || document; const view = window.__view; return ${run}; })()`) as Promise<T>;

const rectOf = (page: Page, selector: string, index = 0): Promise<Rect> =>
  inRoot(page, `(() => { const r = root.querySelectorAll(${JSON.stringify(selector)})[${index}].getBoundingClientRect(); return { x: r.left, y: r.top, width: r.width, height: r.height }; })()`);

/** A contextmenu event on the canvas, wherever the canvas lives. */
const rightClickAt = (page: Page, x: number, y: number): Promise<boolean> =>
  inRoot(
    page,
    `(() => { const e = new MouseEvent('contextmenu', { clientX: ${x}, clientY: ${y}, button: 2, bubbles: true, cancelable: true, composed: true }); root.querySelector('.bym-canvas').dispatchEvent(e); return e.defaultPrevented; })()`,
  );

const focusedText = (page: Page): Promise<string> => inRoot(page, '(root.activeElement ? root.activeElement.textContent : "")');
const menuCount = (page: Page): Promise<number> => inRoot(page, 'root.querySelectorAll(".bym-menu").length');
const itemTexts = (page: Page): Promise<string[]> => inRoot(page, 'Array.from(root.querySelectorAll(".bym-menu button")).map((b) => b.textContent)');

for (const mode of ['open', 'closed'] as const) {
  test.describe(`inside a ${mode} shadow root`, () => {
    test('the cell menu opens, is walked with the arrow keys and takes a press', async ({ page }) => {
      await open(page);
      await show(page, await formatWorkbook(), { shadow: mode } as MountOptions);
      expect(await inRoot(page, '[root !== document, root.mode, root.querySelectorAll("style[data-bymoffice]").length]')).toEqual([true, mode, 1]);
      const box = await rectOf(page, '.bym-canvas');
      const c3 = pointIn(box, 2, 2);

      expect(await rightClickAt(page, c3.x, c3.y)).toBe(true);
      expect(await itemTexts(page)).toEqual(['Cut', 'Copy', 'Paste', 'Clear contents', 'Insert row above', 'Insert column left', 'Delete row', 'Delete column', 'Add note', 'Data validation\u2026']);
      expect(await focusedText(page)).toBe('Cut');
      await page.keyboard.press('ArrowDown');
      expect(await focusedText(page)).toBe('Copy');
      await page.keyboard.press('End');
      expect(await focusedText(page)).toBe('Data validation\u2026');
      await page.keyboard.press('ArrowDown');
      expect(await focusedText(page)).toBe('Cut');
      await page.keyboard.press('Escape');
      expect(await menuCount(page)).toBe(0);
      expect(await inRoot(page, 'root.activeElement === root.querySelector(".bym-input")')).toBe(true);

      // A press on an item reaches it instead of closing the menu as a press "outside".
      await rightClickAt(page, c3.x, c3.y);
      const clear = await rectOf(page, '.bym-menu button', 3);
      await page.mouse.click(clear.x + clear.width / 2, clear.y + clear.height / 2);
      expect(await menuCount(page)).toBe(0);
      expect(await inRoot(page, 'window.BymOffice.getCell(view.workbook.sheets[0], 2, 2)')).toBeUndefined();
      expect(await events(page, 'change')).toHaveLength(1);

      // Chosen with the keyboard.
      await inRoot(page, 'view.undo()');
      await rightClickAt(page, c3.x, c3.y);
      for (let i = 0; i < 3; i++) await page.keyboard.press('ArrowDown');
      expect(await focusedText(page)).toBe('Clear contents');
      await page.keyboard.press('Enter');
      expect(await menuCount(page)).toBe(0);
      expect(await inRoot(page, 'window.BymOffice.getCell(view.workbook.sheets[0], 2, 2)')).toBeUndefined();

      // A press elsewhere in the shadow root closes it, and so does one out on the page.
      await rightClickAt(page, c3.x, c3.y);
      const a8 = pointIn(box, 7, 0);
      await page.mouse.click(a8.x, a8.y);
      expect(await menuCount(page)).toBe(0);
      await rightClickAt(page, c3.x, c3.y);
      expect(await menuCount(page)).toBe(1);
      await page.locator('#after').click();
      expect(await menuCount(page)).toBe(0);
      expect(await events(page, 'change')).toHaveLength(3);
    });
  });
}

test.describe('review: the menu', () => {
  const labels = (page: Page) => menu(page).locator('button');

  test('every row and column item of the cell menu acts on the row or column it names', async ({ page }) => {
    await mount(page, await formatWorkbook());
    await focusGrid(page);
    const box = await canvasBox(page);
    const c3 = pointIn(box, 2, 2);
    const choose = async (index: number, label: string) => {
      await contextMenuAt(page, c3.x, c3.y);
      await expect(labels(page).nth(index)).toHaveText(label);
      await labels(page).nth(index).click();
      await expect(menu(page)).toHaveCount(0);
    };
    const values = async (...refs: string[]) => Promise.all(refs.map(async (ref) => (await cell(page, ref)).value));

    await select(page, 'C3');
    await choose(4, 'Insert row above');
    expect(await values('C3', 'C4', 'B5', 'C1', 'B2')).toEqual([undefined, 7, 'x', 3, 5]);
    await view(page, 'view.undo()');

    await select(page, 'C3');
    await choose(5, 'Insert column left');
    expect(await values('C3', 'D3', 'C1', 'D1', 'B4', 'B1')).toEqual([undefined, 7, undefined, 3, 'x', 2]);
    await view(page, 'view.undo()');

    await select(page, 'C3');
    await choose(6, 'Delete row');
    // Row 3 is gone: row 4 moved up, rows 1 and 2 stayed.
    expect(await values('C3', 'B3', 'B4', 'B2', 'C1')).toEqual([undefined, 'x', undefined, 5, 3]);
    await view(page, 'view.undo()');

    await select(page, 'C3');
    await choose(7, 'Delete column');
    // Column C is gone: B stayed, nothing came from the right.
    expect(await values('C1', 'C3', 'B1', 'B2', 'A1')).toEqual([undefined, undefined, 2, 5, 'Name']);
    await view(page, 'view.undo()');
    expect(await values('C1', 'C3', 'B4')).toEqual([3, 7, 'x']);

    // Several rows and columns: counted from the top-left of the selection.
    await select(page, 'B2:C3');
    await choose(6, 'Delete 2 rows');
    expect(await values('A1', 'B1', 'A2', 'B2', 'C2')).toEqual(['Name', 2, undefined, 'x', undefined]);
    await view(page, 'view.undo()');
    await select(page, 'B2:C3');
    await choose(5, 'Insert 2 columns left');
    expect(await values('B1', 'C1', 'D1', 'E1', 'D2', 'E3')).toEqual([undefined, undefined, 2, 3, 5, 7]);
    await view(page, 'view.undo()');
    await select(page, 'B2:C3');
    await choose(4, 'Insert 2 rows above');
    expect(await values('B1', 'B2', 'B3', 'B4', 'C5', 'B6')).toEqual([2, undefined, undefined, 5, 7, 'x']);
    await view(page, 'view.undo()');
    await select(page, 'B2:C3');
    await choose(7, 'Delete 2 columns');
    expect(await values('A1', 'B1', 'A2', 'B2', 'B4')).toEqual(['Name', undefined, LONG_TEXT, undefined, undefined]);
  });

  for (const height of [300, 200]) {
    test(`in a view ${height} pixels high the whole menu can be reached`, async ({ page }) => {
      await open(page);
      await page.evaluate((height) => {
        document.getElementById('host')!.style.height = `${height}px`;
      }, height);
      await show(page, await formatWorkbook());
      await focusGrid(page);
      const box = await canvasBox(page);
      const c3 = pointIn(box, 2, 2);
      await contextMenuAt(page, c3.x, c3.y);
      const root = await rectOf(page, '.bym-view');
      const inside = async () => {
        const at = await rectOf(page, '.bym-menu');
        expect(at.y, 'top of the menu').toBeGreaterThanOrEqual(root.y - 0.5);
        expect(at.y + at.height, 'bottom of the menu').toBeLessThanOrEqual(root.y + root.height + 0.5);
        return at;
      };
      const at = await inside();
      // With the keyboard the focused item is brought into the visible part of the menu.
      await page.keyboard.press('End');
      await expect(labels(page).last()).toBeFocused();
      const last = await rectOf(page, '.bym-menu button', 8);
      expect(last.y).toBeGreaterThanOrEqual(at.y - 0.5);
      expect(last.y + last.height).toBeLessThanOrEqual(at.y + at.height + 0.5);
      await page.keyboard.press('Home');
      const first = await rectOf(page, '.bym-menu button', 0);
      expect(first.y).toBeGreaterThanOrEqual(at.y - 0.5);

      // And an item near the end takes a press ("Delete column"; the last one opens the field of a note).
      await labels(page).nth(7).click();
      await expect(menu(page)).toHaveCount(0);
      expect([(await cell(page, 'C1')).value, (await cell(page, 'C3')).value, (await cell(page, 'B1')).value]).toEqual([undefined, undefined, 2]);
    });
  }

  test('a menu taller than the view scrolls instead of being cut off', async ({ page }) => {
    await open(page);
    await page.evaluate(() => {
      document.getElementById('host')!.style.height = '160px';
    });
    await show(page, await formatWorkbook());
    const box = await canvasBox(page);
    await contextMenuAt(page, pointIn(box, 1, 1).x, pointIn(box, 1, 1).y);
    const shape = await page.evaluate(() => {
      const element = document.querySelector('.bym-menu') as HTMLElement;
      const buttons = Array.from(element.querySelectorAll('button')).map((b) => b.getBoundingClientRect().height);
      return { scrolls: element.scrollHeight > element.clientHeight + 1, overflow: getComputedStyle(element).overflowY, height: element.offsetHeight, buttons };
    });
    expect(shape.scrolls).toBe(true);
    expect(shape.overflow).toBe('auto');
    expect(shape.height).toBeLessThanOrEqual(160);
    // The items keep their height; they are not squeezed to fit.
    for (const item of shape.buttons) expect(item).toBeGreaterThan(24);
  });
});

test.describe('review: resizing', () => {
  test('Escape gives up a resize that is being dragged', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await formatWorkbook());
    await page.evaluate(() => {
      (window as any).__escapes = 0;
      document.addEventListener('keydown', (event) => {
        if (event.key === 'Escape') (window as any).__escapes++;
      });
    });
    const box = await canvasBox(page);
    const border = columnBorder(box, 1);
    await page.mouse.move(border.x, border.y);
    await page.mouse.down();
    await page.mouse.move(border.x + 40, border.y, { steps: 3 });
    await expect(page.locator('.bym-resize-guide')).toHaveCount(1);
    await page.keyboard.press('Escape');
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
    await page.mouse.move(border.x + 60, border.y);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
    await page.mouse.up();
    expect((await sizes(page, [1])).cols).toEqual([64]);
    expect(await events(page)).toEqual([]);
    expect(await view(page, 'view.history.canUndo')).toBe(false);
    expect(await page.evaluate(() => (window as any).__escapes), 'Escape stays inside the view').toBe(0);

    // The next drag works as usual.
    await drag(page, border, { x: border.x + 16, y: border.y });
    expect((await sizes(page, [1])).cols).toEqual([80]);
  });

  test('every width from 1 to 1785 pixels survives the way to characters and back', async ({ page }) => {
    test.setTimeout(120_000);
    await mount(page, await formatWorkbook());
    const result = await page.evaluate(() => {
      const api = (window as any).BymOffice;
      const view = (window as any).__view;
      const canvas = document.querySelector('.bym-canvas') as HTMLCanvasElement;
      const box = canvas.getBoundingClientRect();
      const origin = box.left + 72;
      const y = box.top + 11;
      const fire = (type: string, x: number) =>
        canvas.dispatchEvent(new PointerEvent(type, { pointerId: 1, pointerType: 'mouse', isPrimary: true, button: 0, buttons: 1, clientX: x, clientY: y, bubbles: true, cancelable: true }));
      const widthOfA = () => api.createSheetMetrics(view.workbook.sheets[0]).cols.size(0);
      // Drags the right border of column A from where it is to `pixels` from the column's start.
      const dragTo = (pixels: number) => {
        // Taken from just beyond the border: a column a few pixels wide has no room of its own to be grabbed by.
        const from = origin + widthOfA();
        fire('pointerdown', from + 1);
        fire('pointermove', from + (pixels >= widthOfA() ? 8 : -8));
        fire('pointermove', origin + pixels);
        fire('pointerup', origin + pixels);
        return widthOfA();
      };
      const wrong: number[] = [];
      for (let pixels = 1; pixels <= 1785; pixels++) if (dragTo(pixels) !== pixels) wrong.push(pixels);
      const stored = view.workbook.sheets[0].columns[0].width;
      return { wrong, top: api.columnWidthToPixels(255), stored, beyond: dragTo(3000) };
    });
    expect(result.wrong).toEqual([]);
    // 255 characters is the widest a file can state, and 1785 pixels is what that is.
    expect(result.top).toBe(1785);
    expect(result.stored).toBeLessThanOrEqual(255);
    expect(result.beyond).toBe(1785);
  });
});

test.describe('review: a finger on a narrow header', () => {
  const finger = (page: Page, type: string, x: number, y: number, id = 81) =>
    page.evaluate(
      ([type, x, y, id]) => {
        document
          .querySelector('.bym-canvas')!
          .dispatchEvent(new PointerEvent(type as string, { pointerId: id as number, pointerType: 'touch', isPrimary: true, clientX: x as number, clientY: y as number, bubbles: true, cancelable: true }));
      },
      [type, x, y, id] as [string, number, number, number],
    );

  test('a tap near the border of a row header selects the row; only a moving finger resizes', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    // The border between rows 2 and 3.
    const border = rowBorder(box, 1);
    await finger(page, 'pointerdown', border.x, border.y - 2);
    await finger(page, 'pointerup', border.x, border.y - 2);
    expect(await rawRange(page)).toEqual({ top: 1, bottom: 1, left: 0, right: MAX_COLS - 1 });
    await finger(page, 'pointerdown', border.x, border.y + 2, 82);
    await finger(page, 'pointerup', border.x, border.y + 2, 82);
    expect(await rawRange(page)).toEqual({ top: 2, bottom: 2, left: 0, right: MAX_COLS - 1 });
    // A wobble of two pixels is still a tap.
    await select(page, 'A1');
    await finger(page, 'pointerdown', border.x, border.y - 2, 83);
    await finger(page, 'pointermove', border.x + 2, border.y - 4, 83);
    await finger(page, 'pointerup', border.x + 2, border.y - 4, 83);
    expect(await rawRange(page)).toEqual({ top: 1, bottom: 1, left: 0, right: MAX_COLS - 1 });
    expect(await events(page, 'change')).toEqual([]);
    expect((await sizes(page, [], [1, 2])).rows).toEqual([20, 20]);

    // The same spot, dragged: the row above the border is resized, and nothing is selected by it.
    await select(page, 'A1');
    await finger(page, 'pointerdown', border.x, border.y - 5, 84);
    await finger(page, 'pointermove', border.x, border.y + 5, 84);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(1);
    await finger(page, 'pointermove', border.x, border.y + 15, 84);
    await finger(page, 'pointerup', border.x, border.y + 15, 84);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
    expect((await sizes(page, [], [0, 1, 2])).rows).toEqual([20, 35, 20]);
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });
    expect(await events(page, 'change')).toHaveLength(1);
  });

  test('the middle third of a narrow header never resizes', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = rowBorder(box, 1);
    // Eight pixels from the border of a 20 pixel row: within a finger's reach by distance, but in the middle third.
    for (const offset of [-8, 8, -10]) {
      await finger(page, 'pointerdown', border.x, border.y + offset, 90 + offset);
      await finger(page, 'pointermove', border.x, border.y + offset + 12, 90 + offset);
      await finger(page, 'pointermove', border.x, border.y + offset + 24, 90 + offset);
      await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
      await finger(page, 'pointerup', border.x, border.y + offset + 24, 90 + offset);
    }
    expect((await sizes(page, [], [0, 1, 2, 3])).rows).toEqual([20, 20, 20, 20]);
    expect(await events(page, 'change')).toEqual([]);
    // Six pixels is the reach on such a row.
    await finger(page, 'pointerdown', border.x, border.y - 6, 99);
    await finger(page, 'pointermove', border.x, border.y + 4, 99);
    await finger(page, 'pointerup', border.x, border.y + 4, 99);
    expect((await sizes(page, [], [1])).rows).toEqual([24]);
  });

  test('a long press near the border of a header opens the header menu', async ({ page }) => {
    await page.clock.install();
    await mount(page, await formatWorkbook());
    const box = await canvasBox(page);
    const border = rowBorder(box, 1);
    await finger(page, 'pointerdown', border.x, border.y - 2);
    await page.clock.fastForward(1000);
    await expect(menu(page).locator('button')).toHaveText(['Insert row above', 'Insert row below', 'Delete row']);
    expect(await rawRange(page)).toEqual({ top: 1, bottom: 1, left: 0, right: MAX_COLS - 1 });
    // The finger moving afterwards does not start a resize under the menu.
    await finger(page, 'pointermove', border.x, border.y + 20);
    await finger(page, 'pointerup', border.x, border.y + 20);
    await expect(page.locator('.bym-resize-guide')).toHaveCount(0);
    expect((await sizes(page, [], [1])).rows).toEqual([20]);
  });
});

test.describe('review: messages', () => {
  test('refusals are announced in the language of the view and handed to onNotice', async ({ page }) => {
    await mount(page, await formatWorkbook(), { locale: 'tr' });
    await focusGrid(page);
    const notices = () => page.evaluate(() => (window as any).__notices.slice());
    await page.evaluate(() => (window as any).__view.select({ top: 0, left: 0, bottom: 99_999, right: 9 }));
    await page.keyboard.press('Control+b');
    await expect(live(page)).toHaveText('Bir kerede biçimlendirilemeyecek kadar çok hücre seçili');
    expect(await notices()).toEqual([{ message: 'Bir kerede biçimlendirilemeyecek kadar çok hücre seçili', kind: 'error' }]);

    // A hint is information, not an error.
    await page.evaluate(() => Object.defineProperty(navigator, 'clipboard', { configurable: true, value: undefined }));
    await select(page, 'A1');
    const box = await canvasBox(page);
    await contextMenuAt(page, pointIn(box, 0, 0).x, pointIn(box, 0, 0).y);
    await menu(page).locator('button', { hasText: 'Yapıştır' }).click();
    await expect(live(page)).toHaveText('Pano buradan okunamıyor. Yapıştırmak için Ctrl+V kullanın.');
    expect((await notices())[1]).toEqual({ message: 'Pano buradan okunamıyor. Yapıştırmak için Ctrl+V kullanın.', kind: 'info' });

    // From code the error itself is thrown, untranslated; no notice.
    const thrown = await page.evaluate(() => {
      const view = (window as any).__view;
      view.select({ top: 0, left: 0, bottom: 99_999, right: 9 });
      try {
        view.format({ bold: true });
        return 'nothing';
      } catch (error) {
        return (error as Error).message;
      }
    });
    expect(thrown).toBe('Too many cells to format at once');
    expect(await notices()).toHaveLength(2);

    await view(page, "view.setLocale('en-US')");
    await page.keyboard.press('Control+b');
    await expect(live(page)).toHaveText('Too many cells to format at once');
    expect((await notices())[2]).toEqual({ message: 'Too many cells to format at once', kind: 'error' });
  });

  test('localizeError knows what the operations refuse, in both languages', async ({ page }) => {
    await mount(page, await formatWorkbook(), { locale: 'tr' });
    const messages = [
      'Too many cells to format at once',
      'Too many rows to format at once',
      'Too many rows to resize at once',
      'Too many rows to change at once',
      'Whole rows or columns cannot be merged',
      'Cannot insert 3 row(s): cells would be pushed off the sheet',
      'Cannot insert 1 col(s): cells would be pushed off the sheet',
      'Cannot insert rows through the array formula in A1:B4',
      'Cannot delete cols through the array formula in C2:C9',
      'Cannot insert rows in a sheet with comments, drawings: not supported yet',
      'Cannot delete cols in a sheet with tables: not supported yet',
      'Cannot paste more than 1000000 cells at once',
      'The workbook has too many cell formats',
      'The range is too large to copy',
      'Something the view has never heard of',
    ];
    const translate = () =>
      page.evaluate((messages) => {
        const view = (window as any).__view;
        return messages.map((message) => view.localizeError(new RangeError(message))).concat([view.localizeError(undefined), view.localizeError('The range is too large to copy')]);
      }, messages);
    expect(await translate()).toEqual([
      'Bir kerede biçimlendirilemeyecek kadar çok hücre seçili',
      'Bir kerede biçimlendirilemeyecek kadar çok satır seçili',
      'Bir kerede boyutlandırılamayacak kadar çok satır seçili',
      'Bir kerede değiştirilemeyecek kadar çok satır seçili',
      'Tam satırlar ya da sütunlar birleştirilemez',
      '3 satır eklenemiyor: hücreler sayfanın dışına itilirdi',
      '1 sütun eklenemiyor: hücreler sayfanın dışına itilirdi',
      'A1:B4 aralığındaki dizi formülünün içinden satır eklenemez',
      'C2:C9 aralığındaki dizi formülünün içinden sütun silinemez',
      'Bu sayfada satır ekleme henüz desteklenmiyor (sayfada bulunan: comments, drawings)',
      'Bu sayfada sütun silme henüz desteklenmiyor (sayfada bulunan: tables)',
      'Bir kerede 1000000 hücreden fazlası yapıştırılamaz',
      'Çalışma kitabında çok fazla hücre biçimi var',
      'Aralık kopyalanamayacak kadar büyük',
      'İşlem tamamlanamadı',
      'İşlem tamamlanamadı',
      'Aralık kopyalanamayacak kadar büyük',
    ]);
    await view(page, "view.setLocale('en')");
    const english = await translate();
    expect(english.slice(0, 5)).toEqual(messages.slice(0, 5));
    expect(english.slice(5, 11)).toEqual([
      'Cannot insert 3 row(s): cells would be pushed off the sheet',
      'Cannot insert 1 column(s): cells would be pushed off the sheet',
      'Cannot insert rows through the array formula in A1:B4',
      'Cannot delete columns through the array formula in C2:C9',
      'Cannot insert rows in a sheet with comments, drawings: not supported yet',
      'Cannot delete columns in a sheet with tables: not supported yet',
    ]);
    expect(english.slice(11)).toEqual([
      'Cannot paste more than 1000000 cells at once',
      'The workbook has too many cell formats',
      'The range is too large to copy',
      'The action could not be completed',
      'The action could not be completed',
      'The range is too large to copy',
    ]);
  });

  test('a refused insert from the menu is announced in Turkish', async ({ page }) => {
    await mount(page, await formatWorkbook(), { locale: 'tr' });
    await focusGrid(page);
    await page.evaluate((rows) => (window as any).__view.setCellInput(rows - 1, 0, 'son'), MAX_ROWS);
    await select(page, 'A1');
    const box = await canvasBox(page);
    await contextMenuAt(page, pointIn(box, 0, 0).x, pointIn(box, 0, 0).y);
    await menu(page).locator('button', { hasText: 'Üste satır ekle' }).click();
    await expect(live(page)).toHaveText('1 satır eklenemiyor: hücreler sayfanın dışına itilirdi');
    expect(await page.evaluate(() => (window as any).__notices.slice())).toEqual([{ message: '1 satır eklenemiyor: hücreler sayfanın dışına itilirdi', kind: 'error' }]);
    expect((await cell(page, 'A1')).value).toBe('Name');
  });
});

test.describe('review: page styles and focus', () => {
  test('page rules for every button, textarea and canvas do not reach into the view', async ({ page }) => {
    await open(page);
    await page.addStyleTag({
      content:
        'button{min-width:120px;display:block;width:100%;min-height:60px;margin:9px;float:left}' +
        'textarea{min-height:80px;min-width:200px;padding:12px;border:3px solid red;margin:5px}' +
        'canvas{width:50%;height:40px;border:5px solid red;margin:7px;max-width:100px}',
    });
    await show(page, await formatWorkbook());
    await focusGrid(page);

    const shape = await page.evaluate(() => {
      const rect = (selector: string, index = 0) => {
        const r = document.querySelectorAll(selector)[index].getBoundingClientRect();
        return { x: r.left, y: r.top, width: r.width, height: r.height };
      };
      const scroller = document.querySelector('.bym-scroller')!;
      return { tab: rect('.bym-tab'), strip: rect('.bym-tabs'), idle: rect('.bym-input'), canvas: rect('.bym-canvas'), client: [scroller.clientWidth, scroller.clientHeight] };
    });
    expect(shape.tab.width).toBeLessThan(110);
    expect(shape.tab.height).toBeGreaterThanOrEqual(32);
    expect(shape.tab.height).toBeLessThanOrEqual(41);
    expect(shape.strip.height).toBeLessThanOrEqual(42);
    expect(shape.tab.y).toBeGreaterThanOrEqual(shape.strip.y);
    expect([shape.idle.width, shape.idle.height]).toEqual([1, 1]);
    expect([Math.round(shape.canvas.width), Math.round(shape.canvas.height)]).toEqual(shape.client);

    await select(page, 'B3');
    await page.keyboard.type('x');
    const box = await canvasBox(page);
    const editor = await page.evaluate(() => {
      const r = document.querySelector('.bym-input')!.getBoundingClientRect();
      return { x: r.left, y: r.top, width: r.width, height: r.height };
    });
    expect(Math.round(editor.x - box.x)).toBe(HEADER_WIDTH + COLUMN_WIDTH - 1);
    expect(Math.round(editor.y - box.y)).toBe(HEADER_HEIGHT + 2 * ROW_HEIGHT - 1);
    expect(editor.width).toBeGreaterThanOrEqual(COLUMN_WIDTH + 1);
    expect(editor.width).toBeLessThan(COLUMN_WIDTH + 12);
    expect(editor.height).toBeGreaterThanOrEqual(ROW_HEIGHT + 1);
    expect(editor.height).toBeLessThan(ROW_HEIGHT + 8);
    await page.keyboard.press('Escape');

    await contextMenuAt(page, pointIn(box, 2, 2).x, pointIn(box, 2, 2).y);
    const items = await page.evaluate(() =>
      Array.from(document.querySelectorAll('.bym-menu button')).map((b) => {
        const r = b.getBoundingClientRect();
        return { x: r.left, y: r.top, width: r.width, height: r.height };
      }),
    );
    expect(items).toHaveLength(10);
    for (let i = 0; i < items.length; i++) {
      expect(items[i].height).toBeLessThan(40);
      expect(items[i].width).toBeLessThan(260);
      expect(items[i].x).toBe(items[0].x);
      if (i > 0) expect(items[i].y).toBeGreaterThanOrEqual(items[i - 1].y + items[i - 1].height - 0.5);
    }
  });

  test('sheet tabs are large enough for a finger where the pointer is coarse', async ({ page, hasTouch, browserName }) => {
    await mount(page, await formatWorkbook());
    const coarse = await page.evaluate(() => window.matchMedia('(pointer: coarse)').matches);
    if (hasTouch && browserName === 'chromium') expect(coarse, 'the emulated phone reports a coarse pointer').toBe(true);
    const height = (await page.locator('.bym-tab').first().boundingBox())!.height;
    expect(height).toBeGreaterThanOrEqual(coarse ? 40 : 32);
    if (!coarse) expect(height).toBeLessThan(40);
  });

  test('focus() shows the focus ring only when asked to', async ({ page }) => {
    await mount(page, await formatWorkbook());
    const ring = () => page.evaluate(() => document.querySelector('.bym-view')!.classList.contains('bym-focus-ring'));
    await page.locator('#after').focus();
    await view(page, 'view.focus()');
    await expect(input(page)).toBeFocused();
    expect(await ring()).toBe(false);

    await page.locator('#after').focus();
    await view(page, 'view.focus({ ring: true })');
    await expect(input(page)).toBeFocused();
    expect(await ring()).toBe(true);
    // Arriving with the keyboard shows it too.
    await page.locator('#after').focus();
    expect(await ring()).toBe(false);
    await input(page).focus();
    expect(await ring()).toBe(true);
    // Calling focus() on a grid that has the focus takes the ring off again: the caller came from a click.
    await view(page, 'view.focus()');
    expect(await ring()).toBe(false);
  });
});
