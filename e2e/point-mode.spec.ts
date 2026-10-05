// Pointing at cells while a formula is typed: a press on a cell, a drag over a
// range, the arrow keys and another sheet put a reference into the formula
// instead of ending the edit.
//
// Pointer positions are written out by hand from the default sizes of a sheet
// (header 72 x 22, columns 64 wide, rows 20 high). Every workbook is generated
// by code.
import type { Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { expect, test } from './helpers';

const HEADER_WIDTH = 72;
const HEADER_HEIGHT = 22;
const COLUMN_WIDTH = 64;
const ROW_HEIGHT = 20;
const MOUSE_ONLY = 'mouse input is covered by the desktop projects';
const TOUCH_ONLY = 'touch input is covered by the mobile projects';
/** Longer than a finger has to rest for a menu (LONG_PRESS_MS in the view, 550). */
const LONG_PRESS_WAIT = 700;

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;

function once<T>(make: () => Promise<T>): () => Promise<T> {
  let value: Promise<T> | undefined;
  return () => (value ??= make());
}

/**
 * "Data": A1=1, B1=2, B2=10, C2=20, B3=30, C3=40, B4=50, C4=60, a merged D5:E6 (5).
 * "Other": A1=3, B2=7 with a filter over A1:B3. "Satış 2026": A1=100.
 */
const pointWorkbook = once(() =>
  buildXlsx({
    sheets: [
      {
        name: 'Data',
        xml: worksheetXml(
          '<sheetData>' +
            `<row r="1">${num('A1', 1)}${num('B1', 2)}</row>` +
            `<row r="2">${num('B2', 10)}${num('C2', 20)}</row>` +
            `<row r="3">${num('B3', 30)}${num('C3', 40)}</row>` +
            `<row r="4">${num('B4', 50)}${num('C4', 60)}</row>` +
            `<row r="5">${num('D5', 5)}</row>` +
            '</sheetData><mergeCells count="1"><mergeCell ref="D5:E6"/></mergeCells>' +
            '<dataValidations count="1"><dataValidation type="whole" showErrorMessage="1" errorTitle="Limit" error="1 to 5 only" sqref="A9">' +
            '<formula1>1</formula1><formula2>5</formula2></dataValidation></dataValidations>',
        ),
      },
      {
        name: 'Other',
        xml: worksheetXml(`<sheetData><row r="1">${num('A1', 3)}${num('B1', 4)}</row><row r="2">${num('A2', 6)}${num('B2', 7)}</row><row r="3">${num('A3', 8)}${num('B3', 9)}</row></sheetData><autoFilter ref="A1:B3"/>`),
      },
      { name: 'Satış 2026', xml: worksheetXml(`<sheetData><row r="1">${num('A1', 100)}</row></sheetData>`) },
    ],
  }),
);

/** A1 = 1; row 1 and column A frozen, column C and row 3 hidden. */
const layoutWorkbook = once(() =>
  buildXlsx({
    sheets: [
      {
        name: 'Layout',
        xml: worksheetXml(
          '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' +
            '<cols><col min="3" max="3" width="9" hidden="1"/></cols>' +
            `<sheetData><row r="1">${num('A1', 1)}</row><row r="3" hidden="1">${num('A3', 3)}</row></sheetData>`,
        ),
      },
    ],
  }),
);

interface MountOptions {
  sheet?: number;
  locale?: string;
  readOnly?: boolean;
}

async function mount(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/view.html');
  await page.waitForFunction(() => typeof (window as any).__mountView === 'function');
  await page.evaluate(
    async ({ base64, options }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountView(data, options);
    },
    { base64: Buffer.from(bytes).toString('base64'), options: { locale: 'en-US', ...options } },
  );
  await expect(page.locator('.bym-canvas')).toBeVisible();
}

async function mountEditor(page: Page, bytes: Uint8Array, options: MountOptions = {}): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(
    async ({ base64, options }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountEditor(data, options, 'host');
      (window as any).__view = (window as any).__editor.view;
    },
    { base64: Buffer.from(bytes).toString('base64'), options: { locale: 'en-US', ...options } },
  );
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
}

const view = <T>(page: Page, run: string): Promise<T> => page.evaluate(`(() => { const view = window.__view; return ${run}; })()`) as Promise<T>;

const input = (page: Page) => page.locator('.bym-input');
const live = (page: Page) => page.locator('.bym-view [aria-live]');
const tab = (page: Page, name: string) => page.locator('.bym-tabs button', { hasText: name });
const addButton = (page: Page) => page.locator('.bym-view > .bym-tab-add');
const menu = (page: Page) => page.locator('.bym-view .bym-menu');
const bar = (page: Page) => page.locator('#host .bym-formula-input');

const editing = (page: Page): Promise<boolean> => view(page, 'view.isEditing');
const activeSheet = (page: Page): Promise<number> => view(page, 'view.activeSheet');
const sheetCount = (page: Page): Promise<number> => view(page, 'view.workbook.sheets.length');

function selection(page: Page): Promise<{ active: string; range: string }> {
  return page.evaluate(() => {
    const api = (window as any).BymOffice;
    const s = (window as any).__view.selection;
    return { active: api.cellName(s.active.row, s.active.col), range: api.rangeName(s.range) };
  });
}

const select = (page: Page, ref: string) =>
  page.evaluate((ref) => {
    const api = (window as any).BymOffice;
    (window as any).__view.select(api.parseRange(ref));
  }, ref);

interface CellInfo {
  value?: number | string | boolean | null;
  formula?: string;
}

function cell(page: Page, ref: string, sheet = 0): Promise<CellInfo> {
  return page.evaluate(
    ([ref, sheet]) => {
      const api = (window as any).BymOffice;
      const view = (window as any).__view;
      const at = api.parseCellRef(ref);
      const found = api.getCell(view.workbook.sheets[sheet as number], at.row, at.col);
      return found ? { value: found.value, formula: found.formula ? found.formula.text : undefined } : {};
    },
    [ref, sheet] as [string, number],
  );
}

/** The widths of the columns of a sheet as the model has them. */
const columns = (page: Page, sheet = 0): Promise<string> => view(page, `JSON.stringify(view.workbook.sheets[${sheet}].columns)`);
const notices = (page: Page): Promise<{ message: string; kind: string }[]> => page.evaluate(() => (window as any).__notices.slice());
const editingEvents = (page: Page): Promise<number> => page.evaluate(() => (window as any).__editing.length);

const changes = (page: Page): Promise<any[]> => page.evaluate(() => (window as any).__events.filter((event: any) => event.type === 'change'));

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

/** A pointer event of a finger on the canvas, at page coordinates. */
function touchEvent(page: Page, type: string, x: number, y: number, id = 31): Promise<void> {
  return page.evaluate(
    ([type, x, y, id]) => {
      document
        .querySelector('.bym-canvas')!
        .dispatchEvent(new PointerEvent(type as string, { pointerId: id as number, pointerType: 'touch', isPrimary: true, clientX: x as number, clientY: y as number, bubbles: true, cancelable: true }));
    },
    [type, x, y, id] as [string, number, number, number],
  );
}

/** Presses a point of the canvas the way the device does: a click, or a tap of a finger. */
async function pressAt(page: Page, hasTouch: boolean | undefined, at: { x: number; y: number }): Promise<void> {
  if (hasTouch) {
    await touchEvent(page, 'pointerdown', at.x, at.y);
    await touchEvent(page, 'pointerup', at.x, at.y);
  } else {
    await page.mouse.click(at.x, at.y);
  }
}

/** Presses the middle of a cell (0-based row and column). */
async function pressCell(page: Page, hasTouch: boolean | undefined, row: number, col: number): Promise<void> {
  await pressAt(page, hasTouch, pointIn(await canvasBox(page), row, col));
}

/** Chooses a tab the way the device does. */
async function pressTab(page: Page, hasTouch: boolean | undefined, name: string): Promise<void> {
  if (hasTouch) await tab(page, name).tap();
  else await tab(page, name).click();
}

/** Selects a cell and begins typing into it, as a user at the keyboard would. */
async function typeIn(page: Page, ref: string, text: string): Promise<void> {
  await input(page).focus();
  await expect(input(page)).toBeFocused();
  await select(page, ref);
  await page.keyboard.type(text);
  await expect(input(page)).toHaveValue(text);
}

/** How many pixels along the top edge of a cell have the colour of the frame around a reference. */
function framePixels(page: Page, row: number, col: number): Promise<number> {
  return page.evaluate(
    ([x, y, width]) => {
      const canvas = document.querySelector('.bym-canvas') as HTMLCanvasElement;
      const ratio = canvas.width / canvas.clientWidth;
      const data = canvas.getContext('2d')!.getImageData(Math.round(x * ratio), Math.round(y * ratio), Math.round(width * ratio), 1).data;
      let found = 0;
      for (let i = 0; i < data.length; i += 4) if (data[i] < 90 && data[i + 1] < 160 && data[i + 2] > 180) found++;
      return found;
    },
    [HEADER_WIDTH + COLUMN_WIDTH * col, HEADER_HEIGHT + ROW_HEIGHT * row, COLUMN_WIDTH],
  );
}

test.describe('pointing at cells with the pointer', () => {
  test('a press on a cell puts its reference into the formula; the next press replaces it', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=B2');
    await expect(live(page)).toHaveText('Reference B2');
    expect(await editing(page)).toBe(true);
    // The selection stays on the cell that is edited, and the focus in its editor.
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });
    await expect(input(page)).toBeFocused();

    // Another press before anything is typed: that cell instead.
    await pressCell(page, hasTouch, 2, 2);
    await expect(input(page)).toHaveValue('=C3');
    await pressCell(page, hasTouch, 2, 2);
    await expect(input(page)).toHaveValue('=C3');

    // After an operator the next press adds a reference.
    await page.keyboard.type('+');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=C3+B2');
    await page.keyboard.press('Enter');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 50, formula: 'C3+B2' });
    expect(await selection(page)).toEqual({ active: 'A7', range: 'A7' });
    // One undo step.
    expect(await view(page, 'view.undo()')).toBe(true);
    expect(await cell(page, 'A6')).toEqual({});
  });

  test('a drag points at a range that follows the pointer', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=SUM(');
    const box = await canvasBox(page);
    const from = pointIn(box, 1, 1);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await expect(input(page)).toHaveValue('=SUM(B2');
    const middle = pointIn(box, 2, 2);
    await page.mouse.move(middle.x, middle.y, { steps: 4 });
    await expect(input(page)).toHaveValue('=SUM(B2:C3');
    const to = pointIn(box, 3, 2);
    await page.mouse.move(to.x, to.y, { steps: 4 });
    await expect(input(page)).toHaveValue('=SUM(B2:C4');
    await page.mouse.up();
    await expect(input(page)).toHaveValue('=SUM(B2:C4');
    await expect(live(page)).toHaveText('Reference B2:C4');
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });

    await page.keyboard.type(')');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A6')).toEqual({ value: 210, formula: 'SUM(B2:C4)' });

    // Dragged up and to the left the range is written the same way.
    await typeIn(page, 'A7', '=SUM(');
    await page.mouse.move(to.x, to.y);
    await page.mouse.down();
    await page.mouse.move(from.x, from.y, { steps: 4 });
    await page.mouse.up();
    await expect(input(page)).toHaveValue('=SUM(B2:C4');
    await page.keyboard.press('Escape');
    expect(await cell(page, 'A7')).toEqual({});
  });

  test('Shift and a press stretch the reference pointed at last', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    const box = await canvasBox(page);
    const first = pointIn(box, 1, 1);
    await page.mouse.click(first.x, first.y);
    await expect(input(page)).toHaveValue('=B2');
    const second = pointIn(box, 2, 2);
    await page.keyboard.down('Shift');
    await page.mouse.click(second.x, second.y);
    await page.keyboard.up('Shift');
    await expect(input(page)).toHaveValue('=B2:C3');
    await page.keyboard.press('Escape');
  });

  test('a press on a header points at the whole column or row', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=SUM(');
    const box = await canvasBox(page);
    // The header of column C, in its middle (the borders resize).
    await pressAt(page, hasTouch, { x: box.x + HEADER_WIDTH + COLUMN_WIDTH * 2.5, y: box.y + HEADER_HEIGHT / 2 });
    await expect(input(page)).toHaveValue('=SUM(C:C');
    // The header of row 3.
    await pressAt(page, hasTouch, { x: box.x + HEADER_WIDTH / 2, y: box.y + HEADER_HEIGHT + ROW_HEIGHT * 2.5 });
    await expect(input(page)).toHaveValue('=SUM(3:3');
    // The corner is no reference.
    await pressAt(page, hasTouch, { x: box.x + HEADER_WIDTH / 2, y: box.y + HEADER_HEIGHT / 2 });
    await expect(input(page)).toHaveValue('=SUM(3:3');
    expect(await editing(page)).toBe(true);
    await pressAt(page, hasTouch, { x: box.x + HEADER_WIDTH + COLUMN_WIDTH * 1.5, y: box.y + HEADER_HEIGHT / 2 });
    await expect(input(page)).toHaveValue('=SUM(B:B');
    await page.keyboard.type(')');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A6')).toEqual({ value: 92, formula: 'SUM(B:B)' });
  });

  test('a drag over headers points at several columns; a border does not resize while a formula is typed', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=SUM(');
    const box = await canvasBox(page);
    const y = box.y + HEADER_HEIGHT / 2;
    await page.mouse.move(box.x + HEADER_WIDTH + COLUMN_WIDTH * 1.5, y);
    await page.mouse.down();
    await page.mouse.move(box.x + HEADER_WIDTH + COLUMN_WIDTH * 3.5, y, { steps: 4 });
    await page.mouse.up();
    await expect(input(page)).toHaveValue('=SUM(B:D');

    // A formula that cannot be written as it stands: the border between two headers does nothing, the
    // editor stays open with the text, nothing is said and no column changes its width.
    const widths = await columns(page);
    const dragBorder = async () => {
      await page.mouse.move(box.x + HEADER_WIDTH + COLUMN_WIDTH * 2, y);
      await page.mouse.down();
      await page.mouse.move(box.x + HEADER_WIDTH + COLUMN_WIDTH * 2 + 30, y, { steps: 3 });
      await page.mouse.up();
    };
    await page.keyboard.type(';');
    await expect(input(page)).toHaveValue('=SUM(B:D;');
    await dragBorder();
    await expect(input(page)).toHaveValue('=SUM(B:D;');
    expect(await editing(page)).toBe(true);
    expect(await columns(page)).toBe(widths);
    expect(await notices(page)).toEqual([]);
    await page.mouse.dblclick(box.x + HEADER_WIDTH + COLUMN_WIDTH * 2, y);
    await expect(input(page)).toHaveValue('=SUM(B:D;');
    expect(await columns(page)).toBe(widths);
    await page.keyboard.press('Escape');

    // A formula that could be written, with the reference pointed at last still in place: the same.
    await typeIn(page, 'A6', '=');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=B2');
    await dragBorder();
    await expect(input(page)).toHaveValue('=B2');
    expect(await editing(page)).toBe(true);
    expect(await columns(page)).toBe(widths);
    expect(await cell(page, 'A6')).toEqual({});
    // The column is where it was: a press where C is still points at C.
    await pressCell(page, hasTouch, 1, 2);
    await expect(input(page)).toHaveValue('=C2');
    await page.keyboard.press('Escape');
    expect(await changes(page)).toEqual([]);

    // Text that is no formula is written by the press, and the border resizes, as before.
    await typeIn(page, 'A6', 'abc');
    await dragBorder();
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 'abc' });
    expect(await columns(page)).not.toBe(widths);
  });

  test('a finger on a border does not resize while a formula is typed, on the sheet of the cell or on another', async ({ page, hasTouch }) => {
    test.skip(!hasTouch, TOUCH_ONLY);
    await mount(page, await pointWorkbook());
    const home = await columns(page, 0);
    const other = await columns(page, 1);
    const box = await canvasBox(page);
    const x = box.x + HEADER_WIDTH + COLUMN_WIDTH * 2;
    const y = box.y + HEADER_HEIGHT / 2;
    const dragBorder = async () => {
      await touchEvent(page, 'pointerdown', x, y);
      await touchEvent(page, 'pointermove', x + 12, y);
      await touchEvent(page, 'pointermove', x + 30, y);
      await touchEvent(page, 'pointerup', x + 30, y);
    };
    await typeIn(page, 'A6', '=');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=B2');
    await dragBorder();
    await expect(input(page)).toHaveValue('=B2');
    expect(await editing(page)).toBe(true);
    expect(await columns(page, 0)).toBe(home);

    // Another sheet is shown: neither its column nor the column of the sheet of the cell is resized.
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    await dragBorder();
    expect(await editing(page)).toBe(true);
    expect(await activeSheet(page)).toBe(1);
    await expect(input(page)).toHaveValue('=B2');
    expect(await columns(page, 0)).toBe(home);
    expect(await columns(page, 1)).toBe(other);
    expect(await notices(page)).toEqual([]);
    await page.keyboard.press('Escape');
    expect(await changes(page)).toEqual([]);

    // Without a formula the same finger resizes.
    await dragBorder();
    expect(await columns(page, 0)).not.toBe(home);
  });

  test('moving inside the cell that is pointed at writes and reports nothing again', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    const before = await editingEvents(page);
    const box = await canvasBox(page);
    const from = pointIn(box, 1, 1, 0.2, 0.3);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await expect(input(page)).toHaveValue('=B2');
    // One reference written: reported once.
    expect(await editingEvents(page)).toBe(before + 1);
    const to = pointIn(box, 1, 1, 0.8, 0.7);
    await page.mouse.move(to.x, to.y, { steps: 12 });
    await page.mouse.move(from.x, from.y, { steps: 12 });
    expect(await editingEvents(page)).toBe(before + 1);
    // The next cell is one more.
    const next = pointIn(box, 1, 2);
    await page.mouse.move(next.x, next.y, { steps: 6 });
    await expect(input(page)).toHaveValue('=B2:C2');
    expect(await editingEvents(page)).toBe(before + 2);
    await page.mouse.up();
    expect(await editingEvents(page)).toBe(before + 2);
    await page.keyboard.press('Escape');
  });

  test('typing while a drag points ends the drag', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    const box = await canvasBox(page);
    const from = pointIn(box, 1, 1);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    const middle = pointIn(box, 2, 2);
    await page.mouse.move(middle.x, middle.y, { steps: 4 });
    await expect(input(page)).toHaveValue('=B2:C3');
    await page.keyboard.type('+');
    await expect(input(page)).toHaveValue('=B2:C3+');
    const to = pointIn(box, 3, 2);
    await page.mouse.move(to.x, to.y, { steps: 4 });
    await page.mouse.up();
    await expect(input(page)).toHaveValue('=B2:C3+');
    expect(await editing(page)).toBe(true);
    // The next press points anew.
    await page.mouse.click(to.x, to.y);
    await expect(input(page)).toHaveValue('=B2:C3+C4');
    await page.keyboard.press('Escape');
  });

  test('a merged cell is pointed at as its top-left cell', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A8', '=');
    // D5:E6 is one cell; the press is on its second row.
    await pressAt(page, hasTouch, pointIn(await canvasBox(page), 5, 3, 0.4));
    await expect(input(page)).toHaveValue('=D5');
    await page.keyboard.type('*2');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A8')).toEqual({ value: 10, formula: 'D5*2' });
  });

  test('the cells pointed at are framed until the edit ends', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    expect(await framePixels(page, 1, 1)).toBe(0);
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=B2');
    await expect.poll(() => framePixels(page, 1, 1)).toBeGreaterThan(10);
    // Typing on takes the frame away: the reference is part of the text now.
    await page.keyboard.type('+');
    await expect.poll(() => framePixels(page, 1, 1)).toBe(0);
    await pressCell(page, hasTouch, 2, 1);
    await expect.poll(() => framePixels(page, 2, 1)).toBeGreaterThan(10);
    await page.keyboard.press('Escape');
    await expect.poll(() => framePixels(page, 2, 1)).toBe(0);
    expect(await cell(page, 'A6')).toEqual({});
  });
});

test.describe('where pointing does nothing', () => {
  test('behind an operand a press ends the edit and selects the cell, as before', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=A1');
    await pressCell(page, hasTouch, 1, 1);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 1, formula: 'A1' });
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2' });

    // Inside a text in quotes too.
    await typeIn(page, 'A7', '="a+');
    await pressCell(page, hasTouch, 2, 1);
    await expect(input(page)).not.toHaveValue('="a+B3');
    // What happens instead: the press tries to write the text, which is no formula that can be read. The
    // editor stays open with it, and that is said.
    await expect(input(page)).toHaveValue('="a+');
    expect(await editing(page)).toBe(true);
    expect(await selection(page)).toEqual({ active: 'A7', range: 'A7' });
    expect((await notices(page)).map((notice) => notice.kind)).toEqual(['error']);
    await page.keyboard.press('Escape');
    expect(await cell(page, 'A7')).toEqual({});
  });

  test('text that is no formula is written by a press, and by a press on a tab', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', 'abc+');
    await pressCell(page, hasTouch, 1, 1);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 'abc+' });
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2' });

    await typeIn(page, 'A7', 'xyz');
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A7')).toEqual({ value: 'xyz' });
    // Nothing keeps the button that adds a sheet from working.
    await expect(addButton(page)).not.toHaveAttribute('aria-disabled', 'true');
  });

  test('a read-only view never points', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook(), { readOnly: true });
    await input(page).focus();
    await select(page, 'A6');
    await page.keyboard.type('=');
    expect(await editing(page)).toBe(false);
    await pressCell(page, hasTouch, 1, 1);
    expect(await editing(page)).toBe(false);
    expect(await selection(page)).toEqual({ active: 'B2', range: 'B2' });
    expect(await cell(page, 'A6')).toEqual({});
    expect(await changes(page)).toEqual([]);
  });

  test('while an input method composes a press adds no reference', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await input(page).evaluate((element) => element.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true })));
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).not.toHaveValue('=B2');
    // What happens instead is what a press always did: the text is written as it stands.
    expect(await cell(page, 'A6')).toEqual({ value: '=' });
    expect(await editing(page)).toBe(false);

    // The end of the composition was never reported (some WebViews do not): the edit that ended took it
    // along, and pointing works again.
    await typeIn(page, 'A7', '=');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=B2');
    await page.keyboard.press('Escape');
  });

  test('the button of a filter is not pressed by pointing at its cell', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook(), { sheet: 1 });
    await typeIn(page, 'D2', '=');
    // The bottom right corner of A1, where its filter button is.
    await pressAt(page, hasTouch, pointIn(await canvasBox(page), 0, 0, 0.9, 0.7));
    await expect(input(page)).toHaveValue('=A1');
    await expect(page.locator('.bym-view .bym-panel')).toHaveCount(0);
    expect(await editing(page)).toBe(true);
    await page.keyboard.press('Enter');
    expect(await cell(page, 'D2', 1)).toEqual({ value: 3, formula: 'A1' });
  });

  test('the wheel scrolls without ending the edit', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    const box = await canvasBox(page);
    await page.mouse.move(box.x + 300, box.y + 200);
    await page.mouse.wheel(0, 400);
    await expect.poll(() => view<number>(page, 'view.getViewState().scrollY')).toBeGreaterThan(0);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('=');
    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
  });
});

test.describe('pointing with the keyboard', () => {
  test('in an edit begun by typing the arrow keys point where the formula takes a reference', async ({ page }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'B6', '=SUM(');
    await page.keyboard.press('ArrowUp');
    // From the cell that is edited.
    await expect(input(page)).toHaveValue('=SUM(B5');
    await expect(live(page)).toHaveText('Reference B5');
    await page.keyboard.press('ArrowUp');
    await expect(input(page)).toHaveValue('=SUM(B4');
    await page.keyboard.press('Shift+ArrowRight');
    await expect(input(page)).toHaveValue('=SUM(B4:C4');
    await page.keyboard.press('Shift+ArrowUp');
    await expect(input(page)).toHaveValue('=SUM(B3:C4');
    expect(await selection(page)).toEqual({ active: 'B6', range: 'B6' });

    // After a separator pointing begins at the edited cell again.
    await page.keyboard.type(',');
    await page.keyboard.press('ArrowLeft');
    await expect(input(page)).toHaveValue('=SUM(B3:C4,A6');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowRight');
    await expect(input(page)).toHaveValue('=SUM(B3:C4,B4');
    await page.keyboard.type(')');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'B6')).toEqual({ value: 230, formula: 'SUM(B3:C4,B4)' });
    expect(await selection(page)).toEqual({ active: 'B7', range: 'B7' });
  });

  test('Ctrl and an arrow jump to the edge of the data', async ({ page }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'B9', '=');
    await page.keyboard.press('Control+ArrowUp');
    await expect(input(page)).toHaveValue('=B4');
    await page.keyboard.press('Control+ArrowUp');
    await expect(input(page)).toHaveValue('=B1');
    await page.keyboard.press('Escape');

    // PageDown points a screen further down, PageUp back; the edited cell stays selected.
    await typeIn(page, 'B9', '=');
    await page.keyboard.press('PageDown');
    await expect(input(page)).toHaveValue(/^=B\d+$/);
    const row = Number((await input(page).inputValue()).substring(2));
    expect(row).toBeGreaterThan(15);
    await page.keyboard.press('PageUp');
    await expect(input(page)).toHaveValue('=B9');
    expect(await selection(page)).toEqual({ active: 'B9', range: 'B9' });
    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
  });

  test('behind an operand, and in an edit begun with F2, the arrow keys move the caret; F2 changes that', async ({ page }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=A1');
    // Behind an operand: the caret moves, as it always did.
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.type('x');
    await expect(input(page)).toHaveValue('=Ax1');
    await page.keyboard.press('ArrowDown');
    expect(await editing(page)).toBe(true);
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });
    await page.keyboard.press('Escape');

    // Begun with F2: the keys are the text's, also where a reference could go.
    await select(page, 'A8');
    await page.keyboard.press('F2');
    expect(await editing(page)).toBe(true);
    await page.keyboard.type('=');
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowDown');
    await expect(input(page)).toHaveValue('=');
    expect(await selection(page)).toEqual({ active: 'A8', range: 'A8' });
    // F2 again: now they point.
    await page.keyboard.press('F2');
    await page.keyboard.press('ArrowUp');
    await expect(input(page)).toHaveValue('=A7');
    // And back.
    await page.keyboard.press('F2');
    await page.keyboard.press('ArrowUp');
    await expect(input(page)).toHaveValue('=A7');
    await page.keyboard.press('Escape');
    expect(await cell(page, 'A8')).toEqual({});
  });

  test('the keys that change sheets show the next sheet without ending the edit', async ({ page }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await page.keyboard.press('Alt+PageDown');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await editing(page)).toBe(true);
    await page.keyboard.press('ArrowDown');
    await expect(input(page)).toHaveValue('=Other!A2');
    await page.keyboard.press('ArrowRight');
    await expect(input(page)).toHaveValue('=Other!B2');
    await page.keyboard.press('Alt+PageUp');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('=Other!B2');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A6')).toEqual({ value: 7, formula: 'Other!B2' });
  });

  test('Ctrl+Z in the editor takes the reference back', async ({ page, hasTouch, browserName }) => {
    test.skip(!!hasTouch, 'the undo of a text field is a key of a desktop keyboard');
    await mount(page, await pointWorkbook());
    for (const how of ['press', 'key']) {
      await typeIn(page, 'A6', '=SUM(');
      if (how === 'press') await pressCell(page, hasTouch, 1, 1);
      else await page.keyboard.press('ArrowUp');
      await expect(input(page)).toHaveValue(how === 'press' ? '=SUM(B2' : '=SUM(A5');
      await page.keyboard.press('ControlOrMeta+z');
      if (browserName === 'webkit') {
        // WebKit takes everything typed into a field in a row for one undo step, with or without a reference
        // pointed at in between: the reference goes, and what was typed before it goes with it. The edit
        // stays open and nothing breaks.
        await expect(input(page)).not.toHaveValue(/[AB]\d/);
        expect((await input(page).inputValue()).startsWith('=SUM(')).toBe(false);
      } else {
        // The reference alone, in one step.
        await expect(input(page)).toHaveValue('=SUM(');
      }
      expect(await editing(page)).toBe(true);
      await page.keyboard.press('Escape');
      expect(await editing(page)).toBe(false);
      expect(await cell(page, 'A6')).toEqual({});
    }
  });

  test('at the edge of the sheet a key that cannot move writes no reference to the cell itself', async ({ page }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A1', '=');
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('Shift+ArrowUp');
    await page.keyboard.press('PageUp');
    await expect(input(page)).toHaveValue('=');
    expect(await editing(page)).toBe(true);
    await page.keyboard.press('ArrowDown');
    await expect(input(page)).toHaveValue('=A2');
    // Once a cell is pointed at, coming back to the edge is a movement like any other.
    await page.keyboard.press('ArrowRight');
    await page.keyboard.press('ArrowUp');
    await expect(input(page)).toHaveValue('=B1');
    await page.keyboard.press('ArrowUp');
    await expect(input(page)).toHaveValue('=B1');
    await page.keyboard.press('Escape');
    expect(await cell(page, 'A1')).toEqual({ value: 1 });
  });

  test('when the edit ends the cell comes back into view, also after typing on', async ({ page }) => {
    await mount(page, await pointWorkbook());
    const scrollY = () => view<number>(page, 'view.getViewState().scrollY');
    for (const key of ['Escape', 'Enter']) {
      await typeIn(page, 'A6', '=');
      await page.keyboard.press('PageDown');
      await page.keyboard.press('PageDown');
      await expect(input(page)).toHaveValue(/^=A\d+$/);
      await page.keyboard.type('+1');
      expect(await scrollY()).toBeGreaterThan(200);
      await page.keyboard.press(key);
      expect(await editing(page)).toBe(false);
      // Row 6 (or 7, after Enter) is at 100 (120) px: in view again.
      await expect.poll(scrollY).toBeLessThanOrEqual(120);
      await select(page, 'A6');
      await page.keyboard.press('Delete');
    }
  });
});

test.describe('pointing at another sheet', () => {
  test('a tab shows its sheet without ending the edit; Enter writes the cell on the sheet it is on', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('=');
    // The editor is out of sight on this sheet.
    await expect(input(page)).toHaveCSS('opacity', '0');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=Other!B2');
    await expect(live(page)).toHaveText('Reference Other!B2');
    await expect.poll(() => framePixels(page, 1, 1)).toBeGreaterThan(10);

    // A sheet whose name needs apostrophes; its cell takes the place of the last reference.
    await pressTab(page, hasTouch, 'Satış 2026');
    await expect.poll(() => activeSheet(page)).toBe(2);
    await expect.poll(() => framePixels(page, 1, 1)).toBe(0);
    await pressCell(page, hasTouch, 0, 0);
    await expect(input(page)).toHaveValue("='Satış 2026'!A1");
    await page.keyboard.type('+');
    await pressTab(page, hasTouch, 'Other');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue("='Satış 2026'!A1+Other!B2");
    // Back on the sheet of the cell a reference has no sheet name.
    await page.keyboard.type('+');
    await pressTab(page, hasTouch, 'Data');
    await expect(input(page)).toHaveCSS('opacity', '1');
    await pressCell(page, hasTouch, 0, 0);
    await expect(input(page)).toHaveValue("='Satış 2026'!A1+Other!B2+A1");
    await pressTab(page, hasTouch, 'Other');

    await page.keyboard.press('Enter');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 108, formula: "'Satış 2026'!A1+Other!B2+A1" });
    expect(await selection(page)).toEqual({ active: 'A7', range: 'A7' });
    // Nothing was written anywhere else.
    expect(await cell(page, 'A6', 1)).toEqual({});
    expect(await cell(page, 'B2', 1)).toEqual({ value: 7 });
    // Undo comes back to the sheet and the cell.
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await view(page, 'view.undo()')).toBe(true);
    expect(await activeSheet(page)).toBe(0);
    expect(await cell(page, 'A6')).toEqual({});
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });
  });

  test('Escape gives up, shows the sheet of the cell again and writes nothing', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await pressTab(page, hasTouch, 'Other');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=Other!B2');
    await page.keyboard.press('Escape');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({});
    expect(await changes(page)).toEqual([]);
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });
    await expect(page.locator('.bym-tabs button[aria-selected="true"]')).toHaveText('Data');
  });

  test('on the other sheet a press where the formula takes no reference writes the cell and selects there', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=A1');
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await editing(page)).toBe(true);
    await pressCell(page, hasTouch, 2, 1);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 1, formula: 'A1' });
    expect(await activeSheet(page)).toBe(1);
    expect(await selection(page)).toEqual({ active: 'B3', range: 'B3' });
  });

  test('while a formula is typed the list of sheets cannot be changed from the tabs', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await expect(addButton(page)).not.toHaveAttribute('aria-disabled', 'true');
    await typeIn(page, 'A6', '=');
    await expect(addButton(page)).toHaveAttribute('aria-disabled', 'true');
    if (hasTouch) await addButton(page).tap({ force: true });
    else await addButton(page).click({ force: true });
    expect(await sheetCount(page)).toBe(3);
    expect(await editing(page)).toBe(true);

    // No menu of a tab.
    if (hasTouch) {
      await tab(page, 'Other').evaluate((element) => {
        const rect = element.getBoundingClientRect();
        const init = { pointerId: 41, pointerType: 'touch', isPrimary: true, clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, bubbles: true, cancelable: true };
        element.dispatchEvent(new PointerEvent('pointerdown', init));
      });
      await page.waitForTimeout(LONG_PRESS_WAIT);
      await tab(page, 'Other').evaluate((element) => element.dispatchEvent(new PointerEvent('pointerup', { pointerId: 41, pointerType: 'touch', bubbles: true })));
    } else {
      await tab(page, 'Other').click({ button: 'right' });
    }
    await expect(menu(page)).toHaveCount(0);
    expect(await editing(page)).toBe(true);
    const longPress = async () => {
      await tab(page, 'Other').evaluate((element) => {
        const rect = element.getBoundingClientRect();
        const init = { pointerId: 41, pointerType: 'touch', isPrimary: true, clientX: rect.left + rect.width / 2, clientY: rect.top + rect.height / 2, bubbles: true, cancelable: true };
        element.dispatchEvent(new PointerEvent('pointerdown', init));
      });
      await page.waitForTimeout(LONG_PRESS_WAIT);
    };

    // Two presses show the sheet and rename nothing.
    await pressTab(page, hasTouch, 'Other');
    await pressTab(page, hasTouch, 'Other');
    await expect(page.locator('.bym-tabs input.bym-tab-input')).toHaveCount(0);
    expect(await editing(page)).toBe(true);
    expect(await activeSheet(page)).toBe(1);
    // Nor does F2 on a tab.
    await tab(page, 'Other').evaluate((element) => element.dispatchEvent(new KeyboardEvent('keydown', { key: 'F2', bubbles: true, cancelable: true })));
    await expect(page.locator('.bym-tabs input.bym-tab-input')).toHaveCount(0);
    await expect(input(page)).toHaveValue('=');

    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
    expect(await activeSheet(page)).toBe(0);
    await expect(addButton(page)).not.toHaveAttribute('aria-disabled', 'true');
    expect(await changes(page)).toEqual([]);
    // Once the edit is over the very same press, held as long, opens the menu of the tab.
    if (hasTouch) {
      await longPress();
      await expect(menu(page)).toBeVisible();
      await tab(page, 'Other').evaluate((element) => element.dispatchEvent(new PointerEvent('pointerup', { pointerId: 41, pointerType: 'touch', bubbles: true })));
    } else {
      await tab(page, 'Other').click({ button: 'right' });
      await expect(menu(page)).toBeVisible();
    }
    await page.keyboard.press('Escape');
    await expect(menu(page)).toHaveCount(0);
    await view(page, 'view.setActiveSheet(0)');
    // And the button adds a sheet again.
    if (hasTouch) await addButton(page).tap();
    else await addButton(page).click();
    expect(await sheetCount(page)).toBe(4);
  });

  test('no menu opens on a sheet that is only shown for pointing', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    await tab(page, 'Other').click();
    await expect.poll(() => activeSheet(page)).toBe(1);
    const at = pointIn(await canvasBox(page), 1, 1);
    await page.mouse.click(at.x, at.y, { button: 'right' });
    await expect(menu(page)).toHaveCount(0);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('=');
    await page.keyboard.press('Escape');
    expect(await activeSheet(page)).toBe(0);
  });
});

test.describe('pointing in other languages and from the formula bar', () => {
  test('in Turkish the reference is the same, and the separator is what the user types', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook(), { locale: 'tr' });
    await typeIn(page, 'A6', '=TOPLA(');
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=TOPLA(B2');
    await expect(live(page)).toHaveText('Başvuru B2');
    await page.keyboard.type(';');
    await pressCell(page, hasTouch, 2, 2);
    await expect(input(page)).toHaveValue('=TOPLA(B2;C3');
    await pressCell(page, hasTouch, 3, 2);
    await expect(input(page)).toHaveValue('=TOPLA(B2;C4');
    await page.keyboard.type(')');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A6')).toEqual({ value: 70, formula: 'SUM(B2,C4)' });
  });

  test('a formula typed in the formula bar is pointed into at the caret of the bar, which keeps the focus', async ({ page, hasTouch }) => {
    await mountEditor(page, await pointWorkbook());
    await select(page, 'A6');
    await bar(page).click();
    await page.keyboard.type('=SUM()');
    await page.keyboard.press('ArrowLeft');
    await pressCell(page, hasTouch, 1, 1);
    await expect(bar(page)).toHaveValue('=SUM(B2)');
    await expect(bar(page)).toBeFocused();
    await expect(input(page)).toHaveValue('=SUM(B2)');
    // The caret is behind the reference: what is typed goes there.
    await page.keyboard.type(';');
    await pressCell(page, hasTouch, 2, 2);
    await expect(bar(page)).toHaveValue('=SUM(B2;C3)');
    await pressCell(page, hasTouch, 3, 2);
    await expect(bar(page)).toHaveValue('=SUM(B2;C4)');
    await expect(bar(page)).toBeFocused();
    expect(await editing(page)).toBe(true);
    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({});

    await select(page, 'A6');
    await bar(page).click();
    await page.keyboard.type('=');
    await pressCell(page, hasTouch, 1, 1);
    await page.keyboard.type('+');
    await pressCell(page, hasTouch, 2, 2);
    await expect(bar(page)).toHaveValue('=B2+C3');
    await page.keyboard.press('Enter');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 50, formula: 'B2+C3' });
  });

  test('from the formula bar another sheet is pointed at, and the bar shows the formula there', async ({ page, hasTouch }) => {
    await mountEditor(page, await pointWorkbook());
    await select(page, 'A6');
    await bar(page).click();
    await page.keyboard.type('=');
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    expect(await editing(page)).toBe(true);
    await expect(bar(page)).toHaveValue('=');
    await expect(bar(page)).toBeFocused();
    await pressCell(page, hasTouch, 1, 1);
    await expect(bar(page)).toHaveValue('=Other!B2');
    await page.keyboard.type('*2');
    await expect(bar(page)).toHaveValue('=Other!B2*2');
    await page.keyboard.press('Enter');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual({ value: 14, formula: 'Other!B2*2' });
    expect(await cell(page, 'A6', 1)).toEqual({});
    await expect(bar(page)).toHaveValue('=Other!B2*2');

    // Escape in the bar on another sheet: back, and nothing written.
    await select(page, 'A7');
    await bar(page).click();
    await page.keyboard.type('=');
    await pressTab(page, hasTouch, 'Other');
    await pressCell(page, hasTouch, 0, 0);
    await expect(bar(page)).toHaveValue('=Other!A1');
    await page.keyboard.press('Escape');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await cell(page, 'A7')).toEqual({});
    await expect(bar(page)).toHaveValue('');
  });

  test('a finger that moves scrolls and adds nothing', async ({ page, hasTouch }) => {
    test.skip(!hasTouch, TOUCH_ONLY);
    await mount(page, await pointWorkbook());
    await typeIn(page, 'A6', '=');
    const box = await canvasBox(page);
    const from = pointIn(box, 1, 1);
    await touchEvent(page, 'pointerdown', from.x, from.y);
    await touchEvent(page, 'pointermove', from.x, from.y + 60);
    await touchEvent(page, 'pointerup', from.x, from.y + 60);
    await expect(input(page)).toHaveValue('=');
    expect(await editing(page)).toBe(true);
    await page.keyboard.press('Escape');
  });
});

test.describe('another sheet is shown and the edit ends some other way', () => {
  /** Types "=" into Data!A6, shows "Other" and points at its B2. */
  async function pointAway(page: Page, hasTouch: boolean | undefined, ref = 'A6'): Promise<void> {
    await typeIn(page, ref, '=');
    await pressTab(page, hasTouch, 'Other');
    await expect.poll(() => activeSheet(page)).toBe(1);
    await pressCell(page, hasTouch, 1, 1);
    await expect(input(page)).toHaveValue('=Other!B2');
  }

  test('a formula that cannot be read is refused on the sheet of the cell, with the editor open', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await page.keyboard.type('+(');
    await page.keyboard.press('Enter');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('=Other!B2+(');
    await expect(input(page)).toHaveCSS('opacity', '1');
    expect(await selection(page)).toEqual({ active: 'A6', range: 'A6' });
    expect((await notices(page)).map((notice) => notice.kind)).toEqual(['error']);
    expect(await cell(page, 'A6')).toEqual({});
    expect(await cell(page, 'A6', 1)).toEqual({});
    expect(await changes(page)).toEqual([]);
    // Put right, it is written.
    await page.keyboard.type('1)');
    await page.keyboard.press('Enter');
    expect(await cell(page, 'A6')).toEqual({ value: 8, formula: 'Other!B2+(1)' });
  });

  test('a value the data validation of the cell stops is refused on the sheet of the cell', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    // A9 takes whole numbers from 1 to 5. A formula is not checked by a data validation (as before), so the
    // text is made a number while the other sheet is still shown.
    await pointAway(page, hasTouch, 'A9');
    await view(page, "view.setEditText('9')");
    expect(await activeSheet(page)).toBe(1);
    await page.keyboard.press('Enter');
    await expect.poll(() => activeSheet(page)).toBe(0);
    expect(await editing(page)).toBe(true);
    await expect(input(page)).toHaveValue('9');
    await expect(input(page)).toHaveCSS('opacity', '1');
    expect(await notices(page)).toEqual([{ message: 'Limit: 1 to 5 only', kind: 'error' }]);
    expect(await cell(page, 'A9')).toEqual({});
    expect(await changes(page)).toEqual([]);
    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
    expect(await activeSheet(page)).toBe(0);
  });

  test('methods of the view, a lost focus and destroy() write the cell of the edit or give it up', async ({ page, hasTouch }) => {
    const written = { value: 7, formula: 'Other!B2' };

    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    // The sheet that is shown for pointing is asked for: the text is written, and that sheet stays.
    await view(page, 'view.setActiveSheet(1)');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await cell(page, 'A6', 1)).toEqual({});
    expect(await activeSheet(page)).toBe(1);
    await expect(page.locator('.bym-tabs button[aria-selected="true"]')).toHaveText('Other');

    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await view(page, 'view.setActiveSheet(2)');
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await activeSheet(page)).toBe(2);

    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await select(page, 'B3');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await activeSheet(page)).toBe(0);
    expect(await selection(page)).toEqual({ active: 'B3', range: 'B3' });
    // The selection of the other sheet is what it was.
    await view(page, 'view.setActiveSheet(1)');
    expect(await selection(page)).toEqual({ active: 'A1', range: 'A1' });

    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await view(page, 'view.setReadOnly(true)');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await activeSheet(page)).toBe(0);
    expect(await view(page, 'view.readOnly')).toBe(true);

    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    expect(await view(page, 'view.insertSheet()')).toBe(3);
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await sheetCount(page)).toBe(4);
    expect(await activeSheet(page)).toBe(3);
    expect(await cell(page, 'A6', 3)).toEqual({});

    // The focus leaves the view.
    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await page.locator('#before').focus();
    await expect.poll(() => editing(page)).toBe(false);
    expect(await cell(page, 'A6')).toEqual(written);
    expect(await activeSheet(page)).toBe(0);

    // destroy() gives the edit up: nothing is written, nothing breaks.
    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    await view(page, 'view.destroy()');
    expect(await cell(page, 'A6')).toEqual({});
    expect(await cell(page, 'A6', 1)).toEqual({});
    expect(await changes(page)).toEqual([]);
    await expect(page.locator('.bym-view')).toHaveCount(0);
  });

  test('a sheet that left the workbook under the view: the edit is given up and that is said', async ({ page, hasTouch }) => {
    await mount(page, await pointWorkbook());
    await pointAway(page, hasTouch);
    // The host page takes the sheet of the cell out of the list behind the back of the view.
    await view(page, 'view.workbook.sheets.splice(0, 1)');
    expect(await view(page, 'view.commitEditing()')).toBe(true);
    expect(await editing(page)).toBe(false);
    expect(await notices(page)).toEqual([{ message: 'The sheet of the cell is gone; what was typed was not written.', kind: 'error' }]);
    // Nothing was written into the sheets that are left.
    expect(await cell(page, 'A6', 0)).toEqual({});
    expect(await cell(page, 'A6', 1)).toEqual({});
    expect(await changes(page)).toEqual([]);
  });
});

test.describe('pointing in a sheet with frozen panes and hidden rows and columns', () => {
  test('a press points at the cell that is drawn there', async ({ page, hasTouch }) => {
    await mount(page, await layoutWorkbook());
    // Column C is hidden: D is drawn in its place. Row 3 is hidden: row 4 is drawn in its place.
    await typeIn(page, 'D1', '=');
    await pressCell(page, hasTouch, 1, 2);
    await expect(input(page)).toHaveValue('=D2');
    await pressCell(page, hasTouch, 2, 1);
    await expect(input(page)).toHaveValue('=B4');
    await pressCell(page, hasTouch, 2, 2);
    await expect(input(page)).toHaveValue('=D4');
    await expect.poll(() => framePixels(page, 2, 2)).toBeGreaterThan(10);

    // Scrolled down: the frozen row is still row 1, the frozen column still A, the rest is further down.
    await page.keyboard.type('+');
    await view(page, 'view.scrollTo(60, 1)');
    await expect.poll(() => view<number>(page, 'view.getViewState().scrollY')).toBeGreaterThan(0);
    expect(await editing(page)).toBe(true);
    await pressCell(page, hasTouch, 0, 1);
    await expect(input(page)).toHaveValue('=D4+B1');
    await pressCell(page, hasTouch, 0, 0);
    await expect(input(page)).toHaveValue('=D4+A1');
    await pressCell(page, hasTouch, 6, 0);
    await expect(input(page)).toHaveValue(/^=D4\+A\d+$/);
    expect(Number((await input(page).inputValue()).substring(5))).toBeGreaterThan(20);
    await pressCell(page, hasTouch, 6, 1);
    await expect(input(page)).toHaveValue(/^=D4\+B\d+$/);
    await page.keyboard.press('Escape');
    expect(await editing(page)).toBe(false);
    expect(await cell(page, 'D1')).toEqual({});
  });

  test('a drag over a hidden column and row takes them into the range; the keys step over them', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, MOUSE_ONLY);
    await mount(page, await layoutWorkbook());
    await typeIn(page, 'D1', '=SUM(');
    const box = await canvasBox(page);
    const from = pointIn(box, 1, 1);
    const to = pointIn(box, 2, 2);
    await page.mouse.move(from.x, from.y);
    await page.mouse.down();
    await page.mouse.move(to.x, to.y, { steps: 4 });
    await page.mouse.up();
    await expect(input(page)).toHaveValue('=SUM(B2:D4');
    await page.keyboard.press('Escape');

    await typeIn(page, 'D2', '=');
    await page.keyboard.press('ArrowLeft');
    await expect(input(page)).toHaveValue('=B2');
    await page.keyboard.press('ArrowDown');
    await expect(input(page)).toHaveValue('=B4');
    await page.keyboard.press('Escape');
  });
});
