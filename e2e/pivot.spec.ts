// Pivot tables in the workbook view: the menu of a cell of a table refreshes it, undo takes the
// refresh back, a refresh that cannot be made says why, a read-only view offers none, and
// mount() refreshes every table with refreshPivots(). The workbooks are made by
// test/pivot-fixtures.ts: the table on "Pivot" (the sheet that is shown) starts in A3 and is a
// stand-in that is out of date, so that a refresh changes it.
//
// Positions are written out by hand from the default sizes of a sheet (header 72 x 22, columns
// 64 wide, rows 20 high), within columns A to D so that a phone shows them.
import type { Page } from '@playwright/test';
import { ZipArchive } from '../src/index';
import { pivotBook, type PivotSpec } from '../test/pivot-fixtures';
import { expect, test } from './helpers';

const HEADER_WIDTH = 72;
const HEADER_HEIGHT = 22;
const COLUMN_WIDTH = 64;
const ROW_HEIGHT = 20;

const DATA: PivotSpec['data'] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
];

async function mount(page: Page, bytes: Uint8Array, options: Record<string, unknown> = {}): Promise<void> {
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

const view = (page: Page, code: string) => page.evaluate((code) => new Function('view', `return ${code}`)((window as any).__view), code);

/** The values of column A and B of "Pivot" from row 3 to row 8. */
const body = (page: Page) =>
  view(
    page,
    '(() => { const sheet = view.workbook.sheets[1]; const out = []; for (let r = 2; r <= 7; r++) { const row = sheet.rows.get(r); out.push([0, 1].map((c) => (row && row.cells[c] ? row.cells[c].value : null))); } return out; })()',
  );

async function openCellMenu(page: Page, col: number, row: number): Promise<void> {
  await page.evaluate(
    ([x, y]) => {
      const canvas = document.querySelector('.bym-canvas') as HTMLElement;
      const rect = canvas.getBoundingClientRect();
      canvas.dispatchEvent(new MouseEvent('contextmenu', { clientX: rect.left + x, clientY: rect.top + y, button: 2, bubbles: true, cancelable: true }));
    },
    [HEADER_WIDTH + (col + 0.5) * COLUMN_WIDTH, HEADER_HEIGHT + (row + 0.5) * ROW_HEIGHT],
  );
  await expect(page.locator('.bym-menu')).toBeVisible();
}

const item = (page: Page, name: string) => page.locator('.bym-menu').getByRole('menuitem', { name, exact: true });
const live = (page: Page) => page.locator('.bym-live');

test.describe('pivot tables', () => {
  test('the menu of a cell of a table refreshes it, and undo takes the refresh back', async ({ page }) => {
    await mount(page, await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
    const before = await body(page);
    expect(before.slice(0, 2)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['Grand Total', 0],
    ]);
    // A cell outside the table has no such command.
    await openCellMenu(page, 3, 10);
    await expect(item(page, 'Refresh pivot table')).toHaveCount(0);
    await page.keyboard.press('Escape');
    await openCellMenu(page, 0, 3);
    await item(page, 'Refresh pivot table').click();
    await expect(live(page)).toHaveText('Pivot table refreshed');
    expect(await body(page)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['East', 15],
      ['West', 27],
      ['Grand Total', 42],
      [null, null],
      [null, null],
    ]);
    expect(await view(page, 'view.pivots()[0].range')).toEqual({ top: 2, left: 0, bottom: 5, right: 1 });
    await view(page, 'view.undo()');
    expect(await body(page)).toEqual(before);
    await view(page, 'view.redo()');
    expect((await body(page))[3]).toEqual(['Grand Total', 42]);
  });

  test('a refresh that cannot be made says why, and changes nothing', async ({ page }) => {
    await mount(page, await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], around: [[4, 0, 'mine']] }));
    const before = await body(page);
    await openCellMenu(page, 1, 2);
    await item(page, 'Refresh pivot table').click();
    await expect(live(page)).toHaveText('Cannot refresh the pivot table "PivotTable1": it would grow over A5, which is not empty');
    expect(await body(page)).toEqual(before);
  });

  test('in Turkish the commands and the reasons are Turkish', async ({ page }) => {
    await mount(page, await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], around: [[4, 0, 'mine']] }), { locale: 'tr' });
    await openCellMenu(page, 1, 2);
    await item(page, 'Pivot tabloyu yenile').click();
    await expect(live(page)).toHaveText('"PivotTable1" pivot tablosu yenilenemiyor: boş olmayan A5 hücresinin üzerine taşardı');
  });

  test('a table whose refresh is known to be refused shows the command disabled', async ({ page }) => {
    await mount(page, await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], tableTail: '<filters count="1"><filter fld="0" type="captionEqual" evalOrder="-1" id="1"/></filters>' }));
    await openCellMenu(page, 0, 2);
    await expect(item(page, 'Refresh pivot table')).toHaveAttribute('aria-disabled', 'true');
  });

  test('the menu looks through the workbook first: a table a pivot chart is made from shows the command disabled', async ({ page }) => {
    const zip = ZipArchive.read(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }));
    zip.set(
      'xl/charts/chart1.xml',
      new TextEncoder().encode('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:pivotSource><c:name>[x.xlsx]Pivot!PivotTable1</c:name></c:pivotSource><c:chart/></c:chartSpace>'),
    );
    await mount(page, await zip.write());
    // Nothing is looked through when the file is opened.
    expect(await view(page, 'view.pivots()[0].refusal === undefined')).toBe(true);
    await openCellMenu(page, 0, 2);
    await expect(item(page, 'Refresh pivot table')).toHaveAttribute('aria-disabled', 'true');
    expect(await view(page, 'view.pivots()[0].refusal')).toBe('a pivot chart is made from a table of its cache');
  });

  test('a read-only view offers no refresh', async ({ page }) => {
    await mount(page, await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] }), { readOnly: true });
    await openCellMenu(page, 0, 3);
    await expect(item(page, 'Refresh pivot table')).toHaveCount(0);
  });

  test('the demo page makes a workbook with two pivot tables, and both are refreshed', async ({ page }) => {
    await page.goto('/demo/index.html');
    await page.locator('#sample-pivot').click();
    await expect(page.locator('#grid .bym-tabs button')).toHaveText(['Veri', 'Pivot']);
    await page.waitForFunction(() => !!(window as any).__view);
    const result = await page.evaluate(() => {
      const view = (window as any).__view;
      const all = view.refreshAllPivots();
      return { all, ranges: view.pivots().map((table: any) => (window as any).BymOffice.rangeName(table.range)) };
    });
    expect(result).toEqual({ all: { refreshed: [0, 1], refused: [] }, ranges: ['A3:D8', 'H3:K13'] });
  });

  test('mount() refreshes every table with refreshPivots()', async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
    const bytes = Array.from(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 2, col: 3 } }));
    const result = await page.evaluate(async (input) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), locale: 'en-US' });
      const out = mounted.refreshPivots();
      const sheet = mounted.workbook.sheets[1];
      return { out, grand: [sheet.rows.get(5).cells[1].value, sheet.rows.get(5).cells[4].value] };
    }, bytes);
    expect(result).toEqual({ out: { refreshed: 2, refused: [] }, grand: [42, 42] });
  });
});
