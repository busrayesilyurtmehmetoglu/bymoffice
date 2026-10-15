// Making pivot tables in the workbook view and changing their fields: the menu of a cell outside
// any table opens the dialog that makes one, the panel of its fields opens beside the cells (along
// the bottom of a narrow view), its boxes and buttons change the table as one undo step each, a
// change that is not built here is said in the panel, and the words follow the language. The
// workbooks are made by test/pivot-fixtures.ts: the data stands from A1 of "Sheet1".
//
// Positions are written out by hand from the default sizes of a sheet (header 72 x 22, columns
// 64 wide, rows 20 high), within columns A to D so that a phone shows them.
import type { Page } from '@playwright/test';
import { bookWith, save, type Datum } from '../test/pivot-fixtures';
import { expect, test } from './helpers';

const HEADER_WIDTH = 72;
const HEADER_HEIGHT = 22;
const COLUMN_WIDTH = 64;
const ROW_HEIGHT = 20;

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
  ['East', 'B', 5],
  ['West', 'A', 7],
];

async function bytes(data: Datum[][] = DATA): Promise<Uint8Array> {
  return save(await bookWith(data));
}

async function mount(page: Page, file: Uint8Array, options: Record<string, unknown> = {}): Promise<void> {
  await page.goto('/e2e/pages/view.html');
  await page.waitForFunction(() => typeof (window as any).__mountView === 'function');
  await page.evaluate(
    async ({ base64, options }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountView(data, options);
    },
    { base64: Buffer.from(file).toString('base64'), options: { locale: 'en-US', ...options } },
  );
  await expect(page.locator('.bym-canvas')).toBeVisible();
}

const view = (page: Page, code: string) => page.evaluate((code) => new Function('view', `return ${code}`)((window as any).__view), code);

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
const panel = (page: Page) => page.locator('[data-bym-panel="pivot"]');
const dialog = (page: Page) => page.locator('[data-bym-dialog="pivot"]');

/** The values of a block of a sheet. */
const cells = (page: Page, sheet: number, top: number, left: number, bottom: number, right: number) =>
  view(
    page,
    `(() => { const s = view.workbook.sheets[${sheet}]; const out = []; for (let r = ${top}; r <= ${bottom}; r++) { const row = s.rows.get(r); const line = []; for (let c = ${left}; c <= ${right}; c++) line.push(row && row.cells[c] ? row.cells[c].value : null); out.push(line); } return out; })()`,
  );

async function makeOne(page: Page): Promise<void> {
  await openCellMenu(page, 1, 2);
  await item(page, 'Insert pivot table…').click();
  await expect(dialog(page)).toBeVisible();
  await expect(dialog(page).locator('input').first()).toHaveValue('A1:C5');
  await dialog(page).locator('[data-bym-action="ok"]').click();
  await expect(panel(page)).toBeVisible();
}

test.describe('making pivot tables and changing their fields', () => {
  test('the menu of a cell makes a table on a new sheet from the block around it; the panel of its fields opens', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    await expect(dialog(page)).toHaveCount(0);
    await expect(live(page)).toHaveText('Pivot table inserted');
    await expect(page.locator('.bym-tab')).toHaveText(['Sheet1', 'Sheet2']);
    expect(await view(page, 'view.activeSheet')).toBe(1);
    expect(await cells(page, 1, 2, 0, 5, 1)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['East', 15],
      ['West', 27],
      ['Grand Total', 42],
    ]);
    // The panel names the fields as text, and the boxes hold what the table has.
    await expect(panel(page).locator('.bym-pivot-field .bym-pivot-name')).toHaveText(['Region', 'Product', 'Sales']);
    await expect(panel(page).locator('[data-bym-box="rows"] .bym-pivot-entry .bym-pivot-name')).toHaveText(['Region']);
    await expect(panel(page).locator('[data-bym-box="values"] .bym-pivot-entry .bym-pivot-name')).toHaveText(['Sum of Sales']);
    // Undone, the table and its sheet are gone, and so is the panel.
    await view(page, 'view.undo()');
    await expect(panel(page)).toHaveCount(0);
    await expect(page.locator('.bym-tab')).toHaveText(['Sheet1']);
  });

  test('the panel changes the table: a field ticked, moved to the columns, a value summarized otherwise; each is one undo step', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    // Product holds text: ticked, it goes into the rows.
    await panel(page).locator('[data-bym-key="use:1"]').click();
    expect(await view(page, 'view.pivotFields(0).rows')).toEqual([0, 1]);
    await expect(live(page)).toHaveText('Pivot table changed');
    // Moved to the columns with the list of its entry.
    await panel(page).getByLabel('Where Product stands').selectOption('cols');
    expect(await view(page, '[view.pivotFields(0).rows, view.pivotFields(0).cols]')).toEqual([[0], [1]]);
    expect(await cells(page, 1, 2, 0, 6, 3)).toEqual([
      ['Sum of Sales', 'Column Labels', null, null],
      ['Row Labels', 'A', 'B', 'Grand Total'],
      ['East', 10, 5, 15],
      ['West', 7, 20, 27],
      ['Grand Total', 17, 25, 42],
    ]);
    await panel(page).getByLabel('Summarize Sum of Sales by').selectOption('count');
    expect(await view(page, 'view.pivotFields(0).values')).toEqual([{ field: 2, subtotal: 'count', name: 'Count of Sales' }]);
    // The focus stays on the control that was used.
    await expect(panel(page).getByLabel('Summarize Count of Sales by')).toBeVisible();
    await view(page, 'view.undo()');
    await expect(panel(page).locator('[data-bym-box="values"] .bym-pivot-name')).toHaveText(['Sum of Sales']);
    await view(page, 'view.undo()');
    expect(await view(page, '[view.pivotFields(0).rows, view.pivotFields(0).cols]')).toEqual([[0, 1], []]);
  });

  test('a change that is not built here is said in the panel; nothing changes', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    await panel(page).getByLabel('Remove Region').click();
    await expect(panel(page).locator('[role="alert"]')).toHaveText('Cannot change the fields of the pivot table "PivotTable1": it needs a field in its rows here');
    expect(await view(page, 'view.pivotFields(0).rows')).toEqual([0]);
    expect(await view(page, 'view.history.canUndo')).toBe(true);
  });

  test('hidden items are left out; a filter stands above the table', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    await panel(page).locator('[data-bym-key="items:0"]').click();
    await panel(page).locator('[data-bym-key="item:0:0"]').click();
    expect(await cells(page, 1, 2, 0, 4, 1)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['West', 27],
      ['Grand Total', 27],
    ]);
    await panel(page).getByLabel('Where Region stands').selectOption('filters');
    await expect(panel(page).locator('[role="alert"]')).toHaveText(/needs a field in its rows here/);
    await panel(page).locator('[data-bym-key="use:1"]').click();
    await panel(page).getByLabel('Where Region stands').selectOption('filters');
    expect(await cells(page, 1, 0, 0, 0, 1)).toEqual([['Region', '(Multiple Items)']]);
  });

  test('Escape closes the panel and the cells have the keys again; the menu of a cell of the table opens it again', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    await panel(page).locator('[data-bym-key="use:1"]').focus();
    await page.keyboard.press('Escape');
    await expect(panel(page)).toHaveCount(0);
    await expect(page.locator('.bym-input')).toBeFocused();
    await openCellMenu(page, 0, 3);
    await item(page, 'Pivot table fields…').click();
    await expect(panel(page)).toBeVisible();
  });

  test('a cell that holds something is refused in the dialog, which stays open', async ({ page }) => {
    await mount(page, await bytes());
    await openCellMenu(page, 0, 0);
    await item(page, 'Insert pivot table…').click();
    await dialog(page).locator('select').selectOption('here');
    const cellField = dialog(page).locator('input').nth(1);
    await cellField.fill('B2');
    await dialog(page).locator('[data-bym-action="ok"]').click();
    await expect(dialog(page).locator('[role="alert"]')).toHaveText(/Cannot make the pivot table: /);
    await cellField.fill('E1');
    await dialog(page).locator('[data-bym-action="ok"]').click();
    await expect(dialog(page)).toHaveCount(0);
    expect(await cells(page, 0, 0, 4, 3, 5)).toEqual([
      ['Row Labels', 'Sum of Sales'],
      ['East', 15],
      ['West', 27],
      ['Grand Total', 42],
    ]);
  });

  test('in Turkish the commands, the panel and the reasons are Turkish', async ({ page }) => {
    await mount(page, await bytes(), { locale: 'tr' });
    await openCellMenu(page, 1, 2);
    await item(page, 'Pivot tablo ekle…').click();
    await dialog(page).locator('[data-bym-action="ok"]').click();
    await expect(panel(page).locator('.bym-pivot-box-title')).toHaveText(['Süzgeçler', 'Sütunlar', 'Satırlar', 'Değerler']);
    await panel(page).getByLabel('Region çıkar').click();
    await expect(panel(page).locator('[role="alert"]')).toHaveText('"PivotTable1" pivot tablosunun alanları değiştirilemiyor: burada satırlarında en az bir alan olmalı');
  });

  test('a narrow view has the panel along its bottom; a wide one at its right', async ({ page }) => {
    await mount(page, await bytes());
    await makeOne(page);
    const box = await panel(page).boundingBox();
    const host = await page.locator('#host').boundingBox();
    expect(box && host).toBeTruthy();
    const width = page.viewportSize()!.width;
    if (width <= 600) {
      expect(Math.abs(box!.width - host!.width)).toBeLessThan(2);
      expect(Math.abs(box!.y + box!.height - (host!.y + host!.height))).toBeLessThan(2);
    } else {
      expect(Math.abs(box!.x + box!.width - (host!.x + host!.width))).toBeLessThan(2);
      expect(box!.width).toBeLessThan(host!.width);
    }
  });

  test('a field is dragged from the list into the columns', async ({ page, hasTouch }) => {
    test.skip(!!hasTouch, 'dragging is for a mouse; fingers use the buttons of an entry');
    await mount(page, await bytes());
    await makeOne(page);
    await panel(page).locator('.bym-pivot-field[data-bym-field="1"]').dragTo(panel(page).locator('[data-bym-box="cols"]'));
    expect(await view(page, 'view.pivotFields(0).cols')).toEqual([1]);
  });

  test('a read-only view offers no table to make', async ({ page }) => {
    await mount(page, await bytes(), { readOnly: true });
    await openCellMenu(page, 1, 2);
    await expect(item(page, 'Insert pivot table…')).toHaveCount(0);
  });
});
