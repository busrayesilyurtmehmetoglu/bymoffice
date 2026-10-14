// Editing commands through the editors as a user reaches them: frozen panes from the menu of
// rows and columns, merged cells from the menu beside the merge button (with the question
// before values are discarded), and the dialog of tab stops of a document. Every file is made
// by code.
import type { Locator, Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { p } from '../test/doc-layout-fixtures';
import { docx } from './fixtures/docx';
import { expect, test } from './helpers';

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;

/** A1..C3 hold numbers; nothing is frozen or merged. */
const book = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Data',
        xml: worksheetXml(
          '<sheetData>' +
            `<row r="1">${num('A1', 1)}${num('B1', 2)}${num('C1', 3)}</row>` +
            `<row r="2">${num('A2', 4)}${num('B2', 5)}${num('C2', 6)}</row>` +
            `<row r="3">${num('A3', 7)}${num('B3', 8)}${num('C3', 9)}</row>` +
            '</sheetData>',
        ),
      },
    ],
  });

async function mountSheet(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(
    async ({ base64 }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountEditor(data, { locale: 'en-US' }, 'host');
    },
    { base64: Buffer.from(bytes).toString('base64') },
  );
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
}

const command = (page: Page, name: string) => page.locator(`#host [data-command="${name}"]`);
const popover = (page: Page) => page.locator('#host .bym-popover');
const view = <T>(page: Page, expression: string): Promise<T> => page.evaluate((code) => new Function('view', `return (${code});`)((window as any).__editor.view), expression);
const select = (page: Page, ref: string) => page.evaluate((r) => (window as any).__editor.view.select((window as any).BymOffice.parseRange(r)), ref);

function presser(hasTouch: boolean | undefined): (target: Locator) => Promise<void> {
  return (target) => (hasTouch ? target.tap() : target.click());
}

async function choose(page: Page, press: (target: Locator) => Promise<void>, button: string, label: string): Promise<void> {
  await press(command(page, button));
  await press(popover(page).getByRole('menuitem', { name: label, exact: true }));
  await expect(popover(page)).toHaveCount(0);
}

test.describe('frozen panes', () => {
  test('the menu of rows and columns freezes at the active cell, the top row, the first column, and unfreezes; undo takes a step back', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book());
    await select(page, 'B3');
    await choose(page, press, 'rowsColumns', 'Freeze panes');
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 2, cols: 1 });
    await choose(page, press, 'rowsColumns', 'Freeze top row');
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 1, cols: 0 });
    await choose(page, press, 'rowsColumns', 'Freeze first column');
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 0, cols: 1 });
    // Unfreezing is offered once something is frozen.
    await press(command(page, 'rowsColumns'));
    await expect(popover(page).locator('.bym-pop-item').last()).toHaveText('Unfreeze panes');
    await press(popover(page).getByRole('menuitem', { name: 'Unfreeze panes', exact: true }));
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 0, cols: 0 });
    await view(page, 'view.undo()');
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 0, cols: 1 });
    // What is saved has the panes; the sheet asks for no calculation.
    const again = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const saved = await (window as any).__file.save();
      const workbook = await (await api.open(saved)).workbook();
      return [workbook.sheets[0].view.frozenRows, workbook.sheets[0].view.frozenCols];
    });
    expect(again).toEqual([0, 1]);
  });

  test('frozen rows stay at the top while the sheet scrolls', async ({ page }) => {
    await mountSheet(page, await book());
    await view(page, 'view.freezePanes(1, 0)');
    const before = await view<{ scrollY: number }>(page, 'view.getViewState()');
    await page.evaluate(() => {
      const scroller = document.querySelector('#host .bym-scroller') as HTMLElement | null;
      if (scroller) scroller.scrollTop = 400;
    });
    await expect.poll(() => view<{ scrollY: number }>(page, 'view.getViewState()').then((state) => state.scrollY)).toBeGreaterThan(before.scrollY);
    expect(await view(page, 'view.frozenPanes')).toEqual({ rows: 1, cols: 0 });
  });
});

test.describe('merged cells', () => {
  test('"Merge and centre" asks before values are discarded: cancel leaves everything, merge merges and centres', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book());
    await select(page, 'A1:B2');
    await choose(page, press, 'mergeOptions', 'Merge and centre');
    const question = page.locator('#host [data-bym-confirm="merge-loss"]');
    await expect(question).toBeVisible();
    await expect(question).toContainText('3 other cells');
    await press(question.locator('[data-bym-action="cancel"]'));
    await expect(question).toHaveCount(0);
    expect(await view(page, 'view.workbook.sheets[0].merges.length')).toBe(0);

    await choose(page, press, 'mergeOptions', 'Merge and centre');
    // The answer that changes something answers only once the question has been there a moment.
    await expect(question.locator('[data-bym-action="overwrite"]:not([aria-disabled])')).toBeVisible();
    await press(question.locator('[data-bym-action="overwrite"]'));
    await expect(question).toHaveCount(0);
    expect(await view(page, 'view.workbook.sheets[0].merges')).toEqual([{ top: 0, left: 0, bottom: 1, right: 1 }]);
    expect(await view(page, 'view.selectionFormat.alignment.horizontal')).toBe('center');
    // The other cells are emptied, and keep the centring (an empty cell with a format of its own).
    expect(await view(page, '[view.workbook.sheets[0].rows.get(0).cells[1].value, view.workbook.sheets[0].rows.get(1).cells[0].value, view.workbook.sheets[0].rows.get(1).cells[1].value]')).toEqual([null, null, null]);
    // One step: undo brings the values and the format back.
    await view(page, 'view.undo()');
    expect(await view(page, 'view.workbook.sheets[0].merges.length')).toBe(0);
    expect(await view(page, 'view.workbook.sheets[0].rows.get(1).cells[1].value')).toBe(5);
  });

  test('cells without values are merged without a question; "Unmerge cells" splits them', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book());
    await select(page, 'E5:F6');
    await choose(page, press, 'mergeOptions', 'Merge cells');
    await expect(page.locator('#host [data-bym-confirm]')).toHaveCount(0);
    expect(await view(page, 'view.workbook.sheets[0].merges')).toEqual([{ top: 4, left: 4, bottom: 5, right: 5 }]);
    await choose(page, press, 'mergeOptions', 'Unmerge cells');
    expect(await view(page, 'view.workbook.sheets[0].merges.length')).toBe(0);
  });

  test('merge() of the view discards without asking and says what it discarded', async ({ page }) => {
    await mountSheet(page, await book());
    await select(page, 'B2:C3');
    const report = await view<{ count: number; losses: { row: number; col: number; text: string }[] }>(page, 'view.merge()');
    expect(report.count).toBe(3);
    expect(report.losses.map((loss) => loss.text)).toEqual(['6', '8', '9']);
  });
});

test.describe('tab stops of a document', () => {
  async function mountDoc(page: Page, bytes: Uint8Array): Promise<void> {
    await page.goto('/e2e/pages/doc-edit.html');
    await expect(page.locator('#state')).toHaveText('ready');
    await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { toolbar: true, locale: 'tr' }] as [number[], unknown]);
    await page.waitForFunction(() => (window as any).__view && (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
  }

  test('the dialog sets a stop and the default distance, in centimetres; cancel changes nothing', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountDoc(page, await docx(p('bir') + p('iki')));
    const button = page.locator('#host .bym-doc-toolbar [data-command="tabStops"]');
    await press(button);
    const dialog = page.locator('#host .bym-tabs-dialog');
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('.bym-tabs-list')).toContainText('Sekme durağı yok');
    await dialog.locator('.bym-tabs-position').fill('2,54');
    await dialog.locator('.bym-tabs-align').selectOption('right');
    await dialog.locator('.bym-tabs-leader').selectOption('dot');
    await press(dialog.locator('[data-bym-action="set"]'));
    await expect(dialog.locator('.bym-tabs-stop')).toHaveCount(1);
    await dialog.locator('.bym-tabs-default').fill('1');
    await press(dialog.locator('[data-bym-action="ok"]'));
    await expect(dialog).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.tabStops)).toEqual([{ pos: 1440, align: 'right', leader: 'dot' }]);
    expect(await page.evaluate(() => (window as any).__view.defaultTabStop)).toBe(567);

    await press(button);
    await expect(dialog.locator('.bym-tabs-stop')).toHaveCount(1);
    await press(dialog.locator('[data-bym-action="clear-all"]'));
    await expect(dialog.locator('.bym-tabs-stop')).toHaveCount(0);
    await press(dialog.locator('[data-bym-action="cancel"]'));
    await expect(dialog).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.tabStops)).toEqual([{ pos: 1440, align: 'right', leader: 'dot' }]);
  });

  test('a position that is no number is said, and nothing is set', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountDoc(page, await docx(p('bir')));
    await press(page.locator('#host .bym-doc-toolbar [data-command="tabStops"]'));
    const dialog = page.locator('#host .bym-tabs-dialog');
    await dialog.locator('.bym-tabs-position').fill('abc');
    await press(dialog.locator('[data-bym-action="set"]'));
    await expect(dialog.locator('.bym-tabs-message')).not.toHaveText('');
    await page.keyboard.press('Escape');
    await expect(dialog).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.tabStops)).toEqual([]);
  });
});
