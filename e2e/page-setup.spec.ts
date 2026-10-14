// The dialog of the page setup as a user reaches it from the toolbars: a workbook (paper,
// orientation, margins, scaling, print area and titles, page breaks at the active cell) and a
// document (paper, orientation and margins of a section or of the whole document, section
// breaks). Only what is changed in the dialog changes; the change is one undo step and is in the
// saved file. Every file is made by code.
import type { Locator, Page } from '@playwright/test';
import { buildXlsx, worksheetXml } from '../test/fixtures';
import { p } from '../test/doc-layout-fixtures';
import { docx } from './fixtures/docx';
import { expect, test } from './helpers';

const num = (ref: string, value: number) => `<c r="${ref}"><v>${value}</v></c>`;

/** A1..C3 hold numbers; the page is A4 upright with Excel's normal margins. */
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
            '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/>',
        ),
      },
    ],
  });

async function mountSheet(page: Page, bytes: Uint8Array, locale = 'en-US'): Promise<void> {
  await page.goto('/e2e/pages/editor.html');
  await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
  await page.evaluate(
    async ({ base64, tag }) => {
      const text = atob(base64);
      const data = new Uint8Array(text.length);
      for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
      await (window as any).__mountEditor(data, { locale: tag }, 'host');
    },
    { base64: Buffer.from(bytes).toString('base64'), tag: locale },
  );
  await expect(page.locator('#host .bym-canvas')).toBeVisible();
}

const sheetView = <T>(page: Page, expression: string): Promise<T> => page.evaluate((code) => new Function('view', `return (${code});`)((window as any).__editor.view), expression);

function presser(hasTouch: boolean | undefined): (target: Locator) => Promise<void> {
  return (target) => (hasTouch ? target.tap() : target.click());
}

const dialogOf = (page: Page) => page.locator('#host .bym-page-setup-dialog');

test.describe('page setup of a sheet', () => {
  test('the toolbar opens the dialog; landscape, a margin and grid lines are set as one step, and are in the saved file', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book());
    await press(page.locator('#host .bym-toolbar [data-command="pageSetup"]'));
    const dialog = dialogOf(page);
    await expect(dialog).toBeVisible();
    await expect(dialog.locator('.bym-page-setup-left')).toHaveValue('0.7');
    await dialog.locator('.bym-page-setup-orientation').selectOption('landscape');
    await dialog.locator('.bym-page-setup-left').fill('1');
    await dialog.locator('.bym-page-setup-gridLines').check();
    await dialog.locator('.bym-page-setup-printArea').fill('A1:B2');
    await dialog.locator('.bym-page-setup-titleRows').fill('$1:$1');
    await press(dialog.locator('[data-bym-action="ok"]'));
    await expect(dialog).toHaveCount(0);
    const state = await sheetView<any>(page, 'view.pageSetup()');
    expect([state.orientation, state.margins.left, state.margins.right, state.gridLines, state.printArea, state.titleRows]).toEqual(['landscape', 1, 0.7, true, [{ top: 0, left: 0, bottom: 1, right: 1 }], { first: 0, last: 0 }]);
    const saved = await page.evaluate(async () => {
      const api = (window as any).BymOffice;
      const bytes = await (window as any).__file.save();
      const workbook = await (await api.open(bytes)).workbook();
      return { setup: workbook.sheets[0].pageSetup, names: workbook.definedNames.map((name: any) => [name.name, name.value]) };
    });
    expect(saved.setup.page.orientation).toBe('landscape');
    expect(saved.setup.margins.left).toBe(1);
    expect(saved.setup.print.gridLines).toBe(true);
    expect(saved.names).toEqual([
      ['_xlnm.Print_Area', 'Data!$A$1:$B$2'],
      ['_xlnm.Print_Titles', 'Data!$1:$1'],
    ]);
    // One step each: the setup and the names together.
    await sheetView(page, 'view.undo()');
    expect(await sheetView(page, 'view.pageSetup().orientation')).toBe('portrait');
    expect(await sheetView(page, 'view.pageSetup().printArea')).toEqual([]);
  });

  test('in centimetres, OK with nothing touched changes nothing; a wrong value is said and keeps the dialog open', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book(), 'tr');
    await press(page.locator('#host .bym-toolbar [data-command="pageSetup"]'));
    const dialog = dialogOf(page);
    await expect(dialog.locator('.bym-page-setup-left')).toHaveValue('1,78');
    await press(dialog.locator('[data-bym-action="ok"]'));
    await expect(dialog).toHaveCount(0);
    expect(await sheetView(page, 'view.canUndo')).toBe(false);
    expect(await sheetView(page, 'view.pageSetup().margins.left')).toBe(0.7);
    await press(page.locator('#host .bym-toolbar [data-command="pageSetup"]'));
    await dialog.locator('.bym-page-setup-top').fill('abc');
    await press(dialog.locator('[data-bym-action="ok"]'));
    await expect(dialog.locator('.bym-page-setup-message')).not.toHaveText('');
    await expect(dialog).toBeVisible();
    await page.keyboard.press('Escape');
    await expect(dialog).toHaveCount(0);
    expect(await sheetView(page, 'view.pageSetup().margins.top')).toBe(0.75);
  });

  test('page breaks at the active cell from the dialog, and from the view', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountSheet(page, await book());
    await page.evaluate(() => (window as any).__editor.view.select((window as any).BymOffice.parseRange('B3')));
    await press(page.locator('#host .bym-toolbar [data-command="pageSetup"]'));
    await press(dialogOf(page).locator('[data-bym-action="action-insert-break"]'));
    await expect(dialogOf(page)).toHaveCount(0);
    expect(await sheetView(page, '[view.pageSetup().rowBreaks, view.pageSetup().colBreaks]')).toEqual([[2], [1]]);
    expect(await sheetView(page, 'view.removePageBreak()')).toBe(true);
    expect(await sheetView(page, '[view.pageSetup().rowBreaks, view.pageSetup().colBreaks]')).toEqual([[], []]);
  });

  test('read-only: the button is dimmed', async ({ page }) => {
    await page.goto('/e2e/pages/editor.html');
    await page.waitForFunction(() => typeof (window as any).__mountEditor === 'function');
    const bytes = await book();
    await page.evaluate(
      async ({ base64 }) => {
        const text = atob(base64);
        const data = new Uint8Array(text.length);
        for (let i = 0; i < text.length; i++) data[i] = text.charCodeAt(i);
        await (window as any).__mountEditor(data, { locale: 'en-US', readOnly: true }, 'host');
      },
      { base64: Buffer.from(bytes).toString('base64') },
    );
    await expect(page.locator('#host .bym-toolbar [data-command="pageSetup"]')).toBeDisabled();
    expect(await sheetView(page, 'view.openPageSetup()')).toBe(false);
  });
});

test.describe('page setup of a document', () => {
  async function mountDoc(page: Page, bytes: Uint8Array): Promise<void> {
    await page.goto('/e2e/pages/doc-edit.html');
    await expect(page.locator('#state')).toHaveText('ready');
    await page.evaluate(([input, opts]) => (window as any).__mountEdit(input, opts), [Array.from(bytes), { toolbar: true, locale: 'tr' }] as [number[], unknown]);
    await page.waitForFunction(() => (window as any).__view && (window as any).__view.layoutDone === true, undefined, { timeout: 120_000 });
  }

  test('the dialog turns the page to landscape A4 and the pages follow; undo takes it back', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountDoc(page, await docx(p('bir') + p('iki')));
    const before = await page.evaluate(() => (window as any).__view.sectionSetup);
    await press(page.locator('#host .bym-doc-toolbar [data-command="pageSetup"]'));
    const dialog = dialogOf(page);
    await expect(dialog).toBeVisible();
    await dialog.locator('.bym-page-setup-paper').selectOption('a4');
    await dialog.locator('.bym-page-setup-orientation').selectOption('landscape');
    await press(dialog.locator('[data-bym-action="ok"]'));
    await expect(dialog).toHaveCount(0);
    const after = await page.evaluate(() => (window as any).__view.sectionSetup);
    expect([after.width, after.height, after.orientation, after.paper]).toEqual([16838, 11906, 'landscape', 'a4']);
    // The margins were not touched: they are what they were.
    expect(after.margins).toEqual(before.margins);
    await page.waitForFunction(() => (window as any).__view.layoutDone === true);
    await page.evaluate(() => (window as any).__view.undo());
    expect(await page.evaluate(() => (window as any).__view.sectionSetup.width)).toBe(before.width);
  });

  test('a section break from the dialog makes a second section', async ({ page, hasTouch }) => {
    const press = presser(hasTouch);
    await mountDoc(page, await docx(p('bir') + p('iki')));
    await page.evaluate(() => (window as any).__view.focus());
    await press(page.locator('#host .bym-doc-toolbar [data-command="pageSetup"]'));
    await press(dialogOf(page).locator('[data-bym-action="action-section-continuous"]'));
    await expect(dialogOf(page)).toHaveCount(0);
    expect(await page.evaluate(() => (window as any).__view.sectionSetup.sections)).toBe(2);
  });
});
