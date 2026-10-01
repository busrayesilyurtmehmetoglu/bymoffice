import { minimalCfb, sampleWorkbook } from '../test/fixtures';
import { expect, test } from './helpers';

// The embedding entry point as a host page would use it.
test.describe('BymOffice.mount', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/e2e/pages/mount.html');
  });

  test('shows a workbook from bytes and reports selection changes', async ({ page }) => {
    const bytes = Array.from(await sampleWorkbook());
    const result = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const events: string[] = [];
      const mounted = await api.mount(document.getElementById('editor'), {
        file: new Uint8Array(input),
        onSelectionChange: (selection: any) => events.push(api.rangeName(selection.range)),
      });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 0, left: 0, bottom: 1, right: 1 });
      return { format: mounted.file.format, sheets: mounted.workbook.sheets.length, events };
    }, bytes);
    expect(result).toEqual({ format: 'xlsx', sheets: 3, events: ['A1:B2'] });
    await expect(page.locator('#editor .bym-tabs button')).toHaveText(['Veri', 'Grafik']);

    await page.evaluate(() => (window as any).__mounted.destroy());
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('the locale option and setLocale change what cells show', async ({ page }) => {
    const bytes = Array.from(await sampleWorkbook());
    await page.evaluate(async (input) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), locale: 'tr' });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 1, left: 1, bottom: 1, right: 1 });
    }, bytes);
    const live = page.locator('#editor [aria-live]');
    await expect(live).toHaveText('B2 12,50 ₺');

    await page.evaluate(() => (window as any).__mounted.view.setLocale('en'));
    await expect(live).toHaveText('B2 12.50 ₺');

    // An incomplete locale object falls back to en-US for what it leaves out.
    await page.evaluate(() => (window as any).__mounted.view.setLocale({ decimal: ';' }));
    await expect(live).toHaveText('B2 12;50 ₺');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('the view speaks the language of its locale', async ({ page, isMobile }) => {
    const bytes = Array.from(await sampleWorkbook());
    await page.evaluate(async (input) => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: new Uint8Array(input), locale: 'tr' });
      (window as any).__mounted = mounted;
      mounted.view.select({ top: 0, left: 0, bottom: 1, right: 1 });
    }, bytes);
    await expect(page.locator('#editor .bym-view')).toHaveAttribute('aria-label', 'Hesap tablosu');
    await expect(page.locator('#editor .bym-input')).toHaveAttribute('aria-label', 'Hücre düzenleyici');
    await expect(page.locator('#editor .bym-scroller')).toHaveAttribute('aria-label', 'Sayfa Veri');
    await expect(page.locator('#editor [aria-live]')).toHaveText('A1 Ürün, A1:B2 seçili');

    if (!isMobile) {
      // Header of row 5: the header strip is 72 px wide; rows 1-2 are 32 and 20 px, row 3 is hidden.
      await page.mouse.click(30, 22 + 32 + 20 + 20 + 10, { button: 'right' });
      await expect(page.locator('#editor .bym-menu button')).toHaveText(['Üste satır ekle', 'Alta satır ekle', 'Satırı sil']);
      await page.keyboard.press('Escape');
    }

    await page.evaluate(() => (window as any).__mounted.view.setLocale('en'));
    await expect(page.locator('#editor .bym-view')).toHaveAttribute('aria-label', 'Spreadsheet');
    await expect(page.locator('#editor .bym-scroller')).toHaveAttribute('aria-label', 'Sheet Veri');
    await page.evaluate(() => (window as any).__mounted.destroy());
  });

  test('loads the file from a URL', async ({ page, pageProblems }) => {
    const bytes = Buffer.from(await sampleWorkbook());
    await page.route('**/files/sample.xlsx', (route) => route.fulfill({ body: bytes, contentType: 'application/octet-stream' }));
    await page.route('**/files/missing.xlsx', (route) => route.fulfill({ status: 404, body: 'nope' }));

    const names = await page.evaluate(async () => {
      const mounted = await (window as any).BymOffice.mount(document.getElementById('editor'), { file: '/files/sample.xlsx' });
      const sheets = mounted.workbook.sheets.map((sheet: any) => sheet.name);
      mounted.destroy();
      return sheets;
    });
    expect(names).toEqual(['Veri', 'Gizli', 'Grafik']);

    const message = await page.evaluate(() =>
      (window as any).BymOffice.mount(document.getElementById('editor'), { file: '/files/missing.xlsx' }).then(
        () => 'resolved',
        (error: Error) => error.message,
      ),
    );
    expect(message).toBe('Could not load the file: HTTP 404');
    // Browsers log the failed request themselves; that one message is expected here.
    expect(pageProblems.every((problem) => problem.includes('missing.xlsx'))).toBe(true);
    pageProblems.length = 0;
    await expect(page.locator('#editor')).toBeEmpty();
  });

  test('rejects files it cannot show and leaves the container empty', async ({ page }) => {
    const xls = Array.from(minimalCfb(new Uint8Array(9 * 512), new Uint8Array(10)));
    const outcome = await page.evaluate(async (input) => {
      const api = (window as any).BymOffice;
      const host = document.getElementById('editor');
      const names: string[] = [];
      for (const file of [new Uint8Array(input), new TextEncoder().encode('plain text')]) {
        await api.mount(host, { file }).catch((error: Error) => names.push(error.name));
      }
      return names;
    }, xls);
    expect(outcome).toEqual(['UnsupportedFileError', 'UnsupportedFileError']);
    await expect(page.locator('#editor')).toBeEmpty();
  });
});
