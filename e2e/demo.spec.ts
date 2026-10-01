import { readFile } from 'node:fs/promises';
import { XLSX_PARTS, XLS_STREAMS, readPartsInNode, textFile, xls, xlsx } from './fixtures/files';
import { sampleWorkbook } from '../test/fixtures';
import { expect, expectResponsive, test } from './helpers';

const upload = (name: string, bytes: Uint8Array) => ({
  name,
  mimeType: 'application/octet-stream',
  buffer: Buffer.from(bytes),
});

test.describe('demo page', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/demo/index.html');
    await expect(page.locator('#save')).toBeDisabled();
  });

  test('lists the parts of an xlsx and enables saving', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('sample.xlsx', await xlsx()));

    await expect(page.locator('#status')).toHaveText('Biçim: xlsx (ooxml)');
    await expect(page.locator('#parts tr th')).toHaveText(['Parça', 'İçerik türü', 'Boyut (bayt)']);
    await expect(page.locator('#parts tr td:first-child')).toHaveText(XLSX_PARTS);
    await expect(page.locator('#save')).toBeEnabled();
  });

  test('"save" downloads bytes that reopen to the same parts', async ({ page }) => {
    const input = await xlsx();
    await page.locator('#file').setInputFiles(upload('sample.xlsx', input));
    await expect(page.locator('#save')).toBeEnabled();

    const downloadPromise = page.waitForEvent('download', { timeout: 10_000 });
    // tap on touch profiles, click elsewhere: the gesture a real user makes
    await page.locator('#save').click();
    const download = await downloadPromise;

    expect(download.suggestedFilename()).toBe('sample.xlsx');
    expect(await download.failure()).toBeNull();
    const saved = new Uint8Array(await readFile(await download.path()));
    expect(saved.length).toBeGreaterThan(0);
    expect(await readPartsInNode(saved)).toEqual(await readPartsInNode(input));
    expect(Array.from(saved)).toEqual(Array.from(input));
  });

  test('lists the streams of an xls (compound file)', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('sample.xls', xls()));

    await expect(page.locator('#status')).toHaveText('Biçim: xls (cfb)');
    await expect(page.locator('#parts tr th')).toHaveText(['Akış', 'Boyut (bayt)']);
    await expect(page.locator('#parts tr td:first-child')).toHaveText(XLS_STREAMS.map((stream) => stream.path));
    await expect(page.locator('#parts tr td:nth-child(2)')).toHaveText(XLS_STREAMS.map((stream) => String(stream.size)));
    await expect(page.locator('#save')).toBeEnabled();
  });

  test('shows an error for a plain text file and keeps working', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('notes.txt', textFile()));

    await expect(page.locator('#status')).toHaveText(/^Açılamadı: .+/);
    await expect(page.locator('#parts tr')).toHaveCount(0);
    await expect(page.locator('#save')).toBeDisabled();
    await expectResponsive(page);

    // The page is still usable: a valid file opens afterwards.
    await page.locator('#file').setInputFiles(upload('sample.xlsx', await xlsx()));
    await expect(page.locator('#status')).toHaveText('Biçim: xlsx (ooxml)');
    await expect(page.locator('#save')).toBeEnabled();
  });
  test('built-in sample opens, round-trips and shows part content on click', async ({ page }) => {
    await page.locator('#sample').click();

    await expect(page.locator('#status')).toHaveText('Biçim: xlsx (ooxml)');
    await expect(page.locator('#files tr')).toHaveCount(1);
    await expect(page.locator('#files tr .badge')).toHaveText('4 parça aynı');
    await expect(page.locator('#files tr .badge')).toHaveClass(/ok/);

    await page.locator('#parts tr', { hasText: '/xl/worksheets/sheet1.xml' }).click();
    await expect(page.locator('#viewer-title')).toContainText('/xl/worksheets/sheet1.xml');
    await expect(page.locator('#viewer')).toContainText('Satır 200 — ğüşiöç');
  });

  test('old-device mode uses the engine codec and still round-trips', async ({ page }) => {
    await page.locator('#legacy').check();
    expect(await page.evaluate(() => typeof (globalThis as any).DecompressionStream)).toBe('undefined');
    await page.locator('#sample').click();
    await expect(page.locator('#files tr .badge')).toHaveText('4 parça aynı');

    // A file written by another codec is read by ours as well.
    await page.locator('#file').setInputFiles(upload('sample.xlsx', await xlsx()));
    await expect(page.locator('#files tr')).toHaveCount(2);
    await expect(page.locator('#files tr .badge').nth(1)).toHaveText(`${XLSX_PARTS.length} parça aynı`);

    await page.locator('#legacy').uncheck();
    expect(await page.evaluate(() => typeof (globalThis as any).DecompressionStream)).toBe('function');
  });

  test('several files at once are listed with a summary; clicking a row selects it', async ({ page }) => {
    await page
      .locator('#file')
      .setInputFiles([upload('a.xlsx', await xlsx()), upload('b.xls', xls()), upload('c.txt', textFile())]);

    await expect(page.locator('#files tr')).toHaveCount(3);
    await expect(page.locator('#summary')).toHaveText('3 dosya: 2 açıldı (2 kayıpsız), 1 desteklenmiyor, 0 hata');
    await expect(page.locator('#status')).toHaveText(/^Açılamadı: /);

    await page.locator('#files tr', { hasText: 'b.xls' }).click();
    await expect(page.locator('#status')).toHaveText('Biçim: xls (cfb)');
    await page.locator('#parts tr', { hasText: 'Small' }).click();
    await expect(page.locator('#viewer-title')).toHaveText('İçerik: Small (100 bayt)');
    await expect(page.locator('#viewer')).toHaveText(/^00000000  ([0-9a-f]{2} ){15}[0-9a-f]{2}  /);

    await page.locator('#clear').click();
    await expect(page.locator('#files tr')).toHaveCount(0);
    await expect(page.locator('#status')).toHaveText('—');
    await expect(page.locator('#save')).toBeDisabled();
  });
  test('xlsx opens in the grid view with sheet tabs and a selection read-out', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('model.xlsx', await sampleWorkbook()));

    await expect(page.locator('#preview-section')).toBeVisible();
    // Hidden sheets get no tab; the workbook's active sheet is hidden, so the first visible one shows.
    await expect(page.locator('#grid .bym-tabs button')).toHaveText(['Veri', 'Grafik']);
    await expect(page.locator('#grid .bym-tabs button[aria-selected="true"]')).toHaveText('Veri');
    await expect(page.locator('#preview-note')).toContainText('Veri: worksheet, 4 satır × 3 sütun');
    await expect(page.locator('#preview-note')).toContainText('dondurulmuş 1 satır / 1 sütun');
    // The file's active cell is B3.
    await expect(page.locator('#cell-info')).toHaveText('Seçim: B3 · B3 = 7,00 ₺ (ham değer: 7)');

    // Programmatic selection goes through the same read-out as a click would.
    await page.evaluate(() => (window as any).__view.select({ top: 3, left: 1, bottom: 3, right: 1 }));
    await expect(page.locator('#cell-info')).toHaveText('Seçim: B4 · B4 = 19,50 ₺ (ham değer: 19.5) · formül: =SUM(B2:B3)');

    // The number format follows the chosen language.
    await page.locator('#locale').selectOption('en');
    await expect(page.locator('#cell-info')).toHaveText('Seçim: B4 · B4 = 19.50 ₺ (ham değer: 19.5) · formül: =SUM(B2:B3)');
    await page.locator('#locale').selectOption('tr');

    const canvas = page.locator('#grid canvas');
    const box = await canvas.boundingBox();
    expect(box && box.width > 200 && box.height > 200).toBe(true);
    const painted = await canvas.evaluate((element) => {
      const c = element as HTMLCanvasElement;
      const data = c.getContext('2d')!.getImageData(0, 0, c.width, c.height).data;
      const colours = new Set<number>();
      for (let i = 0; i < data.length; i += 4 * 97) colours.add((data[i] << 16) | (data[i + 1] << 8) | data[i + 2]);
      return colours.size;
    });
    expect(painted).toBeGreaterThan(3);

    await page.locator('#grid .bym-tabs button', { hasText: 'Grafik' }).click();
    await expect(page.locator('#preview-note')).toContainText('Grafik: chartsheet');
  });

  test('switching files replaces the grid view instead of stacking them', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('a.xlsx', await sampleWorkbook()));
    await expect(page.locator('#grid .bym-view')).toHaveCount(1);
    await page.locator('#file').setInputFiles(upload('b.xlsx', await xlsx()));
    await expect(page.locator('#grid .bym-tabs button')).toHaveText(['Sayfa1']);
    await expect(page.locator('#grid .bym-view')).toHaveCount(1);
    await page.locator('#clear').click();
    await expect(page.locator('#grid .bym-view')).toHaveCount(0);
  });

  test('no cell preview for files without a spreadsheet model', async ({ page }) => {
    await page.locator('#file').setInputFiles(upload('sample.xls', xls()));
    await expect(page.locator('#status')).toHaveText('Biçim: xls (cfb)');
    await expect(page.locator('#preview-section')).toBeHidden();
  });
});
