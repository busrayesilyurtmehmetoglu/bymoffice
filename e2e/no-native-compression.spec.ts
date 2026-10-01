import { ZipArchive } from '../src/zip';
import { XLSX_PARTS, xlsx } from './fixtures/files';
import { expect, gotoHarness, roundTripInPage, test, type Flavour } from './helpers';

// Older WebViews (iOS before 16.4, old Android System WebView) have neither
// DecompressionStream nor CompressionStream. The engine must bring its own codec.
for (const flavour of ['script', 'esm'] as Flavour[]) {
  test(`no native compression (${flavour}): xlsx with deflated entries opens, saves and reopens`, async ({ page }) => {
    await page.addInitScript(() => {
      delete (window as any).DecompressionStream;
      delete (window as any).CompressionStream;
    });
    const input = await xlsx(); // throws unless it contains deflated entries
    await gotoHarness(page, flavour);

    expect(
      await page.evaluate(() => [typeof (window as any).DecompressionStream, typeof (window as any).CompressionStream]),
    ).toEqual(['undefined', 'undefined']);

    const result = await roundTripInPage(page, input);
    expect(result.format).toBe('xlsx');
    expect(result.parts).toEqual(XLSX_PARTS);
    expect(result.reopenedParts).toEqual(XLSX_PARTS);
    expect(result.after).toEqual(result.before);
    expect(result.saved).toEqual(Array.from(input));
  });
}

// Saving an untouched file never compresses anything, so the engine's own
// deflate only runs in the browser when a part is actually replaced.
for (const flavour of ['script', 'esm'] as Flavour[]) {
  test(`no native compression (${flavour}): an edited part is deflated by the engine and readable elsewhere`, async ({ page }) => {
    await page.addInitScript(() => {
      delete (window as any).DecompressionStream;
      delete (window as any).CompressionStream;
    });
    const input = await xlsx();
    await gotoHarness(page, flavour);

    const content = `<?xml version="1.0"?><worksheet><sheetData>${'<row><c><v>ğüş</v></c></row>'.repeat(2000)}</sheetData></worksheet>`;
    const saved = await page.evaluate(
      async ([bytes, xml]) => {
        const api = (window as any).__bym;
        const file = await api.open(new Uint8Array(bytes as number[]));
        file.package.writePart('/xl/worksheets/sheet1.xml', new TextEncoder().encode(xml as string));
        return Array.from((await file.save()) as Uint8Array);
      },
      [Array.from(input), content] as const,
    );

    // Node has the built-in streams: this reads the engine's output with another inflate.
    const zip = ZipArchive.read(new Uint8Array(saved));
    expect(zip.entry('xl/worksheets/sheet1.xml')?.method).toBe(8);
    expect(saved.length).toBeLessThan(content.length / 10);
    expect(new TextDecoder().decode(await zip.read('xl/worksheets/sheet1.xml'))).toBe(content);
  });
}
