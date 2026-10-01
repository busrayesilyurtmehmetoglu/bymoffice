import { XLSX_PARTS, readPartsInNode, xlsx } from './fixtures/files';
import { expect, gotoHarness, roundTripInPage, test, type Flavour } from './helpers';

const EXPORTS = ['open', 'ZipArchive', 'OpcPackage', 'CfbFile', 'parseXml', 'serializeXml', 'UnsupportedFileError'];

for (const flavour of ['script', 'esm'] as Flavour[]) {
  test.describe(`smoke: ${flavour === 'script' ? '<script> dist/bymoffice.min.js' : 'ESM dist/bymoffice.esm.js'}`, () => {
    test('exposes the public API', async ({ page }) => {
      await gotoHarness(page, flavour);
      const seen = await page.evaluate(
        (names) => ({
          exports: names.filter((name) => typeof (window as any).__bym[name] === 'function'),
          global: typeof (window as any).BymOffice,
        }),
        EXPORTS,
      );
      expect(seen.exports).toEqual(EXPORTS);
      // The classic bundle defines the global; the ES module must not leak one.
      expect(seen.global).toBe(flavour === 'script' ? 'object' : 'undefined');
    });

    test('opens an xlsx, saves it and reopens it byte-identically', async ({ page }) => {
      const input = await xlsx();
      await gotoHarness(page, flavour);
      const result = await roundTripInPage(page, input);

      expect(result.format).toBe('xlsx');
      expect(result.kind).toBe('spreadsheet');
      expect(result.container).toBe('ooxml');
      expect(result.parts).toEqual(XLSX_PARTS);
      expect(result.before['/customXml/unknown.bin']).toEqual([0, 1, 2, 3, 255, 254, 253]);

      expect(result.reopenedFormat).toBe('xlsx');
      expect(result.reopenedParts).toEqual(XLSX_PARTS);
      expect(result.after).toEqual(result.before);

      // Nothing was modified, so the archive itself must come back unchanged.
      expect(result.saved).toEqual(Array.from(input));
      // The bytes the browser produced are also readable outside the browser.
      expect(await readPartsInNode(Uint8Array.from(result.saved))).toEqual(result.before);
    });
  });
}
