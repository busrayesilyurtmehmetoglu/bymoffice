import { expect, gotoHarness, test } from './helpers';

// Opening a large sheet must stay well inside what a user tolerates. The file
// is built in the page so that 40 MB of XML never crosses the test protocol.
const ROWS = 100_000;
const COLS = 5;
const LIMIT_MS = 5_000;

test(`performance: a ${ROWS} row sheet is read into the model`, async ({ page }) => {
  test.setTimeout(120_000);
  await gotoHarness(page, 'script');

  const result = await page.evaluate(
    async ([rows, cols]) => {
      const api = (window as any).__bym;
      const utf8 = (text: string) => new TextEncoder().encode(text);
      const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      const main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
      const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
      const chunks: string[] = [];
      for (let r = 1; r <= rows; r++) {
        let row = `<row r="${r}">`;
        for (let c = 0; c < cols; c++) {
          const ref = String.fromCharCode(65 + c) + r;
          row += c === 0 ? `<c r="${ref}" t="s"><v>${r % 100}</v></c>` : `<c r="${ref}"><v>${r * c + 0.5}</v></c>`;
        }
        chunks.push(`${row}</row>`);
      }
      const strings = Array.from({ length: 100 }, (_, i) => `<si><t>Satır ${i} ğüş</t></si>`).join('');
      const zip = new api.ZipArchive();
      zip.set(
        '[Content_Types].xml',
        utf8(
          `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
            '<Default Extension="xml" ContentType="application/xml"/>' +
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>',
        ),
      );
      zip.set(
        '_rels/.rels',
        utf8(`${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
      );
      zip.set('xl/workbook.xml', utf8(`${head}<workbook xmlns="${main}" xmlns:r="${rel}"><sheets><sheet name="Büyük" sheetId="1" r:id="rId1"/></sheets></workbook>`));
      zip.set(
        'xl/_rels/workbook.xml.rels',
        utf8(
          `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>` +
            `<Relationship Id="rId2" Type="${rel}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`,
        ),
      );
      zip.set('xl/sharedStrings.xml', utf8(`${head}<sst xmlns="${main}">${strings}</sst>`));
      zip.set('xl/worksheets/sheet1.xml', utf8(`${head}<worksheet xmlns="${main}"><sheetData>${chunks.join('')}</sheetData></worksheet>`));
      const bytes: Uint8Array = await zip.write();

      const started = performance.now();
      const file = await api.open(bytes);
      const workbook = await file.workbook();
      const ms = performance.now() - started;
      const sheet = workbook.sheets[0];
      return {
        ms,
        size: bytes.length,
        rowCount: sheet.rowCount,
        colCount: sheet.colCount,
        first: api.getCell(sheet, 0, 0).value,
        last: api.getCell(sheet, rows - 1, cols - 1).value,
      };
    },
    [ROWS, COLS] as const,
  );

  expect([result.rowCount, result.colCount]).toEqual([ROWS, COLS]);
  expect(result.first).toBe('Satır 1 ğüş');
  expect(result.last).toBe(ROWS * (COLS - 1) + 0.5);
  test.info().annotations.push({ type: 'open+read ms', description: result.ms.toFixed(0) });
  console.log(`[${test.info().project.name}] ${ROWS} rows: ${result.ms.toFixed(0)} ms, ${result.size} bytes zipped`);
  expect(result.ms).toBeLessThan(LIMIT_MS);
});
