import assert from 'node:assert/strict';
import { test } from 'node:test';
import { UnsupportedFileError, getCell, open, resolveCellStyle, resolveColor } from '../src/index';
import { buildXlsx, minimalCfb, minimalXlsx, sampleWorkbook, worksheetXml } from './fixtures';

test('xlsx: workbook, sheets, cells, styles and names are read into the model', async () => {
  const file = await open(await sampleWorkbook());
  const workbook = await file.workbook();
  assert.equal(await file.workbook(), workbook, 'the model is read once');

  assert.deepEqual(
    workbook.sheets.map((s) => [s.name, s.sheetId, s.state, s.kind, s.part]),
    [
      ['Veri', 1, 'visible', 'worksheet', '/xl/worksheets/sheet1.xml'],
      ['Gizli', 2, 'hidden', 'worksheet', '/xl/worksheets/sheet2.xml'],
      ['Grafik', 3, 'visible', 'chartsheet', '/xl/chartsheets/sheet3.xml'],
    ],
  );
  assert.equal(workbook.date1904, true);
  assert.equal(workbook.activeSheet, 1);
  assert.deepEqual(workbook.definedNames, [
    { name: 'Toplam', value: 'Veri!$B$2:$B$4', localSheet: undefined, hidden: false },
    { name: '_xlnm.Print_Area', value: 'Veri!$A$1:$C$4', localSheet: 0, hidden: true },
  ]);

  const sheet = workbook.sheets[0];
  assert.deepEqual([sheet.rowCount, sheet.colCount], [4, 3]);
  assert.deepEqual(sheet.dimension, { top: 0, left: 0, bottom: 3, right: 2 });
  assert.deepEqual(
    [sheet.view.showGridLines, sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.topLeftCell, sheet.view.activeCell],
    [false, 1, 1, 'B2', 'B3'],
  );
  assert.equal(sheet.defaultRowHeight, 15);
  assert.deepEqual(
    sheet.columns.map((c) => [c.min, c.max, c.width, c.hidden]),
    [
      [0, 0, 18.5, false],
      [2, 3, 9, true],
    ],
  );
  assert.deepEqual([sheet.rows.get(0)?.height, sheet.rows.get(0)?.customHeight, sheet.rows.get(2)?.hidden], [24, true, true]);
  assert.deepEqual(sheet.merges, [{ top: 4, left: 0, bottom: 5, right: 2 }]);

  const values = (row: number) => [0, 1, 2].map((col) => getCell(sheet, row, col)?.value);
  assert.deepEqual(values(0), ['Ürün', 'Fiyat', undefined]);
  assert.deepEqual(values(1), [' Çay ', 12.5, true]);
  assert.deepEqual(values(2), ['Kalın şeker', 7, '#DIV/0!']);
  assert.deepEqual(values(3), ['Toplam', 19.5, undefined]);
  assert.deepEqual(
    [1, 2].map((col) => getCell(sheet, 2, col)?.kind),
    ['number', 'error'],
  );
  assert.equal(getCell(sheet, 1, 2)?.kind, 'boolean');
  assert.deepEqual(getCell(sheet, 3, 1)?.formula, { text: 'SUM(B2:B3)', type: 'normal' });

  const runs = getCell(sheet, 2, 0)?.runs;
  assert.equal(runs?.length, 2);
  assert.deepEqual([runs?.[1].text, runs?.[1].font?.bold, runs?.[1].font?.size, runs?.[1].font?.name], [' şeker', true, 14, 'Arial']);
  assert.equal(runs, workbook.sharedStrings[3].runs, 'rich runs are shared, not copied');

  assert.deepEqual(
    sheet.hyperlinks.map((h) => [h.range.top, h.target, h.location, h.tooltip]),
    [
      [1, 'https://example.com/çay', undefined, 'Çay'],
      [2, undefined, 'Gizli!A1', undefined],
    ],
  );

  const header = resolveCellStyle(workbook.styles, getCell(sheet, 0, 0)!.style);
  assert.deepEqual([header.font.bold, header.font.size, header.alignment.horizontal, header.alignment.wrapText], [true, 12, 'center', true]);
  assert.equal(resolveColor(header.font.color, workbook.styles), '#4472C4');
  assert.equal(header.fill.type === 'pattern' && resolveColor(header.fill.fg, workbook.styles), '#FFFF00');
  assert.deepEqual([header.border.left?.style, header.border.bottom?.style], ['thin', 'medium']);
  assert.equal(resolveColor(header.border.bottom?.color, workbook.styles), '#0000FF');
  assert.equal(resolveCellStyle(workbook.styles, getCell(sheet, 1, 1)!.style).numFmt, '#,##0.00\\ "₺"');

  assert.equal(workbook.sheets[1].rowCount, 0);
  assert.equal(workbook.sheets[2].rows.size, 0);
});

test('xlsx: a workbook without styles, shared strings or theme still reads', async () => {
  const file = await open(await minimalXlsx());
  const workbook = await file.workbook();
  assert.deepEqual(workbook.sheets.map((s) => s.name), ['Sayfa1']);
  assert.equal(workbook.sharedStrings.length, 0);
  assert.equal(workbook.styles.cellXfs.length, 1);
  assert.deepEqual([workbook.date1904, workbook.activeSheet], [false, 0]);
});

test('xlsx: reading the model does not change what save() writes', async () => {
  const bytes = await sampleWorkbook();
  const file = await open(bytes);
  await file.workbook();
  assert.deepEqual(await file.save(), bytes);
});

test('xlsx: broken workbooks are rejected with a clear error', async () => {
  const noSheets = await open(await buildXlsx({ sheets: [] }));
  await assert.rejects(noSheets.workbook(), /no sheets/);

  const badIndex = await open(
    await buildXlsx({
      sharedStrings: '<si><t>a</t></si>',
      sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1" t="s"><v>5</v></c></row></sheetData>') }],
    }),
  );
  await assert.rejects(badIndex.workbook(), Error);

  const activeOutOfRange = await open(
    await buildXlsx({
      workbookHead: '<bookViews><workbookView activeTab="9"/></bookViews>',
      sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }],
    }),
  );
  assert.equal((await activeOutOfRange.workbook()).activeSheet, 0);
});

test('workbook(): an xls file whose stream is not a workbook is reported as damaged', async () => {
  const xls = await open(minimalCfb(new Uint8Array(9 * 512), new Uint8Array(10)));
  await assert.rejects(xls.workbook(), /Invalid xls/);
});

test('styles: a number attribute of a million digits is refused at once', async () => {
  const started = Date.now();
  const bytes = await buildXlsx({
    sheets: [{ name: 'S', xml: worksheetXml('<sheetData/>') }],
    styles: `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><color theme="1" tint="${'1'.repeat(1000000)}x"/></font></fonts></styleSheet>`,
  });
  await (await open(bytes)).workbook();
  assert.ok(Date.now() - started < 3000);
});

test('xlsx: a file whose model cannot be read is still saved as it was', async () => {
  const bytes = await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>abc</v></c></row></sheetData>') }] });
  const file = await open(bytes);
  await assert.rejects(file.workbook(), /is not a number/);
  assert.deepEqual(await file.save(), bytes);
  await assert.rejects(file.workbook(), /is not a number/);
});
