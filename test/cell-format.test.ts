import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR, cellDisplayText, emptyStyleSheet, formatCell, isoToSerial, open, type Workbook } from '../src/index';
import { sampleWorkbook } from './fixtures';

test('cell format: value kinds', () => {
  assert.deepEqual(formatCell(undefined, '0.00'), { text: '' });
  assert.deepEqual(formatCell({ kind: 'blank', value: null, style: 0 }, '0.00'), { text: '' });
  assert.equal(formatCell({ kind: 'number', value: 1234.5, style: 0 }, '#,##0.00').text, '1,234.50');
  assert.equal(formatCell({ kind: 'number', value: 1234.5, style: 0 }, '#,##0.00', { locale: LOCALE_TR }).text, '1.234,50');
  assert.equal(formatCell({ kind: 'string', value: 'metin', style: 0 }, '0.00').text, 'metin');
  assert.equal(formatCell({ kind: 'string', value: 'metin', style: 0 }, '0;0;0;"<"@">"').text, '<metin>');
  assert.equal(formatCell({ kind: 'boolean', value: true, style: 0 }, 'General').text, 'TRUE');
  assert.equal(formatCell({ kind: 'boolean', value: false, style: 0 }, 'General', { locale: LOCALE_TR }).text, 'YANLIŞ');
  // Errors are never run through a format.
  assert.equal(formatCell({ kind: 'error', value: '#DIV/0!', style: 0 }, '0.00').text, '#DIV/0!');

  const negative = formatCell({ kind: 'number', value: -5, style: 0 }, '0.00;[Red](0.00)');
  assert.deepEqual([negative.text, negative.color], ['(5.00)', '#FF0000']);
  assert.equal(formatCell({ kind: 'number', value: -1, style: 0 }, 'yyyy-mm-dd').overflow, true);
});

test('cell format: display text follows the cell style and the workbook date system', async () => {
  const workbook = await (await open(await sampleWorkbook())).workbook();
  const sheet = workbook.sheets[0];
  const cell = (row: number, col: number) => sheet.rows.get(row)?.cells[col];
  // B2 carries the custom format #,##0.00\ "₺".
  assert.equal(cellDisplayText(workbook, cell(1, 1)), '12.50 ₺');
  assert.equal(cellDisplayText(workbook, cell(1, 1), LOCALE_TR), '12,50 ₺');
  assert.equal(cellDisplayText(workbook, cell(0, 0)), 'Ürün');
  assert.equal(cellDisplayText(workbook, cell(1, 2), LOCALE_TR), 'DOĞRU');
  assert.equal(cellDisplayText(workbook, cell(2, 2)), '#DIV/0!');
  assert.equal(cellDisplayText(workbook, undefined), '');

  // The sample uses the 1904 date system: serial 0 is 1 January 1904.
  assert.equal(workbook.date1904, true);
  workbook.styles.numFmts.set(200, 'yyyy-mm-dd');
  workbook.styles.cellXfs.push({ numFmtId: 200, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false });
  const dated = { kind: 'number' as const, value: 0, style: workbook.styles.cellXfs.length - 1 };
  assert.equal(cellDisplayText(workbook, dated), '1904-01-01');
});

test('cell format: ISO date cells are shown as the date they state', () => {
  const iso = (value: string) => ({ kind: 'date' as const, value, style: 0 });
  assert.equal(formatCell(iso('2024-03-15T00:00:00Z'), 'dd/mm/yyyy', { locale: LOCALE_TR }).text, '15.03.2024');
  assert.equal(formatCell(iso('2024-03-15'), 'd mmmm yyyy', { locale: LOCALE_TR }).text, '15 Mart 2024');
  assert.equal(formatCell(iso('2024-03-15T13:45:30.5'), 'yyyy-mm-dd hh:mm:ss.0').text, '2024-03-15 13:45:30.5');
  // Without a date format, and for text that is not a date, the text stays.
  assert.equal(formatCell(iso('2024-03-15T00:00:00Z'), 'General').text, '2024-03-15T00:00:00Z');
  assert.equal(formatCell(iso('2024-02-31'), 'yyyy-mm-dd').text, '2024-02-31');
  assert.equal(formatCell(iso('yarın'), 'yyyy-mm-dd').text, 'yarın');

  assert.equal(isoToSerial('1900-01-01'), 1);
  assert.equal(isoToSerial('1900-02-28'), 59);
  assert.equal(isoToSerial('1900-03-01'), 61);
  assert.equal(isoToSerial('2024-03-15'), 45366);
  assert.equal(isoToSerial('2024-03-15T12:00:00+03:00'), 45366.5);
  assert.equal(isoToSerial('1904-01-01', true), 0);
  assert.equal(isoToSerial('2024-03-15', true), 45366 - 1462);
  assert.equal(isoToSerial('9999-12-31'), 2958465);
  for (const bad of ['', '2024-13-01', '2024-00-10', '2023-02-29', '2024-03-15T24:00:00', '2024-03-15T10:60', '15.03.2024', '2024-3-5']) {
    assert.equal(isoToSerial(bad), undefined, bad);
  }
});

test('cell format: built-in dates follow the locale, the same code written by the file does not', () => {
  const styles = emptyStyleSheet();
  const workbook: Workbook = { sheets: [], styles, sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  const styleOf = (numFmtId: number) => styles.cellXfs.push({ numFmtId, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false }) - 1;
  const builtin = { kind: 'number' as const, value: 45366.25, style: styleOf(14) };
  const builtinTime = { kind: 'number' as const, value: 45366.25, style: styleOf(22) };
  styles.numFmts.set(170, 'mm-dd-yy');
  const custom = { kind: 'number' as const, value: 45366.25, style: styleOf(170) };

  assert.equal(cellDisplayText(workbook, builtin), '3/15/2024');
  assert.equal(cellDisplayText(workbook, builtin, LOCALE_TR), '15.03.2024');
  assert.equal(cellDisplayText(workbook, builtinTime), '3/15/2024 6:00');
  assert.equal(cellDisplayText(workbook, builtinTime, LOCALE_TR), '15.03.2024 06:00');
  assert.equal(cellDisplayText(workbook, custom), '03-15-24');
  assert.equal(cellDisplayText(workbook, custom, LOCALE_TR), '03-15-24');

  // A file that redefines id 14 means its own code, not the built-in.
  styles.numFmts.set(14, 'yyyy/mm/dd');
  assert.equal(cellDisplayText(workbook, builtin), '2024/03/15');
});
