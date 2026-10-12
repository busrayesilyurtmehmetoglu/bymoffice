// A workbook with charts is saved as it was opened: reading a chart into the
// model changes nothing of the file, an edit of a cell leaves the chart and
// drawing parts byte for byte as they were, and rows and columns still cannot
// be moved under a chart.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, open, type DrawingChart } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { History, type Op } from '../src/ops/ops';
import { renameSheet } from '../src/ops/sheets';
import { catAx, chartFrameOf, chartWorkbook, chartXml, group, legend, numRef, ser, sheetData, strRef, title, valAx } from './chart-fixtures';
import { twoCell } from './sheet-drawing-fixtures';

const same = (a: Uint8Array, b: Uint8Array): boolean => a.length === b.length && a.every((value, i) => value === b[i]);

const CHART = chartXml({
  head: title('Satışlar'),
  // Something the reader has no field for: it must still be there afterwards.
  before: '<c:lang val="tr-TR"/><!-- kept --><c:unknownThing xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" odd="1"/>',
  plot: group('barChart', '<c:barDir val="col"/><c:grouping val="clustered"/>', ser(0, { tx: strRef('Sayfa1!$B$1', ['Adet']), cat: strRef('Sayfa1!$A$2:$A$4', ['a', 'b', 'c']), val: numRef('Sayfa1!$B$2:$B$4', [1, 2, 3]) })) + catAx() + valAx(),
  tail: legend('r') + '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>',
});

const book = () => chartWorkbook({ sheet: sheetData([['Ürün', 'Adet'], ['Çay', 10], ['Kahve', 20], ['Süt', 30]]), drawing: twoCell([3, 1], [9, 15], chartFrameOf('rIdC1')), charts: [CHART], chartSheet: CHART });

const CHART_PARTS = [
  'xl/charts/chart1.xml',
  'xl/charts/chart2.xml',
  'xl/drawings/drawing1.xml',
  'xl/drawings/drawing9.xml',
  'xl/drawings/_rels/drawing1.xml.rels',
  'xl/drawings/_rels/drawing9.xml.rels',
  'xl/worksheets/_rels/sheet1.xml.rels',
  'xl/chartsheets/sheet2.xml',
  'xl/chartsheets/_rels/sheet2.xml.rels',
];

test('a workbook with a chart on a sheet and a chart sheet: opened, read into the model and saved, it is the file that was opened', async () => {
  const bytes = await book();
  const file = await open(bytes);
  const workbook = await file.workbook();
  assert.equal((workbook.sheets[0].drawings![0].object as DrawingChart).type, 'chart', 'the chart was read');
  assert.equal(workbook.sheets[1].drawings![0].object.type, 'chart', 'that of the chart sheet too');
  assert.ok(same(await file.save(), bytes), 'saved without a change: byte for byte');
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'written from the model without a change: byte for byte');
});

test('a cell is edited: every part of the charts and drawings is byte for byte what it was; undone, so is the file', async () => {
  const bytes = await book();
  const file = await open(bytes);
  const workbook = await file.workbook();
  const history = new History(workbook);
  history.apply(setCellInput(workbook, 0, 1, 1, '99'), undefined, undefined);
  await writeXlsx(file.package!, workbook);
  const saved = await file.package!.save();
  assert.ok(!same(saved, bytes), 'the edit is in the file');
  const before = ZipArchive.read(bytes);
  const after = ZipArchive.read(saved);
  for (const part of CHART_PARTS) assert.ok(same(await after.read(part), await before.read(part)), `${part} is as it was`);
  assert.ok(new TextDecoder().decode(await after.read('xl/charts/chart1.xml')).includes('<c:unknownThing'), 'what the reader does not know is still in the chart part');
  // The chart of the saved file is read again, and now finds the new number in its cell.
  const reopened = await (await open(saved)).workbook();
  assert.equal(reopened.sheets[0].drawings![0].object.type, 'chart');
  assert.equal(reopened.sheets[0].rows.get(1)!.cells[1].value, 99);

  history.undo();
  await writeXlsx(file.package!, workbook);
  assert.ok(same(await file.package!.save(), bytes), 'undone: byte for byte the file that was opened');
});

test('rows, columns and sheet names still cannot be changed under a chart', async () => {
  const workbook = await (await open(await book())).workbook();
  assert.deepEqual(workbook.sheets[0].anchored, ['chart', 'drawing']);
  const history = new History(workbook);
  for (const op of [
    { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 },
    { type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 },
  ] as Op[]) {
    assert.throws(() => history.apply([op], undefined, undefined), /chart/, `${op.type} is refused`);
  }
  assert.throws(() => renameSheet(workbook, 0, 'Başka'), /chart/, 'renaming is refused');
  assert.equal(workbook.sheets[0].name, 'Sayfa1');
});
