// buildXlsx() on models that are not what a reader hands out: one read from an
// xlsx file and edited since, and models made in code that break what the
// model promises. Whatever the model says, the file is one a schema accepts,
// and what was changed on the way is in the report.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { applyOp, buildXlsx, getCell, open, parseXml, type Workbook, type XlsxBuildReport } from '../src/index';
import { resolveCellStyle } from '../src/model/sheet/color';
import { MAX_COLS, MAX_ROWS, type DrawingShape, type Row } from '../src/model/sheet/types';
import { comment, commentsXml, drawingWorkbook, sp, twoCell, vmlNote, vmlXml } from './sheet-drawing-fixtures';
import { canonical, workbookContent, workbookDifference } from './workbook-content';
import { allNamed, bookOf, build, put, sheetOf, text } from './xlsx-build-fixtures';

// --- rows and columns moved since the file was read -------------------------------------------

/** A5 with a value, a comment, a conditional format, a validation and a shape over it. */
function anchoredBook(): Promise<Uint8Array> {
  return drawingWorkbook({
    sheet:
      '<sheetData><row r="5"><c r="A5"><v>7</v></c><c r="B5"><v>1</v></c></row><row r="9"><c r="A9"><v>9</v></c></row></sheetData>' +
      '<conditionalFormatting sqref="A5:A6"><cfRule type="expression" priority="1"><formula>A5&gt;$B$5</formula></cfRule></conditionalFormatting>' +
      '<dataValidations count="1"><dataValidation type="whole" operator="greaterThan" sqref="A5"><formula1>$B$5</formula1></dataValidation></dataValidations>',
    drawing: twoCell([0, 4], [2, 6], sp({ name: 'Kutu' })) + twoCell([3, 4, 5, 5], [4, 6], sp({ name: 'Sabit' }), 'absolute') + twoCell([0, 8], [1, 9], sp({ name: 'Taşınan' }), 'oneCell'),
    comments: commentsXml(['A'], comment('A5', 0, '<t>beşinci satır</t>') + comment('A9', 0, '<t>dokuzuncu</t>')),
    vml: vmlXml(vmlNote(4, 0, { anchor: '1, 15, 3, 10, 3, 15, 7, 4' }) + vmlNote(8, 0)),
  });
}

/** What moved things look like in a model read from a file: where they are, by the cells they stand on. */
function placed(workbook: Workbook): unknown {
  const sheet = workbook.sheets[0];
  return {
    cells: [...sheet.rows.entries()].sort((a, b) => a[0] - b[0]).map(([index, row]) => [index, Array.from(row.cells, (cell) => (cell ? cell.value : null))]),
    comments: sheet.comments?.map((item) => [item.row, item.col, item.text, item.box]),
    conditional: sheet.conditionalFormats?.map((format) => [format.ranges, format.rules.map((rule) => rule.formulas)]),
    validations: sheet.dataValidations?.map((validation) => [validation.ranges, validation.formula1]),
    drawings: sheet.drawings?.map((drawing) => [drawing.object.name, drawing.anchor.from?.row, drawing.anchor.to?.row, drawing.anchor.from?.col, drawing.anchor.to?.col]),
  };
}

test('buildXlsx: comments, conditional formats, validations and drawings follow rows inserted since the file was read', async () => {
  const file = await open(await anchoredBook());
  const model = await file.workbook();
  applyOp(model, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 3 });
  applyOp(model, { type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 });
  // The model keeps these where the file has them, with the edits beside them.
  assert.equal(model.sheets[0].comments?.[0].row, 4);
  assert.equal(model.sheets[0].structuralEdits?.length, 2);

  const built = await (await open(await buildXlsx(model))).workbook();
  assert.deepEqual(placed(built), {
    cells: [[7, [null, 7, 1]], [11, [null, 9]]],
    comments: [
      [7, 1, 'beşinci satır', { kind: 'twoCell', from: { col: 2, colOff: 15 * 9525, row: 6, rowOff: 10 * 9525 }, to: { col: 4, colOff: 15 * 9525, row: 10, rowOff: 4 * 9525 } }],
      [11, 1, 'dokuzuncu', { kind: 'twoCell', from: { col: 2, colOff: 15 * 9525, row: 10, rowOff: 2 * 9525 }, to: { col: 4, colOff: 15 * 9525, row: 14, rowOff: 16 * 9525 } }],
    ],
    conditional: [[[{ top: 7, left: 1, bottom: 8, right: 1 }], [['B8>$C$8']]]],
    validations: [[[{ top: 7, left: 1, bottom: 7, right: 1 }], '$C$8']],
    drawings: [['Kutu', 7, 9, 1, 3], ['Sabit', 4, 6, 3, 4], ['Taşınan', 11, 12, 1, 2]],
  });
  // The same places as the file its own writer makes of the edited model.
  const saved = await (await open(await file.save())).workbook();
  assert.equal(canonical(placed(built)), canonical(placed(saved)));
  // The model is not changed by being written.
  assert.equal(model.sheets[0].comments?.[0].row, 4);
  assert.equal(model.sheets[0].structuralEdits?.length, 2);
});

test('buildXlsx: what stood on deleted rows goes with them, as in a saved file', async () => {
  const file = await open(await anchoredBook());
  const model = await file.workbook();
  applyOp(model, { type: 'delete', sheet: 0, axis: 'row', at: 3, count: 2 });
  const built = await (await open(await buildXlsx(model))).workbook();
  const saved = await (await open(await file.save())).workbook();
  assert.equal(canonical(placed(built)), canonical(placed(saved)));
  assert.deepEqual(built.sheets[0].comments?.map((item) => [item.row, item.text]), [[6, 'dokuzuncu']]);
  assert.deepEqual(built.sheets[0].dataValidations, undefined);
  assert.deepEqual(built.sheets[0].conditionalFormats?.map((format) => format.ranges), [[{ top: 3, left: 0, bottom: 3, right: 0 }]]);
});

test('buildXlsx: a validation that names another sheet follows the rows moved in that sheet', async () => {
  const data = sheetOf('Veri');
  put(data, 5, 0, { kind: 'number', value: 1 });
  data.structuralEdits = [{ kind: 'insert', axis: 'row', at: 0, count: 2 }];
  const form = sheetOf("Form 'x'", 2);
  form.dataValidations = [{ ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], type: 'list', operator: 'between', formula1: 'Veri!$A$4:$A$6', allowBlank: true, hideDropDown: false, showInputMessage: false, showErrorMessage: false, errorStyle: 'stop' }];
  form.conditionalFormats = [{ ranges: [{ top: 0, left: 0, bottom: 0, right: 0 }], rules: [{ type: 'expression', priority: 1, stopIfTrue: false, formulas: ['Veri!$A$4>A1'], percent: false, bottom: false, aboveAverage: true, equalAverage: false }] }];
  const { after } = await build(bookOf(data, form));
  assert.equal(after.sheets[1].dataValidations?.[0].formula1, 'Veri!$A$6:$A$8');
  assert.deepEqual(after.sheets[1].conditionalFormats?.[0].rules[0].formulas, ['Veri!$A$6>A1']);
});

// --- indexes and tables ------------------------------------------------------------------------

test('buildXlsx: a format index outside the table is written as the default format, and counted', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1, style: 7 });
  put(sheet, 0, 1, { kind: 'number', value: 2, style: -1 });
  put(sheet, 0, 2, { kind: 'number', value: 3, style: 1.5 });
  put(sheet, 1, 0, { kind: 'number', value: 4, style: 1 });
  sheet.rows.get(1)!.style = 9;
  sheet.columns.push({ min: 3, max: 4, style: 5, customWidth: false, hidden: false, outlineLevel: 0, collapsed: false });
  const row: Row = { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [], attrs: [['s', '12']] };
  sheet.rows.set(3, row);
  const workbook = bookOf(sheet);
  workbook.styles.cellXfs.push({ numFmtId: 9, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false });
  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, ['Format: a cell, row or column refers to a format the workbook does not have and is written with the default one (6 times)']);
  assert.deepEqual([0, 1, 2].map((col) => getCell(after.sheets[0], 0, col)?.style), [0, 0, 0]);
  assert.equal(getCell(after.sheets[0], 1, 0)?.style, 1);
  assert.equal(after.sheets[0].rows.get(1)?.style, 0);
  assert.equal(after.sheets[0].columns[0].style, 0);
  const xml = text(await pkg.readPart('/xl/worksheets/sheet1.xml'));
  for (const match of xml.match(/ (s|style)="(\d+)"/g) || []) assert.ok(Number(/\d+/.exec(match)![0]) < 2, match);
  assert.doesNotMatch(xml, /<row r="4"[^>]* s="12"/);
});

test('buildXlsx: a format that refers to a record its table does not have gets the first one, and it is said', async () => {
  const workbook = bookOf(sheetOf('S'));
  workbook.styles.cellXfs.push({ numFmtId: 0, fontId: 5, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false }, { numFmtId: 0, fontId: 0, fillId: 9, borderId: -1, xfId: 3, quotePrefix: false });
  const { after, warnings } = await build(workbook);
  assert.deepEqual(warnings, ['Styles: a format refers to a font, a fill, a border or a parent the tables do not have; the first one is used (2 times)']);
  assert.deepEqual(after.styles.cellXfs.map((xf) => [xf.fontId, xf.fillId, xf.borderId, xf.xfId]), [[0, 0, 0, 0], [0, 0, 0, 0], [0, 0, 0, 0]]);
});

test('buildXlsx: the first two fills of a file are fixed; the fills of the model come after them', async () => {
  const sheet = sheetOf('S');
  const workbook = bookOf(sheet);
  workbook.styles.fills = [
    { type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFF0000' } },
    { type: 'pattern', pattern: 'solid', fg: { rgb: 'FF00FF00' } },
  ];
  workbook.styles.cellStyleXfs = [{ numFmtId: 0, fontId: 0, fillId: 1, borderId: 0, quotePrefix: false }];
  workbook.styles.cellXfs = [
    { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false },
    { numFmtId: 0, fontId: 0, fillId: 1, borderId: 0, xfId: 0, quotePrefix: false },
  ];
  put(sheet, 0, 0, { kind: 'number', value: 1, style: 0 });
  put(sheet, 0, 1, { kind: 'number', value: 2, style: 1 });
  const { after } = await build(workbook);
  assert.deepEqual(after.styles.fills, [{ type: 'pattern', pattern: 'none' }, { type: 'pattern', pattern: 'gray125' }, ...workbook.styles.fills]);
  assert.deepEqual(after.styles.cellXfs.map((xf) => xf.fillId), [2, 3]);
  assert.deepEqual(after.styles.cellStyleXfs.map((xf) => xf.fillId), [3]);
  for (const col of [0, 1]) assert.equal(canonical(resolveCellStyle(after.styles, getCell(after.sheets[0], 0, col)!.style).fill), canonical(workbook.styles.fills[col]));
  assert.equal(await workbookDifference(workbook, after), undefined);

  // Only the first one is missing: the second stays where it is.
  const half = bookOf(sheetOf('S'));
  half.styles.fills = [{ type: 'pattern', pattern: 'solid', fg: { rgb: 'FFFF0000' } }, { type: 'pattern', pattern: 'gray125' }, { type: 'pattern', pattern: 'solid', fg: { rgb: 'FF0000FF' } }];
  half.styles.cellXfs.push({ numFmtId: 0, fontId: 0, fillId: 2, borderId: 0, xfId: 0, quotePrefix: false });
  const built = await build(half);
  assert.deepEqual(built.after.styles.fills.map((fill) => (fill.type === 'pattern' ? fill.pattern : '')), ['none', 'gray125', 'solid', 'gray125', 'solid']);
  assert.deepEqual(built.after.styles.cellXfs.map((xf) => xf.fillId), [2, 4]);
});

// --- what is written differently is said -------------------------------------------------------

test('buildXlsx: a formula or a name that refers to another workbook is written, and the missing link is reported', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1, formula: { text: '[1]Sheet1!$B$2+1', type: 'normal' } });
  put(sheet, 1, 0, { kind: 'number', value: 2, formula: { text: "SUM('[2]Other sheet'!A1:A3)", type: 'normal' } });
  put(sheet, 2, 0, { kind: 'number', value: 3, formula: { text: '[1]!Name*2', type: 'shared', si: 0, ref: 'A3:A4' } });
  put(sheet, 3, 0, { kind: 'number', value: 4, formula: { text: '', type: 'shared', si: 0 } });
  // Not references to a workbook: a text, and a column of a table.
  put(sheet, 4, 0, { kind: 'string', value: 'x', formula: { text: '"[1]Sheet1!A1"&Table1[[#This Row],[1]]', type: 'normal' } });
  const workbook = bookOf(sheet);
  workbook.definedNames.push({ name: 'Dış', value: '[1]Sheet1!$A$1', hidden: false }, { name: 'İç', value: 'S!$A$1', hidden: false });
  const { after, warnings } = await build(workbook);
  assert.deepEqual(warnings, [
    'Formula: a formula refers to another workbook; the link to that workbook is not in the model and is not written (3 times)',
    'Defined name: a name refers to another workbook; the link to that workbook is not in the model and is not written',
  ]);
  assert.equal(getCell(after.sheets[0], 0, 0)?.formula?.text, '[1]Sheet1!$B$2+1');
  assert.equal(after.definedNames[0].value, '[1]Sheet1!$A$1');
});

test('buildXlsx: characters XML cannot carry in a formula, an error or a date are replaced, and counted', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'string', value: 'a\u0001b', formula: { text: '"a\u0001b"&"\ud800"', type: 'normal' } });
  put(sheet, 1, 0, { kind: 'error', value: '#N/A\u0002' });
  put(sheet, 2, 0, { kind: 'date', value: '2024-01-01\u0003' });
  put(sheet, 3, 0, { kind: 'number', value: 1, formula: { text: 'A1', type: 'array', ref: 'A4', attrs: new Map([['aca', '1\u0004']]) } });
  // A line break and a carriage return are carried as they are.
  put(sheet, 4, 0, { kind: 'string', value: 'x', formula: { text: '"satır\n\rsonu"', type: 'normal' } });
  sheet.rows.get(4)!.attrs = [['thickBot', '1\u0005']];
  const { after, warnings } = await build(bookOf(sheet));
  assert.deepEqual(warnings, ['Text: a character XML cannot carry (in a formula, an error value, a name or another attribute) is written as U+FFFD (5 times)']);
  assert.equal(getCell(after.sheets[0], 0, 0)?.formula?.text, '"a�b"&"�"');
  // The text of the cell itself has an escape of its own and is kept.
  assert.equal(getCell(after.sheets[0], 0, 0)?.value, 'a\u0001b');
  assert.equal(getCell(after.sheets[0], 1, 0)?.value, '#N/A�');
  assert.equal(getCell(after.sheets[0], 4, 0)?.formula?.text, '"satır\n\rsonu"');
});

test('buildXlsx: a number that is not finite is written as #NUM!, and counted', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: NaN });
  put(sheet, 0, 1, { kind: 'number', value: -Infinity });
  put(sheet, 0, 2, { kind: 'number', value: 'x' as unknown as number });
  const { after, warnings } = await build(bookOf(sheet));
  assert.deepEqual(warnings, ['Cell: a number that is not finite is written as #NUM! (3 times)']);
  assert.deepEqual([0, 1, 2].map((col) => getCell(after.sheets[0], 0, col)?.value), ['#NUM!', '#NUM!', '#NUM!']);
});

// --- models that break what a model promises ---------------------------------------------------

test('buildXlsx: the used range counts what is written, however sparse or far out the rows of the model are', async () => {
  const sheet = sheetOf('S');
  put(sheet, 5, 2, { kind: 'number', value: 1 });
  put(sheet, 1.5, 0, { kind: 'number', value: 2 });
  put(sheet, MAX_ROWS + 3, 0, { kind: 'number', value: 3 });
  put(sheet, 3, MAX_COLS + 5, { kind: 'number', value: 4 });
  // A row whose first cell is a long way out: nothing may walk up to it.
  put(sheet, 4, 4_000_000_000, { kind: 'number', value: 5 });
  put(sheet, 4, 1, { kind: 'number', value: 6 });
  const started = performance.now();
  const { after, warnings, pkg } = await build(bookOf(sheet));
  assert.ok(performance.now() - started < 2000, 'a sparse row is not walked cell by cell');
  assert.match(text(await pkg.readPart('/xl/worksheets/sheet1.xml')), /<dimension ref="B5:C6"\/>/);
  assert.deepEqual([...after.sheets[0].rows.keys()], [4, 5]);
  assert.deepEqual(warnings, ['Sheet: a row outside the sheet is left out (2 times)', 'Sheet: cells outside the sheet are left out (2 times)']);
});

test('buildXlsx: cells of the view that are not one cell are written as one', async () => {
  const sheet = sheetOf('S');
  sheet.view = { ...sheet.view, activeCell: 'B2:C3', firstVisibleCell: 'D4:E5', frozenRows: 1, topLeftCell: 'A9:B10' };
  const other = sheetOf('T', 2);
  other.view = { ...other.view, activeCell: 'not a cell', firstVisibleCell: '$C$3' };
  const { after, pkg } = await build(bookOf(sheet, other));
  assert.deepEqual([after.sheets[0].view.activeCell, after.sheets[0].view.firstVisibleCell, after.sheets[0].view.topLeftCell], ['B2', 'D4', 'A9']);
  assert.deepEqual([after.sheets[1].view.activeCell, after.sheets[1].view.firstVisibleCell], [undefined, 'C3']);
  assert.match(text(await pkg.readPart('/xl/worksheets/sheet1.xml')), /activeCell="B2" sqref="B2"/);
});

test('buildXlsx: one name per scope, one link per cell', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  sheet.hyperlinks.push(
    { range: { top: 0, left: 0, bottom: 1, right: 1 }, target: 'https://example.com/ilk' },
    { range: { top: 1, left: 1, bottom: 1, right: 1 }, target: 'https://example.com/ikinci' },
    { range: { top: 2, left: 0, bottom: 2, right: 0 }, target: 'https://example.com/a b\nc\r\td\u007f?x=ç' },
  );
  const workbook = bookOf(sheet, sheetOf('T', 2));
  workbook.definedNames.push(
    { name: 'Ad', value: '1', hidden: false },
    { name: 'AD', value: '2', hidden: false },
    { name: 'Ad', value: '3', hidden: false, localSheet: 0 },
    { name: 'Ad', value: '4', hidden: false, localSheet: 1 },
    { name: 'ad', value: '5', hidden: false, localSheet: 1 },
  );
  const { after, warnings, pkg } = await build(workbook);
  assert.deepEqual(warnings, ['Hyperlink: a second link on the same cell is left out', 'Defined name: a second name with the same name and scope is left out (2 times)']);
  assert.deepEqual(after.definedNames.map((name) => [name.name, name.value, name.localSheet]), [['Ad', '1', undefined], ['Ad', '3', 0], ['Ad', '4', 1]]);
  // Control characters of a target are percent-encoded; nothing else of it is touched.
  assert.deepEqual(after.sheets[0].hyperlinks.map((link) => link.target), ['https://example.com/ilk', 'https://example.com/a b%0Ac%0D%09d%7F?x=ç']);
  assert.match(text(await pkg.readPart('/xl/worksheets/_rels/sheet1.xml.rels')), /Target="https:\/\/example.com\/a b%0Ac%0D%09d%7F\?x=ç"/);
});

test('buildXlsx: column ranges that overlap or state a width a file cannot hold', async () => {
  const sheet = sheetOf('S');
  const range = (min: number, max: number, width?: number) => ({ min, max, width, customWidth: true, hidden: false, outlineLevel: 0, collapsed: false });
  sheet.columns.push(range(0, 2, 10), range(1, 5, 20), range(4, 4, 30), range(7, 8, -3), range(9, 9, NaN), range(10, 10, 300), range(12, 11, 5), range(3, 3, 7));
  const { after, warnings, pkg } = await build(bookOf(sheet));
  assert.deepEqual(after.sheets[0].columns.map((item) => [item.min, item.max, item.width]), [[0, 2, 10], [3, 5, 20], [7, 8, undefined], [9, 9, undefined], [10, 10, 255]]);
  assert.deepEqual(warnings, ['Column: a range that overlaps an earlier one, or is not one, is cut to what is left of it (4 times)', 'Column: a width a file cannot state is left out or cut to 255 (3 times)']);
  // Ranges in a file do not overlap and come in order.
  let last = 0;
  for (const col of allNamed(parseXml(await pkg.readPart('/xl/worksheets/sheet1.xml')).root, 'col')) {
    assert.ok(Number(col.attrs.get('min')) > last && Number(col.attrs.get('max')) >= Number(col.attrs.get('min')));
    last = Number(col.attrs.get('max'));
  }
});

test('buildXlsx: formulas a file cannot state as the model has them', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1, formula: { text: 'SUM(B1:B2)', type: 'array', ref: 'ZZZ1:A0' } });
  put(sheet, 1, 0, { kind: 'number', value: 2, formula: { text: 'B2', type: 'array' } });
  put(sheet, 2, 0, { kind: 'number', value: 3, formula: { text: '', type: 'normal' } });
  put(sheet, 3, 0, { kind: 'number', value: 4, formula: { text: '', type: 'array', ref: 'A4' } });
  put(sheet, 4, 0, { kind: 'number', value: 5, formula: { text: '', type: 'dataTable', ref: 'nowhere', attrs: new Map([['r1', 'C1']]) } });
  // Two groups that claim the same number.
  put(sheet, 0, 3, { kind: 'number', value: 6, formula: { text: 'E1*2', type: 'shared', si: 0, ref: 'D1:D2' } });
  put(sheet, 1, 3, { kind: 'number', value: 7, formula: { text: '', type: 'shared', si: 0 } });
  put(sheet, 5, 3, { kind: 'number', value: 8, formula: { text: 'E6*3', type: 'shared', si: 0, ref: 'D6:D7' } });
  put(sheet, 6, 3, { kind: 'number', value: 9, formula: { text: '', type: 'shared', si: 0 } });
  const { after, warnings, pkg } = await build(bookOf(sheet));
  const formulaOf = (row: number, col: number) => getCell(after.sheets[0], row, col)?.formula;
  assert.deepEqual(formulaOf(0, 0), { text: 'SUM(B1:B2)', type: 'array', ref: 'A1' });
  assert.deepEqual(formulaOf(1, 0), { text: 'B2', type: 'array', ref: 'A2' });
  assert.deepEqual([formulaOf(2, 0), getCell(after.sheets[0], 2, 0)?.value], [undefined, 3]);
  assert.deepEqual([formulaOf(3, 0), getCell(after.sheets[0], 3, 0)?.value], [undefined, 4]);
  assert.deepEqual(formulaOf(4, 0), { text: '', type: 'dataTable', ref: 'A5', attrs: new Map([['r1', 'C1']]) });
  assert.deepEqual(formulaOf(0, 3), { text: 'E1*2', type: 'shared', si: 0, ref: 'D1:D2' });
  assert.deepEqual(formulaOf(1, 3), { text: '', type: 'shared', si: 0 });
  // The second group is written out; no cell of it may be read as one of the first.
  assert.deepEqual(formulaOf(5, 3), { text: 'E6*3', type: 'normal' });
  assert.deepEqual(formulaOf(6, 3), { text: 'E7*2', type: 'normal' });
  assert.deepEqual(warnings, [
    'Formula: an array formula or a data table whose range is not one is given its own cell as the range (3 times)',
    'Formula: a formula without a text is left out; the cell keeps its value (2 times)',
  ]);
  assert.doesNotMatch(text(await pkg.readPart('/xl/worksheets/sheet1.xml')), /<f\/>|<f t="array"[^>]*\/>/);
});

test('buildXlsx: attributes of a row that the writer states itself, or that are not names, are not written twice', async () => {
  const sheet = sheetOf('S');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  Object.assign(sheet.rows.get(0)!, { height: 20, attrs: [['r', '9'], ['ht', '99'], ['hidden', '1'], ['spans', '1:1'], ['customFormat', '1'], ['bad name', 'x'], ['a"b', 'y'], ['', 'z'], ['thickTop', '1'], ['thickTop', '0'], ['ph', '1'], ['xmlns', 'urn:x'], ['xml:space', 'preserve']] });
  const { after, warnings, pkg } = await build(bookOf(sheet));
  const xml = text(await pkg.readPart('/xl/worksheets/sheet1.xml'));
  assert.match(xml, /<row r="1" ht="20" thickTop="1" ph="1"><c r="A1"><v>1<\/v><\/c><\/row>/);
  assert.deepEqual([...after.sheets[0].rows.keys()], [0]);
  assert.deepEqual(after.sheets[0].rows.get(0)?.attrs, [['thickTop', '1'], ['ph', '1']]);
  assert.deepEqual(warnings, ['Row: an attribute that cannot be written as it is (one the writer states itself, one of another namespace, or one whose name is not a name) is left out (11 times)']);
});

test('buildXlsx: shape ids of comments stay inside the blocks their drawing names', async () => {
  for (const count of [1023, 1024, 1025, 2046, 2047]) {
    const sheet = sheetOf('S');
    sheet.comments = Array.from({ length: count }, (_, i) => ({ row: i, col: 0, text: 'n' }));
    const { after, pkg } = await build(bookOf(sheet));
    assert.equal(after.sheets[0].comments?.length, count);
    const vml = text(await pkg.readPart('/xl/drawings/vmlDrawing1.vml'));
    const blocks = /<o:idmap v:ext="edit" data="([\d,]+)"\/>/.exec(vml)![1].split(',').map(Number);
    assert.deepEqual(blocks, Array.from({ length: Math.ceil(count / 1023) }, (_, i) => i + 1), `${count} comments`);
    const ids = (vml.match(/_x0000_s\d+/g) || []).map((id) => Number(id.slice(8)));
    assert.equal(new Set(ids).size, count);
    // An id belongs to the block it is in: block n holds n*1024+1 .. n*1024+1023.
    for (const id of ids) assert.ok(blocks.includes(Math.floor(id / 1024)) && id % 1024 !== 0, `${count} comments: id ${id}`);
  }
});

test('buildXlsx: the active sheet of a file is one that is shown', async () => {
  const hidden = sheetOf('Gizli');
  hidden.state = 'hidden';
  const workbook = bookOf(hidden, sheetOf('İlk', 2), sheetOf('İkinci', 3));
  workbook.activeSheet = 0;
  const { after, warnings } = await build(workbook);
  assert.equal(after.activeSheet, 1);
  assert.deepEqual(warnings, ['Workbook: the active sheet is hidden; the first visible one is made the active sheet']);
  workbook.activeSheet = 2;
  assert.deepEqual([(await build(workbook)).after.activeSheet, (await build(workbook)).warnings], [2, []]);
});

// --- the comparison the tests and the corpus rely on --------------------------------------------

/** A model with one of everything the comparison looks at. */
function fullBook(): Workbook {
  const sheet = sheetOf('Her şey');
  put(sheet, 0, 0, { kind: 'number', value: 1 });
  sheet.baseColWidth = 10;
  sheet.defaultRowHeight = 13;
  const shape: DrawingShape = {
    type: 'shape',
    geometry: 'rect',
    name: 'Kutu',
    fill: '#FFFFFF',
    line: { color: '#000000', width: 1 },
    text: { insets: [3, 1, 3, 1], paragraphs: [{ runs: [{ text: 'kalın', bold: true, size: 12, color: '#FF0000', font: 'Arial' }, { text: ' düz\nalt satır', italic: false }] }] },
  };
  sheet.drawings = [
    { anchor: { kind: 'absolute', x: 0, y: 0, width: 952500, height: 476250 }, object: shape },
    { anchor: { kind: 'absolute', x: 0, y: 952500, width: 952500, height: 476250 }, object: { type: 'shape', geometry: 'ellipse' } },
  ];
  sheet.comments = [
    { row: 0, col: 0, author: 'A', text: 'tek parça', runs: [{ text: 'tek parça', font: { name: 'Tahoma', size: 9, bold: true, italic: false, strike: false, underline: 'none' } }] },
    { row: 2, col: 1, text: 'kutulu', visible: true, box: { kind: 'twoCell', from: { col: 2, colOff: 9525 * 3, row: 1, rowOff: 9525 * 4 }, to: { col: 4, colOff: 0, row: 6, rowOff: 9525 * 5 } } },
  ];
  const workbook = bookOf(sheet);
  workbook.styles.cellStyles = [{ name: 'Normal', xfId: 0, builtinId: 0 }, { name: 'Özel <stil>', xfId: 1 }];
  workbook.styles.cellStyleXfs.push({ numFmtId: 2, fontId: 0, fillId: 0, borderId: 0, quotePrefix: false });
  workbook.styles.dxfs = [parseXml('<dxf><font><b/><color rgb="FFFF0000"/></font><fill><patternFill><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>').root];
  return workbook;
}

test('workbook comparison: every field a file carries is looked at', async () => {
  const before = fullBook();
  const bytes = await buildXlsx(before);
  assert.equal(await workbookDifference(before, await (await open(bytes)).workbook()), undefined);
  const changes: [what: string, change: (book: Workbook) => void][] = [
    ['name of a drawing', (book) => (book.sheets[0].drawings![0].object.name = 'Başka')],
    ['given name of a drawing', (book) => (book.sheets[0].drawings![1].object.name = 'Shape 9')],
    ['external picture', (book) => (book.sheets[0].drawings![1].object = { type: 'picture', external: true })],
    ['bold of a run', (book) => ((book.sheets[0].drawings![0].object as DrawingShape).text!.paragraphs[0].runs[0].bold = false)],
    ['size of a run', (book) => ((book.sheets[0].drawings![0].object as DrawingShape).text!.paragraphs[0].runs[0].size = 13)],
    ['colour of a run', (book) => ((book.sheets[0].drawings![0].object as DrawingShape).text!.paragraphs[0].runs[0].color = '#00FF00')],
    ['font of a run', (book) => ((book.sheets[0].drawings![0].object as DrawingShape).text!.paragraphs[0].runs[0].font = 'Tahoma')],
    ['insets', (book) => ((book.sheets[0].drawings![0].object as DrawingShape).text!.insets = [3, 1, 3, 2])],
    ['box of a comment', (book) => (book.sheets[0].comments![1].box!.from!.rowOff = 9525 * 6)],
    ['given box of a comment', (book) => (book.sheets[0].comments![0].box!.to!.row = 9)],
    ['a comment that is a conversation', (book) => (book.sheets[0].comments![0].threaded = true)],
    ['replies', (book) => (book.sheets[0].comments![0].replies = [{ text: 'yanıt' }])],
    ['name of a cell style', (book) => (book.styles.cellStyles[1].name = 'Başka')],
    ['built-in id of a cell style', (book) => (book.styles.cellStyles[0].builtinId = 3)],
    ['format of a cell style', (book) => (book.styles.cellStyleXfs[1].numFmtId = 4)],
    ['a cell style less', (book) => book.styles.cellStyles.pop()],
    ['a differential format', (book) => (book.styles.dxfs[0].children = [])],
    ['default row height', (book) => (book.sheets[0].defaultRowHeight = 15)],
    ['no default row height', (book) => (book.sheets[0].defaultRowHeight = undefined)],
  ];
  for (const [what, change] of changes) {
    const copy = await (await open(bytes)).workbook();
    change(copy);
    assert.notEqual(await workbookDifference(before, copy), undefined, what);
  }
  // What the writer adds is expected as what it adds, not passed over: a sheet without a default height gets 15 points.
  const bare = fullBook();
  bare.sheets[0].defaultRowHeight = undefined;
  bare.styles.cellStyles = [];
  const after = await (await open(await buildXlsx(bare))).workbook();
  assert.equal(await workbookDifference(bare, after), undefined);
  assert.deepEqual([after.sheets[0].defaultRowHeight, after.styles.cellStyles], [15, [{ name: 'Normal', xfId: 0, builtinId: 0 }]]);
  after.sheets[0].defaultRowHeight = 16;
  assert.match((await workbookDifference(bare, after)) || '', /defaultRowHeight/);
  // The two sides are not the same thing: what is expected of a file is not what a file reads as.
  assert.notEqual(canonical(workbookContent(bare)), undefined);
});

test('buildXlsx: what our reader does not read back is in the file all the same', async () => {
  const before = fullBook();
  const report: XlsxBuildReport = { parts: [], warnings: [] };
  const pkg = (await open(await buildXlsx(before, { report }))).package!;
  assert.deepEqual(report.warnings, []);
  // The font of a comment that is one run.
  assert.match(text(await pkg.readPart('/xl/comments1.xml')), /<comment ref="A1" authorId="0"><text><r><rPr><rFont val="Tahoma"\/><b\/><sz val="9"\/><\/rPr><t>tek parça<\/t><\/r><\/text><\/comment>/);
  // The box of a comment, in pixels, and that it is always shown.
  assert.match(text(await pkg.readPart('/xl/drawings/vmlDrawing1.vml')), /visibility:visible[^]*<x:Anchor>2, 3, 1, 4, 4, 0, 6, 5<\/x:Anchor><x:AutoFill>False<\/x:AutoFill><x:Row>2<\/x:Row><x:Column>1<\/x:Column><x:Visible\/>/);
  const drawing = text(await pkg.readPart('/xl/drawings/drawing1.xml'));
  assert.match(drawing, /<a:r><a:rPr lang="en-US" sz="1200" b="1"><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><a:latin typeface="Arial"\/><\/a:rPr><a:t>kalın<\/a:t><\/a:r>/);
  assert.match(drawing, /<a:bodyPr lIns="28575" tIns="9525" rIns="28575" bIns="9525" anchor="t"\/>/);
  assert.match(drawing, /<xdr:cNvPr id="2" name="Kutu"\/>[^]*<xdr:cNvPr id="3" name="Shape 2"\/>/);
  const styles = text(await pkg.readPart('/xl/styles.xml'));
  assert.match(styles, /<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"\/><cellStyle name="Özel &lt;stil>" xfId="1"\/><\/cellStyles>/);
  assert.match(styles, /<dxfs count="1"><dxf><font><b\/><color rgb="FFFF0000"\/><\/font><fill><patternFill><bgColor rgb="FFFFFF00"\/><\/patternFill><\/fill><\/dxf><\/dxfs>/);
});
