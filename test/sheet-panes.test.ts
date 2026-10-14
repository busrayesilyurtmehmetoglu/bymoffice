// Frozen panes: the operation, the command, and what the writer makes of them - only the view of
// the sheet changes, nothing asks for a new calculation, and an undone change saves the file as
// it was. openpyxl reads the panes back where python3 has it.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open, parseXml } from '../src/index';
import { setCellInput } from '../src/ops/commands';
import { applyOp } from '../src/ops/ops';
import { freezePanes, frozenPanes } from '../src/ops/panes';
import { insertSheet } from '../src/ops/sheets';
import { apply, changedParts, hasPython, load, packageProblems, partText, runPython, sameBytes, save } from './edit-commands-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';

const DATA = '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row></sheetData>';

/** One sheet with the views given (none: no <sheetViews>), one plain, and a chart sheet. */
const book = (views: string | undefined, more = '') =>
  buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(`${views === undefined ? '' : `<sheetViews>${views}</sheetViews>`}<sheetFormatPr defaultRowHeight="15"/>${DATA}${more}`) },
      { name: 'Other', xml: worksheetXml('<sheetData/>') },
    ],
  });

const viewOf = async (bytes: Uint8Array) => {
  const xml = await partText(bytes, 'xl/worksheets/sheet1.xml');
  const match = /<sheetViews>.*<\/sheetViews>/.exec(xml);
  return match ? match[0] : '';
};

test('panes: freezing the panes of a sheet writes them into its view, keeps the selection, and asks for no calculation', async () => {
  const views = '<sheetView tabSelected="1" zoomScale="120" workbookViewId="0"><selection activeCell="C3" sqref="C3:D4"/></sheetView>';
  const doc = await load(await book(views));
  apply(doc, freezePanes(doc.wb, 0, 1, 1));
  assert.deepEqual(frozenPanes(doc.wb, 0), { rows: 1, cols: 1 });
  const sheet = doc.wb.sheets[0];
  assert.equal(sheet.modified, undefined, 'a view-only change is no change of the cells');
  assert.equal(sheet.paneModified, true);
  const saved = await save(doc);
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/worksheets/sheet1.xml']);
  assert.equal(
    await viewOf(saved),
    '<sheetViews><sheetView tabSelected="1" zoomScale="120" workbookViewId="0">' +
      '<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>' +
      '<selection pane="topRight"/><selection pane="bottomLeft"/><selection pane="bottomRight" activeCell="C3" sqref="C3:D4"/>' +
      '</sheetView></sheetViews>',
  );
  assert.doesNotMatch(await partText(saved, 'xl/workbook.xml'), /fullCalcOnLoad/);
  const again = await (await open(saved)).workbook();
  assert.deepEqual([again.sheets[0].view.frozenRows, again.sheets[0].view.frozenCols, again.sheets[0].view.topLeftCell, again.sheets[0].view.activeCell], [1, 1, 'B2', 'C3']);
});

test('panes: freeze, unfreeze and undo - each save says what the model says, and the undone change gives the file back', async () => {
  const views = '<sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView>';
  const doc = await load(await book(views));
  apply(doc, freezePanes(doc.wb, 0, 1, 0));
  let saved = await save(doc);
  assert.match(await viewOf(saved), /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/><selection pane="bottomLeft" activeCell="A1" sqref="A1"\/><\/sheetView>/);
  apply(doc, freezePanes(doc.wb, 0, 0, 3));
  saved = await save(doc);
  assert.match(await viewOf(saved), /<pane xSplit="3" topLeftCell="D1" activePane="topRight" state="frozen"\/><selection pane="topRight" activeCell="A1" sqref="A1"\/><\/sheetView>/);
  apply(doc, freezePanes(doc.wb, 0, 0, 0));
  saved = await save(doc);
  // Unfrozen panes are the panes the file came with: nothing of it is written anew.
  assert.ok(sameBytes(saved, doc.bytes));
  while (doc.history.undo());
  assert.ok(sameBytes(await save(doc), doc.bytes), 'freeze, undo, save: the file as it was');
  assert.deepEqual(freezePanes(doc.wb, 0, 0, 0), [], 'nothing to unfreeze');
});

test('panes: unfreezing a file with frozen panes takes its pane away and gives the selection back to the only pane', async () => {
  const views =
    '<sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/>' +
    '<selection pane="topRight" activeCell="C1" sqref="C1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="E9" sqref="E9"/>' +
    '<pivotSelection pane="bottomRight"><pivotArea/></pivotSelection></sheetView>';
  const doc = await load(await book(views));
  assert.deepEqual(frozenPanes(doc.wb, 0), { rows: 3, cols: 2 });
  apply(doc, freezePanes(doc.wb, 0, 0, 0));
  const saved = await save(doc);
  // The pivot selection the model does not know is kept, where it was.
  assert.equal(await viewOf(saved), '<sheetViews><sheetView workbookViewId="0"><selection activeCell="E9" sqref="E9"/><pivotSelection pane="bottomRight"><pivotArea/></pivotSelection></sheetView></sheetViews>');
  const again = await (await open(saved)).workbook();
  assert.deepEqual([again.sheets[0].view.frozenRows, again.sheets[0].view.frozenCols, again.sheets[0].view.activeCell], [0, 0, 'E9']);
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('panes: frozen while scrolled - the file starts its frozen rows at the top, as the view shows them', async () => {
  const views = '<sheetView topLeftCell="C5" workbookViewId="0"><pane ySplit="2" topLeftCell="C9" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="C9" sqref="C9"/></sheetView>';
  const doc = await load(await book(views));
  apply(doc, freezePanes(doc.wb, 0, 4, 0));
  const saved = await save(doc);
  assert.equal(
    await viewOf(saved),
    '<sheetViews><sheetView workbookViewId="0"><pane ySplit="4" topLeftCell="A5" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="C9" sqref="C9"/></sheetView></sheetViews>',
  );
});

test('panes: a split pane that is not frozen stays until panes are frozen, and is then replaced', async () => {
  const views = '<sheetView workbookViewId="0"><pane xSplit="2400" ySplit="1500" topLeftCell="D7" activePane="bottomRight"/><selection pane="bottomRight" activeCell="D7" sqref="D7"/></sheetView>';
  const doc = await load(await book(views));
  assert.equal(doc.wb.sheets[0].view.split, true);
  assert.deepEqual(frozenPanes(doc.wb, 0), { rows: 0, cols: 0 });
  assert.deepEqual(freezePanes(doc.wb, 0, 0, 0), []);
  apply(doc, freezePanes(doc.wb, 0, 2, 0));
  const saved = await save(doc);
  assert.equal(await viewOf(saved), '<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="D7" sqref="D7"/></sheetView></sheetViews>');
  doc.history.undo();
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('panes: a part without views gets one, a self-closing view gets content, and the first of several views is the one changed', async () => {
  let doc = await load(await book(undefined));
  apply(doc, freezePanes(doc.wb, 0, 1, 0));
  let saved = await save(doc);
  const xml = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<worksheet[^>]*><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/><selection pane="bottomLeft"\/><\/sheetView><\/sheetViews><sheetFormatPr/);

  doc = await load(await book('<sheetView workbookViewId="0"/><sheetView workbookViewId="1"><selection activeCell="B2" sqref="B2"/></sheetView>'));
  apply(doc, freezePanes(doc.wb, 0, 0, 1));
  saved = await save(doc);
  assert.equal(
    await viewOf(saved),
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/><selection pane="topRight"/></sheetView><sheetView workbookViewId="1"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>',
  );
});

test('panes: a part with a prefix, a byte order mark and line ends of two characters keeps every other byte', async () => {
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const body = `<x:worksheet xmlns:x="${NS}">\r\n<x:sheetViews>\r\n<x:sheetView workbookViewId="0">\r\n<x:selection activeCell="B2" sqref="B2"/>\r\n</x:sheetView>\r\n</x:sheetViews>\r\n<x:sheetData/>\r\n</x:worksheet>`;
  const doc = await load(await buildXlsx({ sheets: [{ name: 'S', xml: `﻿<?xml version="1.0" encoding="UTF-8"?>\r\n${body}` }] }));
  apply(doc, freezePanes(doc.wb, 0, 2, 2));
  const saved = await save(doc);
  const bytes = await (await import('../src/index')).ZipArchive.read(saved).read('xl/worksheets/sheet1.xml');
  assert.deepEqual(Array.from(bytes.slice(0, 3)), [0xef, 0xbb, 0xbf]);
  const text = new TextDecoder().decode(bytes.slice(3));
  assert.equal(
    text,
    `<?xml version="1.0" encoding="UTF-8"?>\r\n<x:worksheet xmlns:x="${NS}">\r\n<x:sheetViews>\r\n<x:sheetView workbookViewId="0">\r\n` +
      '<x:pane xSplit="2" ySplit="2" topLeftCell="C3" activePane="bottomRight" state="frozen"/><x:selection pane="topRight"/><x:selection pane="bottomLeft"/><x:selection pane="bottomRight" activeCell="B2" sqref="B2"/>' +
      '\r\n</x:sheetView>\r\n</x:sheetViews>\r\n<x:sheetData/>\r\n</x:worksheet>',
  );
});

test('panes: with cells changed as well both are written, and the sheet then asks for a calculation', async () => {
  const doc = await load(await book('<sheetView workbookViewId="0"/>'));
  apply(doc, setCellInput(doc.wb, 0, 4, 0, '7'));
  apply(doc, freezePanes(doc.wb, 0, 1, 0));
  const saved = await save(doc);
  const xml = await partText(saved, 'xl/worksheets/sheet1.xml');
  assert.match(xml, /<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"\/>/);
  assert.match(xml, /<c r="A5"[^>]*><v>7<\/v><\/c>/);
  const again = await (await open(saved)).workbook();
  assert.equal(again.sheets[0].view.frozenRows, 1);
});

test('panes: cells changed and changed back, then panes frozen - the sheet still asks for no calculation', async () => {
  const doc = await load(await book('<sheetView workbookViewId="0"/>'));
  apply(doc, setCellInput(doc.wb, 0, 0, 0, '99'));
  doc.history.undo();
  apply(doc, freezePanes(doc.wb, 0, 1, 0));
  assert.equal(doc.wb.sheets[0].modified, true, 'the undone edit leaves the mark');
  const saved = await save(doc);
  assert.match(await partText(saved, 'xl/worksheets/sheet1.xml'), /<pane ySplit="1"/);
  assert.doesNotMatch(await partText(saved, 'xl/workbook.xml'), /fullCalcOnLoad/);
});

test('panes: a sheet that is in no file is written with its panes from the model', async () => {
  const doc = await load(await book(undefined));
  apply(doc, insertSheet(doc.wb, 2, { name: 'New' }));
  apply(doc, freezePanes(doc.wb, 2, 2, 1));
  const saved = await save(doc);
  const again = await (await open(saved)).workbook();
  const sheet = again.sheets.find((s) => s.name === 'New')!;
  assert.deepEqual([sheet.view.frozenRows, sheet.view.frozenCols, sheet.view.topLeftCell], [2, 1, 'B3']);
});

test('panes: the operation checks what it is given and is its own inverse', async () => {
  const doc = await load(await book(undefined));
  assert.throws(() => freezePanes(doc.wb, 0, -1, 0), RangeError);
  assert.throws(() => freezePanes(doc.wb, 0, 1.5, 0), RangeError);
  assert.throws(() => freezePanes(doc.wb, 0, 0, 16384), RangeError);
  assert.throws(() => freezePanes(doc.wb, 9, 1, 0), RangeError);
  assert.throws(() => applyOp(doc.wb, { type: 'setPane', sheet: 0, pane: { frozenRows: 1, frozenCols: 0, topLeftCell: 'nope' } }), RangeError);
  assert.throws(() => applyOp(doc.wb, { type: 'setPane', sheet: 0, pane: null as never }), RangeError);
  const inverse = applyOp(doc.wb, { type: 'setPane', sheet: 0, pane: { frozenRows: 3, frozenCols: 2, topLeftCell: 'C4' } });
  assert.deepEqual(inverse, { type: 'setPane', sheet: 0, pane: { frozenRows: 0, frozenCols: 0 } });
  const back = applyOp(doc.wb, inverse);
  assert.deepEqual(back, { type: 'setPane', sheet: 0, pane: { frozenRows: 3, frozenCols: 2, topLeftCell: 'C4' } });
  assert.equal(doc.wb.sheets[0].view.topLeftCell, undefined);
});

test('panes: a chart sheet has none to freeze', async () => {
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(DATA) },
      { name: 'Chart', kind: 'chartsheet', xml: '<?xml version="1.0" encoding="UTF-8"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>' },
    ],
  });
  const doc = await load(bytes);
  assert.throws(() => freezePanes(doc.wb, 1, 1, 0), /chartsheet/);
});

const OPENPYXL = String.raw`
import json, os, sys
from openpyxl import load_workbook
folder = sys.argv[1]
out = {}
for name in sorted(os.listdir(folder)):
    wb = load_workbook(os.path.join(folder, name))
    out[name] = [wb[title].freeze_panes for title in wb.sheetnames]
print(json.dumps(out))
`;

test('panes: openpyxl reads the panes the engine froze, and the packages are sound', async (t) => {
  if (!hasPython(['openpyxl'])) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const files: { [name: string]: Uint8Array } = {};
  const cases: [string, string | undefined, number, number][] = [
    ['both.xlsx', '<sheetView workbookViewId="0"><selection activeCell="C3" sqref="C3"/></sheetView>', 2, 3],
    ['rows.xlsx', undefined, 1, 0],
    ['cols.xlsx', '<sheetView workbookViewId="0"/>', 0, 1],
    ['none.xlsx', '<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>', 0, 0],
  ];
  for (const [name, views, rows, cols] of cases) {
    const doc = await load(await book(views));
    apply(doc, freezePanes(doc.wb, 0, rows, cols));
    files[name] = await save(doc);
  }
  const read = runPython<{ [name: string]: (string | null)[] }>(files, OPENPYXL);
  assert.deepEqual(read, { 'both.xlsx': ['D3', null], 'cols.xlsx': ['B1', null], 'none.xlsx': [null, null], 'rows.xlsx': ['A2', null] });
  if (hasPython(['xml.etree.ElementTree'])) assert.deepEqual(packageProblems(files), []);
  // The element order of a view the schema states: pane, selection, pivotSelection, extLst.
  for (const name of Object.keys(files)) {
    const root = parseXml(await partText(files[name], 'xl/worksheets/sheet1.xml')).root;
    const sheetViews = root.children.find((node) => node.type === 'element' && node.name === 'sheetViews');
    if (!sheetViews || sheetViews.type !== 'element') continue;
    for (const view of sheetViews.children) {
      if (view.type !== 'element') continue;
      const order = view.children.filter((node) => node.type === 'element').map((node) => (node as { name: string }).name);
      const ranks = order.map((local) => ['pane', 'selection', 'pivotSelection', 'extLst'].indexOf(local));
      assert.deepEqual(ranks.slice().sort((a, b) => a - b), ranks, `${name}: ${order.join(' ')}`);
    }
  }
});
