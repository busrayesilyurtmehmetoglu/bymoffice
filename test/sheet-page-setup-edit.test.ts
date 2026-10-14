// Changing how a worksheet is printed: the commands (paper, orientation, margins, scale and "fit
// to", what is printed and where, print area and titles, manual page breaks), their operations
// and inverses, what the writer makes of them (only the elements that changed, nothing asks for
// a new calculation, an undone change saves the file as it was), and a reading of the result by
// openpyxl and package_check.py where python3 has them. Every file is built here.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { open } from '../src/index';
import { paginateSheet, printRanges } from '../src/layout/sheet-pages';
import { planSheetPages } from '../src/ui/sheet-print';
import { applyOp, type Op } from '../src/ops/ops';

const insertRows = (_wb: unknown, sheet: number, at: number, count: number): Op[] => [{ type: 'insert', sheet, axis: 'row', at, count }];
const deleteRows = (_wb: unknown, sheet: number, at: number, count: number): Op[] => [{ type: 'delete', sheet, axis: 'row', at, count }];
import { insertPageBreak, pageSetupState, printAreaValue, printTitlesValue, removePageBreak, resetPageBreaks, setPageSetup } from '../src/ops/page-setup';
import { copySheet, copySheetLosses, renameSheet } from '../src/ops/sheets';
import { apply, changedParts, hasPython, load, packageProblems, partText, runPython, sameBytes, save } from './edit-commands-fixtures';
import { buildXlsx, worksheetXml } from './fixtures';

const DATA = '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c></row><row r="30"><c r="C30"><v>4</v></c></row></sheetData>';

/** One sheet with what follows its cells (`after`), a second plain one; `names` go into the workbook. */
const book = (after = '', names = '', name = 'Data', before = '') =>
  buildXlsx({
    sheets: [
      { name, xml: worksheetXml(`${before}<sheetFormatPr defaultRowHeight="15"/>${DATA}${after}`) },
      { name: 'Other', xml: worksheetXml('<sheetData/>') },
    ],
    workbookTail: names === '' ? '' : `<definedNames>${names}</definedNames>`,
  });

const sheetPart = (bytes: Uint8Array) => partText(bytes, 'xl/worksheets/sheet1.xml');
const elementOf = (xml: string, local: string): string => {
  const match = new RegExp(`<${local}[\\s/>][^]*?(?:/>|</${local}>)`).exec(xml);
  return match ? match[0] : '';
};

test('page setup: paper, orientation and margins change only their attributes; nothing else of the part, no calculation', async () => {
  const after = '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" horizontalDpi="300" r:id="rId9"/>';
  const doc = await load(await book(after));
  const ops = setPageSetup(doc.wb, 0, { paperSize: 1, orientation: 'landscape', margins: { left: 1, top: 0.5 } });
  assert.deepEqual(
    ops.map((op) => op.type),
    ['setPageSetup'],
  );
  apply(doc, ops);
  const sheet = doc.wb.sheets[0];
  assert.equal(sheet.modified, undefined, 'a change of the page is no change of the cells');
  assert.equal(sheet.pageSetupModified, true);
  const state = pageSetupState(doc.wb, 0);
  assert.equal(state.paperSize, 1);
  assert.equal(state.orientation, 'landscape');
  assert.deepEqual(state.margins, { left: 1, right: 0.7, top: 0.5, bottom: 0.75, header: 0.3, footer: 0.3 });
  const saved = await save(doc);
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/worksheets/sheet1.xml']);
  const xml = await sheetPart(saved);
  assert.equal(elementOf(xml, 'pageMargins'), '<pageMargins left="1" right="0.7" top="0.5" bottom="0.75" header="0.3" footer="0.3"/>');
  // The printer settings and the resolution stay as the file has them.
  assert.equal(elementOf(xml, 'pageSetup'), '<pageSetup paperSize="1" orientation="landscape" horizontalDpi="300" r:id="rId9"/>');
  assert.equal(xml.replace(/<pageMargins[^>]*\/><pageSetup[^>]*\/>/, ''), (await sheetPart(doc.bytes)).replace(/<pageMargins[^>]*\/><pageSetup[^>]*\/>/, ''));
  assert.doesNotMatch(await partText(saved, 'xl/workbook.xml'), /fullCalcOnLoad/);
  const again = await (await open(saved)).workbook();
  assert.deepEqual(again.sheets[0].pageSetup?.margins, { left: 1, right: 0.7, top: 0.5, bottom: 0.75, header: 0.3, footer: 0.3 });
  assert.equal(again.sheets[0].pageSetup?.page?.paperSize, 1);
  // Undo: the file as it was.
  while (doc.history.undo());
  assert.equal(doc.wb.sheets[0].pageSetup?.page?.paperSize, 9);
  assert.ok(sameBytes(await save(doc), doc.bytes), 'change, undo, save: the bytes of the file');
});

test('page setup: elements the sheet does not have are put where the schema has them', async () => {
  const doc = await load(await book('<drawing r:id="rId3"/>'));
  apply(doc, setPageSetup(doc.wb, 0, { orientation: 'landscape', gridLines: true, horizontalCentered: true, fitTo: { width: 1, height: 0 } }));
  const xml = await sheetPart(await save(doc));
  assert.match(xml, /<sheetPr><pageSetUpPr fitToPage="1"\/><\/sheetPr><sheetFormatPr/);
  assert.match(xml, /<\/sheetData><printOptions horizontalCentered="1" gridLines="1"\/><pageSetup fitToHeight="0" orientation="landscape"\/><drawing r:id="rId3"\/>/);
  const again = await (await open(await save(doc))).workbook();
  const setup = again.sheets[0].pageSetup!;
  assert.equal(setup.fitToPage, true);
  assert.equal(setup.page?.fitToHeight, 0);
  assert.equal(setup.print?.gridLines, true);
});

test('page setup: a scale ends "fit to pages", which a <sheetPr> of the file had among its other children', async () => {
  const before = '<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/><pageSetUpPr fitToPage="1" autoPageBreaks="0"/></sheetPr>';
  const doc = await load(await book('<pageSetup scale="50" fitToWidth="2"/>', '', 'Data', before));
  assert.equal(pageSetupState(doc.wb, 0).fitToPage, true);
  apply(doc, setPageSetup(doc.wb, 0, { scale: 80 }));
  const state = pageSetupState(doc.wb, 0);
  assert.equal(state.fitToPage, false);
  assert.equal(state.scale, 80);
  const xml = await sheetPart(await save(doc));
  assert.match(xml, /<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"\/><pageSetUpPr autoPageBreaks="0"\/><\/sheetPr>/);
  assert.match(xml, /<pageSetup scale="80" fitToWidth="2"\/>/);
  assert.throws(() => setPageSetup(doc.wb, 0, { scale: 80, fitTo: { width: 1, height: 1 } }), /either at a scale/);
  assert.throws(() => setPageSetup(doc.wb, 0, { scale: 401 }), /scale/);
  assert.throws(() => setPageSetup(doc.wb, 0, { paperSize: 4242 }), /paper/);
  assert.throws(() => setPageSetup(doc.wb, 0, { margins: { left: 5, right: 5 } }), /no room/);
  assert.throws(() => setPageSetup(doc.wb, 0, { margins: { left: -1 } }), /margin/);
});

test('page setup: a patch that changes nothing makes no operation', async () => {
  const doc = await load(await book('<printOptions gridLines="1"/><pageSetup paperSize="9" orientation="landscape"/>'));
  assert.deepEqual(setPageSetup(doc.wb, 0, { paperSize: 9, orientation: 'landscape', gridLines: true, headings: false }), []);
  assert.deepEqual(setPageSetup(doc.wb, 0, { printArea: null, titleRows: null }), []);
});

test('print area and titles: defined names local to the sheet, absolute, with the name of the sheet quoted; the workbook part follows and asks for no calculation', async () => {
  const doc = await load(await book('', '', "Q1 'Data'"));
  const ops = setPageSetup(doc.wb, 0, { printArea: [{ top: 0, left: 0, bottom: 19, right: 3 }], titleRows: { first: 0, last: 1 }, titleCols: { first: 0, last: 0 } });
  assert.deepEqual(
    ops.map((op) => op.type),
    ['setDefinedNames'],
  );
  apply(doc, ops);
  const names = doc.wb.definedNames.map((name) => [name.name, name.value, name.localSheet]);
  assert.deepEqual(names, [
    ['_xlnm.Print_Area', "'Q1 ''Data'''!$A$1:$D$20", 0],
    ['_xlnm.Print_Titles', "'Q1 ''Data'''!$A:$A,'Q1 ''Data'''!$1:$2", 0],
  ]);
  assert.deepEqual(printRanges(doc.wb, 0), { areas: [{ top: 0, left: 0, bottom: 19, right: 3 }], titleRows: { first: 0, last: 1 }, titleCols: { first: 0, last: 0 } });
  const saved = await save(doc);
  assert.deepEqual(await changedParts(doc.bytes, saved), ['xl/workbook.xml']);
  const workbookXml = await partText(saved, 'xl/workbook.xml');
  assert.match(workbookXml, /<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">'Q1 ''Data'''!\$A\$1:\$D\$20<\/definedName><definedName name="_xlnm.Print_Titles" localSheetId="0">/);
  assert.doesNotMatch(workbookXml, /fullCalcOnLoad/);
  // Cleared again: the names go, and the file is what it was.
  apply(doc, setPageSetup(doc.wb, 0, { printArea: null, titleRows: null, titleCols: null }));
  assert.deepEqual(doc.wb.definedNames, []);
  assert.ok(sameBytes(await save(doc), doc.bytes));
  while (doc.history.undo());
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('print area: a name of the file is changed in place and removed with its element; whole rows and columns are written short', async () => {
  const names = '<definedName name="Total">Data!$B$2</definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!$A$1:$B$2</definedName>';
  const doc = await load(await book('', names));
  apply(doc, setPageSetup(doc.wb, 0, { printArea: [{ top: 2, left: 0, bottom: 4, right: 16383 }, { top: 0, left: 1, bottom: 1048575, right: 2 }] }));
  assert.equal(doc.wb.definedNames[1].value, 'Data!$3:$5,Data!$B:$C');
  let xml = await partText(await save(doc), 'xl/workbook.xml');
  assert.match(xml, /<definedName name="Total">Data!\$B\$2<\/definedName><definedName name="_xlnm.Print_Area" localSheetId="0">Data!\$3:\$5,Data!\$B:\$C<\/definedName>/);
  apply(doc, setPageSetup(doc.wb, 0, { printArea: null }));
  xml = await partText(await save(doc), 'xl/workbook.xml');
  assert.match(xml, /<definedNames><definedName name="Total">Data!\$B\$2<\/definedName><\/definedNames>/);
  assert.doesNotMatch(xml, /fullCalcOnLoad/);
});

test('print area and titles follow inserted and deleted rows and a renamed sheet', async () => {
  const doc = await load(await book());
  apply(doc, setPageSetup(doc.wb, 0, { printArea: [{ top: 1, left: 0, bottom: 9, right: 2 }], titleRows: { first: 0, last: 0 } }));
  apply(doc, insertRows(doc.wb, 0, 0, 2));
  assert.deepEqual(printRanges(doc.wb, 0), { areas: [{ top: 3, left: 0, bottom: 11, right: 2 }], titleRows: { first: 2, last: 2 } });
  apply(doc, deleteRows(doc.wb, 0, 5, 1));
  assert.deepEqual(printRanges(doc.wb, 0).areas, [{ top: 3, left: 0, bottom: 10, right: 2 }]);
  apply(doc, renameSheet(doc.wb, 0, 'Sales 2026'));
  assert.equal(doc.wb.definedNames[0].value, "'Sales 2026'!$A$4:$C$11");
  const again = await (await open(await save(doc))).workbook();
  assert.deepEqual(printRanges(again, 0), { areas: [{ top: 3, left: 0, bottom: 10, right: 2 }], titleRows: { first: 2, last: 2 } });
});

test('page breaks: inserted at a cell, removed, reset; written as the breaks of the model, with the file\'s other breaks kept', async () => {
  const breaks = '<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks>';
  const doc = await load(await book(`<pageSetup orientation="portrait"/>${breaks}`));
  assert.deepEqual(pageSetupState(doc.wb, 0).rowBreaks, [10]);
  apply(doc, insertPageBreak(doc.wb, 0, { row: 4, col: 2 }));
  let state = pageSetupState(doc.wb, 0);
  assert.deepEqual([state.rowBreaks, state.colBreaks], [[4, 10], [2]]);
  assert.deepEqual(insertPageBreak(doc.wb, 0, { row: 4 }), [], 'a break that is there already');
  assert.deepEqual(insertPageBreak(doc.wb, 0, { row: 0, col: 0 }), [], 'no break above the first row');
  let xml = await sheetPart(await save(doc));
  assert.match(xml, /<pageSetup orientation="portrait"\/><rowBreaks count="2" manualBreakCount="2"><brk id="4" max="16383" man="1"\/><brk id="10" max="16383" man="1"\/><\/rowBreaks><colBreaks count="1" manualBreakCount="1"><brk id="2" max="1048575" man="1"\/><\/colBreaks>/);
  // The layout breaks the pages there.
  const pages = paginateSheet(doc.wb.sheets[0]);
  assert.deepEqual(
    pages.pages.map((page) => [page.rows.first, page.cols.first]),
    [
      [0, 0],
      [4, 0],
      [10, 0],
      [0, 2],
      [4, 2],
      [10, 2],
    ],
  );
  apply(doc, removePageBreak(doc.wb, 0, { row: 10, col: 2 }));
  state = pageSetupState(doc.wb, 0);
  assert.deepEqual([state.rowBreaks, state.colBreaks], [[4], []]);
  xml = await sheetPart(await save(doc));
  assert.doesNotMatch(xml, /colBreaks/);
  apply(doc, resetPageBreaks(doc.wb, 0));
  assert.deepEqual(pageSetupState(doc.wb, 0).rowBreaks, []);
  xml = await sheetPart(await save(doc));
  assert.doesNotMatch(xml, /rowBreaks/);
  while (doc.history.undo());
  assert.ok(sameBytes(await save(doc), doc.bytes));
});

test('page breaks set after rows were inserted move only with the rows inserted after them', async () => {
  const breaks = '<rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks>';
  const doc = await load(await book(breaks));
  apply(doc, insertRows(doc.wb, 0, 0, 3));
  assert.deepEqual(pageSetupState(doc.wb, 0).rowBreaks, [13]);
  apply(doc, insertPageBreak(doc.wb, 0, { row: 20 }));
  assert.equal(doc.wb.sheets[0].pageSetup?.breaksSince, 1);
  assert.deepEqual(pageSetupState(doc.wb, 0).rowBreaks, [13, 20]);
  apply(doc, insertRows(doc.wb, 0, 0, 2));
  assert.deepEqual(pageSetupState(doc.wb, 0).rowBreaks, [15, 22]);
  const xml = await sheetPart(await save(doc));
  const ids = Array.from(xml.matchAll(/<brk id="(\d+)"/g)).map((match) => Number(match[1]));
  // The writer places a break after the last row of its page: 15 and 22 as the reader takes them back.
  assert.deepEqual(ids, [15, 22]);
  const again = await (await open(await save(doc))).workbook();
  assert.deepEqual(pageSetupState(again, 0).rowBreaks, [15, 22]);
  // A copy of the sheet takes the breaks as they are now when nothing moved them since; here rows did.
  const copy = copySheet(doc.wb, 0);
  assert.ok(copy.length > 0);
});

test('page setup: a copy of a sheet whose breaks were set after the last insertion keeps them', async () => {
  const doc = await load(await book());
  apply(doc, insertRows(doc.wb, 0, 0, 1));
  apply(doc, insertPageBreak(doc.wb, 0, { row: 6 }));
  apply(doc, copySheet(doc.wb, 0));
  const copy = doc.wb.sheets[1];
  assert.deepEqual(pageSetupState(doc.wb, 1).rowBreaks, [6]);
  assert.equal(copy.pageSetup?.breaksSince, undefined);
});

test('page setup: refused for a chart sheet; the operation is its own inverse', async () => {
  const bytes = await buildXlsx({
    sheets: [
      { name: 'Data', xml: worksheetXml(DATA) },
      { name: 'Chart', kind: 'chartsheet', xml: '<?xml version="1.0" encoding="UTF-8"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>' },
    ],
  });
  const doc = await load(bytes);
  assert.throws(() => setPageSetup(doc.wb, 1, { orientation: 'landscape' }), /worksheet/);
  assert.throws(() => applyOp(doc.wb, { type: 'setPageSetup', sheet: 1, setup: {} }), /worksheet/);
  const before = doc.wb.sheets[0].pageSetup;
  const inverse = applyOp(doc.wb, { type: 'setPageSetup', sheet: 0, setup: { page: { orientation: 'landscape' } } });
  assert.deepEqual(inverse, { type: 'setPageSetup', sheet: 0, setup: before });
  applyOp(doc.wb, inverse);
  assert.equal(doc.wb.sheets[0].pageSetup, before);
  assert.throws(() => applyOp(doc.wb, { type: 'setDefinedNames', sheet: 0, names: [{ name: '_xlnm.Print_Area', value: 'Data!$A$1', localSheet: 7, hidden: false }] }), /local to no sheet/);
});

test('page setup: the plan of a print job (and so of a PDF) follows the change at once, and undo', async () => {
  const doc = await load(await book('<pageSetup paperSize="9" orientation="portrait"/>'));
  const plan = () => planSheetPages(doc.wb, [0]);
  const first = plan().pages[0];
  assert.equal(first.layout.landscape, false);
  assert.deepEqual([first.page.rows, first.page.cols], [{ first: 0, last: 29 }, { first: 0, last: 2 }]);
  apply(doc, setPageSetup(doc.wb, 0, { orientation: 'landscape', printArea: [{ top: 0, left: 0, bottom: 1, right: 1 }], gridLines: true }));
  const now = plan().pages[0];
  assert.equal(now.layout.landscape, true);
  assert.equal(now.layout.gridLines, true);
  assert.deepEqual([now.page.rows, now.page.cols], [{ first: 0, last: 1 }, { first: 0, last: 1 }]);
  doc.history.undo();
  assert.equal(plan().pages[0].layout.landscape, false);
});

test('page setup: values of the defined names', () => {
  assert.equal(printAreaValue('Sheet1', [{ top: 0, left: 0, bottom: 0, right: 0 }]), 'Sheet1!$A$1');
  assert.equal(printAreaValue('My Sheet', [{ top: 0, left: 0, bottom: 9, right: 1 }]), "'My Sheet'!$A$1:$B$10");
  assert.equal(printTitlesValue('S', { first: 0, last: 2 }, undefined), 'S!$1:$3');
  assert.equal(printTitlesValue('S', undefined, { first: 1, last: 1 }), 'S!$B:$B');
});

const OPENPYXL = String.raw`
import json, os, sys
from openpyxl import load_workbook
folder = sys.argv[1]
out = {}
for name in sorted(os.listdir(folder)):
    wb = load_workbook(os.path.join(folder, name))
    ws = wb[wb.sheetnames[0]]
    ps = ws.page_setup
    m = ws.page_margins
    out[name] = {
        'paperSize': None if ps.paperSize is None else int(ps.paperSize),
        'orientation': ps.orientation,
        'scale': None if ps.scale is None else int(ps.scale),
        'fitToPage': bool(ws.sheet_properties.pageSetUpPr.fitToPage) if ws.sheet_properties.pageSetUpPr is not None else False,
        'fitToHeight': None if ps.fitToHeight is None else int(ps.fitToHeight),
        'margins': [m.left, m.right, m.top, m.bottom, m.header, m.footer],
        'gridLines': bool(ws.print_options.gridLines),
        'horizontalCentered': bool(ws.print_options.horizontalCentered),
        'printArea': ws.print_area,
        'titleRows': ws.print_title_rows,
        'titleCols': ws.print_title_cols,
        'rowBreaks': [b.id for b in ws.row_breaks.brk],
        'colBreaks': [b.id for b in ws.col_breaks.brk],
    }
print(json.dumps(out))
`;

test('page setup: openpyxl reads what the engine set, and the packages are sound', async (t) => {
  if (!hasPython(['openpyxl'])) {
    t.skip('python3 with openpyxl is not available');
    return;
  }
  const files: { [name: string]: Uint8Array } = {};
  {
    const doc = await load(await book('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait"/>'));
    apply(doc, setPageSetup(doc.wb, 0, { paperSize: 8, orientation: 'landscape', margins: { left: 0.25, right: 0.25 }, gridLines: true, horizontalCentered: true }));
    apply(doc, setPageSetup(doc.wb, 0, { printArea: [{ top: 0, left: 0, bottom: 29, right: 2 }], titleRows: { first: 0, last: 0 }, titleCols: { first: 0, last: 1 } }));
    apply(doc, insertPageBreak(doc.wb, 0, { row: 15, col: 1 }));
    files['set.xlsx'] = await save(doc);
  }
  {
    const doc = await load(await book());
    apply(doc, setPageSetup(doc.wb, 0, { fitTo: { width: 1, height: 0 }, margins: { top: 1 } }));
    apply(doc, setPageSetup(doc.wb, 0, { printArea: [{ top: 0, left: 0, bottom: 1, right: 1 }] }));
    files['fit.xlsx'] = await save(doc);
  }
  const read = runPython<{ [name: string]: { [key: string]: unknown } }>(files, OPENPYXL);
  assert.deepEqual(read['set.xlsx'], {
    paperSize: 8,
    orientation: 'landscape',
    scale: null,
    fitToPage: false,
    fitToHeight: null,
    margins: [0.25, 0.25, 0.75, 0.75, 0.3, 0.3],
    gridLines: true,
    horizontalCentered: true,
    printArea: "'Data'!$A$1:$C$30",
    titleRows: '$1:$1',
    titleCols: '$A:$B',
    rowBreaks: [15],
    colBreaks: [1],
  });
  assert.equal(read['fit.xlsx'].fitToPage, true);
  assert.equal(read['fit.xlsx'].fitToHeight, 0);
  assert.deepEqual(read['fit.xlsx'].margins, [0.7, 0.7, 1, 0.75, 0.3, 0.3]);
  assert.equal(read['fit.xlsx'].printArea, "'Data'!$A$1:$B$2");
  if (hasPython(['xml.etree.ElementTree'])) assert.deepEqual(packageProblems(files), []);
});

// ---------------------------------------------------------------------------
// One rule for where a page break goes when rows move (see page-breaks.ts): the model, the pages,
// the file written and the file read back agree.

const FORTY = '<sheetData>' + Array.from({ length: 40 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c></row>`).join('') + '</sheetData>';
const withBreak = () =>
  buildXlsx({
    sheets: [
      { name: 'S', xml: worksheetXml(`${FORTY}<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><rowBreaks count="1" manualBreakCount="1"><brk id="10" max="16383" man="1"/></rowBreaks>`) },
      { name: 'T', xml: worksheetXml('<sheetData/>') },
    ],
  });
const rowsOp = (kind: 'insert' | 'delete', at: number, count = 1): Op[] => [{ type: kind, sheet: 0, axis: 'row', at, count }];

/** The breaks of the first sheet: in the model, in the file written from it, and in that file read back. */
async function breaksEverywhere(doc: Awaited<ReturnType<typeof load>>): Promise<{ model: number[]; file: number[]; reread: number[] }> {
  const model = pageSetupState(doc.wb, 0).rowBreaks;
  const out = await save(doc);
  const xml = await sheetPart(out);
  const file = Array.from(xml.matchAll(/<brk id="(\d+)"/g)).map((match) => Number(match[1]));
  const reread = pageSetupState(await (await open(out)).workbook(), 0).rowBreaks;
  return { model, file, reread };
}

test('page breaks: a break belongs to the row above it, in the model, the file and the layout alike', async () => {
  // Rows inserted right under the break go onto the next page; rows inserted above it, or at the row it follows, push it down.
  // A deleted row the break follows takes the break to the row before the deleted ones (UNSURE).
  const cases: [string, Op[], number[], number][] = [
    ['insert under', rowsOp('insert', 10), [10], 31],
    ['insert at its row', rowsOp('insert', 9), [11], 31],
    ['insert further down', rowsOp('insert', 11), [10], 31],
    ['delete under', rowsOp('delete', 10), [10], 29],
    ['delete its row', rowsOp('delete', 9), [9], 29],
    ['delete above', rowsOp('delete', 0, 3), [7], 27],
    ['delete through', rowsOp('delete', 8, 2), [8], 28],
    ['delete its row and the next', rowsOp('delete', 9, 2), [9], 28],
    ['delete from the top through it', rowsOp('delete', 0, 12), [], 18],
  ];
  for (const [label, ops, expected, later] of cases) {
    for (const set of [false, true]) {
      const doc = await load(await withBreak());
      // Once as the file has it, once with a break set by a command at row 30 (the setup is another object then).
      if (set) apply(doc, insertPageBreak(doc.wb, 0, { row: 30 }));
      apply(doc, ops);
      const want = set ? expected.concat([later]) : expected;
      const seen = await breaksEverywhere(doc);
      assert.deepEqual(seen, { model: want, file: want, reread: want }, `${label}${set ? ' (set)' : ''}`);
      // The pages break there too.
      const starts = paginateSheet(doc.wb.sheets[0]).pages.map((page) => page.rows.first);
      for (const at of want) assert.ok(starts.indexOf(at) >= 0, `${label}: a page starts at ${at}`);
    }
  }
});

test('page breaks: random inserts, deletes and undos keep the model, the file and the file read back the same', async () => {
  let seed = 7;
  const random = (n: number) => {
    seed = (seed * 1103515245 + 12345) % 2147483648;
    return seed % n;
  };
  for (let round = 0; round < 6; round++) {
    const doc = await load(await withBreak());
    for (let step = 0; step < 12; step++) {
      const choice = random(5);
      if (choice === 0) apply(doc, rowsOp('insert', random(30), 1 + random(3)));
      else if (choice === 1) apply(doc, rowsOp('delete', random(30), 1 + random(3)));
      else if (choice === 2) apply(doc, insertPageBreak(doc.wb, 0, { row: 1 + random(35) }));
      else if (choice === 3) doc.history.undo();
      else {
        const now = pageSetupState(doc.wb, 0).rowBreaks;
        if (now.length > 0) apply(doc, removePageBreak(doc.wb, 0, { row: now[random(now.length)] }));
      }
      const seen = await breaksEverywhere(doc);
      assert.deepEqual(seen.file, seen.model, `round ${round} step ${step}: file`);
      assert.deepEqual(seen.reread, seen.model, `round ${round} step ${step}: read back`);
    }
    while (doc.history.undo());
    assert.ok(sameBytes(await save(doc), doc.bytes), `round ${round}: everything undone gives the file back`);
  }
});

test('page breaks: the columns a row break spans follow inserted columns, also once breaks are set anew', async () => {
  const doc = await load(
    await buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml('<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><rowBreaks count="1" manualBreakCount="1"><brk id="10" min="2" max="5" man="1"/></rowBreaks>') }] }),
  );
  apply(doc, [{ type: 'insert', sheet: 0, axis: 'col', at: 0, count: 1 }]);
  assert.match(await sheetPart(await save(doc)), /<brk id="10" min="3" max="6" man="1"\/>/);
  apply(doc, insertPageBreak(doc.wb, 0, { row: 20 }));
  assert.deepEqual(
    doc.wb.sheets[0].pageSetup!.rowBreaks!.map((entry) => [entry.at, entry.min, entry.max]),
    [
      [10, 3, 6],
      [20, 0, 16383],
    ],
  );
  assert.match(await sheetPart(await save(doc)), /<brk id="10" min="3" max="6" man="1"\/><brk id="20" max="16383" man="1"\/>/);
});

test('page breaks: a copy of a sheet takes them where they are now, after rows were inserted', async () => {
  const doc = await load(await withBreak());
  apply(doc, insertPageBreak(doc.wb, 0, { row: 20 }));
  apply(doc, rowsOp('insert', 0, 1));
  assert.deepEqual(copySheetLosses(doc.wb, 0), []);
  apply(doc, copySheet(doc.wb, 0));
  assert.deepEqual(pageSetupState(doc.wb, 1).rowBreaks, [11, 21]);
  const again = await (await open(await save(doc))).workbook();
  assert.deepEqual(pageSetupState(again, 1).rowBreaks, [11, 21]);
});

test('print area of a second sheet: local to it in the file (localSheetId 1)', async () => {
  const doc = await load(await book());
  apply(doc, setPageSetup(doc.wb, 1, { printArea: [{ top: 0, left: 0, bottom: 2, right: 2 }] }));
  const xml = await partText(await save(doc), 'xl/workbook.xml');
  assert.match(xml, /<definedName name="_xlnm.Print_Area" localSheetId="1">Other!\$A\$1:\$C\$3<\/definedName>/);
  const again = await (await open(await save(doc))).workbook();
  assert.deepEqual(printRanges(again, 1).areas, [{ top: 0, left: 0, bottom: 2, right: 2 }]);
  assert.deepEqual(printRanges(again, 0).areas, []);
});

test('only a print area set, and another sheet shown: the file opens on the sheet that was shown', async () => {
  const doc = await load(await book());
  doc.wb.activeSheet = 1;
  apply(doc, setPageSetup(doc.wb, 0, { printArea: [{ top: 0, left: 0, bottom: 1, right: 1 }] }));
  const xml = await partText(await save(doc), 'xl/workbook.xml');
  assert.match(xml, /activeTab="1"/);
  assert.doesNotMatch(xml, /fullCalcOnLoad/);
});

test('print options of a sheet with links go after <hyperlinks>, before <pageMargins>', async () => {
  const doc = await load(await book('<hyperlinks><hyperlink ref="A1" location="Other!A1"/></hyperlinks><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'));
  apply(doc, setPageSetup(doc.wb, 0, { gridLines: true }));
  assert.match(await sheetPart(await save(doc)), /<\/hyperlinks><printOptions gridLines="1"\/><pageMargins /);
});

// ---------------------------------------------------------------------------
// The band a break spans as the file states it (or leaves to the default 0), what else it says,
// and which of two breaks that come to one place stays: in the model as in the file.

const COLS = '<sheetData>' + Array.from({ length: 40 }, (_, i) => `<row r="${i + 1}"><c r="A${i + 1}"><v>${i}</v></c><c r="J${i + 1}"><v>${i}</v></c></row>`).join('') + '</sheetData>';
const breaksBook = (breaks: string) => buildXlsx({ sheets: [{ name: 'S', xml: worksheetXml(`${COLS}${breaks}`) }, { name: 'T', xml: worksheetXml('<sheetData/>') }] });
const colsOp = (kind: 'insert' | 'delete', at: number, count = 1): Op[] => [{ type: kind, sheet: 0, axis: 'col', at, count }];
const breaksXml = async (doc: Awaited<ReturnType<typeof load>>, element = 'rowBreaks') => {
  const match = new RegExp(`<${element}[^]*?</${element}>`).exec(await sheetPart(await save(doc)));
  return match ? match[0] : '';
};

test('page breaks: a band the file leaves to its default (0) moves as the reader reads it', async () => {
  const cases: [string, string, Op[], number[], string][] = [
    ['no band, column A deleted', '<brk id="10" man="1"/>', colsOp('delete', 0), [], ''],
    ['max 2, A:C deleted', '<brk id="10" max="2" man="1"/>', colsOp('delete', 0, 3), [], ''],
    ['max 2, two columns inserted', '<brk id="10" max="2" man="1"/>', colsOp('insert', 0, 2), [10], '<brk id="10" max="4" man="1" min="2"/>'],
    ['the whole width, columns inserted', '<brk id="4" min="0" max="16383" man="1"/>', colsOp('insert', 3, 2), [4], '<brk id="4" min="0" max="16383" man="1"/>'],
    ['all but the last column, columns inserted', '<brk id="4" min="0" max="16382" man="1"/>', colsOp('insert', 0, 2), [4], '<brk id="4" min="2" max="16383" man="1"/>'],
    ['columns 3..6, all of them deleted', '<brk id="4" min="2" max="5" man="1"/>', colsOp('delete', 1, 6), [], ''],
    // A break across the whole width keeps it when columns go.
    ['the whole width, columns deleted', '<brk id="4" max="16383" man="1"/>', colsOp('delete', 0, 3), [4], '<rowBreaks count="1" manualBreakCount="1"><brk id="4" max="16383" man="1"/></rowBreaks>'],
  ];
  for (const [label, brk, ops, want, xml] of cases) {
    const doc = await load(await breaksBook(`<rowBreaks count="1" manualBreakCount="1">${brk}</rowBreaks>`));
    apply(doc, ops);
    const seen = await breaksEverywhere(doc);
    assert.deepEqual(seen, { model: want, file: want, reread: want }, label);
    const element = await breaksXml(doc);
    if (xml === '') assert.equal(element, '', `${label}: no element`);
    else assert.ok(element.indexOf(xml) >= 0, `${label}: ${element}`);
    doc.history.undo();
    assert.ok(sameBytes(await save(doc), doc.bytes), `${label}: undone`);
  }
});

test('page breaks: what else a break of the file says stays after a command; the default band is not written', async () => {
  const doc = await load(await breaksBook('<rowBreaks count="2" manualBreakCount="2"><brk id="5" max="16383" man="1" foo="x"/><brk id="10" max="16383" man="1" xyz="1"/></rowBreaks>'));
  apply(doc, rowsOp('insert', 0, 1));
  assert.equal(await breaksXml(doc), '<rowBreaks count="2" manualBreakCount="2"><brk id="6" max="16383" man="1" foo="x"/><brk id="11" max="16383" man="1" xyz="1"/></rowBreaks>');
  apply(doc, insertPageBreak(doc.wb, 0, { row: 30 }));
  assert.equal(await breaksXml(doc), '<rowBreaks count="3" manualBreakCount="3"><brk id="6" max="16383" man="1" foo="x"/><brk id="11" max="16383" man="1" xyz="1"/><brk id="30" max="16383" man="1"/></rowBreaks>');
});

test('page breaks: of two that come to one place the one the file lists first stays, in the model as in the file', async () => {
  for (const [label, list, kept] of [
    ['in order', '<brk id="5" max="16383" foo="a"/><brk id="6" max="16383" man="1" foo="b"/>', '<brk id="4" max="16383" foo="a"/>'],
    ['out of order', '<brk id="6" max="16383" man="1" foo="b"/><brk id="5" max="16383" foo="a"/>', '<brk id="4" max="16383" man="1" foo="b"/>'],
  ] as const) {
    const doc = await load(await breaksBook(`<rowBreaks count="2" manualBreakCount="1">${list}</rowBreaks>`));
    apply(doc, rowsOp('delete', 4, 2));
    const element = await breaksXml(doc);
    assert.ok(element.indexOf(kept) >= 0 && (element.match(/<brk /g) || []).length === 1, `${label}: ${element}`);
    // A command puts the break of the model into the file: the same one.
    apply(doc, insertPageBreak(doc.wb, 0, { row: 30 }));
    const after = await breaksXml(doc);
    assert.ok(after.indexOf(kept) >= 0, `${label} after a command: ${after}`);
    const model = doc.wb.sheets[0].pageSetup!.rowBreaks!.find((entry) => entry.at === 4)!;
    assert.deepEqual(model.kept, [['foo', label === 'in order' ? 'a' : 'b']]);
  }
});
