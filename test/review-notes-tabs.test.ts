// Review of notes that travel with cells and of the operations on sheet tabs:
// notes outside the used part of a sheet, what a copy of a sheet keeps and
// what it is said to lose, a workbook with two windows whose sheet is hidden,
// what an undo step of a copy weighs, the colour of a tab in odd parts, and
// the rules that no other test tells apart.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { buildXlsx as buildFromModel } from '../src/formats/xlsx/build';
import { writeXlsx } from '../src/formats/xlsx/write';
import { ZipArchive, cellName, createSheet, emptyStyleSheet, open, parseCellRef } from '../src/index';
import type { CellRange, OfficeFile, Sheet, SheetComment, Workbook } from '../src/index';
import { copyRange, cutRange, pasteAt, setCellInput } from '../src/ops/commands';
import { clearNotes, setComment } from '../src/ops/notes';
import { History, type Op } from '../src/ops/ops';
import { copySheet, copySheetLosses, copySheetName, deleteSheet, moveSheet, setSheetState, setTabColor } from '../src/ops/sheets';
import { SheetDrawings } from '../src/render/sheet-drawings';
import { buildXlsx, utf8, worksheetXml } from './fixtures';
import { comment, commentsXml, drawingWorkbook as bareWorkbook, drawingXml, pic, twoCell, vmlNote, vmlXml, type DrawingBookSpec } from './sheet-drawing-fixtures';

const decoder = new TextDecoder();
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

async function drawingWorkbook(spec: DrawingBookSpec): Promise<Uint8Array> {
  const zip = ZipArchive.read(await bareWorkbook(spec));
  const types = decoder.decode(await zip.read('[Content_Types].xml'));
  zip.set('[Content_Types].xml', utf8(types.replace('</Types>', '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>')));
  return zip.write();
}

interface Loaded {
  file: OfficeFile;
  wb: Workbook;
  history: History;
}

async function load(bytes: Uint8Array): Promise<Loaded> {
  const file = await open(bytes);
  const wb = await file.workbook();
  return { file, wb, history: new History(wb) };
}

async function save(doc: Loaded): Promise<Uint8Array> {
  await writeXlsx(doc.file.package!, doc.wb);
  return doc.file.package!.save();
}

const reopen = async (bytes: Uint8Array) => (await open(bytes)).workbook();
const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const text = async (bytes: Uint8Array, name: string) => decoder.decode(await ZipArchive.read(bytes).read(name));
const range = (ref: string): CellRange => {
  const [a, b] = ref.split(':');
  const from = parseCellRef(a)!;
  const to = parseCellRef(b || a)!;
  return { top: from.row, left: from.col, bottom: to.row, right: to.col };
};
const shown = (sheet: Sheet) =>
  new SheetDrawings(sheet)
    .currentComments()
    .map((placed) => `${cellName(placed.row, placed.col)}:${placed.comment.text}`)
    .sort();

function same(actual: Uint8Array, expected: Uint8Array, label: string): void {
  const n = Math.min(actual.length, expected.length);
  let at = 0;
  while (at < n && actual[at] === expected[at]) at++;
  if (at < n || actual.length !== expected.length) assert.fail(`${label}: differs at offset ${at} (lengths ${actual.length} and ${expected.length})`);
}

const rows = (cells: number[][]) => `<sheetData>${cells.map((line, r) => `<row r="${r + 1}">${line.map((value, c) => `<c r="${cellName(r, c)}"><v>${value}</v></c>`).join('')}</row>`).join('')}</sheetData>`;

// --- 1: notes outside the used part ------------------------------------------------------------

/** Data in A1:C6; notes on A1 and C6 (cells with data), on E9 and on B8 (empty cells beyond the data). */
const beyond = () =>
  drawingWorkbook({
    sheet: rows([[1, 2, 3], [4, 5, 6], [7, 8, 9], [1, 1, 1], [2, 2, 2], [3, 3, 3]]),
    comments: commentsXml(['Ada'], comment('A1', 0, '<t>bir</t>') + comment('C6', 0, '<t>iki</t>') + comment('E9', 0, '<t>dışarıda</t>') + comment('B8', 0, '<t>altta</t>')),
    vml: vmlXml(vmlNote(0, 0) + vmlNote(5, 2) + vmlNote(8, 4) + vmlNote(7, 1)),
  });

test('review: a note on an empty cell beyond the data is copied and cut with the selection it is in', async () => {
  const original = await beyond();
  const doc = await load(original);
  const sheet = doc.wb.sheets[0];
  const copied = copyRange(doc.wb, 0, range('A1:E9'));
  assert.deepEqual((copied.notes || []).map((note) => `${cellName(note.row, note.col)}:${note.text}`).sort(), ['A1:bir', 'B8:altta', 'C6:iki', 'E9:dışarıda']);
  // Cut: what goes is what the copy holds.
  apply(doc, cutRange(doc.wb, 0, range('A1:E9')));
  assert.deepEqual(shown(sheet), []);
  const pasted = pasteAt(doc.wb, 0, range('G11'), copied.text, copied);
  apply(doc, pasted.ops);
  // The notes stand where they stood in the selection, and the selection that was pasted is as large as it.
  assert.deepEqual(shown(sheet), ['G11:bir', 'H18:altta', 'I16:iki', 'K19:dışarıda']);
  assert.deepEqual(pasted.range, range('G11:K19'));
  const saved = await save(doc);
  assert.deepEqual(shown((await reopen(saved)).sheets[0]), shown(sheet));
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(shown(sheet), ['A1:bir', 'B8:altta', 'C6:iki', 'E9:dışarıda']);

  // A plain copy takes the same notes, and leaves them where they are.
  apply(doc, pasteAt(doc.wb, 0, range('A20'), copied.text, copied).ops);
  assert.deepEqual(shown(sheet), ['A1:bir', 'A20:bir', 'B27:altta', 'B8:altta', 'C25:iki', 'C6:iki', 'E28:dışarıda', 'E9:dışarıda']);
  doc.history.undo();
  // Whole columns: the notes of the columns, however far down, and no cell beyond the data.
  const columns = copyRange(doc.wb, 0, { top: 0, left: 0, bottom: 1048575, right: 4 });
  assert.equal((columns.notes || []).length, 4);
  assert.equal(columns.cells.length, 6);
});

test('review: a cut leaves the notes of rows the filter hides; a pasted note that would land beyond the sheet is left out; conversations stay', async () => {
  const doc = await load(await beyond());
  const sheet = doc.wb.sheets[0];
  // Row 3 (A3 = 7) is hidden by the filter; it gets a note first.
  apply(doc, setComment(doc.wb, 0, 2, 0, 'gizli'));
  apply(doc, [{ type: 'setAutoFilter', sheet: 0, autoFilter: { range: range('A1:C6'), columns: [{ colId: 0, kind: 'values', values: ['4', '1', '2', '3'] }] }, hide: [2] }]);
  const copied = copyRange(doc.wb, 0, range('A1:C6'));
  assert.deepEqual((copied.notes || []).map((note) => note.text).sort(), ['bir', 'iki']);
  apply(doc, cutRange(doc.wb, 0, range('A1:C6')));
  assert.deepEqual(shown(sheet), ['A3:gizli', 'B8:altta', 'E9:dışarıda']);
  assert.equal(sheet.rows.get(2)!.cells[0].value, 7);
  doc.history.undo();

  // Pasted at the last rows of the sheet: the note of the sixth copied row (C6, the fifth that is shown) has no cell there.
  const edge = pasteAt(doc.wb, 0, { top: 1048573, left: 0, bottom: 1048573, right: 0 }, copied.text, copied);
  const notes = edge.ops.filter((op) => op.type === 'setComments');
  assert.equal(notes.length, 1);
  apply(doc, edge.ops);
  assert.deepEqual(shown(sheet).filter((line) => /1048/.test(line)), ['A1048574:bir']);
  doc.history.undo();

  // clearNotes() leaves a conversation where it is.
  sheet.comments = (sheet.comments || []).concat([{ row: 20, col: 0, text: 'zincir', threaded: true } as SheetComment]);
  const cleared = clearNotes(doc.wb, 0, range('A1:Z30'));
  assert.equal(cleared.length, 1);
  apply(doc, cleared);
  assert.deepEqual(shown(sheet), ['A21:zincir']);
});

test('review: a paste that would give the sheet more notes than it holds is refused as a whole', () => {
  const sheet = createSheet({ name: 'S', sheetId: 1, state: 'visible', kind: 'worksheet', part: '' });
  const many: SheetComment[] = [];
  for (let i = 0; i < 99999; i++) many.push({ row: i + 10, col: 5, text: 'x' });
  many.push({ row: 0, col: 0, text: 'a' }, { row: 1, col: 0, text: 'b' });
  sheet.comments = many.slice(0, 100000);
  const wb: Workbook = { sheets: [sheet], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  const history = new History(wb);
  history.apply(setCellInput(wb, 0, 0, 0, '1'), undefined, undefined);
  const copied = copyRange(wb, 0, range('A1'));
  assert.equal(copied.notes!.length, 1);
  // Onto a cell that has a note: one goes, one comes.
  assert.equal(pasteAt(wb, 0, range('F11'), copied.text, copied).ops.length, 2);
  // Onto a cell without one: one too many, and no cell is written either.
  assert.throws(() => pasteAt(wb, 0, range('B1'), copied.text, copied), /^RangeError: The sheet has too many comments$/);
});

// --- 3: what a copy of a sheet keeps and what it is said to lose -------------------------------

const PROTECTION = '<sheetProtection algorithmName="SHA-512" hashValue="q2w3e4==" saltValue="abc=" spinCount="100000" sheet="1" objects="1" scenarios="1" formatCells="0"/>';
const RANGES = '<protectedRanges><protectedRange sqref="A1:B2" name="Alan" securityDescriptor="O:WDG:WD"/><protectedRange password="CC1A" sqref="D4" name="İki"/></protectedRanges>';
const SPARK =
  '<extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Kor!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>';

/** A protected sheet with everything a sheet part may hold that the model has no field for. */
const guarded = () =>
  buildXlsx({
    sheets: [
      {
        name: 'Kor',
        xml: worksheetXml(
          '<sheetPr codeName="Sayfa1"><outlinePr summaryBelow="0"/></sheetPr><sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" outlineLevelRow="2"/>' +
            '<sheetData><row r="1" spans="1:3" thickBot="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row><row r="2"><c r="A2"><v>0</v></c></row></sheetData>' +
            `${PROTECTION}${RANGES}<sortState ref="A1:A2"><sortCondition ref="A1:A2"/></sortState>` +
            '<customSheetViews><customSheetView guid="{11111111-2222-3333-4444-555555555555}"/></customSheetViews><phoneticPr fontId="1"/>' +
            '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>' +
            `<cellWatches><cellWatch r="A1"/></cellWatches><ignoredErrors><ignoredError sqref="A1" numberStoredAsText="1"/></ignoredErrors>${SPARK}`,
        ),
      },
    ],
  });

test('review: the copy of a protected sheet is protected, with the ranges that are open in it', async () => {
  const original = await guarded();
  const doc = await load(original);
  apply(doc, copySheet(doc.wb, 0));
  const saved = await save(doc);
  const copy = await text(saved, 'xl/worksheets/sheet2.xml');
  assert.ok(copy.indexOf(PROTECTION) >= 0, copy);
  assert.ok(copy.indexOf(RANGES) >= 0, copy);
  // Where the schema has them: behind the cells, before what is printed.
  assert.match(copy, /<\/sheetData><sheetProtection [^>]*\/><protectedRanges>.*<\/protectedRanges><pageMargins /);
  const again = await reopen(saved);
  assert.deepEqual(again.sheets[1].protection, again.sheets[0].protection);
  assert.equal(again.sheets[1].protection!.attrs.length, 8);
  // The sheet of the file is as it was.
  same(await ZipArchive.read(saved).read('xl/worksheets/sheet1.xml'), await ZipArchive.read(original).read('xl/worksheets/sheet1.xml'), 'the source');
  // A file made from the model says the same.
  const built = await (await open(await buildFromModel(doc.wb))).workbook();
  assert.deepEqual(built.sheets[1].protection, doc.wb.sheets[0].protection);
  doc.history.undo();
  same(await save(doc), original, 'undone');
});

test('review: whatever a copy of a sheet does not get is counted, by its kind', async () => {
  const doc = await load(await guarded());
  const losses = copySheetLosses(doc.wb, 0);
  const kinds = losses.map((loss) => loss.what).sort();
  // The row with thickBot is not among them: an attribute without a namespace goes into the copy as the file has it
  // (fourth review; it was counted as lost before, and was).
  assert.deepEqual(kinds, ['cellWatches', 'codeName', 'customSheetViews', 'ignoredErrors', 'outlineLevels', 'outlinePr', 'phoneticPr', 'sortState', 'sparklineGroups']);
  assert.deepEqual(doc.wb.sheets[0].rows.get(0)!.attrs, [['thickBot', '1']]);
  // Protection is not among them: it is copied.
  assert.ok(kinds.indexOf('sheetProtection') < 0 && kinds.indexOf('protectedRanges') < 0);

  // Content of other parts: a conversation, a picture, a table, a pivot table.
  const threaded =
    '<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><threadedComment ref="B2" dT="2024-01-01T00:00:00.00" personId="{P1}" id="{T1}"><text>İlk</text></threadedComment></ThreadedComments>';
  const rich = await load(
    await drawingWorkbook({
      sheet: rows([[1]]),
      drawing: twoCell([0, 0], [2, 2], pic('rIdP')),
      media: { rIdP: new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10]) },
      comments: commentsXml(['Ada'], comment('A1', 0, '<t>not</t>') + comment('B2', 0, '<t>[Threaded comment]</t>')),
      vml: vmlXml(vmlNote(0, 0) + vmlNote(1, 1)),
      threaded,
    }),
  );
  rich.wb.sheets[0].anchored = (rich.wb.sheets[0].anchored || []).concat(['table', 'pivotTable']);
  const counted = copySheetLosses(rich.wb, 0);
  assert.deepEqual(counted, [
    { what: 'drawings', count: 1 },
    { what: 'threadedComments', count: 1 },
    { what: 'table', count: 1 },
    { what: 'pivotTable', count: 1 },
  ]);
  // The conversation is left out of the copy; the note comes along.
  apply(rich, copySheet(rich.wb, 0));
  assert.deepEqual(shown(rich.wb.sheets[1]), ['A1:not']);
  assert.equal((rich.wb.sheets[1].comments || []).some((item) => item.threaded), false);
  assert.equal(rich.wb.sheets[1].drawings, undefined);
  assert.ok(drawingXml('').length > 0);
});

test('review: a copy and its source share no row: editing one leaves the other; the filter comes without what the file kept beside it', async () => {
  const doc = await load(
    await buildXlsx({
      sheets: [{ name: 'Veri', xml: worksheetXml(`${rows([[1, 2], [3, 4]])}<autoFilter ref="A1:B2" xr:uid="{AAAA}" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision"><extLst><ext uri="{X}"/></extLst></autoFilter>`) }],
    }),
  );
  const wb = doc.wb;
  assert.ok(wb.sheets[0].autoFilter!.attrs || wb.sheets[0].autoFilter!.extra);
  apply(doc, copySheet(wb, 0));
  const [source, copy] = wb.sheets;
  assert.deepEqual(copy.autoFilter!.range, range('A1:B2'));
  assert.equal(copy.autoFilter!.attrs, undefined);
  assert.equal(copy.autoFilter!.extra, undefined);
  for (const index of [0, 1]) {
    assert.notEqual(copy.rows.get(index), source.rows.get(index));
    assert.notEqual(copy.rows.get(index)!.cells, source.rows.get(index)!.cells);
  }
  apply(doc, setCellInput(wb, 0, 0, 0, '100'));
  apply(doc, setCellInput(wb, 1, 1, 1, '400'));
  assert.deepEqual([source.rows.get(0)!.cells[0].value, copy.rows.get(0)!.cells[0].value], [100, 1]);
  assert.deepEqual([source.rows.get(1)!.cells[1].value, copy.rows.get(1)!.cells[1].value], [4, 400]);
  apply(doc, [{ type: 'delete', sheet: 1, axis: 'row', at: 0, count: 1 }]);
  assert.equal(source.rows.size, 2);
  const saved = await reopen(await save(doc));
  assert.equal(saved.sheets[0].rows.get(0)!.cells[0].value, 100);
  assert.equal(saved.sheets[1].rows.get(0)!.cells[1].value, 400);
});

test('review: the name of a copy counts on, is cut to 31 characters, and never ends on half a character', () => {
  const sheet = (name: string, id: number) => createSheet({ name, sheetId: id, state: 'visible', kind: 'worksheet', part: '' });
  const wb: Workbook = { sheets: [sheet('Ad', 1), sheet('Ad (2)', 2), sheet('a'.repeat(31), 3), sheet(`${'b'.repeat(26)}😀xyz`, 4)], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  assert.equal(copySheetName(wb, 0), 'Ad (3)');
  assert.equal(copySheetName(wb, 1), 'Ad (3)');
  assert.equal(copySheetName(wb, 2), `${'a'.repeat(27)} (2)`);
  // 27 characters would end between the two halves of the emoji: the name is one shorter.
  assert.equal(copySheetName(wb, 3), `${'b'.repeat(26)} (2)`);
  wb.sheets.push(sheet('AD (3)', 5));
  assert.equal(copySheetName(wb, 0), 'Ad (4)');
});

// --- 4: a workbook with two windows ------------------------------------------------------------

const twoWindows = () =>
  buildXlsx({
    workbookHead: '<bookViews><workbookView xWindow="0" yWindow="0" firstSheet="1" activeTab="1"/><workbookView xWindow="100" yWindow="100" activeTab="2"/></bookViews>',
    sheets: ['Bir', 'İki', 'Üç', 'Dört'].map((name, index) => ({
      name,
      xml: worksheetXml(`<sheetViews><sheetView${index === 1 ? ' tabSelected="1"' : ''} workbookViewId="0"/><sheetView${index === 2 ? ' tabSelected="1"' : ''} workbookViewId="1"/></sheetViews>${rows([[index]])}`),
    })),
  });

/** The sheets the windows of a saved workbook open on, and the sheets each window has marked as selected. */
async function windowsOf(bytes: Uint8Array): Promise<{ active: number[]; first: (number | undefined)[]; selected: number[][]; states: string[] }> {
  const book = await text(bytes, 'xl/workbook.xml');
  const views = [...book.matchAll(/<workbookView [^>]*>/g)].map((match) => match[0]);
  const states = [...book.matchAll(/<sheet [^>]*>/g)].map((match) => (/state="(\w+)"/.exec(match[0]) || ['', 'visible'])[1]);
  const selected: number[][] = views.map(() => []);
  for (let i = 0; i < states.length; i++) {
    const sheet = await text(bytes, `xl/worksheets/sheet${i + 1}.xml`);
    for (const view of sheet.matchAll(/<sheetView [^>]*>/g)) {
      const id = Number((/workbookViewId="(\d+)"/.exec(view[0]) || ['', '0'])[1]);
      if (/tabSelected="1"/.test(view[0])) selected[id].push(i);
    }
  }
  return {
    active: views.map((view) => Number((/activeTab="(\d+)"/.exec(view) || ['', '0'])[1])),
    first: views.map((view) => (/firstSheet="(\d+)"/.exec(view) ? Number(/firstSheet="(\d+)"/.exec(view)![1]) : undefined)),
    selected,
    states,
  };
}

test('review: in a workbook with two windows no window opens on a sheet that was hidden', async () => {
  const original = await twoWindows();
  const doc = await load(original);
  const wb = doc.wb;
  assert.equal(wb.activeSheet, 1);
  // The sheet of the second window.
  apply(doc, setSheetState(wb, 2, 'hidden'));
  let saved = await windowsOf(await save(doc));
  assert.deepEqual(saved.states, ['visible', 'visible', 'hidden', 'visible']);
  for (const active of saved.active) assert.equal(saved.states[active], 'visible', `a window opens on sheet ${active}`);
  assert.equal(saved.active[0], 1);
  saved.active.forEach((active, window) => assert.deepEqual(saved.selected[window], [active], `window ${window}`));
  // And the sheet of the first one, which is the one the model shows.
  apply(doc, setSheetState(wb, 1, 'hidden'));
  saved = await windowsOf(await save(doc));
  assert.deepEqual(saved.states, ['visible', 'hidden', 'hidden', 'visible']);
  for (const active of saved.active) assert.equal(saved.states[active], 'visible', `a window opens on sheet ${active}`);
  for (const first of saved.first) if (first !== undefined) assert.equal(saved.states[first], 'visible', `the first tab shown is sheet ${first}`);
  saved.active.forEach((active, window) => assert.deepEqual(saved.selected[window], [active], `window ${window}`));
  assert.equal(saved.active[0], wb.activeSheet);
  await reopen(await save(doc));
  doc.history.undo();
  doc.history.undo();
  same(await save(doc), original, 'undone');
});

test('review: a sheet that is not the active one is hidden on its own, and the file says so', async () => {
  const original = await buildXlsx({ sheets: ['Bir', 'İki', 'Üç'].map((name) => ({ name, xml: worksheetXml(rows([[1]])) })) });
  const doc = await load(original);
  apply(doc, setSheetState(doc.wb, 2, 'hidden'));
  const saved = await save(doc);
  assert.match(await text(saved, 'xl/workbook.xml'), /<sheet name="Üç" sheetId="3" r:id="rId3" state="hidden"\/>/);
  assert.equal((await reopen(saved)).sheets[2].state, 'hidden');
  for (const name of ['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml']) same(await ZipArchive.read(saved).read(name), await ZipArchive.read(original).read(name), name);
  doc.history.undo();
  same(await save(doc), original, 'undone');
});

test('review: hiding the active sheet shows the next visible one to its right, and the one to its left only when there is none', () => {
  const make = (states: string[], active: number): Workbook => ({
    sheets: states.map((state, i) => createSheet({ name: `S${i}`, sheetId: i + 1, state: state as 'visible', kind: 'worksheet', part: '' })),
    styles: emptyStyleSheet(),
    sharedStrings: [],
    definedNames: [],
    date1904: false,
    activeSheet: active,
  });
  // Visible sheets on both sides, a hidden one in between on the right: the right one wins.
  const both = make(['visible', 'visible', 'hidden', 'visible'], 1);
  new History(both).apply(setSheetState(both, 1, 'hidden'), undefined, undefined);
  assert.equal(both.activeSheet, 3);
  // Nothing visible to the right: the nearest one to the left.
  const left = make(['visible', 'visible', 'visible', 'hidden'], 2);
  new History(left).apply(setSheetState(left, 2, 'hidden'), undefined, undefined);
  assert.equal(left.activeSheet, 1);
  // Another sheet than the active one: the active one stays.
  const other = make(['visible', 'visible', 'visible'], 0);
  new History(other).apply(setSheetState(other, 2, 'hidden'), undefined, undefined);
  assert.equal(other.activeSheet, 0);
});

// --- 6: small things ---------------------------------------------------------------------------

test('review: the step that copies a sheet holds no rows, however large the sheet', () => {
  const big = createSheet({ name: 'Büyük', sheetId: 1, state: 'visible', kind: 'worksheet', part: '' });
  for (let r = 0; r < 20000; r++) big.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: r, style: 0 }] });
  big.rowCount = 20000;
  big.colCount = 1;
  const wb: Workbook = { sheets: [big], styles: emptyStyleSheet(), sharedStrings: [], definedNames: [], date1904: false, activeSheet: 0 };
  const history = new History(wb);
  for (let i = 0; i < 20; i++) history.apply(copySheet(wb, 0), undefined, undefined);
  assert.equal(wb.sheets.length, 21);
  let steps = 0;
  while (history.canUndo) {
    history.undo();
    steps++;
  }
  assert.equal(steps, 20);
  assert.equal(wb.sheets.length, 1);
  // A deleted sheet is held by its step, and weighs what it holds.
  history.apply(copySheet(wb, 0), undefined, undefined);
  history.apply(deleteSheet(wb, 1), undefined, undefined);
  history.undo();
  assert.equal(wb.sheets[1].rows.size, 20000);
});

test('review: a tab colour that is the one of the file again costs the sheet part nothing', async () => {
  const original = await buildXlsx({
    sheets: [
      { name: 'Bir', xml: worksheetXml(`<sheetPr><tabColor theme="4"/></sheetPr>${rows([[1]])}`) },
      { name: 'İki', xml: worksheetXml(rows([[2]])) },
    ],
  });
  const doc = await load(original);
  const pkg = doc.file.package!;
  const reads: string[] = [];
  const readPart = pkg.readPart.bind(pkg);
  pkg.readPart = (part: string) => {
    reads.push(part);
    return readPart(part);
  };
  apply(doc, setTabColor(doc.wb, 0, '#FF0000'));
  apply(doc, setTabColor(doc.wb, 1, '#00FF00'));
  doc.history.undo();
  doc.history.undo();
  reads.length = 0;
  same(await save(doc), original, 'undone');
  assert.deepEqual(reads.filter((part) => /worksheets/.test(part)), []);
  // Changed: the part is read once.
  doc.history.redo();
  reads.length = 0;
  await save(doc);
  assert.deepEqual(reads.filter((part) => /worksheets/.test(part)), ['/xl/worksheets/sheet1.xml']);
});

test('review: a tab colour is written into parts of every make, and undone each is the file again', async () => {
  const root = `xmlns="${NS}" xmlns:r="${NS_REL}"`;
  const parts: [string, string | Uint8Array, RegExp][] = [
    ['children without a colour', `${XML}<worksheet ${root}><sheetPr codeName="K"><outlinePr summaryBelow="0"/><pageSetUpPr fitToPage="1"/></sheetPr><sheetData/></worksheet>`, /<sheetPr codeName="K"><tabColor rgb="FF123456"\/><outlinePr summaryBelow="0"\/><pageSetUpPr fitToPage="1"\/><\/sheetPr>/],
    ['a prefix', `${XML}<x:worksheet xmlns:x="${NS}"><x:sheetPr/><x:sheetData/></x:worksheet>`, /<x:sheetPr><x:tabColor rgb="FF123456"\/><\/x:sheetPr><x:sheetData\/>/],
    ['no properties', `${XML}<x:worksheet xmlns:x="${NS}"><x:sheetData/></x:worksheet>`, /<x:worksheet [^>]*><x:sheetPr><x:tabColor rgb="FF123456"\/><\/x:sheetPr><x:sheetData\/>/],
    ['a byte order mark and CRLF', utf8(`﻿${XML}\r\n<worksheet ${root}>\r\n<sheetPr filterMode="0"/>\r\n<sheetData/>\r\n</worksheet>\r\n`), /^﻿<\?xml[^>]*>\r\n<worksheet [^>]*>\r\n<sheetPr filterMode="0"><tabColor rgb="FF123456"\/><\/sheetPr>\r\n<sheetData\/>\r\n<\/worksheet>\r\n$/],
    ['a colour by index', `${XML}<worksheet ${root}><sheetPr><tabColor indexed="10"/></sheetPr><sheetData/></worksheet>`, /<sheetPr><tabColor rgb="FF123456"\/><\/sheetPr>/],
    ['a root without content', `${XML}<worksheet ${root}/>`, /<worksheet [^>]*><sheetPr><tabColor rgb="FF123456"\/><\/sheetPr><\/worksheet>/],
  ];
  for (const [label, xml, wanted] of parts) {
    const bytes = await buildXlsx({ sheets: [{ name: 'Tek', xml: worksheetXml('<sheetData/>') }, { name: 'Öbür', xml: worksheetXml('<sheetData/>') }] });
    const zip = ZipArchive.read(bytes);
    zip.set('xl/worksheets/sheet1.xml', typeof xml === 'string' ? utf8(xml) : xml);
    const original = await zip.write();
    const doc = await load(original);
    const before = doc.wb.sheets[0].tabColor;
    apply(doc, setTabColor(doc.wb, 0, '#123456'));
    const saved = await save(doc);
    assert.match(new TextDecoder('utf-8', { ignoreBOM: true }).decode(await ZipArchive.read(saved).read('xl/worksheets/sheet1.xml')), wanted, label);
    assert.deepEqual((await reopen(saved)).sheets[0].tabColor, { rgb: 'FF123456' }, label);
    same(await ZipArchive.read(saved).read('xl/worksheets/sheet2.xml'), await ZipArchive.read(original).read('xl/worksheets/sheet2.xml'), `${label}: the other sheet`);
    doc.history.undo();
    assert.deepEqual(doc.wb.sheets[0].tabColor, before, label);
    same(await save(doc), original, `${label}: undone`);
  }
});

test('review: moving a sheet with a tab colour, a note and a hidden neighbour keeps each with its sheet', async () => {
  const doc = await load(await buildXlsx({ sheets: ['Bir', 'İki', 'Üç'].map((name) => ({ name, xml: worksheetXml(rows([[1]])) })) }));
  apply(doc, setTabColor(doc.wb, 0, '#AA0000'));
  apply(doc, setComment(doc.wb, 0, 1, 1, 'not'));
  apply(doc, setSheetState(doc.wb, 1, 'hidden'));
  apply(doc, moveSheet(doc.wb, 0, 2));
  const again = await reopen(await save(doc));
  assert.deepEqual(again.sheets.map((sheet) => `${sheet.name}:${sheet.state}:${sheet.tabColor ? sheet.tabColor.rgb : '-'}:${shown(sheet).join()}`), ['İki:hidden:-:', 'Üç:visible:-:', 'Bir:visible:FFAA0000:B2:not']);
});
