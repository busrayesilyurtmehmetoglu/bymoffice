// Cases found by review, and a randomised check of the property undo rests on.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { getCell, open, type Cell, type Sheet, type Workbook } from '../src/index';
import { clearRange, copyRange, editTextAt, pasteAt, setCellInput } from '../src/ops/commands';
import { History, applyOp, ensureCellFormat, ensureNumberFormat, type Op } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (r: string, v: number) => `<c r="${r}"><v>${v}</v></c>`;
const formula = (r: string, f: string, attrs = '') => `<c r="${r}"><f${attrs}>${f}</f><v>0</v></c>`;
const dependant = (r: string, si = 0) => `<c r="${r}"><f t="shared" si="${si}"/><v>0</v></c>`;
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

async function workbook(): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      styles:
        `<?xml version="1.0"?><styleSheet xmlns="${NS}"><fonts><font/></fonts><fills><fill/></fills><borders><border/></borders><cellStyleXfs><xf/></cellStyleXfs>` +
        '<cellXfs><xf numFmtId="0"/><xf numFmtId="49"/><xf numFmtId="2"><alignment horizontal="center"/></xf></cellXfs></styleSheet>',
      workbookTail: '<definedNames><definedName name="Ad">Veri!$A$1:$A$4</definedName></definedNames>',
      sheets: [
        {
          name: 'Veri',
          xml: worksheetXml(
            '<sheetData>' +
              `<row r="1">${num('A1', 1)}${formula('B1', 'A1*2', ' t="shared" ref="B1:B4" si="0"')}<c r="C1" s="1" t="inlineStr"><is><t>00123</t></is></c></row>` +
              `<row r="2">${num('A2', 2)}${dependant('B2')}</row>` +
              `<row r="3">${num('A3', 3)}${dependant('B3')}</row>` +
              `<row r="4">${num('A4', 4)}${dependant('B4')}</row>` +
              `<row r="6">${formula('A6', 'A1:A2*2', ' t="array" ref="A6:A8"')}</row><row r="7">${num('A7', 0)}</row><row r="8">${num('A8', 0)}</row>` +
              `<row r="9" s="2" customFormat="1">${num('A9', 9)}</row>` +
              '</sheetData><mergeCells><mergeCell ref="D1:E2"/></mergeCells>',
          ),
        },
        {
          name: 'Özet',
          xml: worksheetXml(
            `<sheetData><row r="1">${formula('A1', 'Veri!A2', ' t="shared" ref="A1:A4" si="0"')}${formula('B1', 'C1+1', ' t="shared" ref="B1:B3" si="1"')}</row>` +
              `<row r="2">${dependant('A2')}${dependant('B2', 1)}</row><row r="3">${dependant('A3')}${dependant('B3', 1)}</row><row r="4">${dependant('A4')}</row></sheetData>`,
          ),
        },
      ],
    }),
  );
  return file.workbook();
}

const f = (sheet: Sheet, row: number, col: number) => getCell(sheet, row, col)?.formula;

/** Everything operations may touch, as plain data. */
function picture(wb: Workbook): string {
  return JSON.stringify({
    names: wb.definedNames,
    sheets: wb.sheets.map((sheet) => ({
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      merges: sheet.merges,
      links: sheet.hyperlinks,
      columns: sheet.columns,
      dimension: sheet.dimension,
      edits: sheet.structuralEdits,
      rows: [...sheet.rows.entries()]
        .filter(([, row]) => row.cells.some(() => true) || row.height !== undefined || row.style !== undefined)
        .sort((a, b) => a[0] - b[0])
        .map(([index, row]) => [index, row.height, row.style, row.cells.map((cell, col) => (cell ? [col, cell.kind, cell.value, cell.style, cell.formula] : null)).filter(Boolean)]),
    })),
  });
}

test('ops: shared formulas of other sheets follow a structural edit', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const undo = applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 });
  const summary = wb.sheets[1];
  // A1..A4 referred to Veri!A2..A5; a row inserted at 3 moves all but the first.
  assert.deepEqual([0, 1, 2, 3].map((row) => f(summary, row, 0)?.text), ['Veri!A2', 'Veri!A4', 'Veri!A5', 'Veri!A6']);
  assert.ok([0, 1, 2, 3].every((row) => f(summary, row, 0)?.type === 'normal'));
  // A group that does not refer to the edited sheet stays shared.
  assert.deepEqual([f(summary, 0, 1), f(summary, 1, 1)], [{ text: 'C1+1', type: 'shared', ref: 'B1:B3', si: 1 }, { text: '', type: 'shared', si: 1 }]);
  applyOp(wb, undo);
  assert.equal(picture(wb), before);

  applyOp(wb, { type: 'delete', sheet: 0, axis: 'row', at: 1, count: 1 });
  assert.deepEqual([0, 1, 2, 3].map((row) => f(wb.sheets[1], row, 0)?.text), ['Veri!#REF!', 'Veri!A2', 'Veri!A3', 'Veri!A4']);
});

test('ops: overwriting the cell that holds a shared formula keeps the formulas that depend on it', async () => {
  for (const overwrite of [
    (wb: Workbook) => setCellInput(wb, 0, 0, 1, '5'),
    (wb: Workbook) => clearRange(wb, 0, { top: 0, left: 1, bottom: 0, right: 1 }),
    (wb: Workbook) => pasteAt(wb, 0, { top: 0, left: 1, bottom: 0, right: 1 }, 'x').ops,
  ]) {
    const wb = await workbook();
    const before = picture(wb);
    const history = new History(wb);
    history.apply(overwrite(wb), undefined, undefined);
    const sheet = wb.sheets[0];
    assert.deepEqual([1, 2, 3].map((row) => f(sheet, row, 1)), [
      { text: 'A2*2', type: 'normal' },
      { text: 'A3*2', type: 'normal' },
      { text: 'A4*2', type: 'normal' },
    ]);
    assert.equal(editTextAt(wb, sheet, 1, 1), '=A2*2');
    history.undo();
    assert.equal(picture(wb), before);
  }
  // Overwriting a dependant leaves the group alone.
  const wb = await workbook();
  applyOp(wb, setCellInput(wb, 0, 2, 1, '7')[0]);
  assert.deepEqual([f(wb.sheets[0], 0, 1)?.type, f(wb.sheets[0], 1, 1)?.text, f(wb.sheets[0], 3, 1)?.type], ['shared', '', 'shared']);
  assert.equal(copyRange(wb, 0, { top: 3, left: 1, bottom: 3, right: 1 }).cells[0][0]?.formula?.text, 'A4*2');
});

test('commands: a cell formatted as Text stores what is typed', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  assert.equal(editTextAt(wb, sheet, 0, 2), '00123');
  for (const text of ['0123', '12345678901234567890', '=A2', 'TRUE', '15.03.2024', ' 5 ']) {
    applyOp(wb, setCellInput(wb, 0, 0, 2, text)[0]);
    const cell = getCell(sheet, 0, 2);
    assert.deepEqual([cell?.kind, cell?.value, cell?.formula, cell?.style], ['string', text, undefined, 1], text);
    assert.equal(editTextAt(wb, sheet, 0, 2), text);
  }
  applyOp(wb, setCellInput(wb, 0, 0, 2, '')[0]);
  assert.deepEqual([getCell(sheet, 0, 2)?.kind, getCell(sheet, 0, 2)?.style], ['blank', 1]);
});

test('ops: inserts are refused when a row with only properties would leave the sheet', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  sheet.rows.set(1048575, { height: 40, customHeight: true, hidden: false, outlineLevel: 0, collapsed: false, cells: [] });
  const before = picture(wb);
  assert.throws(() => applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 5 }), /pushed off the sheet/);
  assert.equal(picture(wb), before);
  assert.ok([...sheet.rows.keys()].every((key) => key < 1048576));
});

test('ops: one setCells operation applies as a whole or not at all', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const cell: Cell = { kind: 'number', value: 99, style: 0 };
  for (const bad of [{ row: -1, col: 0 }, { row: 0, col: 16384 }, { row: 0.5, col: 0 }, { row: NaN, col: 0 }]) {
    assert.throws(() => applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell }, { ...bad, cell }] }), RangeError);
    assert.equal(picture(wb), before);
  }
  const history = new History(wb);
  assert.throws(() => history.apply([{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell }, { row: -1, col: 0, cell }] }], undefined, undefined), RangeError);
  assert.equal(history.canUndo, false);
  assert.equal(picture(wb), before);
});

test('ops: a snapshot taken by a structural step is not disturbed by later cell changes', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  const rowBefore = sheet.rows.get(0);
  const cellsBefore = rowBefore?.cells;
  applyOp(wb, setCellInput(wb, 0, 0, 0, '100')[0]);
  // The row object the sheet held before is untouched: changes put a new row in its place.
  assert.equal(rowBefore?.cells, cellsBefore);
  assert.equal(rowBefore?.cells[0].value, 1);
  assert.notEqual(sheet.rows.get(0), rowBefore);

  const undoInsert = applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 });
  const afterInsert = picture(wb);
  const undoEdit = applyOp(wb, setCellInput(wb, 0, 1, 0, '555')[0]);
  applyOp(wb, undoEdit);
  assert.equal(picture(wb), afterInsert);
  const redoInsert = applyOp(wb, undoInsert);
  assert.equal(getCell(sheet, 0, 0)?.value, 100);
  applyOp(wb, redoInsert);
  assert.equal(picture(wb), afterInsert);
});

test('history: steps that hold a lot of rows are dropped before memory runs out', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  // 60,000 formula rows: every insert above them replaces all of them.
  for (let r = 100; r < 60100; r++) {
    sheet.rows.set(r, { customHeight: false, hidden: false, outlineLevel: 0, collapsed: false, cells: [{ kind: 'number', value: 0, style: 0, formula: { text: `A${r}+1`, type: 'normal' } }] });
  }
  sheet.rowCount = 60100;
  const history = new History(wb);
  for (let i = 0; i < 12; i++) history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 50, count: 1 }], undefined, undefined);
  let undone = 0;
  while (history.undo()) undone++;
  assert.ok(undone >= 1 && undone <= 5, `${undone} steps could be undone`);
});

test('ops: array formulas cannot be cut through', async () => {
  const wb = await workbook();
  const before = picture(wb);
  for (const op of [
    { type: 'insert', sheet: 0, axis: 'row', at: 6, count: 1 },
    { type: 'delete', sheet: 0, axis: 'row', at: 5, count: 1 },
    { type: 'delete', sheet: 0, axis: 'row', at: 7, count: 5 },
  ] as Op[]) {
    assert.throws(() => applyOp(wb, op), /array formula/, JSON.stringify(op));
    assert.equal(picture(wb), before);
  }
  // Moving the whole range, or deleting all of it, is allowed.
  applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 5, count: 1 });
  assert.equal(f(wb.sheets[0], 6, 0)?.ref, 'A7:A9');
  applyOp(wb, { type: 'delete', sheet: 0, axis: 'row', at: 6, count: 3 });
  // The row that followed the array is now where it began, and no array formula is left.
  assert.equal(getCell(wb.sheets[0], 6, 0)?.value, 9);
  assert.ok([...wb.sheets[0].rows.values()].every((row) => row.cells.every((cell) => cell.formula?.type !== 'array')));
});

test('commands: emptied cells keep a format that differs from what the row or column gives', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  // Row 9 has format 2; A9 has format 0 of its own. Without the cell the position would show format 2.
  history(wb, clearRange(wb, 0, { top: 8, left: 0, bottom: 8, right: 0 }));
  assert.deepEqual([getCell(sheet, 8, 0)?.kind, getCell(sheet, 8, 0)?.style], ['blank', 0]);
  // A cell typed into that row takes the row's format, and is simply removed when emptied.
  history(wb, setCellInput(wb, 0, 8, 3, '4'));
  assert.equal(getCell(sheet, 8, 3)?.style, 2);
  history(wb, setCellInput(wb, 0, 8, 3, ''));
  assert.equal(getCell(sheet, 8, 3), undefined);
});

function history(wb: Workbook, ops: Op[]): void {
  for (const op of ops) applyOp(wb, op);
}

test('commands: pasted text with short rows or nothing in it leaves other cells alone', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  history(wb, setCellInput(wb, 0, 20, 1, '8'));
  history(wb, pasteAt(wb, 0, { top: 20, left: 0, bottom: 20, right: 0 }, 'x\ny\tz').ops);
  assert.deepEqual([getCell(sheet, 20, 0)?.value, getCell(sheet, 20, 1)?.value, getCell(sheet, 21, 1)?.value], ['x', 8, 'z']);
  assert.deepEqual(pasteAt(wb, 0, { top: 20, left: 0, bottom: 20, right: 0 }, '').ops, []);
  // Our own copy matches the clipboard text whatever line breaks it comes back with.
  const copied = copyRange(wb, 0, { top: 0, left: 0, bottom: 1, right: 1 });
  const pasted = pasteAt(wb, 0, { top: 30, left: 0, bottom: 30, right: 0 }, copied.text.replace(/\n/g, '\r\n'), copied);
  history(wb, pasted.ops);
  assert.equal(f(sheet, 31, 1)?.text, 'A32*2');
  // Pasting at the last column drops what does not fit.
  assert.deepEqual(pasteAt(wb, 0, { top: 0, left: 16383, bottom: 0, right: 16383 }, 'a\tb\tc').range, { top: 0, left: 16383, bottom: 0, right: 16383 });
});

test('styles: equal formats are found whatever objects describe them', async () => {
  const wb = await workbook();
  const styles = wb.styles;
  styles.cellXfs.push({ numFmtId: 14, fontId: 0, fillId: 0, borderId: 0, xfId: 0, quotePrefix: false, alignment: { horizontal: 'center' } });
  const existing = styles.cellXfs.length - 1;
  const source = styles.cellXfs.findIndex((format) => format.numFmtId === 2);
  assert.equal(ensureCellFormat(styles, source, 14), existing);
  assert.equal(ensureNumberFormat(styles, 'General'), 0);
  assert.equal(ensureNumberFormat(styles, 'general'), 0);
  while (styles.cellXfs.length < 65430) styles.cellXfs.push(styles.cellXfs[0]);
  assert.throws(() => ensureCellFormat(styles, source, 22), /too many cell formats/);
});

test('history: random sequences undo back to the start and redo forward to the end', async () => {
  let seed = 20240315;
  const random = (n: number) => {
    seed = (Math.imul(seed, 1103515245) + 12345) & 0x7fffffff;
    return seed % n;
  };
  const inputs = ['1', '2,5', 'metin', '=A1+B2', '=SUM(Veri!A1:A3)', '', '15%', 'x y', '=Özet!A1*2'];
  for (let round = 0; round < 60; round++) {
    const wb = await workbook();
    const hist = new History<number>(wb);
    const states = [picture(wb)];
    let copied = copyRange(wb, 0, { top: 0, left: 0, bottom: 1, right: 1 });
    const steps = 5 + random(20);
    for (let i = 0; i < steps; i++) {
      const sheet = random(2);
      const row = random(12);
      const col = random(6);
      let ops: Op[] = [];
      switch (random(8)) {
        case 0:
        case 1:
          ops = setCellInput(wb, sheet, row, col, inputs[random(inputs.length)]);
          break;
        case 2:
          ops = clearRange(wb, sheet, { top: row, left: col, bottom: row + random(3), right: col + random(3) });
          break;
        case 3:
          copied = copyRange(wb, sheet, { top: row, left: col, bottom: row + random(3), right: col + random(2) });
          break;
        case 4:
          ops = pasteAt(wb, sheet, { top: row, left: col, bottom: row, right: col }, copied.text, random(2) ? copied : undefined).ops;
          break;
        case 5:
          ops = [{ type: random(2) ? 'insert' : 'delete', sheet, axis: 'row', at: row, count: 1 + random(2) }];
          break;
        case 6:
          ops = [{ type: random(2) ? 'insert' : 'delete', sheet, axis: 'col', at: col, count: 1 + random(2) }];
          break;
        default:
          ops = [...setCellInput(wb, sheet, row, col, 'a'), { type: 'insert', sheet, axis: 'row', at: row, count: 1 }, ...setCellInput(wb, sheet, row, col, 'b')];
      }
      const before = picture(wb);
      try {
        hist.apply(ops, i, i);
      } catch (error) {
        // Refused steps (array formulas) must leave the workbook as it was.
        assert.ok(error instanceof RangeError, String(error));
        assert.equal(picture(wb), before, `round ${round} step ${i}: a refused step changed the workbook`);
        continue;
      }
      if (picture(wb) !== before) states.push(picture(wb));
      else if (hist.canUndo && ops.length > 0) states.push(picture(wb));
    }
    const end = picture(wb);
    let undone = 0;
    while (hist.undo()) undone++;
    assert.equal(picture(wb), states[0], `round ${round}: undo did not return to the start`);
    let redone = 0;
    while (hist.redo()) redone++;
    assert.equal(redone, undone);
    assert.equal(picture(wb), end, `round ${round}: redo did not return to the end`);
  }
});

test('commands: a dependant of a shared formula is edited as its formula', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  assert.deepEqual([1, 2, 3].map((row) => editTextAt(wb, sheet, row, 1)), ['=A2*2', '=A3*2', '=A4*2']);
  assert.equal(editTextAt(wb, wb.sheets[1], 2, 0), '=Veri!A4');
  // Writing the edit text back keeps it a formula.
  applyOp(wb, setCellInput(wb, 0, 2, 1, editTextAt(wb, sheet, 2, 1))[0]);
  assert.deepEqual(f(sheet, 2, 1), { text: 'A3*2', type: 'normal' });
});

test('commands: pasting refuses text it could only paste part of, and skips cells under a merge', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  const before = picture(wb);
  assert.throws(() => pasteAt(wb, 0, { top: 0, left: 0, bottom: 0, right: 0 }, 'x\t'.repeat(1000001)), /Cannot paste more than 1000000 cells/);
  assert.equal(picture(wb), before);

  // D1:E2 is merged: D1 is written, E1, D2 and E2 are not.
  history(wb, pasteAt(wb, 0, { top: 0, left: 2, bottom: 0, right: 2 }, 'a\tb\tc\td\ne\tf\tg\th').ops);
  assert.deepEqual(
    [[0, 2], [0, 3], [0, 4], [0, 5], [1, 2], [1, 3], [1, 4], [1, 5]].map(([row, col]) => getCell(sheet, row, col)?.value),
    ['a', 'b', undefined, 'd', 'e', undefined, undefined, 'h'],
  );
});

test('ops: formatting the cell that holds a shared formula keeps the group shared', async () => {
  const wb = await workbook();
  const sheet = wb.sheets[0];
  const master = getCell(sheet, 0, 1)!;
  applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 1, cell: { ...master, style: 2 } }] });
  assert.deepEqual([f(sheet, 0, 1)?.type, f(sheet, 1, 1), f(sheet, 3, 1)], ['shared', { text: '', type: 'shared', si: 0 }, { text: '', type: 'shared', si: 0 }]);
  // Only the row of the changed cell is marked for rewriting.
  assert.deepEqual([...(sheet.changedRows || [])], [0]);
});

test('ops: links in other sheets follow the cells they point at, and come back on undo', async () => {
  const wb = await workbook();
  const other = wb.sheets[1];
  other.hyperlinks = [
    { range: { top: 0, left: 0, bottom: 0, right: 0 }, location: 'Veri!A3' },
    { range: { top: 1, left: 0, bottom: 1, right: 0 }, location: "'Özet'!B2" },
    { range: { top: 2, left: 0, bottom: 2, right: 0 }, target: 'https://example.com/' },
  ];
  const before = picture(wb);
  const linksBefore = other.hyperlinks;
  const undo = applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 1, count: 2 });
  assert.deepEqual(other.hyperlinks.map((link) => link.location), ['Veri!A5', "'Özet'!B2", undefined]);
  assert.equal(other.hyperlinksModified, true);
  assert.equal(linksBefore[0].location, 'Veri!A3', 'the list held for undo is not changed in place');
  const redo = applyOp(wb, undo);
  assert.equal(picture(wb), before);
  applyOp(wb, redo);
  assert.equal(other.hyperlinks[0].location, 'Veri!A5');
});

test('review: a typed formula that cannot be read is refused', async () => {
  const wb = await workbook();
  for (const text of ['=1+', '=SUM(A1;A2)', '=A1*1,5']) {
    assert.throws(() => setCellInput(wb, 0, 20, 0, text), /The formula cannot be read/);
  }
  assert.equal(setCellInput(wb, 0, 20, 0, '=SUM(A1,A2)*2').length, 1);
});
