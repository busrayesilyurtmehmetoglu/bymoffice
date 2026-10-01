import assert from 'node:assert/strict';
import { test } from 'node:test';
import { LOCALE_TR, cellDisplayText, getCell, open, type Cell, type Sheet, type Workbook } from '../src/index';
import { clearRange, copyRange, editTextAt, pasteAt, setCellInput } from '../src/ops/commands';
import { History, applyOp, ensureCellFormat, ensureNumberFormat, type Op } from '../src/ops/ops';
import { buildXlsx, worksheetXml } from './fixtures';

const num = (r: string, v: number, s = '') => `<c r="${r}"${s}><v>${v}</v></c>`;
const formula = (r: string, f: string, v = 0, attrs = '') => `<c r="${r}"><f${attrs}>${f}</f><v>${v}</v></c>`;

/** Two sheets: data with formulas, merges, a link and a shared formula; and a sheet that refers to the first. */
async function workbook(): Promise<Workbook> {
  const file = await open(
    await buildXlsx({
      workbookTail: '<definedNames><definedName name="Toplam">Veri!$B$2:$B$4</definedName><definedName name="Yerel" localSheetId="0">$A$5</definedName></definedNames>',
      sheets: [
        {
          name: 'Veri',
          links: { rId1: 'https://example.com/' },
          xml: worksheetXml(
            '<cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData>' +
              `<row r="1">${num('A1', 1)}${num('B1', 10)}${formula('C1', 'A1+B1', 11, ' t="shared" ref="C1:C3" si="0"')}</row>` +
              `<row r="2" ht="30" customHeight="1">${num('A2', 2)}${num('B2', 20)}<c r="C2"><f t="shared" si="0"/><v>22</v></c></row>` +
              `<row r="3">${num('A3', 3)}${num('B3', 30)}<c r="C3"><f t="shared" si="0"/><v>33</v></c></row>` +
              `<row r="5">${formula('A5', 'SUM(A1:A3)', 6)}${formula('B5', 'SUM($B$1:$B$3)', 60)}${formula('C5', 'A1*2', 2)}</row>` +
              '</sheetData><mergeCells><mergeCell ref="A7:C8"/><mergeCell ref="E1:F1"/></mergeCells>' +
              '<hyperlinks><hyperlink ref="A3" r:id="rId1"/></hyperlinks>',
          ),
        },
        { name: 'Özet', xml: worksheetXml(`<sheetData><row r="1">${formula('A1', 'Veri!A3+1', 4)}${formula('B1', 'SUM(Veri!A1:A3)', 6)}${formula('C1', 'A1', 4)}</row></sheetData>`) },
      ],
    }),
  );
  return file.workbook();
}

const formulaAt = (sheet: Sheet, row: number, col: number) => getCell(sheet, row, col)?.formula?.text;
const valueAt = (sheet: Sheet, row: number, col: number) => getCell(sheet, row, col)?.value;

/** A picture of everything operations may touch, for comparing states. */
function picture(wb: Workbook): unknown {
  return {
    names: wb.definedNames.map((n) => n.value),
    sheets: wb.sheets.map((sheet) => ({
      rowCount: sheet.rowCount,
      colCount: sheet.colCount,
      merges: sheet.merges,
      links: sheet.hyperlinks.map((h) => h.range),
      columns: sheet.columns.map((c) => [c.min, c.max]),
      edits: sheet.structuralEdits,
      rows: [...sheet.rows.entries()]
        .sort((a, b) => a[0] - b[0])
        .map(([index, row]) => [index, row.height, row.cells.map((cell, col) => [col, cell.kind, cell.value, cell.style, cell.formula]).filter(Boolean)]),
    })),
  };
}

test('ops: setCells applies, grows the used range and inverts exactly', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const cell: Cell = { kind: 'string', value: 'yeni', style: 0 };
  const inverse = applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell }, { row: 20, col: 9, cell }, { row: 1, col: 1, cell: undefined }, { row: 50, col: 50, cell: undefined }] });
  const sheet = wb.sheets[0];
  assert.deepEqual([valueAt(sheet, 0, 0), valueAt(sheet, 20, 9), getCell(sheet, 1, 1)], ['yeni', 'yeni', undefined]);
  assert.deepEqual([sheet.rowCount, sheet.colCount, sheet.modified], [21, 10, true]);
  // Removing a cell that is not there is not a change and creates no row.
  assert.equal(sheet.rows.has(50), false);
  assert.equal(sheet.rows.get(1)?.height, 30, 'row properties survive a cell change');

  applyOp(wb, inverse);
  assert.deepEqual(picture(wb), before);
  assert.throws(() => applyOp(wb, { type: 'setCells', sheet: 0, cells: [{ row: 1048576, col: 0, cell }] }), RangeError);
  assert.throws(() => applyOp(wb, { type: 'setCells', sheet: 9, cells: [] }), RangeError);
});

test('ops: inserting rows moves cells, formulas, merges, links and names', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const inverse = applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 1, count: 2 });
  const [data, summary] = wb.sheets;

  assert.deepEqual([valueAt(data, 0, 0), getCell(data, 1, 0), getCell(data, 2, 0), valueAt(data, 3, 0), valueAt(data, 4, 0)], [1, undefined, undefined, 2, 3]);
  assert.equal(data.rows.get(3)?.height, 30);
  // Shared formulas are written out, then adjusted like any other.
  assert.deepEqual([formulaAt(data, 0, 2), formulaAt(data, 3, 2), formulaAt(data, 4, 2)], ['A1+B1', 'A4+B4', 'A5+B5']);
  assert.equal(getCell(data, 3, 2)?.formula?.type, 'normal');
  assert.deepEqual([formulaAt(data, 6, 0), formulaAt(data, 6, 1), formulaAt(data, 6, 2)], ['SUM(A1:A5)', 'SUM($B$1:$B$5)', 'A1*2']);
  assert.deepEqual(data.merges, [{ top: 8, left: 0, bottom: 9, right: 2 }, { top: 0, left: 4, bottom: 0, right: 5 }]);
  assert.deepEqual(data.hyperlinks[0].range, { top: 4, left: 0, bottom: 4, right: 0 });
  assert.equal(data.rowCount, 7);
  assert.deepEqual(data.structuralEdits, [{ kind: 'insert', axis: 'row', at: 1, count: 2 }]);
  // The other sheet's references into this one follow; its own do not.
  assert.deepEqual([formulaAt(summary, 0, 0), formulaAt(summary, 0, 1), formulaAt(summary, 0, 2)], ['Veri!A5+1', 'SUM(Veri!A1:A5)', 'A1']);
  assert.deepEqual(wb.definedNames.map((n) => n.value), ['Veri!$B$4:$B$6', '$A$7']);

  const redo = applyOp(wb, inverse);
  assert.deepEqual(picture(wb), before);
  applyOp(wb, redo);
  assert.equal(formulaAt(wb.sheets[0], 6, 0), 'SUM(A1:A5)');
  assert.equal(formulaAt(wb.sheets[1], 0, 0), 'Veri!A5+1');
});

test('ops: deleting rows removes cells and turns references to them into #REF!', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const inverse = applyOp(wb, { type: 'delete', sheet: 0, axis: 'row', at: 1, count: 2 });
  const [data, summary] = wb.sheets;
  assert.deepEqual([valueAt(data, 0, 0), getCell(data, 1, 0), formulaAt(data, 2, 0)], [1, undefined, 'SUM(A1:A1)']);
  assert.equal(formulaAt(data, 2, 1), 'SUM($B$1:$B$1)');
  assert.deepEqual(data.merges, [{ top: 4, left: 0, bottom: 5, right: 2 }, { top: 0, left: 4, bottom: 0, right: 5 }]);
  assert.deepEqual(data.hyperlinks, []);
  assert.equal(data.rowCount, 3);
  assert.equal(formulaAt(summary, 0, 0), 'Veri!#REF!+1');
  assert.equal(formulaAt(summary, 0, 1), 'SUM(Veri!A1:A1)');
  applyOp(wb, inverse);
  assert.deepEqual(picture(wb), before);
});

test('ops: inserting and deleting columns', async () => {
  const wb = await workbook();
  const before = picture(wb);
  const inverse = applyOp(wb, { type: 'insert', sheet: 0, axis: 'col', at: 1, count: 1 });
  const data = wb.sheets[0];
  assert.deepEqual([valueAt(data, 0, 0), getCell(data, 0, 1), valueAt(data, 0, 2), formulaAt(data, 0, 3)], [1, undefined, 10, 'A1+C1']);
  assert.equal(formulaAt(data, 4, 2), 'SUM($C$1:$C$3)');
  // The column that had a width (B:C) is now C:D; a merge that spans the insert point grows.
  assert.deepEqual(data.columns.map((c) => [c.min, c.max]), [[2, 3]]);
  assert.deepEqual(data.merges, [{ top: 6, left: 0, bottom: 7, right: 3 }, { top: 0, left: 5, bottom: 0, right: 6 }]);
  assert.equal(data.colCount, 4);
  assert.equal(wb.definedNames[0].value, 'Veri!$C$2:$C$4');
  applyOp(wb, inverse);
  assert.deepEqual(picture(wb), before);

  const undoDelete = applyOp(wb, { type: 'delete', sheet: 0, axis: 'col', at: 0, count: 1 });
  assert.deepEqual([valueAt(data, 0, 0), formulaAt(data, 0, 1), formulaAt(data, 4, 0)], [10, '#REF!+A1', 'SUM($A$1:$A$3)']);
  assert.deepEqual(data.columns.map((c) => [c.min, c.max]), [[0, 1]]);
  assert.equal(formulaAt(wb.sheets[1], 0, 0), 'Veri!#REF!+1');
  applyOp(wb, undoDelete);
  assert.deepEqual(picture(wb), before);
});

test('ops: structural operations are validated', async () => {
  const wb = await workbook();
  const before = picture(wb);
  for (const op of [
    { type: 'insert', sheet: 0, axis: 'row', at: -1, count: 1 },
    { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 0 },
    { type: 'insert', sheet: 0, axis: 'row', at: 0.5, count: 1 },
    { type: 'delete', sheet: 0, axis: 'col', at: 16383, count: 2 },
    { type: 'insert', sheet: 5, axis: 'row', at: 0, count: 1 },
    // Would push the used cells past the last row.
    { type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1048572 },
  ] as Op[]) {
    assert.throws(() => applyOp(wb, op), RangeError, JSON.stringify(op));
  }
  assert.deepEqual(picture(wb), before);
  // Inserting below the used cells is fine however many rows it is.
  applyOp(wb, { type: 'insert', sheet: 0, axis: 'row', at: 100, count: 1048000 });
});

test('history: undo and redo across cell and structural steps', async () => {
  const wb = await workbook();
  const history = new History<string>(wb);
  const start = picture(wb);
  assert.equal(history.canUndo, false);
  assert.equal(history.undo(), undefined);

  history.apply(setCellInput(wb, 0, 0, 0, '100'), 'a0', 'a1');
  const afterEdit = picture(wb);
  history.apply([{ type: 'insert', sheet: 0, axis: 'row', at: 0, count: 1 }], 'b0', 'b1');
  history.apply(setCellInput(wb, 0, 0, 0, 'başlık'), 'c0', 'c1');
  const end = picture(wb);
  assert.equal(valueAt(wb.sheets[0], 1, 0), 100);

  assert.deepEqual(history.undo(), { meta: 'c0' });
  assert.deepEqual(history.undo(), { meta: 'b0' });
  assert.deepEqual(picture(wb), afterEdit);
  assert.deepEqual(history.undo(), { meta: 'a0' });
  assert.deepEqual(picture(wb), start);
  assert.equal(history.canUndo, false);

  assert.deepEqual(history.redo(), { meta: 'a1' });
  assert.deepEqual(history.redo(), { meta: 'b1' });
  assert.deepEqual(history.redo(), { meta: 'c1' });
  assert.deepEqual(picture(wb), end);
  assert.equal(history.redo(), undefined);

  // A new change drops what could be redone.
  history.undo();
  history.apply(setCellInput(wb, 0, 5, 5, 'x'), 'd0', 'd1');
  assert.equal(history.canRedo, false);
});

test('history: a step that fails half-way leaves nothing behind, and old steps fall off the end', async () => {
  const wb = await workbook();
  const history = new History<number>(wb, 3);
  const start = picture(wb);
  assert.throws(
    () =>
      history.apply(
        [
          { type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'number', value: 999, style: 0 } }] },
          { type: 'insert', sheet: 0, axis: 'row', at: -5, count: 1 },
        ],
        0,
        0,
      ),
    RangeError,
  );
  assert.deepEqual(picture(wb), start);
  assert.equal(history.canUndo, false);

  for (let i = 0; i < 5; i++) history.apply(setCellInput(wb, 0, 30 + i, 0, String(i)), i, i);
  let undone = 0;
  while (history.undo()) undone++;
  assert.equal(undone, 3);
  assert.equal(valueAt(wb.sheets[0], 31, 0), 1);
});

test('commands: typed input becomes values and picks up formats', async () => {
  const wb = await workbook();
  const history = new History(wb);
  const type = (row: number, col: number, text: string) => history.apply(setCellInput(wb, 0, row, col, text, { locale: LOCALE_TR }), undefined, undefined);
  const sheet = wb.sheets[0];
  const shown = (row: number, col: number) => cellDisplayText(wb, getCell(sheet, row, col), LOCALE_TR);

  type(10, 0, '12,5');
  type(10, 1, '%15');
  type(10, 2, '15%');
  type(10, 3, '15.03.2024');
  type(10, 4, '=A11*2');
  type(10, 5, 'metin');
  type(10, 6, "'0123");
  assert.equal(valueAt(sheet, 10, 0), 12.5);
  assert.equal(valueAt(sheet, 10, 2), 0.15);
  assert.equal(shown(10, 2), '15%');
  assert.equal(valueAt(sheet, 10, 3), 45366);
  assert.equal(shown(10, 3), '15.03.2024');
  assert.deepEqual(getCell(sheet, 10, 4)?.formula, { text: 'A11*2', type: 'normal' });
  assert.deepEqual([valueAt(sheet, 10, 5), valueAt(sheet, 10, 6), getCell(sheet, 10, 6)?.kind], ['metin', '0123', 'string']);

  // What the editor shows turns back into the same cell.
  for (const col of [0, 2, 3, 4, 5, 6]) {
    const before = getCell(sheet, 10, col);
    type(10, col, editTextAt(wb, sheet, 10, col, { locale: LOCALE_TR }));
    const after = getCell(sheet, 10, col);
    assert.deepEqual([after?.kind, after?.value, after?.formula, after?.style], [before?.kind, before?.value, before?.formula, before?.style], `column ${col}`);
  }

  // Emptying a formatted cell keeps the format; emptying a plain one removes it.
  type(10, 3, '');
  assert.deepEqual([getCell(sheet, 10, 3)?.kind, getCell(sheet, 10, 3)?.style !== 0], ['blank', true]);
  type(10, 5, '');
  assert.equal(getCell(sheet, 10, 5), undefined);
  assert.deepEqual(setCellInput(wb, 0, 200, 200, ''), []);
  // A date typed into a cell that already has a format does not replace it.
  const percent = getCell(sheet, 10, 2)!.style;
  type(10, 2, '15.03.2024');
  assert.equal(getCell(sheet, 10, 2)?.style, percent);
});

test('commands: clearing, copying and pasting', async () => {
  const wb = await workbook();
  const history = new History(wb);
  const sheet = wb.sheets[0];

  const copied = copyRange(wb, 0, { top: 0, left: 0, bottom: 2, right: 2 });
  assert.equal(copied.text, '1\t10\t11\n2\t20\t22\n3\t30\t33');
  assert.match(copied.html, /^<table><tr><td>1<\/td>/);

  // Pasting our own copy carries formulas along, moved to the new place.
  const pasted = pasteAt(wb, 0, { top: 10, left: 1, bottom: 10, right: 1 }, copied.text, copied);
  history.apply(pasted.ops, undefined, undefined);
  assert.deepEqual(pasted.range, { top: 10, left: 1, bottom: 12, right: 3 });
  assert.deepEqual([valueAt(sheet, 10, 1), valueAt(sheet, 12, 2)], [1, 30]);
  assert.deepEqual([formulaAt(sheet, 10, 3), formulaAt(sheet, 11, 3)], ['B11+C11', 'B12+C12']);

  // Text from elsewhere is parsed field by field.
  const foreign = pasteAt(wb, 0, { top: 20, left: 0, bottom: 20, right: 0 }, '5\tabc\r\n=A21*2\t7,5\r\n', undefined, { locale: LOCALE_TR });
  history.apply(foreign.ops, undefined, undefined);
  assert.deepEqual([valueAt(sheet, 20, 0), valueAt(sheet, 20, 1), formulaAt(sheet, 21, 0), valueAt(sheet, 21, 1)], [5, 'abc', 'A21*2', 7.5]);
  assert.deepEqual(foreign.range, { top: 20, left: 0, bottom: 21, right: 1 });

  // A selection that is a multiple of the block is filled with repeats.
  const one = copyRange(wb, 0, { top: 0, left: 0, bottom: 0, right: 0 });
  const tiled = pasteAt(wb, 0, { top: 30, left: 0, bottom: 31, right: 2 }, one.text, one);
  history.apply(tiled.ops, undefined, undefined);
  assert.deepEqual([valueAt(sheet, 30, 0), valueAt(sheet, 31, 2)], [1, 1]);
  // Whole columns are not tiled.
  assert.deepEqual(pasteAt(wb, 0, { top: 0, left: 8, bottom: 1048575, right: 8 }, one.text, one).range, { top: 0, left: 8, bottom: 0, right: 8 });

  history.apply(clearRange(wb, 0, { top: 30, left: 0, bottom: 1048575, right: 16383 }), undefined, undefined);
  assert.equal(getCell(sheet, 30, 0), undefined);
  assert.equal(valueAt(sheet, 20, 0), 5);
  assert.deepEqual(clearRange(wb, 0, { top: 500, left: 0, bottom: 600, right: 5 }), []);

  // Copying whole columns stops at the used cells.
  const column = copyRange(wb, 0, { top: 0, left: 0, bottom: 1048575, right: 0 });
  assert.equal(column.cells.length, sheet.rowCount);

  while (history.undo());
  assert.equal(getCell(sheet, 10, 1), undefined);
  assert.equal(getCell(sheet, 20, 0), undefined);
});

test('styles: formats are added once and found again', async () => {
  const wb = await workbook();
  const styles = wb.styles;
  const count = styles.cellXfs.length;
  const percent = ensureCellFormat(styles, 0, 9);
  assert.equal(percent, count);
  assert.equal(ensureCellFormat(styles, 0, 9), percent);
  assert.equal(ensureCellFormat(styles, percent, 9), percent);
  assert.equal(ensureCellFormat(styles, 0, 0), 0);
  assert.equal(styles.modified, true);
  const custom = ensureNumberFormat(styles, '#,##0.00 "₺"');
  assert.ok(custom >= 164);
  assert.equal(ensureNumberFormat(styles, '#,##0.00 "₺"'), custom);
  assert.equal(ensureNumberFormat(styles, '0.0000'), custom + 1);
});
