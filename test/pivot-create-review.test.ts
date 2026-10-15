// What the review of making pivot tables and changing their fields found: date labels keep their
// own format, a refused change leaves the style sheet as it was, the weight of an undo step, the
// panel does not look through the package on every edit, the bridge takes filters with an item,
// default names of values avoid the names of fields, a filter shows one item or hides some, names
// are written as the names of fields are, a defined name of whole columns.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { ZipArchive, type Cell } from '../src/index';
import { History, ensureCellFormat, ensureNumberFormat, type Op } from '../src/ops/ops';
import { createPivot, createPivotProblem } from '../src/ops/pivot-create';
import { pivotFields, setPivotFields, setPivotFieldsProblem } from '../src/ops/pivot-fields';
import { PIVOT_STATS, refreshPivot } from '../src/ops/pivots';
import { tableState } from '../src/model/sheet/pivot';
import { readMadePivot } from '../src/formats/xlsx/pivot';
import { madeTableXml } from '../src/formats/xlsx/pivot-create';
import { bookWith, load, pivotBook, same, save, setData, text, type Datum, type Loaded } from './pivot-fixtures';

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);
const DATES: Datum[][] = [['D', 'S'], [{ date: 45000 }, 1], [{ date: 45001 }, 2], [null, 3], [{ date: 45002 }, 4]];

/** The label cells of column A of a sheet below a row, with the number format of each. */
function labels(doc: Loaded, sheet: number, from: number): [unknown, number][] {
  const out: [unknown, number][] = [];
  const s = doc.wb.sheets[sheet];
  for (const [r, row] of [...s.rows].sort((a, b) => a[0] - b[0])) {
    if (r < from || !row.cells[0]) continue;
    out.push([row.cells[0].value, doc.wb.styles.cellXfs[row.cells[0].style].numFmtId]);
  }
  return out;
}

test('a refreshed table of the file keeps the format its date labels have, beside a "(blank)" label', async () => {
  const doc = await load(await pivotBook({ data: DATES, rows: [0], values: [{ fld: 1 }] }));
  apply(doc, refreshPivot(doc.wb, 0));
  const fmt = ensureCellFormat(doc.wb.styles, 0, ensureNumberFormat(doc.wb.styles, 'dd.mm.yyyy'));
  const sheet = doc.wb.sheets[1];
  const cells: { row: number; col: number; cell: Cell }[] = [];
  for (const [r, row] of sheet.rows) {
    const c = row.cells[0];
    if (c && c.kind === 'number' && r > 2) cells.push({ row: r, col: 0, cell: { ...c, style: fmt } });
  }
  apply(doc, [{ type: 'setCells', sheet: 1, cells }]);
  const custom = doc.wb.styles.cellXfs[fmt].numFmtId;
  setData(doc, 1, 1, 10);
  apply(doc, refreshPivot(doc.wb, 0));
  assert.deepEqual(labels(doc, 1, 3), [
    [45000, custom],
    [45001, custom],
    [45002, custom],
    ['(blank)', 0],
    ['Grand Total', 0],
  ]);
});

test('a table made from dates shows them with the format of the source; saved, opened and refreshed it changes no cell', async () => {
  const doc = await bookWith(DATES);
  const fmt = ensureCellFormat(doc.wb.styles, 0, ensureNumberFormat(doc.wb.styles, 'dd.mm.yyyy'));
  const cells: { row: number; col: number; cell: Cell }[] = [];
  for (let r = 1; r < DATES.length; r++) if (DATES[r][0] !== null) cells.push({ row: r, col: 0, cell: { kind: 'number', value: (DATES[r][0] as { date: number }).date, style: fmt } });
  apply(doc, [{ type: 'setCells', sheet: 0, cells }]);
  const custom = doc.wb.styles.cellXfs[fmt].numFmtId;
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 4, right: 1 } } }).ops);
  // In the order the source has them.
  assert.deepEqual(labels(doc, 1, 3).slice(0, 4), [
    [45000, custom],
    [45001, custom],
    ['(blank)', 0],
    [45002, custom],
  ]);
  const again = await load(await save(doc));
  assert.deepEqual(refreshPivot(again.wb, 0).filter((op) => op.type === 'setCells'), []);
});

test('a refused or only asked-for table and change leave the style sheet as it was; a made one adds its formats as a step that is undone', async () => {
  const doc = await bookWith([['Ad', 'Tarih', 'Tutar'], ['a', { date: 45000 }, 1], ['b', { date: 45001 }, 2]]);
  const base = await save(doc);
  const styles = () => [doc.wb.styles.cellXfs.length, doc.wb.styles.numFmts.size, !!doc.wb.styles.modified];
  const was = styles();
  const source = { sheet: 0, range: { top: 0, left: 0, bottom: 2, right: 2 } };
  assert.ok(createPivotProblem(doc.wb, { source, target: { sheet: 0, row: 1, col: 1 }, rows: [1], values: [{ field: 2, numFmt: '#,##0.000' }] }));
  assert.equal(createPivotProblem(doc.wb, { source, rows: [1], values: [{ field: 2, numFmt: '#,##0.000' }] }), undefined);
  assert.deepEqual(styles(), was);
  assert.ok(same(await save(doc), base), 'nothing was changed');
  apply(doc, createPivot(doc.wb, { source, rows: [0], values: [{ field: 2 }] }).ops);
  assert.ok(setPivotFieldsProblem(doc.wb, 0, { rows: [], values: [{ field: 2, numFmt: '0.0000' }] }));
  assert.equal(setPivotFieldsProblem(doc.wb, 0, { values: [{ field: 2, numFmt: '0.0000' }] }), undefined);
  apply(doc, setPivotFields(doc.wb, 0, { rows: [1], values: [{ field: 2, numFmt: '0.000' }] }));
  assert.notDeepEqual(styles(), was);
  doc.history.undo();
  doc.history.undo();
  assert.deepEqual(styles(), was);
  assert.ok(same(await save(doc), base), 'made with a format and a date, undone: the bytes of the file');
});

test('the weight of a step that changes a large table counts what it keeps: the history drops old steps', async () => {
  const doc = await bookWith([['K', 'V'], ['a', 1]]);
  const cells: { row: number; col: number; cell: Cell }[] = [];
  const n = 40000;
  for (let r = 1; r <= n; r++) {
    cells.push({ row: r, col: 0, cell: { kind: 'string', value: `k${r}`, style: 0 } });
    cells.push({ row: r, col: 1, cell: { kind: 'number', value: r, style: 0 } });
  }
  apply(doc, [{ type: 'setCells', sheet: 0, cells }]);
  const history = new History(doc.wb);
  history.apply(createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: n, right: 1 } } }).ops, undefined, undefined);
  for (let i = 0; i < 12; i++) history.apply(setPivotFields(doc.wb, 0, { sort: [{ field: 0, order: i % 2 ? 'descending' : 'ascending' }] }), undefined, undefined);
  let steps = 0;
  while (history.undo()) steps++;
  assert.ok(steps < 12, `${steps} steps of a table of ${n} items were kept`);
});

test('the fields of a table are listed without looking through the package when asked not to', async () => {
  const doc = await load(await pivotBook({ data: DATES, rows: [0], values: [{ fld: 1 }] }));
  const scans = PIVOT_STATS.scans;
  pivotFields(doc.wb, 0, { look: false });
  setData(doc, 1, 1, 9);
  pivotFields(doc.wb, 0, { look: false });
  assert.equal(PIVOT_STATS.scans, scans);
  pivotFields(doc.wb, 0);
  assert.equal(PIVOT_STATS.scans, scans + 1);
});

test('a default name of a value is not the name of a field: it gets a number', async () => {
  const doc = await bookWith([['Region', 'Sum of Sales', 'Sales'], ['a', 1, 2], ['b', 3, 4]]);
  const source = { sheet: 0, range: { top: 0, left: 0, bottom: 2, right: 2 } };
  apply(doc, createPivot(doc.wb, { source, rows: [0], values: [{ field: 1 }] }).ops);
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2 }] }));
  assert.deepEqual(pivotFields(doc.wb, 0).values.map((v) => v.name), ['Sum of Sales2']);
});

test('a filter shows one item or hides some, never both', async () => {
  const doc = await load(await pivotBook({ data: DATES, rows: [1], values: [{ fld: 1 }], at: { row: 4, col: 0 } }));
  apply(doc, setPivotFields(doc.wb, 0, { filters: [0] }));
  apply(doc, setPivotFields(doc.wb, 0, { hidden: [{ field: 0, items: [0] }] }));
  apply(doc, setPivotFields(doc.wb, 0, { filters: [{ field: 0, item: 1 }] }));
  let info = pivotFields(doc.wb, 0);
  assert.deepEqual(info.filters, [{ field: 0, item: 1 }]);
  assert.ok(info.fields[0].items!.every((item) => !item.hidden), 'one item shown: none hidden');
  apply(doc, setPivotFields(doc.wb, 0, { hidden: [{ field: 0, items: [2] }] }));
  info = pivotFields(doc.wb, 0);
  assert.deepEqual(info.filters, [{ field: 0 }], 'items hidden: no one item shown');
  const part = await text(await save(doc), 'xl/pivotTables/pivotTable1.xml');
  assert.doesNotMatch(part, /<pageField[^>]* item="/);
});

test('names of values and of tables are written as the names of fields are (escapes of ST_Xstring), and read back', async () => {
  const xml = madeTableXml({ name: 'T_x0041_', fields: ['a', 'b'], rows: [0], cols: [], pages: [], values: [{ fld: 1, subtotal: 'sum', name: 'v_x0042_\u0001' }], at: { row: 2, col: 0 }, rowGrandTotals: true, colGrandTotals: true });
  assert.match(xml, /name="T_x005F_x0041_"/);
  assert.match(xml, /name="v_x005F_x0042__x0001_"/);
  const cacheXml = '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheSource type="worksheet"><worksheetSource ref="A1:B3" sheet="S"/></cacheSource><cacheFields count="2"><cacheField name="a"/><cacheField name="b"/></cacheFields></pivotCacheDefinition>';
  const doc = await bookWith([['a', 'b'], ['x', 1]]);
  const read = readMadePivot(cacheXml, xml, doc.wb.sheets[0], false);
  assert.equal(read.table.name, 'T_x0041_');
  assert.equal(read.table.dataFields[0].name, 'v_x0042_\u0001');
});

test('a defined name of whole columns is a source, cut to the cells it holds', async () => {
  const doc = await bookWith(DATES, '<definedName name="Kolonlar">Sheet1!$A:$B</definedName>');
  const made = createPivot(doc.wb, { source: { name: 'Kolonlar' } });
  apply(doc, made.ops);
  assert.deepEqual(doc.wb.pivots!.tables[0].cache.source, { kind: 'range', sheet: 'Sheet1', ref: 'A1:B5' });
  assert.deepEqual(tableState(doc.wb.pivots!.tables[0]).location.ref.top, 2);
});

test('the bridge takes filters as fields or as fields with the item they show', async () => {
  const { base64Encode, createBridge } = await import('../src/embed/bridge');
  const calls: any[] = [];
  const sent: any[] = [];
  const book = (await bookWith(DATES)).wb;
  const view: any = {
    history: { canUndo: false, canRedo: false },
    activeSheet: 0,
    isEditing: false,
    readOnly: false,
    workbook: book,
    selection: { active: { row: 0, col: 0 }, range: { top: 0, left: 0, bottom: 0, right: 0 } },
    commitEditing: () => undefined,
    commitPending: () => true,
    localizeError: (error: Error) => error.message,
    setReadOnly: () => undefined,
    setLocale: () => undefined,
    destroy: () => undefined,
    insertPivot: (options: unknown) => (calls.push(options), 0),
    setPivotFields: (index: number, change: unknown) => (calls.push(change), true),
    pivotFields: () => ({}),
  };
  const bridge = createBridge({
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      sent.push(JSON.parse(json));
      return true;
    },
    open: async () => ({ format: 'xlsx', kind: 'spreadsheet', save: async () => new Uint8Array([1]), workbook: async () => book }) as any,
    createView: () => view,
    createDocumentView: () => ({ destroy: () => undefined }) as any,
    frame: () => undefined,
    timer: () => undefined,
  });
  await bridge.receive({ type: 'load', id: 'l', data: base64Encode(new Uint8Array([80, 75, 3, 4])) });
  await bridge.receive({ type: 'setPivotFields', id: 's', index: 0, filters: [1, { field: 0, item: 2 }, { field: 3 }, 'x'] });
  assert.deepEqual(sent[sent.length - 1], { type: 'done', id: 's', command: 'setPivotFields', changed: true });
  assert.deepEqual(calls[0].filters, [1, { field: 0, item: 2 }, { field: 3 }, NaN]);
  await bridge.receive({ type: 'insertPivot', id: 'i', source: 'A1:B5', filters: [{ field: 0 }] });
  assert.deepEqual(calls[1].filters, [{ field: 0 }]);
  void ZipArchive;
});
