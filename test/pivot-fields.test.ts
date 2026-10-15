// Changing the fields of pivot tables (src/ops/pivot-fields.ts): of a table of the file and of one
// made here; what is refused; undo; what is saved, read back by the engine and by independent readers.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createBlank, type Cell } from '../src/index';
import { createPivot } from '../src/ops/pivot-create';
import { pivotFields, pivotFieldsProblem, setPivotFields, setPivotFieldsProblem } from '../src/ops/pivot-fields';
import { PIVOT_LIMITS, refreshPivot } from '../src/ops/pivots';
import { tableDef, tableState } from '../src/model/sheet/pivot';
import type { Op } from '../src/ops/ops';
import { bookWith, grid, hasPython, independentPivots, load, pivotBook, same, save, text, type Datum, type Loaded, type PivotSpec } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales', 'Units'],
  ['East', 'A', 10, 1],
  ['West', 'B', 20, 2],
  ['East', 'B', 5, 3],
  ['West', 'A', 7, 4],
  ['East', 'A', 1, 5],
];

const apply = (doc: Loaded, ops: Op[]) => doc.history.apply(ops, undefined, undefined);

/** The body of the first table now, as values. */
function body(doc: Loaded): unknown[][] {
  const table = doc.wb.pivots!.tables[0];
  const ref = tableState(table).location.ref;
  return grid(doc.wb, doc.wb.sheets.indexOf(table.sheet), ref.top, ref.left, ref.bottom, ref.right);
}

async function fileTable(spec: Partial<PivotSpec> = {}): Promise<Loaded> {
  return load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], ...spec }));
}

test('pivotFields lists the fields of a table, where they stand, their items, the values', async () => {
  const doc = await fileTable({ cols: [1], hidden: [{ fld: 1, values: ['B'] }] });
  const info = pivotFields(doc.wb, 0);
  assert.equal(info.refusal, undefined);
  assert.deepEqual(info.rows, [0]);
  assert.deepEqual(info.cols, [1]);
  assert.deepEqual(info.filters, []);
  assert.deepEqual(info.values, [{ field: 2, subtotal: 'sum', name: 'Sum of Sales' }]);
  assert.deepEqual(info.fields.map((f) => f.name), ['Region', 'Product', 'Sales', 'Units']);
  assert.deepEqual(info.fields[1].items, [{ text: 'A', hidden: false }, { text: 'B', hidden: true }]);
  assert.equal(info.fields[2].items, undefined);
  assert.equal(info.rowGrandTotals, true);
});

test('a field moved from the columns to the rows; undone, the table is as it was, and saved the file is byte for byte', async () => {
  const bytes = await pivotBook({ data: DATA, rows: [0], cols: [1], values: [{ fld: 2 }] });
  const doc = await load(bytes);
  apply(doc, refreshPivot(doc.wb, 0));
  const before = body(doc);
  const saved = await save(doc);
  apply(doc, setPivotFields(doc.wb, 0, { rows: [0, 1], cols: [] }));
  assert.deepEqual(body(doc), [
    ['Row Labels', 'Sum of Sales'],
    ['East', 16],
    ['A', 11],
    ['B', 5],
    ['West', 27],
    ['A', 7],
    ['B', 20],
    ['Grand Total', 43],
  ]);
  assert.deepEqual(tableDef(doc.wb.pivots!.tables[0]).rowFields, [0, 1]);
  const changed = await save(doc);
  const again = await load(changed);
  const table = again.wb.pivots!.tables[0];
  assert.deepEqual(table.rowFields, [0, 1]);
  assert.deepEqual(table.colFields, []);
  assert.equal(table.fields[1].axis, 'axisRow');
  assert.deepEqual(refreshPivot(again.wb, 0).filter((op) => op.type === 'setCells'), []);
  if (hasPython()) assert.deepEqual(independentPivots({ moved: changed }), []);
  doc.history.undo();
  assert.deepEqual(body(doc), before);
  assert.ok(same(await save(doc), saved));
  doc.history.undo();
  assert.ok(same(await save(doc), bytes), 'changed, refreshed, both undone: the bytes of the file');
});

test('values: a second value stands in the columns; the function of a value with Excel\'s name takes the name along; a number format', async () => {
  const doc = await fileTable();
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2 }, { field: 3, subtotal: 'average' }] }));
  assert.deepEqual(body(doc), [
    ['Row Labels', 'Sum of Sales', 'Average of Units'],
    ['East', 16, 3],
    ['West', 27, 3],
    ['Grand Total', 43, 3],
  ]);
  assert.deepEqual(tableDef(doc.wb.pivots!.tables[0]).colFields, [-2]);
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2, subtotal: 'max', numFmt: '#,##0.00' }] }));
  const info = pivotFields(doc.wb, 0);
  assert.deepEqual(info.values, [{ field: 2, subtotal: 'max', name: 'Max of Sales', numFmt: '#,##0.00' }]);
  const table = doc.wb.pivots!.tables[0];
  const ref = tableState(table).location.ref;
  const cell = table.sheet.rows.get(ref.top + 1)!.cells[ref.left + 1];
  assert.equal(cell.value, 10);
  assert.equal(doc.wb.styles.cellXfs[cell.style].numFmtId, 4);
  // A name of its own stays when the function changes.
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2, name: 'Satış' }] }));
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2, subtotal: 'min' }] }));
  assert.deepEqual(pivotFields(doc.wb, 0).values.map((v) => v.name), ['Satış']);
  const bytes = await save(doc);
  const again = await load(bytes);
  assert.deepEqual(again.wb.pivots!.tables[0].dataFields, [{ name: 'Satış', fld: 2, subtotal: 'min', numFmtId: 4 }]);
  if (hasPython()) assert.deepEqual(independentPivots({ values: bytes }), []);
});

test('the same value twice gets a number, as Excel names it; a name of a field or twice the same name is refused', async () => {
  const doc = await fileTable();
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 2 }, { field: 2 }] }));
  assert.deepEqual(pivotFields(doc.wb, 0).values.map((v) => v.name), ['Sum of Sales', 'Sum of Sales2']);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { values: [{ field: 2, name: 'units' }] }) || '', /name of a field/);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { values: [{ field: 2, name: 'X' }, { field: 3, name: 'x' }] }) || '', /two values have the name "x"/);
});

test('hidden items are left out, totals included; every item hidden is refused', async () => {
  const doc = await fileTable({ rows: [1] });
  apply(doc, setPivotFields(doc.wb, 0, { hidden: [{ field: 1, items: [0] }] }));
  assert.deepEqual(body(doc), [
    ['Row Labels', 'Sum of Sales'],
    ['B', 25],
    ['Grand Total', 25],
  ]);
  assert.deepEqual(pivotFields(doc.wb, 0).fields[1].items, [{ text: 'A', hidden: true }, { text: 'B', hidden: false }]);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { hidden: [{ field: 1, items: [0, 1] }] }) || '', /every item of the field "Product" would be hidden/);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { hidden: [{ field: 3, items: [0] }] }) || '', /lists no items yet/);
  const bytes = await save(doc);
  if (hasPython()) assert.deepEqual(independentPivots({ hidden: bytes }), []);
  apply(doc, setPivotFields(doc.wb, 0, { hidden: [{ field: 1, items: [] }] }));
  assert.deepEqual(body(doc)[3], ['Grand Total', 43]);
});

test('a filter: the body moves down to make room for it; one item shown; taken out again, its cells are cleared', async () => {
  const doc = await fileTable({ rows: [1], at: { row: 0, col: 0 } });
  apply(doc, refreshPivot(doc.wb, 0));
  apply(doc, setPivotFields(doc.wb, 0, { filters: [0] }));
  assert.deepEqual(tableState(doc.wb.pivots!.tables[0]).location.ref, { top: 2, left: 0, bottom: 5, right: 1 });
  assert.deepEqual(grid(doc.wb, 1, 0, 0, 1, 1), [
    ['Region', '(All)'],
    [undefined, undefined],
  ]);
  apply(doc, setPivotFields(doc.wb, 0, { filters: [{ field: 0, item: 1 }] }));
  assert.deepEqual(grid(doc.wb, 1, 0, 0, 5, 1), [
    ['Region', 'West'],
    [undefined, undefined],
    ['Row Labels', 'Sum of Sales'],
    ['A', 7],
    ['B', 20],
    ['Grand Total', 27],
  ]);
  assert.deepEqual(pivotFields(doc.wb, 0).filters, [{ field: 0, item: 1 }]);
  const bytes = await save(doc);
  const again = await load(bytes);
  assert.deepEqual(again.wb.pivots!.tables[0].pageFields, [{ fld: 0, item: 1 }]);
  if (hasPython()) assert.deepEqual(independentPivots({ filter: bytes }), []);
  // Several items hidden in a filter: "(Multiple Items)".
  apply(doc, setPivotFields(doc.wb, 0, { filters: [{ field: 0, item: undefined }], hidden: [{ field: 0, items: [0] }] }));
  assert.deepEqual(grid(doc.wb, 1, 0, 0, 0, 1), [['Region', '(Multiple Items)']]);
  apply(doc, setPivotFields(doc.wb, 0, { filters: [] }));
  assert.deepEqual(grid(doc.wb, 1, 0, 0, 1, 1), [
    [undefined, undefined],
    [undefined, undefined],
  ]);
  assert.equal(tableState(doc.wb.pivots!.tables[0]).location.ref.top, 2, 'the body stays where it went');
});

test('subtotals and grand totals turned off; a field sorted descending', async () => {
  const doc = await fileTable({ rows: [0, 1] });
  apply(doc, setPivotFields(doc.wb, 0, { subtotals: [{ field: 0, on: false }], colGrandTotals: false, sort: [{ field: 1, order: 'descending' }] }));
  assert.deepEqual(body(doc), [
    ['Row Labels', 'Sum of Sales'],
    ['East', undefined],
    ['B', 5],
    ['A', 11],
    ['West', undefined],
    ['B', 20],
    ['A', 7],
  ]);
  const bytes = await save(doc);
  const again = await load(bytes);
  const table = again.wb.pivots!.tables[0];
  assert.equal(table.colGrandTotals, false);
  assert.equal(table.fields[0].defaultSubtotal, false);
  assert.equal(table.fields[1].sortType, 'descending');
  assert.ok(!table.fields[0].items!.some((item) => item.t === 'default'), 'the item of the subtotal goes with it');
  const part = await text(bytes, 'xl/pivotTables/pivotTable1.xml');
  assert.match(part, /colGrandTotals="0"/);
  if (hasPython()) assert.deepEqual(independentPivots({ totals: bytes }), []);
});

test('refused: a table that cannot be refreshed, one with formats of its parts, a layout not built here, a field with too many items', async () => {
  const styled = await load(
    await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], tableTail: '<formats count="1"><format dxfId="0"><pivotArea type="all" dataOnly="0" outline="0" fieldPosition="0"/></format></formats>' }),
  );
  assert.match(pivotFieldsProblem(styled.wb, 0) || '', /formats of parts of the table/);
  assert.match(pivotFields(styled.wb, 0).refusal || '', /formats of parts of the table/);
  assert.throws(() => setPivotFields(styled.wb, 0, { rows: [1] }), /Cannot change the fields of the pivot table "PivotTable1": it has formats/);
  const doc = await fileTable();
  assert.match(setPivotFieldsProblem(doc.wb, 0, { rows: [] }) || '', /needs a field in its rows/);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { cols: [1], values: [{ field: 2 }, { field: 3 }] }) || '', /several values and fields in the columns/);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { rows: [0], cols: [0] }) || '', /two places/);
  const limit = PIVOT_LIMITS.axisItems;
  PIVOT_LIMITS.axisItems = 1;
  try {
    assert.match(setPivotFieldsProblem(doc.wb, 0, { rows: [1] }) || '', /the field "Product" has 2 different values: more than 1 cannot be laid out/);
  } finally {
    PIVOT_LIMITS.axisItems = limit;
  }
  // What a refresh refuses: the header of the source changed.
  apply(doc, [{ type: 'setCells', sheet: 0, cells: [{ row: 0, col: 0, cell: { kind: 'string', value: 'Bölge', style: 0 } }] }]);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { rows: [1] }) || '', /^Cannot change the fields of the pivot table "PivotTable1": the header of A1/);
});

test('a table made here, then changed, saved and read back', async () => {
  const doc = await bookWith(DATA);
  apply(doc, createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 5, right: 3 } } }).ops);
  apply(doc, setPivotFields(doc.wb, 0, { rows: [1], cols: [0], values: [{ field: 3, subtotal: 'count' }] }));
  assert.deepEqual(body(doc), [
    ['Count of Units', 'Column Labels', undefined, undefined],
    ['Row Labels', 'East', 'West', 'Grand Total'],
    ['A', 2, 1, 3],
    ['B', 1, 1, 2],
    ['Grand Total', 3, 2, 5],
  ]);
  const bytes = await save(doc);
  const again = await load(bytes);
  const table = again.wb.pivots!.tables[0];
  assert.deepEqual(table.rowFields, [1]);
  assert.deepEqual(table.colFields, [0]);
  assert.deepEqual(table.dataFields, [{ name: 'Count of Units', fld: 3, subtotal: 'count' }]);
  assert.deepEqual(refreshPivot(again.wb, 0).filter((op) => op.type === 'setCells'), []);
  if (hasPython()) assert.deepEqual(independentPivots({ madeChanged: bytes }), []);
  doc.history.undo();
  doc.history.undo();
  assert.equal(doc.wb.pivots, undefined);
});

test('the styles of values follow them when they are put in another order', async () => {
  const doc = await fileTable({ values: [{ fld: 2, numFmtId: 3 }, { fld: 3 }] });
  apply(doc, refreshPivot(doc.wb, 0));
  const table = doc.wb.pivots!.tables[0];
  const styled = (col: number) => {
    const ref = tableState(table).location.ref;
    return table.sheet.rows.get(ref.top + 1)!.cells[ref.left + col].style;
  };
  const sales = styled(1);
  const units = styled(2);
  apply(doc, setPivotFields(doc.wb, 0, { values: [{ field: 3 }, { field: 2 }] }));
  assert.equal(styled(1), units);
  assert.equal(styled(2), sales);
});

test('100,001 rows of source: a table is made and changed in moments; a field of 100,001 different values is not laid out in the rows', async () => {
  const doc = await load(await createBlank('xlsx'));
  const N = 100001;
  const cells: { row: number; col: number; cell: Cell }[] = [];
  ['Region', 'Product', 'Id', 'Sales'].forEach((name, c) => cells.push({ row: 0, col: c, cell: { kind: 'string', value: name, style: 0 } }));
  for (let r = 1; r <= N; r++) {
    cells.push({ row: r, col: 0, cell: { kind: 'string', value: `R${r % 50}`, style: 0 } });
    cells.push({ row: r, col: 1, cell: { kind: 'string', value: `P${r % 300}`, style: 0 } });
    cells.push({ row: r, col: 2, cell: { kind: 'number', value: r, style: 0 } });
    cells.push({ row: r, col: 3, cell: { kind: 'number', value: (r * 7) % 1000, style: 0 } });
  }
  apply(doc, [{ type: 'setCells', sheet: 0, cells }]);
  const source = { sheet: 0, range: { top: 0, left: 0, bottom: N, right: 3 } };
  // Not a measure of speed, a guard against hanging: about a hundred times what it takes on an idle machine (~0.1 s each).
  let began = Date.now();
  const made = createPivot(doc.wb, { source, rows: [0, 1], values: [{ field: 3 }] });
  apply(doc, made.ops);
  assert.ok(Date.now() - began < 10000, `made in ${Date.now() - began} ms`);
  // A header, 50 regions with the products each has (r % 50 and r % 300 meet in 300 pairs), a grand total.
  assert.equal(made.range.bottom - made.range.top + 1, 1 + 50 + 300 + 1);
  began = Date.now();
  apply(doc, setPivotFields(doc.wb, 0, { rows: [1], cols: [0] }));
  assert.ok(Date.now() - began < 10000, `changed in ${Date.now() - began} ms`);
  assert.match(setPivotFieldsProblem(doc.wb, 0, { rows: [2], cols: [] }) || '', /the field "Id" has 100001 different values: more than 100000 cannot be laid out/);
  assert.throws(() => createPivot(doc.wb, { source, rows: [2] }), /the field "Id" has 100001 different values/);
});
