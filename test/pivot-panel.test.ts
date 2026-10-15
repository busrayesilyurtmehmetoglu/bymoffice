// The panel of the fields of a pivot table (what a move in it changes), the dialog that makes one
// (what its fields are read as), what the commands say in Turkish, and the WebView bridge's
// "insertPivot", "setPivotFields" and "pivotFields".

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type Bridge } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import { createPivot, createPivotProblem } from '../src/ops/pivot-create';
import { pivotFields, setPivotFields, setPivotFieldsProblem, type PivotFieldsInfo } from '../src/ops/pivot-fields';
import { filterChoices, moveChange, toggleChange, valueChange, isUsed } from '../src/ui/pivot-panel';
import { pivotSourceText, readPivotCell, readPivotSource } from '../src/ui/pivot-dialog';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { bookWith, load, pivotBook, type Datum, type Loaded } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales', 'Units'],
  ['East', 'A', 10, 1],
  ['West', 'B', 20, 2],
  ['East', 'B', 5, 3],
];

const INFO: PivotFieldsInfo = {
  fields: [
    { index: 0, name: 'Region', subtotal: true, sort: 'manual', items: [{ text: 'East', hidden: false }] },
    { index: 1, name: 'Product', subtotal: true, sort: 'manual' },
    { index: 2, name: 'Sales', subtotal: true, sort: 'manual' },
    { index: 3, name: 'Units', subtotal: true, sort: 'manual' },
  ],
  rows: [0],
  cols: [],
  filters: [{ field: 1, item: 0 }],
  values: [{ field: 2, subtotal: 'sum', name: 'Sum of Sales' }],
  rowGrandTotals: true,
  colGrandTotals: true,
};

test('panel: a field is put into a box, moved within it and to another, taken out', () => {
  // Into the rows at the end, then before Region.
  assert.deepEqual(moveChange(INFO, 3, undefined, 'rows').rows, [0, 3]);
  assert.deepEqual(moveChange(INFO, 3, undefined, 'rows', 0).rows, [3, 0]);
  // From the filters to the columns: it leaves the filters; the filter that stays keeps its item.
  const moved = moveChange(INFO, 1, { box: 'filters', at: 0 }, 'cols');
  assert.deepEqual([moved.cols, moved.filters], [[1], []]);
  assert.deepEqual(moveChange(INFO, 0, undefined, 'filters').filters, [{ field: 1, item: 0 }, { field: 0 }]);
  assert.deepEqual(moveChange(INFO, 0, undefined, 'filters').rows, []);
  // A value: put in again as a sum; moved within the values keeping its function; taken out.
  assert.deepEqual(moveChange(INFO, 3, undefined, 'values').values, [{ field: 2, subtotal: 'sum' }, { field: 3, subtotal: 'sum' }]);
  const two = { ...INFO, values: [INFO.values[0], { field: 3, subtotal: 'max' as const, name: 'Max of Units' }] };
  assert.deepEqual(moveChange(two, 3, { box: 'values', at: 1 }, 'values', 0).values, [{ field: 3, subtotal: 'max' }, { field: 2, subtotal: 'sum' }]);
  assert.deepEqual(moveChange(two, 2, { box: 'values', at: 0 }, null).values, [{ field: 3, subtotal: 'max' }]);
  // A row moved down past the next one.
  const rows = { ...INFO, rows: [0, 3, 2] };
  assert.deepEqual(moveChange(rows, 0, { box: 'rows', at: 0 }, 'rows', 1).rows, [3, 0, 2]);
  assert.deepEqual(moveChange(rows, 2, { box: 'rows', at: 2 }, 'rows', 1).rows, [0, 2, 3]);
});

test('panel: a ticked field goes where Excel puts it; unticked it leaves every box; a value changes its function or format', () => {
  assert.deepEqual(toggleChange(INFO, 3, true, true).values, [{ field: 2, subtotal: 'sum' }, { field: 3, subtotal: 'sum' }]);
  assert.deepEqual(toggleChange(INFO, 3, true, false).rows, [0, 3]);
  const off = toggleChange(INFO, 2, false, true);
  assert.deepEqual([off.rows, off.values], [[0], []]);
  assert.equal(isUsed(INFO, 1), true);
  assert.equal(isUsed(INFO, 3), false);
  assert.deepEqual(valueChange(INFO, 0, 'count', undefined), { values: [{ field: 2, subtotal: 'count' }] });
  assert.deepEqual(valueChange(INFO, 0, undefined, '#,##0'), { values: [{ field: 2, subtotal: 'sum', numFmt: '#,##0' }] });
});

test('panel changes applied to a table come out as the panel means them', async () => {
  const doc = await bookWith(DATA);
  doc.history.apply(createPivot(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 3, right: 3 } } }).ops, undefined, undefined);
  let info = pivotFields(doc.wb, 0);
  doc.history.apply(setPivotFields(doc.wb, 0, toggleChange(info, 1, true, false)), undefined, undefined);
  info = pivotFields(doc.wb, 0);
  assert.deepEqual(info.rows, [0, 1]);
  doc.history.apply(setPivotFields(doc.wb, 0, moveChange(info, 1, { box: 'rows', at: 1 }, 'cols')), undefined, undefined);
  info = pivotFields(doc.wb, 0);
  assert.deepEqual([info.rows, info.cols], [[0], [1]]);
  doc.history.apply(setPivotFields(doc.wb, 0, valueChange(info, 0, 'count', undefined)), undefined, undefined);
  assert.deepEqual(pivotFields(doc.wb, 0).values, [{ field: 2, subtotal: 'count', name: 'Count of Sales' }]);
  // A second value while a field stands in the columns is not built here: said.
  assert.match(setPivotFieldsProblem(doc.wb, 0, toggleChange(pivotFields(doc.wb, 0), 3, true, true)) || '', /several values and fields in the columns/);
});

test('dialog: what the fields of cells and of the place are read as', async () => {
  const doc = await bookWith(DATA);
  assert.deepEqual(readPivotSource('A1:D4', doc.wb, 0), { sheet: 0, range: { top: 0, left: 0, bottom: 3, right: 3 } });
  assert.deepEqual(readPivotSource("= 'Sheet1'!$A$1:$D$4", doc.wb, 0), { sheet: 0, range: { top: 0, left: 0, bottom: 3, right: 3 } });
  assert.deepEqual(readPivotSource('sheet1!B2', doc.wb, 0), { sheet: 0, range: { top: 1, left: 1, bottom: 1, right: 1 } });
  assert.equal(readPivotSource('Nowhere!A1:B2', doc.wb, 0), undefined);
  assert.equal(readPivotSource('not a range', doc.wb, 0), undefined);
  assert.equal(readPivotSource('', doc.wb, 0), undefined);
  assert.deepEqual(readPivotCell('$F$2'), { row: 1, col: 5 });
  assert.equal(readPivotCell('F2:G3'), undefined);
  assert.equal(pivotSourceText(doc.wb, 0, { top: 0, left: 0, bottom: 3, right: 3 }, 0), 'A1:D4');
  assert.equal(pivotSourceText(doc.wb, 0, { top: 0, left: 0, bottom: 3, right: 3 }, 1), 'Sheet1!A1:D4');
});

test('every reason a table is not made or its fields are not changed is said in Turkish', async () => {
  const messages: string[] = [];
  const add = (message: string | undefined) => {
    assert.ok(message, 'refused');
    messages.push(message as string);
  };
  const doc = await bookWith(DATA);
  const source = { sheet: 0, range: { top: 0, left: 0, bottom: 3, right: 3 } };
  add(createPivotProblem(doc.wb, { source, rows: [] }));
  add(createPivotProblem(doc.wb, { source, values: [] }));
  add(createPivotProblem(doc.wb, { source, rows: [0], cols: [0] }));
  add(createPivotProblem(doc.wb, { source, rows: [9] }));
  add(createPivotProblem(doc.wb, { source, cols: [1], values: [{ field: 2 }, { field: 3 }] }));
  add(createPivotProblem(doc.wb, { source, target: { sheet: 0, row: 2, col: 1 } }));
  add(createPivotProblem(doc.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 0, right: 3 } } }));
  add(createPivotProblem(doc.wb, { source: { sheet: 7, range: source.range } }));
  add(createPivotProblem(doc.wb, { source: { name: 'Nothing' } }));
  add(createPivotProblem(doc.wb, { source, values: [{ field: 2, numFmt: 7 as never }] }));
  const twice = await bookWith([['A', 'a'], [1, 2]]);
  add(createPivotProblem(twice.wb, { source: { sheet: 0, range: { top: 0, left: 0, bottom: 1, right: 1 } } }));
  doc.history.apply(createPivot(doc.wb, { source }).ops, undefined, undefined);
  add(createPivotProblem(doc.wb, { source, name: 'PivotTable1' }));
  add(setPivotFieldsProblem(doc.wb, 0, { values: [{ field: 2, name: 'Units' }] }));
  add(setPivotFieldsProblem(doc.wb, 0, { values: [{ field: 2, name: 'X' }, { field: 3, name: 'x' }] }));
  add(setPivotFieldsProblem(doc.wb, 0, { hidden: [{ field: 3, items: [0] }] }));
  add(setPivotFieldsProblem(doc.wb, 0, { hidden: [{ field: 0, items: [0, 1] }] }));
  add(setPivotFieldsProblem(doc.wb, 0, { hidden: [{ field: 0, items: [5] }] }));
  add(setPivotFieldsProblem(doc.wb, 0, { filters: [{ field: 1, item: 9 }] }));
  const styled = await load(
    await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], tableTail: '<formats count="1"><format dxfId="0"><pivotArea type="all" dataOnly="0" outline="0" fieldPosition="0"/></format></formats>' }),
  );
  add(setPivotFieldsProblem(styled.wb, 0, { rows: [1] }));
  for (const message of messages) {
    const said = STRINGS_TR.pivots.refused(message);
    assert.match(said, /^(Pivot tablo eklenemiyor|"PivotTable1" pivot tablosunun alanları değiştirilemiyor): /, `${message} -> ${said}`);
    assert.doesNotMatch(said, / (it|its|the|a|is|has) /, `${message} -> ${said}`);
    assert.equal(STRINGS_EN.pivots.refused(message), message);
  }
});

interface Fake {
  bridge: Bridge;
  sent: any[];
  calls: any[];
  result: unknown;
}

function fake(): Fake {
  const state: Fake = { bridge: undefined as unknown as Bridge, sent: [], calls: [], result: 0 };
  let workbook: Loaded['wb'] | undefined;
  const view: any = {
    history: { canUndo: false, canRedo: false },
    activeSheet: 0,
    isEditing: false,
    readOnly: false,
    get workbook() {
      return workbook;
    },
    selection: { active: { row: 0, col: 0 }, range: { top: 0, left: 0, bottom: 0, right: 0 } },
    commitEditing: () => undefined,
    commitPending: () => true,
    localizeError: (error: Error) => `Localized: ${error.message}`,
    setReadOnly: () => undefined,
    setLocale: () => undefined,
    destroy: () => undefined,
    refreshAllPivots: () => ({ refreshed: [], refused: [] }),
    insertPivot: (options: unknown, panel: boolean) => {
      state.calls.push(['insertPivot', options, panel]);
      if (state.result instanceof Error) throw state.result;
      return state.result;
    },
    setPivotFields: (index: number, change: unknown) => {
      state.calls.push(['setPivotFields', index, change]);
      if (state.result instanceof Error) throw state.result;
      return state.result;
    },
    pivotFields: (index: number) => {
      state.calls.push(['pivotFields', index]);
      return { ...INFO, fields: INFO.fields.slice(0, 1) };
    },
  };
  state.bridge = createBridge({
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: async () => {
      const file: any = { format: 'xlsx', kind: 'spreadsheet', save: async () => new Uint8Array([1]) };
      workbook = (await bookWith(DATA)).wb;
      file.workbook = async () => workbook;
      return file as OfficeFile;
    },
    createView: () => view,
    createDocumentView: () => ({ destroy: () => undefined }) as any,
    frame: () => undefined,
    timer: () => undefined,
  });
  return state;
}

const DATA64 = base64Encode(new Uint8Array([80, 75, 3, 4]));
const last = (f: Fake) => f.sent[f.sent.length - 1];

test('bridge: insertPivot, setPivotFields and pivotFields reach the view with what the message holds, as lists and numbers only', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'insertPivot', id: 'early' });
  assert.equal(last(f).code, 'not-loaded');
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA64 });
  f.result = 0;
  await f.bridge.receive({ type: 'insertPivot', id: 'i', source: 'A1:D4', rows: [1], values: [{ field: 2, subtotal: 'max', numFmt: '0.00', name: 'ignored', extra: 1 }], cell: 'H2' });
  assert.deepEqual(last(f), { type: 'done', id: 'i', command: 'insertPivot', index: 0, sheet: 0 });
  assert.deepEqual(f.calls[0], [
    'insertPivot',
    { rows: [1], values: [{ field: 2, subtotal: 'max', numFmt: '0.00' }], source: { sheet: 0, range: { top: 0, left: 0, bottom: 3, right: 3 } }, target: { sheet: 0, row: 1, col: 7 } },
    false,
  ]);
  await f.bridge.receive({ type: 'insertPivot', id: 'bad', source: 'not a range' });
  assert.deepEqual(last(f), { type: 'error', id: 'bad', code: 'pivot-refused', message: 'The source of the pivot table is not a range or a name of the workbook' });
  f.result = true;
  await f.bridge.receive({ type: 'setPivotFields', id: 's', index: 0, cols: ['x'], subtotals: [{ field: 0, on: false }], hidden: [{ field: 0, items: [1, 'y'] }], colGrandTotals: false });
  assert.deepEqual(last(f), { type: 'done', id: 's', command: 'setPivotFields', changed: true });
  assert.deepEqual(f.calls[1], ['setPivotFields', 0, { cols: [NaN], colGrandTotals: false, subtotals: [{ field: 0, on: false }], hidden: [{ field: 0, items: [1, NaN] }] }]);
  f.result = new RangeError('Cannot change the fields of the pivot table "P": it needs a value here');
  await f.bridge.receive({ type: 'setPivotFields', id: 'e', index: 0, values: [] });
  assert.deepEqual(last(f), { type: 'error', id: 'e', code: 'pivot-refused', message: 'Localized: Cannot change the fields of the pivot table "P": it needs a value here' });
  await f.bridge.receive({ type: 'setPivotFields', id: 'n' });
  assert.equal(last(f).message, 'The message names no pivot table (index)');
  await f.bridge.receive({ type: 'pivotFields', id: 'f', index: 0 });
  assert.equal(last(f).command, 'pivotFields');
  assert.deepEqual(last(f).fields.fields, [{ index: 0, name: 'Region', subtotal: true, sort: 'manual', items: [{ text: 'East', hidden: false }] }]);
});

test('panel: the list of a filter is cut at 500 items but always holds the one it shows', () => {
  const items = Array.from({ length: 800 }, (_, i) => ({ text: `i${i}` }));
  const choices = filterChoices(items, 700, '(All)');
  assert.equal(choices.length, 502);
  assert.deepEqual(choices[0], ['', '(All)']);
  assert.deepEqual(choices[501], ['700', 'i700']);
  assert.equal(filterChoices(items, 3, '(All)').length, 501);
});
