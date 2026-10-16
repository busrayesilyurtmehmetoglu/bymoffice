// Deleting pivot tables as the user meets it: what a refusal says in Turkish, and the WebView
// bridge's "deletePivot".

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type Bridge } from '../src/embed/bridge';
import { ZipArchive, type OfficeFile } from '../src/index';
import { deletePivotProblem } from '../src/ops/pivot-delete';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { bookWith, load, pivotBook, type Datum } from './pivot-fixtures';

const encoder = new TextEncoder();
const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
];
const CHART_NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

async function edit(bytes: Uint8Array, change: (zip: ZipArchive) => void | Promise<void>): Promise<Uint8Array> {
  const zip = ZipArchive.read(bytes);
  await change(zip);
  return zip.write();
}

test('every reason a deletion is refused with is said in Turkish', async () => {
  const base = () => pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }] });
  const messages: string[] = [];
  const add = (message: string | undefined) => {
    assert.ok(message, 'refused');
    messages.push(message as string);
  };
  const chart = encoder.encode(`<c:chartSpace xmlns:c="${CHART_NS}"><c:pivotSource><c:name>Pivot!PivotTable1</c:name><c:fmtId val="0"/></c:pivotSource><c:chart/></c:chartSpace>`);
  add(deletePivotProblem((await load(await edit(await base(), (zip) => zip.set('xl/charts/chart1.xml', chart)))).wb, 0));
  add(deletePivotProblem((await load(await edit(await base(), (zip) => zip.set('xl/slicerCaches/slicerCache1.xml', encoder.encode('<x/>'))))).wb, 0));
  add(deletePivotProblem((await load(await edit(await base(), async (zip) => zip.set('xl/pivotTables/pivotTable9.xml', await zip.read('xl/pivotTables/pivotTable1.xml'))))).wb, 0));
  // Two tables over each other (a damaged file).
  const over = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], twin: { row: 3, col: 1 } }));
  add(deletePivotProblem(over.wb, 0));
  for (const message of messages) {
    const said = STRINGS_TR.pivots.refused(message);
    assert.match(said, /^"PivotTable1" pivot tablosu silinemiyor: /, `${message} -> ${said}`);
    assert.doesNotMatch(said, / (it|its|the|a|is|has) /, `${message} -> ${said}`);
    assert.equal(STRINGS_EN.pivots.refused(message), message);
  }
  assert.equal(STRINGS_TR.pivots.deleteTable, 'Pivot tabloyu sil');
  assert.equal(STRINGS_EN.pivots.deleteTable, 'Delete pivot table');
});

interface Fake {
  bridge: Bridge;
  sent: any[];
  calls: any[];
  result: unknown;
  readOnly: boolean;
}

function fake(): Fake {
  const state: Fake = { bridge: undefined as unknown as Bridge, sent: [], calls: [], result: true, readOnly: false };
  let workbook: any;
  const view: any = {
    history: { canUndo: false, canRedo: false },
    activeSheet: 0,
    isEditing: false,
    get readOnly() {
      return state.readOnly;
    },
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
    insertPivot: () => 0,
    setPivotFields: () => false,
    pivotFields: () => ({}),
    deletePivot: (index: number) => {
      state.calls.push(['deletePivot', index]);
      if (state.result instanceof Error) throw state.result;
      return state.result;
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

const last = (f: Fake) => f.sent[f.sent.length - 1];

test('bridge: deletePivot reaches the view with the index, says whether anything changed, and what refused it', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'deletePivot', id: 'early', index: 0 });
  assert.equal(last(f).code, 'not-loaded');
  await f.bridge.receive({ type: 'load', id: 'l', data: base64Encode(new Uint8Array([80, 75, 3, 4])) });
  await f.bridge.receive({ type: 'deletePivot', id: 'd', index: 2 });
  assert.deepEqual(last(f), { type: 'done', id: 'd', command: 'deletePivot', changed: true });
  assert.deepEqual(f.calls, [['deletePivot', 2]]);
  await f.bridge.receive({ type: 'deletePivot', id: 'n', index: '2' });
  assert.deepEqual(last(f), { type: 'error', id: 'n', code: 'pivot-refused', message: 'The message names no pivot table (index)' });
  f.result = new RangeError('Cannot delete the pivot table "P": a slicer or a timeline reads the caches of the workbook');
  await f.bridge.receive({ type: 'deletePivot', id: 'e', index: 0 });
  assert.deepEqual(last(f), { type: 'error', id: 'e', code: 'pivot-refused', message: 'Localized: Cannot delete the pivot table "P": a slicer or a timeline reads the caches of the workbook' });
  f.readOnly = true;
  await f.bridge.receive({ type: 'deletePivot', id: 'r', index: 0 });
  assert.equal(last(f).code, 'read-only');
  assert.equal(f.calls.length, 2, 'not asked of a read-only view');
});
