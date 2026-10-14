// What the commands of pivot tables say to the user, and the WebView bridge's "refreshPivots".

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type Bridge } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import { refreshPivot } from '../src/ops/pivots';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { load, pivotBook, setData, type Datum, type PivotSpec } from './pivot-fixtures';

const DATA: Datum[][] = [
  ['Region', 'Product', 'Sales'],
  ['East', 'A', 10],
  ['West', 'B', 20],
];

const reasonOf = (run: () => unknown): string => {
  try {
    run();
  } catch (error) {
    assert.ok(error instanceof RangeError);
    return (error as Error).message;
  }
  throw new assert.AssertionError({ message: 'the refresh was not refused' });
};

test('every reason a refresh is refused with is said in Turkish', async () => {
  const messages: string[] = [];
  const refused = async (spec: Partial<PivotSpec>, change?: (doc: Awaited<ReturnType<typeof load>>) => void) => {
    const doc = await load(await pivotBook({ data: DATA, rows: [0], values: [{ fld: 2 }], ...spec } as PivotSpec));
    if (change) change(doc);
    messages.push(reasonOf(() => refreshPivot(doc.wb, 0)));
  };
  await refused({ around: [[4, 0, 'x']] });
  await refused({}, (doc) => setData(doc, 0, 1, 'Item'));
  await refused({}, (doc) => setData(doc, 0, 1, null));
  await refused({ fieldTail: { 0: '<fieldGroup base="0"/>' } });
  await refused({ tableTail: '<filters count="0"/>' });
  await refused({ tableAttrs: ' dataOnRows="1"' });
  await refused({ tableAttrs: ' gridDropZones="1"' });
  await refused({ tableAttrs: ' mergeItem="1"' });
  await refused({ tableTail: '<chartFormats count="0"/>' });
  await refused({ fieldTail: { 0: '' }, cacheTail: '<calculatedMembers/>' });
  await refused({ cols: [1], values: [{ fld: 2 }, { fld: 2, subtotal: 'max' }] });
  await refused({ form: 'outline', cols: [1] });
  await refused({ pages: [{ fld: 1, item: 'A' }], at: { row: 4, col: 0 } }, (doc) => setData(doc, 1, 1, 'C'));
  for (const message of messages) {
    const said = STRINGS_TR.pivots.refused(message);
    assert.match(said, /^"PivotTable1" pivot tablosu yenilenemiyor: /, message);
    assert.doesNotMatch(said, / (it|its|the|a) /, `${message} -> ${said}`);
    assert.equal(STRINGS_EN.pivots.refused(message), message);
  }
  // A text that is not a refusal of a refresh is left as it is.
  assert.equal(STRINGS_TR.pivots.refused('Something else'), 'Something else');
  assert.equal(STRINGS_TR.pivots.refreshed(1), 'Pivot tablo yenilendi');
  assert.equal(STRINGS_EN.pivots.refreshed(3), '3 pivot tables refreshed');
});

interface Fake {
  bridge: Bridge;
  sent: any[];
  readOnly: boolean;
  result: { refreshed: number[]; refused: { index: number; name: string; reason: string }[] } | Error;
  commits: number;
}

function fake(format: 'xlsx' | 'docx' = 'xlsx', bare = false): Fake {
  const state: Fake = { bridge: undefined as unknown as Bridge, sent: [], readOnly: false, result: { refreshed: [], refused: [] }, commits: 0 };
  const book = { sheets: [{ name: 'Data', state: 'visible', kind: 'worksheet' }] };
  const view: any = {
    history: { canUndo: false, canRedo: false },
    activeSheet: 0,
    isEditing: false,
    get readOnly() {
      return state.readOnly;
    },
    workbook: book,
    commitEditing: () => undefined,
    commitPending: () => {
      state.commits++;
      return true;
    },
    localizeError: (error: Error) => `Localized: ${error.message}`,
    setReadOnly: (value: boolean) => void (state.readOnly = value),
    setLocale: () => undefined,
    destroy: () => undefined,
  };
  if (!bare) {
    view.refreshAllPivots = () => {
      if (state.result instanceof Error) throw state.result;
      return state.result;
    };
  }
  state.bridge = createBridge({
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: async () => {
      const file: any = { format, kind: format === 'docx' ? 'document' : 'spreadsheet', save: async () => new Uint8Array([1]) };
      file.workbook = async () => book;
      if (format === 'docx') file.document = async () => ({ body: [], settings: {} });
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

test('bridge: refreshPivots answers with what was refreshed and what was not', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'refreshPivots', id: 'early' });
  assert.deepEqual(last(f), { type: 'error', id: 'early', code: 'not-loaded', message: 'No file is open' });
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA64 });
  f.result = { refreshed: [0, 2], refused: [{ index: 1, name: 'Satışlar', reason: 'kaynağında başlığın altında satır yok' }] };
  await f.bridge.receive({ type: 'refreshPivots', id: 'p' });
  assert.deepEqual(last(f), { type: 'done', id: 'p', command: 'refreshPivots', refreshed: 2, refused: [{ name: 'Satışlar', reason: 'kaynağında başlığın altında satır yok' }] });
  assert.ok(f.commits > 0, 'what is being typed is applied first');
  // Refused as a whole: editing together.
  f.result = new RangeError('A pivot table cannot be refreshed while the workbook is edited together');
  await f.bridge.receive({ type: 'refreshPivots', id: 'c' });
  assert.deepEqual(last(f), { type: 'error', id: 'c', code: 'pivot-refused', message: 'Localized: A pivot table cannot be refreshed while the workbook is edited together' });
  f.readOnly = true;
  await f.bridge.receive({ type: 'refreshPivots', id: 'r' });
  assert.deepEqual(last(f), { type: 'error', id: 'r', code: 'read-only', message: 'The file is read-only' });
});

test('bridge: refreshPivots of a document, and of a view that cannot', async () => {
  const doc = fake('docx');
  await doc.bridge.receive({ type: 'load', id: 'l', data: DATA64 });
  await doc.bridge.receive({ type: 'refreshPivots', id: 'd' });
  assert.equal(last(doc).code, 'no-sheet');
  const bare = fake('xlsx', true);
  await bare.bridge.receive({ type: 'load', id: 'l', data: DATA64 });
  await bare.bridge.receive({ type: 'refreshPivots', id: 'b' });
  assert.deepEqual(last(bare), { type: 'error', id: 'b', code: 'pivot-refused', message: 'This view cannot refresh pivot tables' });
});
