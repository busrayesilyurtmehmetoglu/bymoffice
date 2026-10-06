// The WebView bridge and the list of sheets: the `sheets` message that follows every change
// of the list, and the commands that insert, rename and delete a sheet.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type Bridge, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import type { WorkbookViewOptions } from '../src/ui/sheet-view';

interface FakeSheet {
  name: string;
  state: string;
  kind: string;
}

interface Fake {
  host: BridgeHost;
  bridge: Bridge;
  sent: any[];
  book: { sheets: FakeSheet[] };
  options(): WorkbookViewOptions;
  calls: string[];
  channel: boolean;
  readOnly: boolean;
  editing: boolean;
  /** Thrown by commitEditing(), as a value the data validation stops is. */
  commitError?: Error;
  timers: { callback: () => void; ms: number }[];
  runRetries(): void;
  /** What the bridge sent since the last call. */
  take(): any[];
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const sheet = (name: string, state = 'visible'): FakeSheet => ({ name, state, kind: 'worksheet' });
const all = (count: number) => Array.from({ length: count }, (_value, index) => index);

interface FakeOptions {
  /** A view without the members that change the list of sheets. */
  bare?: boolean;
}

/** A bridge over a view that changes the list of sheets the way WorkbookView does: a new array, then the callbacks. */
function fake(config: FakeOptions = {}): Fake {
  let options: WorkbookViewOptions | undefined;
  let active = 0;
  const undone: FakeSheet[][] = [];
  const state: Fake = {
    host: undefined as unknown as BridgeHost,
    bridge: undefined as unknown as Bridge,
    sent: [],
    book: { sheets: [sheet('Data'), sheet('Hidden', 'hidden'), sheet('Notes')] },
    options: () => options as WorkbookViewOptions,
    calls: [],
    channel: true,
    readOnly: false,
    editing: false,
    timers: [],
    runRetries() {
      const due = state.timers.filter((entry) => entry.ms === 50);
      state.timers = state.timers.filter((entry) => entry.ms !== 50);
      for (const entry of due) entry.callback();
    },
    take() {
      return state.sent.splice(0, state.sent.length);
    },
  };
  const book = state.book;
  /** Replaces the list, shows a sheet and reports as the view does: the sheet shown first, then the change. */
  const apply = (sheets: FakeSheet[], show: number) => {
    undone.push(book.sheets);
    book.sheets = sheets;
    const moved = show !== active;
    active = show;
    const o = options as WorkbookViewOptions;
    if (moved && o.onSheetChange) o.onSheetChange(show);
    if (o.onChange) o.onChange({ sheets: all(sheets.length) });
  };
  const refuse = (text: string): never => {
    throw new RangeError(text);
  };
  const view: any = {
    history: { canUndo: false, canRedo: false },
    get activeSheet() {
      return active;
    },
    get readOnly() {
      return state.readOnly;
    },
    get isEditing() {
      return state.editing;
    },
    workbook: book,
    commitEditing() {
      if (state.commitError) throw state.commitError;
    },
    setReadOnly: (value: boolean) => void (state.readOnly = value),
    setLocale: () => undefined,
    localizeError: (error: unknown) => `Localized: ${(error as Error).message}`,
    undo() {
      const before = undone.pop();
      if (!before) return;
      book.sheets = before;
      active = Math.min(active, before.length - 1);
      (options as WorkbookViewOptions).onChange!({ sheets: all(before.length) });
    },
    redo: () => undefined,
    destroy: () => state.calls.push('destroy'),
  };
  if (!config.bare) {
    view.insertSheet = (at?: number, name?: string): number => {
      state.calls.push(`insertSheet ${at} ${name}`);
      if (state.readOnly) return -1;
      view.commitEditing();
      const index = at === undefined ? book.sheets.length : at;
      if (index > book.sheets.length) refuse(`No place ${index} for a sheet`);
      const title = name === undefined ? `Sheet${book.sheets.length + 1}` : name;
      if (book.sheets.some((entry) => entry.name === title)) refuse(`Another sheet is called "${title}"`);
      const sheets = book.sheets.slice();
      sheets.splice(index, 0, sheet(title));
      apply(sheets, index);
      return index;
    };
    view.renameSheet = (index: number, name: string): boolean => {
      state.calls.push(`renameSheet ${index} ${name}`);
      if (state.readOnly) return false;
      view.commitEditing();
      if (book.sheets[index].name === name) return false;
      if (name === '' || name.length > 31) refuse('A sheet cannot have this name');
      apply(
        book.sheets.map((entry, at) => (at === index ? { ...entry, name } : entry)),
        active,
      );
      return true;
    };
    view.deleteSheet = (index: number): boolean => {
      state.calls.push(`deleteSheet ${index}`);
      if (state.readOnly) return false;
      view.commitEditing();
      if (book.sheets.filter((entry) => entry.state === 'visible').length === 1 && book.sheets[index].state === 'visible') refuse('The last visible sheet cannot be deleted');
      const sheets = book.sheets.filter((_entry, at) => at !== index);
      apply(sheets, active > index ? active - 1 : Math.min(active, sheets.length - 1));
      return true;
    };
  }
  state.host = {
    container: {} as HTMLElement,
    version: '9.9.9',
    post(json) {
      if (!state.channel) return false;
      state.sent.push(JSON.parse(json));
      return true;
    },
    open: async () => ({ format: 'xlsx', workbook: async () => book, save: async () => new Uint8Array([1]) }) as unknown as OfficeFile,
    createView(_container, _workbook, given) {
      options = given;
      state.readOnly = given.readOnly === true;
      return view;
    },
    frame: () => undefined,
    timer: (callback, ms) => state.timers.push({ callback, ms }),
  };
  state.bridge = createBridge(state.host);
  return state;
}

async function loaded(config: FakeOptions = {}, message: Record<string, unknown> = {}): Promise<Fake> {
  const f = fake(config);
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, ...message });
  f.take();
  return f;
}

const LIST = [sheet('Data'), sheet('Hidden', 'hidden'), sheet('Notes')];

test('bridge sheets: insertSheet adds a sheet, and the list is sent in the form of "loaded" before what else it caused', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'insertSheet', id: 'i1' });
  assert.deepEqual(f.calls, ['insertSheet undefined undefined']);
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [...LIST, sheet('Sheet4')], activeSheet: 3 },
    { type: 'sheet', index: 3, name: 'Sheet4' },
    { type: 'changed', sheets: [0, 1, 2, 3], canUndo: false, canRedo: false },
    { type: 'done', id: 'i1', command: 'insertSheet', index: 3, name: 'Sheet4' },
  ]);
  // At a place and with a name.
  await f.bridge.receive({ type: 'insertSheet', id: 'i2', at: 0, name: 'First' });
  const sent = f.take();
  assert.deepEqual(sent[0], { type: 'sheets', sheets: [sheet('First'), ...LIST, sheet('Sheet4')], activeSheet: 0 });
  assert.deepEqual(sent[sent.length - 1], { type: 'done', id: 'i2', command: 'insertSheet', index: 0, name: 'First' });
});

test('bridge sheets: renameSheet and deleteSheet send the list and are confirmed', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'renameSheet', id: 'r', index: 2, name: 'Remarks' });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Data'), sheet('Hidden', 'hidden'), sheet('Remarks')], activeSheet: 0 },
    { type: 'changed', sheets: [0, 1, 2], canUndo: false, canRedo: false },
    { type: 'done', id: 'r', command: 'renameSheet' },
  ]);
  // The name the sheet has: nothing changes and nothing is sent but the answer.
  await f.bridge.receive({ type: 'renameSheet', id: 'r2', index: 2, name: 'Remarks' });
  assert.deepEqual(f.take(), [{ type: 'done', id: 'r2', command: 'renameSheet' }]);
  await f.bridge.receive({ type: 'deleteSheet', id: 'd', index: 1 });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Data'), sheet('Remarks')], activeSheet: 0 },
    { type: 'changed', sheets: [0, 1], canUndo: false, canRedo: false },
    { type: 'done', id: 'd', command: 'deleteSheet' },
  ]);
});

test('bridge sheets: without an id nothing is confirmed, and the list is sent all the same', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'deleteSheet', index: 2 });
  assert.deepEqual(
    f.take().map((message) => message.type),
    ['sheets', 'changed'],
  );
});

test('bridge sheets: the list follows what the user does in the view, and undo', async () => {
  const f = await loaded();
  const view = f.host.createView({} as HTMLElement, f.book as any, f.options());
  // The view's own tab menu: no command was sent.
  view.renameSheet(0, 'Source');
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Source'), sheet('Hidden', 'hidden'), sheet('Notes')], activeSheet: 0 },
    { type: 'changed', sheets: [0, 1, 2], canUndo: false, canRedo: false },
  ]);
  await f.bridge.receive({ type: 'undo', id: 'u' });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: LIST, activeSheet: 0 },
    { type: 'changed', sheets: [0, 1, 2], canUndo: false, canRedo: false },
    { type: 'done', id: 'u', command: 'undo' },
  ]);
});

test('bridge sheets: a change that leaves the list alone sends no list, and neither does another sheet shown', async () => {
  const f = await loaded();
  f.options().onChange!({ sheets: [0] });
  f.options().onSheetChange!(2);
  assert.deepEqual(
    f.take().map((message) => message.type),
    ['changed', 'sheet'],
  );
});

test('bridge sheets: of the lists that wait for a channel only the newest is sent', async () => {
  const f = await loaded();
  f.channel = false;
  await f.bridge.receive({ type: 'insertSheet', name: 'One' });
  await f.bridge.receive({ type: 'insertSheet', name: 'Two' });
  await f.bridge.receive({ type: 'renameSheet', index: 0, name: 'Source' });
  f.channel = true;
  f.runRetries();
  const lists = f.take().filter((message) => message.type === 'sheets');
  assert.deepEqual(lists, [{ type: 'sheets', sheets: [sheet('Source'), sheet('Hidden', 'hidden'), sheet('Notes'), sheet('One'), sheet('Two')], activeSheet: 4 }]);
});

test('bridge sheets: the list of one file is not sent for the next', async () => {
  const f = await loaded();
  f.channel = false;
  await f.bridge.receive({ type: 'insertSheet', name: 'One' });
  await f.bridge.receive({ type: 'load', id: 'l2', data: DATA });
  f.channel = true;
  f.runRetries();
  const sent = f.take();
  // The list of the first file is still told, before the second file is.
  const types = sent.map((message) => message.type);
  assert.ok(types.indexOf('sheets') >= 0 && types.indexOf('sheets') < types.indexOf('loaded'), types.join(' '));
  // The file that is open now says nothing more until its list changes.
  f.options().onChange!({ sheets: [0] });
  assert.deepEqual(
    f.take().map((message) => message.type),
    ['changed'],
  );
});

test('bridge sheets: fields that are not what the contract says are bad messages, and the view is not asked', async () => {
  const f = await loaded();
  const bad: [Record<string, unknown>, RegExp][] = [
    [{ type: 'insertSheet', at: '1' }, /"at" must be a number/],
    [{ type: 'insertSheet', at: -1 }, /"at" must be a sheet index/],
    [{ type: 'insertSheet', at: 1.5 }, /"at" must be a sheet index/],
    [{ type: 'insertSheet', at: 1e300 }, /"at" must be a sheet index/],
    [{ type: 'insertSheet', name: 5 }, /"name" must be a string/],
    [{ type: 'insertSheet', name: 'x'.repeat(256) }, /"name" is too long/],
    [{ type: 'renameSheet', name: 'x' }, /"index" is missing/],
    [{ type: 'renameSheet', index: 0 }, /"name" is missing/],
    [{ type: 'renameSheet', index: '0', name: 'x' }, /"index" must be a number/],
    [{ type: 'renameSheet', index: 0, name: ['x'] }, /"name" must be a string/],
    [{ type: 'renameSheet', index: 0, name: 'x'.repeat(256) }, /"name" is too long/],
    [{ type: 'deleteSheet' }, /"index" is missing/],
    [{ type: 'deleteSheet', index: -1 }, /"index" must be a sheet index/],
    [{ type: 'deleteSheet', index: 0.5 }, /"index" must be a sheet index/],
    [{ type: 'deleteSheet', index: Number.NaN }, /"index" is missing|"index" must be a sheet index/],
  ];
  for (const [message, text] of bad) {
    await f.bridge.receive(JSON.stringify({ id: 'b', ...message }));
    const sent = f.take();
    assert.equal(sent.length, 1, JSON.stringify(message));
    assert.equal(sent[0].type, 'error');
    assert.equal(sent[0].id, 'b');
    assert.equal(sent[0].code, 'bad-message', JSON.stringify(message));
    assert.match(sent[0].message, text);
  }
  assert.deepEqual(f.calls, []);
  assert.deepEqual(f.book.sheets, LIST);
});

test('bridge sheets: an index no sheet has is answered with no-sheet', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'deleteSheet', id: 'a', index: 3 });
  await f.bridge.receive({ type: 'renameSheet', id: 'b', index: 7, name: 'x' });
  // A sheet may be put after the last one, and no further.
  await f.bridge.receive({ type: 'insertSheet', id: 'c', at: 4 });
  assert.deepEqual(f.take(), [
    { type: 'error', id: 'a', code: 'no-sheet', message: 'No sheet with index 3' },
    { type: 'error', id: 'b', code: 'no-sheet', message: 'No sheet with index 7' },
    { type: 'error', id: 'c', code: 'no-sheet', message: 'A sheet cannot be inserted at index 4: there are 3 sheets' },
  ]);
  assert.deepEqual(f.calls, []);
  await f.bridge.receive({ type: 'insertSheet', id: 'd', at: 3 });
  assert.equal(f.take().pop().type, 'done');
});

test('bridge sheets: what the view refuses is answered with sheet-refused and the reason in the language of the view', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'insertSheet', id: 'a', name: 'Notes' });
  await f.bridge.receive({ type: 'renameSheet', id: 'b', index: 0, name: 'x'.repeat(40) });
  assert.deepEqual(f.take(), [
    { type: 'error', id: 'a', code: 'sheet-refused', message: 'Localized: Another sheet is called "Notes"' },
    { type: 'error', id: 'b', code: 'sheet-refused', message: 'Localized: A sheet cannot have this name' },
  ]);
  await f.bridge.receive({ type: 'deleteSheet', index: 1 });
  f.take();
  await f.bridge.receive({ type: 'deleteSheet', index: 1 });
  f.take();
  await f.bridge.receive({ type: 'deleteSheet', id: 'c', index: 0 });
  assert.deepEqual(f.take(), [{ type: 'error', id: 'c', code: 'sheet-refused', message: 'Localized: The last visible sheet cannot be deleted' }]);
  assert.deepEqual(f.book.sheets, [sheet('Data')]);
  // A cell that is being edited and cannot be written (a value its validation stops) refuses the command as well.
  f.commitError = new RangeError('The value does not match');
  await f.bridge.receive({ type: 'insertSheet', id: 'd' });
  assert.deepEqual(f.take(), [{ type: 'error', id: 'd', code: 'sheet-refused', message: 'Localized: The value does not match' }]);
});

test('bridge sheets: a read-only view is answered with read-only, whoever made it so', async () => {
  const f = await loaded({}, { readOnly: true });
  for (const message of [{ type: 'insertSheet' }, { type: 'renameSheet', index: 0, name: 'x' }, { type: 'deleteSheet', index: 0 }]) {
    await f.bridge.receive({ id: 'r', ...message });
    assert.deepEqual(f.take(), [{ type: 'error', id: 'r', code: 'read-only', message: 'The file is read-only' }]);
  }
  assert.deepEqual(f.calls, []);
  await f.bridge.receive({ type: 'setReadOnly', readOnly: false });
  await f.bridge.receive({ type: 'insertSheet', id: 'i' });
  assert.equal(f.take().pop().type, 'done');
  await f.bridge.receive({ type: 'setReadOnly', readOnly: true });
  await f.bridge.receive({ type: 'insertSheet', id: 'j' });
  assert.deepEqual(f.take(), [{ type: 'error', id: 'j', code: 'read-only', message: 'The file is read-only' }]);
});

test('bridge sheets: without a file the commands are answered with not-loaded; bad fields are complained about first', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'insertSheet', id: 'a' });
  await f.bridge.receive({ type: 'renameSheet', id: 'b', index: 0, name: 'x' });
  await f.bridge.receive({ type: 'deleteSheet', id: 'c', index: 0 });
  await f.bridge.receive({ type: 'deleteSheet', id: 'd' });
  assert.deepEqual(
    f.take().map((message) => [message.type, message.id, message.code]),
    [
      ['ready', undefined, undefined],
      ['error', 'a', 'not-loaded'],
      ['error', 'b', 'not-loaded'],
      ['error', 'c', 'not-loaded'],
      ['error', 'd', 'bad-message'],
    ],
  );
});

test('bridge sheets: a document has no sheets, and a view that cannot change its sheets says so', async () => {
  const f = fake();
  f.host.open = async () => ({ format: 'docx', kind: 'document', document: async () => ({}), save: async () => new Uint8Array([1]) }) as unknown as OfficeFile;
  let change: (() => void) | undefined;
  f.host.createDocumentView = (_container, _doc, options) => {
    change = options.onChange as () => void;
    return { destroy: () => undefined, canUndo: true, canRedo: false };
  };
  await f.bridge.receive({ type: 'load', data: DATA });
  f.take();
  for (const message of [{ type: 'insertSheet' }, { type: 'renameSheet', index: 0, name: 'x' }, { type: 'deleteSheet', index: 0 }]) {
    await f.bridge.receive({ id: 'n', ...message });
    assert.deepEqual(f.take(), [{ type: 'error', id: 'n', code: 'no-sheet', message: 'The open file is a document: it has no sheets' }]);
  }
  // A change of a document says nothing about sheets.
  (change as () => void)();
  assert.deepEqual(f.take(), [{ type: 'changed', sheets: [], canUndo: true, canRedo: false }]);

  const g = await loaded({ bare: true });
  await g.bridge.receive({ type: 'insertSheet', id: 'i' });
  assert.deepEqual(g.take(), [{ type: 'error', id: 'i', code: 'sheet-refused', message: 'This view cannot change its sheets' }]);
});

test('bridge sheets: a list that waits for a channel is brought up to date where it stands, before what follows it', async () => {
  const f = await loaded();
  f.channel = false;
  await f.bridge.receive({ type: 'insertSheet', id: 'i', name: 'One' });
  await f.bridge.receive({ type: 'renameSheet', id: 'r', index: 0, name: 'Source' });
  f.channel = true;
  f.runRetries();
  const sent = f.take();
  const types = sent.map((message) => message.type);
  // One list, the newest, and nothing that speaks of its sheets comes before it.
  assert.equal(types.filter((type) => type === 'sheets').length, 1);
  assert.equal(types[0], 'sheets', types.join(' '));
  assert.deepEqual(sent[0], { type: 'sheets', sheets: [sheet('Source'), sheet('Hidden', 'hidden'), sheet('Notes'), sheet('One')], activeSheet: 3 });
  for (const type of ['sheet', 'changed', 'done']) assert.ok(types.indexOf(type) > 0, `${type} after the list: ${types.join(' ')}`);
  assert.deepEqual(
    sent.filter((message) => message.type === 'done').map((message) => message.id),
    ['i', 'r'],
  );
  // The size of what waits is counted as before: nothing is left over once it is sent.
  f.channel = false;
  await f.bridge.receive({ type: 'renameSheet', index: 0, name: 'Again' });
  f.channel = true;
  f.runRetries();
  assert.deepEqual(
    f.take().map((message) => message.type),
    ['sheets', 'changed'],
  );
});

test('bridge: "author" of a load reaches the view of a workbook, where it is who the notes added to cells are by', async () => {
  const named = await loaded({}, { author: 'Ada Lovelace' });
  assert.equal(named.options().author, 'Ada Lovelace');
  const plain = await loaded();
  assert.equal('author' in plain.options(), false);
  const long = fake();
  await long.bridge.receive({ type: 'load', id: 'l', data: DATA, author: 'x'.repeat(256) });
  assert.deepEqual(
    long
      .take()
      .filter((message) => message.type === 'error')
      .map((message) => message.code),
    ['bad-message'],
  );
});
