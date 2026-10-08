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
    view.moveSheet = (from: number, to: number): boolean => {
      state.calls.push(`moveSheet ${from} ${to}`);
      if (state.readOnly) return false;
      view.commitEditing();
      if (from === to) return false;
      const sheets = book.sheets.slice();
      sheets.splice(to, 0, sheets.splice(from, 1)[0]);
      apply(sheets, sheets.indexOf(book.sheets[active]));
      return true;
    };
    view.setSheetState = (index: number, value: string): boolean => {
      state.calls.push(`setSheetState ${index} ${value}`);
      if (state.readOnly) return false;
      view.commitEditing();
      if (book.sheets[index].state === value) return false;
      if (value !== 'visible' && book.sheets.filter((entry) => entry.state === 'visible').length === 1 && book.sheets[index].state === 'visible') refuse('The last visible sheet cannot be hidden');
      apply(
        book.sheets.map((entry, at) => (at === index ? { ...entry, state: value } : entry)),
        active,
      );
      return true;
    };
    view.copySheet = (index: number, given: { at?: number; name?: string } = {}): number => {
      state.calls.push(`copySheet ${index} ${JSON.stringify(given)}`);
      if (state.readOnly) return -1;
      view.commitEditing();
      const at = given.at === undefined ? index + 1 : given.at;
      const title = given.name === undefined ? `${book.sheets[index].name} (2)` : given.name;
      if (book.sheets.some((entry) => entry.name === title)) refuse(`Another sheet is called "${title}"`);
      const sheets = book.sheets.slice();
      sheets.splice(at, 0, sheet(title));
      apply(sheets, at);
      return at;
    };
    view.setTabColor = (index: number, color: string | undefined): boolean => {
      state.calls.push(`setTabColor ${index} ${color}`);
      if (state.readOnly) return false;
      view.commitEditing();
      apply(
        book.sheets.map((entry, at) => {
          if (at !== index) return entry;
          const next: any = { ...entry };
          if (color === undefined) delete next.tabColor;
          else next.tabColor = { rgb: `FF${color.slice(1).toUpperCase()}` };
          return next;
        }),
        active,
      );
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

// --- second review -------------------------------------------------------------------

test('review 2: a selection that waits to be told is told by the index its sheet has once the list of sheets changed', async () => {
  const f = fake();
  const frames: (() => void)[] = [];
  f.host.frame = (callback: () => void) => void frames.push(callback);
  f.bridge = createBridge(f.host);
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA });
  f.take();
  const view = f.host.createView({} as HTMLElement, f.book as any, f.options()) as any;
  // What the view answers when asked what is selected: the sheet it shows now.
  view.selection = { range: { top: 4, left: 2, bottom: 4, right: 2 }, active: { row: 4, col: 2 } };
  // A selection on the third sheet is noted for the next frame...
  f.options().onSheetChange!(2);
  f.options().onSelectionChange!({ range: { top: 1, left: 1, bottom: 1, right: 1 }, active: { row: 1, col: 1 } }, 2);
  f.take();
  // ...and before that frame a sheet is inserted in front: the sheet that was third is the fourth.
  view.insertSheet(0, 'First');
  f.take();
  for (const callback of frames.splice(0)) callback();
  const told = f.take().filter((message) => message.type === 'selection');
  assert.equal(told.length, 1);
  assert.equal(told[0].sheet, view.activeSheet, 'the index of the list the application knows');
  assert.deepEqual([told[0].range, told[0].active], ['C5', 'C5']);
});

test('review 2: save applies what is being typed through commitPending, and saves nothing while a note or a name cannot be applied', async () => {
  const f = await loaded();
  const view = f.host.createView({} as HTMLElement, f.book as any, f.options()) as any;
  let pending = true;
  view.commitPending = () => {
    f.calls.push('commitPending');
    return pending;
  };
  await f.bridge.receive({ type: 'save', id: 's' });
  assert.deepEqual(f.calls, ['commitPending']);
  assert.deepEqual(
    f.take().map((message) => [message.type, message.id]),
    [['saved', 's']],
  );
  pending = false;
  await f.bridge.receive({ type: 'save', id: 't' });
  const answer = f.take();
  assert.deepEqual(
    answer.map((message) => [message.type, message.id, message.code]),
    [['error', 't', 'save-failed']],
  );
  // A view of before the method: its cell editor is written as it was.
  delete view.commitPending;
  f.editing = true;
  f.commitError = new RangeError('The value is not allowed');
  await f.bridge.receive({ type: 'save', id: 'u' });
  assert.deepEqual(
    f.take().map((message) => [message.type, message.id, message.code]),
    [['error', 'u', 'save-failed']],
  );
});

test('bridge sheets: moveSheet, setSheetState, copySheet and setTabColor go through the view, send the list and are confirmed', async () => {
  const f = await loaded();
  await f.bridge.receive({ type: 'moveSheet', id: 'm', from: 0, to: 2 });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Hidden', 'hidden'), sheet('Notes'), sheet('Data')], activeSheet: 2 },
    { type: 'sheet', index: 2, name: 'Data' },
    { type: 'changed', sheets: [0, 1, 2], canUndo: false, canRedo: false },
    { type: 'done', id: 'm', command: 'moveSheet' },
  ]);
  await f.bridge.receive({ type: 'setSheetState', id: 's', index: 0, state: 'visible' });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Hidden'), sheet('Notes'), sheet('Data')], activeSheet: 2 },
    { type: 'changed', sheets: [0, 1, 2], canUndo: false, canRedo: false },
    { type: 'done', id: 's', command: 'setSheetState' },
  ]);
  await f.bridge.receive({ type: 'copySheet', id: 'c', index: 1 });
  assert.deepEqual(f.take(), [
    { type: 'sheets', sheets: [sheet('Hidden'), sheet('Notes'), sheet('Notes (2)'), sheet('Data')], activeSheet: 2 },
    { type: 'changed', sheets: [0, 1, 2, 3], canUndo: false, canRedo: false },
    { type: 'done', id: 'c', command: 'copySheet', index: 2, name: 'Notes (2)' },
  ]);
  await f.bridge.receive({ type: 'copySheet', id: 'c2', index: 0, at: 4, name: 'Last' });
  assert.deepEqual(f.take().pop(), { type: 'done', id: 'c2', command: 'copySheet', index: 4, name: 'Last' });
  // The colour of a tab is part of the list, as a CSS colour, for the sheets that have one.
  await f.bridge.receive({ type: 'setTabColor', id: 't', index: 0, color: '#ff8800' });
  const told = f.take();
  assert.deepEqual(told[0].sheets[0], { name: 'Hidden', state: 'visible', kind: 'worksheet', tabColor: '#FF8800' });
  assert.equal(told[0].sheets[1].tabColor, undefined);
  assert.deepEqual(told[told.length - 1], { type: 'done', id: 't', command: 'setTabColor' });
  await f.bridge.receive({ type: 'setTabColor', id: 't2', index: 0, color: null });
  assert.equal(f.take()[0].sheets[0].tabColor, undefined);
  assert.deepEqual(f.calls.filter((call) => /^(move|setSheet|copy|setTab)/.test(call)), ['moveSheet 0 2', 'setSheetState 0 visible', 'copySheet 1 {}', 'copySheet 0 {"at":4,"name":"Last"}', 'setTabColor 0 #ff8800', 'setTabColor 0 undefined']);
});

test('bridge sheets: the new commands refuse what is not in the contract, an index no sheet has, a read-only view and what the view refuses', async () => {
  const f = await loaded();
  const bad: [Record<string, unknown>, RegExp][] = [
    [{ type: 'moveSheet', to: 1 }, /"from" is missing/],
    [{ type: 'moveSheet', from: 0 }, /"to" is missing/],
    [{ type: 'moveSheet', from: 0, to: 1.5 }, /"to" must be a sheet index/],
    [{ type: 'setSheetState', index: 0 }, /"state" is missing/],
    [{ type: 'setSheetState', index: 0, state: 'gone' }, /"state" must be "visible", "hidden" or "veryHidden"/],
    [{ type: 'copySheet' }, /"index" is missing/],
    [{ type: 'copySheet', index: 0, name: 7 }, /"name" must be a string/],
    [{ type: 'setTabColor', index: 0, color: 'red' }, /"color" must be written #RRGGBB/],
    [{ type: 'setTabColor', index: 0, color: 7 }, /"color" must be a string/],
  ];
  for (const [message, reason] of bad) {
    await f.bridge.receive({ ...message, id: 'x' });
    const answer = f.take();
    assert.equal(answer.length, 1, JSON.stringify(message));
    assert.equal(answer[0].code, 'bad-message', JSON.stringify(message));
    assert.match(answer[0].message, reason);
  }
  assert.deepEqual(f.calls.filter((call) => /^(move|setSheet|copy|setTab)/.test(call)), []);
  await f.bridge.receive({ type: 'moveSheet', id: 'a', from: 3, to: 0 });
  await f.bridge.receive({ type: 'moveSheet', id: 'b', from: 0, to: 3 });
  await f.bridge.receive({ type: 'copySheet', id: 'c', index: 0, at: 4 });
  assert.deepEqual(f.take().map((message) => message.code), ['no-sheet', 'no-sheet', 'no-sheet']);
  await f.bridge.receive({ type: 'setSheetState', id: 'd', index: 2, state: 'hidden' });
  f.take();
  await f.bridge.receive({ type: 'setSheetState', id: 'e', index: 0, state: 'hidden' });
  assert.deepEqual(f.take(), [{ type: 'error', id: 'e', code: 'sheet-refused', message: 'Localized: The last visible sheet cannot be hidden' }]);
  await f.bridge.receive({ type: 'copySheet', id: 'f', index: 0, name: 'Notes' });
  assert.deepEqual(f.take(), [{ type: 'error', id: 'f', code: 'sheet-refused', message: 'Localized: Another sheet is called "Notes"' }]);

  const locked = await loaded({}, { readOnly: true });
  for (const message of [{ type: 'moveSheet', from: 0, to: 1 }, { type: 'setSheetState', index: 0, state: 'hidden' }, { type: 'copySheet', index: 0 }, { type: 'setTabColor', index: 0, color: '#000000' }]) {
    await locked.bridge.receive({ ...message, id: 'r' });
    assert.deepEqual(locked.take().map((answer) => answer.code), ['read-only'], message.type);
  }
  const bare = await loaded({ bare: true });
  await bare.bridge.receive({ type: 'moveSheet', id: 'g', from: 0, to: 1 });
  assert.deepEqual(bare.take().map((answer) => answer.code), ['sheet-refused']);
  const none = fake();
  await none.bridge.receive({ type: 'setTabColor', id: 'h', index: 0 });
  assert.equal(none.take().pop().code, 'not-loaded');
});
