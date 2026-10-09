import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Decode, base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import { open as openFile, type OfficeFile } from '../src/index';
import type { SaveCommand } from '../src/ui/save-command';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';
import { cat, number, workbookStream, xlsFile } from './xls-fixtures';

// The bridge page for a `load` with `save`: the Save button and Ctrl+S hand the
// file to the application in a `saved` message that says it comes from the
// user, and `saveResult` says what became of it.

interface ViewOptions {
  readOnly?: boolean;
  onChange?(change?: { sheets: number[] }): void;
  saveCommand?: SaveCommand;
}

interface Fake {
  host: BridgeHost;
  sent: any[];
  /** The options of every view that was made, workbook or document. */
  views: ViewOptions[];
  notices: [kind: string, message: string][];
  timers: { callback: () => void; ms: number }[];
  /** What the container was asked to listen to. */
  listeners: { [type: string]: ((event: any) => void)[] };
  /** What the view answers when it is asked to apply what is being typed. */
  pending: 'none' | 'stuck' | 'throws';
  commits: number;
  format: 'xlsx' | 'docx' | 'xls';
  content: number[];
  destroyed: number;
}

function fake(): Fake {
  const state: Fake = { host: undefined as unknown as BridgeHost, sent: [], views: [], notices: [], timers: [], listeners: {}, pending: 'none', commits: 0, format: 'xlsx', content: [1, 2, 3, 250], destroyed: 0 };
  const showNotice = (message: string, kind: string = 'info') => void state.notices.push([kind, message]);
  state.host = {
    container: {
      addEventListener: (type: string, listener: (event: any) => void) => void (state.listeners[type] || (state.listeners[type] = [])).push(listener),
      removeEventListener: (type: string, listener: (event: any) => void) => {
        const list = state.listeners[type] || [];
        const at = list.indexOf(listener);
        if (at >= 0) list.splice(at, 1);
      },
    } as unknown as HTMLElement,
    version: '9.9.9',
    post(json) {
      state.sent.push(JSON.parse(json));
      return true;
    },
    async open() {
      const format = state.format;
      const file: any = { format, kind: format === 'docx' ? 'document' : 'spreadsheet', save: async () => new Uint8Array(state.content) };
      file.workbook = async () => ({ sheets: [{ name: 'Veri', state: 'visible', kind: 'worksheet' }] });
      if (format === 'docx') file.document = async () => ({ body: [], settings: {} });
      return file as OfficeFile;
    },
    createView(_container, _workbook, options) {
      state.views.push(options as ViewOptions);
      return {
        history: { canUndo: true, canRedo: false },
        activeSheet: 0,
        isEditing: false,
        commitEditing: () => undefined,
        commitPending: () => {
          state.commits++;
          if (state.pending === 'throws') throw new Error('The formula cannot be read');
          return state.pending === 'none';
        },
        localizeError: (error: Error) => `localized: ${error.message}`,
        showNotice,
        setReadOnly: () => undefined,
        setLocale: () => undefined,
        destroy: () => void state.destroyed++,
      } as any;
    },
    createDocumentView(_container, _doc, options) {
      state.views.push(options as ViewOptions);
      let changes = 0;
      return {
        destroy: () => void state.destroyed++,
        setLocale: () => undefined,
        commitEditing: () => void state.commits++,
        get changeCount() {
          return changes;
        },
        markSaved: () => void (changes = 0),
        canUndo: true,
        canRedo: false,
        showNotice,
      };
    },
    timer: (callback, ms) => void state.timers.push({ callback, ms }),
    frame: () => undefined,
  };
  return state;
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const types = (f: Fake) => f.sent.map((message) => message.type);
const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};
const last = (f: Fake) => f.sent[f.sent.length - 1];
/** The id of the last save the user started: what its `saveResult` names. */
const lastId = (f: Fake): string => {
  const saves = f.sent.filter((message) => message.type === 'saved' && message.by === 'user');
  return saves.length > 0 ? saves[saves.length - 1].saveId : 'none';
};
const key = (more: object = {}) => {
  const event = { key: 's', ctrlKey: true, metaKey: false, altKey: false, shiftKey: false, prevented: false, preventDefault: () => void (event.prevented = true), ...more };
  return event;
};
const press = (f: Fake, event: ReturnType<typeof key>) => {
  for (const listener of f.listeners.keydown || []) listener(event);
  return event.prevented;
};

test('bridge: a load with save gives the view a save command and says so in loaded; without it nothing changes', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'a', data: DATA });
  assert.equal(f.views[0].saveCommand, undefined);
  assert.equal('save' in f.sent[1], false);
  assert.equal(press(f, key()), false, 'Ctrl+S is left to the browser');

  await bridge.receive({ type: 'load', id: 'b', data: DATA, save: true, name: 'liste.xlsx' });
  const command = f.views[1].saveCommand as SaveCommand;
  assert.equal(typeof command.run, 'function');
  assert.deepEqual([command.dirty, command.saving], [false, false]);
  assert.equal(f.sent[2].type, 'loaded');
  assert.equal(f.sent[2].save, true);

  await bridge.receive({ type: 'load', id: 'c', data: DATA, save: 'yes' });
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'c']);
});

test('bridge: the user saves: saved carries the file, its format and by "user", without an id; saveResult ok makes it saved', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'l', data: DATA, save: true, name: 'liste.xlsx', locale: 'tr' });
  const options = f.views[0];
  const command = options.saveCommand as SaveCommand;
  const seen: string[] = [];
  command.watch(() => seen.push(`${command.dirty} ${command.saving}`));

  // Nothing was changed: there is nothing to hand over.
  command.run();
  await settle();
  assert.deepEqual(types(f), ['ready', 'loaded']);

  options.onChange!({ sheets: [0] });
  assert.deepEqual([command.dirty, command.saving], [true, false]);
  command.run();
  await settle();
  assert.deepEqual(last(f), { type: 'saved', by: 'user', saveId: 'u1', data: base64Encode(new Uint8Array([1, 2, 3, 250])), format: 'xlsx', name: 'liste.xlsx' });
  assert.equal('id' in last(f), false);
  assert.deepEqual([command.dirty, command.saving], [true, true]);
  assert.deepEqual(f.timers.filter((entry) => entry.ms === 60000).length, 1, 'the page waits a minute for the result');

  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.deepEqual([command.dirty, command.saving], [false, false]);
  assert.deepEqual(f.notices, [['info', STRINGS_TR.saved]]);
  assert.deepEqual(seen, ['true false', 'true true', 'false false']);
  // The time limit that is still waiting has nothing left to end.
  for (const entry of f.timers) if (entry.ms === 60000) entry.callback();
  await settle();
  assert.deepEqual([command.dirty, command.saving, f.notices.length], [false, false, 1]);
});

test('bridge: saveResult that is not ok: could not be saved, with one line of what the application said; the file is still unsaved', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), id: 'r', ok: false, message: `Disk full${String.fromCharCode(10)}<b>${'x'.repeat(400)}` });
  await settle();
  assert.deepEqual(last(f), { type: 'done', id: 'r', command: 'saveResult' });
  assert.deepEqual([command.dirty, command.saving], [true, false]);
  assert.equal(f.notices.length, 1);
  assert.equal(f.notices[0][0], 'error');
  assert.equal(f.notices[0][1], STRINGS_EN.saveFailed(`Disk full <b>${'x'.repeat(187)}`), 'cut to 200 characters, on one line, and only ever shown as text');

  // Without a message.
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: false });
  await settle();
  assert.deepEqual(f.notices[1], ['error', 'The file could not be saved']);
  // And the next one succeeds.
  command.run();
  await settle();
  assert.equal(f.sent.filter((message) => message.type === 'saved').length, 3);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.deepEqual([command.dirty, f.notices[2]], [false, ['info', 'Saved']]);
});

test('bridge: no saveResult in time: the save has failed, and a result that comes late changes nothing', async () => {
  const f = fake();
  f.host.saveResultTimeoutMs = 9000;
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  const limit = f.timers.filter((entry) => entry.ms === 9000);
  assert.equal(limit.length, 1);
  limit[0].callback();
  await settle();
  assert.deepEqual([command.dirty, command.saving], [true, false]);
  assert.deepEqual(f.notices, [['error', STRINGS_EN.saveTimedOut]]);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), id: 'late', ok: true });
  await settle();
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'late']);
  assert.deepEqual([command.dirty, command.saving, f.notices.length], [true, false, 1]);
});

test('bridge review 2: a result that comes late is not taken for the result of the next save', async () => {
  const f = fake();
  f.host.saveResultTimeoutMs = 9000;
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  const first = lastId(f);
  f.timers.filter((entry) => entry.ms === 9000)[0].callback();
  await settle();
  // The user tries again, and now the answer to the first save arrives.
  command.run();
  await settle();
  const second = lastId(f);
  assert.deepEqual([first, second], ['u1', 'u2']);
  await bridge.receive({ type: 'saveResult', saveId: first, ok: true });
  await settle();
  assert.deepEqual([command.dirty, command.saving], [true, true], 'the second save still waits for its own result');
  assert.deepEqual([last(f).type, last(f).code], ['error', 'bad-message']);
  // Its own result is the one that counts.
  await bridge.receive({ type: 'saveResult', saveId: second, ok: false, message: 'no room' });
  await settle();
  assert.deepEqual([command.dirty, command.saving], [true, false]);
  assert.deepEqual(f.notices[1], ['error', 'The file could not be saved (no room)']);
  // A result without an id, or with one that is no text, answers nothing.
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', ok: true });
  await bridge.receive({ type: 'saveResult', saveId: 3, ok: true });
  await settle();
  assert.equal(command.saving, true);
  assert.deepEqual(f.sent.slice(-2).map((message) => [message.type, message.code]), [
    ['error', 'bad-message'],
    ['error', 'bad-message'],
  ]);
});

test('bridge: saveResult must say ok and name its save; one that nothing waits for is refused and does nothing', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'saveResult', saveId: 'u1', id: 'early', ok: true });
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'early']);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), id: 'bad' });
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'bad']);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), id: 'worse', ok: 'true' });
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'worse']);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), id: 'text', ok: false, message: 7 });
  assert.deepEqual([last(f).type, last(f).code, last(f).id], ['error', 'bad-message', 'text']);
});

test('bridge: a save the application asks for works as before, beside the user\'s: it has its id, no "by", and leaves the changes unsaved', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  await bridge.receive({ type: 'save', id: 's' });
  assert.deepEqual(last(f), { type: 'saved', id: 's', data: base64Encode(new Uint8Array([1, 2, 3, 250])) });
  assert.equal(command.dirty, true, 'the application was given the bytes; whether it stored them the page does not know');
  // While the user's save waits for its result, the application's own save is answered.
  command.run();
  await settle();
  await bridge.receive({ type: 'save', id: 't' });
  assert.deepEqual(f.sent.slice(-2).map((message) => [message.type, message.by, message.id]), [
    ['saved', 'user', undefined],
    ['saved', undefined, 't'],
  ]);
  assert.equal(command.saving, true);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.deepEqual([command.dirty, command.saving], [false, false]);
});

test('bridge: what is typed while the application stores the file is not called saved; two presses hand the file over once', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  const saved = () => f.sent.filter((message) => message.type === 'saved');
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  command.run();
  await settle();
  assert.equal(saved().length, 1);
  f.views[0].onChange!({ sheets: [0] });
  f.content = [7, 7];
  command.run();
  await settle();
  assert.equal(saved().length, 1, 'the next save waits for the result of the one under way');
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.equal(saved().length, 2);
  assert.deepEqual(Array.from(base64Decode(saved()[1].data)), [7, 7]);
  assert.deepEqual([command.dirty, command.saving, f.notices.length], [true, true, 0]);
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.deepEqual([command.dirty, command.saving, f.notices], [false, false, [['info', 'Saved']]]);
});

test('bridge: Ctrl+S and Cmd+S in the page save as the button does, and the browser does not get the key', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  f.views[0].onChange!({ sheets: [0] });
  assert.equal(press(f, key()), true);
  await settle();
  assert.equal(last(f).by, 'user');
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  f.views[0].onChange!({ sheets: [0] });
  assert.equal(press(f, key({ ctrlKey: false, metaKey: true, key: 'S' })), true);
  await settle();
  assert.equal(f.sent.filter((message) => message.type === 'saved').length, 2);
  // Other keys are not taken.
  for (const other of [key({ ctrlKey: false }), key({ shiftKey: true }), key({ altKey: true }), key({ key: 'a' })]) assert.equal(press(f, other), false, JSON.stringify(other));

  bridge.destroy();
  assert.deepEqual(f.listeners.keydown, [], 'the listener goes with the bridge');
});

test('bridge: a file that cannot be written has no save command, and its Ctrl+S is left alone', async () => {
  const f = fake();
  f.format = 'xls';
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  assert.equal(f.sent[1].type, 'loaded');
  assert.equal(f.sent[1].readOnly, true);
  assert.equal('save' in f.sent[1], false);
  assert.equal(f.views[0].saveCommand, undefined);
  assert.equal(press(f, key()), false);
});

test('bridge: what is being typed and cannot be applied keeps the user\'s save from being handed over', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  f.pending = 'stuck';
  command.run();
  await settle();
  assert.deepEqual(types(f), ['ready', 'loaded', 'changed']);
  assert.deepEqual(f.notices, [], 'the view says by itself why a note or a name stays open');
  f.pending = 'throws';
  command.run();
  await settle();
  assert.deepEqual(types(f), ['ready', 'loaded', 'changed']);
  assert.deepEqual(f.notices, [['error', 'localized: The formula cannot be read']]);
  f.pending = 'none';
  command.run();
  await settle();
  assert.equal(last(f).by, 'user');
});

test('bridge: closing the file, or loading another, ends the wait for a result; nothing is said about the old file', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const first = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  first.run();
  await settle();
  await bridge.receive({ type: 'close' });
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: false, message: 'too late' });
  for (const entry of f.timers) if (entry.ms === 60000) entry.callback();
  await settle();
  assert.deepEqual(f.notices, []);
  first.run();
  await settle();
  assert.equal(f.sent.filter((message) => message.type === 'saved').length, 1, 'the command of a file that is closed saves nothing');

  await bridge.receive({ type: 'load', data: DATA, save: true });
  const second = f.views[1].saveCommand as SaveCommand;
  f.views[1].onChange!({ sheets: [0] });
  second.run();
  await settle();
  // The result of the save of the file that was closed arrives now: it is not the result of this one.
  await bridge.receive({ type: 'saveResult', saveId: 'u1', ok: true });
  await settle();
  assert.deepEqual([second.dirty, second.saving, f.notices], [true, true, []]);
  assert.deepEqual([last(f).type, last(f).code], ['error', 'bad-message']);
  assert.equal(lastId(f), 'u2');
  await bridge.receive({ type: 'saveResult', saveId: 'u2', ok: true });
  await settle();
  assert.deepEqual([second.dirty, f.notices], [false, [['info', 'Saved']]]);
});

test('bridge: setLocale changes the language a save is spoken of in', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true, locale: 'en' });
  const command = f.views[0].saveCommand as SaveCommand;
  await bridge.receive({ type: 'setLocale', locale: 'tr-TR' });
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: false });
  await settle();
  assert.deepEqual(f.notices, [['error', 'Dosya kaydedilemedi']]);
});

test('bridge: a document is saved by the user as a workbook is', async () => {
  const f = fake();
  f.format = 'docx';
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true, name: 'mektup.docx' });
  assert.deepEqual([f.sent[1].type, f.sent[1].kind, f.sent[1].save], ['loaded', 'document', true]);
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!();
  assert.equal(command.dirty, true);
  command.run();
  await settle();
  assert.deepEqual(last(f), { type: 'saved', by: 'user', saveId: 'u1', data: base64Encode(new Uint8Array([1, 2, 3, 250])), format: 'docx', name: 'mektup.docx' });
  assert.equal(f.commits, 2, 'a composition under way is ended before the file is taken');
  await bridge.receive({ type: 'saveResult', saveId: lastId(f), ok: true });
  await settle();
  assert.deepEqual([command.dirty, f.notices], [false, [['info', 'Saved']]]);

  // A document that is loaded read-only can be made editable later: it has the command all the same.
  await bridge.receive({ type: 'load', data: DATA, save: true, readOnly: true });
  assert.equal(typeof (f.views[1].saveCommand as SaveCommand).run, 'function');
});

test('bridge: a converted file handed over by the user names both formats', async () => {
  const f = fake();
  f.host.open = (bytes) => openFile(bytes);
  const bridge = createBridge(f.host);
  const xls = xlsFile(workbookStream({ sheets: [{ name: 'Old', records: cat(number(0, 0, 41)) }] }));
  await bridge.receive({ type: 'load', data: base64Encode(xls), name: 'old.xls', convert: true, save: true });
  assert.deepEqual([f.sent[1].type, f.sent[1].format, f.sent[1].convertedFrom, f.sent[1].save], ['loaded', 'xlsx', 'xls', true]);
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  for (let i = 0; i < 200 && last(f).type !== 'saved'; i++) await new Promise((resolve) => setTimeout(resolve, 5));
  const saved = last(f);
  assert.deepEqual({ ...saved, data: '' }, { type: 'saved', by: 'user', saveId: 'u1', data: '', format: 'xlsx', name: 'old.xls', convertedFrom: 'xls' });
  assert.equal((await openFile(base64Decode(saved.data))).format, 'xlsx');
  // Without convert the same file cannot be written, and has no command.
  await bridge.receive({ type: 'load', data: base64Encode(xls), save: true });
  assert.equal(f.views[1].saveCommand, undefined);
});
