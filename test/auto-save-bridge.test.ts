import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Decode, base64Encode, createBridge, type Bridge, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import { watchHiding } from '../src/embed/save-to';
import type { SaveCommand } from '../src/ui/save-command';
import { STRINGS_EN } from '../src/ui/strings';

// The bridge page for a `load` with `save: { auto }`: the page hands the file to
// the application by itself, in `saved` messages that say `by: "auto"`.

interface ViewOptions {
  onChange?(change?: { sheets: number[] }): void;
  saveCommand?: SaveCommand;
}

interface Fake {
  bridge: Bridge;
  sent: any[];
  views: ViewOptions[];
  notices: [kind: string, message: string][];
  timers: { callback: () => void; at: number; off: boolean }[];
  listeners: { [type: string]: ((event: any) => void)[] };
  typing: boolean;
  readOnly: boolean;
  commits: number;
  questions: number;
  format: 'xlsx' | 'docx';
  content: number[];
  clock: number;
  advance(ms: number): Promise<void>;
  edit(): void;
}

const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};

function fake(): Fake {
  const state: Fake = {
    bridge: undefined as unknown as Bridge,
    sent: [],
    views: [],
    notices: [],
    timers: [],
    listeners: {},
    typing: false,
    readOnly: false,
    commits: 0,
    questions: 0,
    format: 'xlsx',
    content: [1],
    clock: 50000,
    async advance(ms) {
      const end = state.clock + ms;
      for (let guard = 0; guard < 10000; guard++) {
        let next: Fake['timers'][number] | undefined;
        for (const timer of state.timers) if (!timer.off && timer.at <= end && (!next || timer.at < next.at)) next = timer;
        if (!next) break;
        state.clock = Math.max(state.clock, next.at);
        next.off = true;
        next.callback();
        await settle();
      }
      state.clock = end;
      await settle();
    },
    edit() {
      state.content = [state.content[0] + 1];
      const options = state.views[state.views.length - 1];
      if (options.onChange) options.onChange({ sheets: [0] });
    },
  };
  const showNotice = (message: string, kind: string = 'info') => void state.notices.push([kind, message]);
  const host: BridgeHost = {
    container: {
      addEventListener: (type: string, listener: (event: any) => void) => void (state.listeners[type] || (state.listeners[type] = [])).push(listener),
      removeEventListener: () => undefined,
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
        get readOnly() {
          return state.readOnly;
        },
        commitEditing: () => undefined,
        commitPending: () => {
          state.commits++;
          return true;
        },
        hasPendingInput: () => state.typing,
        localizeError: (error: Error) => error.message,
        showNotice,
        setReadOnly: (value: boolean) => void (state.readOnly = value),
        setLocale: () => undefined,
        destroy: () => undefined,
      } as any;
    },
    createDocumentView(_container, _doc, options) {
      state.views.push(options as ViewOptions);
      return {
        destroy: () => undefined,
        commitEditing: () => void state.commits++,
        hasPendingInput: () => state.typing,
        changeCount: 0,
        markSaved: () => undefined,
        canUndo: true,
        canRedo: false,
        showNotice,
      };
    },
    askOverwrite: () => void state.questions++,
    timer: (callback, ms) => void state.timers.push({ callback, at: state.clock + ms, off: false }),
    now: () => state.clock,
    frame: () => undefined,
  };
  state.bridge = createBridge(host);
  return state;
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const saves = (f: Fake) => f.sent.filter((message) => message.type === 'saved');
const states = (f: Fake) => f.sent.filter((message) => message.type === 'saveState').map((message) => `${message.dirty} ${message.saving}`);
const lastOf = (f: Fake) => f.sent.filter((message) => message.type !== 'saveState').pop();

test('bridge auto-save: the page hands the file over by itself, as "saved" by "auto" with an id of its own, and saveResult ends it', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, name: 'liste.xlsx', save: { auto: true } });
  const loaded = lastOf(f);
  assert.deepEqual([loaded.type, loaded.save, loaded.autoSave], ['loaded', true, true]);
  assert.equal(typeof (f.views[0].saveCommand as SaveCommand).run, 'function', 'the Save button is there as for save: true');

  f.edit();
  await f.advance(1999);
  assert.equal(saves(f).length, 0);
  await f.advance(1);
  assert.equal(saves(f).length, 1);
  const saved = saves(f)[0];
  assert.deepEqual([saved.by, saved.saveId, saved.format, saved.name, 'id' in saved, 'overwrite' in saved], ['auto', 'a1', 'xlsx', 'liste.xlsx', false, false]);
  assert.deepEqual(Array.from(base64Decode(saved.data)), [2]);
  assert.equal(f.commits, 0, 'nothing that is being typed was applied for it');
  assert.equal(states(f).pop(), 'true true');

  await f.bridge.receive({ type: 'saveResult', id: 'r', saveId: 'a1', ok: true });
  await settle();
  assert.deepEqual([lastOf(f).type, lastOf(f).command], ['done', 'saveResult']);
  assert.equal(states(f).pop(), 'false false');
  assert.deepEqual(f.notices, [], 'a save nobody asked for is not announced every time: the state says it');

  // The save of the user that follows has the next number, and its own letter.
  f.edit();
  (f.views[0].saveCommand as SaveCommand).run();
  await settle();
  assert.deepEqual([saves(f)[1].by, saves(f)[1].saveId], ['user', 'u2']);
  assert.equal(f.commits > 0, true, 'the save of the user applies what is being typed');
});

test('bridge auto-save: the times can be given; values that cannot be are a bad message, and the file is not loaded', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'a', data: DATA, save: { auto: { delayMs: 500, maxWaitMs: 1000 } } });
  f.edit();
  await f.advance(500);
  assert.equal(saves(f).length, 1);

  for (const save of [{ auto: { delayMs: 10 } }, { auto: 'yes' }, { auto: { delayMs: 1000, maxWaitMs: 10 } }, 'yes', 1]) {
    const g = fake();
    await g.bridge.receive({ type: 'load', id: 'b', data: DATA, save });
    assert.deepEqual([lastOf(g).type, lastOf(g).code, lastOf(g).id], ['error', 'bad-message', 'b'], JSON.stringify(save));
    assert.equal(g.views.length, 0);
  }
});

test('bridge auto-save: save: {} and save: { auto: false } are a Save button and nothing more', async () => {
  for (const save of [{}, { auto: false }, true]) {
    const f = fake();
    await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save });
    assert.deepEqual([lastOf(f).save, 'autoSave' in lastOf(f)], [true, false]);
    f.edit();
    await f.advance(120000);
    assert.equal(saves(f).length, 0, JSON.stringify(save));
  }
});

test('bridge auto-save: nothing is handed over while something is being typed, in a workbook and in a document', async () => {
  for (const format of ['xlsx', 'docx'] as const) {
    const f = fake();
    f.format = format;
    await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
    f.edit();
    f.typing = true;
    await f.advance(60000);
    assert.equal(saves(f).length, 0, format);
    f.typing = false;
    await f.advance(2000);
    assert.equal(saves(f).length, 1, format);
    assert.equal(saves(f)[0].by, 'auto');
    assert.equal(f.commits, 0, format);
  }
});

test('bridge auto-save: a result that says "conflict" is told, not asked about, and stops it; the Save of the user is what asks', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
  f.edit();
  await f.advance(2000);
  await f.bridge.receive({ type: 'saveResult', saveId: 'a1', ok: false, conflict: true, message: 'changed by Ada' });
  await settle();
  assert.equal(f.questions, 0, 'no question by itself');
  assert.deepEqual(f.notices.pop(), ['error', STRINGS_EN.saveConflict]);
  assert.equal(states(f).pop(), 'true false');
  f.edit();
  await f.advance(600000);
  assert.equal(saves(f).length, 1, 'nothing is handed over by itself after a conflict');

  (f.views[0].saveCommand as SaveCommand).run();
  await settle();
  assert.deepEqual([saves(f)[1].by, saves(f)[1].saveId, 'overwrite' in saves(f)[1]], ['user', 'u2', false]);
  await f.bridge.receive({ type: 'saveResult', saveId: 'u2', ok: false, conflict: true });
  await settle();
  assert.equal(f.questions, 1);
});

test('bridge auto-save: a result that says "not saved" stops it; no result in time is tried again', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
  f.edit();
  await f.advance(2000);
  // No saveResult: sixty seconds later the save has failed, and five seconds after that it is tried again.
  await f.advance(60000);
  assert.equal(saves(f).length, 1);
  assert.deepEqual(f.notices.pop(), ['error', STRINGS_EN.saveTimedOut]);
  await f.advance(5000);
  assert.equal(saves(f).length, 2);
  assert.equal(saves(f)[1].saveId, 'a2');
  await f.bridge.receive({ type: 'saveResult', saveId: 'a2', ok: false, message: 'disk full' });
  await settle();
  f.edit();
  await f.advance(600000);
  assert.equal(saves(f).length, 2, 'what the application refused is not offered again by itself');
});

test('bridge auto-save: a file that is only read hands nothing over; setReadOnly(false) starts the delay again', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
  f.edit();
  await f.bridge.receive({ type: 'setReadOnly', readOnly: true });
  await f.advance(60000);
  assert.equal(saves(f).length, 0);
  await f.bridge.receive({ type: 'setReadOnly', readOnly: false });
  await f.advance(1999);
  assert.equal(saves(f).length, 0);
  await f.advance(1);
  assert.equal(saves(f).length, 1);
});

test('bridge auto-save: a page that is hidden hands over what waits at once; a file that was closed hands nothing over', async () => {
  const f = fake();
  await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
  f.bridge.pageHidden();
  await settle();
  assert.equal(saves(f).length, 0, 'nothing to save');
  f.edit();
  f.bridge.pageHidden();
  await settle();
  assert.equal(saves(f).length, 1);
  await f.bridge.receive({ type: 'saveResult', saveId: 'a1', ok: true });
  await settle();

  f.edit();
  await f.bridge.receive({ type: 'close' });
  f.bridge.pageHidden();
  await f.advance(60000);
  assert.equal(saves(f).length, 1);
});

test('watchHiding: tells when the page is hidden, not when it is shown again, and stops when asked', () => {
  const listeners: (() => void)[] = [];
  const doc = {
    visibilityState: 'visible',
    addEventListener: (_type: 'visibilitychange', listener: () => void) => void listeners.push(listener),
    removeEventListener: (_type: 'visibilitychange', listener: () => void) => void listeners.splice(listeners.indexOf(listener), 1),
  };
  let hidden = 0;
  const stop = watchHiding(doc, () => void hidden++);
  doc.visibilityState = 'hidden';
  listeners[0]();
  doc.visibilityState = 'visible';
  listeners[0]();
  assert.equal(hidden, 1);
  stop();
  assert.equal(listeners.length, 0);
});
