import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Encode, createBridge, type BridgeHost } from '../src/embed/bridge';
import type { ConflictQuestion } from '../src/embed/save-to';
import type { OfficeFile } from '../src/index';
import type { SaveCommand } from '../src/ui/save-command';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

// The bridge page and a save the application could not make because the file it holds was changed
// meanwhile: `saveResult` says so with `conflict: true`, the page asks the user, and the answer
// "overwrite" hands the file over once more in a `saved` that says `overwrite: true`.

interface ViewOptions {
  readOnly?: boolean;
  onChange?(change?: { sheets: number[] }): void;
  saveCommand?: SaveCommand;
}

interface Fake {
  host: BridgeHost;
  sent: any[];
  views: ViewOptions[];
  notices: [kind: string, message: string][];
  /** The questions the page asked, with the language each was asked in and how often it was taken away. */
  asked: { question: ConflictQuestion; locale: string | undefined; closed: number }[];
  format: 'xlsx' | 'docx';
  content: number[];
}

function fake(ask = true): Fake {
  const state: Fake = { host: undefined as unknown as BridgeHost, sent: [], views: [], notices: [], asked: [], format: 'xlsx', content: [1, 2, 3, 250] };
  const showNotice = (message: string, kind: string = 'info') => void state.notices.push([kind, message]);
  state.host = {
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
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
        commitPending: () => true,
        showNotice,
        setReadOnly: () => undefined,
        setLocale: () => undefined,
        destroy: () => undefined,
      } as any;
    },
    createDocumentView(_container, _doc, options) {
      state.views.push(options as ViewOptions);
      return { destroy: () => undefined, setReadOnly: () => undefined, commitEditing: () => undefined, changeCount: 0, markSaved: () => undefined, showNotice };
    },
    timer: () => undefined,
    frame: () => undefined,
  };
  if (ask) {
    state.host.askOverwrite = (question, locale) => {
      const entry = { question, locale, closed: 0 };
      state.asked.push(entry);
      return () => void entry.closed++;
    };
  }
  return state;
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};
const saves = (f: Fake) => f.sent.filter((message) => message.type === 'saved' && message.by === 'user');
const states = (f: Fake) => f.sent.filter((message) => message.type === 'saveState').map((message) => `${message.dirty} ${message.saving}`);
const errors = (f: Fake) => f.sent.filter((message) => message.type === 'error');

/** A file with a change whose save the application answered with a conflict. */
async function conflicted(f: Fake, more: object = {}, load: object = {}): Promise<{ bridge: ReturnType<typeof createBridge>; command: SaveCommand }> {
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true, name: 'liste.xlsx', ...load });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  assert.equal(saves(f).length, 1, 'the file was handed over');
  await bridge.receive({ type: 'saveResult', saveId: saves(f)[0].saveId, ok: false, conflict: true, ...more });
  await settle();
  return { bridge, command };
}

test('bridge: saveResult with conflict asks the user, in the language of the file, and says nothing new to the application', async () => {
  const f = fake();
  const { command } = await conflicted(f, { id: 'r' }, { locale: 'tr' });
  assert.equal(f.asked.length, 1, 'the question was asked');
  assert.equal(f.asked[0].locale, 'tr', 'in the language the file was loaded with');
  assert.equal(f.asked[0].question.detail, '', 'nothing was said of the conflict');
  assert.deepEqual(f.notices, [['error', STRINGS_TR.saveConflict]], 'the view says that the file was not saved, as a conflict');
  assert.deepEqual(f.sent.filter((message) => message.type === 'done'), [{ type: 'done', id: 'r', command: 'saveResult' }], 'the result is confirmed');
  assert.deepEqual(errors(f), [], 'and no error is sent');
  assert.deepEqual(states(f).slice(-2), ['true true', 'true false'], 'the application hears that the save is over and the file unsaved');
  assert.deepEqual([command.dirty, command.saving], [true, false], 'the file is unsaved, and nothing is under way');
  assert.equal('overwrite' in saves(f)[0], false, 'the first save says nothing of overwriting');
  assert.equal(saves(f).length, 1, 'nothing is handed over while the question is open');
});

test('bridge: the answer "overwrite" hands the file over again with a new saveId and overwrite: true; its result ends the save', async () => {
  const f = fake();
  const { bridge, command } = await conflicted(f, { message: `changed by Ada${String.fromCharCode(10)}<i>` });
  assert.equal(f.asked[0].question.detail, 'changed by Ada <i>', 'what the application said is shown as one line of text');
  f.content = [7, 7];
  f.asked[0].question.answer(true);
  await settle();
  assert.equal(f.asked[0].closed, 1, 'the question is taken away');
  assert.equal(saves(f).length, 2, 'the file is handed over once more');
  const again = saves(f)[1];
  assert.deepEqual(again, { type: 'saved', by: 'user', saveId: again.saveId, overwrite: true, data: base64Encode(new Uint8Array([7, 7])), format: 'xlsx', name: 'liste.xlsx' }, 'as a save of the user that overwrites, with the file as it is now');
  assert.notEqual(again.saveId, saves(f)[0].saveId, 'under an id of its own');
  assert.deepEqual([command.dirty, command.saving], [true, true], 'the overwrite is under way');
  await bridge.receive({ type: 'saveResult', saveId: again.saveId, ok: true });
  await settle();
  assert.deepEqual([command.dirty, command.saving], [false, false], 'saved');
  assert.deepEqual(f.notices[f.notices.length - 1], ['info', STRINGS_EN.saved], 'and the user is told');
  assert.deepEqual(states(f).slice(-3), ['true false', 'true true', 'false false'], 'the application is told the state as of any save');
  assert.deepEqual(errors(f), [], 'no error was sent');
});

test('bridge: the answer "cancel" sends no message; the next save is handed over without overwrite and may be asked about again', async () => {
  const f = fake();
  const { bridge, command } = await conflicted(f);
  const before = f.sent.length;
  f.asked[0].question.answer(false);
  await settle();
  assert.equal(f.sent.length, before, 'nothing is sent for a question that was cancelled');
  assert.deepEqual([command.dirty, command.saving], [true, false], 'the file stays unsaved');
  command.run();
  await settle();
  assert.equal(saves(f).length, 2, 'the next save is handed over');
  assert.equal('overwrite' in saves(f)[1], false, 'without overwrite: the application decides again');
  assert.equal(f.asked.length, 1, 'nothing is asked before the application has answered');
  await bridge.receive({ type: 'saveResult', saveId: saves(f)[1].saveId, ok: false, conflict: true });
  await settle();
  assert.equal(f.asked.length, 2, 'the same answer asks again');
});

test('bridge: while the question is open the Save button hands nothing over and asks nothing more', async () => {
  const f = fake();
  const { command } = await conflicted(f);
  command.run();
  command.run();
  await settle();
  assert.equal(saves(f).length, 1, 'nothing is handed over');
  assert.equal(f.asked.length, 1, 'one question');
});

test('bridge: an overwrite the application answers with a conflict again is told and not asked about', async () => {
  const f = fake();
  const { bridge, command } = await conflicted(f);
  f.asked[0].question.answer(true);
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: saves(f)[1].saveId, ok: false, conflict: true });
  await settle();
  assert.equal(f.asked.length, 1, 'no second question for a save that was told to overwrite');
  assert.deepEqual([command.dirty, command.saving], [true, false], 'the file stays unsaved');
  assert.deepEqual(f.notices[f.notices.length - 1], ['error', STRINGS_EN.saveConflict], 'the conflict is told');
});

test('bridge: without the conflict field, with conflict false, and with ok true, a result is what it always was', async () => {
  const plain = fake();
  const first = createBridge(plain.host);
  await first.receive({ type: 'load', data: DATA, save: true });
  plain.views[0].onChange!({ sheets: [0] });
  (plain.views[0].saveCommand as SaveCommand).run();
  await settle();
  await first.receive({ type: 'saveResult', saveId: saves(plain)[0].saveId, ok: false, message: 'HTTP 412' });
  await settle();
  assert.equal(plain.asked.length, 0, 'no question without the field');
  assert.deepEqual(plain.notices, [['error', STRINGS_EN.saveFailed('HTTP 412')]], 'the failure is told as before');

  const off = fake();
  const second = createBridge(off.host);
  await second.receive({ type: 'load', data: DATA, save: true });
  off.views[0].onChange!({ sheets: [0] });
  (off.views[0].saveCommand as SaveCommand).run();
  await settle();
  await second.receive({ type: 'saveResult', saveId: saves(off)[0].saveId, ok: false, conflict: false });
  await settle();
  assert.equal(off.asked.length, 0, 'no question for conflict: false');
  assert.deepEqual(off.notices, [['error', STRINGS_EN.saveFailed('')]], 'a plain failure');

  const fine = fake();
  const third = createBridge(fine.host);
  await third.receive({ type: 'load', data: DATA, save: true });
  fine.views[0].onChange!({ sheets: [0] });
  (fine.views[0].saveCommand as SaveCommand).run();
  await settle();
  await third.receive({ type: 'saveResult', saveId: saves(fine)[0].saveId, ok: true, conflict: true });
  await settle();
  assert.equal(fine.asked.length, 0, 'a save that succeeded is no conflict');
  assert.deepEqual(fine.notices, [['info', STRINGS_EN.saved]], 'it was saved');
});

test('bridge: a conflict that is no boolean is a bad message, and the save goes on waiting for its result', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!({ sheets: [0] });
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: saves(f)[0].saveId, id: 'r', ok: false, conflict: 'yes' });
  await settle();
  assert.deepEqual([errors(f)[0].code, errors(f)[0].id], ['bad-message', 'r'], 'refused as a bad message');
  assert.equal(command.saving, true, 'the save still waits');
  assert.equal(f.asked.length, 0, 'nothing is asked');
});

test('bridge: a page that cannot ask tells a conflict as a conflict and nothing more', async () => {
  const f = fake(false);
  const { command } = await conflicted(f);
  assert.deepEqual(f.notices, [['error', STRINGS_EN.saveConflict]], 'the view says what kind of failure it was');
  assert.deepEqual([command.dirty, command.saving], [true, false], 'unsaved');
});

test('bridge: the question goes away with a new load, with close, with setReadOnly and with destroy; a late answer hands nothing over', async () => {
  for (const how of ['load', 'close', 'setReadOnly', 'destroy'] as const) {
    const f = fake();
    const { bridge } = await conflicted(f);
    if (how === 'load') await bridge.receive({ type: 'load', data: DATA, save: true });
    else if (how === 'close') await bridge.receive({ type: 'close' });
    else if (how === 'setReadOnly') await bridge.receive({ type: 'setReadOnly', readOnly: true });
    else bridge.destroy();
    await settle();
    assert.equal(f.asked[0].closed, 1, `${how}: the question is taken away`);
    f.asked[0].question.answer(true);
    await settle();
    assert.equal(saves(f).length, 1, `${how}: an answer that comes late hands nothing over`);
  }
});

test('bridge: a document is asked about as a workbook is', async () => {
  const f = fake();
  f.format = 'docx';
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, save: true });
  const command = f.views[0].saveCommand as SaveCommand;
  f.views[0].onChange!();
  command.run();
  await settle();
  await bridge.receive({ type: 'saveResult', saveId: saves(f)[0].saveId, ok: false, conflict: true });
  await settle();
  assert.equal(f.asked.length, 1, 'the question was asked');
  f.asked[0].question.answer(true);
  await settle();
  assert.deepEqual([saves(f)[1].overwrite, saves(f)[1].format], [true, 'docx'], 'the overwrite is handed over');
});

test('bridge: a file that is only read is not asked about; its conflict is told, and nothing is handed over again', async () => {
  for (const format of ['xlsx', 'docx'] as const) {
    const f = fake();
    f.format = format;
    // A view that knows whether it is read-only, as the views of the engine do.
    let readOnly = false;
    const createView = f.host.createView;
    f.host.createView = (container, workbook, options) => {
      const view = createView.call(f.host, container, workbook, options) as any;
      view.setReadOnly = (value: boolean) => void (readOnly = value);
      Object.defineProperty(view, 'readOnly', { get: () => readOnly });
      return view;
    };
    const bridge = createBridge(f.host);
    await bridge.receive({ type: 'load', data: DATA, save: true });
    const command = f.views[0].saveCommand as SaveCommand;
    f.views[0].onChange!({ sheets: [0] });
    await bridge.receive({ type: 'setReadOnly', readOnly: true });
    // The key still saves what was changed before (as it always did).
    command.run();
    await settle();
    assert.equal(saves(f).length, 1, `${format}: the file is handed over`);
    await bridge.receive({ type: 'saveResult', saveId: saves(f)[0].saveId, ok: false, conflict: true });
    await settle();
    assert.equal(f.asked.length, 0, `${format}: nothing is asked over a view that is only read`);
    assert.deepEqual(f.notices, [['error', STRINGS_EN.saveConflict]], `${format}: the conflict is told`);
    assert.deepEqual([command.dirty, command.saving], [true, false], `${format}: the file stays unsaved, and nothing is under way`);
    assert.equal(saves(f).length, 1, `${format}: nothing more is handed over`);

    // Editable again: the next conflict is asked about.
    await bridge.receive({ type: 'setReadOnly', readOnly: false });
    command.run();
    await settle();
    await bridge.receive({ type: 'saveResult', saveId: saves(f)[1].saveId, ok: false, conflict: true });
    await settle();
    assert.equal(f.asked.length, 1, `${format}: a view that can be edited is asked`);
  }
});

test('bridge: whoever asks is given the element of the view that is shown, and nothing where the view names none', async () => {
  const withElement = fake();
  const element = { nodeType: 1 } as unknown as HTMLElement;
  const createView = withElement.host.createView;
  withElement.host.createView = (container, workbook, options) => Object.assign(createView.call(withElement.host, container, workbook, options) as object, { element }) as any;
  const given: unknown[] = [];
  withElement.host.askOverwrite = (_question, _locale, shownElement) => void given.push(shownElement);
  await conflicted(withElement);
  assert.deepEqual(given, [element], 'the element of the view');

  const without = fake();
  const none: unknown[] = [];
  without.host.askOverwrite = (_question, _locale, shownElement) => void none.push(shownElement);
  await conflicted(without);
  assert.deepEqual(none, [undefined], 'a view without one leaves the place to who asks');
});
