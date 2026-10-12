import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  SAVE_TIMEOUT_MS,
  SaveController,
  SaveError,
  bindSaveToPage,
  resolveAutoSave,
  type ConflictQuestion,
  type SaveHost,
  type SaveInfo,
  type SaveState,
  type SaveTarget,
} from '../src/embed/save-to';
import { watchDragEnd } from '../src/ui/drag-end';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

// Auto-save after its review: it only goes when the request can carry a condition, it says
// "saved" only after it said that something was wrong, a clock that is set back does not
// silence it, and what holds the bytes of a save lets go of them when the save is over.

interface Sent {
  at: number;
  init: RequestInit;
  answer(status: number, headers?: { [name: string]: string }): void;
  /** Answers 200 with an ETag no answer before it had. */
  ok(): void;
  breakOff(): void;
}

interface Rig {
  controller: SaveController;
  sent: Sent[];
  states: SaveState[];
  notices: [kind: string, message: string][];
  errors: unknown[];
  timers: { callback: (() => void) | undefined; at: number; off: boolean }[];
  questions: ConflictQuestion[];
  content: number[];
  pending: boolean;
  /** How often the host was asked whether something is being typed. */
  asked: number;
  clock: number;
  advance(ms: number): Promise<void>;
  edit(): void;
}

const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};

const request = (more: Partial<Extract<SaveTarget, { kind: 'request' }>> = {}): SaveTarget => ({
  kind: 'request',
  url: 'https://app.example.com/files/7',
  method: 'PUT',
  headers: [],
  credentials: 'same-origin',
  sameAsFile: true,
  timeoutMs: SAVE_TIMEOUT_MS,
  allowConverted: false,
  onConflict: 'ask',
  ...more,
});

let serial = 0;

function rig(target: SaveTarget = request(), more: Partial<SaveHost> = {}, info: SaveInfo = { format: 'docx' }): Rig {
  const state: Rig = {
    controller: undefined as unknown as SaveController,
    sent: [],
    states: [],
    notices: [],
    errors: [],
    timers: [],
    questions: [],
    content: [1],
    pending: false,
    asked: 0,
    clock: 100000,
    async advance(ms) {
      const end = state.clock + ms;
      for (let guard = 0; guard < 100000; guard++) {
        let next: Rig['timers'][number] | undefined;
        for (const timer of state.timers) if (!timer.off && timer.at <= end && (!next || timer.at < next.at)) next = timer;
        if (!next) break;
        state.clock = Math.max(state.clock, next.at);
        next.off = true;
        const run = next.callback;
        next.callback = undefined;
        if (run) run();
        await settle();
      }
      state.clock = end;
      await settle();
    },
    edit() {
      state.content = [state.content[0] + 1];
      state.controller.changed();
    },
  };
  const fakeFetch = (_url: unknown, init?: RequestInit) =>
    new Promise<Response>((resolve, reject) => {
      const answer = (status: number, headers: { [name: string]: string } = {}) => {
        const lower: { [name: string]: string } = {};
        for (const name of Object.keys(headers)) lower[name.toLowerCase()] = headers[name];
        resolve({ status, headers: { get: (name: string) => (name.toLowerCase() in lower ? lower[name.toLowerCase()] : null) } } as unknown as Response);
      };
      state.sent.push({ at: state.clock, init: init || {}, answer, ok: () => answer(200, { ETag: `"s${++serial}"` }), breakOff: () => reject(new TypeError('Failed to fetch')) });
    });
  const host: SaveHost = {
    target,
    writable: () => true,
    commit: () => true,
    bytes: async () => new Uint8Array(state.content),
    bytesAsIs: async () => new Uint8Array(state.content),
    pending: () => {
      state.asked++;
      return state.pending;
    },
    readOnly: () => false,
    info: () => info,
    strings: () => STRINGS_EN,
    notice: (message, kind) => state.notices.push([kind, message]),
    onState: (value) => state.states.push(value),
    onError: (error) => state.errors.push(error),
    askOverwrite: (question) => void state.questions.push(question),
    etag: '"v1"',
    fetch: fakeFetch as typeof fetch,
    now: () => state.clock,
    timer: (callback, ms) => {
      const entry: Rig['timers'][number] = { callback, at: state.clock + ms, off: false };
      state.timers.push(entry);
      // As clearTimeout: the callback is let go of.
      return () => {
        entry.off = true;
        entry.callback = undefined;
      };
    },
    autoSave: { delayMs: 2000, maxWaitMs: 30000 },
    ...more,
  };
  state.controller = new SaveController(host);
  return state;
}

const headersOf = (sent: Sent) => sent.init.headers as { [name: string]: string };
const codes = (r: Rig) => r.errors.map((error) => (error instanceof SaveError ? error.code : String(error)));

// ------------------------------------------------------------ 1: only with a condition

test('auto-save review 1: an answer without an ETag leaves nothing to guard the next save with: auto-save sends nothing, says so once, and the user can still save', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  assert.equal(headersOf(r.sent[0])['If-Match'], '"v1"');
  // The server takes the file and names no version of it.
  r.sent[0].answer(200);
  await settle();
  assert.equal(r.controller.currentETag, undefined);
  assert.equal(r.controller.state.autoPaused, true, 'the state says that auto-save waits');

  // Someone else may have written the file since: nothing goes over it unasked.
  r.edit();
  await r.advance(60000);
  assert.equal(r.sent.length, 1, 'no request without a condition');
  assert.deepEqual(r.notices, [['error', STRINGS_EN.autoSaveNoVersion]], 'said once, as a notice');
  assert.deepEqual(r.errors, [], 'no save failed: onError hears nothing');
  assert.deepEqual([r.controller.state.auto, r.controller.state.autoPaused, r.controller.state.dirty, 'error' in r.controller.state], [true, true, true, false]);
  assert.equal(r.states[r.states.length - 1].autoPaused, true, 'onSaveState was told');
  r.edit();
  await r.advance(60000);
  assert.equal(r.notices.length, 1, 'not said again for every change');

  // The user saves, as always; the answer names a version, and auto-save is back.
  const saved = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 2);
  assert.equal('If-Match' in headersOf(r.sent[1]), false, 'the save of the user goes as it did before auto-save existed');
  r.sent[1].answer(200, { ETag: '"v7"' });
  assert.equal(await saved, true);
  assert.equal('autoPaused' in r.controller.state, false);
  r.edit();
  await r.advance(2000);
  assert.equal(r.sent.length, 3);
  assert.equal(headersOf(r.sent[2])['If-Match'], '"v7"');
});

test('auto-save review 1: a weak ETag, a file that came without one and a file saved to another address are not guarded either', async () => {
  const weak = rig();
  weak.edit();
  await weak.advance(2000);
  weak.sent[0].answer(200, { ETag: 'W/"x"' });
  await settle();
  weak.edit();
  await weak.advance(60000);
  assert.equal(weak.sent.length, 1);

  const bare = rig(request(), { etag: null });
  assert.equal(bare.controller.state.autoPaused, true);
  bare.edit();
  await bare.advance(60000);
  assert.equal(bare.sent.length, 0);
  assert.deepEqual(bare.notices, [['error', STRINGS_EN.autoSaveNoVersion]]);

  const elsewhere = rig(request({ sameAsFile: false }));
  elsewhere.edit();
  await elsewhere.advance(60000);
  assert.equal(elsewhere.sent.length, 0);
  assert.equal(typeof STRINGS_TR.autoSaveNoVersion, 'string');
  assert.notEqual(STRINGS_TR.autoSaveNoVersion, STRINGS_EN.autoSaveNoVersion);
});

test('auto-save review 1: a page hidden with nothing to guard the save with sends nothing either', async () => {
  const r = rig(request(), { etag: null });
  r.edit();
  r.controller.pageHidden();
  await settle();
  assert.equal(r.sent.length, 0);
});

test('auto-save review 1: { unconditional: true } is the page saying that overwriting is acceptable: then it goes without If-Match', async () => {
  assert.deepEqual(resolveAutoSave({ unconditional: true }), { delayMs: 2000, maxWaitMs: 30000, unconditional: true });
  assert.deepEqual(resolveAutoSave({ unconditional: false }), { delayMs: 2000, maxWaitMs: 30000 });
  assert.throws(() => resolveAutoSave({ unconditional: 'yes' }), TypeError);
  const r = rig(request(), { etag: null, autoSave: { delayMs: 2000, maxWaitMs: 30000, unconditional: true } });
  assert.equal('autoPaused' in r.controller.state, false);
  r.edit();
  await r.advance(2000);
  assert.equal(r.sent.length, 1);
  assert.equal('If-Match' in headersOf(r.sent[0]), false);
  assert.deepEqual(r.notices, []);
  // Turned off again while it runs: back to "only with a condition".
  r.sent[0].answer(200);
  await settle();
  r.controller.setAutoSave(true);
  r.edit();
  await r.advance(60000);
  assert.equal(r.sent.length, 1);
});

test('auto-save review 1: a function of the application and a request with the page’s own If-Match are the page’s to guard', async () => {
  let calls = 0;
  const fn = rig({ kind: 'function', save: async () => void calls++ }, { etag: null });
  assert.equal('autoPaused' in fn.controller.state, false);
  fn.edit();
  await fn.advance(2000);
  assert.equal(calls, 1);

  const own = rig(request({ sameAsFile: false, headers: [['if-match', '"mine"']] }), { etag: null });
  own.edit();
  await own.advance(2000);
  assert.equal(own.sent.length, 1);
  assert.equal(headersOf(own.sent[0])['if-match'], '"mine"');
});

// ------------------------------------------------------------ 6: "saved" is not said every few seconds

test('auto-save review 6: a save by itself says "saved" to the state only; after a failure was told, the first success is told once', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].ok();
  await settle();
  assert.deepEqual(r.notices, [], 'no notice for a save nobody asked for');
  assert.equal(typeof r.states[r.states.length - 1].lastSavedAt, 'number', 'the state says when');

  r.edit();
  await r.advance(2000);
  r.sent[1].breakOff();
  await settle();
  await r.advance(5000);
  r.sent[2].ok();
  await settle();
  assert.deepEqual(r.notices, [
    ['error', STRINGS_EN.saveNoConnection],
    ['info', STRINGS_EN.saved],
  ]);
  r.edit();
  await r.advance(2000);
  r.sent[3].ok();
  await settle();
  assert.equal(r.notices.length, 2, 'and then quiet again');

  // The save of the user is told as it always was.
  r.edit();
  const user = r.controller.saveNow();
  await settle();
  r.sent[4].ok();
  await user;
  assert.deepEqual(r.notices[2], ['info', STRINGS_EN.saved]);
});

// ------------------------------------------------------------ 4: the clock

test('auto-save review 4: a clock that is set back an hour does not silence auto-save for an hour', async () => {
  const r = rig();
  r.edit();
  r.pending = true;
  await r.advance(2000);
  r.pending = false;
  r.clock -= 3600000;
  const from = r.clock;
  r.edit();
  await r.advance(4000);
  assert.equal(r.sent.length, 1, 'saved within the delay it was held back for, counted from now');
  assert.equal(r.sent[0].at - from <= 2000, true, `after ${r.sent[0].at - from} ms`);

  // The same for the wait after a failure.
  r.sent[0].breakOff();
  await settle();
  r.clock -= 3600000;
  const again = r.clock;
  r.edit();
  await r.advance(6000);
  assert.equal(r.sent.length, 2);
  assert.equal(r.sent[1].at - again <= 5000, true, `after ${r.sent[1].at - again} ms`);
});

// ------------------------------------------------------------ 7: what the first tests did not hold

test('auto-save review 7: while something is being typed it is looked at once a delay, not without end', async () => {
  const r = rig();
  r.edit();
  r.pending = true;
  await r.advance(2000);
  const before = r.asked;
  const timers = r.timers.length;
  await r.advance(60000);
  assert.equal(r.asked - before <= 31 && r.asked - before >= 29, true, `asked ${r.asked - before} times in a minute`);
  assert.equal(r.timers.length - timers <= 31, true, `${r.timers.length - timers} timers in a minute`);
  assert.equal(r.sent.length, 0);
});

test('auto-save review 7: stopped, then a save of the user that does not get through: auto-save tries again', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].answer(403);
  await settle();
  assert.equal(r.controller.autoSaveStopped, true);
  assert.equal(r.controller.state.autoPaused, true);
  const user = r.controller.saveNow();
  await settle();
  r.sent[1].breakOff();
  assert.equal(await user, false);
  assert.equal(r.controller.autoSaveStopped, false);
  assert.deepEqual(codes(r), ['denied', 'network']);
  const failedAt = r.clock;
  await r.advance(4999);
  assert.equal(r.sent.length, 2);
  await r.advance(1);
  assert.equal(r.sent.length, 3, 'five seconds after the failure');
  assert.equal(r.sent[2].at - failedAt, 5000);
});

test('auto-save review 7: asked for from inside the telling of a change, the file is not taken half way through that change', async () => {
  let changing = false;
  const taken: boolean[] = [];
  let controller: SaveController | undefined;
  const r = rig(request(), {
    onState: () => {
      // A page that hides (or a host that hurries) while the change is still being made.
      if (controller) controller.pageHidden();
    },
    bytesAsIs: async () => {
      taken.push(changing);
      return new Uint8Array([9]);
    },
  });
  controller = r.controller;
  changing = true;
  r.controller.changed();
  changing = false;
  await settle();
  await r.advance(2000);
  assert.deepEqual(taken, [false], 'the file was written once, after the change was over');
});

// ------------------------------------------------------------ 2: what holds the bytes

test('auto-save review 2: a function that has saved leaves no timer behind that holds the save', async () => {
  const r = rig({ kind: 'function', save: async () => undefined });
  for (let i = 0; i < 5; i++) {
    r.edit();
    await r.advance(2000);
  }
  assert.equal(r.controller.dirty, false);
  assert.deepEqual(
    r.timers.filter((timer) => !timer.off || timer.callback !== undefined),
    [],
    'every time limit was called off, and its callback let go of',
  );
});

// ------------------------------------------------------------ 3: what mount() ties

function page() {
  const listeners: { [type: string]: ((event: any) => void)[] } = {};
  const target = {
    visibilityState: 'visible',
    addEventListener: (type: string, listener: (event: any) => void) => void (listeners[type] || (listeners[type] = [])).push(listener),
    removeEventListener: (type: string, listener: (event: any) => void) => {
      const list = listeners[type] || [];
      const at = list.indexOf(listener);
      if (at >= 0) list.splice(at, 1);
    },
  };
  const fire = (type: string) => {
    const event = { type, prevented: false, returnValue: undefined as unknown, preventDefault: () => void (event.prevented = true) };
    for (const listener of (listeners[type] || []).slice()) listener(event);
    return event;
  };
  const count = () => Object.keys(listeners).reduce((sum, type) => sum + listeners[type].length, 0);
  return { target, fire, count };
}

test('auto-save review 3: text that is typed and not applied is unsaved work: hasUnsavedChanges and the question before leaving count it, the state of the file does not', async () => {
  const r = rig();
  const win = page();
  const doc = page();
  let typed = false;
  const view = { setReadOnly: (_readOnly: boolean) => undefined };
  const bound = bindSaveToPage(r.controller, { view, unapplied: () => typed, window: win.target as any, document: doc.target as any, warnOnLeave: true });
  assert.equal(bound.unsaved(), false);
  assert.equal(win.fire('beforeunload').prevented, false);

  // Saved by itself, and then half a cell is typed.
  r.edit();
  await r.advance(2000);
  r.sent[0].ok();
  await settle();
  typed = true;
  assert.equal(r.controller.dirty, false, 'the file is saved');
  assert.equal(bound.unsaved(), true, 'what is typed is not');
  assert.equal(win.fire('beforeunload').prevented, true);
  typed = false;
  assert.equal(win.fire('beforeunload').prevented, false);
  // A host that fails at saying is not in the way of leaving.
  const failing = bindSaveToPage(r.controller, {
    view,
    unapplied: () => {
      throw new Error('no view');
    },
  });
  assert.equal(failing.unsaved(), false);
  failing.release();

  bound.release();
  assert.deepEqual([win.count(), doc.count()], [0, 0], 'nothing is left listening');
  bound.release();
});

test('auto-save review 3: warnOnLeave false asks nothing; a page that is hidden saves at once; editing turned on again starts the delay anew', async () => {
  const r = rig();
  const win = page();
  const doc = page();
  let readOnly = false;
  const seen: boolean[] = [];
  const view = { setReadOnly: (value: boolean) => void seen.push((readOnly = value)) };
  const bound = bindSaveToPage(r.controller, { view, unapplied: () => true, window: win.target as any, document: doc.target as any, warnOnLeave: false });
  assert.equal(win.count(), 0, 'no question before leaving');
  r.edit();
  doc.target.visibilityState = 'hidden';
  doc.fire('visibilitychange');
  await settle();
  assert.equal(r.sent.length, 1, 'hidden: saved at once');
  r.sent[0].ok();
  await settle();

  // The view is told as before, and the controller hears of it.
  let woken = 0;
  let dismissed = 0;
  const wake = r.controller.wake.bind(r.controller);
  const dismiss = r.controller.dismissConflict.bind(r.controller);
  r.controller.wake = () => (woken++, wake());
  r.controller.dismissConflict = () => (dismissed++, dismiss());
  view.setReadOnly(true);
  view.setReadOnly(false);
  assert.deepEqual([seen, readOnly, woken, dismissed], [[true, false], false, 1, 1]);
  bound.release();
  assert.equal(doc.count(), 0);
});

// ------------------------------------------------------------ 5: a drag that ends

test('auto-save review 5: a drag ends when the pointer is let go, lost or cancelled, when the window loses the focus and when the page is hidden', () => {
  for (const [where, type, pointerEvents] of [
    ['window', 'pointerup', true],
    ['window', 'pointercancel', true],
    ['window', 'mouseup', false],
    ['window', 'blur', true],
    ['window', 'blur', false],
    ['element', 'lostpointercapture', true],
    ['document', 'visibilitychange', true],
  ] as [string, string, boolean][]) {
    const targets: { [name: string]: ReturnType<typeof page> } = { window: page(), document: page(), element: page() };
    let ended = 0;
    const stop = watchDragEnd({ window: targets.window.target as any, document: targets.document.target as any, element: targets.element.target as any, pointerEvents }, () => void ended++);
    if (type === 'visibilitychange') {
      targets.document.fire(type);
      assert.equal(ended, 0, 'a page that is shown again ends nothing');
      targets.document.target.visibilityState = 'hidden';
    }
    targets[where].fire(type);
    assert.equal(ended, 1, `${type} on the ${where}`);
    stop();
    assert.deepEqual([targets.window.count(), targets.document.count(), targets.element.count()], [0, 0, 0], `${type}: nothing is left listening`);
    targets[where].fire(type);
    assert.equal(ended, 1);
    stop();
  }
});

// ------------------------------------------------------------ what a second look at the tests showed

test('auto-save review 1: when the server stops naming a version a second time, that is said a second time', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].answer(200);
  await settle();
  r.edit();
  await r.advance(10000);
  assert.equal(r.notices.length, 1);
  // The user saves and the answer brings an ETag; the save by itself after it brings none again.
  const user = r.controller.saveNow();
  await settle();
  r.sent[1].ok();
  await user;
  r.edit();
  await r.advance(2000);
  r.sent[2].answer(200);
  await settle();
  r.edit();
  await r.advance(10000);
  assert.equal(r.sent.length, 3);
  assert.deepEqual(
    r.notices.map((notice) => notice[1]),
    [STRINGS_EN.autoSaveNoVersion, STRINGS_EN.saved, STRINGS_EN.autoSaveNoVersion],
  );
});

test('auto-save review 4: a clock set back while a save is under way does not put what was changed meanwhile off for as long', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.edit();
  r.clock -= 3600000;
  const from = r.clock;
  r.sent[0].ok();
  await settle();
  assert.equal(r.controller.dirty, true);
  await r.advance(2000);
  assert.equal(r.sent.length, 2, 'sent within the delay, counted from now');
  assert.equal(r.sent[1].at - from <= 2000, true, `after ${r.sent[1] ? r.sent[1].at - from : 0} ms`);
});

test('auto-save review: turned off and on again while a save is under way, a page that was hidden before does not make the next save hurry', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.edit();
  r.controller.pageHidden();
  r.controller.setAutoSave(false);
  r.controller.setAutoSave(true);
  r.sent[0].ok();
  await settle();
  assert.equal(r.sent.length, 1, 'what the hidden page asked for went with auto-save being turned off');
  await r.advance(2000);
  assert.equal(r.sent.length, 2);
});
