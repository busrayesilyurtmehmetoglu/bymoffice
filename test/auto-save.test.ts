import assert from 'node:assert/strict';
import { test } from 'node:test';
import {
  AUTO_SAVE_DELAY_MS,
  AUTO_SAVE_MAX_RETRY_MS,
  AUTO_SAVE_MAX_WAIT_MS,
  AUTO_SAVE_RETRY_MS,
  SAVE_TIMEOUT_MS,
  SaveController,
  SaveError,
  resolveAutoSave,
  type ConflictQuestion,
  type SaveHost,
  type SaveInfo,
  type SaveState,
  type SaveTarget,
} from '../src/embed/save-to';
import { STRINGS_EN } from '../src/ui/strings';

// Saving by itself: the controller with a clock and a fetch of the test's own.
// What is counted everywhere is what reaches the server and when.

interface Sent {
  at: number;
  init: RequestInit;
  answer(status: number, headers?: { [name: string]: string }): void;
  breakOff(): void;
}

interface Rig {
  controller: SaveController;
  sent: Sent[];
  states: SaveState[];
  notices: [kind: string, message: string][];
  errors: unknown[];
  timers: { callback: () => void; at: number; off: boolean }[];
  questions: ConflictQuestion[];
  content: number[];
  commits: number;
  /** How often the file was written with and without what is being typed applied. */
  written: { applied: number; asIs: number };
  pending: boolean;
  readOnly: boolean;
  writable: boolean;
  /** Counts the versions the server named. */
  served: number;
  clock: number;
  /** Lets `ms` pass, running every timer that comes due on the way. */
  advance(ms: number): Promise<void>;
  /** Changes the file. */
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
    commits: 0,
    written: { applied: 0, asIs: 0 },
    pending: false,
    readOnly: false,
    writable: true,
    served: 0,
    clock: 100000,
    async advance(ms) {
      const end = state.clock + ms;
      for (let guard = 0; guard < 10000; guard++) {
        let next: Rig['timers'][number] | undefined;
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
      state.controller.changed();
    },
  };
  const fakeFetch = (_url: unknown, init?: RequestInit) =>
    new Promise<Response>((resolve, reject) => {
      const signal = init && init.signal;
      if (signal) signal.addEventListener('abort', () => reject(new Error('aborted')));
      state.sent.push({
        at: state.clock,
        init: init || {},
        answer(status, headers = {}) {
          const lower: { [name: string]: string } = {};
          for (const name of Object.keys(headers)) lower[name.toLowerCase()] = headers[name];
          // A server that takes a file names its new version, unless the test says what it names: without
          // one auto-save has nothing to guard its next save with (that has tests of its own).
          if (status >= 200 && status < 300 && !('etag' in lower)) lower.etag = `"s${++state.served}"`;
          resolve({ status, headers: { get: (name: string) => (name.toLowerCase() in lower ? lower[name.toLowerCase()] : null) } } as unknown as Response);
        },
        breakOff: () => reject(new TypeError('Failed to fetch')),
      });
    });
  const host: SaveHost = {
    target,
    writable: () => state.writable,
    commit: () => {
      state.commits++;
      return true;
    },
    bytes: async () => {
      state.written.applied++;
      return new Uint8Array(state.content);
    },
    bytesAsIs: async () => {
      state.written.asIs++;
      return new Uint8Array(state.content);
    },
    pending: () => state.pending,
    readOnly: () => state.readOnly,
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
      const entry = { callback, at: state.clock + ms, off: false };
      state.timers.push(entry);
      return () => void (entry.off = true);
    },
    autoSave: { delayMs: 2000, maxWaitMs: 30000 },
    ...more,
  };
  state.controller = new SaveController(host);
  return state;
}

/** Lets time pass, a quarter of a second at a time, until the next request is sent (or `limit` is over). */
const untilSent = async (r: Rig, count: number, limit: number) => {
  for (let passed = 0; passed < limit && r.sent.length < count; passed += 250) await r.advance(250);
};

const headersOf = (sent: Sent) => sent.init.headers as { [name: string]: string };
const bodyOf = (sent: Sent) => Array.from(sent.init.body as unknown as Uint8Array);
const codes = (r: Rig) => r.errors.map((error) => (error instanceof SaveError ? error.code : String(error)));

// ------------------------------------------------------------ the option

test('autoSave: true is two seconds after the last change and thirty at most; off unless asked for', () => {
  assert.equal(resolveAutoSave(undefined), undefined);
  assert.equal(resolveAutoSave(false), undefined);
  assert.equal(resolveAutoSave(null), undefined);
  assert.deepEqual(resolveAutoSave(true), { delayMs: 2000, maxWaitMs: 30000 });
  assert.deepEqual([AUTO_SAVE_DELAY_MS, AUTO_SAVE_MAX_WAIT_MS], [2000, 30000]);
  assert.deepEqual(resolveAutoSave({}), { delayMs: 2000, maxWaitMs: 30000 });
  assert.deepEqual(resolveAutoSave({ delayMs: 250 }), { delayMs: 250, maxWaitMs: 30000 });
  assert.deepEqual(resolveAutoSave({ delayMs: 600000 }), { delayMs: 600000, maxWaitMs: 600000 }, 'the longest wait is never shorter than the delay');
  assert.deepEqual(resolveAutoSave({ delayMs: 500, maxWaitMs: 500 }), { delayMs: 500, maxWaitMs: 500 });
});

test('autoSave: a delay outside 250 ms to ten minutes, a longest wait below the delay and anything that is no number are TypeErrors', () => {
  for (const bad of [{ delayMs: 249 }, { delayMs: 600001 }, { delayMs: NaN }, { delayMs: '2000' }, { delayMs: 2000, maxWaitMs: 1999 }, { maxWaitMs: 1000 }, { maxWaitMs: Infinity }, { maxWaitMs: '1' }, 'yes', 1, []]) {
    assert.throws(() => resolveAutoSave(bad), TypeError, JSON.stringify(bad));
  }
});

// ------------------------------------------------------------ when it saves

test('auto-save: the file is sent the delay after the last change, and every change starts the delay again', async () => {
  const r = rig();
  assert.equal(r.controller.state.auto, true);
  r.edit();
  await r.advance(1500);
  r.edit();
  await r.advance(1999);
  assert.equal(r.sent.length, 0, 'nothing is sent before the delay after the last change is over');
  await r.advance(1);
  assert.equal(r.sent.length, 1);
  assert.equal(r.sent[0].at, 100000 + 3500);
  assert.equal(headersOf(r.sent[0])['If-Match'], '"v1"', 'guarded as the save of the button is');
  assert.deepEqual(bodyOf(r.sent[0]), [3]);
  assert.equal(r.controller.saving, true);
  r.sent[0].answer(204, { ETag: '"v2"' });
  await settle();
  assert.deepEqual([r.controller.dirty, r.controller.saving], [false, false]);
  assert.equal(r.controller.currentETag, '"v2"');
  await r.advance(60000);
  assert.equal(r.sent.length, 1, 'a file that is saved is not sent again');
});

test('auto-save: off by default, and nothing is sent then', async () => {
  const r = rig(request(), { autoSave: undefined });
  assert.equal('auto' in r.controller.state, false);
  r.edit();
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  assert.equal(r.timers.length, 0, 'no timer is set for a controller that does not save by itself');
});

test('auto-save: who never stops typing is saved the longest wait after the first change', async () => {
  const r = rig(request(), { autoSave: { delayMs: 2000, maxWaitMs: 5000 } });
  for (let i = 0; i < 5; i++) {
    r.edit();
    await r.advance(1000);
  }
  assert.equal(r.sent.length, 1, 'five seconds after the first change');
  assert.equal(r.sent[0].at, 105000);
  // The changes go on while the file is under way: they are counted from when that save took the file.
  r.edit();
  await r.advance(500);
  r.sent[0].answer(200);
  await settle();
  assert.equal(r.controller.dirty, true);
  for (let i = 0; i < 6 && r.sent.length < 2; i++) {
    r.edit();
    await r.advance(1000);
  }
  assert.equal(r.sent.length, 2);
  assert.equal(r.sent[1].at, 110000, 'five seconds after the save before it took the file');
});

// ------------------------------------------------------------ what is being typed

test('auto-save: nothing is sent, applied or closed while something is being typed; it is looked at again a delay later', async () => {
  const r = rig();
  r.edit();
  r.pending = true;
  await r.advance(60000);
  assert.equal(r.sent.length, 0, 'not even when the longest wait is over');
  assert.equal(r.commits, 0, 'what is being typed is never applied by a save nobody asked for');
  r.pending = false;
  await r.advance(2000);
  assert.equal(r.sent.length, 1);
  assert.equal(r.commits, 0);
  assert.deepEqual(r.written, { applied: 0, asIs: 1 }, 'the file is written as the model is, without applying anything');
});

test('auto-save: a host that fails at saying whether something is typed is taken to be typing', async () => {
  const r = rig(request(), {
    pending: () => {
      throw new Error('no view');
    },
  });
  r.edit();
  await r.advance(10000);
  assert.equal(r.sent.length, 0);
});

test('auto-save: the save of the user applies what is being typed, as it always did', async () => {
  const r = rig();
  r.edit();
  void r.controller.saveNow();
  await settle();
  assert.equal(r.commits, 1);
  assert.deepEqual(r.written, { applied: 1, asIs: 0 });
});

// ------------------------------------------------------------ a conflict

test('auto-save: a conflict is told and not asked about, and nothing is sent again until the user saves; then the user is asked', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].answer(412);
  await settle();
  assert.equal(r.questions.length, 0, 'no question opens under the hands of someone who is typing');
  assert.equal(r.controller.asking, false);
  assert.deepEqual(codes(r), ['conflict']);
  assert.deepEqual(r.notices, [['error', STRINGS_EN.saveConflict]]);
  assert.equal(r.controller.state.error, STRINGS_EN.saveConflict);
  assert.equal(r.controller.autoSaveStopped, true);

  // More changes and more time: nothing is sent by itself.
  for (let i = 0; i < 5; i++) {
    r.edit();
    await r.advance(40000);
  }
  assert.equal(r.sent.length, 1);

  // The Save button: the same guarded request, and now the question.
  const saved = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 2);
  assert.equal(headersOf(r.sent[1])['If-Match'], '"v1"', 'the save of the user is guarded as well: nothing was overwritten unasked');
  r.sent[1].answer(412);
  await settle();
  assert.equal(r.questions.length, 1);
  assert.equal(r.controller.asking, true);
  await r.advance(60000);
  assert.equal(r.sent.length, 2, 'nothing is sent while the question is open');
  r.questions[0].answer(true);
  await settle();
  assert.equal(r.sent.length, 3);
  assert.equal('If-Match' in headersOf(r.sent[2]), false);
  r.sent[2].answer(200, { ETag: '"v9"' });
  assert.equal(await saved, true);
  assert.equal(r.controller.autoSaveStopped, false);

  // Saving by itself goes on from there.
  r.edit();
  await r.advance(2000);
  assert.equal(r.sent.length, 4);
  assert.equal(headersOf(r.sent[3])['If-Match'], '"v9"');
});

test('auto-save: after a conflict the user said "cancel" to, nothing is sent by itself either', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].answer(412);
  await settle();
  void r.controller.saveNow();
  await settle();
  r.sent[1].answer(412);
  await settle();
  r.questions[0].answer(false);
  await settle();
  r.edit();
  await r.advance(120000);
  assert.equal(r.sent.length, 2);
});

test('auto-save: the user who presses Save while a save by itself is under way is asked about its conflict', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  const joined = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 1, 'one request: the save under way holds everything');
  r.sent[0].answer(412);
  await settle();
  assert.equal(r.questions.length, 1, 'the save is the user’s now');
  r.questions[0].answer(false);
  assert.equal(await joined, false);
});

// ------------------------------------------------------------ failures

test('auto-save: a request that did not get through is tried again after 5, 10, 20 seconds, and told once', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  const waits: number[] = [];
  for (let i = 0; i < 4; i++) {
    const failedAt = r.clock;
    r.sent[i].breakOff();
    await settle();
    assert.equal(r.controller.state.error, STRINGS_EN.saveNoConnection, 'the state says so every time');
    await untilSent(r, i + 2, AUTO_SAVE_MAX_RETRY_MS);
    assert.equal(r.sent.length, i + 2, `try ${i + 2}`);
    waits.push(r.sent[i + 1].at - failedAt);
  }
  assert.deepEqual(waits, [5000, 10000, 20000, 40000]);
  assert.equal(AUTO_SAVE_RETRY_MS, 5000);
  assert.deepEqual(r.notices, [['error', STRINGS_EN.saveNoConnection]], 'the same failure is not said again and again');
  assert.deepEqual(codes(r), ['network']);

  // A save that succeeds: the next failure is told again and waits five seconds again.
  r.sent[4].answer(200);
  await settle();
  r.edit();
  await r.advance(2000);
  r.sent[5].breakOff();
  await settle();
  assert.deepEqual(codes(r), ['network', 'network']);
  const failedAt = r.clock;
  await r.advance(5000);
  assert.equal(r.sent.length, 7);
  assert.equal(r.sent[6].at - failedAt, 5000);
});

test('auto-save: the wait between tries grows to five minutes and no further, and a change does not shorten it', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  let last = 0;
  for (let i = 0; i < 9; i++) {
    const failedAt = r.clock;
    r.sent[i].breakOff();
    await settle();
    // Typing on does not bring the next try forward.
    r.edit();
    await r.advance(2500);
    assert.equal(r.sent.length, i + 1, 'no try before its time');
    await untilSent(r, i + 2, AUTO_SAVE_MAX_RETRY_MS);
    assert.equal(r.sent.length, i + 2);
    last = r.sent[i + 1].at - failedAt;
  }
  assert.equal(last, AUTO_SAVE_MAX_RETRY_MS);
  assert.equal(AUTO_SAVE_MAX_RETRY_MS, 300000);
});

test('auto-save: a 5xx answer and a request without an answer in time are tried again; another failure is told as it changes', async () => {
  const r = rig(request(), { timeoutMs: 1000 });
  r.edit();
  await r.advance(2000);
  r.sent[0].answer(503);
  await settle();
  await r.advance(5000);
  assert.equal(r.sent.length, 2);
  // No answer: the time limit of the request ends it.
  await r.advance(1000);
  assert.deepEqual(codes(r), ['http', 'timeout']);
  await r.advance(10000);
  assert.equal(r.sent.length, 3);
});

for (const [name, status, code] of [
  ['denied (403)', 403, 'denied'],
  ['denied (401)', 401, 'denied'],
  ['locked (423)', 423, 'locked'],
  ['any other 4xx (404)', 404, 'http'],
  ['too large (413)', 413, 'http'],
  ['a redirect (302)', 302, 'redirected'],
] as [string, number, string][]) {
  test(`auto-save: ${name} stops it until the user saves`, async () => {
    const r = rig();
    r.edit();
    await r.advance(2000);
    r.sent[0].answer(status);
    await settle();
    assert.deepEqual(codes(r), [code]);
    assert.equal(r.controller.autoSaveStopped, true);
    r.edit();
    await r.advance(900000);
    assert.equal(r.sent.length, 1, 'no try by itself');
    // The user saves, and the server takes it: saving by itself is back.
    void r.controller.saveNow();
    await settle();
    r.sent[1].answer(200);
    await settle();
    r.edit();
    await r.advance(2000);
    assert.equal(r.sent.length, 3);
  });
}

test('auto-save: a function of the application that rejects stops it; one that is given the file knows that nobody pressed Save', async () => {
  const given: SaveInfo[] = [];
  let fail = true;
  const r = rig({
    kind: 'function',
    save: async (_bytes, info) => {
      given.push(info);
      if (fail) throw new Error('disk full');
    },
  });
  r.edit();
  await r.advance(2000);
  assert.equal(given.length, 1);
  assert.equal(given[0].auto, true);
  assert.deepEqual(codes(r), ['refused']);
  r.edit();
  await r.advance(600000);
  assert.equal(given.length, 1);
  fail = false;
  assert.equal(await r.controller.saveNow(), true);
  assert.equal('auto' in given[1], false, 'the save of the user does not say "auto"');
});

// ------------------------------------------------------------ what is never tried

test('auto-save: a converted file that may not replace its original is never tried, and nothing is said about it', async () => {
  const r = rig(request(), {}, { format: 'docx', convertedFrom: 'doc' });
  assert.equal(r.controller.blockedConversion, true);
  r.edit();
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  assert.deepEqual(r.notices, []);
  assert.deepEqual(r.errors, []);
  // With the server's word that it expects the new format it is saved as any other file.
  const allowed = rig(request({ allowConverted: true }), {}, { format: 'docx', convertedFrom: 'doc' });
  allowed.edit();
  await allowed.advance(2000);
  assert.equal(allowed.sent.length, 1);
});

test('auto-save: nothing is sent from a view that is only read; when editing is turned on again the delay starts anew', async () => {
  const r = rig();
  r.edit();
  r.readOnly = true;
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  r.readOnly = false;
  r.controller.wake();
  await r.advance(1999);
  assert.equal(r.sent.length, 0);
  await r.advance(1);
  assert.equal(r.sent.length, 1);
});

test('auto-save: nothing is sent for a file that cannot be written', async () => {
  const r = rig();
  r.writable = false;
  r.edit();
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
});

test('auto-save: destroy() calls its timer off, and nothing is sent afterwards', async () => {
  const r = rig();
  r.edit();
  r.controller.destroy();
  assert.equal(
    r.timers.every((timer) => timer.off),
    true,
    'no timer is left running',
  );
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  r.controller.pageHidden();
  r.controller.setAutoSave(true);
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  assert.equal(
    r.timers.every((timer) => timer.off),
    true,
    'and none is set again',
  );
});

// ------------------------------------------------------------ the page is hidden

test('auto-save: a page that is hidden saves what waits at once, but never what is being typed', async () => {
  const r = rig();
  r.edit();
  await r.advance(300);
  r.controller.pageHidden();
  await settle();
  assert.equal(r.sent.length, 1, 'the last chance before the tab is gone');
  assert.equal(r.sent[0].at, 100300);
  r.sent[0].answer(200);
  await settle();
  r.controller.pageHidden();
  await settle();
  assert.equal(r.sent.length, 1, 'nothing to save, nothing sent');

  r.edit();
  r.pending = true;
  r.controller.pageHidden();
  await settle();
  assert.equal(r.sent.length, 1);
  assert.equal(r.commits, 0);
});

test('auto-save: a page hidden while a save is under way saves what was changed since as soon as that save ends', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.edit();
  r.controller.pageHidden();
  await settle();
  assert.equal(r.sent.length, 1, 'one request at a time');
  r.sent[0].answer(200);
  await settle();
  assert.equal(r.sent.length, 2, 'without waiting for the delay');
  assert.deepEqual(bodyOf(r.sent[1]), [3]);
});

test('auto-save: a hidden page does not save where saving by itself is off or was stopped', async () => {
  const off = rig(request(), { autoSave: undefined });
  off.edit();
  off.controller.pageHidden();
  await settle();
  assert.equal(off.sent.length, 0);

  const stopped = rig();
  stopped.edit();
  await stopped.advance(2000);
  stopped.sent[0].answer(412);
  await settle();
  stopped.controller.pageHidden();
  await settle();
  assert.equal(stopped.sent.length, 1);
});

// ------------------------------------------------------------ on and off

test('setAutoSave: turned off, what waited is not sent; turned on, a file with changes is sent a delay later', async () => {
  const r = rig();
  r.edit();
  r.controller.setAutoSave(false);
  assert.equal('auto' in r.controller.state, false);
  assert.equal(r.controller.autoSave, undefined);
  await r.advance(120000);
  assert.equal(r.sent.length, 0);
  r.controller.setAutoSave({ delayMs: 500 });
  assert.deepEqual(r.controller.autoSave, { delayMs: 500, maxWaitMs: 30000 });
  assert.equal(r.states[r.states.length - 1].auto, true, 'the state says that it is on');
  await r.advance(499);
  assert.equal(r.sent.length, 0);
  await r.advance(1);
  assert.equal(r.sent.length, 1);
});

test('setAutoSave: a value that cannot be is a TypeError and leaves everything as it was', async () => {
  const r = rig();
  assert.throws(() => r.controller.setAutoSave({ delayMs: 1 }), TypeError);
  assert.deepEqual(r.controller.autoSave, { delayMs: 2000, maxWaitMs: 30000 });
  r.edit();
  await r.advance(2000);
  assert.equal(r.sent.length, 1);
});

// ------------------------------------------------------------ one request at a time

test('auto-save: its time comes while the save of the user is under way: nothing is sent beside it, and what changed since goes later', async () => {
  const r = rig();
  r.edit();
  const user = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 1);
  r.edit();
  await r.advance(60000);
  assert.equal(r.sent.length, 1, 'never two requests');
  r.sent[0].answer(200, { ETag: '"v2"' });
  assert.equal(await user, true);
  assert.equal(r.controller.dirty, true);
  await r.advance(2000);
  assert.equal(r.sent.length, 2);
  assert.deepEqual(bodyOf(r.sent[1]), [3]);
  assert.equal(headersOf(r.sent[1])['If-Match'], '"v2"');
});

test('auto-save: the user saves while a save by itself is under way: what was typed since is applied and sent after it', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.edit();
  const user = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 1);
  assert.equal(r.commits, 1, 'the save of the user applies what is being typed');
  r.sent[0].answer(200, { ETag: '"v2"' });
  await settle();
  assert.equal(r.sent.length, 2);
  assert.deepEqual(r.written, { applied: 1, asIs: 1 });
  r.sent[1].answer(200);
  assert.equal(await user, true);
  assert.equal(r.controller.dirty, false);
  await r.advance(60000);
  assert.equal(r.sent.length, 2);
});

test('auto-save: saving is never left hanging: every try ends in a state that is not "saving"', async () => {
  const r = rig();
  r.edit();
  await r.advance(2000);
  r.sent[0].breakOff();
  await settle();
  assert.equal(r.controller.saving, false);
  await r.advance(5000);
  r.sent[1].answer(500);
  await settle();
  assert.equal(r.controller.saving, false);
  await r.advance(10000);
  r.sent[2].answer(200);
  await settle();
  assert.deepEqual([r.controller.saving, r.controller.dirty, r.controller.state.error], [false, false, undefined]);
});
