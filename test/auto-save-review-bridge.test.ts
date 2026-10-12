import assert from 'node:assert/strict';
import { test } from 'node:test';
import { setFlagsFromString } from 'node:v8';
import { runInNewContext } from 'node:vm';
import { base64Encode, createBridge, installBridge, type Bridge, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';

// The bridge page after the review of auto-save: what holds the bytes of a save lets go of
// them when the save is over, and the page that is hidden is heard by the installed bridge.

setFlagsFromString('--expose-gc');
const collect = runInNewContext('gc') as () => void;

const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};
const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));

interface Timer {
  callback: (() => void) | undefined;
  at: number;
  off: boolean;
}

interface Fake {
  bridge: Bridge;
  sent: any[];
  timers: Timer[];
  /** Every file that was written, held weakly: what is still reachable shows who holds on to it. */
  files: WeakRef<Uint8Array>[];
  onChange: (() => void) | undefined;
  clock: number;
  advance(ms: number): Promise<void>;
}

/** `cancellable`: the timer of the host gives back what calls it off, as the installed bridge does; otherwise nothing, as hosts written before did. */
function fake(cancellable: boolean): Fake {
  const state: Fake = {
    bridge: undefined as unknown as Bridge,
    sent: [],
    timers: [],
    files: [],
    onChange: undefined,
    clock: 50000,
    async advance(ms) {
      const end = state.clock + ms;
      for (let guard = 0; guard < 10000; guard++) {
        let next: Timer | undefined;
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
  };
  const host: BridgeHost = {
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
    version: '9.9.9',
    post(json) {
      // The application takes the message and keeps nothing of it.
      const message = JSON.parse(json);
      delete message.data;
      state.sent.push(message);
      return true;
    },
    async open() {
      const file: any = {
        format: 'xlsx',
        kind: 'spreadsheet',
        save: async () => {
          const bytes = new Uint8Array(1 << 20);
          state.files.push(new WeakRef(bytes));
          return bytes;
        },
      };
      file.workbook = async () => ({ sheets: [{ name: 'Veri', state: 'visible', kind: 'worksheet' }] });
      return file as OfficeFile;
    },
    createView(_container, _workbook, options) {
      state.onChange = () => (options as any).onChange({ sheets: [0] });
      return {
        history: { canUndo: true, canRedo: false },
        activeSheet: 0,
        readOnly: false,
        commitPending: () => true,
        hasPendingInput: () => false,
        showNotice: () => undefined,
        setReadOnly: () => undefined,
        setLocale: () => undefined,
        destroy: () => undefined,
      } as any;
    },
    timer: (callback, ms) => {
      const entry: Timer = { callback, at: state.clock + ms, off: false };
      state.timers.push(entry);
      if (!cancellable) return undefined;
      return () => {
        entry.off = true;
        entry.callback = undefined;
      };
    },
    now: () => state.clock,
    frame: () => undefined,
  };
  state.bridge = createBridge(host);
  return state;
}

const reachable = async (f: Fake): Promise<number> => {
  for (let i = 0; i < 3; i++) {
    await new Promise((resolve) => setTimeout(resolve, 20));
    collect();
  }
  return f.files.filter((file) => file.deref() !== undefined).length;
};

for (const cancellable of [true, false]) {
  test(`bridge auto-save review 2: the bytes of a save that was answered are let go of at once (${cancellable ? 'a timer that can be called off' : 'a timer of a host that cannot call it off'})`, async () => {
    const f = fake(cancellable);
    await f.bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
    const count = 30;
    for (let i = 1; i <= count; i++) {
      (f.onChange as () => void)();
      await f.advance(2000);
      const saved = f.sent.filter((message) => message.type === 'saved').pop();
      await f.bridge.receive({ type: 'saveResult', saveId: saved.saveId, ok: true });
      await settle();
    }
    assert.equal(f.files.length, count, 'every save wrote the file');
    // Sixty simulated seconds have passed in all: neither the minute the page waits for a result nor the
    // ten minutes a function may take are over for most of these saves.
    const alive = await reachable(f);
    assert.equal(alive <= 2, true, `${alive} of ${count} files of 1 MiB are still held although their saves are over`);
    if (cancellable) {
      const waiting = f.timers.filter((timer) => !timer.off);
      assert.equal(waiting.length <= 1, true, `${waiting.length} timers are still running for saves that are over`);
    }
  });
}

test('bridge auto-save review 2: without a result the save still fails after a minute, with either kind of timer', async () => {
  for (const cancellable of [true, false]) await withoutResult(fake(cancellable));
});

async function withoutResult(f: Fake): Promise<void> {
  const states = () => f.sent.filter((message) => message.type === 'saveState').map((message) => `${message.dirty} ${message.saving}`);
  await f.bridge.receive({ type: 'load', id: 'm', data: DATA, save: { auto: true } });
  (f.onChange as () => void)();
  await f.advance(2000);
  assert.equal(states().pop(), 'true true');
  await f.advance(59999);
  assert.equal(states().pop(), 'true true');
  await f.advance(1);
  assert.equal(states().pop(), 'true false', 'failed for want of an answer');
}

// ------------------------------------------------------------ the installed bridge and the hidden page

test('bridge auto-save review 7: the installed bridge hears the page being hidden and hands the file over at once; destroy() stops listening', async () => {
  const listeners: { [type: string]: (() => void)[] } = {};
  const documentListeners: { [type: string]: (() => void)[] } = {};
  const listen = (map: { [type: string]: (() => void)[] }) => ({
    addEventListener: (type: string, listener: () => void) => void (map[type] || (map[type] = [])).push(listener),
    removeEventListener: (type: string, listener: () => void) => {
      const list = map[type] || [];
      const at = list.indexOf(listener);
      if (at >= 0) list.splice(at, 1);
    },
  });
  const posted: any[] = [];
  const doc = { visibilityState: 'visible', ...listen(documentListeners) };
  const timeouts: (() => void)[] = [];
  const win: any = {
    ...listen(listeners),
    document: doc,
    location: { href: 'https://app.example.com/bridge.html' },
    // No time passes in this test: whatever is sent is sent without a timer.
    setTimeout: (callback: () => void) => timeouts.push(callback),
    clearTimeout: () => undefined,
    BymOfficeHost: { postMessage: (json: string) => void posted.push(JSON.parse(json)) },
  };
  win.parent = win;
  let change: (() => void) | undefined;
  const bridge = installBridge({
    window: win as Window,
    container: { addEventListener: () => undefined, removeEventListener: () => undefined } as unknown as HTMLElement,
    version: '9.9.9',
    open: async () => {
      const file: any = { format: 'xlsx', kind: 'spreadsheet', save: async () => new Uint8Array([7]) };
      file.workbook = async () => ({ sheets: [{ name: 'Veri', state: 'visible', kind: 'worksheet' }] });
      return file as OfficeFile;
    },
    createView: (_container, _workbook, options) => {
      change = () => (options as any).onChange({ sheets: [0] });
      return { history: {}, activeSheet: 0, readOnly: false, commitPending: () => true, hasPendingInput: () => false, showNotice: () => undefined, setReadOnly: () => undefined, setLocale: () => undefined, destroy: () => undefined } as any;
    },
  });
  await bridge.receive({ type: 'load', id: 'l', data: DATA, save: { auto: true } });
  assert.equal((documentListeners.visibilitychange || []).length, 1, 'the page is watched');
  (change as () => void)();
  const saves = () => posted.filter((message) => message.type === 'saved');
  // Shown again: nothing happens.
  for (const listener of documentListeners.visibilitychange) listener();
  await settle();
  assert.equal(saves().length, 0);
  doc.visibilityState = 'hidden';
  for (const listener of documentListeners.visibilitychange) listener();
  await settle();
  assert.equal(saves().length, 1, 'handed over at once, without a timer');
  assert.equal(saves()[0].by, 'auto');
  bridge.destroy();
  assert.equal((documentListeners.visibilitychange || []).length, 0, 'destroy() stops listening');
});
