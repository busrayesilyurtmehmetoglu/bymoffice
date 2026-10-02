import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Decode, base64Encode, createBridge, isAllowedUrl, normalizeOrigin, resolveLoadUrl, type BridgeHost } from '../src/embed/bridge';
import type { OfficeFile } from '../src/index';
import type { WorkbookViewOptions } from '../src/ui/sheet-view';

/** Deterministic bytes that do not repeat with a short period. */
function noise(length: number): Uint8Array {
  const out = new Uint8Array(length);
  let seed = 0x2545f491;
  for (let i = 0; i < length; i++) {
    seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
    out[i] = seed >>> 24;
  }
  return out;
}

const char = String.fromCharCode;

test('base64: encodes like the platform for every length around the step size', () => {
  for (const length of [0, 1, 2, 3, 4, 5, 6, 255, 12287, 12288, 12289, 24576, 24577, 40000]) {
    const bytes = noise(length);
    const text = base64Encode(bytes);
    assert.equal(text, Buffer.from(bytes).toString('base64'), `length ${length}`);
    assert.deepEqual(base64Decode(text), bytes, `length ${length}`);
  }
});

test('base64: decodes wrapped, unpadded and URL-safe text', () => {
  assert.deepEqual(base64Decode(`aGVs${char(13, 10)}bG8g d29y${char(9)}bGQ=${char(10)}`), new TextEncoder().encode('hello world'));
  assert.deepEqual(base64Decode('aGVsbG8gd29ybGQ'), new TextEncoder().encode('hello world'));
  assert.deepEqual(base64Decode('-_-_'), new Uint8Array([0xfb, 0xff, 0xbf]));
  assert.deepEqual(base64Decode('+/+/'), new Uint8Array([0xfb, 0xff, 0xbf]));
  assert.deepEqual(base64Decode(''), new Uint8Array(0));
  assert.deepEqual(base64Decode('AA=='), new Uint8Array([0]));
});

test('base64: rejects text that is not base64', () => {
  for (const bad of ['a', 'abcde', 'ab*d', 'abç=', 'ab=c', 'a===', 'abc==', 'abcd=', '====', `ab${char(0)}d`, 'data:;base64,AAAA']) {
    assert.throws(() => base64Decode(bad), /Invalid base64/, JSON.stringify(bad));
  }
});

test('base64: 20 MB goes there and back without a stack overflow, in linear time', () => {
  const bytes = noise(20 * 1024 * 1024 + 1);
  const started = Date.now();
  const text = base64Encode(bytes);
  const back = base64Decode(text);
  const elapsed = Date.now() - started;
  assert.equal(text.length, Math.ceil(bytes.length / 3) * 4);
  assert.equal(back.length, bytes.length);
  assert.equal(Buffer.compare(Buffer.from(back.buffer), Buffer.from(bytes.buffer)), 0);
  assert.ok(elapsed < 5000, `took ${elapsed} ms`);
});

test('resolveLoadUrl: a page served over http(s) loads from its own origin, blob: and data: only', () => {
  const base = 'https://app.example.com/editor/bymoffice.webview.html?x=1';
  const good: [string, string][] = [
    ['files/a.xlsx', 'https://app.example.com/editor/files/a.xlsx'],
    ['/a.xlsx', 'https://app.example.com/a.xlsx'],
    ['../a.xlsx?x=1#y', 'https://app.example.com/a.xlsx?x=1#y'],
    ['https://app.example.com/a.xlsx', 'https://app.example.com/a.xlsx'],
    ['HTTPS://APP.example.com:443/a.xlsx', 'https://app.example.com/a.xlsx'],
    ['//app.example.com/a.xlsx', 'https://app.example.com/a.xlsx'],
    ['blob:https://app.example.com/1-2', 'blob:https://app.example.com/1-2'],
    ['data:application/octet-stream;base64,AAAA', 'data:application/octet-stream;base64,AAAA'],
  ];
  for (const [url, resolved] of good) assert.equal(resolveLoadUrl(url, base), resolved, url);
  const bad = [
    '',
    'https://evil.example/a.xlsx',
    'http://app.example.com/a.xlsx',
    'https://app.example.com:8443/a.xlsx',
    '//evil.example/x',
    `${char(92, 92)}evil${char(92)}x`,
    `/${char(92)}evil.example/x`,
    'https://user:pw@app.example.com/a.xlsx',
    'https://user@app.example.com/a.xlsx',
    'javascript:alert(1)',
    'JaVaScript:alert(1)',
    ' javascript:alert(1)',
    `java${char(9)}script:alert(1)`,
    `${char(1)}javascript:x`,
    'file:///etc/passwd',
    'ftp://app.example.com/a',
    'vbscript:x',
    'a b.xlsx',
    'content://x/y',
  ];
  for (const url of bad) assert.equal(isAllowedUrl(url, base), false, JSON.stringify(url));
});

test('resolveLoadUrl: a page loaded from a file or an asset scheme loads blob: and data: only', () => {
  for (const base of ['file:///android_asset/flutter_assets/assets/bymoffice.webview.html', 'app-asset://local/bymoffice.webview.html', undefined]) {
    for (const url of ['a.xlsx', '/a.xlsx', '../../etc/passwd', 'file:///etc/passwd', 'https://files.example/a.xlsx', '//files.example/a.xlsx']) {
      assert.equal(isAllowedUrl(url, base), false, `${url} from ${base}`);
    }
    assert.equal(isAllowedUrl('data:application/octet-stream;base64,AAAA', base), true);
    assert.equal(isAllowedUrl('blob:null/1-2', base), true);
  }
});

test('normalizeOrigin: the form browsers report, and nothing for what is not an http(s) origin', () => {
  assert.equal(normalizeOrigin('https://App.Example.com:443/'), 'https://app.example.com');
  assert.equal(normalizeOrigin('http://localhost:8080/path?x'), 'http://localhost:8080');
  for (const bad of ['', 'null', '*', 'app.example.com', 'file:///x', 'data:text/plain,x', 'blob:https://app.example.com/1']) {
    assert.equal(normalizeOrigin(bad), '', bad);
  }
});

interface Fake {
  host: BridgeHost;
  /** Messages the bridge sent, parsed. */
  sent: any[];
  states: string[];
  locales: string[];
  calls: string[];
  opened: Uint8Array[];
  views: WorkbookViewOptions[];
  timers: { callback: () => void; ms: number }[];
  frames: (() => void)[];
  fetches: { url: string; signal?: AbortSignal; resolve(bytes: Uint8Array): void; reject(error: Error): void }[];
  /** Whether the application's channel exists. */
  channel: boolean;
  history: { canUndo: boolean; canRedo: boolean };
  editing: boolean;
  saveResult: () => Promise<Uint8Array>;
  openError?: Error;
  viewError?: Error;
  sheetNames: string[];
  /** Runs the retry timers that are waiting (not the fetch time limit). */
  runRetries(): void;
}

function fake(): Fake {
  const state: Fake = {
    sent: [],
    states: [],
    locales: [],
    calls: [],
    opened: [],
    views: [],
    timers: [],
    frames: [],
    fetches: [],
    channel: true,
    history: { canUndo: false, canRedo: false },
    editing: false,
    saveResult: async () => new Uint8Array([1, 2, 3, 250]),
    sheetNames: ['Veri', 'Gizli'],
    host: undefined as unknown as BridgeHost,
    runRetries() {
      const due = state.timers.filter((entry) => entry.ms === 50);
      state.timers = state.timers.filter((entry) => entry.ms !== 50);
      for (const entry of due) entry.callback();
    },
  };
  state.host = {
    container: {} as HTMLElement,
    version: '9.9.9',
    baseUrl: 'https://app.example.com/bymoffice.webview.html',
    fetchTimeoutMs: 7000,
    post(json) {
      if (!state.channel) return false;
      assert.equal(typeof json, 'string');
      state.sent.push(JSON.parse(json));
      return true;
    },
    async open(bytes) {
      state.opened.push(bytes);
      if (state.openError) throw state.openError;
      const names = state.sheetNames;
      return {
        format: 'xlsx',
        workbook: async () => ({ sheets: names.map((name, index) => ({ name, state: index === 1 ? 'hidden' : 'visible', kind: 'worksheet' })) }),
        save: () => state.saveResult(),
      } as unknown as OfficeFile;
    },
    createView(_container, _workbook, options) {
      state.calls.push('view');
      if (state.viewError) throw state.viewError;
      state.views.push(options);
      return {
        history: state.history,
        activeSheet: options.sheet === undefined ? 0 : options.sheet,
        get isEditing() {
          return state.editing;
        },
        commitEditing: () => state.calls.push('commit'),
        setReadOnly: (value: boolean) => state.calls.push(`setReadOnly ${value}`),
        setLocale: (value: string) => state.calls.push(`setLocale ${value}`),
        undo: () => state.calls.push('undo'),
        redo: () => state.calls.push('redo'),
        destroy: () => state.calls.push('destroy'),
      } as any;
    },
    fetchBytes: (url, limits) =>
      new Promise<Uint8Array>((resolve, reject) => {
        state.fetches.push({ url, signal: limits.signal, resolve, reject });
      }),
    onState: (value, code) => state.states.push(code === undefined ? value : `${value}: ${code}`),
    onLocale: (locale) => state.locales.push(locale),
    frame: (callback) => state.frames.push(callback),
    timer: (callback, ms) => state.timers.push({ callback, ms }),
  };
  return state;
}

const DATA = base64Encode(new Uint8Array([80, 75, 3, 4]));
const types = (f: Fake) => f.sent.map((message) => message.type);
/** Lets promise callbacks that are already due run. */
const settle = async () => {
  for (let i = 0; i < 20; i++) await Promise.resolve();
};

test('bridge: start sends ready once, with the version; hello asks for it again', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  bridge.start();
  bridge.start();
  await bridge.receive('{"type":"close"}');
  assert.deepEqual(f.sent, [{ type: 'ready', version: '9.9.9' }]);
  await bridge.receive('{"type":"hello"}');
  assert.deepEqual(f.sent, [{ type: 'ready', version: '9.9.9' }, { type: 'ready', version: '9.9.9' }]);

  // A hello that is the first thing to arrive is answered once.
  const g = fake();
  await createBridge(g.host).receive({ type: 'hello' });
  assert.deepEqual(g.sent, [{ type: 'ready', version: '9.9.9' }]);
});

test('bridge: ready waits for a channel that appears later, and comes first', async () => {
  const f = fake();
  f.channel = false;
  const bridge = createBridge(f.host);
  bridge.start();
  assert.equal(f.timers.length, 1);
  // Still no channel: one more try is scheduled each time.
  f.runRetries();
  assert.equal(f.timers.length, 1);
  assert.deepEqual(f.sent, []);
  await bridge.receive('{"type":"save","id":"s"}');
  f.channel = true;
  f.runRetries();
  assert.deepEqual(types(f), ['ready', 'error']);
  assert.equal(f.timers.length, 0);
});

test('bridge: the search for a channel ends, and receive() still delivers ready afterwards', async () => {
  const f = fake();
  f.channel = false;
  const bridge = createBridge(f.host);
  bridge.start();
  let rounds = 0;
  while (f.timers.length > 0 && rounds < 10000) {
    f.runRetries();
    rounds++;
  }
  assert.ok(rounds >= 100 && rounds < 1000, `${rounds} rounds`);
  f.channel = true;
  await bridge.receive('{"type":"close"}');
  assert.deepEqual(f.sent, [{ type: 'ready', version: '9.9.9' }]);
});

test('bridge: every delivery gives the search for a channel its full time back', async () => {
  const f = fake();
  f.channel = false;
  const bridge = createBridge(f.host);
  bridge.start();
  // Most of the time is used up before the channel appears.
  for (let i = 0; i < 150; i++) f.runRetries();
  f.channel = true;
  f.runRetries();
  assert.deepEqual(types(f), ['ready']);
  // The channel goes away again: the next message is looked after for the whole time, not the rest of it.
  f.channel = false;
  await bridge.receive('{"type":"save"}');
  let rounds = 0;
  while (f.timers.length > 0 && rounds < 10000) {
    f.runRetries();
    rounds++;
  }
  assert.ok(rounds >= 200, `${rounds} rounds`);
});

test('bridge: a channel that throws is treated as missing', async () => {
  const f = fake();
  let broken = true;
  const post = f.host.post;
  f.host.post = (json) => {
    if (broken) throw new Error('not attached');
    return post(json);
  };
  const bridge = createBridge(f.host);
  bridge.start();
  broken = false;
  f.runRetries();
  assert.deepEqual(f.sent, [{ type: 'ready', version: '9.9.9' }]);
  bridge.destroy();
});

test('bridge: without a channel, status messages are merged and answers are all kept, in order', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', id: 'l', data: DATA }));
  f.channel = false;
  const options = f.views[0];
  const cell = (row: number) => ({ active: { row, col: 0 }, range: { top: row, left: 0, bottom: row, right: 0 } });
  for (let i = 0; i < 100; i++) {
    options.onChange!({ sheets: [i % 2] });
    options.onSheetChange!(i % 2);
    options.onSelectionChange!(cell(i), 0);
    f.frames.shift()!();
    await bridge.receive(`{"type":"explode","id":"e${i}"}`);
  }
  await bridge.receive('{"type":"undo","id":"u"}');
  await bridge.receive('{"type":"save","id":"s"}');
  f.channel = true;
  f.runRetries();
  const late = f.sent.slice(2);
  // 100 errors, one done, one saved, and one of each kind of status: the newest.
  assert.equal(late.filter((m) => m.type === 'error').length, 100);
  assert.deepEqual(late.filter((m) => m.type === 'error').map((m) => m.id), Array.from({ length: 100 }, (_, i) => `e${i}`));
  assert.deepEqual(late.filter((m) => m.type === 'changed'), [{ type: 'changed', sheets: [0, 1], canUndo: false, canRedo: false }]);
  assert.deepEqual(late.filter((m) => m.type === 'sheet'), [{ type: 'sheet', index: 1, name: 'Gizli' }]);
  assert.deepEqual(late.filter((m) => m.type === 'selection'), [{ type: 'selection', sheet: 0, range: 'A100', active: 'A100' }]);
  assert.deepEqual(late.slice(-2).map((m) => [m.type, m.id]), [['done', 'u'], ['saved', 's']]);
  assert.equal(late.length, 105);
});

test('bridge: past the size limit the oldest saved files become save-failed errors; nothing else is dropped', async () => {
  const f = fake();
  f.host.maxPendingChars = 3000;
  f.saveResult = async () => new Uint8Array(900);
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', id: 'l', data: DATA }));
  f.channel = false;
  for (const id of ['s1', 's2', 's3', 's4']) await bridge.receive(`{"type":"save","id":"${id}"}`);
  await bridge.receive('{"type":"redo","id":"r"}');
  f.channel = true;
  f.runRetries();
  assert.deepEqual(
    f.sent.slice(2).map((m) => [m.type, m.id, m.code]),
    [
      ['error', 's1', 'save-failed'],
      ['error', 's2', 'save-failed'],
      ['saved', 's3', undefined],
      ['saved', 's4', undefined],
      ['done', 'r', undefined],
    ],
  );
});

test('bridge: load opens the decoded bytes, shows them and answers with loaded', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', id: '7', data: DATA, name: 'a.xlsx', locale: 'tr', readOnly: true, sheet: 1 }));
  assert.deepEqual(Array.from(f.opened[0]), [80, 75, 3, 4]);
  assert.equal(f.views[0].locale, 'tr');
  assert.equal(f.views[0].readOnly, true);
  assert.equal(f.views[0].sheet, 1);
  assert.deepEqual(f.sent[1], {
    type: 'loaded',
    id: '7',
    format: 'xlsx',
    sheets: [
      { name: 'Veri', state: 'visible', kind: 'worksheet' },
      { name: 'Gizli', state: 'hidden', kind: 'worksheet' },
    ],
    // A hidden sheet can be made the shown one this way.
    activeSheet: 1,
  });
  assert.deepEqual(f.states, ['loading', 'open']);
  assert.deepEqual(f.locales, ['tr']);
});

test('bridge: a message may arrive as an object as well as text', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  assert.deepEqual(types(f), ['ready', 'loaded']);
  assert.equal('id' in f.sent[1], false);
});

test('bridge: messages that break the contract get bad-message and change nothing', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const bad: unknown[] = [
    '{not json',
    '',
    '[]',
    'null',
    '42',
    42,
    null,
    undefined,
    '{}',
    '{"type":7}',
    '{"type":"explode"}',
    '{"type":"load"}',
    '{"type":"load","data":"AAAA","url":"a.xlsx"}',
    '{"type":"load","data":5}',
    '{"type":"load","data":"***"}',
    '{"type":"load","data":"AAAA","sheet":-1}',
    '{"type":"load","data":"AAAA","sheet":1.5}',
    '{"type":"load","data":"AAAA","readOnly":"yes"}',
    '{"type":"load","data":"AAAA","locale":3}',
    '{"type":"load","url":"javascript:alert(1)"}',
    '{"type":"load","url":"file:///etc/passwd"}',
    '{"type":"load","url":"//evil.example/a.xlsx"}',
    '{"type":"load","url":"https://user:pw@app.example.com/a.xlsx"}',
    '{"type":"save","id":5}',
    '{"type":"close","id":5}',
    '{"type":"hello","id":{}}',
    '{"type":"setReadOnly"}',
    '{"type":"setReadOnly","readOnly":"true"}',
    '{"type":"setLocale"}',
  ];
  for (const message of bad) await bridge.receive(message);
  assert.equal(f.sent.length, bad.length + 1);
  for (let i = 0; i < bad.length; i++) {
    assert.equal(f.sent[i + 1].type, 'error', String(bad[i]));
    assert.equal(f.sent[i + 1].code, 'bad-message', String(bad[i]));
    assert.equal(typeof f.sent[i + 1].message, 'string');
  }
  assert.equal(f.opened.length, 0);
  assert.equal(f.fetches.length, 0);
  assert.deepEqual(f.states, []);
});

test('bridge: errors carry the id of the message they answer', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive('{"type":"explode","id":"x1"}');
  await bridge.receive('{"type":"save","id":"x2"}');
  assert.deepEqual(f.sent.slice(1).map((m) => [m.id, m.code]), [['x1', 'bad-message'], ['x2', 'not-loaded']]);
});

test('bridge: commands without a file answer not-loaded; close is always fine', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  for (const message of ['{"type":"save"}', '{"type":"undo"}', '{"type":"redo"}', '{"type":"setReadOnly","readOnly":true}', '{"type":"setLocale","locale":"tr"}']) {
    await bridge.receive(message);
  }
  await bridge.receive('{"type":"close","id":"c"}');
  assert.deepEqual(f.sent.slice(1, 6).map((m) => m.code), ['not-loaded', 'not-loaded', 'not-loaded', 'not-loaded', 'not-loaded']);
  assert.deepEqual(f.sent[6], { type: 'done', id: 'c', command: 'close' });
  assert.deepEqual(f.calls, []);
});

test('bridge: files the engine refuses are unsupported, files it cannot read are corrupt', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  f.openError = Object.assign(new Error('Not a Word or Excel file'), { name: 'UnsupportedFileError' });
  await bridge.receive(JSON.stringify({ type: 'load', id: 'a', data: DATA }));
  f.openError = new Error('Central directory not found');
  await bridge.receive(JSON.stringify({ type: 'load', id: 'b', data: DATA }));
  f.openError = undefined;
  f.sheetNames = [];
  await bridge.receive(JSON.stringify({ type: 'load', id: 'c', data: DATA }));
  assert.deepEqual(f.sent.slice(1), [
    { type: 'error', id: 'a', code: 'unsupported', message: 'Not a Word or Excel file' },
    { type: 'error', id: 'b', code: 'corrupt', message: 'Central directory not found' },
    { type: 'error', id: 'c', code: 'corrupt', message: 'The workbook has no sheets' },
  ]);
  assert.deepEqual(f.states, ['loading', 'error: unsupported', 'loading', 'error: corrupt', 'loading', 'error: corrupt']);
  assert.deepEqual(f.calls, []);
  await bridge.receive('{"type":"save"}');
  assert.equal(f.sent[4].code, 'not-loaded');
});

test('bridge: a load that fails leaves the open file open, shown and savable', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA, name: 'kept.xlsx' }));
  f.openError = new Error('Central directory not found');
  await bridge.receive(JSON.stringify({ type: 'load', id: 'bad1', data: DATA }));
  f.openError = undefined;
  await bridge.receive(JSON.stringify({ type: 'load', id: 'bad2', data: DATA, sheet: 9 }));
  await bridge.receive(JSON.stringify({ type: 'load', id: 'bad3', data: '***' }));
  const missing = bridge.receive('{"type":"load","id":"bad4","url":"/missing.xlsx"}');
  await settle();
  f.fetches[0].reject(new Error('HTTP 404'));
  await missing;
  await bridge.receive('{"type":"save","id":"s"}');
  assert.deepEqual(f.sent.slice(2).map((m) => [m.id, m.code || m.type]), [
    ['bad1', 'corrupt'],
    ['bad2', 'bad-message'],
    ['bad3', 'bad-message'],
    ['bad4', 'load-failed'],
    ['s', 'saved'],
  ]);
  assert.equal(f.sent[6].name, 'kept.xlsx');
  // The view was created once and never destroyed; the page never left the open state.
  assert.deepEqual(f.calls, ['view']);
  assert.deepEqual(f.states, ['loading', 'open']);
});

test('bridge: load from a URL fetches the resolved address; a failed fetch is load-failed', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const first = bridge.receive('{"type":"load","url":"files/a.xlsx"}');
  await settle();
  assert.equal(f.fetches[0].url, 'https://app.example.com/files/a.xlsx');
  f.fetches[0].resolve(new Uint8Array([80, 75]));
  await first;
  const second = bridge.receive('{"type":"load","url":"files/missing.xlsx","id":"m"}');
  await settle();
  f.fetches[1].reject(new Error('HTTP 404'));
  await second;
  assert.equal(f.sent[1].type, 'loaded');
  assert.deepEqual(f.sent[2], { type: 'error', id: 'm', code: 'load-failed', message: 'Could not load the file: HTTP 404' });

  const g = fake();
  delete g.host.fetchBytes;
  await createBridge(g.host).receive('{"type":"load","url":"a.xlsx"}');
  assert.equal(g.sent[1].code, 'load-failed');
  assert.deepEqual(g.states, ['loading', 'error: load-failed']);
});

test('bridge: a framed page refuses every load with a url', async () => {
  const f = fake();
  f.host.allowUrl = false;
  const bridge = createBridge(f.host);
  await bridge.receive('{"type":"load","id":"u","url":"/files/a.xlsx"}');
  await bridge.receive('{"type":"load","id":"d","url":"data:application/octet-stream;base64,AAAA"}');
  assert.deepEqual(f.sent.slice(1).map((m) => [m.id, m.code]), [['u', 'bad-message'], ['d', 'bad-message']]);
  assert.equal(f.fetches.length, 0);
});

test('bridge: a fetch that hangs times out, is aborted, and the next message is handled', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const hung = bridge.receive('{"type":"load","id":"h","url":"/slow.xlsx"}');
  const after = bridge.receive('{"type":"save","id":"s"}');
  await settle();
  assert.equal(f.sent.length, 1);
  const limit = f.timers.find((entry) => entry.ms === 7000)!;
  limit.callback();
  await Promise.all([hung, after]);
  assert.equal(f.fetches[0].signal!.aborted, true);
  assert.deepEqual(f.sent.slice(1), [
    { type: 'error', id: 'h', code: 'load-failed', message: 'Could not load the file: no answer within 7 s' },
    { type: 'error', id: 's', code: 'not-loaded', message: 'No file is open' },
  ]);
  // An answer that arrives after all is ignored.
  f.fetches[0].resolve(new Uint8Array([80, 75]));
  await settle();
  assert.equal(f.sent.length, 3);
  assert.equal(f.opened.length, 0);
});

test('bridge: a fetched file larger than the limit is load-failed', async () => {
  const f = fake();
  f.host.maxFetchBytes = 10;
  const bridge = createBridge(f.host);
  const load = bridge.receive('{"type":"load","id":"big","url":"/big.xlsx"}');
  await settle();
  f.fetches[0].resolve(new Uint8Array(11));
  await load;
  assert.deepEqual(f.sent[1], { type: 'error', id: 'big', code: 'load-failed', message: 'Could not load the file: the file is larger than 10 bytes' });
  assert.equal(f.opened.length, 0);
});

test('bridge: close does not wait behind a hanging load, and cancels it', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const hung = bridge.receive('{"type":"load","id":"h","url":"/slow.xlsx"}');
  await settle();
  assert.deepEqual(f.states, ['loading']);
  await bridge.receive('{"type":"close","id":"c"}');
  // Answered at once, before the load has ended.
  assert.deepEqual(f.sent[1], { type: 'done', id: 'c', command: 'close' });
  assert.equal(f.fetches[0].signal!.aborted, true);
  await hung;
  assert.deepEqual(f.sent[2], { type: 'error', id: 'h', code: 'load-failed', message: 'The load was cancelled' });
  assert.deepEqual(f.states, ['loading', 'empty']);
  assert.equal(f.opened.length, 0);
  // The bridge goes on working.
  await bridge.receive(JSON.stringify({ type: 'load', id: 'next', data: DATA }));
  assert.equal(f.sent[3].type, 'loaded');
});

test('bridge: a second load cancels the first, hanging or waiting, and only it is shown', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const first = bridge.receive('{"type":"load","id":"1","url":"/slow.xlsx"}');
  await settle();
  assert.equal(f.fetches.length, 1);
  const second = bridge.receive(JSON.stringify({ type: 'load', id: '2', data: DATA }));
  const third = bridge.receive(JSON.stringify({ type: 'load', id: '3', data: DATA, name: 'third.xlsx' }));
  await Promise.all([first, second, third]);
  assert.deepEqual(f.sent.slice(1).map((m) => [m.id, m.code || m.type]), [
    ['1', 'load-failed'],
    ['2', 'load-failed'],
    ['3', 'loaded'],
  ]);
  assert.equal(f.fetches[0].signal!.aborted, true);
  assert.deepEqual(f.calls, ['view']);
  // A cancelled load without an id says nothing.
  const fourth = bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  const fifth = bridge.receive(JSON.stringify({ type: 'load', id: '5', data: DATA }));
  await Promise.all([fourth, fifth]);
  assert.deepEqual(f.sent.slice(4).map((m) => m.id), ['5']);
});

test('bridge: a load that is overtaken while the file is being read shows nothing', async () => {
  const f = fake();
  let release: () => void = () => undefined;
  const open = f.host.open;
  f.host.open = async (bytes) => {
    await new Promise<void>((resolve) => (release = resolve));
    return open(bytes);
  };
  const bridge = createBridge(f.host);
  const load = bridge.receive(JSON.stringify({ type: 'load', id: 'slow', data: DATA }));
  await settle();
  await bridge.receive('{"type":"close"}');
  release();
  await load;
  assert.deepEqual(f.sent[1], { type: 'error', id: 'slow', code: 'load-failed', message: 'The load was cancelled' });
  assert.deepEqual(f.calls, []);
});

test('bridge: destroy during a load aborts the fetch and never creates a view', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const hung = bridge.receive('{"type":"load","id":"h","url":"/slow.xlsx"}');
  await settle();
  bridge.destroy();
  assert.equal(f.fetches[0].signal!.aborted, true);
  f.fetches[0].resolve(new Uint8Array([80, 75]));
  await hung;
  assert.deepEqual(f.calls, []);
  assert.equal(f.opened.length, 0);
  assert.equal(f.sent.length, 1);

  // The same while the file is being read.
  const g = fake();
  let release: () => void = () => undefined;
  const open = g.host.open;
  g.host.open = async (bytes) => {
    await new Promise<void>((resolve) => (release = resolve));
    return open(bytes);
  };
  const other = createBridge(g.host);
  const load = other.receive(JSON.stringify({ type: 'load', data: DATA }));
  await settle();
  other.destroy();
  release();
  await load;
  assert.deepEqual(g.calls, []);
});

test('bridge: a view that cannot be created, or a page callback that throws, leaves a consistent state', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  f.viewError = new Error('no canvas');
  await bridge.receive(JSON.stringify({ type: 'load', id: 'v', data: DATA }));
  assert.deepEqual(f.sent[1], { type: 'error', id: 'v', code: 'corrupt', message: 'no canvas' });
  await bridge.receive('{"type":"save","id":"s"}');
  assert.equal(f.sent[2].code, 'not-loaded');

  f.viewError = undefined;
  f.host.onState = () => {
    throw new Error('page callback');
  };
  f.host.onLocale = () => {
    throw new Error('page callback');
  };
  await bridge.receive(JSON.stringify({ type: 'load', id: 'ok', data: DATA, locale: 'tr' }));
  await bridge.receive('{"type":"setLocale","id":"l","locale":"en"}');
  await bridge.receive('{"type":"save","id":"s2"}');
  await bridge.receive('{"type":"close","id":"c"}');
  assert.deepEqual(f.sent.slice(3).map((m) => [m.id, m.type]), [['ok', 'loaded'], ['l', 'done'], ['s2', 'saved'], ['c', 'done']]);
});

test('bridge: a sheet index the file does not have is refused', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA, sheet: 2 }));
  assert.deepEqual(f.sent[1], { type: 'error', code: 'bad-message', message: 'No sheet with index 2' });
  assert.deepEqual(f.calls, []);
  assert.deepEqual(f.states, ['loading', 'empty']);
});

test('bridge: save writes an open edit first and returns the bytes with the id and the name', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  // Sent without waiting: the save is handled after the load.
  const first = bridge.receive(JSON.stringify({ type: 'load', data: DATA, name: 'rapor.xlsx' }));
  f.editing = true;
  const second = bridge.receive('{"type":"save","id":"s1"}');
  await Promise.all([first, second]);
  assert.deepEqual(types(f).slice(1), ['loaded', 'saved']);
  assert.deepEqual(f.sent[2], { type: 'saved', id: 's1', data: 'AQID+g==', name: 'rapor.xlsx' });
  assert.deepEqual(f.calls, ['view', 'commit']);

  f.saveResult = async () => {
    throw new Error('disk on fire');
  };
  await bridge.receive('{"type":"save","id":"s2"}');
  assert.deepEqual(f.sent[3], { type: 'error', id: 's2', code: 'save-failed', message: 'disk on fire' });
});

test('bridge: changes are reported with the state of the undo history, sheet changes with the name', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  f.history.canUndo = true;
  f.views[0].onChange!({ sheets: [0] });
  f.views[0].onSheetChange!(1);
  assert.deepEqual(f.sent.slice(2), [
    { type: 'changed', sheets: [0], canUndo: true, canRedo: false },
    { type: 'sheet', index: 1, name: 'Gizli' },
  ]);
});

test('bridge: selection is sent once per frame, the last one only, and not twice the same', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  const select = f.views[0].onSelectionChange!;
  const cell = (row: number, col: number) => ({ active: { row, col }, range: { top: row, left: col, bottom: row + 1, right: col + 1 } });
  select(cell(0, 0), 0);
  select(cell(1, 1), 0);
  select(cell(2, 2), 1);
  assert.equal(f.frames.length, 1);
  assert.equal(f.sent.length, 2);
  f.frames.shift()!();
  assert.deepEqual(f.sent[2], { type: 'selection', sheet: 1, range: 'C3:D4', active: 'C3' });
  select(cell(2, 2), 1);
  f.frames.shift()!();
  assert.equal(f.sent.length, 3);
  // A frame that comes after the file was closed sends nothing.
  select(cell(5, 5), 0);
  await bridge.receive('{"type":"close"}');
  f.frames.shift()!();
  assert.equal(f.sent.length, 3);
});

test('bridge: commands reach the view; with an id they are confirmed, without one they are silent', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  for (const message of ['{"type":"setReadOnly","readOnly":true}', '{"type":"setLocale","locale":"en"}', '{"type":"undo"}', '{"type":"redo"}']) {
    await bridge.receive(message);
  }
  assert.deepEqual(f.calls, ['view', 'setReadOnly true', 'setLocale en', 'undo', 'redo']);
  assert.deepEqual(f.locales, ['en']);
  assert.equal(f.sent.length, 2);

  for (const message of ['{"type":"setReadOnly","readOnly":false,"id":"a"}', '{"type":"setLocale","locale":"tr","id":"b"}', '{"type":"undo","id":"c"}', '{"type":"redo","id":"d"}', '{"type":"close","id":"e"}', '{"type":"close"}']) {
    await bridge.receive(message);
  }
  assert.deepEqual(f.sent.slice(2), [
    { type: 'done', id: 'a', command: 'setReadOnly' },
    { type: 'done', id: 'b', command: 'setLocale' },
    { type: 'done', id: 'c', command: 'undo' },
    { type: 'done', id: 'd', command: 'redo' },
    { type: 'done', id: 'e', command: 'close' },
  ]);
  assert.deepEqual(f.states, ['loading', 'open', 'empty', 'empty']);
  assert.equal(f.calls[f.calls.length - 1], 'destroy');
});

test('bridge: a second load replaces the first file', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA, name: 'one.xlsx' }));
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  await bridge.receive('{"type":"save"}');
  // The first view goes only once the second file has been read.
  assert.deepEqual(f.calls, ['view', 'destroy', 'view']);
  assert.deepEqual(f.states, ['loading', 'open', 'open']);
  // The name of the first file is not carried over.
  assert.equal('name' in f.sent[3], false);
});

test('bridge: a view that throws does not break the bridge', async () => {
  const f = fake();
  const createView = f.host.createView;
  f.host.createView = (container, workbook, options) => {
    const view = createView(container, workbook, options);
    view.undo = () => {
      throw new Error('boom');
    };
    return view;
  };
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  await bridge.receive('{"type":"undo","id":"u"}');
  await bridge.receive('{"type":"redo"}');
  assert.deepEqual(f.sent[2], { type: 'error', id: 'u', code: 'bad-message', message: 'boom' });
  assert.equal(f.calls[f.calls.length - 1], 'redo');
});

test('bridge: after destroy nothing is handled or sent', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  bridge.destroy();
  bridge.destroy();
  await bridge.receive('{"type":"save"}');
  await bridge.receive('{"type":"hello"}');
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  assert.deepEqual(f.calls, ['view', 'destroy']);
  assert.equal(f.sent.length, 2);
});

// ------------------------------------------------------------------ documents

interface FakeDocument {
  file: OfficeFile;
  model: object;
  views: { container: unknown; doc: unknown; options: import('../src/ui/doc-view').DocumentViewOptions }[];
  calls: string[];
  documentError?: Error;
  viewError?: Error;
}

/** Makes the fake host open a docx and able to show it. */
function withDocument(f: Fake, canShow = true): FakeDocument {
  const state: FakeDocument = { file: undefined as unknown as OfficeFile, model: { body: { blocks: [] } }, views: [], calls: [] };
  state.file = {
    format: 'docx',
    kind: 'document',
    container: 'ooxml',
    package: { readPart: async () => new Uint8Array(0) },
    workbook: () => Promise.reject(Object.assign(new Error('This file is a document, not a spreadsheet'), { name: 'UnsupportedFileError' })),
    document: async () => {
      if (state.documentError) throw state.documentError;
      return state.model;
    },
    save: () => f.saveResult(),
  } as unknown as OfficeFile;
  f.host.open = async (bytes) => {
    f.opened.push(bytes);
    return state.file;
  };
  if (canShow) {
    f.host.createDocumentView = (container, doc, options) => {
      state.calls.push('view');
      if (state.viewError) throw state.viewError;
      state.views.push({ container, doc, options });
      return { destroy: () => state.calls.push('destroy'), setLocale: (locale: string) => state.calls.push(`setLocale ${locale}`) };
    };
  }
  return state;
}

test('bridge: a document is shown read-only and answered with loaded', async () => {
  const f = fake();
  const d = withDocument(f);
  const bridge = createBridge(f.host);
  // `readOnly` and `sheet` mean nothing for a document and are not complained about.
  await bridge.receive(JSON.stringify({ type: 'load', id: 'd1', data: DATA, name: 'a.docx', locale: 'tr', readOnly: false, sheet: 3 }));
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'd1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.deepEqual(f.states, ['loading', 'open']);
  assert.deepEqual(f.locales, ['tr']);
  assert.equal(d.views.length, 1);
  assert.equal(d.views[0].container, f.host.container);
  assert.equal(d.views[0].doc, d.model);
  assert.equal(d.views[0].options.locale, 'tr');
  // Pictures are read from the package of the file.
  assert.equal(d.views[0].options.package, d.file.package);
  // The spreadsheet view was never asked for.
  assert.deepEqual(f.calls, []);
});

test('bridge: a link in a document is reported to the application and never opened by the page', async () => {
  const f = fake();
  const d = withDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  const handler = d.views[0].options.onLinkClick!;
  assert.equal(handler('https://example.com/a?b=1'), false);
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'link', url: 'https://example.com/a?b=1' });
  // Every click is reported, also the same link twice.
  assert.equal(handler('https://example.com/a?b=1'), false);
  assert.deepEqual(types(f), ['ready', 'loaded', 'link', 'link']);
});

test('bridge: save returns the bytes of a document; it never sends changed', async () => {
  const f = fake();
  withDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, name: 'a.docx' });
  await bridge.receive({ type: 'save', id: 's1' });
  assert.deepEqual(f.sent[2], { type: 'saved', id: 's1', data: base64Encode(new Uint8Array([1, 2, 3, 250])), name: 'a.docx' });
  f.saveResult = () => Promise.reject(new Error('disk full'));
  await bridge.receive({ type: 'save', id: 's2' });
  assert.deepEqual(f.sent[3], { type: 'error', id: 's2', code: 'save-failed', message: 'disk full' });
  assert.equal(types(f).indexOf('changed'), -1);
  assert.equal(types(f).indexOf('selection'), -1);
});

test('bridge: commands on a document are confirmed; only the language has an effect', async () => {
  const f = fake();
  const d = withDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'setLocale', id: 'a', locale: 'en' });
  await bridge.receive({ type: 'setReadOnly', id: 'b', readOnly: false });
  await bridge.receive({ type: 'undo', id: 'c' });
  await bridge.receive({ type: 'redo', id: 'd' });
  await bridge.receive({ type: 'setLocale' });
  assert.deepEqual(d.calls, ['view', 'setLocale en']);
  assert.deepEqual(f.locales, ['en']);
  assert.deepEqual(f.sent.slice(2), [
    { type: 'done', id: 'a', command: 'setLocale' },
    { type: 'done', id: 'b', command: 'setReadOnly' },
    { type: 'done', id: 'c', command: 'undo' },
    { type: 'done', id: 'd', command: 'redo' },
    { type: 'error', code: 'bad-message', message: '"locale" is missing' },
  ]);
  await bridge.receive({ type: 'close', id: 'e' });
  assert.deepEqual(d.calls, ['view', 'setLocale en', 'destroy']);
  await bridge.receive({ type: 'save', id: 'f' });
  assert.equal(f.sent[f.sent.length - 1].code, 'not-loaded');
});

test('bridge: a document replaces a workbook and a workbook a document', async () => {
  const f = fake();
  const open = f.host.open;
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  assert.deepEqual(f.calls, ['view']);
  const d = withDocument(f);
  await bridge.receive({ type: 'load', data: DATA });
  assert.deepEqual(f.calls, ['view', 'destroy']);
  assert.deepEqual(d.calls, ['view']);
  f.host.open = open;
  await bridge.receive({ type: 'load', data: DATA });
  assert.deepEqual(d.calls, ['view', 'destroy']);
  assert.deepEqual(f.calls, ['view', 'destroy', 'view']);
  // Commands reach the workbook again.
  await bridge.receive({ type: 'undo' });
  assert.deepEqual(f.calls, ['view', 'destroy', 'view', 'undo']);
  bridge.destroy();
  assert.deepEqual(f.calls, ['view', 'destroy', 'view', 'undo', 'destroy']);
});

test('bridge: a page that cannot show documents refuses them as unsupported and keeps what is open', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  withDocument(f, false);
  await bridge.receive({ type: 'load', id: 'x', data: DATA });
  assert.deepEqual(f.sent[2], { type: 'error', id: 'x', code: 'unsupported', message: 'This kind of document cannot be shown' });
  // The workbook that was open stays open.
  assert.deepEqual(f.calls, ['view']);
  assert.deepEqual(f.states, ['loading', 'open']);

  // A document format without a reader yet (doc) has no document(): unsupported as well.
  const g = fake();
  const d = withDocument(g);
  delete (d.file as { document?: unknown }).document;
  const second = createBridge(g.host);
  await second.receive({ type: 'load', id: 'y', data: DATA });
  assert.equal(g.sent[1].code, 'unsupported');
  assert.deepEqual(g.states, ['loading', 'error: unsupported']);
  assert.deepEqual(d.calls, []);
});

test('bridge: a document that cannot be read or shown is corrupt', async () => {
  const f = fake();
  const d = withDocument(f);
  const bridge = createBridge(f.host);
  d.documentError = new Error('Invalid XML at offset 12');
  await bridge.receive({ type: 'load', id: 'a', data: DATA });
  assert.deepEqual(f.sent[1], { type: 'error', id: 'a', code: 'corrupt', message: 'Invalid XML at offset 12' });
  assert.deepEqual(d.calls, []);
  d.documentError = undefined;
  d.viewError = new Error('no canvas');
  await bridge.receive({ type: 'load', id: 'b', data: DATA });
  assert.deepEqual(f.sent[2], { type: 'error', id: 'b', code: 'corrupt', message: 'no canvas' });
  assert.deepEqual(f.states, ['loading', 'error: corrupt', 'loading', 'error: corrupt']);
  d.viewError = undefined;
  await bridge.receive({ type: 'load', id: 'c', data: DATA });
  assert.equal(f.sent[3].type, 'loaded');
  assert.deepEqual(d.calls, ['view', 'view']);
});

test('bridge: a document load that is overtaken shows nothing', async () => {
  const f = fake();
  const d = withDocument(f);
  let release: (() => void) | undefined;
  const reading = new Promise<void>((resolve) => (release = resolve));
  const read = d.file.document!;
  (d.file as { document: unknown }).document = async () => {
    await reading;
    return read();
  };
  const bridge = createBridge(f.host);
  const first = bridge.receive({ type: 'load', id: 'slow', data: DATA });
  await settle();
  await bridge.receive({ type: 'close', id: 'c' });
  release!();
  await first;
  assert.deepEqual(d.calls, []);
  assert.deepEqual(
    f.sent.slice(1).map((message) => [message.type, message.id, message.code]),
    [
      ['done', 'c', undefined],
      ['error', 'slow', 'load-failed'],
    ],
  );
});
