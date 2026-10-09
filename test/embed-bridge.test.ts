import assert from 'node:assert/strict';
import { test } from 'node:test';
import { base64Decode, base64Encode, createBridge, isAllowedUrl, normalizeOrigin, resolveLoadUrl, type BridgeHost } from '../src/embed/bridge';
import { open as openFile, type OfficeFile } from '../src/index';
import { openConverted } from '../src/formats/convert';
import { simpleDoc } from './doc-binary-fixtures';
import { cat, number, workbookStream, xlsFile } from './xls-fixtures';
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
  assert.ok(elapsed < 50000, `took ${elapsed} ms`);
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

test('bridge: a docx is shown for editing and answered with loaded', async () => {
  const f = fake();
  const d = withDocument(f);
  const bridge = createBridge(f.host);
  // `sheet` means nothing for a document and is not complained about.
  await bridge.receive(JSON.stringify({ type: 'load', id: 'd1', data: DATA, name: 'a.docx', locale: 'tr', readOnly: false, sheet: 3 }));
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'd1', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.deepEqual(f.states, ['loading', 'open']);
  assert.deepEqual(f.locales, ['tr']);
  assert.equal(d.views.length, 1);
  assert.equal(d.views[0].container, f.host.container);
  assert.equal(d.views[0].doc, d.model);
  assert.equal(d.views[0].options.locale, 'tr');
  assert.equal(d.views[0].options.readOnly, false);
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

test('bridge: save returns the bytes of a document; a document that is not edited never sends changed', async () => {
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

test('bridge: commands on a document view that cannot edit are confirmed; only the language has an effect', async () => {
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

/** Makes the fake host open every file as an xls workbook. */
function withXls(f: Fake): void {
  f.host.open = async (bytes) => {
    f.opened.push(bytes);
    return {
      format: 'xls',
      workbook: async () => ({ sheets: [{ name: 'Sheet1', state: 'visible', kind: 'worksheet' }] }),
      save: () => f.saveResult(),
    } as unknown as OfficeFile;
  };
}

test('bridge: an xls file is shown read-only whatever the load asks for, and loaded says so', async () => {
  for (const asked of [false, true, undefined]) {
    const f = fake();
    withXls(f);
    const bridge = createBridge(f.host);
    const message: Record<string, unknown> = { type: 'load', id: 'x', data: DATA };
    if (asked !== undefined) message.readOnly = asked;
    await bridge.receive(JSON.stringify(message));
    // The view is never created editable.
    assert.equal(f.views[0].readOnly, true, String(asked));
    assert.deepEqual(f.sent[1], { type: 'loaded', id: 'x', format: 'xls', sheets: [{ name: 'Sheet1', state: 'visible', kind: 'worksheet' }], activeSheet: 0, readOnly: true });
  }
  // A file that can be edited states nothing: what the load asked for holds.
  const g = fake();
  await createBridge(g.host).receive(JSON.stringify({ type: 'load', id: 'y', data: DATA, readOnly: false }));
  assert.equal(g.views[0].readOnly, false);
  assert.equal('readOnly' in g.sent[1], false);
});

test('bridge: setReadOnly cannot make an xls file editable; the confirmation says that it stays read-only', async () => {
  const f = fake();
  withXls(f);
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', data: DATA }));
  await bridge.receive('{"type":"setReadOnly","readOnly":false,"id":"a"}');
  await bridge.receive('{"type":"setReadOnly","readOnly":false}');
  await bridge.receive('{"type":"setReadOnly","readOnly":true,"id":"b"}');
  // The view is told "read-only" every time, never "editable".
  assert.deepEqual(f.calls, ['view', 'setReadOnly true', 'setReadOnly true', 'setReadOnly true']);
  assert.deepEqual(f.sent.slice(2), [
    { type: 'done', id: 'a', command: 'setReadOnly', readOnly: true },
    // What was asked for is what holds: confirmed as for any other file.
    { type: 'done', id: 'b', command: 'setReadOnly' },
  ]);
  // The other commands are not touched by it.
  await bridge.receive('{"type":"undo","id":"u"}');
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'done', id: 'u', command: 'undo' });

  // An xlsx file follows the request, and its confirmation carries nothing more.
  const g = fake();
  const other = createBridge(g.host);
  await other.receive(JSON.stringify({ type: 'load', data: DATA }));
  await other.receive('{"type":"setReadOnly","readOnly":false,"id":"a"}');
  assert.deepEqual(g.calls, ['view', 'setReadOnly false']);
  assert.deepEqual(g.sent[g.sent.length - 1], { type: 'done', id: 'a', command: 'setReadOnly' });
});

// --------------------------------------------------------- editing a document

interface EditableDocument extends FakeDocument {
  /** What the view says about itself; the tests change it as an edit would. */
  state: { canUndo: boolean; canRedo: boolean; changeCount: number; saved: number };
  /** The options of the views made with the toolbar. */
  editors: import('../src/ui/doc-editor').DocumentEditorOptions[];
}

/** Makes the fake host open a document of `format` in a view that can edit. */
function withEditableDocument(f: Fake, format = 'docx'): EditableDocument {
  const d = withDocument(f) as EditableDocument;
  (d.file as { format: string }).format = format;
  d.state = { canUndo: false, canRedo: false, changeCount: 0, saved: 0 };
  d.editors = [];
  const make = (kind: string) => (container: unknown, doc: unknown, options: any) => {
    d.calls.push(kind);
    if (d.viewError) throw d.viewError;
    if (kind === 'editor') d.editors.push(options);
    d.views.push({ container, doc, options });
    return {
      destroy: () => d.calls.push('destroy'),
      setLocale: (locale: string) => d.calls.push(`setLocale ${locale}`),
      setReadOnly: (value: boolean) => d.calls.push(`setReadOnly ${value}`),
      undo: () => d.calls.push('undo'),
      redo: () => d.calls.push('redo'),
      commitEditing: () => d.calls.push('commit'),
      markSaved: () => d.state.saved++,
      get canUndo() {
        return d.state.canUndo;
      },
      get canRedo() {
        return d.state.canRedo;
      },
      get changeCount() {
        return d.state.changeCount;
      },
    };
  };
  f.host.createDocumentView = make('view');
  f.host.createDocumentEditor = make('editor');
  return d;
}

test('bridge: a change of a document is reported like one of a workbook, with no sheets', async () => {
  const f = fake();
  const d = withEditableDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'l', data: DATA });
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'l', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.equal(d.views[0].options.readOnly, false);
  d.state.canUndo = true;
  d.views[0].options.onChange!();
  d.state.canUndo = false;
  d.state.canRedo = true;
  d.views[0].options.onChange!();
  assert.deepEqual(f.sent.slice(2), [
    { type: 'changed', sheets: [], canUndo: true, canRedo: false },
    { type: 'changed', sheets: [], canUndo: false, canRedo: true },
  ]);
  // Without a channel the newest one stands for those that waited.
  f.channel = false;
  d.views[0].options.onChange!();
  d.state.canUndo = true;
  d.views[0].options.onChange!();
  f.channel = true;
  f.runRetries();
  assert.deepEqual(f.sent.slice(4), [{ type: 'changed', sheets: [], canUndo: true, canRedo: true }]);
  // After the file was closed a late callback of the view says nothing.
  await bridge.receive({ type: 'close' });
  const count = f.sent.length;
  d.views[0].options.onChange!();
  assert.equal(f.sent.length, count);
});

test('bridge: undo, redo and setReadOnly reach a document; readOnly in the load opens it read-only', async () => {
  const f = fake();
  const d = withEditableDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, readOnly: true });
  assert.equal(d.views[0].options.readOnly, true);
  // A docx says nothing about being read-only: what the load asked for holds, as for an xlsx.
  assert.equal('readOnly' in f.sent[1], false);
  await bridge.receive({ type: 'setReadOnly', id: 'a', readOnly: false });
  await bridge.receive({ type: 'undo', id: 'b' });
  await bridge.receive({ type: 'redo' });
  await bridge.receive({ type: 'setReadOnly', readOnly: true });
  await bridge.receive({ type: 'setLocale', locale: 'tr' });
  assert.deepEqual(d.calls, ['view', 'setReadOnly false', 'undo', 'redo', 'setReadOnly true', 'setLocale tr']);
  assert.deepEqual(f.sent.slice(2), [
    { type: 'done', id: 'a', command: 'setReadOnly' },
    { type: 'done', id: 'b', command: 'undo' },
  ]);
});

test('bridge: saving a document ends a composition first; a change made while it is written is not marked as saved', async () => {
  const f = fake();
  const d = withEditableDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, name: 'a.docx' });
  await bridge.receive({ type: 'save', id: 's1' });
  assert.deepEqual(d.calls, ['view', 'commit']);
  assert.deepEqual(f.sent[2], { type: 'saved', id: 's1', data: 'AQID+g==', name: 'a.docx' });
  assert.equal(d.state.saved, 1);

  // The user types while the file is being written.
  let release: ((bytes: Uint8Array) => void) | undefined;
  f.saveResult = () => new Promise<Uint8Array>((resolve) => (release = resolve));
  const saving = bridge.receive({ type: 'save', id: 's2' });
  await settle();
  d.state.changeCount++;
  release!(new Uint8Array([9]));
  await saving;
  assert.equal(f.sent[3].type, 'saved');
  assert.equal(d.state.saved, 1);

  // A save that fails marks nothing either.
  f.saveResult = () => Promise.reject(new Error('disk full'));
  await bridge.receive({ type: 'save', id: 's3' });
  assert.deepEqual(f.sent[4], { type: 'error', id: 's3', code: 'save-failed', message: 'disk full' });
  assert.equal(d.state.saved, 1);
});

test('bridge: doc and RTF stay read-only whatever is asked, and loaded and done say so', async () => {
  for (const format of ['doc', 'rtf']) {
    const f = fake();
    const d = withEditableDocument(f, format);
    const bridge = createBridge(f.host);
    await bridge.receive({ type: 'load', id: 'l', data: DATA, readOnly: false });
    assert.equal(d.views[0].options.readOnly, true, format);
    assert.deepEqual(f.sent[1], { type: 'loaded', id: 'l', format, kind: 'document', sheets: [], activeSheet: 0, readOnly: true });
    await bridge.receive({ type: 'setReadOnly', id: 'a', readOnly: false });
    await bridge.receive({ type: 'setReadOnly', readOnly: false });
    await bridge.receive({ type: 'setReadOnly', id: 'b', readOnly: true });
    // The view is told "read-only" every time, never "editable".
    assert.deepEqual(d.calls, ['view', 'setReadOnly true', 'setReadOnly true', 'setReadOnly true']);
    assert.deepEqual(f.sent.slice(2), [
      { type: 'done', id: 'a', command: 'setReadOnly', readOnly: true },
      { type: 'done', id: 'b', command: 'setReadOnly' },
    ]);
  }
});

// ------------------------------------------------- convert, blank and toolbar

const RTF = new TextEncoder().encode('{\\rtf1\\ansi Hello {\\b rich} text\\par}');
const docBytes = () => simpleDoc('First paragraph\rSecond paragraph\r').build();
const xlsBytes = () => xlsFile(workbookStream({ sheets: [{ name: 'Old', records: cat(number(0, 0, 41)) }, { name: 'Chart1', type: 2 }] }));

/** A host that opens files with the engine itself, in fake views that can edit. */
function real(): { f: Fake; d: EditableDocument; opened: OfficeFile[] } {
  const f = fake();
  const d = withEditableDocument(f);
  const opened: OfficeFile[] = [];
  f.host.open = async (bytes) => {
    const file = await openFile(bytes);
    opened.push(file);
    return file;
  };
  return { f, d, opened };
}

test('bridge: convert opens doc, RTF and xls for editing as docx and xlsx; loaded and saved name both formats', async () => {
  for (const [from, bytes, to] of [
    ['rtf', RTF, 'docx'],
    ['doc', docBytes(), 'docx'],
    ['xls', xlsBytes(), 'xlsx'],
  ] as const) {
    const { f, d, opened } = real();
    const bridge = createBridge(f.host);
    await bridge.receive({ type: 'load', id: 'l', data: base64Encode(bytes), name: `old.${from}`, convert: true });
    const loaded = f.sent[1];
    assert.equal(loaded.type, 'loaded', `${from}: ${JSON.stringify(loaded)}`);
    assert.equal(loaded.format, to, from);
    assert.equal(loaded.convertedFrom, from);
    assert.equal('readOnly' in loaded, false, from);
    assert.ok(Array.isArray(loaded.warnings), from);
    for (const warning of loaded.warnings) assert.equal(typeof warning, 'string');
    // The file given and the file made from it.
    assert.deepEqual(opened.map((file) => file.format), [from, to]);
    if (to === 'docx') {
      assert.equal(loaded.kind, 'document');
      assert.equal(d.views[0].options.readOnly, false, from);
      assert.deepEqual(loaded.warnings, [], from);
    } else {
      assert.deepEqual(loaded.sheets.map((sheet: { name: string }) => sheet.name), ['Old', 'Chart1']);
      assert.equal(f.views[0].readOnly, undefined);
      // The chart sheet is not in the model.
      assert.equal(loaded.warnings.length, 1);
      assert.match(loaded.warnings[0], /chart sheet/);
    }
    // It can be made read-only and editable again like any docx or xlsx.
    await bridge.receive({ type: 'setReadOnly', id: 'r', readOnly: false });
    assert.deepEqual(f.sent[2], { type: 'done', id: 'r', command: 'setReadOnly' });

    await bridge.receive({ type: 'save', id: 's' });
    const saved = f.sent[3];
    assert.deepEqual({ ...saved, data: '' }, { type: 'saved', id: 's', data: '', name: `old.${from}`, format: to, convertedFrom: from });
    assert.equal((await openFile(base64Decode(saved.data))).format, to);
  }
});

test('bridge: convert does nothing for a read-only load and for files that need no conversion', async () => {
  const { f, d, opened } = real();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'a', data: base64Encode(RTF), convert: true, readOnly: true });
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'a', format: 'rtf', kind: 'document', sheets: [], activeSheet: 0, readOnly: true });
  assert.equal(d.views[0].options.readOnly, true);
  await bridge.receive({ type: 'load', id: 'b', data: base64Encode(RTF), convert: false });
  assert.equal(f.sent[2].format, 'rtf');
  await bridge.receive({ type: 'load', id: 'c', data: base64Encode(RTF) });
  assert.equal(f.sent[3].format, 'rtf');
  await bridge.receive({ type: 'save', id: 's' });
  // The file as it came, and a message without the fields of a conversion.
  assert.deepEqual(f.sent[4], { type: 'saved', id: 's', data: base64Encode(RTF) });

  await bridge.receive({ type: 'load', id: 'd', blank: 'docx', convert: true });
  assert.deepEqual(f.sent[5], { type: 'loaded', id: 'd', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.deepEqual(opened.map((file) => file.format), ['rtf', 'rtf', 'rtf', 'docx']);
});

test('bridge: a file that cannot be converted is shown as it is, read-only, and loaded says why', async () => {
  const { f, d } = real();
  const open = f.host.open;
  let calls = 0;
  f.host.open = (bytes) => (++calls === 2 ? Promise.reject(new Error('no room')) : open(bytes));
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'l', data: base64Encode(RTF), convert: true });
  assert.deepEqual(f.sent[1], {
    type: 'loaded',
    id: 'l',
    format: 'rtf',
    kind: 'document',
    sheets: [],
    activeSheet: 0,
    readOnly: true,
    warnings: ['The file could not be converted and is shown read-only: no room'],
  });
  assert.equal(d.views[0].options.readOnly, true);
  await bridge.receive({ type: 'setReadOnly', id: 'r', readOnly: false });
  assert.deepEqual(f.sent[2], { type: 'done', id: 'r', command: 'setReadOnly', readOnly: true });
  assert.deepEqual(f.states, ['loading', 'open']);
});

test('openConverted: converts what needs it, passes on what does not, and falls back when the new file cannot be read', async () => {
  const docx = await openFile(await (await openFile(RTF)).saveAsDocx!());
  assert.deepEqual(await openConverted(docx, openFile), { file: docx });

  const rtf = await openFile(RTF);
  const converted = await openConverted(rtf, openFile);
  assert.equal(converted.file.format, 'docx');
  assert.deepEqual(converted.conversion, { from: 'rtf', to: 'docx', warnings: [] });
  assert.equal(converted.failed, undefined);

  // The new file opens as something else, or its model cannot be read.
  const wrong = await openConverted(rtf, async () => rtf);
  assert.deepEqual(wrong, { file: rtf, failed: 'The converted file is not a docx file' });
  const unreadable = await openConverted(rtf, async (bytes) => {
    const file = await openFile(bytes);
    return { ...file, document: () => Promise.reject(new Error('Invalid XML')) };
  });
  assert.deepEqual(unreadable, { file: rtf, failed: 'Invalid XML' });

  // A file whose own model cannot be read is not something to fall back to.
  const broken = { ...rtf, document: () => Promise.reject(new Error('Invalid RTF')) };
  await assert.rejects(openConverted(broken, openFile), /Invalid RTF/);
});

test('bridge: a conversion that is overtaken shows nothing and answers its id', async () => {
  const { f, d } = real();
  const open = f.host.open;
  let release: (() => void) | undefined;
  const held = new Promise<void>((resolve) => (release = resolve));
  let calls = 0;
  let reached: (() => void) | undefined;
  const second = new Promise<void>((resolve) => (reached = resolve));
  // The converted file is being opened when the application closes.
  f.host.open = async (bytes) => {
    if (++calls === 2) {
      reached!();
      await held;
    }
    return open(bytes);
  };
  const bridge = createBridge(f.host);
  const first = bridge.receive({ type: 'load', id: 'slow', data: base64Encode(RTF), convert: true });
  await second;
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
  assert.deepEqual(f.states, ['loading', 'empty']);
});

test('bridge: blank starts with a new, empty docx or xlsx in the language of the load', async () => {
  const { f, d, opened } = real();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'x', blank: 'xlsx', locale: 'tr', name: 'yeni.xlsx' });
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'x', format: 'xlsx', sheets: [{ name: 'Sayfa1', state: 'visible', kind: 'worksheet' }], activeSheet: 0 });
  assert.equal(f.views[0].locale, 'tr');
  assert.equal(f.fetches.length, 0);
  await bridge.receive({ type: 'load', id: 'e', blank: 'xlsx' });
  assert.equal(f.sent[2].sheets[0].name, 'Sheet1');

  await bridge.receive({ type: 'load', id: 'w', blank: 'docx' });
  assert.deepEqual(f.sent[3], { type: 'loaded', id: 'w', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
  assert.equal(d.views[0].options.readOnly, false);
  await bridge.receive({ type: 'load', blank: 'docx', readOnly: true });
  assert.equal(d.views[1].options.readOnly, true);
  assert.deepEqual(opened.map((file) => file.format), ['xlsx', 'xlsx', 'docx', 'docx']);
  // The new file is saved like any other.
  await bridge.receive({ type: 'save', id: 's' });
  assert.equal((await openFile(base64Decode(f.sent[f.sent.length - 1].data))).format, 'docx');
});

test('bridge: a blank load that is overtaken shows nothing', async () => {
  const { f } = real();
  const bridge = createBridge(f.host);
  const first = bridge.receive({ type: 'load', id: 'one', blank: 'xlsx' });
  const second = bridge.receive({ type: 'load', id: 'two', blank: 'xlsx' });
  await Promise.all([first, second]);
  assert.deepEqual(
    f.sent.slice(1).map((message) => [message.type, message.id, message.code]),
    [
      ['error', 'one', 'load-failed'],
      ['loaded', 'two', undefined],
    ],
  );
  assert.deepEqual(f.calls, ['view']);
});

test('bridge: the new fields of load are checked; exactly one of data, url and blank is needed', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const bad = [
    { type: 'load', blank: 'pdf' },
    { type: 'load', blank: '' },
    { type: 'load', blank: 5 },
    { type: 'load', blank: true },
    { type: 'load', blank: 'docx', data: DATA },
    { type: 'load', blank: 'xlsx', url: 'a.xlsx' },
    { type: 'load', blank: 'xlsx', url: 'a.xlsx', data: DATA },
    { type: 'load', data: DATA, convert: 'yes' },
    { type: 'load', data: DATA, convert: 1 },
    { type: 'load', data: DATA, toolbar: 'true' },
    { type: 'load', data: DATA, toolbar: 0 },
    { type: 'load', data: DATA, toolbar: {} },
  ];
  for (const message of bad) await bridge.receive(message);
  assert.equal(f.sent.length, bad.length + 1);
  for (let i = 0; i < bad.length; i++) assert.equal(f.sent[i + 1].code, 'bad-message', JSON.stringify(bad[i]));
  assert.equal(f.sent[1].message, '"blank" must be "docx" or "xlsx"');
  assert.equal(f.sent[5].message, '"load" needs exactly one of "data", "url" and "blank"');
  assert.equal(f.sent[8].message, '"convert" must be a boolean');
  assert.equal(f.opened.length, 0);
  assert.deepEqual(f.states, []);
  // null counts as "not given", as for every other optional field.
  await bridge.receive({ type: 'load', data: DATA, blank: null, convert: null, toolbar: null });
  assert.equal(f.sent[f.sent.length - 1].type, 'loaded');
});

test('bridge: toolbar chooses the editor with its shell; without it, or on a host without one, the plain view is used', async () => {
  const f = fake();
  const editors: WorkbookViewOptions[] = [];
  const createView = f.host.createView;
  f.host.createEditor = (container, workbook, options) => {
    editors.push(options);
    f.calls.push('editor');
    return createView(container, workbook, options);
  };
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'load', data: DATA, toolbar: false });
  assert.deepEqual(f.calls.filter((call) => call !== 'destroy'), ['view', 'view']);
  await bridge.receive({ type: 'load', id: 't', data: DATA, toolbar: true, locale: 'tr', readOnly: true, sheet: 1 });
  assert.deepEqual(f.calls.filter((call) => call !== 'destroy'), ['view', 'view', 'editor', 'view']);
  assert.equal(editors.length, 1);
  assert.deepEqual([editors[0].locale, editors[0].readOnly, editors[0].sheet], ['tr', true, 1]);
  // The answer and the messages are the same in both modes.
  assert.deepEqual(f.sent[3], { type: 'loaded', id: 't', format: 'xlsx', sheets: f.sent[1].sheets, activeSheet: 1 });
  f.history.canUndo = true;
  editors[0].onChange!({ sheets: [1] });
  assert.deepEqual(f.sent[4], { type: 'changed', sheets: [1], canUndo: true, canRedo: false });
  await bridge.receive({ type: 'undo', id: 'u' });
  assert.equal(f.calls[f.calls.length - 1], 'undo');

  // Documents.
  const d = withEditableDocument(f);
  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'load', data: DATA, toolbar: true });
  assert.deepEqual(d.calls.filter((call) => call !== 'destroy'), ['view', 'editor']);
  d.state.canUndo = true;
  d.editors[0].onChange!();
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'changed', sheets: [], canUndo: true, canRedo: false });
  await bridge.receive({ type: 'redo' });
  assert.equal(d.calls[d.calls.length - 1], 'redo');

  // A host that only has the plain views shows them, whatever is asked.
  const g = fake();
  const e = withDocument(g);
  const plain = createBridge(g.host);
  await plain.receive({ type: 'load', data: DATA, toolbar: true });
  assert.deepEqual(e.calls, ['view']);
  const h = fake();
  await createBridge(h.host).receive({ type: 'load', data: DATA, toolbar: true });
  assert.deepEqual(h.calls, ['view']);
});

test('review: a change waiting for a channel is not joined with the change of the next file', async () => {
  const { f, d } = real();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'x', blank: 'xlsx' });
  f.channel = false;
  f.views[0].onChange!({ sheets: [0] });
  await bridge.receive({ type: 'load', id: 'w', blank: 'docx' });
  d.views[0].options.onChange!();
  d.views[0].options.onChange!();
  f.channel = true;
  f.runRetries();
  const late = f.sent.slice(2);
  // The workbook's change, then the document: its changes name no sheet, and the two of them are one.
  assert.deepEqual(
    late.map((m) => [m.type, JSON.stringify(m.sheets)]),
    [
      ['changed', '[0]'],
      ['loaded', '[]'],
      ['changed', '[]'],
    ],
  );
});

test('bridge: print starts the print job of a document and answers when its dialog was closed; other messages do not wait', async () => {
  const f = fake();
  const d = withEditableDocument(f);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  const jobs: { options: unknown; resolve(): void; reject(error: unknown): void }[] = [];
  const view = d.views[0];
  const make = f.host.createDocumentView!;
  // The view of the next load can print.
  f.host.createDocumentView = (container, doc, options) => {
    const pages = make.call(f.host, container, doc, options);
    pages.print = (printOptions) => new Promise<void>((resolve, reject) => jobs.push({ options: printOptions, resolve, reject }));
    return pages;
  };
  // A view without print(): the page says that it cannot.
  assert.ok(view);
  await bridge.receive({ type: 'print', id: 'p0' });
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'error', id: 'p0', code: 'print-unavailable', message: 'This view cannot print' });

  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'print', id: 'p1' });
  await bridge.receive({ type: 'print', id: 'p2', dpi: 150 });
  assert.deepEqual(
    jobs.map((job) => job.options),
    [undefined, { dpi: 150 }],
  );
  // The dialog is still open: no answer yet, and the next message is handled all the same.
  const before = f.sent.length;
  await bridge.receive({ type: 'undo', id: 'u' });
  assert.deepEqual(f.sent.slice(before), [{ type: 'done', id: 'u', command: 'undo' }]);
  jobs[0].resolve();
  jobs[1].reject(Object.assign(new Error('The window has no print()'), { code: 'unavailable' }));
  await settle();
  assert.deepEqual(f.sent.slice(before + 1), [
    { type: 'done', id: 'p1', command: 'print' },
    { type: 'error', id: 'p2', code: 'print-unavailable', message: 'The window has no print()' },
  ]);
  // Any other failure, and one that is thrown at once.
  await bridge.receive({ type: 'print', id: 'p3' });
  jobs[2].reject(new Error('Page 2 could not be painted'));
  await settle();
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'error', id: 'p3', code: 'print-failed', message: 'Page 2 could not be painted' });
  // Without an id a print that succeeds is silent, like every command.
  const count = f.sent.length;
  await bridge.receive({ type: 'print' });
  jobs[3].resolve();
  await settle();
  assert.equal(f.sent.length, count);
});

test('bridge: print is refused without a file, by a view that cannot print and with a wrong dpi or a wrong choice of sheets', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'print', id: 'a' });
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'error', id: 'a', code: 'not-loaded', message: 'No file is open' });
  await bridge.receive({ type: 'load', data: DATA });
  // The view of this host has no print().
  await bridge.receive({ type: 'print', id: 'b' });
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'error', id: 'b', code: 'print-unavailable', message: 'This view cannot print' });
  for (const dpi of [0, -5, '200', true]) {
    await bridge.receive({ type: 'print', id: 'c', dpi });
    assert.equal(f.sent[f.sent.length - 1].code, 'bad-message', String(dpi));
  }
  for (const sheets of ['', 'visible', 0, true, ['all']]) {
    await bridge.receive({ type: 'print', id: 'd', sheets });
    assert.equal(f.sent[f.sent.length - 1].code, 'bad-message', String(sheets));
  }
  // The workbook is still there.
  assert.equal(f.calls.includes('destroy'), false);
});

test('bridge: print of a workbook hands dpi, the choice of sheets and the name of the file to the view', async () => {
  const f = fake();
  const jobs: { options: unknown; resolve(): void; reject(error: unknown): void }[] = [];
  const make = f.host.createView;
  f.host.createView = (container, workbook, options) => {
    const view = make.call(f.host, container, workbook, options);
    view.print = ((printOptions: unknown) => new Promise<void>((resolve, reject) => jobs.push({ options: printOptions, resolve, reject }))) as unknown as typeof view.print;
    return view;
  };
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA, name: 'Budget 2026.xlsx' });
  await bridge.receive({ type: 'print', id: 'p1' });
  await bridge.receive({ type: 'print', id: 'p2', dpi: 150, sheets: 'all' });
  await bridge.receive({ type: 'print', id: 'p3', sheets: 'active' });
  assert.deepEqual(
    jobs.map((job) => job.options),
    [{ fileName: 'Budget 2026.xlsx' }, { dpi: 150, sheets: 'all', fileName: 'Budget 2026.xlsx' }, { sheets: 'active', fileName: 'Budget 2026.xlsx' }],
  );
  // The answers come when the jobs end, each with its id; other messages do not wait.
  const before = f.sent.length;
  await bridge.receive({ type: 'undo', id: 'u' });
  jobs[0].resolve();
  jobs[1].reject(Object.assign(new Error('There is nothing to print'), { code: 'empty' }));
  jobs[2].reject(Object.assign(new Error('The window has no print()'), { code: 'unavailable' }));
  await settle();
  assert.deepEqual(f.sent.slice(before), [
    { type: 'done', id: 'u', command: 'undo' },
    { type: 'done', id: 'p1', command: 'print' },
    { type: 'error', id: 'p2', code: 'print-failed', message: 'There is nothing to print' },
    { type: 'error', id: 'p3', code: 'print-unavailable', message: 'The window has no print()' },
  ]);
  // A file without a name: no file name is handed on. A print() that throws is a failure like any other.
  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'print' });
  assert.deepEqual(jobs[3].options, {});
  f.host.createView = (container, workbook, options) => {
    const view = make.call(f.host, container, workbook, options);
    view.print = (() => {
      throw new RangeError('The print resolution must be a positive number of dots per inch');
    }) as unknown as typeof view.print;
    return view;
  };
  await bridge.receive({ type: 'load', data: DATA });
  await bridge.receive({ type: 'print', id: 'p4' });
  await settle();
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'error', id: 'p4', code: 'print-failed', message: 'The print resolution must be a positive number of dots per inch' });
});

test('bridge: a merged "changed" does not name a sheet that was deleted while it waited', async () => {
  const f = fake();
  // One workbook object, so that the test can take a sheet out of it as the view would.
  const book = { sheets: ['A', 'B', 'C'].map((name) => ({ name, state: 'visible', kind: 'worksheet' })) };
  f.host.open = async () => ({ format: 'xlsx', workbook: async () => book, save: async () => new Uint8Array([1]) }) as unknown as OfficeFile;
  const bridge = createBridge(f.host);
  await bridge.receive(JSON.stringify({ type: 'load', id: 'l', data: DATA }));
  f.channel = false;
  const options = f.views[0];
  options.onChange!({ sheets: [2] });
  options.onChange!({ sheets: [1] });
  // The last sheet is deleted: the view reports the sheets there are now.
  book.sheets.pop();
  options.onChange!({ sheets: [0, 1] });
  f.channel = true;
  f.runRetries();
  assert.deepEqual(f.sent.filter((m) => m.type === 'changed'), [{ type: 'changed', sheets: [0, 1], canUndo: false, canRedo: false }]);
  bridge.destroy();
});

// ------------------------------------------- the tracked changes of a document

/** Makes the views of an editable document know about tracked changes: how many there are, and whether tracking is on. */
function withRevisions(f: Fake, count: number, format = 'docx'): { d: EditableDocument; state: { count: number; tracking: boolean; refuse: boolean } } {
  const d = withEditableDocument(f, format);
  const state = { count, tracking: false, refuse: false };
  for (const name of ['createDocumentView', 'createDocumentEditor'] as const) {
    const inner = f.host[name] as NonNullable<BridgeHost['createDocumentView']>;
    f.host[name] = ((container: HTMLElement, doc: never, options: never) => {
      const view = inner(container, doc, options) as unknown as Record<string, unknown>;
      view.revisions = () => new Array(state.count).fill({ kind: 'insert' });
      const resolve = (call: string) => () => {
        d.calls.push(call);
        if (state.refuse) return false;
        state.count = 0;
        d.views[d.views.length - 1].options.onChange!();
        return true;
      };
      view.acceptAllRevisions = resolve('acceptAll');
      view.rejectAllRevisions = resolve('rejectAll');
      view.setTrackChanges = (on: boolean) => {
        d.calls.push(`track ${on}`);
        if (state.refuse) return false;
        state.tracking = on;
        d.views[d.views.length - 1].options.onChange!();
        return true;
      };
      Object.defineProperty(view, 'trackingChanges', { get: () => state.tracking });
      return view as never;
    }) as never;
  }
  return { d, state };
}

test('bridge: `loaded` says how many tracked changes a document has, when it has any', async () => {
  const f = fake();
  withRevisions(f, 3);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', id: 'l', data: DATA });
  assert.deepEqual(f.sent[1], { type: 'loaded', id: 'l', format: 'docx', kind: 'document', sheets: [], activeSheet: 0, revisions: 3 });

  const g = fake();
  withRevisions(g, 0);
  const other = createBridge(g.host);
  await other.receive({ type: 'load', id: 'l', data: DATA });
  assert.deepEqual(g.sent[1], { type: 'loaded', id: 'l', format: 'docx', kind: 'document', sheets: [], activeSheet: 0 });
});

test('bridge: acceptAllRevisions and rejectAllRevisions reach the document; `revisions` says the new count before `changed`, `done` what holds now', async () => {
  for (const [command, call] of [
    ['acceptAllRevisions', 'acceptAll'],
    ['rejectAllRevisions', 'rejectAll'],
  ]) {
    const f = fake();
    const { d, state } = withRevisions(f, 3);
    const bridge = createBridge(f.host);
    await bridge.receive({ type: 'load', data: DATA });
    const from = f.sent.length;
    await bridge.receive({ type: command, id: 'a' });
    assert.ok(d.calls.indexOf(call) >= 0, command);
    assert.deepEqual(f.sent.slice(from), [
      { type: 'revisions', count: 0, trackChanges: false },
      { type: 'changed', sheets: [], canUndo: false, canRedo: false },
      { type: 'done', id: 'a', command, revisions: 0, trackChanges: false },
    ]);
    // A change of the document that leaves the count as it is says nothing about the changes.
    const next = f.sent.length;
    d.views[0].options.onChange!();
    assert.deepEqual(f.sent.slice(next), [{ type: 'changed', sheets: [], canUndo: false, canRedo: false }]);
    // An edit that makes a tracked change does.
    state.count = 1;
    d.views[0].options.onChange!();
    assert.deepEqual(f.sent.slice(next + 1), [
      { type: 'revisions', count: 1, trackChanges: false },
      { type: 'changed', sheets: [], canUndo: false, canRedo: false },
    ]);
    // A view that refuses (a protection): the answer says that nothing changed.
    state.refuse = true;
    await bridge.receive({ type: command, id: 'b' });
    assert.deepEqual(f.sent[f.sent.length - 1], { type: 'done', id: 'b', command, revisions: 1, trackChanges: false });
  }
});

test('bridge: setTrackChanges turns tracking on and off and needs `on`', async () => {
  const f = fake();
  const { d, state } = withRevisions(f, 0);
  const bridge = createBridge(f.host);
  await bridge.receive({ type: 'load', data: DATA });
  const from = f.sent.length;
  await bridge.receive({ type: 'setTrackChanges', on: true, id: 't' });
  assert.ok(d.calls.indexOf('track true') >= 0);
  assert.deepEqual(f.sent.slice(from), [
    { type: 'revisions', count: 0, trackChanges: true },
    { type: 'changed', sheets: [], canUndo: false, canRedo: false },
    { type: 'done', id: 't', command: 'setTrackChanges', revisions: 0, trackChanges: true },
  ]);
  // Where it cannot be turned off, the answer says that it is still on.
  state.refuse = true;
  await bridge.receive({ type: 'setTrackChanges', on: false, id: 'u' });
  assert.deepEqual(f.sent[f.sent.length - 1], { type: 'done', id: 'u', command: 'setTrackChanges', revisions: 0, trackChanges: true });
  await bridge.receive({ type: 'setTrackChanges', id: 'v' });
  assert.equal((f.sent[f.sent.length - 1] as { code?: string }).code, 'bad-message');
  await bridge.receive({ type: 'setTrackChanges', on: 'yes', id: 'w' });
  assert.equal((f.sent[f.sent.length - 1] as { code?: string }).code, 'bad-message');
});

test('bridge: the commands of the tracked changes without a file, with a workbook, and with a document that stays read-only', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  for (const type of ['acceptAllRevisions', 'rejectAllRevisions', 'setTrackChanges']) {
    await bridge.receive({ type, on: true, id: 'n' });
    assert.deepEqual([(f.sent[f.sent.length - 1] as { code?: string }).code, (f.sent[f.sent.length - 1] as { id?: string }).id], ['not-loaded', 'n'], type);
  }
  await bridge.receive({ type: 'load', data: DATA });
  for (const type of ['acceptAllRevisions', 'rejectAllRevisions', 'setTrackChanges']) {
    await bridge.receive({ type, on: true, id: 'w' });
    assert.equal((f.sent[f.sent.length - 1] as { code?: string }).code, 'no-document', type);
  }
  // An RTF file that was not converted cannot be edited: neither can its changes be resolved.
  const g = fake();
  const { d } = withRevisions(g, 2, 'rtf');
  const other = createBridge(g.host);
  await other.receive({ type: 'load', data: DATA });
  assert.equal((g.sent[1] as { revisions?: number }).revisions, 2);
  for (const type of ['acceptAllRevisions', 'rejectAllRevisions', 'setTrackChanges']) {
    await other.receive({ type, on: true, id: 'r' });
    assert.equal((g.sent[g.sent.length - 1] as { code?: string }).code, 'read-only', type);
  }
  assert.deepEqual(d.calls.filter((call) => /acceptAll|rejectAll|track/.test(call)), []);
  // So is a document that was opened read-only.
  const h = fake();
  const opened = withRevisions(h, 2);
  const third = createBridge(h.host);
  await third.receive({ type: 'load', data: DATA, readOnly: true });
  await third.receive({ type: 'acceptAllRevisions', id: 'r' });
  assert.equal((h.sent[h.sent.length - 1] as { code?: string }).code, 'read-only');
  // Editing turned on, it works.
  await third.receive({ type: 'setReadOnly', readOnly: false });
  await third.receive({ type: 'acceptAllRevisions', id: 's' });
  assert.deepEqual(h.sent[h.sent.length - 1], { type: 'done', id: 's', command: 'acceptAllRevisions', revisions: 0, trackChanges: false });
  assert.ok(opened.d.calls.indexOf('acceptAll') >= 0);
});

test('bridge: an application that answers ready from inside its postMessage is served: one ready, the file loaded, nothing thrown', async () => {
  const f = fake();
  const bridge = createBridge(f.host);
  const post = f.host.post;
  const thrown: unknown[] = [];
  let answered = false;
  f.host.post = (json) => {
    const delivered = post(json);
    // As a page does that owns the frame of the editor: the answer to `ready` is given in the same call.
    if (delivered && !answered && JSON.parse(json).type === 'ready') {
      answered = true;
      try {
        void bridge.receive({ type: 'hello' });
        void bridge.receive({ type: 'load', id: 'file', data: DATA, readOnly: true });
      } catch (error) {
        thrown.push(error);
      }
    }
    return delivered;
  };
  let failed: unknown;
  try {
    bridge.start();
  } catch (error) {
    failed = error;
  }
  for (let i = 0; i < 5; i++) await settle();
  f.runRetries();
  for (let i = 0; i < 5; i++) await settle();
  assert.equal(failed, undefined, 'start() does not throw');
  assert.deepEqual(thrown, [], 'receive() inside postMessage does not throw');
  // The hello asked for ready again; the load was carried out and answered once, and nothing is sent twice by the pass that was under way.
  assert.deepEqual(types(f).filter((type) => type === 'loaded'), ['loaded']);
  assert.deepEqual(types(f).slice(0, 1), ['ready']);
  assert.ok(types(f).filter((type) => type === 'ready').length <= 2, types(f).join());
  assert.equal(f.opened.length, 1);
  const loaded = f.sent.find((message) => message.type === 'loaded');
  assert.equal(loaded.id, 'file');
});
