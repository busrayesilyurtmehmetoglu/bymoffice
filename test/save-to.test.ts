import assert from 'node:assert/strict';
import { test } from 'node:test';
import { SAVE_MEDIA_TYPES, SAVE_TIMEOUT_MS, SaveController, SaveError, nameFromUrl, resolveSaveTarget, strongETag, type SaveHost, type SaveInfo, type SaveState, type SaveTarget } from '../src/embed/save-to';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

const PAGE = 'https://app.example.com/editor/index.html?x=1';
const char = String.fromCharCode;

interface Sent {
  url: string;
  init: RequestInit;
  /** Answers the request. */
  answer(status: number, headers?: { [name: string]: string }): void;
  /** Makes the request fail as a connection that broke. */
  breakOff(): void;
}

interface Rig {
  controller: SaveController;
  host: SaveHost;
  sent: Sent[];
  states: SaveState[];
  notices: [kind: string, message: string][];
  errors: unknown[];
  saved: Uint8Array[];
  timers: { callback: () => void; ms: number; off: boolean }[];
  /** How many bodies of answers were let go of without being read. */
  dropped: number;
  /** What the file is written as; the test changes it to tell one save from the next. */
  content: number[];
  commits: number;
  clock: number;
}

/** Lets promise callbacks that are already due run. */
const settle = async () => {
  for (let i = 0; i < 30; i++) await Promise.resolve();
};

function rig(target: SaveTarget, more: Partial<SaveHost> = {}, info: SaveInfo = { format: 'docx' }): Rig {
  const state: Rig = { controller: undefined as unknown as SaveController, host: undefined as unknown as SaveHost, sent: [], states: [], notices: [], errors: [], saved: [], timers: [], dropped: 0, content: [1, 2, 3], commits: 0, clock: 1000 };
  const fakeFetch = (url: unknown, init?: RequestInit) =>
    new Promise<Response>((resolve, reject) => {
      const signal = init && init.signal;
      // As a browser does: an aborted request rejects.
      if (signal) signal.addEventListener('abort', () => reject(new Error('aborted by the signal')));
      state.sent.push({
        url: String(url),
        init: init || {},
        answer(status, headers = {}) {
          const lower: { [name: string]: string } = {};
          for (const name of Object.keys(headers)) lower[name.toLowerCase()] = headers[name];
          resolve({
            status,
            ok: status >= 200 && status < 300,
            headers: { get: (name: string) => (name.toLowerCase() in lower ? lower[name.toLowerCase()] : null) },
            body: { cancel: async () => void state.dropped++ },
            // Reading the body is what must not happen.
            arrayBuffer: () => assert.fail('the body of the answer was read'),
            text: () => assert.fail('the body of the answer was read'),
          } as unknown as Response);
        },
        breakOff: () => reject(new TypeError('Failed to fetch <secret of the browser>')),
      });
    });
  state.host = {
    target,
    writable: () => true,
    commit: () => {
      state.commits++;
      return true;
    },
    bytes: async () => new Uint8Array(state.content),
    info: () => info,
    strings: () => STRINGS_EN,
    notice: (message, kind) => state.notices.push([kind, message]),
    onState: (value) => state.states.push(value),
    onSaved: (bytes) => state.saved.push(bytes),
    onError: (error) => state.errors.push(error),
    fetch: fakeFetch as typeof fetch,
    now: () => state.clock,
    timer: (callback, ms) => {
      const entry = { callback, ms, off: false };
      state.timers.push(entry);
      return () => void (entry.off = true);
    },
    ...more,
  };
  state.controller = new SaveController(state.host);
  return state;
}

const request = (url = 'https://app.example.com/files/7', more: Partial<Extract<SaveTarget, { kind: 'request' }>> = {}): SaveTarget => ({ kind: 'request', url, method: 'PUT', headers: [], credentials: 'same-origin', sameAsFile: true, timeoutMs: SAVE_TIMEOUT_MS, allowConverted: true, ...more });
const headersOf = (sent: Sent) => sent.init.headers as { [name: string]: string };

// ------------------------------------------------------------ the option

test('saveTo: true saves to the address the file was fetched from, with PUT and the credentials the file is fetched with', () => {
  assert.deepEqual(resolveSaveTarget(true, '/files/7/download?v=2#top', PAGE), {
    kind: 'request',
    url: 'https://app.example.com/files/7/download?v=2',
    method: 'PUT',
    headers: [],
    credentials: 'same-origin',
    sameAsFile: true,
    timeoutMs: SAVE_TIMEOUT_MS,
    allowConverted: false,
    onConflict: 'ask',
  });
  // A file that was given as bytes has no address to go back to.
  assert.throws(() => resolveSaveTarget(true, undefined, PAGE), TypeError);
  // Nor has one that came from a blob: or data: address.
  assert.throws(() => resolveSaveTarget(true, 'blob:https://app.example.com/1-2', PAGE), TypeError);
  assert.throws(() => resolveSaveTarget(true, 'data:application/octet-stream;base64,AAAA', PAGE), TypeError);
});

test('saveTo: an address is resolved against the page; one that is where the file came from is known as such', () => {
  const other = resolveSaveTarget('../files/8', '/files/7', PAGE);
  assert.deepEqual([other.kind, other.kind === 'request' && other.url, other.kind === 'request' && other.sameAsFile], ['request', 'https://app.example.com/files/8', false]);
  const same = resolveSaveTarget('https://APP.example.com:443/files/7', '/files/7', PAGE);
  assert.deepEqual([same.kind === 'request' && same.url, same.kind === 'request' && same.sameAsFile], ['https://app.example.com/files/7', true]);
  // Another origin is allowed: it is the page that says where its files go.
  const far = resolveSaveTarget('https://store.example.net/put/7', undefined, PAGE);
  assert.equal(far.kind === 'request' && far.url, 'https://store.example.net/put/7');
  assert.equal(far.kind === 'request' && far.credentials, 'same-origin', 'nothing of the user goes to another origin unless the page asks for it');
  // Without a page to resolve against, only a whole address will do.
  assert.throws(() => resolveSaveTarget('/files/8', undefined, undefined), TypeError);
  assert.equal((resolveSaveTarget('http://localhost:8080/f', undefined, undefined) as { url: string }).url, 'http://localhost:8080/f');
});

test('saveTo: a request takes its url from the file when it has none, and its method, headers and credentials as given', () => {
  assert.deepEqual(resolveSaveTarget({ method: 'POST', headers: { 'X-CSRF-Token': 'abc', Authorization: 'Bearer t' }, credentials: 'include' }, 'https://app.example.com/files/7', PAGE), {
    kind: 'request',
    url: 'https://app.example.com/files/7',
    method: 'POST',
    headers: [
      ['X-CSRF-Token', 'abc'],
      ['Authorization', 'Bearer t'],
    ],
    credentials: 'include',
    sameAsFile: true,
    timeoutMs: SAVE_TIMEOUT_MS,
    allowConverted: false,
    onConflict: 'ask',
  });
  const withUrl = resolveSaveTarget({ url: '/store', credentials: 'omit' }, undefined, PAGE);
  assert.deepEqual([withUrl.kind === 'request' && withUrl.url, withUrl.kind === 'request' && withUrl.method, withUrl.kind === 'request' && withUrl.credentials], ['https://app.example.com/store', 'PUT', 'omit']);
  assert.throws(() => resolveSaveTarget({}, undefined, PAGE), TypeError, 'a request without url for a file given as bytes');
  for (const bad of [{ method: 'DELETE' }, { method: 'put' }, { credentials: 'yes' }, { url: 7 }, { headers: 'X: y' }, { headers: ['X', 'y'] }, { headers: { 'Bad Name': 'x' } }, { headers: { '': 'x' } }, { headers: { X: 7 } }, { headers: { X: `a${char(13, 10)}Y: b` } }, { headers: { X: `a${char(0)}` } }]) {
    assert.throws(() => resolveSaveTarget(bad, '/files/7', PAGE), TypeError, JSON.stringify(bad));
  }
});

test('saveTo: a function is the application saving by itself; anything else is refused', () => {
  const own = async () => undefined;
  assert.deepEqual(resolveSaveTarget(own, undefined, undefined), { kind: 'function', save: own });
  for (const bad of [7, null, undefined, false, [], ['/files/7']]) assert.throws(() => resolveSaveTarget(bad, '/files/7', PAGE), TypeError, JSON.stringify(bad));
});

test('saveTo: only http(s) addresses are saved to', () => {
  const bad = [
    '',
    'javascript:alert(1)',
    'JaVaScript:alert(1)',
    ' javascript:alert(1)',
    `java${char(9)}script:alert(1)`,
    `${char(1)}javascript:x`,
    'data:text/plain,x',
    'file:///etc/passwd',
    'blob:https://app.example.com/1-2',
    'ftp://app.example.com/a',
    'vbscript:x',
    'content://x/y',
    'https://user:pw@app.example.com/files/7',
    'https://user@app.example.com/files/7',
    'https://',
  ];
  for (const url of bad) {
    assert.throws(() => resolveSaveTarget(url, '/files/7', PAGE), TypeError, JSON.stringify(url));
    assert.throws(() => resolveSaveTarget({ url }, '/files/7', PAGE), TypeError, JSON.stringify(url));
  }
  // A page that is not served over http(s) has no relative address to save to.
  for (const base of ['file:///android_asset/page.html', 'about:blank', 'data:text/html,x']) assert.throws(() => resolveSaveTarget('files/7', undefined, base), TypeError, base);
});

test('strongETag: only an ETag that If-Match can be sent with', () => {
  assert.equal(strongETag('"abc-1"'), '"abc-1"');
  assert.equal(strongETag('""'), '""');
  for (const not of ['W/"abc"', 'abc', '"a"b"', `"a${char(10)}"`, '', null, undefined]) assert.equal(strongETag(not), undefined, JSON.stringify(not));
});

test('nameFromUrl: the last part of the path, decoded; not of the query', () => {
  assert.equal(nameFromUrl('/files/Rapor%202024.docx?v=2', PAGE), 'Rapor 2024.docx');
  assert.equal(nameFromUrl('https://app.example.com/files/7/', PAGE), undefined);
  assert.equal(nameFromUrl('blob:https://app.example.com/1-2', PAGE), undefined);
  assert.equal(nameFromUrl('/files/%E0%A4%A', PAGE), undefined, 'a name that cannot be decoded is no name');
  assert.equal(nameFromUrl('a.xlsx', undefined), undefined);
});

// ----------------------------------------------------------- the request

test('request: the bytes go out with PUT, the media type of the format and X-Bym-Format; 2xx means saved', async () => {
  const r = rig(request());
  assert.deepEqual(r.controller.state, { dirty: false, saving: false });
  r.controller.changed();
  assert.deepEqual(r.states, [{ dirty: true, saving: false }]);
  const done = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 1);
  assert.equal(r.sent[0].url, 'https://app.example.com/files/7');
  assert.equal(r.sent[0].init.method, 'PUT');
  assert.equal(r.sent[0].init.credentials, 'same-origin');
  assert.deepEqual(Array.from(r.sent[0].init.body as Uint8Array), [1, 2, 3]);
  assert.deepEqual(headersOf(r.sent[0]), { 'Content-Type': SAVE_MEDIA_TYPES.docx, 'X-Bym-Format': 'docx' });
  assert.deepEqual(r.controller.state, { dirty: true, saving: true });
  r.sent[0].answer(204);
  assert.equal(await done, true);
  assert.deepEqual(r.controller.state, { dirty: false, saving: false, lastSavedAt: 1000 });
  assert.deepEqual(r.states.slice(1), [
    { dirty: true, saving: true },
    { dirty: false, saving: false, lastSavedAt: 1000 },
  ]);
  assert.deepEqual(r.notices, [['info', 'Saved']]);
  assert.deepEqual(r.errors, []);
  assert.equal(r.dropped, 1, 'the body of the answer is let go of, not read');
  assert.deepEqual(r.timers.map((entry) => [entry.ms, entry.off]), [[SAVE_TIMEOUT_MS, true]], 'the time limit is the one of a fetch, and is called off');
});

test('request: an xlsx, POST, credentials and headers of the page; a header of the page takes the place of one of the engine', async () => {
  const target = request('https://store.example.net/put', { method: 'POST', credentials: 'include', sameAsFile: false, headers: [['X-CSRF-Token', 'abc'], ['content-type', 'application/octet-stream']] });
  const r = rig(target, { etag: '"v1"' }, { format: 'xlsx' });
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  assert.equal(r.sent[0].init.method, 'POST');
  assert.equal(r.sent[0].init.credentials, 'include');
  // No If-Match: the ETag is of the address the file came from, and this is another.
  assert.deepEqual(headersOf(r.sent[0]), { 'X-Bym-Format': 'xlsx', 'X-CSRF-Token': 'abc', 'content-type': 'application/octet-stream' });
  r.sent[0].answer(200);
  assert.equal(await done, true);
});

test('request: a file that was converted says in a header what it was', async () => {
  for (const [format, from] of [['docx', 'doc'], ['docx', 'rtf'], ['xlsx', 'xls']] as const) {
    const r = rig(request(), {}, { format, convertedFrom: from, name: `old.${format}` });
    r.controller.changed();
    const done = r.controller.saveNow();
    await settle();
    assert.deepEqual(headersOf(r.sent[0]), { 'Content-Type': SAVE_MEDIA_TYPES[format], 'X-Bym-Format': format, 'X-Bym-Converted-From': from });
    r.sent[0].answer(201);
    assert.equal(await done, true, from);
  }
});

test('ETag: the one the file came with is sent as If-Match, and the one of the answer the next time', async () => {
  const r = rig(request(), { etag: '"v1"' });
  r.controller.changed();
  let done = r.controller.saveNow();
  await settle();
  assert.equal(headersOf(r.sent[0])['If-Match'], '"v1"');
  r.sent[0].answer(200, { ETag: '"v2"' });
  assert.equal(await done, true);
  assert.equal(r.controller.currentETag, '"v2"');

  r.controller.changed();
  done = r.controller.saveNow();
  await settle();
  assert.equal(headersOf(r.sent[1])['If-Match'], '"v2"');
  // An answer without one: the old ETag names a file that is gone, and is not sent again.
  r.sent[1].answer(204);
  assert.equal(await done, true);
  assert.equal(r.controller.currentETag, undefined);

  r.controller.changed();
  done = r.controller.saveNow();
  await settle();
  assert.equal('If-Match' in headersOf(r.sent[2]), false);
  r.sent[2].answer(204, { etag: 'W/"weak"' });
  assert.equal(await done, true);
  assert.equal(r.controller.currentETag, undefined, 'a weak ETag never matches in If-Match');
});

test('ETag: a weak one the file came with is not sent; the page can send its own If-Match', async () => {
  const weak = rig(request(), { etag: 'W/"v1"' });
  weak.controller.changed();
  void weak.controller.saveNow();
  await settle();
  assert.equal('If-Match' in headersOf(weak.sent[0]), false);

  const own = rig(request(undefined, { headers: [['if-match', '*']] }), { etag: '"v1"' });
  own.controller.changed();
  void own.controller.saveNow();
  await settle();
  assert.deepEqual(headersOf(own.sent[0]), { 'Content-Type': SAVE_MEDIA_TYPES.docx, 'X-Bym-Format': 'docx', 'if-match': '*' });
});

test('412: the file changed on the server; nothing is overwritten, the ETag stays, and the file is still unsaved', async () => {
  const r = rig(request(), { etag: '"v1"' });
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  r.sent[0].answer(412, { ETag: '"v9"' });
  assert.equal(await done, false);
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveConflict });
  assert.deepEqual(r.notices, [['error', STRINGS_EN.saveConflict]]);
  assert.equal(r.errors.length, 1);
  const error = r.errors[0] as SaveError;
  assert.deepEqual([error instanceof SaveError, error.name, error.code, error.status], [true, 'SaveError', 'conflict', 412]);
  assert.equal(r.controller.currentETag, '"v1"');
  assert.deepEqual(r.saved, [], 'onSave is not called for a save that failed');
  assert.equal(r.dropped, 1);
});

test('401, 403 and 423: not allowed, and locked; 500 and the rest: could not be saved, with the status', async () => {
  const cases: [number, string, string][] = [
    [401, 'denied', STRINGS_EN.saveDenied],
    [403, 'denied', STRINGS_EN.saveDenied],
    [423, 'locked', STRINGS_EN.saveLocked],
    [500, 'http', 'The file could not be saved (HTTP 500)'],
    [409, 'http', 'The file could not be saved (HTTP 409)'],
    [404, 'http', 'The file could not be saved (HTTP 404)'],
    [199, 'http', 'The file could not be saved (HTTP 199)'],
  ];
  for (const [status, code, text] of cases) {
    const r = rig(request());
    r.controller.changed();
    const done = r.controller.saveNow();
    await settle();
    r.sent[0].answer(status);
    assert.equal(await done, false, String(status));
    assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: text }, String(status));
    assert.deepEqual(r.notices, [['error', text]], String(status));
    assert.deepEqual([(r.errors[0] as SaveError).code, (r.errors[0] as SaveError).status], [code, status]);
  }
});

test('the words of a failure are those of the language of the view', async () => {
  const r = rig(request(), { strings: () => STRINGS_TR });
  r.controller.changed();
  let done = r.controller.saveNow();
  await settle();
  r.sent[0].answer(412);
  await done;
  done = r.controller.saveNow();
  await settle();
  r.sent[1].answer(503);
  await done;
  done = r.controller.saveNow();
  await settle();
  r.sent[2].answer(204);
  await done;
  assert.deepEqual(r.notices, [
    ['error', STRINGS_TR.saveConflict],
    ['error', 'Dosya kaydedilemedi (HTTP 503)'],
    ['info', 'Kaydedildi'],
  ]);
});

test('a request that does not get through: said without what the browser says of it', async () => {
  const r = rig(request());
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  r.sent[0].breakOff();
  assert.equal(await done, false);
  assert.deepEqual(r.notices, [['error', STRINGS_EN.saveNoConnection]]);
  const error = r.errors[0] as SaveError;
  assert.equal(error.code, 'network');
  assert.doesNotMatch(error.message + r.notices[0][1], /secret/);
  // A place without fetch is the same failure.
  const none = rig(request(), { fetch: undefined });
  const globalFetch = globalThis.fetch;
  try {
    (globalThis as { fetch?: unknown }).fetch = undefined;
    none.controller.changed();
    assert.equal(await none.controller.saveNow(), false);
  } finally {
    globalThis.fetch = globalFetch;
  }
  assert.equal((none.errors[0] as SaveError).code, 'network');
});

test('no answer in time: the request is aborted and the save has failed; a late answer changes nothing', async () => {
  const r = rig(request(), { timeoutMs: 5000 });
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  assert.equal(r.timers[0].ms, 5000);
  const signal = r.sent[0].init.signal as AbortSignal;
  assert.equal(signal.aborted, false);
  r.timers[0].callback();
  assert.equal(await done, false);
  assert.equal(signal.aborted, true);
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveTimedOut });
  assert.equal((r.errors[0] as SaveError).code, 'timeout');
  r.sent[0].answer(204);
  await settle();
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveTimedOut });
  assert.equal(r.notices.length, 1);
});

// ------------------------------------------------------- what is saved

test('nothing to save: no request, and the answer is true', async () => {
  const r = rig(request());
  assert.equal(await r.controller.saveNow(), true);
  assert.equal(r.sent.length, 0);
  assert.deepEqual([r.states, r.notices, r.saved], [[], [], []]);
  assert.equal(r.commits, 1, 'what is being typed is applied before it is known whether there is anything to save');
});

test('what is typed while the file is on its way is not called saved, and the next save holds it', async () => {
  const r = rig(request());
  r.controller.changed();
  const first = r.controller.saveNow();
  await settle();
  r.controller.changed();
  r.content = [4, 5];
  assert.deepEqual(r.controller.state, { dirty: true, saving: true });
  r.sent[0].answer(204);
  assert.equal(await first, true);
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, lastSavedAt: 1000 });
  assert.deepEqual(r.notices, [], '"Saved" is not said while there is something to save');
  assert.deepEqual(r.saved.map((bytes) => Array.from(bytes)), [[1, 2, 3]]);
  r.clock = 2000;
  const second = r.controller.saveNow();
  await settle();
  assert.deepEqual(Array.from(r.sent[1].init.body as Uint8Array), [4, 5]);
  r.sent[1].answer(204);
  assert.equal(await second, true);
  assert.deepEqual(r.controller.state, { dirty: false, saving: false, lastSavedAt: 2000 });
  assert.deepEqual(r.notices, [['info', 'Saved']]);
});

test('two saves one after the other send one request; with a change between them the second waits and sends once more', async () => {
  const r = rig(request());
  r.controller.changed();
  const first = r.controller.saveNow();
  const again = r.controller.saveNow();
  r.controller.run();
  await settle();
  assert.equal(r.sent.length, 1, 'a second press sends nothing');
  r.sent[0].answer(204);
  assert.deepEqual([await first, await again], [true, true]);
  assert.equal(r.sent.length, 1);

  r.controller.changed();
  const third = r.controller.saveNow();
  await settle();
  r.controller.changed();
  r.content = [9];
  const fourth = r.controller.saveNow();
  const fifth = r.controller.saveNow();
  await settle();
  assert.equal(r.sent.length, 2, 'the save that waits does not start before the one under way has its answer');
  r.sent[1].answer(204);
  assert.equal(await third, true);
  await settle();
  assert.equal(r.sent.length, 3, 'two presses that waited are one save');
  assert.deepEqual(Array.from(r.sent[2].init.body as Uint8Array), [9]);
  r.sent[2].answer(204);
  assert.deepEqual([await fourth, await fifth], [true, true]);
  assert.deepEqual(r.controller.state.dirty, false);
});

test('a second press while a save that will fail is under way does not send the same file again', async () => {
  const r = rig(request());
  r.controller.changed();
  const first = r.controller.saveNow();
  const again = r.controller.saveNow();
  await settle();
  r.sent[0].answer(500);
  assert.deepEqual([await first, await again], [false, false]);
  await settle();
  assert.equal(r.sent.length, 1);
  // The user decides to try again.
  const retry = r.controller.saveNow();
  await settle();
  assert.deepEqual(r.controller.state, { dirty: true, saving: true }, 'the error goes when a save starts');
  r.sent[1].answer(204);
  assert.equal(await retry, true);
});

test('a file that cannot be written (read-only, a format that was not converted) is never sent', async () => {
  const r = rig(request(), { writable: () => false });
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), false);
  assert.equal(r.sent.length, 0);
  assert.equal(r.commits, 0);
  assert.deepEqual(r.notices, []);
});

test('what is being typed and cannot be applied keeps the file from being saved', async () => {
  let open = true;
  const r = rig(request(), { commit: () => !open });
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), false);
  assert.equal(r.sent.length, 0);
  assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: STRINGS_EN.savePending });
  assert.equal((r.errors[0] as SaveError).code, 'pending');
  assert.deepEqual(r.notices, [], 'the view said why where it happened');
  // A commit that throws is the same.
  const thrower = rig(request(), {
    commit: () => {
      throw new Error('cell');
    },
  });
  thrower.controller.changed();
  assert.equal(await thrower.controller.saveNow(), false);
  assert.equal(thrower.sent.length, 0);
  open = false;
  const done = r.controller.saveNow();
  await settle();
  r.sent[0].answer(204);
  assert.equal(await done, true);
});

test('a file that cannot be written out: could not be saved, and nothing is sent', async () => {
  const r = rig(request(), { bytes: async () => Promise.reject(new Error('Sheet "Maaşlar" cannot be written')) });
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), false);
  assert.equal(r.sent.length, 0);
  assert.deepEqual(r.notices, [['error', 'The file could not be saved']], 'what the writer said (it may name what is in the file) is not shown');
  const error = r.errors[0] as SaveError;
  assert.equal(error.code, 'write');
  assert.equal((error.cause as Error).message, 'Sheet "Maaşlar" cannot be written');
});

test('onSave is called with the bytes of every save that succeeded, after the state says so; its failure does not undo the save', async () => {
  const order: string[] = [];
  const r = rig(request(), {
    onState: (state) => order.push(`state ${state.dirty} ${state.saving}`),
    onSaved: (bytes) => {
      order.push(`onSave ${Array.from(bytes).join()}`);
      throw new Error('the page failed');
    },
  });
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  r.sent[0].answer(204);
  assert.equal(await done, true);
  assert.deepEqual(order, ['state true false', 'state true true', 'state false false', 'onSave 1,2,3']);
  assert.equal(r.controller.dirty, false);
  assert.equal((r.errors[0] as Error).message, 'the page failed');
});

// ------------------------------------------------- the application saves

test('a function: it is given the bytes and what they are; resolving, or returning nothing, means saved', async () => {
  const taken: [number[], SaveInfo][] = [];
  const info: SaveInfo = { format: 'xlsx', convertedFrom: 'xls', name: 'liste.xlsx' };
  const later = rig({ kind: 'function', save: async (bytes, what) => void taken.push([Array.from(bytes), what]) }, {}, info);
  later.controller.changed();
  assert.equal(await later.controller.saveNow(), true);
  const plain = rig({ kind: 'function', save: (bytes, what) => void taken.push([Array.from(bytes), what]) }, {}, info);
  plain.controller.changed();
  assert.equal(await plain.controller.saveNow(), true);
  assert.deepEqual(taken, [
    [[1, 2, 3], info],
    [[1, 2, 3], info],
  ]);
  assert.deepEqual(plain.controller.state, { dirty: false, saving: false, lastSavedAt: 1000 });
  assert.deepEqual(plain.notices, [['info', 'Saved']]);
  assert.equal(later.sent.length + plain.sent.length, 0, 'nothing is fetched');
  assert.deepEqual(plain.timers.map((entry) => [entry.ms, entry.off]), [[600000, true]], 'a long time limit, called off when the function answers');
});

test('a function that rejects or throws: the save has failed, with one line of what it said', async () => {
  const long = `Disk ${'x'.repeat(300)}`;
  const cases: [() => void | Promise<void>, string][] = [
    [() => Promise.reject(new Error('quota exceeded')), 'The file could not be saved (quota exceeded)'],
    [
      () => {
        throw new Error(`first line${char(10)}second${char(9)}line`);
      },
      'The file could not be saved (first line second line)',
    ],
    [() => Promise.reject('just text'), 'The file could not be saved (just text)'],
    [() => Promise.reject(undefined), 'The file could not be saved'],
    [() => Promise.reject({ html: '<b>x</b>' }), 'The file could not be saved'],
    [() => Promise.reject(new Error(long)), `The file could not be saved (${long.slice(0, 200)}…)`],
  ];
  for (const [save, text] of cases) {
    const r = rig({ kind: 'function', save });
    r.controller.changed();
    assert.equal(await r.controller.saveNow(), false, text);
    assert.deepEqual(r.controller.state, { dirty: true, saving: false, error: text });
    assert.deepEqual(r.notices, [['error', text]]);
    assert.equal((r.errors[0] as SaveError).code, 'refused');
    assert.deepEqual(r.saved, []);
  }
});

// ------------------------------------------------------------- the end

test('destroy: the request under way is aborted, the save answers false and nothing more is said', async () => {
  const r = rig(request());
  let heard = 0;
  const unwatch = r.controller.watch(() => heard++);
  r.controller.changed();
  assert.equal(heard, 1);
  const done = r.controller.saveNow();
  await settle();
  const before = [r.states.length, r.notices.length, r.errors.length, heard];
  r.controller.destroy();
  r.controller.destroy();
  assert.equal((r.sent[0].init.signal as AbortSignal).aborted, true);
  assert.equal(await done, false);
  r.sent[0].answer(204);
  r.controller.changed();
  await settle();
  assert.deepEqual([r.states.length, r.notices.length, r.errors.length, heard], before);
  assert.deepEqual(r.saved, []);
  assert.equal(await r.controller.saveNow(), false);
  assert.equal(r.sent.length, 1);
  assert.equal(r.timers[0].off, true, 'the time limit is called off');
  unwatch();
});

test('destroy while the application saves: its answer is not waited for', async () => {
  let finish: () => void = () => undefined;
  const r = rig({ kind: 'function', save: () => new Promise<void>((resolve) => (finish = resolve)) });
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  r.controller.destroy();
  finish();
  assert.equal(await done, false);
  assert.deepEqual(r.notices, []);
});

test('the Save button is told of every change of what it shows, and of nothing after its watch ended', async () => {
  const r = rig(request());
  const seen: string[] = [];
  const unwatch = r.controller.watch(() => seen.push(`${r.controller.dirty} ${r.controller.saving}`));
  r.controller.changed();
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  r.sent[0].answer(204);
  await done;
  assert.deepEqual(seen, ['true false', 'true true', 'false false']);
  unwatch();
  r.controller.changed();
  assert.equal(seen.length, 3);
});
