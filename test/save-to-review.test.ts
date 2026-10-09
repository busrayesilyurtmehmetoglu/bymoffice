import assert from 'node:assert/strict';
import { createServer, type Server } from 'node:http';
import { test } from 'node:test';
import * as saveTo from '../src/embed/save-to';
import { SaveController, SaveError, resolveSaveTarget, type SaveHost, type SaveInfo, type SaveTarget } from '../src/embed/save-to';
import { STRINGS_EN, STRINGS_TR } from '../src/ui/strings';

// What the review of saveTo found: saves that were called saved and were not
// (a redirect to a login page), a save that could never succeed again (the
// answer of the one before it was lost), headers that pass and cannot be sent,
// a converted file written over its original. Where it matters the request is
// made with the fetch of Node against a server of this test, so that what a
// faked fetch would hide shows.

const PAGE = 'https://app.example.com/editor/index.html';
const char = String.fromCharCode;
const settle = async () => {
  for (let i = 0; i < 30; i++) await Promise.resolve();
};

interface Made {
  controller: SaveController;
  notices: string[];
  errors: SaveError[];
  timers: { callback: () => void; ms: number; off: boolean }[];
}

function make(target: SaveTarget, more: Partial<SaveHost> = {}, info: SaveInfo = { format: 'xlsx' }): Made {
  const made: Made = { controller: undefined as unknown as SaveController, notices: [], errors: [], timers: [] };
  made.controller = new SaveController({
    target,
    writable: () => true,
    commit: () => true,
    bytes: async () => new Uint8Array([1, 2, 3]),
    info: () => info,
    strings: () => STRINGS_EN,
    notice: (message, kind) => void made.notices.push(`${kind}: ${message}`),
    onError: (error) => void made.errors.push(error as SaveError),
    ...more,
  });
  return made;
}

/** A server that redirects to a login page, as an application whose session ran out does, and one that keeps a file with an ETag. */
async function withServer(run: (origin: string, log: string[], state: { slow: number; version: number; stored: number }) => Promise<void>): Promise<void> {
  const log: string[] = [];
  const state = { slow: 0, version: 1, stored: 0 };
  const server: Server = createServer((req, res) => {
    const chunks: Buffer[] = [];
    req.on('data', (chunk) => chunks.push(chunk));
    req.on('end', () => {
      const body = Buffer.concat(chunks);
      log.push(`${req.method} ${req.url} ${body.length} ${req.headers['if-match'] ?? '-'} ${req.headers['x-csrf-token'] ?? '-'}`);
      if (req.url === '/login') {
        res.writeHead(200, { 'content-type': 'text/html' });
        return res.end('<form>login</form>');
      }
      const redirect = /^\/go(30[12378])$/.exec(req.url || '');
      if (redirect) {
        res.writeHead(Number(redirect[1]), { location: '/login' });
        return res.end();
      }
      if (req.url === '/file') {
        const match = req.headers['if-match'];
        if (match && match !== `"v${state.version}"`) {
          res.writeHead(412);
          return res.end();
        }
        state.stored = body.length;
        state.version++;
        const answer = () => {
          res.writeHead(204, { etag: `"v${state.version}"` });
          res.end();
        };
        if (state.slow > 0) setTimeout(answer, state.slow);
        else answer();
        return;
      }
      res.writeHead(404);
      res.end();
    });
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address() as { port: number };
  try {
    await run(`http://127.0.0.1:${address.port}`, log, state);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
}

// ------------------------------------------------------------- redirects

test('review 1: a save that the server redirects (a session that ran out) is not saved, and the redirect is not followed', async () => {
  await withServer(async (origin, log) => {
    const cases: [unknown, string][] = [
      [{ url: '/go302', method: 'POST', headers: { 'X-CSRF-Token': 't' } }, 'POST /go302 3 - t'],
      ['/go303', 'PUT /go303 3 - -'],
      ['/go302', 'PUT /go302 3 - -'],
      ['/go307', 'PUT /go307 3 - -'],
      ['/go308', 'PUT /go308 3 - -'],
    ];
    for (const [target, line] of cases) {
      log.length = 0;
      const made = make(resolveSaveTarget(target, undefined, `${origin}/page`));
      made.controller.changed();
      assert.equal(await made.controller.saveNow(), false, line);
      assert.equal(made.controller.dirty, true, line);
      assert.deepEqual(made.notices, [`error: ${STRINGS_EN.saveRedirected}`], line);
      assert.deepEqual(made.errors.map((error) => error.code), ['redirected'], line);
      // The body and the headers of the page went to the address they were meant for, and nowhere else.
      assert.deepEqual(log, [line]);
    }
  });
});

test('review 1: the request asks not to be redirected; an answer that was redirected all the same is no save', async () => {
  const answers: Partial<Response>[] = [
    { status: 0, type: 'opaqueredirect' as ResponseType },
    { status: 200, redirected: true },
    { status: 204, type: 'opaqueredirect' as ResponseType },
    { status: 301 },
  ];
  for (const answer of answers) {
    let asked: RequestInit | undefined;
    const made = make(resolveSaveTarget('/files/7', '/files/7', PAGE), {
      etag: '"v1"',
      fetch: (async (_url: unknown, init?: RequestInit) => {
        asked = init;
        return { headers: { get: () => '"v2"' }, ...answer } as unknown as Response;
      }) as typeof fetch,
    });
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), false, JSON.stringify(answer));
    assert.equal(asked && asked.redirect, 'manual');
    assert.deepEqual(made.errors.map((error) => error.code), ['redirected'], JSON.stringify(answer));
    assert.equal(made.controller.currentETag, '"v1"', 'the ETag of an answer that was no save is not taken');
    assert.equal(made.controller.dirty, true);
  }
  assert.match(STRINGS_TR.saveRedirected, /yönlendirdi/);
});

// ------------------------------------------- a save whose answer is lost

test('review 3: the server wrote the file but its answer came too late: the next save is told as a conflict, and overwriting is the way out', async () => {
  await withServer(async (origin, log, state) => {
    state.slow = 300;
    const made = make(resolveSaveTarget({ timeoutMs: 60 }, '/file', `${origin}/page`), { etag: '"v1"' });
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), false);
    assert.deepEqual(made.errors.map((error) => error.code), ['timeout']);
    await new Promise((resolve) => setTimeout(resolve, 400));
    assert.equal(state.stored, 3, 'the server has the file all the same');
    state.slow = 0;
    assert.equal(await made.controller.saveNow(), false);
    assert.deepEqual(made.errors.map((error) => error.code), ['timeout', 'conflict']);
    assert.equal(made.notices[1], `error: ${STRINGS_EN.saveConflict}`);
    assert.match(STRINGS_EN.saveConflict, /earlier save/, 'the words do not blame somebody else alone');
    // Nothing overwrites by itself.
    assert.equal(await made.controller.saveNow(), false);
    // The user says so: no If-Match, and the file is saved.
    assert.equal(await made.controller.saveNow({ overwrite: true }), true);
    assert.equal(made.controller.dirty, false);
    assert.deepEqual(
      log.map((line) => line.split(' ')[3]),
      ['"v1"', '"v1"', '"v1"', '-'],
    );
    // From here on the ETag of the answer is followed again.
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), true);
    assert.equal(log[4].split(' ')[3], '"v3"');
  });
});

test('review 3: overwrite sends nothing when there is nothing to save, and waits for a save that is under way', async () => {
  const sent: RequestInit[] = [];
  const release: (() => void)[] = [];
  const made = make(resolveSaveTarget(true, '/files/7', PAGE), {
    etag: '"v1"',
    fetch: ((_url: unknown, init?: RequestInit) =>
      new Promise<Response>((resolve) => {
        sent.push(init || {});
        release.push(() => resolve({ status: sent.length === 1 ? 412 : 204, headers: { get: () => null } } as unknown as Response));
      })) as typeof fetch,
  });
  assert.equal(await made.controller.saveNow({ overwrite: true }), true);
  assert.equal(sent.length, 0);
  made.controller.changed();
  const first = made.controller.saveNow();
  await settle();
  const second = made.controller.saveNow({ overwrite: true });
  await settle();
  assert.equal(sent.length, 1, 'one request at a time');
  release[0]();
  assert.equal(await first, false);
  await settle();
  assert.equal(sent.length, 2);
  assert.equal('If-Match' in (sent[1].headers as object), false);
  release[1]();
  assert.equal(await second, true);
});

test('review 3: how long a save may take is the page\'s to say; the default leaves time for the upload', () => {
  assert.equal(saveTo.SAVE_TIMEOUT_MS, 120000);
  const target = resolveSaveTarget({ timeoutMs: 5000 }, '/files/7', PAGE);
  assert.equal(target.kind === 'request' && target.timeoutMs, 5000);
  const plain = resolveSaveTarget(true, '/files/7', PAGE);
  assert.equal(plain.kind === 'request' && plain.timeoutMs, 120000);
  for (const bad of [0, -1, NaN, Infinity, '30', null]) assert.throws(() => resolveSaveTarget({ timeoutMs: bad }, '/files/7', PAGE), TypeError, String(bad));
});

// --------------------------------------------------------------- headers

test('review 4: a header value that cannot be sent is refused when the option is read, not at every save', () => {
  for (const value of ['Büşra', 'İstanbul', '€', `a${char(0x2028)}b`, `a${char(127)}`, `a${char(10)}b`, `a${char(0)}`]) {
    assert.throws(() => resolveSaveTarget({ headers: { 'X-User': value } }, '/files/7', PAGE), TypeError, JSON.stringify(value));
  }
  // What a ByteString holds is taken: Latin-1 and a tab.
  const target = resolveSaveTarget({ headers: { 'X-User': 'Gül Ünal', 'X-Tab': `a${char(9)}b` } }, '/files/7', PAGE);
  assert.deepEqual(target.kind === 'request' && target.headers, [
    ['X-User', 'Gül Ünal'],
    ['X-Tab', `a${char(9)}b`],
  ]);
});

test('review 5: a header named __proto__ is refused, not lost', () => {
  assert.throws(() => resolveSaveTarget({ headers: JSON.parse('{"__proto__":"x"}') }, '/files/7', PAGE), TypeError);
  assert.throws(() => resolveSaveTarget({ headers: { ['__proto__']: 'x' } }, '/files/7', PAGE), TypeError);
  const target = resolveSaveTarget({ headers: JSON.parse('{"constructor":"y","toString":"z"}') }, '/files/7', PAGE);
  assert.deepEqual(target.kind === 'request' && target.headers, [
    ['constructor', 'y'],
    ['toString', 'z'],
  ]);
});

test('review 4: a request that fetch refuses to make is not called a server that could not be reached', async () => {
  const made = make(resolveSaveTarget('/files/7', undefined, PAGE), {
    fetch: (() => {
      throw new TypeError('Invalid header value <secret>');
    }) as unknown as typeof fetch,
  });
  made.controller.changed();
  assert.equal(await made.controller.saveNow(), false);
  assert.deepEqual(made.errors.map((error) => error.code), ['refused']);
  assert.deepEqual(made.notices, ['error: The file could not be saved']);
  assert.match(made.errors[0].message, /could not be made/);
});

// ------------------------------------------------ the application's own

test('review 8: a function that never answers does not keep the editor saving for ever', async () => {
  let calls = 0;
  const timers: { callback: () => void; ms: number; off: boolean }[] = [];
  const made = make(
    { kind: 'function', save: () => (++calls === 1 ? new Promise<void>(() => undefined) : undefined) },
    {
      timer: (callback, ms) => {
        const entry = { callback, ms, off: false };
        timers.push(entry);
        return () => void (entry.off = true);
      },
    },
  );
  made.controller.changed();
  const first = made.controller.saveNow();
  await settle();
  assert.equal(made.controller.saving, true);
  assert.deepEqual(timers.map((entry) => entry.ms), [saveTo.SAVE_FUNCTION_TIMEOUT_MS]);
  assert.equal(saveTo.SAVE_FUNCTION_TIMEOUT_MS, 600000);
  timers[0].callback();
  assert.equal(await first, false);
  assert.deepEqual(made.errors.map((error) => error.code), ['timeout']);
  assert.deepEqual(made.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveTimedOut });
  // The next one starts, and its time limit is called off when it answers.
  assert.equal(await made.controller.saveNow(), true);
  assert.equal(timers[1].off, true);
});

// ------------------------------------------------------ converted files

test('review 9: a converted file is not written over its original unless the page allows it', async () => {
  const sent: string[] = [];
  const fetcher = (async (url: unknown) => {
    sent.push(String(url));
    return { status: 204, headers: { get: () => null } } as unknown as Response;
  }) as typeof fetch;
  const converted: SaveInfo = { format: 'docx', convertedFrom: 'doc', name: 'old.docx' };

  for (const option of [true, '/files/old.doc', { method: 'POST' }, { url: 'https://app.example.com/files/old.doc#x', allowConverted: false }]) {
    const made = make(resolveSaveTarget(option, '/files/old.doc', PAGE), { fetch: fetcher, etag: '"v1"' }, converted);
    assert.equal(made.controller.blockedConversion, true, JSON.stringify(option));
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), false, JSON.stringify(option));
    assert.deepEqual(made.errors.map((error) => error.code), ['converted']);
    assert.deepEqual(made.notices, [`error: ${STRINGS_EN.saveConvertedNotAllowed('docx')}`]);
    assert.deepEqual(made.controller.state, { dirty: true, saving: false, error: STRINGS_EN.saveConvertedNotAllowed('docx') });
  }
  assert.deepEqual(sent, [], 'nothing was sent');

  // Allowed by the page, to another address, or through its own function: saved, and told what it was.
  const allowed = make(resolveSaveTarget({ allowConverted: true }, '/files/old.doc', PAGE), { fetch: fetcher }, converted);
  const elsewhere = make(resolveSaveTarget('/files/new.docx', '/files/old.doc', PAGE), { fetch: fetcher }, converted);
  let given: SaveInfo | undefined;
  const own = make({ kind: 'function', save: (_bytes, info) => void (given = info) }, {}, converted);
  for (const made of [allowed, elsewhere, own]) {
    assert.equal(made.controller.blockedConversion, false);
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), true);
  }
  assert.deepEqual(sent, ['https://app.example.com/files/old.doc', 'https://app.example.com/files/new.docx']);
  assert.deepEqual(given, converted);
  // A file that was not converted is not asked about.
  const plain = make(resolveSaveTarget(true, '/files/a.docx', PAGE), { fetch: fetcher }, { format: 'docx' });
  assert.equal(plain.controller.blockedConversion, false);
  assert.throws(() => resolveSaveTarget({ allowConverted: 'yes' }, '/files/old.doc', PAGE), TypeError);
  assert.match(STRINGS_TR.saveConvertedNotAllowed('xlsx'), /xlsx/);
});

// ------------------------------------------------------ what mount() ties

test('review 7: the ETag of a file that was fetched through a redirect is not kept', () => {
  const headers = { get: (name: string) => (name.toLowerCase() === 'etag' ? '"v1"' : null) };
  assert.equal(saveTo.fileETag({ redirected: false, headers }), '"v1"');
  assert.equal(saveTo.fileETag({ headers }), '"v1"');
  assert.equal(saveTo.fileETag({ redirected: true, headers }), null, 'it is of the address the answer came from, not of the one that is saved to');
  assert.equal(saveTo.fileETag({ redirected: false, headers: { get: () => null } }), null);
});

test('review 10: the name a saved file is known by follows its conversion', () => {
  assert.deepEqual(saveTo.saveInfoOf('xlsx', undefined, 'liste.xlsx'), { format: 'xlsx', name: 'liste.xlsx' });
  assert.deepEqual(saveTo.saveInfoOf('docx', undefined, undefined), { format: 'docx' });
  assert.deepEqual(saveTo.saveInfoOf('docx', { from: 'doc', to: 'docx' }, 'Rapor.DOC'), { format: 'docx', convertedFrom: 'doc', name: 'Rapor.docx' });
  assert.deepEqual(saveTo.saveInfoOf('docx', { from: 'rtf', to: 'docx' }, 'teklif.v2'), { format: 'docx', convertedFrom: 'rtf', name: 'teklif.v2.docx' });
  assert.deepEqual(saveTo.saveInfoOf('xlsx', { from: 'xls', to: 'xlsx' }, undefined), { format: 'xlsx', convertedFrom: 'xls' });
});

test('review 10: a protection that allows no editing locks the file; tracked changes do not', () => {
  assert.equal(saveTo.locksSaving(undefined), false);
  assert.equal(saveTo.locksSaving('trackedChanges'), false);
  for (const kind of ['readOnly', 'comments', 'forms']) assert.equal(saveTo.locksSaving(kind), true, kind);
});

test('review 10: leaving the page asks only while there is something unsaved, and the listener goes when the watch ends', () => {
  const listeners: ((event: { preventDefault(): void; returnValue?: unknown }) => void)[] = [];
  const win = {
    addEventListener: (type: string, listener: (event: any) => void) => void (type === 'beforeunload' && listeners.push(listener)),
    removeEventListener: (type: string, listener: (event: any) => void) => {
      const at = listeners.indexOf(listener);
      if (type === 'beforeunload' && at >= 0) listeners.splice(at, 1);
    },
  };
  let dirty = false;
  const unwatch = saveTo.watchLeaving(win, () => dirty);
  assert.equal(listeners.length, 1);
  const leave = () => {
    const event = { prevented: false, returnValue: undefined as unknown, preventDefault: () => void (event.prevented = true) };
    for (const listener of listeners.slice()) listener(event);
    return [event.prevented, event.returnValue];
  };
  assert.deepEqual(leave(), [false, undefined]);
  dirty = true;
  assert.deepEqual(leave(), [true, '']);
  unwatch();
  unwatch();
  assert.equal(listeners.length, 0);
  assert.deepEqual(leave(), [false, undefined]);
});
