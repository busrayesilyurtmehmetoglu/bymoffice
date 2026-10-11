import assert from 'node:assert/strict';
import { createServer, request as httpRequest, type Server } from 'node:http';
import { test } from 'node:test';
import { SaveController, SaveError, resolveSaveTarget, type SaveHost, type SaveInfo, type SaveState, type SaveTarget } from '../src/embed/save-to';
import { STRINGS_EN } from '../src/ui/strings';

// What the second review of saveTo found: a page that saves from inside
// onSaveState had the same file sent twice (and was told of a conflict that
// was its own second request), a "could not be saved" that outlived the save
// that succeeded, an If-Match of the page that an overwrite still sent, a
// redirect that an old fetch followed without saying so, and addresses that
// were the same file and were not taken for it.

const PAGE = 'https://app.example.com/docs/page';
const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};

interface Sent {
  url: string;
  headers: { [name: string]: string };
  answer(status: number, more?: { etag?: string; url?: string; redirected?: boolean }): void;
}

interface Made {
  controller: SaveController;
  sent: Sent[];
  states: SaveState[];
  notices: string[];
  errors: SaveError[];
}

function make(target: SaveTarget, more: Partial<SaveHost> = {}, info: SaveInfo = { format: 'xlsx' }): Made {
  const made: Made = { controller: undefined as unknown as SaveController, sent: [], states: [], notices: [], errors: [] };
  const fetcher = (url: unknown, init?: RequestInit) =>
    new Promise<Response>((resolve) => {
      made.sent.push({
        url: String(url),
        headers: ((init && init.headers) || {}) as { [name: string]: string },
        answer(status, extra = {}) {
          const answer: { [name: string]: unknown } = { status, headers: { get: (name: string) => (name.toLowerCase() === 'etag' && extra.etag !== undefined ? extra.etag : null) } };
          if (extra.url !== undefined) answer.url = extra.url;
          if (extra.redirected !== undefined) answer.redirected = extra.redirected;
          resolve(answer as unknown as Response);
        },
      });
    });
  made.controller = new SaveController({
    target,
    writable: () => true,
    commit: () => true,
    bytes: async () => new Uint8Array([1, 2, 3]),
    info: () => info,
    strings: () => STRINGS_EN,
    notice: (message, kind) => void made.notices.push(`${kind}: ${message}`),
    onState: (state) => void made.states.push(state),
    onError: (error) => void made.errors.push(error as SaveError),
    fetch: fetcher as typeof fetch,
    timer: () => () => undefined,
    ...more,
  });
  return made;
}

const fileTarget = (): SaveTarget => resolveSaveTarget(true, 'https://app.example.com/files/7.xlsx', undefined);

// ------------------------------------------------- 1. re-entry from callbacks

test('review2 1: saveNow() from inside onSaveState ("save whenever something is unsaved") sends the file once', async () => {
  let made: Made;
  const results: boolean[] = [];
  made = make(fileTarget(), {
    etag: '"v1"',
    onState: (state) => {
      made.states.push(state);
      if (state.dirty) void made.controller.saveNow().then((ok) => results.push(ok));
    },
  });
  made.controller.changed();
  await settle();
  assert.equal(made.sent.length, 1, 'one request for one change');
  assert.equal(made.sent[0].headers['If-Match'], '"v1"');
  made.sent[0].answer(200, { etag: '"v2"' });
  await settle();
  assert.equal(made.sent.length, 1, 'and none after it was answered');
  assert.deepEqual(made.errors, [], 'no conflict of the page with itself');
  assert.deepEqual(made.notices, ['info: Saved']);
  assert.deepEqual([made.controller.dirty, made.controller.saving, made.controller.currentETag], [false, false, '"v2"']);
  assert.equal(results.every((ok) => ok === true), true, `every call was answered true: ${JSON.stringify(results)}`);
});

test('review2 1: the save under way is known before any callback of the host is called', async () => {
  // Each callback asks for a save when it is called while a save is being started or is under way.
  for (const from of ['onState', 'bytes', 'info', 'strings', 'notice', 'onSaved', 'commit'] as const) {
    let made: Made;
    let inside = 0;
    const again = (): void => {
      // The call from inside the call it makes itself is not followed any deeper.
      if (inside > 0) return;
      inside++;
      void made.controller.saveNow();
      inside--;
    };
    const more: Partial<SaveHost> = {};
    if (from === 'onState') more.onState = () => again();
    if (from === 'bytes') {
      more.bytes = async () => {
        again();
        return new Uint8Array([1]);
      };
    }
    if (from === 'info') {
      more.info = () => {
        again();
        return { format: 'xlsx' };
      };
    }
    if (from === 'strings') {
      more.strings = () => {
        again();
        return STRINGS_EN;
      };
    }
    if (from === 'notice') more.notice = () => again();
    if (from === 'onSaved') more.onSaved = () => again();
    if (from === 'commit') {
      more.commit = () => {
        again();
        return true;
      };
    }
    made = make(fileTarget(), more);
    made.controller.changed();
    void made.controller.saveNow();
    await settle();
    assert.equal(made.sent.length, 1, `${from}: one request while the first is not answered`);
    made.sent[0].answer(200);
    await settle();
    assert.equal(made.sent.length, 1, `${from}: nothing changed since, nothing more is sent`);
    assert.deepEqual([made.controller.dirty, made.controller.saving], [false, false], from);
  }
});

test('review2 1: saveNow() from inside onError does not start a second request, and a save that cannot start does not call itself for ever', async () => {
  let made: Made;
  made = make(fileTarget(), {
    etag: '"v1"',
    onError: (error) => {
      made.errors.push(error as SaveError);
      if (made.errors.length < 3) void made.controller.saveNow({ overwrite: true });
    },
  });
  made.controller.changed();
  void made.controller.saveNow();
  await settle();
  made.sent[0].answer(412);
  await settle();
  // The overwrite the page asked for from inside onError: one request, after the one that failed.
  assert.equal(made.sent.length, 2);
  assert.equal(made.sent[1].headers['If-Match'], undefined);
  made.sent[1].answer(200);
  await settle();
  assert.deepEqual([made.controller.dirty, made.controller.saving, made.sent.length], [false, false, 2]);

  // What is being typed cannot be applied, and onError asks again every time: the call from inside says no and reports nothing more.
  const errors: string[] = [];
  let stuck: Made;
  stuck = make(fileTarget(), {
    commit: () => false,
    onError: (error) => {
      errors.push((error as SaveError).code);
      void stuck.controller.saveNow();
    },
  });
  stuck.controller.changed();
  assert.equal(await stuck.controller.saveNow(), false);
  assert.deepEqual(errors, ['pending'], 'told once for the one press');
  assert.equal(stuck.sent.length, 0);
});

test('review2 1: a host whose callback throws while the save ends does not leave the editor saving for ever', async () => {
  let fail = false;
  const made = make(fileTarget(), {
    strings: () => {
      if (fail) throw new Error('no strings');
      return STRINGS_EN;
    },
  });
  made.controller.changed();
  const first = made.controller.saveNow();
  await settle();
  fail = true;
  made.sent[0].answer(500);
  assert.equal(await first, false, 'the save answers, and does not reject');
  assert.equal(made.controller.saving, false);
  fail = false;
  const second = made.controller.saveNow();
  await settle();
  assert.equal(made.sent.length, 2, 'the next save is sent');
  made.sent[1].answer(200);
  assert.equal(await second, true);
});

test('review2 1: saveNow() from inside the onState of a change saves when the change is over, not while the view is still making it', async () => {
  // As a workbook view is while a cell is written: its editor is open until the change was told, and nothing can be applied meanwhile.
  let committing = false;
  let made: Made;
  made = make(fileTarget(), {
    commit: () => !committing,
    onState: (state) => {
      made.states.push(state);
      if (state.dirty && !state.saving) void made.controller.saveNow();
    },
  });
  committing = true;
  made.controller.changed();
  assert.equal(made.sent.length, 0, 'nothing is taken from a view that is half way through its change');
  committing = false;
  await settle();
  assert.deepEqual(made.errors, [], 'and it is not told that something could not be applied');
  assert.equal(made.sent.length, 1);
  made.sent[0].answer(200);
  await settle();
  assert.deepEqual([made.controller.dirty, made.controller.state.error, made.sent.length], [false, undefined, 1]);
  // Destroyed in between: the call that waited answers false and sends nothing.
  committing = true;
  made.controller.changed();
  made.controller.destroy();
  await settle();
  assert.equal(made.sent.length, 1);
});

// ------------------------------------------------------ 2. the pending error

test('review2 2: "could not be applied" said while a save is under way is gone when that save succeeds', async () => {
  let stuck = false;
  const made = make(fileTarget(), { commit: () => !stuck });
  made.controller.changed();
  const first = made.controller.saveNow();
  await settle();
  stuck = true;
  assert.equal(await made.controller.saveNow(), false);
  assert.deepEqual(made.controller.state, { dirty: true, saving: true, error: STRINGS_EN.savePending });
  assert.deepEqual(made.errors.map((error) => error.code), ['pending']);
  stuck = false;
  made.sent[0].answer(200);
  assert.equal(await first, true);
  const state = made.controller.state;
  assert.equal(state.error, undefined, 'saved and "could not be saved" are not said together');
  assert.deepEqual([state.dirty, state.saving, typeof state.lastSavedAt], [false, false, 'number']);
  assert.deepEqual(made.states[made.states.length - 1], state, 'and the page was told so');

  // A save that fails says why it failed, whatever was said before it.
  made.controller.changed();
  const second = made.controller.saveNow();
  await settle();
  stuck = true;
  await made.controller.saveNow();
  made.sent[1].answer(423);
  assert.equal(await second, false);
  assert.equal(made.controller.state.error, STRINGS_EN.saveLocked);
});

// ---------------------------------------------------------- 7. overwrite

test('review2 7: overwrite sends no If-Match, also not the one the page gave in its headers, however it is spelled', async () => {
  for (const name of ['If-Match', 'if-match', 'IF-MATCH', 'iF-mAtCh']) {
    const made = make(resolveSaveTarget({ headers: { [name]: '"mine"', 'X-Token': 't' } }, 'https://app.example.com/files/7.xlsx', undefined), { etag: '"v1"' });
    made.controller.changed();
    void made.controller.saveNow();
    await settle();
    assert.deepEqual(
      Object.keys(made.sent[0].headers).filter((header) => header.toLowerCase() === 'if-match').map((header) => made.sent[0].headers[header]),
      ['"mine"'],
      `${name}: a save that does not overwrite carries the header of the page`,
    );
    made.sent[0].answer(412);
    await settle();
    void made.controller.saveNow({ overwrite: true });
    await settle();
    const headers = made.sent[1].headers;
    assert.deepEqual(Object.keys(headers).filter((header) => header.toLowerCase() === 'if-match'), [], `${name}: the overwrite carries none`);
    assert.equal(headers['X-Token'], 't', 'the other headers of the page stay');
  }
});

// ------------------------------------------- 8. a fetch that follows redirects

test('review2 8: an answer that names another address than the one asked is a redirect, whatever else it says', async () => {
  const cases: [answer: string | undefined, saved: boolean][] = [
    [undefined, true],
    ['', true],
    ['https://app.example.com/files/7.xlsx', true],
    ['https://app.example.com/files/7.xlsx#top', true],
    ['https://APP.example.com:443/files/./7.xlsx', true],
    ['not an address ::', true],
    ['https://app.example.com/login', false],
    ['https://app.example.com/files/7.xlsx?next=1', false],
    ['https://other.example.com/files/7.xlsx', false],
    ['http://app.example.com/files/7.xlsx', false],
  ];
  for (const [url, saved] of cases) {
    const made = make(fileTarget(), { etag: '"v1"' });
    made.controller.changed();
    const done = made.controller.saveNow();
    await settle();
    made.sent[0].answer(200, { etag: '"of-the-answer"', url });
    assert.equal(await done, saved, String(url));
    assert.deepEqual(made.errors.map((error) => error.code), saved ? [] : ['redirected'], String(url));
    assert.equal(made.controller.currentETag, saved ? '"of-the-answer"' : '"v1"', `${url}: the ETag of a login page is not taken for the file's`);
    assert.equal(made.controller.dirty, !saved, String(url));
  }
  // An address with an empty query is the address without one.
  const made = make(resolveSaveTarget('https://app.example.com/files/7.xlsx?', undefined, undefined));
  made.controller.changed();
  const done = made.controller.saveNow();
  await settle();
  made.sent[0].answer(200, { url: 'https://app.example.com/files/7.xlsx' });
  assert.equal(await done, true, 'an empty query');
});

test('review2 8: a fetch that does not know "redirect" and follows to the login page (a real server): not saved', async () => {
  const hits: string[] = [];
  const server: Server = createServer((req, res) => {
    req.resume();
    req.on('end', () => {
      hits.push(`${req.method} ${req.url}`);
      if (req.url === '/file') res.writeHead(302, { location: '/login' }).end();
      else res.writeHead(200, { etag: '"login-page"' }).end('<form>login</form>');
    });
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  try {
    const url = `http://127.0.0.1:${(server.address() as { port: number }).port}/file`;
    // What a fetch made of XMLHttpRequest does: the redirect is followed inside the request, and the answer has
    // neither `redirected` nor `type`; only its `url` (responseURL) tells.
    const old = (address: string, init: { method: string; headers: { [name: string]: string }; body: Uint8Array }) =>
      new Promise<unknown>((resolve, reject) => {
        const go = (target: string, method: string, body: Uint8Array | undefined, hops: number): void => {
          const outgoing = httpRequest(target, { method, headers: body ? init.headers : {} }, (incoming) => {
            incoming.resume();
            if (incoming.statusCode === 302 && hops < 5) return go(new URL(incoming.headers.location as string, target).href, 'GET', undefined, hops + 1);
            resolve({ status: incoming.statusCode, url: target, headers: { get: (name: string) => (incoming.headers[name.toLowerCase()] as string | undefined) ?? null } });
          });
          outgoing.on('error', reject);
          outgoing.end(body ? Buffer.from(body) : undefined);
        };
        go(address, init.method, init.body, 0);
      });
    const errors: string[] = [];
    const controller = new SaveController({
      target: resolveSaveTarget(true, url, undefined),
      writable: () => true,
      commit: () => true,
      bytes: async () => new Uint8Array([1]),
      info: () => ({ format: 'xlsx' }),
      strings: () => STRINGS_EN,
      onError: (error) => void errors.push((error as SaveError).code),
      fetch: old as unknown as typeof fetch,
    });
    controller.changed();
    assert.equal(await controller.saveNow(), false);
    assert.deepEqual(hits, ['PUT /file', 'GET /login']);
    assert.deepEqual(errors, ['redirected']);
    assert.deepEqual([controller.dirty, controller.currentETag], [true, undefined]);
  } finally {
    await new Promise<void>((resolve) => server.close(() => resolve()));
  }
});

// ------------------------------------------------ 9. the address of the file

test('review2 9: an address is the one the file came from when origin, path and query are the same, however it is written', () => {
  const same = (file: string, to: string): boolean => (resolveSaveTarget(to, file, PAGE) as Extract<SaveTarget, { kind: 'request' }>).sameAsFile;
  for (const [file, to] of [
    ['/files/42.doc', '/files/42.doc'],
    ['/files/42.doc', 'https://app.example.com/files/42.doc'],
    ['/files/42.doc', 'https://APP.example.com:443/files/42.doc#x'],
    ['/files/42.doc', '/files/./42.doc'],
    ['/files/42.doc', '//app.example.com/files/42.doc'],
    ['/files/42.doc', '/files/42.doc?'],
    ['/files/42.doc?', '/files/42.doc'],
    ['/files/42.doc?v=3', '/files/42.doc?v=3#top'],
  ]) {
    assert.equal(same(file, to), true, `${file} and ${to}`);
  }
  // What a server may tell apart is told apart: the query, the case of the path, how it is encoded, a slash at the end.
  for (const [file, to] of [
    ['/files/42.doc', '/files/42.doc?save=1'],
    ['/files/42.doc?v=3', '/files/42.doc'],
    ['/files/42.doc', '/files/42.DOC'],
    ['/files/42', '/files/42/'],
    ['/files/%41.doc', '/files/A.doc'],
    ['/files/42.doc', 'http://app.example.com/files/42.doc'],
    ['/files/42.doc', 'https://app.example.com:8443/files/42.doc'],
  ]) {
    assert.equal(same(file, to), false, `${file} and ${to}`);
  }
});

test('review2 9: an address written another way gets If-Match and the guard of a converted file all the same', async () => {
  const converted: SaveInfo = { format: 'docx', convertedFrom: 'doc' };
  const guarded = make(resolveSaveTarget('/files/42.doc?', '/files/42.doc', PAGE), {}, converted);
  assert.equal(guarded.controller.blockedConversion, true);
  const made = make(resolveSaveTarget('/files/7.xlsx?', '/files/7.xlsx', PAGE), { etag: '"v1"' });
  made.controller.changed();
  void made.controller.saveNow();
  await settle();
  assert.equal(made.sent[0].headers['If-Match'], '"v1"');
  assert.equal(made.sent[0].url, 'https://app.example.com/files/7.xlsx?', 'sent to the address as the page wrote it');
});

test('review2 9: a converted file that was given as bytes is not saved to an address that ends in the old extension unless the page allows it', async () => {
  const converted: SaveInfo = { format: 'docx', convertedFrom: 'doc' };
  for (const to of ['/files/42.doc', '/files/42.DOC?v=1', '/files/old.rtf', 'https://app.example.com/a/b.xls#x', { url: '/files/42.doc', method: 'POST' }]) {
    const made = make(resolveSaveTarget(to, undefined, PAGE), {}, converted);
    assert.equal(made.controller.blockedConversion, true, JSON.stringify(to));
    made.controller.changed();
    assert.equal(await made.controller.saveNow(), false, JSON.stringify(to));
    assert.deepEqual(made.errors.map((error) => error.code), ['converted'], JSON.stringify(to));
    assert.equal(made.sent.length, 0, 'nothing is sent');
  }
  for (const to of ['/files/42.docx', '/files/42', '/files/42.doc/save', '/save?name=42.doc', { url: '/files/42.doc', allowConverted: true }]) {
    const made = make(resolveSaveTarget(to, undefined, PAGE), {}, converted);
    assert.equal(made.controller.blockedConversion, false, JSON.stringify(to));
  }
  // A file that was not converted is saved wherever the page says.
  const plain = make(resolveSaveTarget('/files/42.doc', undefined, PAGE), {}, { format: 'docx' });
  assert.equal(plain.controller.blockedConversion, false);
  // A file that was fetched: only its own address is guarded, as before.
  const fetched = make(resolveSaveTarget('/copies/42.doc', '/files/42.doc', PAGE), {}, converted);
  assert.equal(fetched.controller.blockedConversion, false);
  // The target reads as before for an address that is not one of an old format.
  assert.deepEqual(Object.keys(resolveSaveTarget('/files/42.docx', undefined, PAGE)).sort(), ['allowConverted', 'credentials', 'headers', 'kind', 'method', 'onConflict', 'sameAsFile', 'timeoutMs', 'url']);
});
