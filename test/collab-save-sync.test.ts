// What a participant says about a save is asked of the server before it is believed; and who
// saves when the room asks. With the real save controller, a server the test plays, and
// participants who lie.

import assert from 'node:assert/strict';
import { test } from 'node:test';
import { createSaveSync, type SyncSession } from '../src/collab/save-sync';
import { SAVE_TIMEOUT_MS, SaveController, strongETag, type SaveHost, type SaveTarget } from '../src/embed/save-to';
import { STRINGS_EN } from '../src/ui/strings';

const settle = async () => {
  for (let i = 0; i < 40; i++) await Promise.resolve();
};

/** A file on a server that names it by an ETag, the controller of one participant, and that participant's session as the test sets it. */
function rig(options: { head?: boolean; typing?: () => boolean } = {}) {
  const server = { etag: '"v1"', puts: [] as { match: string | undefined }[], heads: 0, gets: 0, answer: undefined as number | undefined };
  const session: { state: SyncSession['state']; rev: number; pending: number; leads: boolean } = { state: 'online', rev: 3, pending: 0, leads: true };
  const asked: string[] = [];
  const serve = (async (_url: unknown, init?: RequestInit) => {
    const method = (init && init.method) || 'GET';
    const headers = ((init && init.headers) || {}) as { [name: string]: string };
    const reply = (status: number, etag?: string) => ({ status, ok: status >= 200 && status < 300, headers: { get: (name: string) => (name.toLowerCase() === 'etag' && etag !== undefined ? etag : null) }, body: { cancel: async () => undefined } }) as unknown as Response;
    if (method === 'HEAD') {
      server.heads++;
      return options.head === false ? reply(405) : reply(200, server.etag);
    }
    if (method === 'GET') {
      server.gets++;
      return headers['If-None-Match'] === server.etag ? reply(304) : reply(200, server.etag);
    }
    server.puts.push({ match: headers['If-Match'] });
    if (server.answer !== undefined) return reply(server.answer);
    if (headers['If-Match'] !== undefined && headers['If-Match'] !== server.etag) return reply(412);
    server.etag = `"v${server.puts.length + 1}"`;
    return reply(204, server.etag);
  }) as typeof fetch;
  const target: SaveTarget = { kind: 'request', url: 'https://app.example.com/files/7', method: 'PUT', headers: [], credentials: 'same-origin', sameAsFile: true, timeoutMs: SAVE_TIMEOUT_MS, allowConverted: true, onConflict: 'ask' };
  const host: SaveHost = {
    target,
    writable: () => true,
    commit: () => true,
    bytes: async () => new Uint8Array([1]),
    pending: options.typing,
    info: () => ({ format: 'xlsx' }),
    strings: () => STRINGS_EN,
    askOverwrite: () => void asked.push('asked'),
    etag: '"v1"',
    fetch: serve,
    timer: () => () => undefined,
  };
  const controller = new SaveController(host);
  let readOnly = false;
  const sync = createSaveSync({ controller, session: () => session, target: { url: target.url, credentials: 'same-origin' }, strong: strongETag, readOnly: () => readOnly, fetch: serve });
  return { server, session, controller, sync, asked, setReadOnly: (value: boolean) => (readOnly = value) };
}

test('a save that the server confirms is taken: its ETag, and a clean workbook when it is the workbook of that step', async () => {
  const r = rig();
  r.controller.changed(true);
  r.server.etag = '"v2"';
  await r.sync.saved({ versions: ['e:"v2"', 'h:1:x'], rev: 3 });
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], ['"v2"', false], 'the ETag of the server, and nothing left to save');
  // Changes this side made itself are never taken to be saved on someone's word.
  r.controller.changed();
  r.server.etag = '"v2b"';
  await r.sync.saved({ versions: ['e:"v2b"'], rev: 3 });
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], ['"v2b"', true], 'the ETag is taken, the changes made here stay to be saved');
  assert.equal(await r.controller.saveNow(), true, 'and are saved');
  r.controller.changed(true);
  r.server.etag = '"v3"';
  await r.sync.saved({ versions: ['e:"v3"'], rev: 2 });
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], ['"v3"', true], 'a save of an earlier step leaves what came after it to be saved');
  r.session.pending = 1;
  r.server.etag = '"v4"';
  await r.sync.saved({ versions: ['e:"v4"'], rev: 3 });
  assert.equal(r.controller.dirty, true, 'so does a step of this participant that still waits');
  await r.sync.saved({ versions: ['e:"v4"'] });
  assert.equal(r.controller.dirty, true, 'and a save that names no step');
});

test('a participant who lies about a save makes nothing clean and no save unguarded', async () => {
  const r = rig();
  r.controller.changed();
  // The server still has "v1"; someone says the file was saved as "forged", at this very step.
  await r.sync.saved({ versions: ['e:"forged"'], rev: 3 });
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], ['"v1"', true], 'nothing is believed');
  // Someone names the ETag this side knows already, with changes made here since: that file does not hold them.
  await r.sync.saved({ versions: ['e:"v1"'], rev: 3 });
  assert.equal(r.controller.dirty, true, 'the file this side read is not the file with its changes in it');
  // Someone changed the file outside the room and announces its ETag as a save of the room.
  r.server.etag = '"outside"';
  await r.sync.saved({ versions: ['e:"outside"'], rev: 3 });
  assert.equal(r.controller.currentETag, '"outside"', 'the server says so too: the ETag is what the file is known by');
  assert.equal(r.controller.dirty, true, 'but changes made here are not in it');
  // A weak ETag, a version that is no ETag, a session that is not online: nothing is asked or taken.
  const heads = r.server.heads;
  await r.sync.saved({ versions: ['e:W/"weak"'], rev: 3 });
  await r.sync.saved({ versions: ['h:1:abc'], rev: 3 });
  assert.equal(r.server.heads, heads, 'no question for what is no ETag');
  r.session.state = 'offline';
  r.server.etag = '"later"';
  await r.sync.saved({ versions: ['e:"later"'], rev: 3 });
  assert.equal(r.controller.currentETag, '"outside"', 'what arrives for a session that is not online is not taken');
});

test('an ETag that could not be confirmed is only a condition to try once after a conflict, and the question stays the user\'s', async () => {
  const r = rig();
  r.controller.changed();
  r.server.etag = '"v2"';
  // The server cannot be asked right now (it answers the question with an error).
  const real = r.server.etag;
  r.server.etag = '"hidden"';
  await r.sync.saved({ versions: ['e:"v2"'], rev: 3 });
  r.server.etag = real;
  assert.equal(r.controller.currentETag, '"v1"', 'not taken');
  const done = r.controller.saveNow();
  await settle();
  assert.deepEqual(r.server.puts.map((put) => put.match), ['"v1"', '"v2"'], 'the save meets a conflict and is tried once with what was said, as a condition');
  assert.equal(await done, true, 'which the server takes');
  // The same with a lie: the condition does not hold, and the user is asked.
  r.controller.changed();
  r.server.etag = '"elsewhere"';
  r.controller.hint('"wrong"');
  const again = r.controller.saveNow();
  await settle();
  assert.equal(r.asked.length, 1, 'a conflict is asked about');
  assert.equal(r.server.puts.every((put) => put.match !== undefined), true, 'and nothing was ever sent without a condition');
  // A save that is confirmed while the question is open does not answer it.
  r.server.etag = '"v9"';
  await r.sync.saved({ versions: ['e:"v9"'], rev: 3 });
  assert.equal(r.controller.asking, true, 'the question is still the user\'s to answer');
  r.controller.dismissConflict();
  assert.equal(await again, false, 'and the save did not happen');
});

test('a server without HEAD is asked with a GET that fetches nothing', async () => {
  const r = rig({ head: false });
  r.controller.changed(true);
  r.server.etag = '"v2"';
  await r.sync.saved({ versions: ['e:"v2"'], rev: 3 });
  assert.deepEqual([r.server.gets, r.controller.currentETag, r.controller.dirty], [1, '"v2"', false], 'confirmed by "not modified"');
  await r.sync.saved({ versions: ['e:"other"'], rev: 3 });
  assert.equal(r.controller.currentETag, '"v2"', 'and not confirmed by a file that is another one');
});

test('when the room asks for a save, the one who leads saves, as a save nobody asked for', async () => {
  let typing = false;
  const r = rig({ typing: () => typing });
  r.session.leads = false;
  assert.equal(await r.sync.sync(), false, 'who does not lead does not save');
  r.session.leads = true;
  r.setReadOnly(true);
  assert.equal(await r.sync.sync(), false, 'a view that is only read does not save');
  r.setReadOnly(false);
  r.session.pending = 1;
  assert.equal(await r.sync.sync(), false, 'nor one with a step on its way: the file would be of no step');
  r.session.pending = 0;
  typing = true;
  assert.equal(await r.sync.sync(), false, 'nor while something is being typed: it is not applied for a save nobody asked for');
  typing = false;
  assert.equal(await r.sync.sync(), true, 'the one who leads saves, also a workbook it had nothing to save of');
  assert.deepEqual(r.server.puts.map((put) => put.match), ['"v1"'], 'guarded by the ETag it knows');
  r.server.etag = '"elsewhere"';
  r.controller.changed(true);
  assert.equal(await r.sync.sync(), false, 'a conflict is not written over');
  assert.equal(r.asked.length, 0, 'and nobody is asked about a save nobody asked for');
});
