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
  const server = { etag: '"v1"', puts: [] as { match: string | undefined }[], heads: 0, gets: 0, answer: undefined as number | undefined, down: false };
  const session: { state: SyncSession['state']; rev: number; pending: number; leads: boolean } = { state: 'online', rev: 3, pending: 0, leads: true };
  const asked: string[] = [];
  const serve = (async (_url: unknown, init?: RequestInit) => {
    const method = (init && init.method) || 'GET';
    const headers = ((init && init.headers) || {}) as { [name: string]: string };
    const reply = (status: number, etag?: string) => ({ status, ok: status >= 200 && status < 300, headers: { get: (name: string) => (name.toLowerCase() === 'etag' && etag !== undefined ? etag : null) }, body: { cancel: async () => undefined } }) as unknown as Response;
    if (method === 'HEAD') {
      if (server.down) throw new TypeError('no connection');
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
  // The server cannot be asked right now (the question does not get through).
  r.server.down = true;
  await r.sync.saved({ versions: ['e:"v2"'], rev: 3 });
  r.server.down = false;
  assert.equal(r.controller.currentETag, '"v1"', 'not taken');
  const done = r.controller.saveNow();
  await settle();
  assert.deepEqual(r.server.puts.map((put) => put.match), ['"v1"', '"v2"'], 'the save meets a conflict and is tried once with what was said, as a condition');
  assert.equal(await done, true, 'which the server takes');
  // A name the server itself says is not the file's is no hint: a save that meets a conflict is not sent again with it.
  r.controller.changed();
  await r.sync.saved({ versions: ['e:"never-was"'], rev: 3 });
  r.server.etag = '"moved"';
  const before = r.server.puts.length;
  const refused = r.controller.saveNow();
  await settle();
  assert.deepEqual(r.server.puts.slice(before).map((put) => put.match), ['"v3"'], 'sent once, with the name it knows; not again with what was refuted');
  r.controller.dismissConflict();
  assert.equal(await refused, false, 'the conflict is the user\'s');
  r.asked.length = 0;
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

test('a server that names a file by its content names it so again when the content comes back: the name is taken again, and the next save is no conflict', async () => {
  const r = rig();
  // X -> Y -> X, each said by someone and confirmed by the server.
  r.controller.changed(true);
  r.server.etag = '"Y"';
  await r.sync.saved({ versions: ['e:"Y"'], rev: 3 });
  assert.equal(r.controller.currentETag, '"Y"', 'the file is known by its new name');
  r.controller.changed(true);
  r.server.etag = '"v1"';
  await r.sync.saved({ versions: ['e:"v1"'], rev: 3 });
  assert.equal(r.controller.currentETag, '"v1"', 'and by the name it had before, when the server says so now');
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), true, 'a save is taken');
  assert.deepEqual([r.server.puts.map((put) => put.match), r.asked], [['"v1"'], []], 'sent once with the name the file has, and nobody is asked');
  // A name that is only said (the server could not be asked) and is a past one is still no hint.
  r.server.down = true;
  await r.sync.saved({ versions: ['e:"Y"'], rev: 3 });
  r.server.down = false;
  r.controller.changed();
  r.server.etag = '"elsewhere"';
  const before = r.server.puts.length;
  const refused = r.controller.saveNow();
  await settle();
  assert.equal(r.server.puts.length, before + 1, 'a conflict is not tried again with a name the file had before');
  r.controller.dismissConflict();
  assert.equal(await refused, false, 'it is the user\'s');
});

test('a word of a save that comes again after this side saved (late, twice, or a lie) is asked of the server anew: it moves nothing and makes nothing clean', async () => {
  const r = rig();
  // Someone saves; the server confirms.
  r.controller.changed(true);
  r.server.etag = '"Y"';
  await r.sync.saved({ versions: ['e:"Y"'], rev: 3 });
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], ['"Y"', false], 'taken');
  // This side saves its own work: the file has another name. Then work of someone else arrives.
  r.controller.changed();
  assert.equal(await r.controller.saveNow(), true, 'saved');
  const mine = r.controller.currentETag;
  assert.notEqual(mine, '"Y"', 'the file has another name now');
  r.controller.changed(true);
  const heads = r.server.heads;
  // The same word again, a moment later.
  await r.sync.saved({ versions: ['e:"Y"'], rev: 3 });
  assert.equal(r.server.heads, heads + 1, 'the server is asked again: what it said of that name was said of a file that is gone');
  assert.deepEqual([r.controller.currentETag, r.controller.dirty], [mine, true], 'the name stays the one the server gave, and the work of the other is still to be saved');
  r.asked.length = 0;
  assert.equal(await r.controller.saveNow(), true, 'the next save is taken');
  assert.deepEqual(r.asked, [], 'and nobody is asked about a conflict');
  // While the file is still known by that name, the same word within a moment asks nothing.
  r.controller.changed(true);
  r.server.etag = '"Z"';
  await r.sync.saved({ versions: ['e:"Z"'], rev: 3 });
  const asked = r.server.heads;
  r.controller.changed(true);
  await r.sync.saved({ versions: ['e:"Z"'], rev: 3 });
  assert.equal(r.server.heads, asked, 'no second question for the name the file is known by');
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

/**
 * Participants of one room, each with the real controller and the real glue of mount(), on one file on a
 * server that takes a save only when it is conditional on what the file is known by. A workbook is
 * the list of what was typed, in the order of the room; its file is that list as bytes, with the step.
 */
function room(count = 2, options: { auto?: boolean[]; clean?: (name: string) => void; put?: (name: string, texts: string[]) => void; latency?: () => number; /** Sides whose own file is never byte for byte the one the others write (a document whose model and file differ in a detail). */ apart?: number[] } = {}) {
  const encoder = new TextEncoder();
  const fileOf = (book: string[], rev: number): Uint8Array => encoder.encode(JSON.stringify({ rev, book }));
  const server = { etag: '"v1"', n: 1, body: fileOf([], 0), puts: [] as { by: string; match: string | undefined; status: number }[], gets: [] as string[], asked: [] as (string | undefined)[] };
  let clock = 1000;
  const order: string[] = [];
  const sides = Array.from({ length: count }, (_unused, index) => {
    // A name of its own for every participant, however many there are.
    const name = index < 26 ? 'abcdefghijklmnopqrstuvwxyz'[index] : `p${index}`;
    const session = { state: 'online' as SyncSession['state'], rev: 0, pending: 0, leads: index === 0 };
    const book: string[] = [];
    let timers: { run: () => void; at: number }[] = [];
    const errors: string[] = [];
    const states: (string | undefined)[] = [];
    const serve = (async (_url: unknown, init?: RequestInit) => {
      const method = (init && init.method) || 'GET';
      const headers = ((init && init.headers) || {}) as { [name: string]: string };
      const reply = (status: number, etag?: string, body?: Uint8Array) =>
        ({
          status,
          ok: status >= 200 && status < 300,
          headers: { get: (key: string) => (key.toLowerCase() === 'etag' && etag !== undefined ? etag : key.toLowerCase() === 'content-length' && body ? String(body.length) : null) },
          body: { cancel: async () => undefined },
          arrayBuffer: async () => (body || new Uint8Array(0)).slice().buffer,
        }) as unknown as Response;
      // An answer of the server takes its time, when the test says so: saves and words of saves cross each other.
      if (options.latency) await new Promise<void>((resolve) => void timer(resolve, options.latency!()));
      if (method !== 'PUT') server.asked.push(headers['X-Who']);
      if (method === 'HEAD') return reply(200, server.etag);
      if (method !== 'PUT') {
        server.gets.push(name);
        return reply(200, server.etag, server.body);
      }
      // The answer takes a moment: two saves that were sent at once are both on their way before either is answered.
      await Promise.resolve();
      const ok = headers['If-Match'] === server.etag;
      if (ok) {
        server.etag = `"v${++server.n}"`;
        server.body = (init as { body: Uint8Array }).body;
        if (options.put) options.put(name, stored());
      }
      server.puts.push({ by: name, match: headers['If-Match'], status: ok ? 204 : 412 });
      return ok ? reply(204, server.etag) : reply(412);
    }) as typeof fetch;
    const target: SaveTarget = { kind: 'request', url: 'https://app.example.com/files/7', method: 'PUT', headers: [], credentials: 'same-origin', sameAsFile: true, timeoutMs: SAVE_TIMEOUT_MS, allowConverted: true, onConflict: 'ask' };
    const asked: string[] = [];
    // eslint-disable-next-line prefer-const
    let sync: ReturnType<typeof createSaveSync>;
    const timer = (run: () => void, ms: number): (() => void) => {
      const entry = { run, at: clock + ms };
      timers.push(entry);
      return () => {
        timers = timers.filter((other) => other !== entry);
      };
    };
    const host: SaveHost = {
      target,
      writable: () => true,
      commit: () => true,
      bytes: async () => fileOf(book, session.rev),
      info: () => ({ format: 'xlsx' }),
      strings: () => STRINGS_EN,
      askOverwrite: () => void asked.push('asked'),
      onError: (error) => void errors.push(error.code),
      onState: (state) => {
        states.push(state.autoWaits);
        if (!state.dirty && !state.saving && options.clean) options.clean(name);
      },
      etag: '"v1"',
      fetch: serve,
      autoSave: options.auto && options.auto[index] === false ? undefined : { delayMs: 1000, maxWaitMs: 5000 },
      now: () => clock,
      timer,
      // As mount() ties them: one participant saves for all, and a conflict with a save of the room is none.
      autoAllowed: () => sync.autoAllowed(),
      autoWaiting: () => sync.waiting(),
      conflictWait: () => sync.conflictWait(),
      // Tells the others, as the session does.
      onSaved: () => {
        sync.ownSaved();
        const tag = controller.currentETag as string;
        for (const other of sides) if (other.name !== name) void other.sync.saved({ versions: [`e:${tag}`], rev: session.rev });
      },
    };
    const controller: SaveController = new SaveController(host);
    sync = createSaveSync({
      controller,
      session: () => session,
      target: { url: target.url, credentials: 'same-origin', headers: [['X-Who', 'a user'], ['Content-Type', 'x/y'], ['If-Match', 'never']] },
      people: () => count,
      strong: strongETag,
      readOnly: () => false,
      fetch: serve,
      timer,
      now: () => clock,
      random: () => ((index * 7919 + 13) % 97) / 97,
      ownBytes: async () => ({ bytes: options.apart && options.apart.indexOf(index) >= 0 ? encoder.encode(JSON.stringify({ rev: session.rev, book, apart: true })) : fileOf(book, session.rev), rev: session.pending === 0 ? session.rev : undefined }),
    });
    const fire = (): boolean => {
      const due = timers.filter((entry) => entry.at <= clock);
      timers = timers.filter((entry) => entry.at > clock);
      for (const entry of due) entry.run();
      return due.length > 0;
    };
    return { name, session, controller, sync, errors, asked, fire, book, states, serve };
  });
  /** Time passes, and everything that is due happens. */
  const pass = async (ms: number): Promise<void> => {
    const end = clock + ms;
    while (clock < end) {
      clock = Math.min(end, clock + 250);
      for (let round = 0; round < 20; round++) {
        let any = false;
        for (const side of sides) if (side.fire()) any = true;
        await settle();
        if (!any) break;
      }
    }
  };
  /** One participant types something: a change of its own there, one of someone else for the others; a step for all. */
  const edit = (who: number, text = `t${order.length}`): void => {
    order.push(text);
    for (const side of sides) {
      side.session.rev++;
      side.book.push(text);
    }
    sides[who].controller.changed();
    sides[who].sync.localChanged();
    sides.forEach((side, index) => {
      if (index !== who) side.sync.remoteChanged();
    });
  };
  /** What the stored file holds. */
  function stored(): string[] {
    try {
      return (JSON.parse(new TextDecoder().decode(server.body)) as { book: string[] }).book;
    } catch {
      return [];
    }
  }
  return { server, sides, pass, edit, stored, fileOf };
}

test('auto-save on both sides of a room: one participant saves, once, and nobody is told of a conflict', async () => {
  const r = room();
  const [a, b] = r.sides;
  // The one who does not save for all changes something.
  r.edit(1);
  await r.pass(1500);
  assert.deepEqual(r.server.puts.map((put) => [put.by, put.status]), [['a', 204]], 'one save, by the one who saves for all');
  assert.deepEqual([a.errors, b.errors, a.asked, b.asked], [[], [], [], []], 'nobody is told of a conflict or asked');
  assert.deepEqual([a.controller.dirty, a.controller.currentETag], [false, r.server.etag], 'the saver has nothing left to save');
  assert.equal(b.controller.currentETag, r.server.etag, 'the other knows the file by its new name');
  assert.equal(b.controller.dirty, false, 'and has nothing left to save either: it looked at the file, which holds what it changed');
  assert.deepEqual(r.stored(), ['t0'], 'as it does');
  // Then the saver changes something, then the other again: still one save each time.
  r.edit(0);
  await r.pass(1500);
  r.edit(1);
  await r.pass(1500);
  assert.deepEqual(r.server.puts.map((put) => [put.by, put.status]), [['a', 204], ['a', 204], ['a', 204]], 'one save for every change');
  assert.deepEqual([a.errors, b.errors], [[], []], 'and still no conflict');
  // The other saves by hand at the moment the saver saves by itself: one of the two is refused, and neither is told so.
  r.edit(1);
  r.server.puts.length = 0;
  await r.pass(750);
  const byHand = b.controller.saveNow();
  await r.pass(5000);
  assert.equal(await byHand, true, 'the save by hand went through');
  assert.equal(r.server.puts.some((put) => put.status === 412), true, 'one save met the other (or the test does not test what it says)');
  assert.deepEqual([a.errors, b.errors, a.asked, b.asked], [[], [], [], []], 'a conflict with a save of the room is no conflict');
  assert.equal(a.controller.state.autoPaused === true, false, 'and auto-save goes on');
});

test('auto-save in a room: when the one who saves for all leaves, the next one does', async () => {
  const r = room();
  const [a, b] = r.sides;
  r.edit(1);
  // The saver is gone before it saved.
  a.session.state = 'closed';
  a.controller.destroy();
  b.session.leads = true;
  b.sync.roomChanged();
  await r.pass(1500);
  assert.deepEqual(r.server.puts.filter((put) => put.by === 'b').map((put) => put.status), [204], 'the one who is left saves what it changed');
  assert.equal(b.controller.dirty, false, 'and has nothing left to save');
  // Offline, nobody knows who saves: nothing is sent.
  b.session.state = 'offline';
  b.session.leads = false;
  b.controller.changed();
  await r.pass(6000);
  assert.equal(r.server.puts.filter((put) => put.by === 'b').length, 1, 'nothing is saved by itself while the room cannot be reached');
  b.session.state = 'online';
  b.session.leads = true;
  b.sync.roomChanged();
  await r.pass(6000);
  assert.equal(r.server.puts.filter((put) => put.by === 'b').length, 2, 'and it is, once the room is back');
});

test('whoever says the file was saved, and whatever it is in the room: a participant\'s own work is saved when the file holds it, and not before', async () => {
  const r = room();
  const [a, b] = r.sides;
  // The victim (b) types; its step is placed. Someone writes an OLD file to the server, and says it is the file of this step.
  r.edit(1, 'the work of the victim');
  r.server.etag = '"old-file-put-again"';
  r.server.body = r.fileOf([], b.session.rev);
  a.controller.destroy();
  await b.sync.saved({ versions: [`e:${r.server.etag}`], rev: b.session.rev });
  await r.pass(2000);
  assert.equal(b.controller.dirty, true, 'the work of the victim is in no file: it stays to be saved, whoever said otherwise');
  assert.equal(b.controller.currentETag, '"old-file-put-again"', 'the file is known by its new name (the server says so)');
  assert.equal(b.controller.state.autoWaits, 'room', 'and it says that it waits for the one who saves for all');
  // Nobody saves it: after a while the victim does, itself, and the file holds the work.
  await r.pass(20000);
  assert.deepEqual(r.stored(), ['the work of the victim'], 'the work is in the stored file');
  assert.deepEqual([b.controller.dirty, b.errors, b.asked], [false, [], []], 'saved, by a save that was conditional on the file as it was, and without a word of a conflict');
  assert.deepEqual(r.server.puts.map((put) => [put.by, put.status]), [['b', 204]], 'one save');
});

test('the proof of a save: the same step and the same bytes, by the name that was said, and nothing less', async () => {
  for (const [name, spoil, clean] of [
    ['the file of the step', () => undefined, true],
    ['a file of another workbook', (r: ReturnType<typeof room>) => void (r.server.body = r.fileOf(['something else'], 1)), false],
    ['a file of another step', (r: ReturnType<typeof room>) => void (r.server.body = r.fileOf(['mine'], 0)), false],
    ['a file that is another one by now', (r: ReturnType<typeof room>) => void (r.server.etag = '"moved-on"'), false],
    ['a workbook that was changed since', (r: ReturnType<typeof room>) => r.edit(1, 'later'), false],
    ['a step of this side on its way', (r: ReturnType<typeof room>) => void (r.sides[1].session.pending = 1), false],
    ['a room that cannot be reached', (r: ReturnType<typeof room>) => void (r.sides[1].session.state = 'offline'), false],
  ] as [string, (r: ReturnType<typeof room>) => void, boolean][]) {
    const r = room(2, { auto: [false, false] });
    const b = r.sides[1];
    r.edit(1, 'mine');
    r.server.etag = '"saved-by-someone"';
    r.server.body = r.fileOf(['mine'], 1);
    await b.sync.saved({ versions: ['e:"saved-by-someone"'], rev: 1 });
    assert.equal(b.controller.dirty, true, `${name}: not on the word alone`);
    spoil(r);
    await r.pass(1000);
    assert.equal(b.controller.dirty, !clean, `${name}: ${clean ? 'the work is in the file, and nothing is left to save' : 'the work stays to be saved'}`);
  }
});

test('auto-save in a room never stops without a word: when the saver does not save, the others do after a while; offline it waits, and says so', async () => {
  // The one who would save for all has auto-save off.
  const off = room(2, { auto: [false, true] });
  off.edit(1, 'typed by b');
  await off.pass(3000);
  assert.deepEqual([off.server.puts.length, off.sides[1].controller.state.autoWaits], [0, 'room'], 'at first b waits for the saver, and says so');
  await off.pass(20000);
  assert.deepEqual(off.stored(), ['typed by b'], 'then it saves its work itself');
  assert.equal(off.sides[1].controller.state.autoWaits, undefined, 'and waits for nothing');
  // Its work again: the wait starts over.
  off.edit(1, 'more');
  await off.pass(3000);
  assert.equal(off.server.puts.length, 1, 'a new change waits its own time');
  await off.pass(20000);
  assert.deepEqual(off.stored(), ['typed by b', 'more'], 'and is saved in its turn');
  // Three, the saver is stuck: the two others do not both write over each other.
  const three = room(3, { auto: [false, true, true] });
  three.edit(1, 'one');
  three.edit(2, 'two');
  await three.pass(30000);
  assert.deepEqual(three.stored(), ['one', 'two'], 'the file holds the work of both');
  assert.deepEqual(three.sides.slice(1).map((side) => [side.controller.dirty, side.errors, side.asked]), [[false, [], []], [false, [], []]], 'both have nothing left to save, and neither was told of a conflict');
  // The room cannot be reached: nothing is saved by itself, the state says why, and Save still saves.
  const away = room(2);
  away.sides[0].session.state = 'offline';
  away.sides[0].session.leads = false;
  away.sides[0].controller.changed();
  away.sides[0].sync.localChanged();
  await away.pass(60000);
  assert.deepEqual([away.server.puts.length, away.sides[0].controller.state.autoWaits, away.sides[0].controller.dirty], [0, 'offline', true], 'offline: not saved by itself, however long; said; and still to be saved');
  assert.equal(away.sides[0].states.indexOf('offline') >= 0, true, 'whoever listens to the state was told');
  assert.equal(await away.sides[0].controller.saveNow(), true, 'Save saves');
});

test('a participant whose own file is never the very file the others write (a content control that lost its id one way here and another there): it saves its work itself, once, and nothing is left to save; no save goes round', async () => {
  const r = room(2, { auto: [true, true], apart: [1] });
  r.edit(1, 'typed by b');
  await r.pass(60000);
  assert.deepEqual(r.stored(), ['typed by b'], 'the work is in the file');
  const puts = r.server.puts.length;
  assert.ok(puts >= 1 && puts <= 2, `saved, by the saver and at most once by itself: ${puts}`);
  assert.deepEqual(r.sides.map((side) => [side.controller.dirty, side.controller.state.autoWaits, side.errors, side.asked]), [[false, undefined, [], []], [false, undefined, [], []]], 'nobody has anything left to save, waits, or was told of a conflict');
  await r.pass(600000);
  assert.equal(r.server.puts.length, puts, 'and nothing is saved again while nothing changes');
});

const RUNS = Math.max(1, Math.min(Number(process.env.COLLAB_SEEDS) || 150, 100000));

test(`whenever a participant is told it has nothing to save, the stored file holds everything it typed - also with someone in the room who writes old files and says they are new (${RUNS} runs)`, async () => {
  let cleaned = 0;
  let lies = 0;
  for (let seed = 1; seed <= RUNS; seed++) {
    let state = seed * 2654435761;
    const random = (): number => {
      state = (state + 0x6d2b79f5) >>> 0;
      let t = state;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const typed = new Map<string, string[]>();
    const bad: string[] = [];
    const count = 2 + (seed % 3);
    // The last one lies (in half of the runs); it has no auto-save of its own, and leads in some runs.
    const liar = seed % 2 === 0 ? count - 1 : -1;
    const r: ReturnType<typeof room> = room(count, {
      auto: Array.from({ length: count }, (_unused, index) => index !== liar && random() < 0.8),
      // What a participant's own save put into the file was saved, whatever becomes of the file afterwards.
      put: (name, texts) => typed.set(name, (typed.get(name) || []).filter((text) => texts.indexOf(text) < 0)),
      clean: (name) => {
        cleaned++;
        const have = r ? r.stored() : [];
        // What it typed since it last had nothing to save. (What was saved once and written over later by someone who
        // can write to where the file is stored is beyond what a room can guard.)
        for (const text of typed.get(name) || []) if (have.indexOf(text) < 0) bad.push(`seed ${seed}: ${name} was told it has nothing to save, and "${text}" is not in the stored file`);
        typed.set(name, []);
      },
    });
    if (liar >= 0 && random() < 0.5) {
      r.sides[0].session.leads = false;
      r.sides[liar].session.leads = true;
    }
    for (let tick = 0; tick < 40; tick++) {
      const roll = random();
      const who = Math.floor(random() * count);
      if (roll < 0.45) {
        if (who === liar) continue;
        const text = `s${seed}-${tick}`;
        typed.set(r.sides[who].name, (typed.get(r.sides[who].name) || []).concat(text));
        r.edit(who, text);
      } else if (roll < 0.6 && liar >= 0) {
        // An old file under a new name, said to be the file of the step everybody is at.
        lies++;
        r.server.etag = `"lie-${seed}-${tick}"`;
        r.server.body = r.fileOf(r.stored().slice(0, Math.floor(random() * (r.stored().length + 1))), r.sides[0].session.rev);
        // What was typed and is no longer in the file is unsaved again as far as this test knows; who was clean before is not looked at.
        for (const side of r.sides) if (side.name !== r.sides[liar].name) void side.sync.saved({ versions: [`e:${r.server.etag}`], rev: side.session.rev });
      } else if (roll < 0.7) {
        if (who !== liar) void r.sides[who].controller.saveNow();
      } else await r.pass(Math.floor(random() * 6000));
    }
    await r.pass(60000);
    assert.deepEqual(bad, [], 'nobody was told so before the file held their work');
      // Among honest people nobody is asked about a conflict: every save that meets another is one of the room.
    if (liar < 0) for (const side of r.sides) assert.deepEqual([side.asked, side.errors], [[], []], `seed ${seed}: ${side.name} was never asked about, or told of, a conflict with a save of the room`);
    // And whoever can save by itself has, in the end, nothing left to save.
    if (liar < 0) for (const side of r.sides) if (side.controller.autoSave) assert.equal(side.controller.dirty, false, `seed ${seed}: ${side.name} has nothing left to save in the end`);
  }
  assert.ok(cleaned > RUNS && lies > RUNS / 4, `participants were told they are saved (${cleaned}) and lies were told (${lies})`);
});

test('after the room could not be reached for a while, those who waited do not all save at once; nobody is stopped, and everybody ends with nothing to save', async () => {
  for (const count of [2, 4, 8, 12, 16, 64]) {
    const r = room(count);
    // Everybody typed; then the relay is gone for longer than anybody waits.
    for (let who = 0; who < count; who++) r.edit(who, `typed by ${who}`);
    for (const side of r.sides) {
      side.session.state = 'offline';
      side.sync.roomChanged();
    }
    await r.pass(120000);
    assert.equal(r.server.puts.length, 0, `${count}: nothing is saved while the room cannot be reached`);
    for (const side of r.sides) {
      side.session.state = 'online';
      side.sync.roomChanged();
    }
    await r.pass(180000);
    const refused = r.server.puts.filter((put) => put.status === 412).length;
    assert.ok(refused <= count, `${count}: no storm of saves that meet each other (${refused} of ${r.server.puts.length} were refused)`);
    assert.deepEqual(r.sides.filter((side) => side.controller.state.autoPaused).map((side) => side.name), [], `${count}: auto-save is stopped for nobody`);
    assert.deepEqual(r.sides.filter((side) => side.controller.dirty).map((side) => side.name), [], `${count}: everybody has nothing left to save`);
    assert.deepEqual(r.sides.reduce((all: string[], side) => all.concat(side.errors), []).filter((code) => code !== 'conflict'), [], `${count}: no other error`);
    assert.deepEqual(r.stored(), r.sides[0].book, `${count}: the stored file is the workbook`);
    if (process.env.COLLAB_STATS) console.log(`STORM ${count}: ${r.server.puts.length} PUT, ${refused} refused`);
  }
});

test('work that was proven to be saved is saved like anybody\'s from then on: the file is not fetched again for every save of the room', async () => {
  const r = room(3);
  const [, b, c] = r.sides;
  // b types once; the saver saves; b looks at the file, once, and has nothing to save.
  r.edit(1, 'typed by b');
  await r.pass(3000);
  assert.deepEqual([b.controller.dirty, r.server.gets.filter((who) => who === 'b').length], [false, 1], 'looked at once, and saved');
  assert.equal(b.controller.ownChanges, false, 'its work is in a file: nothing of its own is unsaved');
  // From here on the others type, for a long time: b takes the word of each save (the server confirms the name), and fetches nothing.
  for (let i = 0; i < 20; i++) {
    r.edit(i % 2 === 0 ? 0 : 2, `more ${i}`);
    await r.pass(3000);
  }
  assert.equal(r.server.gets.filter((who) => who === 'b').length, 1, 'the file was fetched no further time by who typed once');
  assert.equal(b.controller.dirty, false, 'and it has nothing to save');
  assert.ok(r.server.gets.filter((who) => who === 'c').length <= 10, 'who types is the one who looks');
  // What the page sends with a save to say who it is goes with the questions as well; what is about a body or a condition does not.
  assert.deepEqual(Array.from(new Set(r.server.asked)), ['a user'], 'every HEAD and GET carried the header of the page');
});

test(`among honest people, with a server that takes its time, nobody is ever asked whether to write over a save of the room (${RUNS * 4} runs); a change from outside the room still is`, async () => {
  let saves = 0;
  let met = 0;
  for (let seed = 1; seed <= RUNS * 4; seed++) {
    let state = seed * 40503;
    const random = (): number => {
      state = (state + 0x6d2b79f5) >>> 0;
      let t = state;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const count = 2 + (seed % 3);
    const r = room(count, { auto: Array.from({ length: count }, () => random() < 0.7), latency: () => 10 + Math.floor(random() * 390) });
    for (let tick = 0; tick < 30; tick++) {
      const roll = random();
      const who = Math.floor(random() * count);
      if (roll < 0.45) r.edit(who, `s${seed}-${tick}`);
      else if (roll < 0.65) void r.sides[who].controller.saveNow();
      else await r.pass(Math.floor(random() * 3000));
    }
    await r.pass(120000);
    saves += r.server.puts.length;
    met += r.server.puts.filter((put) => put.status === 412).length;
    for (const side of r.sides) assert.deepEqual(side.asked, [], `seed ${seed}: ${side.name} was asked about a conflict with a save of the room (puts ${r.server.puts.map((put) => `${put.by}${put.status}`).join(' ')})`);
  }
  assert.ok(met > 0 && saves > RUNS * 4, `saves met each other (${met} of ${saves}), and nobody was asked`);
  // Someone outside the room writes the file: that is a conflict, and the user who saves is asked, soon.
  const r = room(2, { auto: [false, false] });
  r.edit(1, 'mine');
  r.server.etag = '"written elsewhere"';
  const saving = r.sides[1].controller.saveNow();
  await r.pass(6000);
  assert.deepEqual(r.sides[1].asked, ['asked'], 'asked, within a few seconds');
  r.sides[1].controller.dismissConflict();
  assert.equal(await saving, false, 'and nothing is written over what was written elsewhere');
});

test('a connection that comes and goes does not keep work unsaved: the wait goes on through short breaks; and those who only read are not waited for', async () => {
  // The saver has auto-save off; the other loses the room for a second every ten.
  const r = room(2, { auto: [false, true] });
  const b = r.sides[1];
  r.edit(1, 'typed by b');
  for (let i = 0; i < 12 && r.server.puts.length === 0; i++) {
    await r.pass(9000);
    b.session.state = 'offline';
    b.sync.roomChanged();
    await r.pass(1000);
    b.session.state = 'online';
    b.sync.roomChanged();
  }
  await r.pass(5000);
  assert.deepEqual(r.stored(), ['typed by b'], 'saved, though the room was lost every ten seconds');
  assert.ok(r.server.puts.length >= 1 && b.controller.dirty === false, 'by the one who waited');
  // The spread among those who wait is of those who may stand in, and half a minute at most.
  const { WATCH_MIN_MS, WATCH_SPREAD_MAX_MS, createSaveSync: make } = await import('../src/collab/save-sync');
  for (const [people, most] of [[2, WATCH_MIN_MS * 1.25], [5, WATCH_MIN_MS + 12000], [500, WATCH_MIN_MS + WATCH_SPREAD_MAX_MS]] as [number, number][]) {
    let clock = 0;
    let due: { run: () => void; at: number } | undefined;
    let rearmed = -1;
    const controller = { currentETag: '"x"', dirty: true, ownChanges: true, adopt: () => undefined, hint: () => undefined, changed: () => undefined, saveQuietly: async () => false, rearm: () => void (rearmed = rearmed < 0 && due === undefined ? clock : rearmed), autoSave: { delayMs: 1000 }, changes: 0, saving: false };
    const sync = make({ controller, session: () => ({ state: 'online', rev: 1, pending: 0, leads: false }), target: undefined, strong: (tag) => tag || undefined, readOnly: () => false, people: () => people, random: () => 0.999, now: () => clock, timer: (run, ms) => ((due = { run, at: clock + ms }), () => (due = undefined)) });
    sync.localChanged();
    assert.ok(due !== undefined && (due as { at: number }).at <= most + 1 && (due as { at: number }).at > WATCH_MIN_MS, `${people} who may write: the longest wait is ${due ? (due as { at: number }).at : '?'} ms, no more than ${most}`);
  }
});
