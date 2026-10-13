// Collaborative editing, around the session: the WebSocket transport (a connection that
// comes and goes), the example relay of server/ with two participants on real sockets, and
// what the save controller does when someone else saved the file.

import assert from 'node:assert/strict';
import { createHash, randomBytes } from 'node:crypto';
import net from 'node:net';
import { test } from 'node:test';
import { SaveController, SAVE_TIMEOUT_MS, type SaveHost, type SaveState, type SaveTarget } from '../src/embed/save-to';
import { getCell, open } from '../src/index';
import { contentVersion } from '../src/collab/protocol';
import { CollabSession, type CollabNotice } from '../src/collab/session';
import { isSocketUrl, socketTransport } from '../src/collab/transport';
import { setCellInput } from '../src/ops/commands';
import { STRINGS_EN } from '../src/ui/strings';
import { canon, startFile } from './collab-fixtures';
// @ts-expect-error The relay is plain JavaScript.
import { createCollabServer } from '../server/collab-server.mjs';

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
async function until(what: () => boolean, label: string, ms = 5000): Promise<void> {
  const end = Date.now() + ms;
  while (!what()) {
    if (Date.now() > end) assert.fail(`Never happened: ${label}`);
    await wait(5);
  }
}

// ------------------------------------------------------------ the transport

/** What a browser's WebSocket looks like to the transport; the test opens it, closes it and reads what was sent. */
class FakeSocket {
  static made: FakeSocket[] = [];
  readyState = 0;
  sent: string[] = [];
  closed = false;
  onopen: ((event: unknown) => void) | null = null;
  onclose: ((event: unknown) => void) | null = null;
  onerror: ((event: unknown) => void) | null = null;
  onmessage: ((event: { data: unknown }) => void) | null = null;
  constructor(readonly url: string) {
    FakeSocket.made.push(this);
  }
  send(data: string): void {
    this.sent.push(data);
  }
  close(): void {
    this.closed = true;
    this.readyState = 3;
  }
  accept(): void {
    this.readyState = 1;
    if (this.onopen) this.onopen({});
  }
  drop(): void {
    this.readyState = 3;
    if (this.onclose) this.onclose({});
  }
}

test('an address of a relay is ws:// or wss://, and nothing else', () => {
  assert.equal(isSocketUrl('wss://example.com/collab/42'), true, 'wss');
  assert.equal(isSocketUrl('ws://127.0.0.1:8787/x'), true, 'ws');
  for (const bad of ['https://example.com/x', 'javascript:alert(1)', '', 'wss://', 'ws://a b', 5, null, undefined]) assert.equal(isSocketUrl(bad), false, `not one: ${String(bad)}`);
  assert.throws(() => socketTransport('https://example.com/x', { WebSocket: FakeSocket }), TypeError, 'the transport refuses it as well');
});

test('the socket transport says when it is open, passes text on, and connects again after the connection is lost', async () => {
  FakeSocket.made = [];
  const states: string[] = [];
  const got: string[] = [];
  const transport = socketTransport('ws://relay.example/room', { WebSocket: FakeSocket, minDelayMs: 5, maxDelayMs: 20, pingMs: 10, idleMs: 60, random: () => 1 });
  transport.onMessage((text) => got.push(text));
  assert.equal(FakeSocket.made.length, 0, 'nothing connects before someone listens to the state');
  transport.onState!((state) => states.push(state));
  assert.equal(FakeSocket.made.length, 1, 'one socket');
  transport.send('lost');
  assert.deepEqual(FakeSocket.made[0].sent, [], 'what is sent before the socket is open is dropped');
  FakeSocket.made[0].accept();
  assert.deepEqual(states, ['open'], 'open');
  transport.send('hello');
  FakeSocket.made[0].onmessage!({ data: 'text' });
  FakeSocket.made[0].onmessage!({ data: new Uint8Array(3) });
  assert.deepEqual([FakeSocket.made[0].sent, got], [['hello'], ['text']], 'text goes both ways; what is not text is not passed on');
  await until(() => FakeSocket.made[0].sent.indexOf('{"t":"ping"}') >= 0, 'a ping was sent');

  FakeSocket.made[0].drop();
  assert.deepEqual(states, ['open', 'closed'], 'closed');
  await until(() => FakeSocket.made.length === 2, 'the transport connected again');
  FakeSocket.made[1].drop();
  await until(() => FakeSocket.made.length === 3, 'and again after a failure before the socket was open');
  assert.deepEqual(states, ['open', 'closed'], 'a socket that never opened changes no state');
  FakeSocket.made[2].accept();
  assert.deepEqual(states, ['open', 'closed', 'open'], 'open again');

  // A relay that says nothing any more is taken to be gone.
  await until(() => states.length === 4, 'the silent connection was given up');
  assert.equal(FakeSocket.made[2].closed, true, 'its socket was closed');
  await until(() => FakeSocket.made.length === 4, 'and another one made');

  transport.suspend();
  const count = FakeSocket.made.length;
  await wait(60);
  assert.equal(FakeSocket.made.length, count, 'a suspended transport does not connect');
  transport.resume();
  assert.equal(FakeSocket.made.length, count + 1, 'resume connects at once');
  transport.close();
  await wait(60);
  assert.equal(FakeSocket.made.length, count + 1, 'a closed transport stays closed');
  assert.equal(FakeSocket.made[count].closed, true, 'and its socket is closed');
});

// ---------------------------------------------- the example relay, for real

/** A WebSocket client on a plain socket, as far as the transport needs one: text frames, masked, and what the relay sends. */
class MiniSocket {
  readyState = 0;
  onopen: ((event: unknown) => void) | null = null;
  onclose: ((event: unknown) => void) | null = null;
  onerror: ((event: unknown) => void) | null = null;
  onmessage: ((event: { data: unknown }) => void) | null = null;
  private readonly socket: net.Socket;
  private buffer = Buffer.alloc(0);
  private upgraded = false;
  private parts: Buffer[] = [];

  constructor(url: string) {
    const target = new URL(url);
    const key = randomBytes(16).toString('base64');
    this.socket = net.connect(Number(target.port), target.hostname, () => {
      this.socket.write(`GET ${target.pathname} HTTP/1.1\r\nHost: ${target.host}\r\nUpgrade: websocket\r\nConnection: Upgrade\r\nSec-WebSocket-Key: ${key}\r\nSec-WebSocket-Version: 13\r\n\r\n`);
    });
    this.socket.on('data', (chunk) => {
      this.buffer = Buffer.concat([this.buffer, chunk]);
      if (!this.upgraded) {
        const end = this.buffer.indexOf('\r\n\r\n');
        if (end < 0) return;
        const head = this.buffer.subarray(0, end).toString('latin1');
        const accept = createHash('sha1').update(`${key}258EAFA5-E914-47DA-95CA-C5AB0DC85B11`).digest('base64');
        if (!/^HTTP\/1\.1 101/.test(head) || head.toLowerCase().indexOf(`sec-websocket-accept: ${accept.toLowerCase()}`) < 0) return this.socket.destroy();
        this.buffer = this.buffer.subarray(end + 4);
        this.upgraded = true;
        this.readyState = 1;
        if (this.onopen) this.onopen({});
      }
      this.frames();
    });
    this.socket.on('close', () => {
      this.readyState = 3;
      if (this.onclose) this.onclose({});
    });
    this.socket.on('error', () => undefined);
  }

  private frames(): void {
    for (;;) {
      if (this.buffer.length < 2) return;
      const fin = (this.buffer[0] & 0x80) !== 0;
      const opcode = this.buffer[0] & 0x0f;
      let length = this.buffer[1] & 0x7f;
      let at = 2;
      if (length === 126) {
        if (this.buffer.length < 4) return;
        length = this.buffer.readUInt16BE(2);
        at = 4;
      } else if (length === 127) {
        if (this.buffer.length < 10) return;
        length = Number(this.buffer.readBigUInt64BE(2));
        at = 10;
      }
      if (this.buffer.length < at + length) return;
      const payload = this.buffer.subarray(at, at + length);
      this.buffer = this.buffer.subarray(at + length);
      if (opcode === 9) this.frame(10, payload);
      else if (opcode === 8) this.socket.end();
      else if (opcode === 0 || opcode === 1) {
        this.parts.push(payload);
        if (fin) {
          const text = Buffer.concat(this.parts).toString('utf8');
          this.parts = [];
          if (this.onmessage) this.onmessage({ data: text });
        }
      }
    }
  }

  private frame(opcode: number, payload: Buffer): void {
    const mask = randomBytes(4);
    const head = payload.length < 126 ? Buffer.from([0x80 | opcode, 0x80 | payload.length]) : payload.length < 65536 ? Buffer.from([0x80 | opcode, 0x80 | 126, payload.length >> 8, payload.length & 255]) : Buffer.concat([Buffer.from([0x80 | opcode, 0x80 | 127]), (() => {
      const size = Buffer.alloc(8);
      size.writeBigUInt64BE(BigInt(payload.length));
      return size;
    })()]);
    const masked = Buffer.from(payload);
    for (let i = 0; i < masked.length; i++) masked[i] ^= mask[i & 3];
    this.socket.write(Buffer.concat([head, mask, masked]));
  }

  send(data: string): void {
    this.frame(1, Buffer.from(data, 'utf8'));
  }

  close(): void {
    this.readyState = 3;
    this.socket.destroy();
  }
}

test('two participants on the example relay, over real sockets, end with the same workbook; one that drops out catches up', async () => {
  const relay = createCollabServer({ pingIntervalMs: 60000 });
  const { port } = await relay.listen(0, '127.0.0.1');
  const bytes = await startFile();
  try {
    const members = [];
    for (const name of ['a', 'b']) {
      const wb = await (await open(bytes)).workbook();
      const notices: CollabNotice[] = [];
      const transport = socketTransport(`ws://127.0.0.1:${port}/files/7`, { WebSocket: MiniSocket, minDelayMs: 10, maxDelayMs: 50 });
      const session = new CollabSession({ workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice) }, { transport, versions: [contentVersion(bytes)], user: { name }, token: 'a secret nobody logs' });
      members.push({ wb, session, transport, notices });
    }
    const [a, b] = members;
    await until(() => a.session.state === 'online' && b.session.state === 'online', 'both are online');
    await until(() => a.session.peers.length === 1 && b.session.peers.length === 1, 'and see each other');
    assert.equal(a.session.peers[0].name, 'b', 'by name');
    for (let i = 0; i < 20; i++) {
      a.session.submit(setCellInput(a.wb, 0, 12, 0, `a${i}`));
      b.session.submit(setCellInput(b.wb, 0, 12, 0, `b${i}`));
      a.session.submit([{ type: 'insert', sheet: 0, axis: 'row', at: 2, count: 1 }]);
      b.session.submit(setCellInput(b.wb, 1, i, 5, `=Data!A${i + 1}`));
      if (i % 5 === 0) await wait(3);
    }
    await until(() => a.session.pending === 0 && b.session.pending === 0, 'everything was placed');
    await until(() => a.session.rev === b.session.rev, 'and both are at the same step');
    assert.equal(canon(a.wb), canon(b.wb), 'the same workbook');
    assert.equal(relay.rooms()[0].connections, 2, 'one room with both in it');

    // One drops out, both go on, and it comes back.
    b.transport.suspend();
    await until(() => b.session.state === 'offline' && a.session.peers.length === 0, 'b is gone');
    b.session.submit(setCellInput(b.wb, 1, 0, 0, 'while away'));
    a.session.submit(setCellInput(a.wb, 1, 0, 1, 'meanwhile'));
    await until(() => a.session.pending === 0, 'a went on');
    b.transport.resume();
    await until(() => b.session.state === 'online' && b.session.pending === 0 && a.session.rev === b.session.rev, 'b caught up');
    assert.equal(canon(a.wb), canon(b.wb), 'the same workbook again');
    assert.equal(getCell(a.wb.sheets[1], 0, 0)?.value, 'while away', 'with what b did while away');
    assert.deepEqual([a.notices, b.notices], [[], []], 'nothing was dropped and nothing went wrong');
    for (const member of members) member.session.leave();
  } finally {
    await relay.close();
  }
});

test('a relay that takes nobody ends the session, which says so and never names the token', async () => {
  const logged: unknown[] = [];
  const relay = createCollabServer({ authorize: () => false, log: (event: unknown) => logged.push(event) });
  const { port } = await relay.listen(0, '127.0.0.1');
  try {
    const wb = await (await open(await startFile())).workbook();
    const notices: CollabNotice[] = [];
    const session = new CollabSession({ workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice) }, { transport: socketTransport(`ws://127.0.0.1:${port}/x`, { WebSocket: MiniSocket }), versions: ['v'], token: 'tok-12345' });
    await until(() => session.state === 'closed', 'the session is over');
    assert.deepEqual(notices, [{ code: 'refused' }], 'and it was said');
    assert.equal(JSON.stringify([notices, logged]).indexOf('tok-12345'), -1, 'the token is in no notice and no log');
  } finally {
    await relay.close();
  }
});

// ------------------------------------------------- the save and the others

const settle = async () => {
  for (let i = 0; i < 30; i++) await Promise.resolve();
};

function saver(etag: string): { controller: SaveController; sent: { headers: { [name: string]: string }; answer(status: number, headers?: { [name: string]: string }): void }[]; states: SaveState[]; errors: unknown[]; asked: number } {
  const out = { controller: undefined as unknown as SaveController, sent: [] as { headers: { [name: string]: string }; answer(status: number, headers?: { [name: string]: string }): void }[], states: [] as SaveState[], errors: [] as unknown[], asked: 0 };
  const target: SaveTarget = { kind: 'request', url: 'https://app.example.com/files/7', method: 'PUT', headers: [], credentials: 'same-origin', sameAsFile: true, timeoutMs: SAVE_TIMEOUT_MS, allowConverted: true, onConflict: 'ask' };
  const host: SaveHost = {
    target,
    writable: () => true,
    commit: () => true,
    bytes: async () => new Uint8Array([1, 2, 3]),
    info: () => ({ format: 'xlsx' }),
    strings: () => STRINGS_EN,
    onState: (state) => out.states.push(state),
    onError: (error) => out.errors.push(error),
    askOverwrite: () => {
      out.asked++;
    },
    etag,
    fetch: ((_url: unknown, init?: RequestInit) =>
      new Promise<Response>((resolve) => {
        out.sent.push({
          headers: (init && init.headers) as { [name: string]: string },
          answer: (status, headers = {}) => resolve({ status, ok: status >= 200 && status < 300, headers: { get: (name: string) => { for (const key of Object.keys(headers)) if (key.toLowerCase() === name.toLowerCase()) return headers[key]; return null; } }, body: { cancel: async () => undefined } } as unknown as Response),
        });
      })) as typeof fetch,
    timer: () => () => undefined,
  };
  out.controller = new SaveController(host);
  return out;
}

test('a save of someone else: its ETag is what the next save here is sent with, and a workbook that is what was saved has nothing to save', async () => {
  const r = saver('"v1"');
  r.controller.changed(true);
  assert.equal(r.controller.dirty, true, 'what someone else changed is there to be saved');
  r.controller.adopt('"v2"', true);
  assert.equal(r.controller.dirty, false, 'the file that was saved is this one');
  assert.equal(r.controller.currentETag, '"v2"', 'and it is known by its new ETag');
  r.controller.changed();
  r.controller.adopt('"v3"', false);
  assert.equal(r.controller.dirty, true, 'a save that does not hold what was changed here leaves it to be saved');
  const done = r.controller.saveNow();
  await settle();
  assert.equal(r.sent[0].headers['If-Match'], '"v3"', 'the save is guarded by the ETag the other one named');
  r.sent[0].answer(204, { ETag: '"v4"' });
  assert.equal(await done, true, 'saved');
  r.controller.adopt('W/"weak"', false);
  assert.equal(r.controller.currentETag, '"v4"', 'an ETag that If-Match cannot be sent with is not taken');
});

test('a conflict with a file that someone else saved while this save was on its way is no conflict: the save is sent again', async () => {
  const r = saver('"v1"');
  r.controller.changed();
  const done = r.controller.saveNow();
  await settle();
  assert.equal(r.sent[0].headers['If-Match'], '"v1"', 'sent with what the file was known by');
  r.controller.adopt('"v2"', false);
  r.sent[0].answer(412);
  await settle();
  assert.equal(r.sent.length, 2, 'sent again');
  assert.equal(r.sent[1].headers['If-Match'], '"v2"', 'with the ETag of the save that got there first');
  r.sent[1].answer(204, { ETag: '"v3"' });
  assert.equal(await done, true, 'and taken');
  assert.deepEqual([r.errors, r.asked], [[], 0], 'no error was told and nobody was asked');

  // A conflict nobody announced is still one.
  r.controller.changed();
  const again = r.controller.saveNow();
  await settle();
  r.sent[2].answer(412);
  await settle();
  assert.equal(r.asked, 1, 'the user is asked');
  r.controller.adopt('"v9"', false);
  assert.equal(r.controller.asking, true, 'the question stays open when someone else saved: whether to write over a file is the user\'s to say');
  r.controller.dismissConflict();
  assert.equal(await again, false, 'the save did not happen');
});
