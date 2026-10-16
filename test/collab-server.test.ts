import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import http from 'node:http';
import net from 'node:net';
import { test } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createCollabServer, type CollabServerOptions } from '../server/collab-server.mjs';
import { PROTOCOL_VERSION } from '../src/collab/room';
import { CollabSession } from '../src/collab/session';
import { open } from '../src/index';
import { startFile } from './collab-fixtures';

// The example relay of collaborative editing: the WebSocket it speaks by hand
// (handshake, frames, limits) and the contract of a room (join, op, nack,
// reset, saved, presence). The client below is as raw as the server: it can say
// things no honest client says.

const SERVER_FILE = fileURLToPath(new URL('../server/collab-server.mjs', import.meta.url));

/** A guard against a hang, not a measure of speed. */
const WAIT = 20000;

interface Frame {
  fin: boolean;
  opcode: number;
  masked: boolean;
  payload: Buffer;
}

interface FrameShape {
  fin?: boolean;
  mask?: boolean;
  rsv?: number;
  /** The form of the length, whatever the payload needs. */
  form?: 7 | 16 | 64;
  /** The length the header says, when it is not the payload's. */
  declared?: number | bigint;
}

function frameBytes(opcode: number, payload: Buffer | string = '', shape: FrameShape = {}): Buffer {
  const body = Buffer.from(payload);
  const length = shape.declared ?? body.length;
  const form = shape.form ?? (length < 126 ? 7 : length < 65536 ? 16 : 64);
  let head: Buffer;
  if (form === 7) {
    head = Buffer.from([0, Number(length)]);
  } else if (form === 16) {
    head = Buffer.alloc(4);
    head[1] = 126;
    head.writeUInt16BE(Number(length), 2);
  } else {
    head = Buffer.alloc(10);
    head[1] = 127;
    head.writeBigUInt64BE(BigInt(length), 2);
  }
  head[0] = (shape.fin === false ? 0 : 0x80) | ((shape.rsv ?? 0) << 4) | opcode;
  if (shape.mask === false) return Buffer.concat([head, body]);
  head[1] |= 0x80;
  const key = randomBytes(4);
  for (let i = 0; i < body.length; i++) body[i] ^= key[i & 3];
  return Buffer.concat([head, key, body]);
}

class Client {
  readonly frames: Frame[] = [];
  /** Every text the server sent, kept for the tests that look for a leak. */
  readonly seen: string[] = [];
  closed = false;
  /** The `key` it joined with: a secret of its own, which the tests also use to tell its operations apart. */
  key = '';
  private rest: Buffer = Buffer.alloc(0);
  private wake: (() => void) | null = null;

  constructor(
    readonly socket: net.Socket,
    readonly status: number,
    readonly head: string,
  ) {}

  feed(data: Buffer): void {
    this.rest = Buffer.concat([this.rest, data]);
    for (;;) {
      if (this.rest.length < 2) break;
      let length = this.rest[1] & 127;
      let at = 2;
      if (length === 126) {
        if (this.rest.length < 4) break;
        length = this.rest.readUInt16BE(2);
        at = 4;
      } else if (length === 127) {
        if (this.rest.length < 10) break;
        length = Number(this.rest.readBigUInt64BE(2));
        at = 10;
      }
      const masked = (this.rest[1] & 0x80) !== 0;
      if (masked) at += 4;
      if (this.rest.length < at + length) break;
      const payload = Buffer.from(this.rest.subarray(at, at + length));
      const opcode = this.rest[0] & 15;
      this.frames.push({ fin: (this.rest[0] & 0x80) !== 0, opcode, masked, payload });
      if (opcode === 1) this.seen.push(payload.toString('utf8'));
      this.rest = this.rest.subarray(at + length);
    }
    this.wake?.();
  }

  markClosed(): void {
    this.closed = true;
    this.wake?.();
  }

  private async wait(what: string, ready: () => boolean, ms = WAIT): Promise<void> {
    const end = Date.now() + ms;
    while (!ready()) {
      const left = end - Date.now();
      if (left <= 0) throw new Error(`timed out waiting for ${what}`);
      await new Promise<void>((resolve) => {
        const done = (): void => {
          clearTimeout(timer);
          this.wake = null;
          resolve();
        };
        const timer = setTimeout(done, left);
        this.wake = done;
      });
    }
  }

  async frame(): Promise<Frame> {
    await this.wait('a frame', () => this.frames.length > 0 || this.closed);
    const frame = this.frames.shift();
    if (!frame) throw new Error('the connection closed before a frame arrived');
    return frame;
  }

  /** The next message; pings of the server are not messages. */
  async json(): Promise<any> {
    const frame = await this.frame();
    assert.equal(frame.opcode, 1, `a text frame was expected, opcode ${frame.opcode} came`);
    assert.equal(frame.fin, true, 'the server does not fragment');
    return JSON.parse(frame.payload.toString('utf8'));
  }

  /** The close frame the server sends, after the `error` message when one is named. */
  async closes(code: number, error?: string): Promise<void> {
    if (error !== undefined) assert.deepEqual(await this.json(), { t: 'error', code: error });
    const frame = await this.frame();
    assert.equal(frame.opcode, 8, `a close frame was expected, opcode ${frame.opcode} came`);
    assert.equal(frame.payload.readUInt16BE(0), code, 'the status code of the close frame');
    await this.gone();
  }

  gone(): Promise<void> {
    return this.wait('the socket to close', () => this.closed);
  }

  /** The id of an operation of this connection. */
  id(name: string): string {
    return `${this.key}:${name}`;
  }

  send(message: unknown): void {
    this.socket.write(frameBytes(1, JSON.stringify(message)));
  }

  bytes(data: Buffer): void {
    this.socket.write(data);
  }

  /** `ping` is answered in order, so a `pong` proves everything before it was read. */
  async settled(): Promise<void> {
    this.send({ t: 'ping' });
    assert.deepEqual(await this.json(), { t: 'pong' }, 'nothing else was to arrive before the pong');
  }

  end(): void {
    this.socket.destroy();
  }
}

interface ConnectOptions {
  path?: string;
  /** Added to, or with `null` taken from, the headers of an honest handshake. */
  headers?: Record<string, string | null>;
}

function connect(port: number, options: ConnectOptions = {}): Promise<Client> {
  return new Promise((resolve, reject) => {
    const socket = net.connect(port, '127.0.0.1');
    const headers: Record<string, string | null> = {
      Host: `127.0.0.1:${port}`,
      Upgrade: 'websocket',
      Connection: 'Upgrade',
      'Sec-WebSocket-Key': randomBytes(16).toString('base64'),
      'Sec-WebSocket-Version': '13',
      ...options.headers,
    };
    let got: Buffer = Buffer.alloc(0);
    let client: Client | null = null;
    socket.on('connect', () => {
      let text = `GET ${options.path ?? '/doc'} HTTP/1.1\r\n`;
      for (const name of Object.keys(headers)) {
        if (headers[name] !== null) text += `${name}: ${headers[name]}\r\n`;
      }
      socket.write(text + '\r\n');
    });
    socket.on('data', (data: Buffer) => {
      if (client) return client.feed(data);
      got = Buffer.concat([got, data]);
      const end = got.indexOf('\r\n\r\n');
      if (end < 0) return;
      const head = got.subarray(0, end).toString('latin1');
      client = new Client(socket, Number(head.slice(9, 12)), head);
      client.feed(got.subarray(end + 4));
      resolve(client);
    });
    socket.on('error', (error) => {
      if (!client) reject(error);
    });
    socket.on('close', () => {
      if (client) client.markClosed();
      else reject(new Error('the connection closed before the handshake was answered'));
    });
  });
}

type Server = ReturnType<typeof createCollabServer>;

/** A server that keeps what it has to say to itself, unless the test listens. */
async function withServer(
  options: CollabServerOptions,
  body: (port: number, server: Server) => Promise<void>,
): Promise<void> {
  const server = createCollabServer({ log: () => undefined, ...options });
  const { port } = await server.listen(0, '127.0.0.1');
  try {
    await body(port, server);
  } finally {
    await server.close();
  }
}

let keys = 0;

/** Joins with a key of its own, unless the join names one (`key: undefined` joins without). */
async function joined(port: number, join: object = { v: ['v1'] }, path = '/doc'): Promise<{ c: Client; hello: any }> {
  const c = await connect(port, { path });
  const message: { key?: string } = { key: `c${++keys}`, ...join };
  c.key = message.key ?? '';
  c.send({ t: 'join', ...message });
  const hello = await c.json();
  assert.equal(hello.t, 'hello', `a hello was expected, ${JSON.stringify(hello)} came`);
  return { c, hello };
}

async function until(what: string, ready: () => boolean, ms = WAIT): Promise<void> {
  const end = Date.now() + ms;
  while (!ready()) {
    if (Date.now() > end) throw new Error(`timed out waiting for ${what}`);
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

/** The name of a room, as `hello` tells it and as a `join` may say it. */
const EPOCH = /^[A-Za-z0-9_-]{1,64}$/;

/**
 * Joins as a client does whose file says where it was saved: with the revision and with the name the room had
 * then. Only to read, so that nobody is asked for anything. The `hello` it gets; its `v` are those of `v` the
 * room has heard of.
 */
async function comesBack(port: number, ep: string, since: number, v: string[], path = '/doc'): Promise<any> {
  const { hello } = await joined(port, { v, since, ep, ro: true }, path);
  assert.equal(hello.from, since, `a join with the name of the room is caught up from the revision it names: ${JSON.stringify(hello)}`);
  assert.equal(hello.ep, ep, 'the room is the one that was named');
  return hello;
}

/** Sends `count` operations one after another and waits for each to come back. */
async function operations(c: Client, from: number, count: number, ops: unknown[] = [1]): Promise<void> {
  for (let i = 0; i < count; i++) {
    c.send({ t: 'op', base: from + i, id: c.id(`op${from + i}`), ops });
    const back = await c.json();
    assert.equal(back.t, 'op', `the operation was to come back, ${JSON.stringify(back).slice(0, 200)} came`);
    assert.equal(back.rev, from + i + 1, 'the revision the operation made');
  }
}

// --- Handshake ---------------------------------------------------------------

test('the handshake answers with the accept value of RFC 6455 and takes no subprotocol or extension', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port, {
      headers: {
        'Sec-WebSocket-Key': 'dGhlIHNhbXBsZSBub25jZQ==',
        'Sec-WebSocket-Protocol': 'chat',
        'Sec-WebSocket-Extensions': 'permessage-deflate',
      },
    });
    assert.equal(c.status, 101);
    assert.match(c.head, /\r\nSec-WebSocket-Accept: s3pPLMBiTxaQ9kYGzzhZRbK\+xOo=(\r\n|$)/);
    assert.match(c.head, /\r\nUpgrade: websocket(\r\n|$)/i);
    assert.match(c.head, /\r\nConnection: Upgrade(\r\n|$)/i);
    assert.doesNotMatch(c.head, /Sec-WebSocket-Protocol/i);
    assert.doesNotMatch(c.head, /Sec-WebSocket-Extensions/i);
    c.end();
  });
});

test('a handshake with a missing or wrong header is answered with 400', async () => {
  await withServer({}, async (port) => {
    const cases: Record<string, Record<string, string | null>> = {
      'no key': { 'Sec-WebSocket-Key': null },
      'a key that is not sixteen bytes': { 'Sec-WebSocket-Key': 'c2hvcnQ=' },
      'no version': { 'Sec-WebSocket-Version': null },
      'version 8': { 'Sec-WebSocket-Version': '8' },
      'another protocol': { Upgrade: 'h2c' },
    };
    for (const name of Object.keys(cases)) {
      const c = await connect(port, { headers: cases[name] });
      assert.equal(c.status, 400, name);
      await c.gone();
    }
    // The server is still there.
    const { c } = await joined(port);
    c.end();
  });
});

test('a plain HTTP request to the server of listen() is answered with 426', async () => {
  await withServer({}, async (port) => {
    const status = await new Promise<number>((resolve, reject) => {
      http
        .get({ host: '127.0.0.1', port, path: '/doc', agent: false }, (response) => {
          response.resume();
          resolve(response.statusCode ?? 0);
        })
        .on('error', reject);
    });
    assert.equal(status, 426);
  });
});

test('with a list of origins, only those are let in', async () => {
  await withServer({ origins: ['https://app.example'] }, async (port) => {
    const allowed = await connect(port, { headers: { Origin: 'https://app.example' } });
    assert.equal(allowed.status, 101, 'the origin on the list');
    allowed.end();
    const other = await connect(port, { headers: { Origin: 'https://evil.example' } });
    assert.equal(other.status, 403, 'an origin that is not on the list');
    await other.gone();
    const none = await connect(port);
    assert.equal(none.status, 403, 'no Origin header');
    await none.gone();
  });
  await withServer({}, async (port) => {
    const any = await connect(port, { headers: { Origin: 'https://evil.example' } });
    assert.equal(any.status, 101, 'without a list every origin is let in');
    any.end();
  });
});

test('with a path, only the paths under it are upgraded', async () => {
  await withServer({ path: '/collab/' }, async (port) => {
    const inside = await connect(port, { path: '/collab/doc?x=1' });
    assert.equal(inside.status, 101);
    inside.end();
    const outside = await connect(port, { path: '/other' });
    assert.equal(outside.status, 404);
    await outside.gone();
  });
});

test('attach() serves the upgrades of a server that already exists and close() lets go of it', async () => {
  const web = http.createServer((_request, response) => response.end('page'));
  await new Promise<void>((resolve) => web.listen(0, '127.0.0.1', resolve));
  const port = (web.address() as net.AddressInfo).port;
  const server = createCollabServer();
  try {
    server.attach(web);
    const { c, hello } = await joined(port);
    assert.equal(hello.rev, 0);
    await server.close();
    await c.gone();
    assert.equal(web.listenerCount('upgrade'), 0, 'the listener is taken off');
    assert.equal(web.listening, true, 'the server of the application is not closed');
  } finally {
    await server.close();
    await new Promise((resolve) => web.close(resolve));
  }
});

// --- Frames ------------------------------------------------------------------

test('a frame that is not masked closes the connection with 1002', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    c.bytes(frameBytes(1, '{"t":"join","v":[]}', { mask: false }));
    await c.closes(1002);
  });
});

test('reserved bits, unknown opcodes and lengths that are not minimal close with 1002', async () => {
  await withServer({}, async (port) => {
    const cases: Record<string, Buffer> = {
      'a reserved bit': frameBytes(1, '{}', { rsv: 4 }),
      'opcode 3': frameBytes(3, '{}'),
      'opcode 11': frameBytes(11, ''),
      'a continuation of nothing': frameBytes(0, '{}'),
      'a fragmented ping': frameBytes(9, 'x', { fin: false }),
      'a ping of 126 bytes': frameBytes(9, Buffer.alloc(126)),
      '16 bits for a short length': frameBytes(1, '{}', { form: 16 }),
      '64 bits for a short length': frameBytes(1, '{}', { form: 64 }),
      'a close frame of one byte': frameBytes(8, 'x'),
    };
    for (const name of Object.keys(cases)) {
      const c = await connect(port);
      c.bytes(cases[name]);
      const frame = await c.frame();
      assert.equal(frame.opcode, 8, name);
      assert.equal(frame.payload.readUInt16BE(0), 1002, name);
      await c.gone();
    }
    const c = await connect(port);
    c.bytes(Buffer.concat([frameBytes(1, '{"t":', { fin: false }), frameBytes(1, '"ping"}')]));
    await c.closes(1002);
  });
});

test('a binary frame closes with 1003', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    c.bytes(frameBytes(2, Buffer.from([1, 2, 3])));
    await c.closes(1003);
  });
});

test('a text that is not UTF-8 closes with 1007', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    c.bytes(frameBytes(1, Buffer.from([0x7b, 0xff, 0xfe, 0x7d])));
    await c.closes(1007);
    // A sequence that is cut by the end of the message is as wrong.
    const cut = await connect(port);
    cut.bytes(frameBytes(1, Buffer.from([0x22, 0xc3])));
    await cut.closes(1007);
  });
});

test('a message comes in fragments, with a ping between them', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    const text = Buffer.from(JSON.stringify({ t: 'join', v: ['ğüşiöç-v1'] }));
    // The second cut falls inside a character of two bytes.
    const cut = text.indexOf(Buffer.from('ğ')) + 1;
    c.bytes(frameBytes(1, text.subarray(0, 5), { fin: false }));
    c.bytes(frameBytes(9, 'still here'));
    c.bytes(frameBytes(0, text.subarray(5, cut), { fin: false }));
    c.bytes(frameBytes(0, '', { fin: false }));
    c.bytes(frameBytes(0, text.subarray(cut)));
    const pong = await c.frame();
    assert.equal(pong.opcode, 10, 'the ping is answered before the message is whole');
    assert.equal(pong.payload.toString(), 'still here');
    assert.equal(pong.masked, false, 'the server does not mask');
    const hello = await c.json();
    assert.equal(hello.t, 'hello');
    assert.deepEqual(hello.v, ['ğüşiöç-v1']);
  });
});

test('a frame that arrives a byte at a time is read, and so are two frames in one write', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    const join = frameBytes(1, JSON.stringify({ t: 'join', v: ['v1'], key: 'k', user: { name: 'x'.repeat(40) } }));
    for (let i = 0; i < join.length; i++) {
      c.bytes(join.subarray(i, i + 1));
      if (i % 16 === 0) await new Promise((resolve) => setImmediate(resolve));
    }
    assert.equal((await c.json()).t, 'hello');
    c.bytes(
      Buffer.concat([
        frameBytes(1, JSON.stringify({ t: 'op', base: 0, id: 'k:a', ops: [1] })),
        frameBytes(1, JSON.stringify({ t: 'op', base: 1, id: 'k:b', ops: [2] })),
        frameBytes(1, JSON.stringify({ t: 'ping' })).subarray(0, 3),
      ]),
    );
    assert.deepEqual((await c.json()).id, 'k:a');
    assert.deepEqual((await c.json()).id, 'k:b');
  });
});

test('lengths of 16 and of 64 bits are read', async () => {
  await withServer({}, async (port) => {
    const { c, hello } = await joined(port);
    const medium = 'm'.repeat(300);
    const large = 'L'.repeat(70000);
    c.bytes(frameBytes(1, JSON.stringify({ t: 'op', base: 0, id: c.id('medium'), ops: [medium] })));
    assert.deepEqual(await c.json(), { t: 'op', rev: 1, id: c.id('medium'), from: hello.you, ops: [medium] });
    c.bytes(frameBytes(1, JSON.stringify({ t: 'op', base: 1, id: c.id('large'), ops: [large] })));
    const back = await c.frame();
    assert.equal(back.payload.length > 70000, true, 'the answer needs a length of 64 bits too');
    assert.deepEqual(JSON.parse(back.payload.toString()), { t: 'op', rev: 2, id: c.id('large'), from: hello.you, ops: [large] });
  });
});

test('a message over the limit closes with 1009 as soon as its header says so', async () => {
  await withServer({}, async (port) => {
    // Only the header is sent: the server cannot be waiting for the payload.
    const huge = await connect(port);
    huge.bytes(frameBytes(1, '', { declared: 2 * 1024 * 1024, form: 64 }));
    await huge.closes(1009);
    const absurd = await connect(port);
    absurd.bytes(frameBytes(1, '', { declared: 2n ** 62n, form: 64 }));
    await absurd.closes(1009);
    const whole = await connect(port);
    whole.bytes(frameBytes(1, Buffer.alloc(1024 * 1024 + 1, 0x20)));
    await whole.closes(1009);
  });
  await withServer({ maxMessageBytes: 100 }, async (port) => {
    const medium = await connect(port);
    medium.bytes(frameBytes(1, '', { declared: 300, form: 16 }));
    await medium.closes(1009);
    // The parts are counted together.
    const parts = await connect(port);
    parts.bytes(frameBytes(1, ' '.repeat(60), { fin: false }));
    parts.bytes(frameBytes(0, '', { declared: 41 }));
    await parts.closes(1009);
    const fits = await connect(port);
    fits.bytes(frameBytes(1, ' '.repeat(60), { fin: false }));
    fits.bytes(frameBytes(0, JSON.stringify({ t: 'join', v: [] }).padEnd(40)));
    assert.equal((await fits.json()).t, 'hello', 'a hundred bytes are within a limit of a hundred');
  });
});

test('a ping is answered with a pong that carries the same bytes', async () => {
  await withServer({}, async (port) => {
    const c = await connect(port);
    const payload = randomBytes(125);
    c.bytes(frameBytes(9, payload));
    const pong = await c.frame();
    assert.equal(pong.opcode, 10);
    assert.deepEqual(pong.payload, payload);
    c.bytes(frameBytes(9, ''));
    assert.equal((await c.frame()).payload.length, 0, 'an empty ping has an empty pong');
    // A pong nobody asked for is not an error.
    c.bytes(frameBytes(10, 'unasked'));
    c.send({ t: 'join', v: [] });
    assert.equal((await c.json()).t, 'hello');
  });
});

test('a close frame is answered with a close frame and the socket is closed', async () => {
  await withServer({}, async (port, server) => {
    const { c } = await joined(port);
    const payload = Buffer.concat([Buffer.from([0x03, 0xe8]), Buffer.from('bye')]);
    c.bytes(frameBytes(8, payload));
    const frame = await c.frame();
    assert.equal(frame.opcode, 8);
    assert.equal(frame.payload.readUInt16BE(0), 1000, 'the status code comes back');
    await c.gone();
    await until('the room to be empty', () => server.rooms()[0].connections === 0);

    const bare = await connect(port);
    bare.bytes(frameBytes(8, ''));
    const empty = await bare.frame();
    assert.equal(empty.opcode, 8);
    assert.equal(empty.payload.length, 0, 'a close without a code is answered without one');
    await bare.gone();
  });
});

test('the server pings, and a connection that does not answer is cut', async () => {
  await withServer({ pingIntervalMs: 40 }, async (port, server) => {
    const silent = (await joined(port)).c;
    const { c: lively, hello } = await joined(port);
    let pongs = 0;
    const answer = setInterval(() => {
      for (let i = lively.frames.length - 1; i >= 0; i--) {
        if (lively.frames[i].opcode !== 9) continue;
        lively.bytes(frameBytes(10, lively.frames[i].payload));
        lively.frames.splice(i, 1);
        pongs++;
      }
    }, 5);
    try {
      await silent.gone();
      assert.equal(silent.frames.filter((frame) => frame.opcode === 9).length >= 2, true, 'two pings went unanswered');
      await until('some more pings', () => pongs >= 6);
      assert.equal(lively.closed, false, 'the connection that answers stays');
      assert.equal(server.rooms()[0].connections, 1);
      const left = lively.frames.find((frame) => frame.opcode === 1 && frame.payload.includes('"left"'));
      assert.notEqual(left, undefined, 'the others are told that the silent one left');
      assert.notEqual(hello.peers[0].id, undefined);
    } finally {
      clearInterval(answer);
    }
  });
});

// --- The first message -------------------------------------------------------

test('an op before a join is a bad message', async () => {
  await withServer({}, async (port, server) => {
    const c = await connect(port);
    c.send({ t: 'op', base: 0, id: 'a', ops: [1] });
    await c.closes(1008, 'bad-message');
    const ping = await connect(port);
    ping.send({ t: 'ping' });
    await ping.closes(1008, 'bad-message');
    assert.deepEqual(server.rooms(), [], 'no room was made');
  });
});

test('a connection that does not join in time is closed', async () => {
  await withServer({ joinTimeoutMs: 60 }, async (port) => {
    const idle = await connect(port);
    await idle.closes(1008, 'join-timeout');
    // Having joined, a connection may be as quiet as it likes.
    const { c } = await joined(port);
    await new Promise((resolve) => setTimeout(resolve, 200));
    await c.settled();
  });
});

test('messages that are not what the contract says are bad messages', async () => {
  await withServer({}, async (port) => {
    const first: Record<string, string> = {
      'not JSON': '{"t":"join"',
      'an array': '[]',
      'a string': '"join"',
      null: 'null',
      'join without v': '{"t":"join"}',
      'v that is not a list': '{"t":"join","v":"v1"}',
      'nine versions': JSON.stringify({ t: 'join', v: ['1', '2', '3', '4', '5', '6', '7', '8', '9'] }),
      'an empty version': '{"t":"join","v":[""]}',
      'a version of 201 characters': JSON.stringify({ t: 'join', v: ['x'.repeat(201)] }),
      'a version that is a number': '{"t":"join","v":[1]}',
      'a negative since': '{"t":"join","v":[],"since":-1}',
      'a fractional since': '{"t":"join","v":[],"since":1.5}',
      'a since that is not safe': '{"t":"join","v":[],"since":9007199254740993}',
      'a since that is a string': '{"t":"join","v":[],"since":"0"}',
      'a token that is a number': '{"t":"join","v":[],"token":1}',
      'a token of 4097 characters': JSON.stringify({ t: 'join', v: [], token: 't'.repeat(4097) }),
      'ro that is a string': '{"t":"join","v":[],"ro":"yes"}',
      'a room of 201 characters': JSON.stringify({ t: 'join', v: [], room: 'r'.repeat(201) }),
      'a user that is a string': '{"t":"join","v":[],"user":"Ada"}',
    };
    for (const name of Object.keys(first)) {
      const c = await connect(port);
      c.bytes(frameBytes(1, first[name]));
      await c.closes(1008, 'bad-message').catch((error) => {
        throw new Error(`${name}: ${error.message}`);
      });
    }
    const later: Record<string, unknown> = {
      'an unknown type': { t: 'hello' },
      'no type': { base: 0 },
      'op without base': { t: 'op', id: 'a', ops: [1] },
      'op with a fractional base': { t: 'op', base: 0.5, id: 'a', ops: [1] },
      'op with a negative base': { t: 'op', base: -1, id: 'a', ops: [1] },
      'op without id': { t: 'op', base: 0, ops: [1] },
      'op with an empty id': { t: 'op', base: 0, id: '', ops: [1] },
      'op with an id of 101 characters': { t: 'op', base: 0, id: 'i'.repeat(101), ops: [1] },
      'op whose ops are an object': { t: 'op', base: 0, id: 'a', ops: {} },
      'op with no ops': { t: 'op', base: 0, id: 'a', ops: [] },
      'op with 10002 ops': { t: 'op', base: 0, id: 'a', ops: new Array(10002).fill(0) },
      'saved without v': { t: 'saved' },
      'saved with no versions': { t: 'saved', v: [] },
      'saved with a rev that is a string': { t: 'saved', v: ['v2'], rev: '1' },
    };
    for (const name of Object.keys(later)) {
      const { c } = await joined(port, { v: ['v1'] }, `/bad/${encodeURIComponent(name)}`);
      c.send(later[name]);
      await c.closes(1008, 'bad-message').catch((error) => {
        throw new Error(`${name}: ${error.message}`);
      });
    }
  });
});

// --- The contract of a room --------------------------------------------------

test('an operation on the current revision goes to everyone, its sender too, and a stale one is refused', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], user: { name: 'Ada', color: '#c00' } });
    assert.deepEqual(a.hello, { t: 'hello', pv: 1, ep: a.hello.ep, rev: 0, from: 0, you: a.hello.you, user: a.hello.user, ro: false, v: ['v1'], peers: [] });
    assert.equal(typeof a.hello.user, 'object', 'hello tells the participant who it is to the room');
    assert.equal(typeof a.hello.you, 'string');
    assert.match(a.hello.ep, EPOCH, 'hello tells the name of the room');
    const b = await joined(port, { v: ['v1'] });
    assert.equal(b.hello.ep, a.hello.ep, 'everyone in the room is told the same name');
    assert.deepEqual(b.hello.peers, [{ id: a.hello.you, user: { name: 'Ada', color: '#c00' }, ro: false }]);
    assert.notEqual(b.hello.you, a.hello.you, 'every connection has its own id');
    assert.deepEqual(await a.c.json(), { t: 'peer', id: b.hello.you, user: {}, ro: false });

    a.c.send({ t: 'op', base: 0, id: a.c.id('a1'), ops: [{ set: 'A1', nested: [null, 1.5, 'ş'] }], extra: 'not passed on' });
    const first = { t: 'op', rev: 1, id: a.c.id('a1'), from: a.hello.you, ops: [{ set: 'A1', nested: [null, 1.5, 'ş'] }] };
    assert.deepEqual(await a.c.json(), first, 'the sender gets its own operation back');
    assert.deepEqual(await b.c.json(), first);

    b.c.send({ t: 'op', base: 0, id: b.c.id('b1'), ops: [1] });
    assert.deepEqual(await b.c.json(), { t: 'nack', id: b.c.id('b1'), rev: 1 });
    await a.c.settled();
    b.c.send({ t: 'op', base: 1, id: b.c.id('b1'), ops: [1] });
    assert.deepEqual(await a.c.json(), { t: 'op', rev: 2, id: b.c.id('b1'), from: b.hello.you, ops: [1] });
    assert.deepEqual(await b.c.json(), { t: 'op', rev: 2, id: b.c.id('b1'), from: b.hello.you, ops: [1] });
    b.c.send({ t: 'op', base: 3, id: b.c.id('b2'), ops: [1] });
    assert.deepEqual(await b.c.json(), { t: 'nack', id: b.c.id('b2'), rev: 2 }, 'a base from the future is as stale');

    // Both write on revision 2: one wins, and both hear the same order.
    a.c.send({ t: 'op', base: 2, id: a.c.id('race'), ops: ['a'] });
    b.c.send({ t: 'op', base: 2, id: b.c.id('race'), ops: ['b'] });
    const won = await a.c.json();
    assert.equal(won.t, 'op', 'the operation that won is heard before the refusal of the one that lost');
    assert.equal(won.rev, 3);
    assert.deepEqual(await b.c.json(), won);
    const loser = won.id === a.c.id('race') ? b.c : a.c;
    assert.deepEqual(await loser.json(), { t: 'nack', id: loser.id('race'), rev: 3 });
    await a.c.settled();
    await b.c.settled();
    const room = server.rooms()[0];
    assert.deepEqual(
      { key: room.key, rev: room.rev, floor: room.floor, connections: room.connections, logEntries: room.logEntries },
      { key: '/doc', rev: 3, floor: 0, connections: 2, logEntries: 3 },
    );
    assert.equal(room.logBytes > 0, true, 'the log is weighed');
  });
});

test('a connection that comes back with `since` gets the operations it missed', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port);
    await operations(a.c, 0, 3);
    const back = await connect(port);
    back.send({ t: 'join', v: ['something else'], since: 1, ep: a.hello.ep, key: 'back' });
    const hello = await back.json();
    assert.deepEqual(
      { t: hello.t, ep: hello.ep, rev: hello.rev, from: hello.from, v: hello.v },
      { t: 'hello', ep: a.hello.ep, rev: 3, from: 1, v: [] },
      'since decides in the room it names, whatever the versions say',
    );
    assert.deepEqual(await back.json(), { t: 'op', rev: 2, id: a.c.id('op1'), from: a.hello.you, ops: [1] });
    assert.deepEqual(await back.json(), { t: 'op', rev: 3, id: a.c.id('op2'), from: a.hello.you, ops: [1] });
    await back.settled();
    // Having caught up, it writes on.
    back.send({ t: 'op', base: 3, id: 'back:next', ops: [2] });
    assert.equal((await back.json()).rev, 4);

    const current = await connect(port);
    current.send({ t: 'join', v: [], since: 4, ep: a.hello.ep });
    assert.equal((await current.json()).from, 4);
    await current.settled();
  });
});

test('a version that was saved is not joined from by itself: its client names the revision and the room', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v1', 'v1-etag'] });
    assert.deepEqual(a.hello.v, ['v1', 'v1-etag'], 'the first to come writes its versions at revision 0');
    const ep = a.hello.ep;
    await operations(a.c, 0, 2);
    const b = await joined(port, { v: ['v1-etag'] });
    assert.deepEqual({ from: b.hello.from, rev: b.hello.rev, v: b.hello.v }, { from: 0, rev: 2, v: ['v1-etag'] });
    assert.equal((await b.c.json()).rev, 1);
    assert.equal((await b.c.json()).rev, 2);
    assert.equal((await a.c.json()).t, 'peer');

    a.c.send({ t: 'saved', v: ['v2'], rev: 2 });
    assert.deepEqual(await b.c.json(), { t: 'saved', from: a.hello.you, v: ['v2'], rev: 2 });
    await a.c.settled();
    await operations(a.c, 2, 1);
    assert.equal((await b.c.json()).rev, 3);

    // The table knows v2 as revision 2, and that is no ground to start from: anybody who may write could have said it.
    const late = await knock(port, { v: ['v2'], key: 'late' });
    assert.deepEqual(late.answer, { t: 'reset', rev: 3, v: ['v2'] }, 'a version of a save does not get in by itself');
    assert.deepEqual(await a.c.json(), { t: 'sync' });
    assert.deepEqual(await b.c.json(), { t: 'sync' });
    // The same client, with what its file says of itself: the revision, and the room it was saved in.
    late.c.send({ t: 'join', v: ['v2'], since: 2, ep, key: 'late' });
    const hello = await late.c.json();
    assert.deepEqual(
      { t: hello.t, ep: hello.ep, from: hello.from, rev: hello.rev, v: hello.v },
      { t: 'hello', ep, from: 2, rev: 3, v: ['v2'] },
      'the room says that it has heard of the version',
    );
    assert.deepEqual(await late.c.json(), { t: 'op', rev: 3, id: a.c.id('op2'), from: a.hello.you, ops: [1] });
    await late.c.settled();
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await b.c.json()).t, 'peer');

    // With `since`, hello names every version the room has heard of; without, only the foundation counts.
    const named = await comesBack(port, ep, 2, ['v1', 'v2', 'unknown']);
    assert.deepEqual(named.v, ['v1', 'v2'], 'those of the foundation and those of a save, and no other');
    const both = await joined(port, { v: ['v2', 'v1', 'unknown'] });
    assert.deepEqual({ from: both.hello.from, v: both.hello.v }, { from: 0, v: ['v1'] }, 'the save beside it does not move the join');

    // A save that names a revision the room does not have is passed on without it.
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await b.c.json()).t, 'peer');
    assert.equal((await b.c.json()).t, 'peer');
    b.c.send({ t: 'saved', v: ['v9'], rev: 99 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['v9'] });
    b.c.send({ t: 'saved', v: ['v8'] });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['v8'] });
    b.c.send({ t: 'saved', v: ['v7'], rev: 1.5 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['v7'] });
    const stranger = await connect(port);
    stranger.send({ t: 'join', v: ['v9', 'v8', 'v7'], ro: true });
    assert.deepEqual(await stranger.json(), { t: 'reset', rev: 3, v: ['v2'] }, 'the newest the room knows is still v2');
    assert.deepEqual((await comesBack(port, ep, 3, ['v9', 'v8', 'v7'])).v, [], 'none of them was written down');
  });
});

test('the first to come to a fresh room sets its versions; later strangers are told to reset', async () => {
  await withServer({}, async (port, server) => {
    const first = await joined(port, { v: ['old'] });
    first.c.end();
    await until('the room to be empty', () => server.rooms()[0].connections === 0);
    // Nothing was written and nobody is there: the room is still fresh.
    const second = await joined(port, { v: ['new', 'new'] });
    assert.deepEqual({ from: second.hello.from, v: second.hello.v, peers: second.hello.peers }, { from: 0, v: ['new'], peers: [] });
    const old = await connect(port);
    old.send({ t: 'join', v: ['old'] });
    assert.deepEqual(await old.json(), { t: 'reset', rev: 0, v: ['new'] }, 'the versions of the first visitor are gone');

    // A fresh room keeps the versions even when the first to come names `since: 0`.
    const other = await joined(port, { v: ['x1'], since: 0 }, '/other');
    assert.deepEqual({ from: other.hello.from, v: other.hello.v }, { from: 0, v: ['x1'] });
    const same = await joined(port, { v: ['x1'] }, '/other');
    assert.equal(same.hello.from, 0);

    // No versions and no since, in a room that is not fresh.
    const blind = await connect(port, { path: '/other' });
    blind.send({ t: 'join', v: [] });
    assert.deepEqual(await blind.json(), { t: 'reset', rev: 0, v: ['x1'] });
  });
});

test('after a reset the connection hears and says nothing until it joins with the right version', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'] });
    await operations(a.c, 0, 2);
    const c = await connect(port);
    c.send({ t: 'join', v: ['v0'], since: 7, key: 'cem', user: { name: 'Cem' } });
    assert.deepEqual(await c.json(), { t: 'reset', rev: 2, v: ['v1'] });
    assert.deepEqual(await a.c.json(), { t: 'sync' }, 'the room is asked to save');
    assert.equal(c.closed, false, 'the connection stays open');
    assert.equal(server.rooms()[0].connections, 1, 'it is not in the room');

    c.send({ t: 'op', base: 2, id: 'cem:ignored', ops: [1] });
    c.send({ t: 'presence', data: { at: 1 } });
    c.send({ t: 'saved', v: ['v5'], rev: 2 });
    await c.settled();
    await operations(a.c, 2, 1);
    await a.c.settled();
    await c.settled();
    assert.equal(server.rooms()[0].rev, 3, 'its operation was not taken');

    c.send({ t: 'join', v: ['v1'], key: 'cem', user: { name: 'Cem' } });
    const hello = await c.json();
    assert.deepEqual({ t: hello.t, rev: hello.rev, from: hello.from, v: hello.v }, { t: 'hello', rev: 3, from: 0, v: ['v1'] });
    assert.deepEqual(hello.peers, [{ id: a.hello.you, user: {}, ro: false }]);
    for (const rev of [1, 2, 3]) assert.equal((await c.json()).rev, rev);
    assert.deepEqual(await a.c.json(), { t: 'peer', id: hello.you, user: { name: 'Cem' }, ro: false });
    c.send({ t: 'op', base: 3, id: 'cem:now', ops: [1] });
    assert.equal((await a.c.json()).rev, 4);
  });
});

test('the log is cut at its limits, and what was cut can no longer be joined from', async () => {
  await withServer({ maxLogEntries: 3 }, async (port, server) => {
    const a = await joined(port, { v: ['v1'] });
    await operations(a.c, 0, 2);
    a.c.send({ t: 'saved', v: ['v2'], rev: 2 });
    await a.c.settled();
    await operations(a.c, 2, 3);
    const room = server.rooms()[0];
    assert.deepEqual({ rev: room.rev, floor: room.floor, logEntries: room.logEntries }, { rev: 5, floor: 2, logEntries: 3 });

    const cut = await connect(port);
    const ep = a.hello.ep;
    cut.send({ t: 'join', v: [], since: 1, ep });
    assert.deepEqual(await cut.json(), { t: 'reset', rev: 5, v: ['v2'] }, 'revision 1 is below the floor');
    assert.deepEqual(await a.c.json(), { t: 'sync' });
    cut.send({ t: 'join', v: ['v1'] });
    assert.deepEqual(await cut.json(), { t: 'reset', rev: 5, v: ['v2'] }, 'the version of revision 0 was forgotten');
    assert.deepEqual(await a.c.json(), { t: 'sync' }, 'the room is asked twice a minute');
    cut.send({ t: 'join', v: [], since: 2, ep });
    assert.equal((await cut.json()).from, 2, 'the floor itself can be joined from');
    for (const rev of [3, 4, 5]) assert.equal((await cut.json()).rev, rev);
    await cut.settled();
    assert.equal((await a.c.json()).t, 'peer');

    await operations(a.c, 5, 1);
    assert.equal(server.rooms()[0].floor, 3);
    const late = await connect(port);
    late.send({ t: 'join', v: ['v2'], since: 2, ep, ro: true });
    assert.deepEqual(await late.json(), { t: 'reset', rev: 6, v: [] }, 'v2 was saved at revision 2, below the floor');
    assert.deepEqual((await comesBack(port, ep, 3, ['v1', 'v2'])).v, ['v1'], 'the table let go of v2; the foundation is not forgotten');
  });
  await withServer({ maxLogBytes: 4000 }, async (port, server) => {
    const a = await joined(port);
    await operations(a.c, 0, 10, ['x'.repeat(1000)]);
    const room = server.rooms()[0];
    assert.equal(room.logBytes <= 4000, true, `the log weighs ${room.logBytes} bytes`);
    assert.equal(room.logEntries, 3, 'three operations of a little over 1000 bytes fit');
    assert.equal(room.floor, 7);
  });
});

test('presence is passed on, kept for those who come later, and never logged', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], user: { name: 'Ada' } });
    const b = await joined(port, { v: ['v1'] });
    assert.equal((await a.c.json()).t, 'peer');
    a.c.send({ t: 'presence', data: { cell: 'B2', deep: [1, { x: null }] } });
    assert.deepEqual(await b.c.json(), { t: 'presence', from: a.hello.you, data: { cell: 'B2', deep: [1, { x: null }] } });
    await a.c.settled();

    // Over 16 KiB it is dropped, and the connection is not.
    a.c.send({ t: 'presence', data: 'x'.repeat(16 * 1024 - 1) });
    a.c.send({ t: 'presence', data: 'x'.repeat(16 * 1024 - 2) });
    assert.equal((await b.c.json()).data.length, 16 * 1024 - 2, 'with its quotes this one is 16 KiB to the byte');
    a.c.send({ t: 'presence', data: 'second' });
    assert.deepEqual(await b.c.json(), { t: 'presence', from: a.hello.you, data: 'second' });

    const c = await joined(port, { v: ['v1'] });
    assert.deepEqual(c.hello.peers, [
      { id: a.hello.you, user: { name: 'Ada' }, ro: false, presence: 'second' },
      { id: b.hello.you, user: {}, ro: false },
    ]);
    assert.deepEqual(server.rooms()[0].logEntries, 0);

    a.c.send({ t: 'presence' });
    assert.equal((await b.c.json()).t, 'peer');
    assert.deepEqual(await b.c.json(), { t: 'presence', from: a.hello.you, data: null }, 'no data is null');
  });
});

test('the others hear of who comes and who goes, however it goes', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port);
    const long = { name: 'n'.repeat(70), color: 'c'.repeat(40), role: 'dropped' };
    const b = await joined(port, { v: ['v1'], user: long, ro: true });
    assert.deepEqual(await a.c.json(), {
      t: 'peer',
      id: b.hello.you,
      user: { name: 'n'.repeat(64), color: 'c'.repeat(32) },
      ro: true,
    });
    const odd = await joined(port, { v: ['v1'], user: { name: 7, color: null } });
    assert.deepEqual(await a.c.json(), { t: 'peer', id: odd.hello.you, user: {}, ro: false });
    b.c.end();
    assert.deepEqual(await a.c.json(), { t: 'left', id: b.hello.you }, 'a socket that drops');
    odd.c.bytes(frameBytes(8, Buffer.from([0x03, 0xe9])));
    assert.deepEqual(await a.c.json(), { t: 'left', id: odd.hello.you }, 'a connection that closes');
    const rude = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    rude.c.send({ t: 'nonsense' });
    assert.deepEqual(await a.c.json(), { t: 'left', id: rude.hello.you }, 'a connection the server closes');
  });
});

test('rooms are told apart by the path and by the room of the join, not by the query', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'] }, '/docs/1?token=abc');
    const b = await joined(port, { v: ['v1'] }, '/docs/1?other=1');
    assert.equal(b.hello.peers.length, 1, 'the query does not name a room');
    assert.equal((await a.c.json()).t, 'peer');
    const c = await joined(port, { v: ['v1'] }, '/docs/2');
    assert.equal(c.hello.peers.length, 0);
    const d = await joined(port, { v: ['v1'], room: 'side' }, '/docs/1');
    assert.equal(d.hello.peers.length, 0);
    const e = await joined(port, { v: ['v1'], room: '' }, '/docs/1');
    assert.equal(e.hello.peers.length, 2, 'an empty room name is no room name');
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/docs/1', '/docs/2', '/docs/1#side'],
    );
    assert.deepEqual(await a.c.json(), { t: 'peer', id: e.hello.you, user: {}, ro: false });
    await operations(c.c, 0, 1);
    await a.c.settled().catch(() => {
      throw new Error('an operation of another room was heard');
    });
  });
});

test('a room that is full and a server that is full turn the connection away', async () => {
  await withServer({ maxConnectionsPerRoom: 2 }, async (port) => {
    const a = await joined(port);
    await joined(port);
    const third = await connect(port);
    third.send({ t: 'join', v: ['v1'] });
    await third.closes(1013, 'room-full');
    // A place that is given up can be taken.
    a.c.end();
    const again = await joined(port);
    assert.equal(again.hello.peers.length, 1);
  });
  await withServer({ maxRooms: 1 }, async (port) => {
    await joined(port, { v: ['v1'] }, '/one');
    const other = await connect(port, { path: '/two' });
    other.send({ t: 'join', v: ['v1'] });
    await other.closes(1013, 'server-full');
    const same = await joined(port, { v: ['v1'] }, '/one');
    assert.equal(same.hello.peers.length, 1, 'the room that exists still takes connections');
  });
});

test('a connection that asked to be read-only is refused its operations and its saves', async () => {
  await withServer({}, async (port, server) => {
    const writer = await joined(port);
    const reader = await joined(port, { v: ['v1'], ro: true });
    assert.equal(reader.hello.ro, true);
    assert.equal((await writer.c.json()).ro, true, 'the peer message says so');
    reader.c.send({ t: 'op', base: 0, id: 'r1', ops: [1] });
    assert.deepEqual(await reader.c.json(), { t: 'nack', id: 'r1', rev: 0, reason: 'read-only' });
    reader.c.send({ t: 'saved', v: ['v2'], rev: 0 });
    reader.c.send({ t: 'presence', data: 'reading' });
    assert.deepEqual(
      await writer.c.json(),
      { t: 'presence', from: reader.hello.you, data: 'reading' },
      'the save was not passed on; presence is',
    );
    const stranger = await connect(port);
    stranger.send({ t: 'join', v: ['v2'] });
    assert.equal((await stranger.json()).t, 'reset', 'and the version was not written down');
    assert.deepEqual(await writer.c.json(), { t: 'sync' });
    // The reader is not asked to save: it could not.
    assert.equal(server.rooms()[0].rev, 0);
    await operations(writer.c, 0, 1);
    assert.equal((await reader.c.json()).rev, 1, 'it still hears the operations of the others');
  });
});

test('authorize decides who writes, who reads and who is turned away', async () => {
  const asked: any[] = [];
  const options: CollabServerOptions = {
    authorize(question) {
      asked.push(question);
      if (question.token === 'writer') return true;
      if (question.token === 'reader') return { readOnly: true };
      if (question.token === 'slow-reader') {
        return new Promise((resolve) => setTimeout(() => resolve({ readOnly: true }), 30));
      }
      if (question.token === 'thrower') throw new Error('the directory is down');
      if (question.token === 'rejecter') return Promise.reject(new Error('the directory is down'));
      return false;
    },
  };
  await withServer(options, async (port) => {
    const writer = await joined(port, { v: ['v1'], token: 'writer', room: 'r' });
    assert.equal(writer.hello.ro, false);
    assert.equal(asked[0].room, '/doc#r');
    assert.equal(asked[0].origin, undefined);
    assert.equal(asked[0].request.url, '/doc');

    const reader = await connect(port, { path: '/doc?token=reader', headers: { Origin: 'https://app.example' } });
    reader.send({ t: 'join', v: ['v1'], room: 'r' });
    assert.equal((await reader.json()).ro, true, 'the token of the query is used');
    assert.equal(asked[1].origin, 'https://app.example');
    reader.send({ t: 'op', base: 0, id: 'r1', ops: [1] });
    assert.deepEqual(await reader.json(), { t: 'nack', id: 'r1', rev: 0, reason: 'read-only' });

    const both = await connect(port, { path: '/doc?token=nobody' });
    both.send({ t: 'join', v: ['v1'], room: 'r', token: 'writer' });
    assert.equal((await both.json()).ro, false, 'the token of the message comes first');

    // The answer is awaited, and what was sent meanwhile waits its turn.
    const slow = await connect(port);
    slow.bytes(
      Buffer.concat([
        frameBytes(1, JSON.stringify({ t: 'join', v: ['v1'], room: 'r', token: 'slow-reader' })),
        frameBytes(1, JSON.stringify({ t: 'op', base: 0, id: 's1', ops: [1] })),
        frameBytes(1, JSON.stringify({ t: 'ping' })),
      ]),
    );
    assert.equal((await slow.json()).t, 'hello');
    assert.deepEqual(await slow.json(), { t: 'nack', id: 's1', rev: 0, reason: 'read-only' });
    assert.deepEqual(await slow.json(), { t: 'pong' });

    for (const token of ['nobody', 'thrower', 'rejecter', undefined]) {
      const refused = await connect(port);
      refused.send({ t: 'join', v: ['v1'], room: 'r', token });
      await refused.closes(1008, 'unauthorized').catch((error) => {
        throw new Error(`${token}: ${error.message}`);
      });
    }
  });
});

test('the token is in no message of the server and in nothing it logs', async () => {
  const secret = 'S3cr3t-T0ken-that-must-not-leak';
  const events: unknown[] = [];
  const options: CollabServerOptions = {
    authorize: ({ token }) => {
      if (token === secret) return true;
      if (token === `${secret}-thrown`) throw new Error(`no such token: ${token}`);
      return false;
    },
    log: (event) => events.push(event),
  };
  await withServer(options, async (port) => {
    const clients: Client[] = [];
    const a = await connect(port, { path: `/doc?token=${secret}` });
    a.send({ t: 'join', v: ['v1'], key: 'ada', token: secret, user: { name: 'Ada' } });
    assert.equal((await a.json()).t, 'hello');
    const b = await connect(port, { path: `/doc?token=${secret}` });
    b.send({ t: 'join', v: ['v1'] });
    assert.equal((await b.json()).t, 'hello');
    clients.push(a, b);
    a.send({ t: 'op', base: 0, id: 'ada:a1', ops: [1], token: secret });
    assert.equal((await b.json()).t, 'op');
    a.send({ t: 'presence', data: 1, token: secret });
    assert.equal((await b.json()).t, 'presence');
    a.send({ t: 'saved', v: ['v2'], rev: 1, token: secret });
    assert.equal((await b.json()).t, 'saved');
    for (const token of [`${secret}-wrong`, `${secret}-thrown`]) {
      const refused = await connect(port);
      refused.send({ t: 'join', v: ['v1'], token });
      await refused.closes(1008, 'unauthorized');
      clients.push(refused);
    }
    const reset = await connect(port, { path: `/doc?token=${secret}` });
    reset.send({ t: 'join', v: ['v0'], token: secret });
    assert.equal((await reset.json()).t, 'reset');
    reset.send({ t: 'nonsense', token: secret });
    await reset.closes(1008, 'bad-message');
    clients.push(reset);
    b.end();
    assert.equal((await a.json()).t, 'peer');
    await a.settled().catch(() => undefined);

    for (const client of clients) {
      for (const text of client.seen) assert.equal(text.includes(secret), false, `a message carries the token: ${text}`);
    }
    assert.equal(events.length > 0, true, 'the log was written to');
    const seen = new Set<unknown>();
    const texts: string[] = [];
    const walk = (value: unknown, depth: number): void => {
      if (typeof value === 'string') texts.push(value);
      if (typeof value !== 'object' || value === null || seen.has(value) || depth > 6) return;
      seen.add(value);
      for (const key of Object.keys(value)) {
        texts.push(key);
        walk((value as any)[key], depth + 1);
      }
    };
    walk(events, 0);
    for (const text of texts) assert.equal(text.includes(secret), false, `the log carries the token: ${text}`);
  });
});

test('a client that stops reading is cut off when too much waits for it', async () => {
  // The rate of operations is another limit, with a test of its own.
  await withServer({ maxBufferedBytes: 256 * 1024, opsPerSecond: 0 }, async (port, server) => {
    const stalled = await joined(port);
    const busy = await joined(port);
    stalled.c.socket.pause();
    const chunk = 'x'.repeat(32 * 1024);
    let sent = 0;
    let left: any = null;
    while (!left && sent < 4000) {
      busy.c.send({ t: 'op', base: sent, id: busy.c.id(`op${sent}`), ops: [chunk] });
      sent++;
      for (;;) {
        const message = await busy.c.json();
        if (message.t === 'left') left = message;
        else if (message.t === 'op' && message.id === busy.c.id(`op${sent - 1}`)) break;
        else if (message.t !== 'op') throw new Error(`unexpected ${JSON.stringify(message).slice(0, 100)}`);
      }
    }
    if (!left) left = await busy.c.json();
    assert.deepEqual(left, { t: 'left', id: stalled.hello.you }, `after ${sent} operations of 32 KiB`);
    assert.equal(server.rooms()[0].connections, 1);
    assert.equal(busy.c.closed, false, 'the client that reads is not cut off');
    stalled.c.socket.resume();
    await stalled.c.gone();
  });
});

test('a log longer than the limit of waiting bytes is still played to a client that reads', async () => {
  await withServer({ maxBufferedBytes: 64 * 1024 }, async (port) => {
    const a = await joined(port);
    const chunk = 'y'.repeat(32 * 1024);
    await operations(a.c, 0, 100, [chunk]);
    const b = await joined(port);
    for (let rev = 1; rev <= 100; rev++) assert.equal((await b.c.json()).rev, rev);
    await b.c.settled();
    await operations(b.c, 100, 1, [chunk]);
  });
});

test('an empty room is forgotten after a while, and not while someone is in it', async () => {
  await withServer({ idleRoomMs: 60 }, async (port, server) => {
    const a = await joined(port);
    await operations(a.c, 0, 1);
    const b = await joined(port);
    a.c.end();
    await new Promise((resolve) => setTimeout(resolve, 200));
    assert.equal(server.rooms().length, 1, 'someone is still there');
    b.c.end();
    await until('the room to be forgotten', () => server.rooms().length === 0);
    const fresh = await joined(port, { v: ['other'] });
    assert.deepEqual({ rev: fresh.hello.rev, v: fresh.hello.v }, { rev: 0, v: ['other'] }, 'the room starts again');
    // Coming back before the time is up keeps the room.
    fresh.c.send({ t: 'op', base: 0, id: fresh.c.id('kept'), ops: [1] });
    assert.equal((await fresh.c.json()).rev, 1);
    fresh.c.end();
    await until('the room to be empty', () => server.rooms()[0]?.connections === 0);
    const back = await joined(port, { v: ['other'] });
    assert.equal(back.hello.rev, 1);
    await new Promise((resolve) => setTimeout(resolve, 200));
    assert.equal(server.rooms().length, 1, 'the timer was stopped by the one who came back');
  });
});

// --- Hardening: versions of the protocol, keys, one room -----------------------

/** Connects and joins by hand; the answer is whatever the server says. */
async function knock(port: number, join: object, path = '/doc'): Promise<{ c: Client; answer: any }> {
  const c = await connect(port, { path });
  c.send({ t: 'join', ...join });
  return { c, answer: await c.json() };
}

/** Connects and joins by hand, and is turned away with `code`. */
async function turnedAway(port: number, join: object, status: number, code: string, path = '/doc'): Promise<void> {
  const c = await connect(port, { path });
  c.send({ t: 'join', ...join });
  await c.closes(status, code);
}

test('the protocol version of a room is that of the first to join it, and another version is turned away', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], pv: 3 });
    assert.equal(a.hello.pv, 3, 'hello says the version of the room');
    const same = await joined(port, { v: ['v1'], pv: 3 });
    assert.equal(same.hello.pv, 3);
    assert.equal(server.rooms()[0].pv, 3);
    await turnedAway(port, { v: ['v1'], pv: 2 }, 1008, 'version');
    await turnedAway(port, { v: ['v1'] }, 1008, 'version').catch((error) => {
      throw new Error(`no pv is version 1: ${error.message}`);
    });
    assert.equal(server.rooms()[0].connections, 2, 'neither of them got in');

    // The version stays for as long as the room is kept, with nobody in it too.
    a.c.end();
    same.c.end();
    await until('the room to be empty', () => server.rooms()[0].connections === 0);
    await turnedAway(port, { v: ['v1'], pv: 1 }, 1008, 'version');
    assert.equal((await joined(port, { v: ['v1'], pv: 3 })).hello.pv, 3);

    const plain = await joined(port, { v: ['v1'] }, '/plain');
    assert.equal(plain.hello.pv, 1, 'a join that names no version is of version 1');
    assert.equal((await joined(port, { v: ['v1'], pv: 1 }, '/plain')).hello.pv, 1);
    await turnedAway(port, { v: ['v1'], pv: 1000 }, 1008, 'version', '/plain');
    assert.equal((await joined(port, { v: ['v1'], pv: 1000 }, '/last')).hello.pv, 1000);

    for (const pv of [0, 1001, 1.5, -1, '1', null, true]) {
      await turnedAway(port, { v: ['v1'], pv }, 1008, 'bad-message', '/odd').catch((error) => {
        throw new Error(`pv ${JSON.stringify(pv)}: ${error.message}`);
      });
    }
    assert.equal(server.rooms().some((room) => room.key === '/odd'), false, 'a join that is refused makes no room');
  });
});

test('a session of this package that comes to a room of another version is told "version" and says so: it is not a refusal', async () => {
  await withServer({}, async (port) => {
    const older = await joined(port, { v: ['v1'], pv: PROTOCOL_VERSION - 1 });
    assert.equal(older.hello.pv, PROTOCOL_VERSION - 1);
    // The session speaks over a raw connection: what the server sends is passed on as it comes.
    const c = await connect(port);
    let listener: (text: string) => void = () => undefined;
    const transport = { send: (text: string) => c.send(JSON.parse(text)), onMessage: (given: (text: string) => void) => void (listener = given), close: () => c.end() };
    const wb = await (await open(await startFile())).workbook();
    const notices: string[] = [];
    const session = new CollabSession(
      { workbook: () => wb, changed: () => undefined, notice: (notice) => notices.push(notice.code) },
      { transport, versions: ['v1'], timer: () => () => undefined },
    );
    while (session.state !== 'closed') {
      const frame = await c.frame();
      if (frame.opcode === 1) listener(frame.payload.toString('utf8'));
    }
    assert.deepEqual([session.reason, notices], ['version', ['version']]);
    older.c.end();
  });
});

test('a key is one to sixty-four letters, digits, underscores and hyphens', async () => {
  await withServer({}, async (port) => {
    for (const key of ['', 'k'.repeat(65), 'a b', 'a:b', 'ş', 'a.b', 'a#b', 7, null, ['a']]) {
      await turnedAway(port, { v: ['v1'], key }, 1008, 'bad-message').catch((error) => {
        throw new Error(`key ${JSON.stringify(key)}: ${error.message}`);
      });
    }
    for (const key of ['a', 'Az09_-', 'k'.repeat(40), 'k'.repeat(64)]) {
      const { answer } = await knock(port, { v: ['v1'], key });
      assert.equal(answer.t, 'hello', `key ${key}`);
    }
  });
});

test('the id of an operation is any text of one to a hundred characters, with a key or without', async () => {
  await withServer({}, async (port, server) => {
    const me = await joined(port, { v: ['v1'], key: 'me' });
    const other = await joined(port, { v: ['v1'], key: 'other' });
    assert.equal((await me.c.json()).t, 'peer');
    let rev = 0;
    for (const id of ['other:1', 'me', 'ME:1', ' me:1', ':1', '1', 'ş', 'x'.repeat(100)]) {
      me.c.send({ t: 'op', base: rev, id, ops: [1] });
      rev++;
      assert.deepEqual(await me.c.json(), { t: 'op', rev, id, from: me.hello.you, ops: [1] }, `id ${id}`);
      assert.deepEqual(await other.c.json(), { t: 'op', rev, id, from: me.hello.you, ops: [1] }, `id ${id}, as the other hears it`);
    }
    // A connection that gave no key writes like any other.
    const keyless = await joined(port, { v: ['v1'], key: undefined });
    for (let index = 1; index <= rev; index++) assert.equal((await keyless.c.json()).rev, index, 'the log is played to it');
    keyless.c.send({ t: 'op', base: rev, id: 'no-key', ops: [1] });
    assert.deepEqual(await keyless.c.json(), { t: 'op', rev: rev + 1, id: 'no-key', from: keyless.hello.you, ops: [1] });
    assert.equal(server.rooms()[0].rev, rev + 1);
    // What is no id at all is still a bad message.
    for (const id of ['', 'x'.repeat(101), 7, null]) {
      const odd = await joined(port, { v: ['v1'] }, '/odd');
      odd.c.send({ t: 'op', base: 0, id, ops: [1] });
      await odd.c.closes(1008, 'bad-message').catch((error) => {
        throw new Error(`id ${JSON.stringify(id)}: ${error.message}`);
      });
    }
    // A connection that only reads is told so.
    const reader = await joined(port, { v: ['v1'], key: undefined, ro: true });
    reader.c.send({ t: 'op', base: rev + 1, id: 'x:1', ops: [1] });
    for (;;) {
      const back = await reader.c.json();
      if (back.t === 'op') continue;
      assert.deepEqual(back, { t: 'nack', id: 'x:1', rev: rev + 1, reason: 'read-only' });
      break;
    }
  });
});

test('whoever joins with the key of a connection of the room takes its place', async () => {
  await withServer({ maxConnectionsPerRoom: 2 }, async (port, server) => {
    const old = await joined(port, { v: ['v1'], key: 'same', user: { name: 'Ada' } });
    const witness = await joined(port, { v: ['v1'], key: 'witness' });
    assert.equal((await old.c.json()).t, 'peer');

    // The room is full, and the one that comes back gets in all the same: it is not one more.
    const back = await joined(port, { v: ['v1'], key: 'same', user: { name: 'Ada' } });
    await old.c.closes(1008, 'replaced');
    assert.deepEqual(await witness.c.json(), { t: 'left', id: old.hello.you }, 'the others hear the old one go');
    assert.deepEqual(
      await witness.c.json(),
      { t: 'peer', id: back.hello.you, user: { name: 'Ada' }, ro: false },
      'and then the new one come',
    );
    assert.deepEqual(
      back.hello.peers.map((peer: any) => peer.id),
      [witness.hello.you],
      'the one it replaced is not among its peers',
    );
    assert.notEqual(back.hello.you, old.hello.you, 'it is a connection of its own');
    assert.equal(server.rooms()[0].connections, 2);
    back.c.send({ t: 'op', base: 0, id: 'after', ops: [1] });
    assert.equal((await witness.c.json()).id, 'after', 'the new one writes');

    // Another key is one more, and the room is full.
    await turnedAway(port, { v: ['v1'], key: 'third' }, 1013, 'room-full');
    // A join that is turned back with `reset` takes nobody's place.
    const stale = await knock(port, { v: ['v0'], key: 'same', ro: true });
    assert.equal(stale.answer.t, 'reset');
    await back.c.json().then((message) => assert.equal(message.id, 'after', 'its own operation comes back to it'));
    await back.c.settled().catch(() => {
      throw new Error('a join that did not get in closed the connection whose key it had');
    });
    assert.equal(server.rooms()[0].connections, 2);
  });
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], key: 'same' });
    // Another room has its own keys.
    const elsewhere = await joined(port, { v: ['v1'], key: 'same' }, '/elsewhere');
    const side = await joined(port, { v: ['v1'], key: 'same', room: 'side' });
    // Two connections without a key do not share one.
    const one = await joined(port, { v: ['v1'], key: undefined });
    const two = await joined(port, { v: ['v1'], key: undefined });
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await one.c.json()).t, 'peer');
    for (const client of [a.c, elsewhere.c, side.c, one.c, two.c]) {
      await client.settled().catch(() => {
        throw new Error('a connection of another room or one without a key was replaced');
      });
    }
    assert.equal(server.rooms().find((room) => room.key === '/doc')?.connections, 3);
  });
});

test('a connection joins one room, once', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], room: 'one' });
    const b = await joined(port, { v: ['v1'], room: 'one' });
    assert.equal((await a.c.json()).t, 'peer');
    b.c.send({ t: 'join', v: ['v1'], room: 'two', key: b.c.key });
    await b.c.closes(1008, 'bad-message');
    assert.deepEqual(await a.c.json(), { t: 'left', id: b.hello.you });
    // The same room again is as wrong.
    a.c.send({ t: 'join', v: ['v1'], room: 'one', key: a.c.key });
    await a.c.closes(1008, 'bad-message');
    assert.deepEqual(
      server.rooms().map((room) => [room.key, room.connections]),
      [['/doc#one', 0]],
      'the second join made no room',
    );
  });
});

test('one connection cannot fill the table of rooms', async () => {
  const asked: string[] = [];
  await withServer({ maxRooms: 1000 }, async (port, server) => {
    const flood = await connect(port, { path: '/flood' });
    const frames: Buffer[] = [];
    for (let i = 0; i < 1100; i++) frames.push(frameBytes(1, JSON.stringify({ t: 'join', v: ['v'], room: `r${i}` })));
    flood.bytes(Buffer.concat(frames));
    assert.equal((await flood.json()).t, 'hello', 'its first join is answered');
    await flood.closes(1008, 'bad-message');
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/flood#r0'],
    );
    const honest = await joined(port, { v: ['v'] }, '/honest');
    assert.equal(honest.hello.t, 'hello', 'there is room for an honest client');

    // Joins that are refused leave nothing behind either.
    const before = server.rooms().length;
    await turnedAway(port, { v: ['v'], pv: 0 }, 1008, 'bad-message', '/refused/1');
    await turnedAway(port, { v: ['v'], key: 'not a key' }, 1008, 'bad-message', '/refused/2');
    await turnedAway(port, { v: ['v'], room: 'a#b' }, 1008, 'bad-message', '/refused/3');
    assert.equal(server.rooms().length, before);
  });
  await withServer(
    {
      authorize: ({ room }) => {
        asked.push(room);
        return false;
      },
    },
    async (port, server) => {
      for (let i = 0; i < 5; i++) await turnedAway(port, { v: ['v'], room: `r${i}` }, 1008, 'unauthorized');
      assert.equal(asked.length, 5);
      assert.deepEqual(server.rooms(), [], 'a join that is not authorized makes no room');
    },
  );
});

test('a connection that leaves its room may join that room again over the same connection, and no other', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v1'], room: 'r' });
    const b = await joined(port, { v: ['v1'], room: 'r', key: 'b' });
    assert.equal((await a.c.json()).t, 'peer', 'the first hears of the second');
    b.c.send({ t: 'leave' });
    assert.deepEqual(await a.c.json(), { t: 'left', id: b.hello.you }, 'the room is told that it left');
    b.c.send({ t: 'join', v: ['v1'], room: 'r', key: 'b' });
    const hello = await b.c.json();
    assert.deepEqual({ t: hello.t, from: hello.from, ep: hello.ep }, { t: 'hello', from: 0, ep: b.hello.ep }, 'it joins the room again, over the connection it has');
    assert.equal((await a.c.json()).t, 'peer', 'and the room hears of it again');
    // In no room, leaving is nothing; and a room that was left is the only one to come back to.
    b.c.send({ t: 'leave' });
    assert.equal((await a.c.json()).t, 'left');
    b.c.send({ t: 'leave' });
    b.c.send({ t: 'join', v: ['v1'], room: 'other' });
    await b.c.closes(1008, 'bad-message');
  });
});

test('who a participant is, is what `authorize` says: its name stands in place of what the client says of itself', async () => {
  await withServer(
    {
      authorize: ({ token }) => (token === 'known' ? { user: { name: 'Named by the application', color: '#112233' } } : token === 'reader' ? { readOnly: true, user: { name: 'A reader' } } : true),
    },
    async (port) => {
      const first = await joined(port, { v: ['v0'], token: 'anyone', user: { name: 'What I say I am', color: '#ABCDEF' } });
      assert.deepEqual(first.hello.user, { name: 'What I say I am', color: '#ABCDEF' }, 'without a word of the application a participant is who it says it is, and is told so');
      const known = await joined(port, { v: ['v0'], token: 'known', user: { name: 'Somebody else entirely', color: '#FFFFFF' } });
      assert.deepEqual(known.hello.user, { name: 'Named by the application', color: '#112233' }, 'the name and the colour of the application, told to the participant itself');
      assert.deepEqual((await first.c.json()).user, { name: 'Named by the application', color: '#112233' }, 'and to the room');
      assert.deepEqual(known.hello.peers.map((peer) => peer.user), [{ name: 'What I say I am', color: '#ABCDEF' }], 'the others as the room has them');
      const reader = await joined(port, { v: ['v0'], token: 'reader', user: { name: 'An editor, honest', color: '#00FF00' } });
      assert.deepEqual([reader.hello.user, reader.hello.ro], [{ name: 'A reader', color: '#00FF00' }, true], 'a name alone leaves the colour the client gave');
    },
  );
});

test('a connection joins only so many times a minute: one that leaves and joins over and over is closed', async () => {
  let time = 1000;
  await withServer({ now: () => time }, async (port) => {
    const a = await joined(port, { v: ['v1'] });
    const b = await joined(port, { v: ['v1'], ro: true });
    assert.equal((await a.c.json()).t, 'peer');
    // Nine more times it leaves and joins (ten joins in all): each is told to the room.
    for (let i = 0; i < 9; i++) {
      b.c.send({ t: 'leave' });
      assert.equal((await a.c.json()).t, 'left', `leave ${i + 1} is told`);
      b.c.send({ t: 'join', v: ['v1'], ro: true });
      assert.equal((await b.c.json()).t, 'hello', `join ${i + 2} of the minute gets in`);
      assert.equal((await a.c.json()).t, 'peer');
    }
    b.c.send({ t: 'leave' });
    assert.equal((await a.c.json()).t, 'left');
    b.c.send({ t: 'join', v: ['v1'], ro: true });
    await b.c.closes(1008, 'join-rate');
    // The limit is of the connection, and of the minute.
    const c = await joined(port, { v: ['v1'], ro: true });
    assert.equal(c.hello.t, 'hello', 'another connection joins');
    time += 61000;
    c.c.send({ t: 'leave' });
    c.c.send({ t: 'join', v: ['v1'], ro: true });
    assert.equal((await c.c.json()).t, 'hello', 'and a minute later a connection joins again');
  });
  await withServer({ joinsPerMinute: 0 }, async (port) => {
    const b = await joined(port, { v: ['v1'] });
    for (let i = 0; i < 15; i++) {
      b.c.send({ t: 'leave' });
      b.c.send({ t: 'join', v: ['v1'] });
      assert.equal((await b.c.json()).t, 'hello', `0 is no limit: join ${i + 2}`);
    }
  });
});

test('after a reset a connection may join the same room again, and no other', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'], room: 'r' });
    await operations(a.c, 0, 1);

    const wanderer = await knock(port, { v: ['v0'], room: 'r' });
    assert.deepEqual(wanderer.answer, { t: 'reset', rev: 1, v: ['v1'] });
    wanderer.c.send({ t: 'join', v: ['v0'], room: 'other' });
    await wanderer.c.closes(1008, 'bad-message');
    const bare = await knock(port, { v: ['v0'], room: 'r' });
    bare.c.send({ t: 'join', v: ['v0'] });
    await bare.c.closes(1008, 'bad-message').catch((error) => {
      throw new Error(`the room of the path alone is another room: ${error.message}`);
    });
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/doc#r'],
      'no room was made on the way',
    );

    const patient = await knock(port, { v: ['v0'], room: 'r', key: 'p' });
    assert.equal(patient.answer.t, 'reset');
    patient.c.send({ t: 'join', v: ['still wrong'], room: 'r', key: 'p' });
    assert.deepEqual(await patient.c.json(), { t: 'reset', rev: 1, v: ['v1'] }, 'a second reset is no offence');
    patient.c.send({ t: 'join', v: ['v1'], room: 'r', key: 'p' });
    const hello = await patient.c.json();
    assert.deepEqual({ t: hello.t, from: hello.from, rev: hello.rev }, { t: 'hello', from: 0, rev: 1 });
    assert.equal((await patient.c.json()).rev, 1);
    // Now it is in, and a third join is one too many.
    patient.c.send({ t: 'join', v: ['v1'], room: 'r', key: 'p' });
    await patient.c.closes(1008, 'bad-message');
  });
});

test('a connection that was reset and does not join in time is closed', async () => {
  await withServer({ joinTimeoutMs: 1000 }, async (port) => {
    const a = await joined(port);
    await operations(a.c, 0, 1);
    const lost = await knock(port, { v: ['v0'] });
    assert.equal(lost.answer.t, 'reset');
    await lost.c.closes(1008, 'join-timeout');

    const found = await knock(port, { v: ['v0'] });
    assert.equal(found.answer.t, 'reset');
    found.c.send({ t: 'join', v: ['v1'] });
    assert.equal((await found.c.json()).t, 'hello');
    assert.equal((await found.c.json()).rev, 1);
    await new Promise((resolve) => setTimeout(resolve, 2500));
    await found.c.settled().catch(() => {
      throw new Error('a connection that joined after its reset was closed all the same');
    });
  });
});

// --- Hardening: the table of versions ----------------------------------------

test('a version is written down once: a save cannot move the version a room was founded on', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['orig', 'orig-etag'] });
    await operations(a.c, 0, 5);
    const forger = await joined(port, { v: ['orig'] });
    for (const rev of [1, 2, 3, 4, 5]) assert.equal((await forger.c.json()).rev, rev);
    assert.equal((await a.c.json()).t, 'peer');

    forger.c.send({ t: 'saved', v: ['orig'], rev: 5 });
    assert.deepEqual(
      await a.c.json(),
      { t: 'saved', from: forger.hello.you, v: ['orig'] },
      'the save is passed on without the revision the table contradicts',
    );
    const late = await joined(port, { v: ['orig'] });
    assert.deepEqual({ from: late.hello.from, rev: late.hello.rev }, { from: 0, rev: 5 }, 'orig is still revision 0');
    for (const rev of [1, 2, 3, 4, 5]) assert.equal((await late.c.json()).rev, rev, 'every operation is played to it');
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await forger.c.json()).t, 'peer');

    // A version that a save wrote is as fixed, and a save that agrees is passed on whole.
    a.c.send({ t: 'saved', v: ['v2'], rev: 2 });
    assert.deepEqual(await forger.c.json(), { t: 'saved', from: a.hello.you, v: ['v2'], rev: 2 });
    assert.equal((await late.c.json()).t, 'saved');
    forger.c.send({ t: 'saved', v: ['v2'], rev: 5 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: forger.hello.you, v: ['v2'] });
    forger.c.send({ t: 'saved', v: ['v2'], rev: 2 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: forger.hello.you, v: ['v2'], rev: 2 });
    // One version that is taken spoils the whole save.
    forger.c.send({ t: 'saved', v: ['v5', 'orig-etag'], rev: 5 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: forger.hello.you, v: ['v5', 'orig-etag'] });
    forger.c.send({ t: 'saved', v: ['v2'], rev: 2 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: forger.hello.you, v: ['v2'], rev: 2 }, 'v2 is still revision 2');
    const third = await knock(port, { v: ['v5'] });
    assert.equal(third.answer.t, 'reset', 'v5 is no ground to join on');
    assert.deepEqual(third.answer.v, ['v2']);
    assert.deepEqual((await comesBack(port, a.hello.ep, 2, ['v2', 'v5'])).v, ['v2'], 'v2 is known and v5 was not written down');
  });
});

test('the versions a room was founded on stay taken when the log has moved past them', async () => {
  await withServer({ maxLogEntries: 2 }, async (port, server) => {
    const a = await joined(port, { v: ['orig'] });
    await operations(a.c, 0, 4);
    assert.equal(server.rooms()[0].floor, 2);
    a.c.send({ t: 'saved', v: ['orig'], rev: 4 });
    a.c.send({ t: 'saved', v: ['v3'], rev: 3 });
    await a.c.settled();
    const old = await knock(port, { v: ['orig'] });
    assert.deepEqual(old.answer, { t: 'reset', rev: 4, v: ['v3'] }, 'the original file cannot be given a later revision');
  });
});

test('a save names only a revision its sender has seen', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v1'] });
    await operations(a.c, 0, 3);
    const b = await joined(port, { v: [], since: 2, ep: a.hello.ep });
    assert.equal(b.hello.from, 2);
    assert.equal((await b.c.json()).rev, 3);
    assert.equal((await a.c.json()).t, 'peer');

    b.c.send({ t: 'saved', v: ['below'], rev: 1 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['below'] }, 'it joined from 2: 1 it never saw');
    b.c.send({ t: 'saved', v: ['beyond'], rev: 4 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['beyond'] }, 'the room is at 3');
    b.c.send({ t: 'saved', v: ['at-two'], rev: 2 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['at-two'], rev: 2 });
    b.c.send({ t: 'saved', v: ['at-three'], rev: 3 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: b.hello.you, v: ['at-three'], rev: 3 });

    // The one who saw everything may name any of it.
    a.c.send({ t: 'saved', v: ['at-one'], rev: 1 });
    assert.deepEqual(await b.c.json(), { t: 'saved', from: a.hello.you, v: ['at-one'], rev: 1 });

    for (const version of ['below', 'beyond']) {
      const stranger = await knock(port, { v: [version] });
      assert.deepEqual(stranger.answer, { t: 'reset', rev: 3, v: ['at-three'] }, `${version} is no ground to join on`);
    }
    assert.deepEqual(
      (await comesBack(port, a.hello.ep, 2, ['below', 'beyond', 'at-two', 'at-three', 'at-one'])).v,
      ['at-two', 'at-three', 'at-one'],
      'what named a revision its sender had not seen was not written down',
    );
  });
});

test('a room remembers sixteen versions; the oldest save is forgotten first, the versions it was founded on last', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    await operations(a.c, 0, 3);
    for (const rev of [1, 2, 3]) {
      const names = [0, 1, 2, 3, 4, 5, 6, 7].map((index) => `r${rev}-${index}`);
      a.c.send({ t: 'saved', v: names, rev });
    }
    await a.c.settled();
    const first = await knock(port, { v: ['r1-0', 'r1-7', 'r2-0'] });
    assert.equal(first.answer.t, 'reset');
    first.c.send({ t: 'join', v: ['v0'] });
    assert.equal((await first.c.json()).from, 0, 'the version of the foundation is kept');
    assert.deepEqual(
      (await comesBack(port, a.hello.ep, 3, ['r1-0', 'r1-7', 'r2-0', 'r2-1', 'r3-0', 'r3-7', 'v0'])).v,
      ['r2-1', 'r3-0', 'r3-7', 'v0'],
      'the nine oldest saves made room for newer ones',
    );
  });
});

test('the writers of a room are asked to save when someone who may write has a file the room does not know', async () => {
  let time = 1000;
  const options: CollabServerOptions = {
    now: () => time,
    authorize: ({ token }) => (token === 'viewer' ? { readOnly: true } : true),
  };
  await withServer(options, async (port, server) => {
    const a = await joined(port, { v: ['v1'] });
    const b = await joined(port, { v: ['v1'] });
    const reader = await joined(port, { v: ['v1'], ro: true });
    const viewer = await joined(port, { v: ['v1'], token: 'viewer' });
    assert.equal(viewer.hello.ro, true, 'authorize made it read-only');
    for (const count of [3, 2, 1]) {
      const client = [a, b, reader][3 - count].c;
      for (let i = 0; i < count; i++) assert.equal((await client.json()).t, 'peer');
    }
    await operations(a.c, 0, 1);
    for (const client of [b.c, reader.c, viewer.c]) assert.equal((await client.json()).rev, 1);
    const quiet = async (why: string): Promise<void> => {
      for (const client of [a.c, b.c, reader.c, viewer.c]) {
        await client.settled().catch(() => {
          throw new Error(why);
        });
      }
    };

    // Whoever could not write what it is asking for has nothing to ask.
    const readOnly = await knock(port, { v: ['v0'], ro: true });
    assert.equal(readOnly.answer.t, 'reset');
    await quiet('a join that asked to read made the room save');
    const viewing = await knock(port, { v: ['v0'], token: 'viewer' });
    assert.equal(viewing.answer.t, 'reset');
    await quiet('a join that authorize lets only read made the room save');

    const first = await knock(port, { v: ['v0'] });
    assert.equal(first.answer.t, 'reset');
    assert.deepEqual(await a.c.json(), { t: 'sync' });
    assert.deepEqual(await b.c.json(), { t: 'sync' }, 'every writer hears it');
    await quiet('a connection that cannot save was asked to');

    // Two a minute for the room, whoever is reset and however often.
    first.c.send({ t: 'join', v: ['v0'] });
    assert.equal((await first.c.json()).t, 'reset');
    assert.deepEqual(await a.c.json(), { t: 'sync' });
    assert.deepEqual(await b.c.json(), { t: 'sync' });
    for (let i = 0; i < 5; i++) {
      first.c.send({ t: 'join', v: ['v0'] });
      assert.equal((await first.c.json()).t, 'reset');
    }
    const second = await knock(port, { v: [], since: 9 });
    assert.equal(second.answer.t, 'reset');
    await quiet('a third sync followed within the minute');
    await first.c.settled().catch(() => {
      throw new Error('a connection that is in no room heard something');
    });
    time += 29000;
    first.c.send({ t: 'join', v: ['v0'] });
    assert.equal((await first.c.json()).t, 'reset');
    await quiet('half a token is no token');
    time += 2000;
    first.c.send({ t: 'join', v: ['v0'] });
    assert.equal((await first.c.json()).t, 'reset');
    assert.deepEqual(await a.c.json(), { t: 'sync' }, 'after half a minute there is one more');
    assert.deepEqual(await b.c.json(), { t: 'sync' });
    first.c.send({ t: 'join', v: ['v0'] });
    assert.equal((await first.c.json()).t, 'reset');
    await quiet('and only one');

    // A join that gets in asks for nothing.
    const fine = await joined(port, { v: ['v1'] }, '/quiet');
    await joined(port, { v: ['v1'] }, '/quiet');
    assert.equal((await fine.c.json()).t, 'peer');
    await fine.c.settled();

    // With no writer in the room there is nobody to ask, and nothing is spent.
    time += 600000;
    a.c.end();
    b.c.end();
    await until('the writers to be gone', () => server.rooms().find((room) => room.key === '/doc')?.connections === 2);
    for (const client of [reader.c, viewer.c]) {
      assert.equal((await client.json()).t, 'left');
      assert.equal((await client.json()).t, 'left');
    }
    for (let i = 0; i < 3; i++) {
      first.c.send({ t: 'join', v: ['v0'] });
      assert.equal((await first.c.json()).t, 'reset');
    }
    for (const client of [reader.c, viewer.c]) {
      await client.settled().catch(() => {
        throw new Error('readers were asked to save');
      });
    }
    const writer = await joined(port, { v: ['v1'] });
    assert.equal((await writer.c.json()).rev, 1);
    for (let i = 0; i < 2; i++) {
      first.c.send({ t: 'join', v: ['v0'] });
      assert.equal((await first.c.json()).t, 'reset');
      assert.deepEqual(await writer.c.json(), { t: 'sync' }, 'the asking that reached nobody was not counted');
    }
  });
});

test('a message carries up to 10001 operations', async () => {
  await withServer({}, async (port) => {
    const { c } = await joined(port);
    c.send({ t: 'op', base: 0, id: c.id('most'), ops: new Array(10001).fill(0) });
    const back = await c.json();
    assert.deepEqual({ t: back.t, rev: back.rev, count: back.ops.length }, { t: 'op', rev: 1, count: 10001 });
    c.send({ t: 'op', base: 1, id: c.id('more'), ops: new Array(10002).fill(0) });
    await c.closes(1008, 'bad-message');
  });
});

// --- Hardening: what a client can use up -------------------------------------

/** Joins as soon as the server has a place: a connection that ended is not gone at once. */
async function joinedInTheEnd(port: number, path = '/doc'): Promise<Client> {
  const end = Date.now() + WAIT;
  for (;;) {
    const { c, answer } = await knock(port, { v: ['v1'] }, path);
    if (answer.t === 'hello') return c;
    assert.deepEqual(answer, { t: 'error', code: 'server-full' });
    if (Date.now() > end) throw new Error('timed out waiting for a place on the server');
    await new Promise((resolve) => setTimeout(resolve, 20));
  }
}

test('a server that has all the connections it takes turns the next away', async () => {
  await withServer({ maxConnections: 2 }, async (port, server) => {
    const a = await joined(port, { v: ['v1'] }, '/a');
    const silent = await connect(port, { path: '/b' });
    const third = await connect(port, { path: '/c' });
    assert.equal(third.status, 101, 'the handshake is answered, so that the client can be told why');
    await third.closes(1013, 'server-full');
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/a'],
    );
    silent.end();
    const next = await joinedInTheEnd(port, '/a');
    await next.settled();
    await a.c.settled().catch(() => undefined);
  });
});

test('one address has only so many connections', async () => {
  await withServer({ maxConnectionsPerIp: 2 }, async (port) => {
    const a = await joined(port);
    const b = await joined(port);
    const third = await connect(port);
    await third.closes(1013, 'server-full');
    a.c.end();
    const next = await joinedInTheEnd(port);
    await next.settled();
    await b.c.settled().catch(() => undefined);
  });
  await withServer({}, async (port) => {
    for (let i = 0; i < 12; i++) await joined(port);
    assert.equal((await joined(port)).hello.peers.length, 12, 'by default an address has as many as it likes');
  });
});

test('the logs of all rooms together have a limit, and the oldest operations go first', async () => {
  await withServer({ maxTotalLogBytes: 4500 }, async (port, server) => {
    const big = ['x'.repeat(1000)];
    const old = await joined(port, { v: ['v1'] }, '/old');
    const young = await joined(port, { v: ['v1'] }, '/young');
    await operations(old.c, 0, 2, big);
    await operations(young.c, 0, 2, big);
    const total = (): number => server.rooms().reduce((sum, room) => sum + room.logBytes, 0);
    assert.equal(total() <= 4500, true, `four operations weigh ${total()} bytes`);
    assert.deepEqual(
      server.rooms().map((room) => [room.key, room.floor, room.logEntries]),
      [
        ['/old', 0, 2],
        ['/young', 0, 2],
      ],
    );
    await operations(young.c, 2, 1, big);
    assert.deepEqual(
      server.rooms().map((room) => [room.key, room.floor, room.logEntries]),
      [
        ['/old', 1, 1],
        ['/young', 0, 3],
      ],
      'the oldest operation of all was in /old',
    );
    await operations(young.c, 3, 2, big);
    assert.deepEqual(
      server.rooms().map((room) => [room.key, room.rev, room.floor, room.logEntries]),
      [
        ['/old', 2, 2, 0],
        ['/young', 5, 1, 4],
      ],
      'then the oldest of /young',
    );
    assert.equal(total() <= 4500, true, `the logs weigh ${total()} bytes`);
    // What was cut can no longer be joined from; the room goes on.
    const late = await knock(port, { v: ['v1'] }, '/old');
    assert.deepEqual(late.answer, { t: 'reset', rev: 2, v: [] });
    assert.deepEqual(await old.c.json(), { t: 'sync' });
    await operations(old.c, 2, 1, big);
    assert.equal(total() <= 4500, true, `the logs weigh ${total()} bytes`);
  });
  // A room that is forgotten gives its bytes back.
  await withServer({ maxTotalLogBytes: 2500, idleRoomMs: 50 }, async (port, server) => {
    const big = ['x'.repeat(1000)];
    const gone = await joined(port, { v: ['v1'] }, '/gone');
    await operations(gone.c, 0, 2, big);
    gone.c.end();
    await until('the room to be forgotten', () => server.rooms().length === 0);
    const kept = await joined(port, { v: ['v1'] }, '/kept');
    await operations(kept.c, 0, 2, big);
    assert.deepEqual(
      server.rooms().map((room) => [room.key, room.floor, room.logEntries]),
      [['/kept', 0, 2]],
      'the bytes of the forgotten room are not held against the others',
    );
  });
});

test('a connection writes only so many operations a second', async () => {
  await withServer({ opsPerSecond: 0.001, opsBurst: 3 }, async (port, server) => {
    const a = await joined(port);
    const b = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 3);
    a.c.send({ t: 'op', base: 3, id: a.c.id('fourth'), ops: [1] });
    assert.deepEqual(await a.c.json(), { t: 'nack', id: a.c.id('fourth'), rev: 3, reason: 'rate' });
    a.c.send({ t: 'op', base: 0, id: a.c.id('stale'), ops: [1] });
    assert.deepEqual(await a.c.json(), { t: 'nack', id: a.c.id('stale'), rev: 3, reason: 'rate' }, 'rate is said first');
    assert.equal(a.c.closed, false, 'the connection is not closed');
    assert.equal(server.rooms()[0].rev, 3);
    for (const rev of [1, 2, 3]) assert.equal((await b.c.json()).rev, rev);
    // The limit is of the connection, not of the room.
    await operations(b.c, 3, 3);
    b.c.send({ t: 'op', base: 6, id: b.c.id('fourth'), ops: [1] });
    assert.deepEqual(await b.c.json(), { t: 'nack', id: b.c.id('fourth'), rev: 6, reason: 'rate' });
  });
  await withServer({ opsPerSecond: 20, opsBurst: 1 }, async (port) => {
    const { c } = await joined(port);
    await operations(c, 0, 1);
    // The bucket fills again: sooner or later the next one is taken.
    const end = Date.now() + WAIT;
    let nacks = 0;
    for (;;) {
      c.send({ t: 'op', base: 1, id: c.id('again'), ops: [1] });
      const back = await c.json();
      if (back.t === 'op') break;
      assert.deepEqual(back, { t: 'nack', id: c.id('again'), rev: 1, reason: 'rate' });
      nacks++;
      if (Date.now() > end) throw new Error(`the bucket never filled again (${nacks} refusals)`);
      await new Promise((resolve) => setTimeout(resolve, 20));
    }
  });
  await withServer({}, async (port) => {
    const { c } = await joined(port);
    const frames: Buffer[] = [];
    for (let i = 0; i < 260; i++) frames.push(frameBytes(1, JSON.stringify({ t: 'op', base: i, id: c.id(`n${i}`), ops: [1] })));
    c.bytes(Buffer.concat(frames));
    const kinds = { op: 0, rate: 0 };
    for (let i = 0; i < 260; i++) {
      const back = await c.json();
      if (back.t === 'op') kinds.op++;
      else if (back.t === 'nack' && back.reason === 'rate') kinds.rate++;
      else if (back.t !== 'nack') throw new Error(`unexpected ${JSON.stringify(back)}`);
    }
    assert.equal(kinds.op >= 200 && kinds.op < 260, true, `the default burst is 200: ${kinds.op} were taken`);
    assert.equal(kinds.rate >= 1, true, 'and what is over it is refused for its rate');
  });
});

test('presence that comes too fast is dropped, and not kept', async () => {
  await withServer({ presencePerSecond: 0.001 }, async (port) => {
    const a = await joined(port);
    const b = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    a.c.send({ t: 'presence', data: 'first' });
    a.c.send({ t: 'presence', data: 'second' });
    a.c.send({ t: 'presence', data: 'third' });
    assert.deepEqual(await b.c.json(), { t: 'presence', from: a.hello.you, data: 'first' });
    await a.c.settled();
    assert.equal(a.c.closed, false, 'the connection is not closed');
    await b.c.settled().catch(() => {
      throw new Error('presence over the rate was passed on');
    });
    const late = await joined(port);
    assert.equal(late.hello.peers.find((peer: any) => peer.id === a.hello.you).presence, 'first', 'what was dropped is not kept');
  });
  await withServer({}, async (port) => {
    const a = await joined(port);
    const b = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    const frames: Buffer[] = [];
    for (let i = 0; i < 500; i++) frames.push(frameBytes(1, JSON.stringify({ t: 'presence', data: i })));
    a.c.bytes(Buffer.concat(frames));
    await a.c.settled();
    b.c.send({ t: 'ping' });
    let heard = 0;
    for (;;) {
      const message = await b.c.json();
      if (message.t === 'pong') break;
      assert.equal(message.t, 'presence', `presence or the pong was expected, ${JSON.stringify(message)} came`);
      heard++;
    }
    assert.equal(heard >= 20 && heard < 500, true, `the default is 20 a second: ${heard} of 500 were passed on`);
  });
});

test('the path of a room is the path the request resolves to', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'] }, '/x/../y');
    const b = await joined(port, { v: ['v1'] }, '/y');
    assert.equal(b.hello.peers.length, 1, '/x/../y is /y');
    const c = await joined(port, { v: ['v1'] }, '/z/./../y?q=1');
    assert.equal(c.hello.peers.length, 2);
    const d = await joined(port, { v: ['v1'] }, `http://127.0.0.1:${port}/w/../y`);
    assert.equal(d.hello.peers.length, 3, 'the absolute form of a target');
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/y'],
    );
    const refused = ['/x/%2e%2e/y', '/x/%2E%2E/y', '/x/.%2e/y', '/x/%2e./y', '/x/%2e/y', '/x/%2e', '/p#q', '/p?x=1#q', '/#'];
    // What a segment may not be once it is decoded, and what cannot be decoded.
    refused.push('/p/a%2fb', '/p/%2F', '/p/a%5cb', '/p/%5C', '/p/%00', '/p/a%00', '/p/%1f', '/p/%0a', '/p/%7f', '/p/%c2%85');
    refused.push('/p/%', '/p/%zz', '/p/%6', '/p/%c3', '/p/%ff', '/p/%ed%a0%80', '/p%23q', '/%23');
    for (const path of refused) {
      const odd = await connect(port, { path });
      assert.equal(odd.status, 400, path);
      await odd.gone();
    }
    // Dots that are part of a name are a name.
    const dotted = await joined(port, { v: ['v1'] }, '/a..b/.c/d%2ee');
    assert.equal(dotted.hello.peers.length, 0);
    // The path cannot be made to look like a path and a room: a `#` is in no path, however it is written.
    const side = await joined(port, { v: ['v1'], room: 'q' }, '/p');
    assert.equal(side.hello.peers.length, 0);
    await turnedAway(port, { v: ['v1'], room: 'a#b' }, 1008, 'bad-message', '/p');
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/y', '/a..b/.c/d.e', '/p#q'],
    );
    a.c.end();
  });
  await withServer({}, async (port, server) => {
    // A room has one name, however its path is spelled.
    const plain = await joined(port, { v: ['v1'] }, '/p/a');
    const spelled = await joined(port, { v: ['v1'] }, '/p/%61');
    assert.equal(spelled.hello.peers.length, 1, '/p/%61 is /p/a');
    const more = await joined(port, { v: ['v1'] }, '/%70/a?x=%2f');
    assert.equal(more.hello.peers.length, 2, 'the query is not the path');
    const wide = await joined(port, { v: ['v1'] }, '/p/%C5%9F%20x');
    const lower = await joined(port, { v: ['v1'] }, '/p/%c5%9f%20x');
    assert.equal(lower.hello.peers.length, 1, 'the case of the digits does not matter');
    // It is decoded once: what is left is a name.
    const twice = await joined(port, { v: ['v1'] }, '/p/%2561');
    assert.equal(twice.hello.peers.length, 0);
    assert.deepEqual(
      server.rooms().map((room) => room.key),
      ['/p/a', '/p/ş x', '/p/%61'],
    );
    plain.c.end();
    wide.c.end();
  });
  await withServer({ path: '/collab/' }, async (port) => {
    const spelled = await connect(port, { path: '/%63ollab/doc' });
    assert.equal(spelled.status, 101, 'the prefix is looked for in the decoded path');
    spelled.end();
    const other = await connect(port, { path: '/collab%2fdoc' });
    assert.equal(other.status, 400, 'an encoded slash does not make the prefix');
    await other.gone();
  });
  await withServer({ path: '/collab/' }, async (port) => {
    const out = await connect(port, { path: '/collab/../other' });
    assert.equal(out.status, 404, 'a path that leaves the prefix is not under it');
    await out.gone();
    const inside = await connect(port, { path: '/other/../collab/doc' });
    assert.equal(inside.status, 101);
    inside.end();
  });
});

// --- Hardening: the warning --------------------------------------------------

/** What `console.warn` was told while `body` ran. */
async function warned(body: () => Promise<void>): Promise<string[]> {
  const said: string[] = [];
  const original = console.warn;
  console.warn = (...parts: unknown[]) => void said.push(parts.join(' '));
  try {
    await body();
  } finally {
    console.warn = original;
  }
  return said;
}

const WARNING = /anyone who can reach this relay can read and change every file edited through it; set `authorize` \(and `origins`\) in production/;

test('a relay without `authorize` and without `origins` says so, once', async () => {
  const web = http.createServer();
  await new Promise<void>((resolve) => web.listen(0, '127.0.0.1', resolve));
  try {
    const said = await warned(async () => {
      const server = createCollabServer();
      const { port } = await server.listen(0, '127.0.0.1');
      server.attach(web);
      await joined(port);
      await server.close();
    });
    assert.equal(said.length, 1, `one warning, not ${JSON.stringify(said)}`);
    assert.match(said[0], WARNING);

    const attached = await warned(async () => {
      const server = createCollabServer();
      server.attach(web);
      server.attach(web);
      await server.close();
    });
    assert.equal(attached.length, 1, 'attach() warns as listen() does');
    assert.match(attached[0], WARNING);

    const made = await warned(async () => void createCollabServer());
    assert.deepEqual(made, [], 'a relay that serves nothing has nothing to warn of');

    // With a log, the log is told.
    const events: any[] = [];
    const logged = await warned(async () => {
      const server = createCollabServer({ log: (event) => events.push(event) });
      await server.listen(0, '127.0.0.1');
      server.attach(web);
      await server.close();
    });
    assert.deepEqual(logged, []);
    const warnings = events.filter((event) => event.type === 'warning');
    assert.equal(warnings.length, 1);
    assert.match(warnings[0].message, WARNING);

    const options: Record<string, CollabServerOptions> = {
      authorize: { authorize: () => true },
      origins: { origins: ['https://app.example'] },
      both: { authorize: () => true, origins: ['https://app.example'] },
    };
    for (const name of Object.keys(options)) {
      const seen: any[] = [];
      const quiet = await warned(async () => {
        const server = createCollabServer(options[name]);
        await server.listen(0, '127.0.0.1');
        await server.close();
        const told = createCollabServer({ ...options[name], log: (event) => seen.push(event) });
        await told.listen(0, '127.0.0.1');
        await told.close();
      });
      assert.deepEqual(quiet, [], `with ${name} there is no warning`);
      assert.deepEqual(seen.filter((event) => event.type === 'warning'), [], `with ${name} the log is told of none`);
    }
  } finally {
    await new Promise((resolve) => web.close(resolve));
  }
});

// --- Garbage -----------------------------------------------------------------

/** A generator that gives the same garbage on every run. */
function random(seed: number): () => number {
  let state = seed >>> 0;
  return () => {
    state = (state + 0x6d2b79f5) >>> 0;
    let t = state;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

test('garbage does not bring the server down', async () => {
  const next = random(20261013);
  const pick = <T>(list: T[]): T => list[Math.floor(next() * list.length)];
  const noise = (length: number): Buffer => {
    const bytes = Buffer.alloc(length);
    for (let i = 0; i < length; i++) bytes[i] = Math.floor(next() * 256);
    return bytes;
  };
  const value = (depth: number): unknown => {
    const kind = Math.floor(next() * (depth > 3 ? 6 : 8));
    if (kind === 0) return null;
    if (kind === 1) return next() < 0.5;
    if (kind === 2) return pick([0, 1, -1, 1.5, 2 ** 53, 1e300, 3]);
    if (kind === 3) return pick(['', 'v1', 'join', '__proto__', 'x'.repeat(300), '\u0000', '\ud83d']);
    if (kind === 4) return Math.floor(next() * 5);
    if (kind === 5) return pick(['constructor', 'toString', 'op']);
    if (kind === 6) return Array.from({ length: Math.floor(next() * 4) }, () => value(depth + 1));
    const object: Record<string, unknown> = {};
    for (let i = Math.floor(next() * 4); i > 0; i--) object[pick(FIELDS)] = value(depth + 1);
    return object;
  };
  const FIELDS = ['t', 'v', 'since', 'user', 'token', 'ro', 'room', 'base', 'id', 'ops', 'data', 'rev', '__proto__', 'name'];
  const message = (): unknown => {
    const object: Record<string, unknown> = { t: pick(['join', 'op', 'presence', 'saved', 'ping', 'hello', 7, null]) };
    for (let i = Math.floor(next() * 6); i > 0; i--) object[pick(FIELDS)] = value(0);
    return object;
  };
  await withServer({ maxMessageBytes: 4096, joinTimeoutMs: 2000 }, async (port, server) => {
    for (let round = 0; round < 240; round++) {
      const c = await connect(port, { path: `/fuzz/${round % 5}` });
      const kind = round % 4;
      if (kind === 0) {
        // Bytes that were never a frame.
        c.bytes(noise(1 + Math.floor(next() * 600)));
      } else if (kind === 1) {
        // Frames that are well made around a payload that is not.
        for (let i = 0; i < 4; i++) {
          c.bytes(frameBytes(pick([0, 1, 1, 1, 2, 8, 9, 10, 5]), noise(Math.floor(next() * 200)), { fin: next() < 0.8 }));
        }
      } else {
        // Messages that are JSON and nothing more, some after an honest join.
        if (kind === 3) c.send({ t: 'join', v: ['v1'] });
        for (let i = 0; i < 6; i++) c.send(message());
        if (next() < 0.3) c.bytes(frameBytes(1, '['.repeat(4000)));
      }
      c.send({ t: 'ping' });
      // Either it is still served or it was closed: both are fine. A hang is not.
      await Promise.race([c.gone(), new Promise((resolve) => setTimeout(resolve, 15))]);
      c.end();
    }
    assert.equal(server.rooms().length <= 6, true, 'garbage made no rooms beyond the paths it used');
    const a = await joined(port, { v: ['honest'] }, '/after');
    const b = await joined(port, { v: ['honest'] }, '/after');
    a.c.send({ t: 'op', base: 0, id: a.c.id('fine'), ops: [{ ok: true }] });
    assert.deepEqual(await b.c.json(), { t: 'op', rev: 1, id: a.c.id('fine'), from: a.hello.you, ops: [{ ok: true }] });
  });
});

test('operations nested deeper than 256 levels are a bad message, however deep', async () => {
  await withServer({}, async (port) => {
    const nested = (id: string, depth: number): Buffer =>
      frameBytes(1, `{"t":"op","base":0,"id":"${id}","ops":${'['.repeat(depth - 1)}{"k":"[[[{{{"}${']'.repeat(depth - 1)}}`);
    const { c } = await joined(port);
    c.bytes(nested('deepest', 256));
    assert.equal((await c.json()).rev, 1, '256 levels, `ops` counted, are passed on');
    for (const depth of [257, 10000, 200000]) {
      const odd = await joined(port, { v: ['v1'] }, `/deep/${depth}`);
      odd.c.bytes(nested('deep', depth));
      await odd.c.closes(1008, 'bad-message').catch((error) => {
        throw new Error(`${depth} levels: ${error.message}`);
      });
    }
    const after = await joined(port, { v: ['v1'] }, '/after');
    await after.c.settled();
  });
});

// --- Shutting down -----------------------------------------------------------

test('close() closes every socket, forgets every room and can be called twice', async () => {
  const server = createCollabServer({ idleRoomMs: 60000, pingIntervalMs: 60000, joinTimeoutMs: 60000 });
  const { port } = await server.listen(0, '127.0.0.1');
  const a = await joined(port);
  const b = await joined(port, { v: ['v1'] }, '/other');
  const silent = await connect(port);
  b.c.end();
  await until('the room to be empty', () => server.rooms().some((room) => room.connections === 0));
  await server.close();
  await a.c.gone();
  await silent.gone();
  assert.deepEqual(server.rooms(), []);
  await server.close();
  await assert.rejects(connect(port), 'nothing listens any more');
});

function run(args: string[], input: { killOn?: RegExp } = {}): Promise<{ code: number | null; out: string; err: string }> {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    let out = '';
    let err = '';
    let asked = false;
    const guard = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`the process did not end by itself\nstdout: ${out}\nstderr: ${err}`));
    }, WAIT);
    child.stdout.on('data', (data) => {
      out += data;
      if (input.killOn && !asked && input.killOn.test(out)) {
        asked = true;
        child.kill('SIGTERM');
      }
    });
    child.stderr.on('data', (data) => (err += data));
    child.on('error', reject);
    child.on('close', (code) => {
      clearTimeout(guard);
      resolve({ code, out, err });
    });
  });
}

test('after close() nothing keeps the process alive, and importing the module listens to nothing', async () => {
  const script = `
    import net from 'node:net';
    import { createCollabServer } from ${JSON.stringify(pathToFileURL(SERVER_FILE).href)};
    const server = createCollabServer();
    const { port } = await server.listen(0, '127.0.0.1');
    const socket = net.connect(port, '127.0.0.1');
    socket.on('error', () => {});
    socket.write('GET /doc HTTP/1.1\\r\\nHost: x\\r\\nUpgrade: websocket\\r\\nConnection: Upgrade\\r\\n'
      + 'Sec-WebSocket-Key: dGhlIHNhbXBsZSBub25jZQ==\\r\\nSec-WebSocket-Version: 13\\r\\n\\r\\n');
    await new Promise((resolve) => socket.once('data', resolve));
    console.log('rooms ' + server.rooms().length);
    await server.close();
    console.log('closed');
  `;
  const result = await run(['--input-type=module', '-e', script]);
  assert.match(result.err, WARNING, 'it was made without `authorize` and without `origins`');
  assert.equal(result.err.split('\n').filter(Boolean).length, 1, `one line of warning, not ${result.err}`);
  assert.equal(result.out, 'rooms 0\nclosed\n');
  assert.equal(result.code, 0);
});

test('run as a program it listens where it is told and stops when it is asked to', async () => {
  const result = await run([SERVER_FILE, '--port', '0', '--host', '127.0.0.1', '--origin', 'https://app.example'], {
    killOn: /listening on ws:\/\/127\.0\.0\.1:\d+/,
  });
  assert.match(result.out, /listening on ws:\/\/127\.0\.0\.1:\d+/);
  assert.equal(result.err, '');
  assert.equal(result.code, 0, 'SIGTERM ends it cleanly');

  const open = await run([SERVER_FILE, '--port', '0', '--host', '127.0.0.1'], {
    killOn: /listening on ws:\/\/127\.0\.0\.1:\d+/,
  });
  assert.match(open.err, WARNING, 'without --origin it warns');
  assert.equal(open.code, 0);

  const wrong = await run([SERVER_FILE, '--port', 'eighty']);
  assert.equal(wrong.code, 2);
  assert.match(wrong.err, /--port/);
});

test('without a host, the program and listen() listen on 127.0.0.1 only', async () => {
  const program = await run([SERVER_FILE, '--port', '0', '--origin', 'https://app.example'], {
    killOn: /listening on ws:\/\/\S+:\d+/,
  });
  assert.match(program.out, /listening on ws:\/\/127\.0\.0\.1:\d+/);
  assert.equal(program.code, 0);

  const server = createCollabServer({ origins: ['https://app.example'] });
  try {
    const { host } = await server.listen(0);
    assert.equal(host, '127.0.0.1');
  } finally {
    await server.close();
  }
});

// --- The client of the platform ----------------------------------------------

test('the WebSocket of the platform talks to the server', async (t) => {
  const WebSocketClass = (globalThis as any).WebSocket;
  if (typeof WebSocketClass !== 'function') return t.skip('no global WebSocket');
  await withServer({}, async (port, server) => {
    const open = (path: string): Promise<{ socket: any; inbox: any[]; next(): Promise<any> }> =>
      new Promise((resolve, reject) => {
        const socket = new WebSocketClass(`ws://127.0.0.1:${port}${path}`);
        const inbox: any[] = [];
        let wake: (() => void) | null = null;
        socket.onmessage = (event: any) => {
          inbox.push(JSON.parse(event.data));
          wake?.();
        };
        socket.onerror = () => reject(new Error('the WebSocket failed'));
        socket.onopen = () =>
          resolve({
            socket,
            inbox,
            async next() {
              const end = Date.now() + WAIT;
              while (!inbox.length) {
                if (Date.now() > end) throw new Error('timed out waiting for a message');
                await new Promise<void>((done) => {
                  const timer = setTimeout(done, 50);
                  wake = () => {
                    clearTimeout(timer);
                    done();
                  };
                });
              }
              return inbox.shift();
            },
          });
      });
    const a = await open('/doc?x=1');
    const b = await open('/doc');
    a.socket.send(JSON.stringify({ t: 'join', v: ['v1'], key: 'a', user: { name: 'Ada' } }));
    const helloA = await a.next();
    b.socket.send(JSON.stringify({ t: 'join', v: ['v1'], key: 'b' }));
    const helloB = await b.next();
    assert.deepEqual(helloB.peers, [{ id: helloA.you, user: { name: 'Ada' }, ro: false }]);
    assert.equal((await a.next()).t, 'peer');
    const big = 'ğ'.repeat(100000);
    a.socket.send(JSON.stringify({ t: 'op', base: 0, id: 'a:1', ops: [big] }));
    assert.deepEqual(await b.next(), { t: 'op', rev: 1, id: 'a:1', from: helloA.you, ops: [big] });
    assert.equal((await a.next()).rev, 1);
    b.socket.send(JSON.stringify({ t: 'op', base: 0, id: 'b:1', ops: [1] }));
    assert.deepEqual(await b.next(), { t: 'nack', id: 'b:1', rev: 1 });
    const closed = new Promise<any>((resolve) => (b.socket.onclose = resolve));
    b.socket.close(1000, 'done');
    const event = await closed;
    assert.equal(event.code, 1000);
    assert.equal(event.wasClean, true, 'the closing handshake was completed');
    assert.deepEqual(await a.next(), { t: 'left', id: helloB.you });
    await until('the room to have one connection', () => server.rooms()[0].connections === 1);
    a.socket.close();
  });
});

// --- Second hardening --------------------------------------------------------

/** Every string and every name of a property in `value`. */
function textsOf(value: unknown): string[] {
  const seen = new Set<unknown>();
  const texts: string[] = [];
  const walk = (item: unknown, depth: number): void => {
    if (typeof item === 'string') texts.push(item);
    if (typeof item !== 'object' || item === null || seen.has(item) || depth > 8) return;
    seen.add(item);
    for (const name of Object.keys(item)) {
      texts.push(name);
      walk((item as any)[name], depth + 1);
    }
  };
  walk(value, 0);
  return texts;
}

test('the key of a join is in no message of the server, in nothing it logs and not in rooms()', async () => {
  const mark = 'K3y-0f-';
  const events: unknown[] = [];
  await withServer({ log: (event) => events.push(event), savedPerMinute: 0, syncPerMinute: 0 }, async (port, server) => {
    const clients: Client[] = [];
    const a = await joined(port, { v: ['v1'], key: `${mark}ada`, user: { name: 'Ada' } });
    const b = await joined(port, { v: ['v1'], key: `${mark}bob` });
    const late: Client[] = [a.c, b.c];
    assert.equal((await a.c.json()).t, 'peer');
    a.c.send({ t: 'op', base: 0, id: 'first', ops: [1] });
    assert.equal((await b.c.json()).t, 'op');
    a.c.send({ t: 'op', base: 0, id: 'stale', ops: [1] });
    a.c.send({ t: 'presence', data: { at: 1 } });
    assert.equal((await b.c.json()).t, 'presence');
    a.c.send({ t: 'saved', v: ['v2'], rev: 1 });
    assert.equal((await b.c.json()).t, 'saved');
    // One that is reset and then gets in, one that is replaced, one that is turned away for what it says.
    const reset = await knock(port, { v: ['v0'], key: `${mark}cem` });
    assert.equal(reset.answer.t, 'reset');
    reset.c.send({ t: 'join', v: ['v2'], since: 1, ep: a.hello.ep, key: `${mark}cem` });
    assert.equal((await reset.c.json()).t, 'hello');
    /** Whatever else it was told before, the connection is turned away with `code`. */
    const turned = async (c: Client, code: string): Promise<void> => {
      for (;;) {
        const message = await c.json();
        if (message.t !== 'error') continue;
        assert.deepEqual(message, { t: 'error', code });
        return c.gone();
      }
    };
    const again = await joined(port, { v: ['v1'], key: `${mark}bob` });
    await turned(b.c, 'replaced');
    const odd = await joined(port, { v: ['v1'], key: `${mark}odd` });
    odd.c.send({ t: 'nonsense', key: `${mark}odd` });
    await turned(odd.c, 'bad-message');
    const second = await joined(port, { v: ['v1'], key: `${mark}two` });
    second.c.send({ t: 'join', v: ['v1'], key: `${mark}two` });
    await turned(second.c, 'bad-message');
    const deep = await joined(port, { v: ['v1'], key: `${mark}deep` });
    deep.c.bytes(frameBytes(1, `{"t":"presence","key":"${mark}deep","data":${'['.repeat(200000)}${']'.repeat(200000)}}`));
    await deep.c.settled().catch(() => undefined);
    clients.push(...late, reset.c, again.c, odd.c, second.c, deep.c);
    await a.c.settled().catch(() => undefined);
    assert.deepEqual(
      textsOf(server.rooms()).filter((text) => text.includes(mark)),
      [],
      'rooms() carries a key',
    );
    a.c.end();
    again.c.end();
    reset.c.end();
    deep.c.end();
    await until('everyone to be gone', () => server.rooms().every((room) => room.connections === 0));

    let messages = 0;
    for (const client of clients) {
      for (const text of client.seen) {
        messages++;
        assert.equal(text.includes(mark), false, `a message carries a key: ${text.slice(0, 300)}`);
      }
    }
    assert.equal(messages > 15, true, `the clients were told things: ${messages} messages`);
    assert.equal(events.length > 15, true, `the log was written to: ${events.length} events`);
    assert.deepEqual(
      textsOf(events).filter((text) => text.includes(mark)),
      [],
      'the log carries a key',
    );
  });
});

/** The next message that is not about who comes and goes or a call to save. */
async function next(c: Client): Promise<any> {
  for (;;) {
    const message = await c.json();
    if (message.t !== 'peer' && message.t !== 'left' && message.t !== 'sync') return message;
  }
}

/** Says `saved` and gives back what `listener` hears of it. */
async function says(speaker: Client, listener: Client, v: string[], rev: number): Promise<any> {
  speaker.send({ t: 'saved', v, rev });
  const heard = await next(listener);
  assert.equal(heard.t, 'saved', `a save was to be passed on, ${JSON.stringify(heard)} came`);
  assert.deepEqual(heard.v, v, 'the versions are passed on whatever becomes of the revision');
  return heard;
}

test('a version that fell out of the full table cannot be written again, at any revision', async () => {
  await withServer({ savedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['e0'] });
    const b = await joined(port, { v: ['e0'] });
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 3);
    for (let rev = 1; rev <= 3; rev++) assert.equal((await b.c.json()).rev, rev);
    assert.equal((await says(a.c, b.c, ['e1', 'h1'], 3)).rev, 3, 'the true save');
    await operations(a.c, 3, 6);
    for (let rev = 4; rev <= 9; rev++) assert.equal((await b.c.json()).rev, rev);
    const ep = a.hello.ep;
    assert.deepEqual((await comesBack(port, ep, 3, ['e1'])).v, ['e1'], 'while it is in the table, the room has heard of e1');

    // One who may write says so straight away: the table says otherwise.
    assert.equal('rev' in (await says(b.c, a.c, ['e1'], 9)), false, 'e1 is of revision 3');
    // It fills the table until e1 falls out, and says it again.
    let junk = 0;
    for (let message = 0; message < 4; message++) {
      const names = [0, 1, 2, 3, 4, 5, 6, 7].map(() => `junk${junk++}`);
      assert.equal((await says(b.c, a.c, names, 9)).rev, 9);
    }
    assert.deepEqual((await comesBack(port, ep, 3, ['e1'])).v, [], 'e1 fell out of the table');
    for (const version of ['e1', 'h1', 'junk0']) {
      for (const rev of [9, 3, 0]) {
        const heard = await says(b.c, a.c, [version], rev);
        assert.equal('rev' in heard, false, `${version} is written at ${rev} after it fell out: ${JSON.stringify(heard)}`);
      }
    }
    // One of them spoils a save, as a version that is taken does.
    assert.equal('rev' in (await says(b.c, a.c, ['brand-new', 'e1'], 9)), false);
    assert.deepEqual(
      (await comesBack(port, ep, 9, ['e1', 'h1', 'junk0', 'brand-new', 'junk31'])).v,
      ['junk31'],
      'whoever has one of them is not told that the room knows its file',
    );
    // The honest save that follows is written as before.
    a.c.send({ t: 'op', base: 9, id: a.c.id('op9'), ops: [1] });
    assert.equal((await next(a.c)).rev, 10);
    assert.equal((await next(b.c)).rev, 10);
    assert.equal((await says(a.c, b.c, ['e2'], 10)).rev, 10);
    assert.deepEqual((await comesBack(port, ep, 10, ['e2'])).v, ['e2']);
  });
});

test('a version that the log moved past cannot be written again', async () => {
  await withServer({ maxLogEntries: 2, savedPerMinute: 0 }, async (port, server) => {
    const a = await joined(port, { v: ['orig'] });
    const b = await joined(port, { v: ['orig'] });
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 3);
    for (let rev = 1; rev <= 3; rev++) assert.equal((await b.c.json()).rev, rev);
    assert.equal(server.rooms()[0].floor, 1);
    assert.equal((await says(a.c, b.c, ['s1'], 1)).rev, 1);
    assert.equal((await says(a.c, b.c, ['s3'], 3)).rev, 3);
    await operations(a.c, 3, 2);
    for (let rev = 4; rev <= 5; rev++) assert.equal((await b.c.json()).rev, rev);
    assert.equal(server.rooms()[0].floor, 3, 'revision 1 is no longer in the log');
    for (const rev of [5, 3, 4]) {
      assert.equal('rev' in (await says(a.c, b.c, ['s1'], rev)), false, `s1 is written at ${rev} after the log moved past it`);
    }
    const stale = await knock(port, { v: ['s1'], ro: true });
    assert.deepEqual(stale.answer, { t: 'reset', rev: 5, v: ['s3'] });
    assert.deepEqual((await comesBack(port, a.hello.ep, 3, ['s1', 's3'])).v, ['s3'], 's3 is still in the log');
  });
});

test('the versions of the foundation and those of the last save do not fall out of the full table', async () => {
  await withServer({ savedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['f0', 'f1'] });
    const b = await joined(port, { v: ['f0'] });
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 2);
    assert.equal((await b.c.json()).rev, 1);
    assert.equal((await b.c.json()).rev, 2);
    assert.equal((await says(a.c, b.c, ['x'], 1)).rev, 1);
    assert.equal((await says(a.c, b.c, [1, 2, 3, 4, 5, 6, 7].map((index) => `j${index}`), 1)).rev, 1);
    assert.equal((await says(a.c, b.c, [8, 9, 10, 11, 12, 13].map((index) => `j${index}`), 1)).rev, 1);
    // The table is full: f0 f1 x j1..j13. This save names x, the oldest, again, and seven new ones.
    const last = ['x', 'n1', 'n2', 'n3', 'n4', 'n5', 'n6', 'n7'];
    assert.equal((await says(a.c, b.c, last, 1)).rev, 1);
    const ep = a.hello.ep;
    assert.deepEqual(
      (await comesBack(port, ep, 1, ['x', 'n7', 'j8', 'j7', 'j1'])).v,
      ['x', 'n7', 'j8'],
      'x is of the last save: it did not make room; the seven oldest of the others did',
    );
    // A save of eight new ones: the last save is this one now, and the foundation stays.
    const newest = [1, 2, 3, 4, 5, 6, 7, 8].map((index) => `m${index}`);
    assert.equal((await says(a.c, b.c, newest, 2)).rev, 2);
    assert.deepEqual(
      (await comesBack(port, ep, 2, ['m1', 'm8', 'f1', 'f0', 'x'])).v,
      ['m1', 'm8', 'f1', 'f0'],
      'x is of the last save no more; the foundation is kept',
    );
    assert.equal((await joined(port, { v: ['f1'] })).hello.from, 0, 'the foundation is joined from');
    assert.equal((await joined(port, { v: ['f0'] })).hello.from, 0);
  });
});

test('a room remembers only the latest versions that fell out, and goes on writing saves down', async () => {
  await withServer({ maxEvictedVersions: 2, savedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['f'] });
    const b = await joined(port, { v: ['f'] });
    const ep = a.hello.ep;
    assert.equal((await a.c.json()).t, 'peer', 'the first hears of the second');
    await operations(a.c, 0, 2);
    assert.equal((await b.c.json()).rev, 1, 'the first operation reaches the other');
    assert.equal((await b.c.json()).rev, 2, 'the second operation reaches the other');
    const names = (prefix: string): string[] => [0, 1, 2, 3, 4, 5, 6, 7].map((index) => `${prefix}${index}`);
    assert.equal((await says(a.c, b.c, names('a'), 1)).rev, 1, 'eight versions beside the foundation');
    assert.equal((await says(a.c, b.c, names('b'), 1)).rev, 1, 'one fell out: a0');
    assert.equal((await says(a.c, b.c, names('c'), 1)).rev, 1, 'eight more fell out: a1 to a7 and b0, of which a7 and b0 are remembered');
    // Nine have left and two are remembered: the room writes a new save down all the same.
    const heard = await says(a.c, b.c, ['new'], 2);
    assert.equal(heard.rev, 2, `a room that has forgotten versions still writes a save down: ${JSON.stringify(heard)}`);
    assert.deepEqual((await comesBack(port, ep, 2, ['new'])).v, ['new'], 'whoever comes back with the newest save is told the room has heard of it');
    // `new` pushed b1 out: b0 and b1 are the two that left last, and neither is written again.
    for (const version of ['b1', 'b0']) {
      for (const rev of [2, 1]) {
        const again = await says(a.c, b.c, [version], rev);
        assert.equal('rev' in again, false, `${version} left a moment ago and is written at ${rev}: ${JSON.stringify(again)}`);
      }
    }
    assert.equal('rev' in (await says(b.c, a.c, ['b1'], 2)), false, 'whoever says it');
    assert.equal('rev' in (await says(a.c, b.c, ['other', 'b0'], 2)), false, 'one of them spoils a save');
    assert.deepEqual((await comesBack(port, ep, 2, ['b0', 'b1', 'other'])).v, [], 'none of them is in the table');
    // a0 left long ago and is remembered no more: it is written like a version the room never had.
    assert.equal((await says(a.c, b.c, ['a0'], 2)).rev, 2, 'a version that was pushed out of the memory can be written again');
    assert.deepEqual((await comesBack(port, ep, 2, ['a0', 'new', 'b2'])).v, ['a0', 'new'], 'a0 is in the table again, and b2, the oldest in it, made room');
    // And the room goes on like that: every new save is written down and found by who comes back with it.
    for (let save = 0; save < 6; save++) {
      const versions = names(`n${save}-`);
      assert.equal((await says(save % 2 ? b.c : a.c, save % 2 ? a.c : b.c, versions, 2)).rev, 2, `save ${save} after many versions have left`);
      assert.deepEqual((await comesBack(port, ep, 2, versions)).v, versions, `save ${save} is in the table`);
    }
    // A save that only says what the table says is passed on whole, and the foundation is joined from.
    assert.equal((await says(a.c, b.c, ['n5-0', 'n5-7'], 2)).rev, 2, 'what the table says is passed on with its revision');
    assert.equal('rev' in (await says(a.c, b.c, ['n5-0'], 1)), false, 'the table says otherwise of n5-0');
    assert.equal((await joined(port, { v: ['f'] })).hello.from, 0, 'the foundation is joined from');
  });
  // To remember none of them is to write no save down at all.
  await withServer({ maxEvictedVersions: 0, savedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['f'] });
    const b = await joined(port, { v: ['f'] });
    assert.equal((await a.c.json()).t, 'peer', 'the first hears of the second');
    await operations(a.c, 0, 1);
    assert.equal((await b.c.json()).rev, 1, 'the operation reaches the other');
    for (const rev of [1, 0]) {
      const heard = await says(a.c, b.c, ['s1'], rev);
      assert.equal('rev' in heard, false, `with maxEvictedVersions 0 a save was written down at ${rev}: ${JSON.stringify(heard)}`);
    }
    assert.equal('rev' in (await says(b.c, a.c, ['s2'], 1)), false, 'whoever says it');
    assert.deepEqual((await comesBack(port, a.hello.ep, 1, ['s1', 's2', 'f'])).v, ['f'], 'only the foundation is in the table');
    assert.equal((await says(a.c, b.c, ['f'], 0)).rev, 0, 'what the table says is passed on with its revision');
    assert.equal((await joined(port, { v: ['f'] })).hello.from, 0, 'the foundation is joined from');
  });
});

test('a client that saves every two seconds for ten minutes is always known by its newest save', async () => {
  // The defaults, and a clock that is moved by hand: what saves by itself saves every two seconds at most.
  let time = 1000;
  await withServer({ now: () => time }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    const ep = a.hello.ep;
    assert.equal((await a.c.json()).t, 'peer', 'the first hears of the second');
    const SAVES = 300;
    for (let n = 1; n <= SAVES; n++) {
      time += 2000;
      // Those who came back to look are still in the room: what is said of them is skipped.
      a.c.send({ t: 'op', base: n - 1, id: a.c.id(`op${n}`), ops: [n] });
      assert.equal((await next(a.c)).rev, n, `operation ${n} comes back to its sender`);
      assert.equal((await next(b.c)).rev, n, `operation ${n} reaches the other`);
      const versions = [`e:${n}`, `h:${n}`];
      const heard = await says(a.c, b.c, versions, n);
      assert.equal(heard.rev, n, `save ${n} of ${SAVES} was passed on without its revision: ${JSON.stringify(heard)}`);
      if (n <= SAVES - 20) continue;
      const hello = await comesBack(port, ep, n, versions);
      assert.deepEqual(hello.v, versions, `the room has not heard of save ${n}, which it was just told of: ${JSON.stringify(hello)}`);
    }
    assert.equal(a.c.closed, false, 'the connection that saved is not closed');
  });
});

test('a connection has only so many saves a minute written down', async () => {
  let time = 5000;
  await withServer({ now: () => time, savedPerMinute: 6 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 1);
    assert.equal((await b.c.json()).rev, 1);
    for (let i = 1; i <= 6; i++) assert.equal((await says(a.c, b.c, [`s${i}`], 1)).rev, 1, `save ${i} of the six`);
    // What only repeats the table costs nothing; what is refused for another reason neither.
    assert.equal((await says(a.c, b.c, ['s1', 's6'], 1)).rev, 1);
    assert.equal('rev' in (await says(a.c, b.c, ['s1'], 0)), false);
    const seventh = await says(a.c, b.c, ['s7'], 1);
    assert.equal('rev' in seventh, false, `the seventh save of the minute was written down: ${JSON.stringify(seventh)}`);
    assert.equal(a.c.closed, false, 'the connection is not closed');
    assert.deepEqual((await comesBack(port, a.hello.ep, 1, ['s6', 's7'])).v, ['s6'], 's7 is not in the table');
    // The limit is of the connection.
    assert.equal((await says(b.c, a.c, ['t1'], 1)).rev, 1);
    time += 9000;
    assert.equal('rev' in (await says(a.c, b.c, ['s8'], 1)), false, 'nine seconds are not a save');
    time += 2000;
    assert.equal((await says(a.c, b.c, ['s9'], 1)).rev, 1, 'eleven seconds are');
    assert.equal('rev' in (await says(a.c, b.c, ['s10'], 1)), false, 'and only one');
    // What was not written for its rate was not spoiled: it can be said again.
    time += 10000;
    assert.equal((await says(a.c, b.c, ['s7'], 1)).rev, 1);
    assert.deepEqual((await comesBack(port, a.hello.ep, 1, ['s7'])).v, ['s7']);
  });
  await withServer({ now: () => time, savedPerMinute: 2 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await says(a.c, b.c, ['s1'], 0)).rev, 0);
    assert.equal((await says(a.c, b.c, ['s2'], 0)).rev, 0);
    assert.equal('rev' in (await says(a.c, b.c, ['s3'], 0)), false, 'the rate is what the option says');
  });
  // Without the option a connection has forty a minute, and the room more than that.
  await withServer({ now: () => time }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer', 'the first hears of the second');
    for (let i = 1; i <= 40; i++) assert.equal((await says(a.c, b.c, [`s${i}`], 0)).rev, 0, `save ${i} of the default forty`);
    const over = await says(a.c, b.c, ['s41'], 0);
    assert.equal('rev' in over, false, `the default is forty a minute for a connection: ${JSON.stringify(over)}`);
    assert.equal((await says(b.c, a.c, ['t1'], 0)).rev, 0, 'the room has more than one connection may use');
  });
});

test('a connection adds only so many formats a minute', async () => {
  let time = 0;
  const table = (count: number): unknown => ({ t: 'st', f: new Array(count).fill({ b: 1 }) });
  await withServer({ now: () => time, formatsPerMinute: 10 }, async (port, server) => {
    const a = await joined(port);
    const b = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    const send = async (c: Client, base: number, id: string, ops: unknown[]): Promise<any> => {
      c.send({ t: 'op', base, id, ops });
      return c.json();
    };
    assert.equal((await send(a.c, 0, 'six', [table(6), { n: 1 }])).rev, 1, 'six of ten');
    assert.deepEqual(await send(a.c, 1, 'five', [table(5)]), { t: 'nack', id: 'five', rev: 1, reason: 'rate' }, 'four are left');
    assert.equal(server.rooms()[0].rev, 1, 'what is refused is not placed');
    // An operation that is not placed adds nothing, and costs nothing.
    assert.deepEqual(await send(a.c, 0, 'stale', [table(4)]), { t: 'nack', id: 'stale', rev: 1 });
    assert.equal((await send(a.c, 1, 'four', [table(4)])).rev, 2, 'the refusal took nothing from the bucket');
    assert.deepEqual(await send(a.c, 2, 'one', [table(1)]), { t: 'nack', id: 'one', rev: 2, reason: 'rate' });
    // Only the first operation of a message is looked at, and only when it is a table of formats.
    assert.equal((await send(a.c, 2, 'none', [table(0)])).rev, 3);
    assert.equal((await send(a.c, 3, 'later', [{ n: 1 }, table(50)])).rev, 4);
    assert.equal((await send(a.c, 4, 'other', [{ t: 'st', f: 'text' }, { t: 'sx', f: [1, 2, 3] }])).rev, 5);
    // The limit is of the connection.
    for (let rev = 1; rev <= 5; rev++) assert.equal((await b.c.json()).rev, rev);
    assert.equal((await send(b.c, 5, 'ten', [table(10)])).rev, 6);
    assert.equal((await a.c.json()).rev, 6);
    // The bucket fills at ten a minute, up to ten.
    time += 30000;
    assert.deepEqual(await send(a.c, 6, 'early', [table(6)]), { t: 'nack', id: 'early', rev: 6, reason: 'rate' });
    assert.equal((await send(a.c, 6, 'half', [table(5)])).rev, 7, 'half a minute is five formats');
    time += 600000;
    assert.deepEqual(await send(a.c, 7, 'huge', [table(11)]), { t: 'nack', id: 'huge', rev: 7, reason: 'rate' }, 'more than a full bucket');
    assert.equal((await send(a.c, 7, 'full', [table(10)])).rev, 8);
    assert.equal(a.c.closed, false, 'the connection is not closed');
  });
  await withServer({ now: () => time }, async (port) => {
    const { c } = await joined(port);
    c.send({ t: 'op', base: 0, id: 'most', ops: [table(2000)] });
    assert.equal((await c.json()).rev, 1, 'the default is two thousand');
    c.send({ t: 'op', base: 1, id: 'more', ops: [table(1)] });
    assert.deepEqual(await c.json(), { t: 'nack', id: 'more', rev: 1, reason: 'rate' });
  });
  await withServer({ now: () => time, formatsPerMinute: 0 }, async (port) => {
    const { c } = await joined(port);
    c.send({ t: 'op', base: 0, id: 'any', ops: [table(5000)] });
    assert.equal((await c.json()).rev, 1, '0 is no limit');
  });
});

test('the operations are passed on in the very text they came in', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port);
    const b = await joined(port);
    assert.equal((await a.c.json()).t, 'peer');
    const cases: [string, string, string][] = [
      ['numbers and the order of names', '{"t":"op","base":0,"id":"n1","ops":[{"b":1,"a":1e2},1.0,1e400,-0,0.10,1E+2]}', '[{"b":1,"a":1e2},1.0,1e400,-0,0.10,1E+2]'],
      ['white space', '{ "ops" :\t[ 1 ,\n{ "k" : [ ] } ] , "t":"op","base":1,"id":"n2"}', '[ 1 ,\n{ "k" : [ ] } ]'],
      ['escapes', '{"t":"op","base":2,"id":"n3","ops":["\\u00e9\\ud83d\\ude00\\/\\"]}","\\ud800","é😀"]}', '["\\u00e9\\ud83d\\ude00\\/\\"]}","\\ud800","é😀"]'],
      ['a name written twice: the last one counts', '{"t":"op","ops":[1.0],"base":3,"id":"n4","ops":[2.0]}', '[2.0]'],
      ['a name written with an escape', '{"t":"op","base":4,"id":"n5","\\u006fps":[3.0]}', '[3.0]'],
      ['the same text inside a string and inside another value', '{"t":"op","id":"\\"ops\\":[9]","x":{"ops":[8.0]},"base":5,"ops":[4.0],"y":["ops",{"ops":[7.0]}]}', '[4.0]'],
      ['names that mean something to an object', '{"t":"op","base":6,"id":"n7","ops":[{"__proto__":{"x":1.0},"constructor":2.0}]}', '[{"__proto__":{"x":1.0},"constructor":2.0}]'],
    ];
    const told: string[] = [];
    for (const [name, text, ops] of cases) {
      a.c.bytes(frameBytes(1, text));
      const back = await a.c.json();
      assert.equal(back.t, 'op', `${name}: the operation was to be placed, ${JSON.stringify(back)} came`);
      const id = JSON.parse(text).id;
      const expected = `{"t":"op","rev":${back.rev},"id":${JSON.stringify(id)},"from":${JSON.stringify(a.hello.you)},"ops":${ops}}`;
      told.push(expected);
      assert.equal(a.c.seen[a.c.seen.length - 1], expected, `${name}: as the sender hears it`);
      assert.equal((await b.c.json()).rev, back.rev);
      assert.equal(b.c.seen[b.c.seen.length - 1], expected, `${name}: as the other hears it`);
    }
    // The log plays the same text.
    const late = await joined(port);
    for (let i = 0; i < cases.length; i++) assert.equal((await late.c.json()).rev, i + 1);
    assert.deepEqual(late.c.seen.slice(1), told, 'the log');
    assert.equal((await a.c.json()).t, 'peer');
    // What the text says is still checked: `ops` that is no list is a bad message.
    a.c.bytes(frameBytes(1, '{"t":"op","base":7,"id":"n8","ops":[1],"ops":{"0":1}}'));
    await a.c.closes(1008, 'bad-message');
  });
});

test('a room that is founded again forgets nothing it has let go of', async () => {
  await withServer({ savedPerMinute: 0, idleRoomMs: 600000 }, async (port, server) => {
    const empty = (): Promise<void> => until('the room to be empty', () => server.rooms()[0].connections === 0);
    const first = await joined(port, { v: ['f'] });
    const witness = await joined(port, { v: ['f'] });
    assert.equal((await says(first.c, witness.c, ['s'], 0)).rev, 0);
    first.c.end();
    witness.c.end();
    await empty();
    // Nothing was written: whoever comes now founds the room, on its own file.
    const b = await joined(port, { v: ['g'] });
    const c = await joined(port, { v: ['g'] });
    assert.equal(c.hello.from, 0);
    assert.equal((await b.c.json()).t, 'peer');
    await operations(b.c, 0, 2);
    for (const rev of [1, 2]) assert.equal((await c.c.json()).rev, rev);
    for (const version of ['f', 's']) {
      const heard = await says(b.c, c.c, [version], 2);
      assert.equal('rev' in heard, false, `${version}, which the room knew as revision 0, is written at 2: ${JSON.stringify(heard)}`);
      assert.equal((await knock(port, { v: [version], ro: true })).answer.t, 'reset');
    }
    assert.deepEqual((await comesBack(port, b.hello.ep, 2, ['f', 's', 'g'])).v, ['g'], 'the room has heard of neither any more');
  });
  await withServer({ idleRoomMs: 600000 }, async (port, server) => {
    const empty = (): Promise<void> => until('the room to be empty', () => server.rooms()[0].connections === 0);
    for (const version of ['f', 'g']) {
      const { c } = await joined(port, { v: [version] });
      c.end();
      await empty();
    }
    // A room is founded on what the first to come has, whatever it has let go of before.
    const a = await joined(port, { v: ['f'] });
    const b = await joined(port, { v: ['f'] });
    assert.equal(b.hello.from, 0);
    assert.equal((await says(a.c, b.c, ['f'], 0)).rev, 0, 'and that is revision 0 again');
  });
});

// --- The name of a room ---------------------------------------------------------

test('hello tells the name of the room, and a join names it in the letters of a key or is a bad message', async () => {
  await withServer({}, async (port, server) => {
    const a = await joined(port, { v: ['v1'] });
    assert.equal(typeof a.hello.ep, 'string', `hello carries the name of the room: ${JSON.stringify(a.hello)}`);
    assert.match(a.hello.ep, EPOCH, 'the name is made of what a join may send back');
    const b = await joined(port, { v: ['v1'] });
    assert.equal(b.hello.ep, a.hello.ep, 'the second to come is told the same name');
    const other = await joined(port, { v: ['v1'] }, '/other');
    assert.match(other.hello.ep, EPOCH);
    assert.notEqual(other.hello.ep, a.hello.ep, 'another room has another name');

    for (const ep of ['', 'e'.repeat(65), 'a b', 'a:b', 'ş', 'a.b', 'a#b', 'a/b', 7, null, true, ['a'], { ep: 'a' }]) {
      await turnedAway(port, { v: ['v1'], since: 0, ep }, 1008, 'bad-message').catch((error) => {
        throw new Error(`ep ${JSON.stringify(ep)}: ${error.message}`);
      });
      await turnedAway(port, { v: ['v1'], ep }, 1008, 'bad-message', '/nowhere').catch((error) => {
        throw new Error(`ep ${JSON.stringify(ep)} in a room that does not exist: ${error.message}`);
      });
    }
    assert.equal(server.rooms().some((room) => room.key === '/nowhere'), false, 'a join that is refused makes no room');
    // A name that is well made and not the room's is no fault: it is only not this room.
    for (const ep of ['a', 'Az09_-', 'e'.repeat(64)]) {
      const { answer } = await knock(port, { v: ['v1'], since: 0, ep });
      assert.equal(answer.t, 'hello', `ep ${ep}: ${JSON.stringify(answer)}`);
      assert.equal(answer.ep, a.hello.ep, 'hello tells the name the room has, not the one that was sent');
    }
  });
});

test('a `since` without the name of the room, or with another name, is not looked at', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const ep = a.hello.ep;
    await operations(a.c, 0, 3);
    a.c.send({ t: 'saved', v: ['s3'], rev: 3 });
    await a.c.settled();
    const elsewhere = await joined(port, { v: ['v0'] }, '/elsewhere');
    await operations(elsewhere.c, 0, 3);
    assert.notEqual(elsewhere.hello.ep, ep);
    const wrong: [string, object][] = [
      ['no name', {}],
      ['a name nobody has', { ep: 'not-this-room' }],
      ['the name of another room', { ep: elsewhere.hello.ep }],
      ['the name in other letters', { ep: ep.toUpperCase() === ep ? ep.toLowerCase() : ep.toUpperCase() }],
    ];
    for (const [name, extra] of wrong) {
      // A file the room was not founded on: with nothing to go by, it is to be loaded again.
      for (const [v, since] of [[['s3'], 3], [[], 2], [['unknown'], 0], [['s3'], 99]] as [string[], number][]) {
        const { answer } = await knock(port, { v, since, ...extra, ro: true });
        assert.deepEqual(
          answer,
          { t: 'reset', rev: 3, v: ['s3'] },
          `${name}, v ${JSON.stringify(v)}, since ${since}: the join was to be turned back`,
        );
      }
      // The file the room was founded on gets in as it would without `since`: from the beginning.
      const { c, answer } = await knock(port, { v: ['v0', 's3'], since: 3, ...extra, ro: true });
      assert.deepEqual(
        { t: answer.t, ep: answer.ep, from: answer.from, rev: answer.rev, v: answer.v },
        { t: 'hello', ep, from: 0, rev: 3, v: ['v0'] },
        `${name}: the foundation is joined from 0, whatever since says`,
      );
      for (const rev of [1, 2, 3]) assert.equal((await c.json()).rev, rev, `${name}: every operation is played to it`);
      await c.settled();
    }
    // The control: the same joins with the name of the room.
    const right = await knock(port, { v: ['s3'], since: 3, ep, ro: true });
    assert.deepEqual(
      { t: right.answer.t, from: right.answer.from, v: right.answer.v },
      { t: 'hello', from: 3, v: ['s3'] },
      'with the name, since decides',
    );
    await right.c.settled();
    const founder = await knock(port, { v: ['v0', 's3'], since: 3, ep, ro: true });
    assert.deepEqual({ from: founder.answer.from, v: founder.answer.v }, { from: 3, v: ['v0', 's3'] });
    await founder.c.settled();
    const beyond = await knock(port, { v: ['s3'], since: 4, ep, ro: true });
    assert.deepEqual(beyond.answer, { t: 'reset', rev: 3, v: ['s3'] }, 'the name does not make a revision the room has not reached');
  });
});

test('a room that is founded again has another name; a room that is kept keeps its name', async () => {
  await withServer({ idleRoomMs: 600000 }, async (port, server) => {
    const empty = (): Promise<void> => until('the room to be empty', () => server.rooms()[0].connections === 0);
    const first = await joined(port, { v: ['f'] });
    const beside = await joined(port, { v: ['f'] });
    assert.equal(beside.hello.ep, first.hello.ep);
    first.c.end();
    beside.c.end();
    await empty();
    // Nothing was written: with the name it had, the room is still the room it was.
    const same = await joined(port, { v: ['other'], since: 0, ep: first.hello.ep });
    assert.deepEqual(
      { ep: same.hello.ep, from: same.hello.from, v: same.hello.v },
      { ep: first.hello.ep, from: 0, v: [] },
      'a join that names the room does not found it again',
    );
    same.c.end();
    await empty();
    // Another file, and no name: it is founded again.
    const second = await joined(port, { v: ['g'] });
    assert.match(second.hello.ep, EPOCH);
    assert.notEqual(second.hello.ep, first.hello.ep, 'the room that is founded again is another room');
    assert.deepEqual({ from: second.hello.from, rev: second.hello.rev, v: second.hello.v }, { from: 0, rev: 0, v: ['g'] });
    const old = await knock(port, { v: ['f'], since: 0, ep: first.hello.ep, ro: true });
    assert.deepEqual(old.answer, { t: 'reset', rev: 0, v: ['g'] }, 'revision 0 of the room before it is not revision 0 of this one');
    // The same file, with nothing written since, finds the room it was: whoever was in it can come back to it.
    second.c.end();
    await empty();
    const third = await joined(port, { v: ['g'] });
    assert.equal(third.hello.ep, second.hello.ep, 'a room that is started again on the file it was founded on keeps its name');
    assert.notEqual(third.hello.ep, first.hello.ep, 'which is not the name of the room before it');

    // Once something was written, the room keeps its name with nobody in it, for as long as it is kept.
    await operations(third.c, 0, 3);
    third.c.end();
    await empty();
    const back = await joined(port, { v: ['saved-at-3'], since: 3, ep: third.hello.ep });
    assert.deepEqual({ ep: back.hello.ep, from: back.hello.from, rev: back.hello.rev }, { ep: third.hello.ep, from: 3, rev: 3 });
    back.c.end();
    await empty();
    const again = await joined(port, { v: ['g'] });
    assert.deepEqual({ ep: again.hello.ep, from: again.hello.from, rev: again.hello.rev }, { ep: third.hello.ep, from: 0, rev: 3 });
  });
  await withServer({ idleRoomMs: 60 }, async (port, server) => {
    const a = await joined(port, { v: ['v0'] });
    const old = a.hello.ep;
    await operations(a.c, 0, 3);
    a.c.end();
    await until('the room to be forgotten', () => server.rooms().length === 0);
    // The room is let go of. A client of it comes back with what it knew: revision 3 of a room that is no more.
    const back = await knock(port, { v: ['saved-at-3'], since: 3, ep: old, key: 'back' });
    assert.equal(back.answer.t, 'hello', `whoever comes first founds the room: ${JSON.stringify(back.answer)}`);
    assert.notEqual(back.answer.ep, old, 'the room is a new one');
    assert.match(back.answer.ep, EPOCH);
    assert.deepEqual(
      { from: back.answer.from, rev: back.answer.rev, v: back.answer.v },
      { from: 0, rev: 0, v: ['saved-at-3'] },
      'it is not caught up from revision 3 of the room that was let go of',
    );
    await operations(back.c, 0, 4);
    // Now the new room has a revision 3 of its own, and the old name still does not open it.
    for (const v of [['v0'], ['unknown'], []]) {
      const second = await knock(port, { v, since: 3, ep: old, ro: true });
      assert.deepEqual(second.answer, { t: 'reset', rev: 4, v: ['saved-at-3'] }, `v ${JSON.stringify(v)}: the old name opened the new room`);
    }
    const mine = await knock(port, { v: ['saved-at-3'], since: 3, ep: old, ro: true });
    assert.deepEqual(
      { t: mine.answer.t, from: mine.answer.from },
      { t: 'hello', from: 0 },
      'the file the new room was founded on is joined from 0, not from the 3 of the old room',
    );
    for (const rev of [1, 2, 3, 4]) assert.equal((await mine.c.json()).rev, rev);
  });
});

test('what a writer said of a version lets nobody in: a forged save is not joined from', async () => {
  await withServer({}, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const ep = a.hello.ep;
    await operations(a.c, 0, 3);
    const liar = await joined(port, { v: ['v0'] });
    for (const rev of [1, 2, 3]) assert.equal((await liar.c.json()).rev, rev);
    assert.equal((await a.c.json()).t, 'peer');

    // The liar may write and has seen revision 1: the table takes its word that X holds revision 1.
    liar.c.send({ t: 'saved', v: ['X'], rev: 1 });
    assert.deepEqual(await a.c.json(), { t: 'saved', from: liar.hello.you, v: ['X'], rev: 1 });
    // The one who really saved X says where: the table contradicts, and the revision is not passed on.
    a.c.send({ t: 'saved', v: ['X'], rev: 3 });
    assert.deepEqual(await liar.c.json(), { t: 'saved', from: a.hello.you, v: ['X'] }, 'the honest save is passed on without its revision');

    // Whoever loads X now has the file of revision 3. From 1 it would apply the operations 2 and 3 twice.
    const blind = await knock(port, { v: ['X'], key: 'blind' });
    assert.deepEqual(blind.answer, { t: 'reset', rev: 3, v: ['X'] }, 'a join with X alone is never caught up from where the liar put it');
    assert.deepEqual(await a.c.json(), { t: 'sync' });
    assert.deepEqual(await liar.c.json(), { t: 'sync' });

    // The file carries where it was saved; its client says so, and is believed for this room.
    const honest = await knock(port, { v: ['X'], since: 3, ep, key: 'honest' });
    assert.deepEqual(
      { t: honest.answer.t, ep: honest.answer.ep, from: honest.answer.from, rev: honest.answer.rev, v: honest.answer.v },
      { t: 'hello', ep, from: 3, rev: 3, v: ['X'] },
      'with the revision and the name of the room it gets in where its file is',
    );
    await honest.c.settled().catch(() => {
      throw new Error('an operation was played to a file that holds it already');
    });
    assert.equal(honest.c.seen.filter((text) => text.includes('"t":"op"')).length, 0, 'no operation was played to it');
    assert.equal((await a.c.json()).t, 'peer');
    // And it writes on from there.
    honest.c.send({ t: 'op', base: 3, id: 'honest:next', ops: [1] });
    assert.equal((await a.c.json()).rev, 4);
  });
});

test('a room has only so many saves a minute written down, whoever says them', async () => {
  let time = 7000;
  await withServer({ now: () => time, savedPerMinute: 0, roomSavedPerMinute: 3 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    const ep = a.hello.ep;
    assert.equal((await a.c.json()).t, 'peer');
    await operations(a.c, 0, 1);
    assert.equal((await b.c.json()).rev, 1);
    assert.equal((await says(a.c, b.c, ['s1'], 1)).rev, 1);
    assert.equal((await says(a.c, b.c, ['s2'], 1)).rev, 1);
    assert.equal((await says(b.c, a.c, ['t1'], 1)).rev, 1, 'the third save of the room, the first of its connection');
    // What only repeats the table costs nothing; what is refused for another reason neither.
    assert.equal((await says(a.c, b.c, ['s1', 't1'], 1)).rev, 1);
    assert.equal('rev' in (await says(a.c, b.c, ['s1'], 0)), false);
    const fourth = await says(a.c, b.c, ['s4'], 1);
    assert.equal('rev' in fourth, false, `the fourth save of the minute was written down: ${JSON.stringify(fourth)}`);
    assert.equal('rev' in (await says(b.c, a.c, ['t2'], 1)), false, 'the limit is of the room, whoever says it');
    // One who comes new has a full bucket of its own, and the room is as full as it was.
    const c = await joined(port, { v: ['v0'] });
    assert.equal((await next(c.c)).rev, 1);
    assert.equal('rev' in (await says(c.c, a.c, ['u1'], 1)), false, 'a new connection does not write for the room either');
    assert.deepEqual(await next(b.c), { t: 'saved', from: c.hello.you, v: ['u1'] }, 'everyone else hears it the same');
    assert.equal(a.c.closed, false, 'the connection is not closed');
    assert.deepEqual((await comesBack(port, ep, 1, ['s2', 's4', 't1', 't2', 'u1'])).v, ['s2', 't1'], 'what came over the limit is not in the table');
    time += 19000;
    assert.equal('rev' in (await says(a.c, b.c, ['s5'], 1)), false, 'nineteen seconds are not a save');
    time += 2000;
    assert.equal((await says(b.c, a.c, ['t3'], 1)).rev, 1, 'twenty-one seconds are');
    assert.equal('rev' in (await says(a.c, b.c, ['s6'], 1)), false, 'and only one');
    // What was not written for the rate was not spoiled: it can be said again.
    time += 20000;
    assert.equal((await says(a.c, b.c, ['s4'], 1)).rev, 1);
    assert.deepEqual((await comesBack(port, ep, 1, ['s4', 's5', 't3'])).v, ['s4', 't3']);
    // The bucket holds no more than a minute of saves.
    time += 600000;
    for (const version of ['w1', 'w2', 'w3']) assert.equal((await says(a.c, b.c, [version], 1)).rev, 1);
    assert.equal('rev' in (await says(a.c, b.c, ['w4'], 1)), false, 'ten minutes are three saves, not thirty');
  });
  // A save that the connection may not write takes nothing from the room, and the other way round.
  await withServer({ now: () => time, savedPerMinute: 1, roomSavedPerMinute: 2 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer');
    assert.equal((await says(a.c, b.c, ['s1'], 0)).rev, 0);
    assert.equal('rev' in (await says(a.c, b.c, ['s2'], 0)), false, 'the connection has one save a minute');
    assert.equal('rev' in (await says(a.c, b.c, ['s3'], 0)), false);
    assert.equal((await says(b.c, a.c, ['t1'], 0)).rev, 0, 'what the connection could not write was not taken from the room');
    const c = await joined(port, { v: ['v0'] });
    const d = await joined(port, { v: ['v0'] });
    assert.equal('rev' in (await says(c.c, a.c, ['u1'], 0)), false, 'the room is full');
    assert.equal('rev' in (await says(c.c, a.c, ['u2'], 0)), false);
    time += 30000;
    assert.equal((await says(c.c, a.c, ['u3'], 0)).rev, 0, 'what the room could not write was not taken from the connection');
    assert.equal('rev' in (await says(d.c, a.c, ['x1'], 0)), false, 'half a minute is one save of the room');
  });
  await withServer({ now: () => time, savedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer');
    for (let i = 1; i <= 120; i++) {
      assert.equal((await says(i % 2 ? a.c : b.c, i % 2 ? b.c : a.c, [`s${i}`], 0)).rev, 0, `save ${i} of the default hundred and twenty`);
    }
    assert.equal('rev' in (await says(a.c, b.c, ['s121'], 0)), false, 'the default is a hundred and twenty a minute for a room');
    // The limit is of the room: another room writes.
    const other = await joined(port, { v: ['v0'] }, '/other');
    const beside = await joined(port, { v: ['v0'] }, '/other');
    assert.equal((await says(other.c, beside.c, ['s121'], 0)).rev, 0, 'another room writes');
  });
  await withServer({ now: () => time, savedPerMinute: 0, roomSavedPerMinute: 0 }, async (port) => {
    const a = await joined(port, { v: ['v0'] });
    const b = await joined(port, { v: ['v0'] });
    assert.equal((await a.c.json()).t, 'peer');
    for (let i = 1; i <= 40; i++) assert.equal((await says(a.c, b.c, [`s${i}`], 0)).rev, 0, `save ${i}: 0 is no limit`);
  });
  assert.throws(() => createCollabServer({ roomSavedPerMinute: -1 }), TypeError, 'a rate that is negative is refused');
});
