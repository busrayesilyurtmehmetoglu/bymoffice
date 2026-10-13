#!/usr/bin/env node
// An example relay for collaborative editing. It needs nothing but Node: the
// WebSocket handshake and framing (RFC 6455) are written by hand below.
//
// NOTHING IS STORED. Rooms, their logs and the versions they know live in
// memory and are gone when the process ends or when a room has been empty for
// `idleRoomMs`. The relay never looks inside an operation: it gives accepted
// operations a number, in the order they arrive, and passes them on. Saving the
// file is the business of the clients and of the application's own storage.
//
// SECURITY. Whoever gets into a room reads the file that is edited there and
// changes it: the relay cannot tell an honest operation from another. Without
// `authorize`, everyone who can reach the relay gets into every room. In
// production `authorize` is NOT OPTIONAL (and `origins` should be set as well);
// a relay that is started without both says so once. The limits below (of
// connections, rooms, logs and rates) keep a client from using up the process;
// they do not decide who may come in.
//
// A room is the path of the request plus the `room` of the `join` message. The
// path is taken as it resolves (`/x/../y` is `/y`), without the query, with
// every segment percent-decoded: `/p/%61` and `/p/a` are one room. A path that
// cannot be decoded, or in which a decoded segment holds `/`, `\`, `#` or a
// control character, or is `.` or `..`, is answered with 400. `path` and
// `authorize` see the decoded path.
//
// A room has a revision `rev` that starts at 0, a log of the operations that
// made the revisions after `floor`, a version of the protocol `pv` (that of the
// first to join it) and a small table that says at which revision a version of
// the file (an ETag, a hash: any text) was saved. An operation is accepted only
// when its `base` is the room's `rev`; otherwise its sender gets a `nack` and
// tries again on top of what it missed.
//
// The table of versions. What it says of a version is written once and never
// changed: the versions the room was founded on are revision 0 for as long as
// it lives. It holds sixteen versions; the oldest save makes room, the versions
// of the foundation and those of the last save that was written never do. A
// version that left the table (it was full, or the log moved past its
// revision) is remembered, and is never written again at any revision: the room
// has not heard of it any more, whoever names it. A room remembers
// `maxEvictedVersions` of them; from then on it writes no new version at all. A
// `saved` whose revision the table contradicts, that names a version that left,
// that comes after the room stopped writing, or faster than `savedPerMinute`
// for its connection or than `roomSavedPerMinute` for the room (all its
// connections together), is passed on WITHOUT its `rev`, and writes nothing.
//
// What the table is for. It is what those who may write have SAID, so nobody
// gets into the room on its word (see the rules of `join`). It decides the
// `rev` of the `saved` messages that are passed on, it tells a client that
// comes back by rule a. which of its versions the room has heard of (the `v`
// of `hello`), and its newest versions are named in `reset`.
//
// A connection is in one room at most, and joins it once. A `join` may carry a
// `key`: a SECRET of the client, that the relay puts in no message, in no line
// of its log and not in `rooms()`. A join that gets in with the key of a
// connection that is in the room takes its place: the old connection is told
// `replaced` and closed. That is how a client that lost its connection gets
// back in before the relay has noticed. The `id` of an operation is any text
// the client likes; it is passed on and means nothing to the relay.
//
// The relay passes the operations on in the very text they came in. It looks at
// one thing: an operation that opens with `{ "t": "st", "f": [...] }` adds
// formats to the file, and a connection adds `formatsPerMinute` of them (the
// length of `f`); over it the operation is refused with `nack`, reason `rate`.
//
// A room has a name of its own, `ep`, made up when it is founded and told in
// every `hello`: a room that is founded again (the relay was restarted, the
// room was let go of after `idleRoomMs`, or everybody left it before anything
// was written and somebody came with no claim on it) is another room, whose
// revisions are counted anew. A room in which something was written keeps its
// name for as long as it is kept. A `join` may say the name back (`ep`, in the
// letters of a `key`: one to sixty-four of A-Z, a-z, 0-9, `_` and `-`).
//
// The four rules of `join`, tried in this order:
//   a. `since` is given with the `ep` of this room and floor <= since <= rev:
//      the connection is caught up from `since`. That is a client coming back,
//      or one whose file was saved in this room and says so itself. A `since`
//      without `ep`, or with another one, is not looked at: the join goes on
//      to the rules below as if it had named none. The `v` of its `hello` are
//      those of its versions the room has heard of (foundation or table).
//   b. the room is fresh (rev 0, nobody else in it): it is founded, with a new
//      `ep`; the versions `v` are written down as revision 0 and it starts there.
//   c. one of its versions `v` is one the room was founded on, and the log
//      still starts at 0: it is caught up from 0. (The table of versions that
//      `saved` fills is not used for this: what it says is what somebody said.
//      It tells who joins by rule a. whether the room has heard of its file.)
//   d. otherwise it gets `reset` with the newest versions the room knows. It
//      stays connected but is in no room until it has loaded that file and
//      joined the same room again, for which it has `joinTimeoutMs`. If it may
//      write (it did not ask for `ro`, and `authorize` did not say read-only),
//      those in the room who may write are sent `sync`: a save of theirs gives
//      the newcomer a version to load. A room is asked `syncPerMinute` times.
//
// Use as a library:
//   import { createCollabServer } from './collab-server.mjs';
//   const relay = createCollabServer({ origins: ['https://app.example'], authorize });
//   relay.attach(httpServer);            // or: await relay.listen(8787, '127.0.0.1')
// or as a program:
//   node collab-server.mjs --port 8787 --host 127.0.0.1 --origin https://app.example

import { createHash, randomBytes } from 'node:crypto';
import { realpathSync } from 'node:fs';
import http from 'node:http';
import { fileURLToPath } from 'node:url';

const HANDSHAKE_GUID = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';

const OP_CONTINUATION = 0;
const OP_TEXT = 1;
const OP_BINARY = 2;
const OP_CLOSE = 8;
const OP_PING = 9;
const OP_PONG = 10;

const MAX_PATH_LENGTH = 1024;
const MAX_ROOM_NAME = 200;
const MAX_VERSIONS_PER_MESSAGE = 8;
const MAX_VERSION_LENGTH = 200;
const MAX_VERSIONS_PER_ROOM = 16;
const MAX_USER_NAME = 64;
const MAX_USER_COLOR = 32;
const MAX_TOKEN_LENGTH = 4096;
const MAX_OP_ID = 100;
/** Ten thousand operations and the table of formats that may come with them. */
const MAX_OPS_PER_MESSAGE = 10001;
/**
 * How deep the lists and objects of `ops` may be nested, `ops` itself counted.
 * The relay passes the text on without building it again, so nothing else
 * would keep a text from the clients that their parsers cannot walk.
 */
const MAX_OPS_DEPTH = 256;
const MAX_PROTOCOL_VERSION = 1000;
const USER_KEY = /^[A-Za-z0-9_-]{1,64}$/;
/** A segment of a path that is `.` or `..` with a dot that is written `%2e`. */
const ENCODED_DOT_SEGMENT = /(?:^|[/\\])(?:%2e\.?|\.%2e|%2e%2e)(?=[/\\]|$)/i;
/** What a segment of a path may not hold once it is decoded: separators, the mark of a room, control characters. */
const FORBIDDEN_IN_SEGMENT = /[/\\#\u0000-\u001f\u007f-\u009f]/;
const MAX_PRESENCE_BYTES = 16 * 1024;
/** How long a connection that was sent a close frame may take to go away. */
const CLOSE_LINGER_MS = 2000;
/** A buffer of a message this small is kept for the next message. */
const KEPT_BUFFER_BYTES = 64 * 1024;

const DEFAULTS = {
  maxMessageBytes: 1024 * 1024,
  maxConnectionsPerRoom: 64,
  maxRooms: 1000,
  maxLogEntries: 20000,
  maxLogBytes: 32 * 1024 * 1024,
  idleRoomMs: 5 * 60 * 1000,
  joinTimeoutMs: 10 * 1000,
  pingIntervalMs: 30 * 1000,
  maxBufferedBytes: 8 * 1024 * 1024,
  maxConnections: 10000,
  /** 0: as many as it likes. */
  maxConnectionsPerIp: 0,
  maxTotalLogBytes: 512 * 1024 * 1024,
  /** 0: no limit. */
  opsPerSecond: 50,
  opsBurst: 200,
  /** 0: no limit. */
  presencePerSecond: 20,
  /** 0: no limit. */
  syncPerMinute: 2,
  /** 0: no limit. */
  savedPerMinute: 6,
  /** 0: no limit. */
  roomSavedPerMinute: 30,
  /** 0: no limit. */
  formatsPerMinute: 2000,
  /** 0: no save is ever written down. */
  maxEvictedVersions: 4096,
};

const WARNING =
  'collab server: anyone who can reach this relay can read and change every file edited through it; ' +
  'set `authorize` (and `origins`) in production';

const EMPTY = Buffer.alloc(0);
const utf8 = new TextDecoder('utf-8', { fatal: true, ignoreBOM: true });

function noop() {}

/** Milliseconds that only go forward. */
function monotonic() {
  return performance.now();
}

/**
 * The path of a request with every segment decoded, so that a room has one
 * name however its path is spelled; null when a segment cannot be decoded or
 * would not be one segment any more (see FORBIDDEN_IN_SEGMENT), or is a dot
 * segment after all.
 */
function decodePath(pathname) {
  const segments = pathname.split('/');
  for (let index = 0; index < segments.length; index++) {
    let segment;
    try {
      segment = decodeURIComponent(segments[index]);
    } catch {
      return null;
    }
    if (segment === '.' || segment === '..' || FORBIDDEN_IN_SEGMENT.test(segment)) return null;
    segments[index] = segment;
  }
  return segments.join('/');
}

function isObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isCount(value) {
  return Number.isSafeInteger(value) && value >= 0;
}

/** Whether `code` is white space between the tokens of JSON. */
function isJsonSpace(code) {
  return code === 0x20 || code === 0x0a || code === 0x0d || code === 0x09;
}

/**
 * The text of the value that the name `ops` has in `text`, an object that
 * JSON.parse has read without complaint: the same value JSON.parse found (the
 * last one, if the name is there more than once), as its sender wrote it, and
 * how deep it is nested. Null when it is not found, which is not to happen.
 *
 * Only the outermost object is walked, with a counter and no recursion, and
 * nothing is decided here that JSON.parse has not decided already.
 */
function rawOps(text) {
  const length = text.length;
  let at = 0;
  /** The index after the string that opens at `from`. */
  const afterString = (from) => {
    let index = from + 1;
    while (index < length) {
      const code = text.charCodeAt(index);
      if (code === 0x5c) index += 2;
      else if (code === 0x22) return index + 1;
      else index++;
    }
    return -1;
  };
  const skipSpace = () => {
    while (at < length && isJsonSpace(text.charCodeAt(at))) at++;
  };
  skipSpace();
  if (text.charCodeAt(at) !== 0x7b) return null;
  at++;
  let found = null;
  for (;;) {
    skipSpace();
    const first = text.charCodeAt(at);
    if (first === 0x7d) return found;
    if (first === 0x2c) {
      at++;
      continue;
    }
    if (first !== 0x22) return null;
    const nameEnd = afterString(at);
    if (nameEnd < 0) return null;
    const name = text.slice(at, nameEnd);
    at = nameEnd;
    skipSpace();
    if (text.charCodeAt(at) !== 0x3a) return null;
    at++;
    skipSpace();
    const start = at;
    const opens = text.charCodeAt(at);
    let deepest = 0;
    if (opens === 0x22) {
      at = afterString(at);
      if (at < 0) return null;
    } else if (opens === 0x5b || opens === 0x7b) {
      let depth = 0;
      while (at < length) {
        const code = text.charCodeAt(at);
        if (code === 0x22) {
          at = afterString(at);
          if (at < 0) return null;
          continue;
        }
        if (code === 0x5b || code === 0x7b) {
          depth++;
          if (depth > deepest) deepest = depth;
        } else if (code === 0x5d || code === 0x7d) depth--;
        at++;
        if (depth === 0) break;
      }
      if (depth !== 0) return null;
    } else {
      // A number, true, false or null: it ends where the next token begins.
      while (at < length) {
        const code = text.charCodeAt(at);
        if (code === 0x2c || code === 0x7d || isJsonSpace(code)) break;
        at++;
      }
    }
    // A name may be written with escapes: `"\u006fps"` is `ops` as well.
    if (name === '"ops"' || (name.includes('\\') && JSON.parse(name) === 'ops')) {
      found = { text: text.slice(start, at), depth: deepest };
    }
  }
}

/** A server frame: never masked, never fragmented. */
function encodeFrame(opcode, payload) {
  const length = payload.length;
  let head;
  if (length < 126) {
    head = Buffer.allocUnsafe(2);
    head[1] = length;
  } else if (length < 65536) {
    head = Buffer.allocUnsafe(4);
    head[1] = 126;
    head.writeUInt16BE(length, 2);
  } else {
    head = Buffer.allocUnsafe(10);
    head[1] = 127;
    head.writeUInt32BE(Math.floor(length / 4294967296), 2);
    head.writeUInt32BE(length >>> 0, 6);
  }
  head[0] = 0x80 | opcode;
  return Buffer.concat([head, payload]);
}

function encodeMessage(message) {
  return encodeFrame(OP_TEXT, Buffer.from(JSON.stringify(message), 'utf8'));
}

/** `text` cut to `max` UTF-16 units, without leaving half of a pair behind. */
function clip(text, max) {
  if (text.length <= max) return text;
  const last = text.charCodeAt(max - 1);
  return text.slice(0, last >= 0xd800 && last <= 0xdbff ? max - 1 : max);
}

/** The version strings of a message without repeats, or null when they are not what the contract says. */
function versionList(value, least) {
  if (!Array.isArray(value) || value.length < least || value.length > MAX_VERSIONS_PER_MESSAGE) return null;
  const list = [];
  for (const item of value) {
    if (typeof item !== 'string' || item.length < 1 || item.length > MAX_VERSION_LENGTH) return null;
    if (!list.includes(item)) list.push(item);
  }
  return list;
}

function readOptions(options) {
  const out = { ...DEFAULTS };
  for (const name of Object.keys(DEFAULTS)) {
    const value = options[name];
    if (value === undefined) continue;
    if (typeof value !== 'number' || !(value >= 0)) throw new TypeError(`collab server: ${name} must be a number that is not negative`);
    out[name] = value;
  }
  out.path = typeof options.path === 'string' ? options.path : '';
  out.origins = Array.isArray(options.origins) ? options.origins.map(String) : null;
  out.authorize = typeof options.authorize === 'function' ? options.authorize : null;
  out.log = typeof options.log === 'function' ? options.log : null;
  out.now = typeof options.now === 'function' ? options.now : monotonic;
  return out;
}

/**
 * Makes a relay. It listens to nothing until `attach()` or `listen()` is called.
 */
export function createCollabServer(options = {}) {
  const opt = readOptions(options);
  /** Room key → room. */
  const rooms = new Map();
  const connections = new Set();
  /** Remote address → how many connections it has. */
  const perAddress = new Map();
  /** The bytes of the logs of all rooms. */
  let totalLogBytes = 0;
  /** Counts the operations of all rooms, to tell which log entry is the oldest. */
  let logSequence = 0;
  let warned = false;
  const opsBurst = Math.max(1, opt.opsBurst);
  const presenceBurst = Math.max(1, Math.ceil(opt.presencePerSecond));
  const syncBurst = Math.max(1, Math.ceil(opt.syncPerMinute));
  const savedBurst = Math.max(1, Math.ceil(opt.savedPerMinute));
  const roomSavedBurst = Math.max(1, Math.ceil(opt.roomSavedPerMinute));
  const formatsBurst = Math.max(1, Math.ceil(opt.formatsPerMinute));
  const now = opt.now;

  /**
   * A bucket of `burst` tokens that fills at `rate` a second; false when it
   * holds fewer than `count`, and then nothing is taken.
   */
  function takeToken(bucket, rate, burst, count = 1) {
    const time = now();
    // A clock that is handed in may stand still or step back; the bucket does not empty for it.
    const passed = time > bucket.at ? time - bucket.at : 0;
    bucket.tokens = Math.min(burst, bucket.tokens + (passed * rate) / 1000);
    bucket.at = time;
    if (bucket.tokens < count) return false;
    bucket.tokens -= count;
    return true;
  }
  /** The servers whose upgrades are served → the listener that was added. */
  const attached = new Map();
  let ownServer = null;
  let pingTimer = null;
  let closed = false;

  function log(event) {
    if (!opt.log) return;
    try {
      opt.log(event);
    } catch {
      // A log that fails must not take a connection with it.
    }
  }

  /** Says once, when the relay starts to serve, that it lets everyone in. */
  function warnIfOpen() {
    if (warned || opt.origins || opt.authorize) return;
    warned = true;
    if (opt.log) log({ type: 'warning', message: WARNING });
    else console.warn(WARNING);
  }

  // --- Sockets ---------------------------------------------------------------

  function refuse(socket, status, text, extra = '') {
    try {
      socket.write(`HTTP/1.1 ${status} ${text}\r\nConnection: close\r\nContent-Length: 0\r\n${extra}\r\n`);
      socket.destroySoon();
    } catch {
      socket.destroy();
    }
  }

  function onUpgrade(server, request, socket, head) {
    socket.on('error', noop);
    const target = typeof request.url === 'string' ? request.url : '';
    // The path names the room, so it is resolved as a browser resolves it
    // (`/x/../y` is `/y`). What could name one room in two ways is refused: a
    // fragment, and dots that are written `%2e`.
    let url;
    try {
      // Not the absolute form of a request target: the path is all of it.
      url = new URL(target.startsWith('/') ? `http://relay${target}` : target);
    } catch {
      return refuse(socket, 400, 'Bad Request');
    }
    const mark = target.indexOf('?');
    if (target.includes('#') || ENCODED_DOT_SEGMENT.test(mark < 0 ? target : target.slice(0, mark))) {
      return refuse(socket, 400, 'Bad Request');
    }
    // One room, one name: `/p/%61` is `/p/a`. The prefix is looked for in the
    // decoded path, and so is everything that is said of the path from here on.
    const pathname = decodePath(url.pathname);
    if (pathname === null) return refuse(socket, 400, 'Bad Request');
    if (opt.path && !pathname.startsWith(opt.path)) {
      // Somebody else may be serving this path on the same server.
      if (server.listenerCount('upgrade') > 1) {
        socket.removeListener('error', noop);
        return;
      }
      return refuse(socket, 404, 'Not Found');
    }
    if (closed) return refuse(socket, 503, 'Service Unavailable');
    const headers = request.headers;
    const key = headers['sec-websocket-key'];
    const wanted =
      request.method === 'GET' &&
      typeof headers.upgrade === 'string' &&
      headers.upgrade.trim().toLowerCase() === 'websocket' &&
      typeof headers.connection === 'string' &&
      headers.connection.toLowerCase().split(',').some((token) => token.trim() === 'upgrade') &&
      headers['sec-websocket-version'] === '13' &&
      typeof key === 'string' &&
      /^[A-Za-z0-9+/]{22}==$/.test(key) &&
      url.pathname.length <= MAX_PATH_LENGTH;
    if (!wanted) return refuse(socket, 400, 'Bad Request', 'Sec-WebSocket-Version: 13\r\n');
    const origin = typeof headers.origin === 'string' ? headers.origin : undefined;
    if (opt.origins && (origin === undefined || !opt.origins.includes(origin))) {
      log({ type: 'refused', reason: 'origin', path: pathname, origin });
      return refuse(socket, 403, 'Forbidden');
    }
    let queryToken;
    const found = url.searchParams.get('token');
    if (found !== null) queryToken = found;
    // No subprotocol and no extension is ever agreed to: neither is named here.
    const accept = createHash('sha1').update(key + HANDSHAKE_GUID).digest('base64');
    socket.write(
      'HTTP/1.1 101 Switching Protocols\r\nUpgrade: websocket\r\nConnection: Upgrade\r\n' +
        `Sec-WebSocket-Accept: ${accept}\r\n\r\n`,
    );
    socket.setNoDelay(true);
    socket.setTimeout(0);

    const conn = {
      id: randomBytes(6).toString('base64url'),
      socket,
      request,
      path: pathname,
      origin,
      queryToken,
      address: typeof socket.remoteAddress === 'string' ? socket.remoteAddress : '',
      /** The room it is in, once a `join` was answered with `hello`. */
      room: null,
      /** The key of the room a `join` of it was answered with `reset`: the only room it may still join. */
      resetFrom: null,
      /** The `key` of its join, while it is in a room: a secret, said to nobody. */
      key: undefined,
      /** The revision it joined from: what came before, it has not seen. */
      from: 0,
      opTokens: { tokens: opsBurst, at: now() },
      presenceTokens: { tokens: presenceBurst, at: now() },
      savedTokens: { tokens: savedBurst, at: now() },
      formatTokens: { tokens: formatsBurst, at: now() },
      /** A `join` was received: the first message has been what it must be. */
      greeted: false,
      ro: false,
      user: {},
      presence: undefined,
      hasPresence: false,
      /** Pings that were sent since the last pong. */
      unanswered: 0,
      /** A close frame was sent; what still arrives is thrown away. */
      closing: false,
      dead: false,
      /** Waiting for `authorize`: what arrives is held, in order. */
      blocked: false,
      held: [],
      header: Buffer.alloc(14),
      headerLength: 0,
      headerNeed: 2,
      frame: null,
      /** The text message that is being put together, over frames and chunks. */
      message: null,
      messageLength: 0,
      inMessage: false,
      /** Bytes of a replayed log that may wait in the socket beyond the limit. */
      allowance: 0,
      joinTimer: null,
      closeTimer: null,
    };
    connections.add(conn);
    const ofAddress = (perAddress.get(conn.address) || 0) + 1;
    perAddress.set(conn.address, ofAddress);
    // Asked before this one was counted: the handshake is answered all the
    // same, because a browser can read a message and cannot read a status.
    const full =
      connections.size > opt.maxConnections || (opt.maxConnectionsPerIp > 0 && ofAddress > opt.maxConnectionsPerIp);
    if (!full) startJoinTimer(conn);
    startPinging();
    log({ type: 'connect', id: conn.id, path: pathname, origin });

    const guarded = (work) => {
      try {
        work();
      } catch (error) {
        log({ type: 'error', id: conn.id, message: describe(error, conn) });
        drop(conn, 'error');
      }
    };
    socket.removeListener('error', noop);
    socket.on('error', () => drop(conn, 'socket-error'));
    socket.on('close', () => cleanup(conn, 'closed'));
    socket.on('end', () => {
      // The other side will say no more; neither will we.
      if (conn.closing) socket.destroy();
      else drop(conn, 'ended');
    });
    socket.on('drain', () => {
      conn.allowance = 0;
    });
    socket.on('data', (chunk) =>
      guarded(() => {
        if (conn.closing || conn.dead) return;
        if (conn.blocked || conn.held.length) conn.held.push(chunk);
        else feed(conn, chunk);
      }),
    );
    if (full) return fail(conn, 'server-full', 1013);
    if (head && head.length) guarded(() => feed(conn, head));
  }

  /** A connection that is in no room has this long to get into one. */
  function startJoinTimer(conn) {
    clearTimeout(conn.joinTimer);
    conn.joinTimer = null;
    if (opt.joinTimeoutMs > 0) {
      conn.joinTimer = setTimeout(() => fail(conn, 'join-timeout', 1008), opt.joinTimeoutMs);
    }
  }

  /** What an error says, without the token and the key of the connection if they found their way in. */
  function describe(error, conn, token, key) {
    let text = String((error && error.message) || error);
    for (const secret of [token, conn.queryToken]) {
      if (typeof secret === 'string' && secret) text = text.split(secret).join('[token]');
    }
    for (const secret of [key, conn.key]) {
      if (typeof secret === 'string' && secret) text = text.split(secret).join('[key]');
    }
    return text.slice(0, 300);
  }

  /** Writes a frame; a connection whose queue is too long is cut off. */
  function write(conn, frame) {
    if (conn.dead || conn.closing) return false;
    try {
      conn.socket.write(frame);
    } catch {
      drop(conn, 'socket-error');
      return false;
    }
    if (conn.socket.writableLength > opt.maxBufferedBytes + conn.allowance) {
      log({ type: 'slow', id: conn.id, buffered: conn.socket.writableLength });
      drop(conn, 'slow');
      return false;
    }
    return true;
  }

  function send(conn, message) {
    return write(conn, encodeMessage(message));
  }

  /** Sends a close frame and lets the socket finish; the room is left at once. */
  function closeWith(conn, code, reason = '') {
    if (conn.dead || conn.closing) return;
    let payload = EMPTY;
    if (code !== null) {
      const text = Buffer.from(reason, 'utf8').subarray(0, 123);
      payload = Buffer.allocUnsafe(2 + text.length);
      payload.writeUInt16BE(code, 0);
      text.copy(payload, 2);
    }
    write(conn, encodeFrame(OP_CLOSE, payload));
    if (conn.dead) return;
    conn.closing = true;
    clearTimeout(conn.joinTimer);
    conn.held.length = 0;
    conn.message = null;
    leave(conn);
    try {
      // The rest of what the client sends is read and thrown away, so that the
      // close frame is not lost to a reset of the connection.
      conn.socket.resume();
      conn.socket.end();
    } catch {
      // The socket is gone already.
    }
    conn.closeTimer = setTimeout(() => drop(conn, 'close-timeout'), CLOSE_LINGER_MS);
    conn.closeTimer.unref();
  }

  /** An `error` message and then the close frame. */
  function fail(conn, code, status) {
    if (conn.dead || conn.closing) return;
    log({ type: 'rejected', id: conn.id, code });
    send(conn, { t: 'error', code });
    closeWith(conn, status, code);
  }

  function drop(conn, why) {
    if (conn.dead) return;
    conn.socket.destroy();
    cleanup(conn, why);
  }

  function cleanup(conn, why) {
    if (conn.dead) return;
    conn.dead = true;
    clearTimeout(conn.joinTimer);
    clearTimeout(conn.closeTimer);
    conn.held.length = 0;
    conn.message = null;
    conn.frame = null;
    connections.delete(conn);
    const ofAddress = perAddress.get(conn.address) - 1;
    if (ofAddress > 0) perAddress.set(conn.address, ofAddress);
    else perAddress.delete(conn.address);
    leave(conn);
    log({ type: 'disconnect', id: conn.id, reason: why });
  }

  function startPinging() {
    if (pingTimer || !(opt.pingIntervalMs > 0)) return;
    const ping = encodeFrame(OP_PING, EMPTY);
    pingTimer = setInterval(() => {
      for (const conn of [...connections]) {
        if (conn.closing) continue;
        if (conn.unanswered >= 2) {
          drop(conn, 'no-pong');
          continue;
        }
        conn.unanswered++;
        write(conn, ping);
      }
      if (!connections.size) {
        clearInterval(pingTimer);
        pingTimer = null;
      }
    }, opt.pingIntervalMs);
    pingTimer.unref();
  }

  // --- Frames ----------------------------------------------------------------

  /** Reads what arrived, as far as it goes: a frame may come in any number of pieces. */
  function feed(conn, chunk) {
    let at = 0;
    while (at < chunk.length) {
      if (conn.closing || conn.dead) return;
      if (conn.blocked) {
        conn.held.unshift(chunk.subarray(at));
        return;
      }
      if (!conn.frame) {
        const header = conn.header;
        const take = Math.min(conn.headerNeed - conn.headerLength, chunk.length - at);
        chunk.copy(header, conn.headerLength, at, at + take);
        conn.headerLength += take;
        at += take;
        if (conn.headerLength < conn.headerNeed) return;
        const opcode = header[0] & 0x0f;
        const fin = (header[0] & 0x80) !== 0;
        let length = header[1] & 0x7f;
        if (conn.headerNeed === 2) {
          // Everything that can be refused is refused before a byte more is read.
          if (header[0] & 0x70) return closeWith(conn, 1002, 'reserved bits');
          if (!(header[1] & 0x80)) return closeWith(conn, 1002, 'frame is not masked');
          if (opcode >= 8) {
            if (opcode > OP_PONG) return closeWith(conn, 1002, 'unknown opcode');
            if (!fin || length > 125) return closeWith(conn, 1002, 'bad control frame');
          } else if (opcode === OP_BINARY) {
            return closeWith(conn, 1003, 'text only');
          } else if (opcode > OP_BINARY) {
            return closeWith(conn, 1002, 'unknown opcode');
          } else if ((opcode === OP_CONTINUATION) !== conn.inMessage) {
            return closeWith(conn, 1002, 'fragments out of order');
          }
          conn.headerNeed = 6 + (length === 126 ? 2 : length === 127 ? 8 : 0);
          continue;
        }
        let maskAt = 2;
        if (length === 126) {
          length = header.readUInt16BE(2);
          maskAt = 4;
          if (length < 126) return closeWith(conn, 1002, 'length is not minimal');
        } else if (length === 127) {
          const high = header.readUInt32BE(2);
          length = header.readUInt32BE(6);
          maskAt = 10;
          if (high !== 0) return closeWith(conn, 1009, 'message too big');
          if (length < 65536) return closeWith(conn, 1002, 'length is not minimal');
        }
        if (opcode < 8 && conn.messageLength + length > opt.maxMessageBytes) {
          return closeWith(conn, 1009, 'message too big');
        }
        conn.frame = {
          opcode,
          fin,
          left: length,
          done: 0,
          mask: [header[maskAt], header[maskAt + 1], header[maskAt + 2], header[maskAt + 3]],
          control: opcode >= 8 ? Buffer.allocUnsafe(length) : null,
        };
        if (opcode === OP_TEXT) conn.inMessage = true;
        conn.headerLength = 0;
        conn.headerNeed = 2;
        if (length === 0) endFrame(conn);
        continue;
      }
      const frame = conn.frame;
      const take = Math.min(frame.left, chunk.length - at);
      let target;
      let offset;
      if (frame.control) {
        target = frame.control;
        offset = frame.done;
      } else {
        target = roomFor(conn, take);
        offset = conn.messageLength;
        conn.messageLength += take;
      }
      const mask = frame.mask;
      const phase = frame.done;
      for (let i = 0; i < take; i++) target[offset + i] = chunk[at + i] ^ mask[(phase + i) & 3];
      frame.done += take;
      frame.left -= take;
      at += take;
      if (frame.left === 0) endFrame(conn);
    }
  }

  /**
   * The buffer of the message with room for `more` bytes. It grows with what
   * has arrived, never with what a header promises.
   */
  function roomFor(conn, more) {
    const need = conn.messageLength + more;
    const have = conn.message ? conn.message.length : 0;
    if (have >= need) return conn.message;
    const size = Math.max(need, Math.min(Math.max(have * 2, 1024), opt.maxMessageBytes));
    const next = Buffer.allocUnsafe(size);
    if (conn.message) conn.message.copy(next, 0, 0, conn.messageLength);
    conn.message = next;
    return next;
  }

  function endFrame(conn) {
    const frame = conn.frame;
    conn.frame = null;
    if (frame.opcode === OP_CLOSE) return onCloseFrame(conn, frame.control);
    if (frame.opcode === OP_PING) return void write(conn, encodeFrame(OP_PONG, frame.control));
    if (frame.opcode === OP_PONG) {
      conn.unanswered = 0;
      return;
    }
    if (!frame.fin) return;
    conn.inMessage = false;
    const bytes = conn.message ? conn.message.subarray(0, conn.messageLength) : EMPTY;
    conn.messageLength = 0;
    let text;
    try {
      text = utf8.decode(bytes);
    } catch {
      return closeWith(conn, 1007, 'text is not UTF-8');
    }
    if (conn.message && conn.message.length > KEPT_BUFFER_BYTES) conn.message = null;
    onText(conn, text);
  }

  function onCloseFrame(conn, payload) {
    if (payload.length === 0) return closeWith(conn, null);
    if (payload.length === 1) return closeWith(conn, 1002, 'bad close frame');
    const code = payload.readUInt16BE(0);
    const known = (code >= 1000 && code <= 1003) || (code >= 1007 && code <= 1014) || (code >= 3000 && code <= 4999);
    if (!known) return closeWith(conn, 1002, 'bad close code');
    try {
      utf8.decode(payload.subarray(2));
    } catch {
      return closeWith(conn, 1007, 'reason is not UTF-8');
    }
    closeWith(conn, code);
  }

  // --- Rooms -----------------------------------------------------------------

  function broadcast(room, message, except) {
    if (!room.conns.size) return;
    const frame = encodeMessage(message);
    for (const conn of [...room.conns]) {
      if (conn !== except) write(conn, frame);
    }
  }

  function leave(conn) {
    const room = conn.room;
    if (!room) return;
    conn.room = null;
    conn.key = undefined;
    conn.presence = undefined;
    conn.hasPresence = false;
    room.conns.delete(conn);
    log({ type: 'leave', id: conn.id, room: room.key });
    broadcast(room, { t: 'left', id: conn.id });
    if (!room.conns.size) startIdleTimer(room);
  }

  function startIdleTimer(room) {
    clearTimeout(room.idleTimer);
    room.idleTimer = setTimeout(() => {
      room.idleTimer = null;
      if (room.conns.size || rooms.get(room.key) !== room) return;
      rooms.delete(room.key);
      totalLogBytes -= room.logBytes;
      log({ type: 'room-removed', room: room.key, rev: room.rev });
    }, opt.idleRoomMs);
    room.idleTimer.unref();
  }

  /**
   * Remembers that a version left the table: it is never written again. Once
   * the room remembers as many as it may, it writes no version at all any more
   * (see `sealed`), and so has no need to remember what leaves after that.
   */
  function forget(room, version) {
    if (room.evicted.size < opt.maxEvictedVersions) room.evicted.add(version);
  }

  /** The room can no longer tell a version that left its table from one it never had: nothing new is written. */
  function sealed(room) {
    return room.evicted.size >= opt.maxEvictedVersions;
  }

  /**
   * Writes down where a version was saved. The oldest save makes room in a
   * full table; the versions the room was founded on and those of the save
   * that is being written (`room.lastSaved`) never do.
   */
  function setVersion(room, version, rev) {
    if (room.versions.has(version) || room.founding.has(version)) return;
    room.versions.set(version, rev);
    if (room.versions.size <= MAX_VERSIONS_PER_ROOM) return;
    for (const old of room.versions.keys()) {
      if (room.founding.has(old) || room.lastSaved.has(old)) continue;
      room.versions.delete(old);
      forget(room, old);
      return;
    }
  }

  /** The revision the table gives a version: 0 for ever for those the room was founded on. */
  function revisionOf(room, version) {
    const rev = room.versions.get(version);
    if (rev !== undefined) return rev;
    return room.founding.has(version) ? 0 : undefined;
  }

  /** The strings of the newest version the room knows. */
  function latestVersions(room) {
    let newest = -1;
    for (const rev of room.versions.values()) newest = Math.max(newest, rev);
    const list = [];
    for (const [version, rev] of room.versions) {
      if (rev === newest) list.push(version);
    }
    return list;
  }

  /** Takes the `count` oldest entries off the log of a room: the floor rises. */
  function cutLog(room, count) {
    if (!count) return;
    for (let index = 0; index < count; index++) {
      room.logBytes -= room.log[index].bytes;
      totalLogBytes -= room.log[index].bytes;
    }
    room.floor = room.log[count - 1].rev;
    room.log.splice(0, count);
    // A version that was founded on stays in `founding`: it cannot be joined
    // from any more, and it cannot be given another revision either. Neither
    // can a version of a save, which is remembered as one that left.
    for (const [version, rev] of room.versions) {
      if (rev >= room.floor) continue;
      room.versions.delete(version);
      if (!room.founding.has(version)) forget(room, version);
    }
  }

  function trimLog(room) {
    let cut = 0;
    let bytes = room.logBytes;
    while (cut < room.log.length && (room.log.length - cut > opt.maxLogEntries || bytes > opt.maxLogBytes)) {
      bytes -= room.log[cut].bytes;
      cut++;
    }
    cutLog(room, cut);
    if (totalLogBytes > opt.maxTotalLogBytes) trimAllLogs();
  }

  /** Cuts the oldest operations of the whole relay until the logs fit together. */
  function trimAllLogs() {
    while (totalLogBytes > opt.maxTotalLogBytes) {
      // The room whose oldest operation is the oldest of all, and how old the
      // oldest of the other rooms is: up to there this room is cut in one go.
      let oldest = null;
      let next = Infinity;
      for (const room of rooms.values()) {
        if (!room.log.length) continue;
        const sequence = room.log[0].sequence;
        if (!oldest || sequence < oldest.log[0].sequence) {
          if (oldest) next = oldest.log[0].sequence;
          oldest = room;
        } else if (sequence < next) {
          next = sequence;
        }
      }
      if (!oldest) return;
      let cut = 0;
      let bytes = totalLogBytes;
      while (cut < oldest.log.length && bytes > opt.maxTotalLogBytes && oldest.log[cut].sequence < next) {
        bytes -= oldest.log[cut].bytes;
        cut++;
      }
      cutLog(oldest, cut);
    }
  }

  // --- Messages --------------------------------------------------------------

  function onText(conn, text) {
    let message;
    try {
      message = JSON.parse(text);
    } catch {
      return fail(conn, 'bad-message', 1008);
    }
    if (!isObject(message)) return fail(conn, 'bad-message', 1008);
    if (!conn.greeted && message.t !== 'join') return fail(conn, 'bad-message', 1008);
    let fine;
    try {
      if (message.t === 'join') fine = onJoin(conn, message);
      else if (message.t === 'op') fine = onOp(conn, message, text);
      else if (message.t === 'presence') fine = onPresence(conn, message);
      else if (message.t === 'saved') fine = onSaved(conn, message);
      else if (message.t === 'ping') fine = send(conn, { t: 'pong' }) || true;
      else fine = false;
    } catch (error) {
      // What cannot be written back (too deep for the stack, say) is refused.
      log({ type: 'error', id: conn.id, message: describe(error, conn, message.token, message.key) });
      fine = false;
    }
    if (fine === false) fail(conn, 'bad-message', 1008);
  }

  function onJoin(conn, message) {
    const versions = versionList(message.v, 0);
    if (!versions) return false;
    const since = message.since;
    if (since !== undefined && !isCount(since)) return false;
    const ep = message.ep;
    if (ep !== undefined && (typeof ep !== 'string' || !USER_KEY.test(ep))) return false;
    if (message.ro !== undefined && typeof message.ro !== 'boolean') return false;
    const name = message.room;
    // A `#` in the name would let two pairs of path and name make one key.
    if (name !== undefined) {
      if (typeof name !== 'string' || name.length > MAX_ROOM_NAME || name.includes('#')) return false;
    }
    const pv = message.pv === undefined ? 1 : message.pv;
    if (!Number.isInteger(pv) || pv < 1 || pv > MAX_PROTOCOL_VERSION) return false;
    const userKey = message.key;
    if (userKey !== undefined && (typeof userKey !== 'string' || !USER_KEY.test(userKey))) return false;
    const user = {};
    if (message.user !== undefined && message.user !== null) {
      if (!isObject(message.user)) return false;
      if (typeof message.user.name === 'string') user.name = clip(message.user.name, MAX_USER_NAME);
      if (typeof message.user.color === 'string') user.color = clip(message.user.color, MAX_USER_COLOR);
    }
    const token = message.token !== undefined ? message.token : conn.queryToken;
    if (token !== undefined && (typeof token !== 'string' || token.length > MAX_TOKEN_LENGTH)) return false;

    const key = name ? `${conn.path}#${name}` : conn.path;
    // One room for a connection: it joins once, and after a `reset` it may try
    // the same room again and no other.
    if (conn.room || (conn.resetFrom !== null && conn.resetFrom !== key)) return false;
    conn.greeted = true;
    clearTimeout(conn.joinTimer);
    conn.joinTimer = null;
    const join = { key, versions, since, ep, user, ro: message.ro === true, pv, userKey };
    if (!opt.authorize) return enter(conn, join, true);

    let verdict;
    try {
      verdict = opt.authorize({ room: key, token, origin: conn.origin, request: conn.request });
    } catch {
      // What it threw is not logged: it may quote the token.
      log({ type: 'authorize-failed', id: conn.id, room: key });
      verdict = false;
    }
    if (!verdict || typeof verdict.then !== 'function') return enter(conn, join, verdict);
    // Nothing more is read from this connection until the answer is there.
    conn.blocked = true;
    conn.socket.pause();
    Promise.resolve(verdict)
      .then(
        (value) => value,
        () => {
          log({ type: 'authorize-failed', id: conn.id, room: key });
          return false;
        },
      )
      .then((value) => {
        if (conn.dead || conn.closing) return;
        conn.blocked = false;
        try {
          enter(conn, join, value);
          conn.socket.resume();
          while (conn.held.length && !conn.blocked && !conn.dead && !conn.closing) feed(conn, conn.held.shift());
        } catch (error) {
          log({ type: 'error', id: conn.id, message: describe(error, conn, token, userKey) });
          drop(conn, 'error');
        }
      });
    return true;
  }

  /** Applies the four rules of a join, once `authorize` has given its verdict. */
  function enter(conn, join, verdict) {
    if (verdict !== true && !isObject(verdict)) {
      fail(conn, 'unauthorized', 1008);
      return true;
    }
    const known = rooms.get(join.key);
    // The connection of the room that joined with the same key: whoever knows
    // the key is that client, come back, and takes its place if it gets in.
    let holder = null;
    if (known) {
      if (known.pv !== join.pv) {
        fail(conn, 'version', 1008);
        return true;
      }
      if (join.userKey !== undefined) {
        for (const other of known.conns) {
          if (other.key === join.userKey) holder = other;
        }
      }
      if (known.conns.size - (holder ? 1 : 0) >= opt.maxConnectionsPerRoom) {
        fail(conn, 'room-full', 1013);
        return true;
      }
    } else if (rooms.size >= opt.maxRooms) {
      fail(conn, 'server-full', 1013);
      return true;
    }
    // A room that does not exist is put in the table only when the join gets
    // in: a join that is refused leaves nothing behind.
    const room = known || {
      key: join.key,
      /** The name of the room as it was founded: a new one every time, so that "revision 3" of the room before it means nothing here. */
      ep: randomBytes(9).toString('base64url'),
      pv: join.pv,
      rev: 0,
      floor: 0,
      log: [],
      logBytes: 0,
      versions: new Map(),
      /** The versions of the first to come: revision 0, whatever anyone says later. */
      founding: new Set(),
      /** Versions that left the table (it was full, or the log moved past them): never written again. */
      evicted: new Set(),
      /** The versions of the last save that was written down: they do not make room in a full table. */
      lastSaved: new Set(),
      conns: new Set(),
      idleTimer: null,
      syncTokens: { tokens: syncBurst, at: now() },
      /** The saves of everyone in it that may still be written down. */
      savedTokens: { tokens: roomSavedBurst, at: now() },
    };

    const fresh = room.rev === 0 && room.conns.size === 0;
    const { since, versions } = join;
    let from = -1;
    let matched = [];
    // `since` counts only for this room as it was founded (`ep`): a room that
    // was founded again counts its revisions again.
    if (since !== undefined && join.ep === room.ep && since >= room.floor && since <= room.rev) {
      from = since;
      matched = versions.filter((version) => room.versions.has(version) || room.founding.has(version));
    } else if (fresh) {
      // Nothing was written here yet: the room is founded (again), under a new
      // name. What it knew until now leaves the table like any other version.
      // On the file it was founded on before, with nothing written since, it is
      // the room it was: whoever was in it comes back to it.
      if (known && !versions.some((version) => room.founding.has(version))) room.ep = randomBytes(9).toString('base64url');
      for (const version of [...room.versions.keys(), ...room.founding]) {
        if (!versions.includes(version)) forget(room, version);
      }
      room.versions.clear();
      room.lastSaved = new Set();
      room.founding = new Set(versions);
      for (const version of versions) {
        room.versions.set(version, 0);
        room.evicted.delete(version);
      }
      from = 0;
      matched = versions;
    } else if (room.floor === 0 && versions.some((version) => room.founding.has(version))) {
      // Otherwise only with the file the room was founded on. What a version
      // was said to hold later is no ground to start from (anybody who may write
      // can say it); a file that was saved in the room names its revision itself,
      // and its client joins with `since` and `ep`.
      from = 0;
      matched = versions.filter((version) => room.founding.has(version));
    }
    if (from < 0) {
      // Only a room that exists turns a join away like this.
      log({ type: 'reset', id: conn.id, room: room.key, rev: room.rev });
      conn.resetFrom = room.key;
      startJoinTimer(conn);
      send(conn, { t: 'reset', rev: room.rev, v: latestVersions(room) });
      // Those who can save are asked to: what they save is a version the
      // newcomer can load. Only one who could write is worth a save of the
      // room, and the room is asked only so often, whoever is turned back.
      if (!join.ro && verdict.readOnly !== true) {
        const writers = [...room.conns].filter((other) => !other.ro);
        if (writers.length && (!(opt.syncPerMinute > 0) || takeToken(room.syncTokens, opt.syncPerMinute / 60, syncBurst))) {
          const frame = encodeMessage({ t: 'sync' });
          for (const other of writers) write(other, frame);
        }
      }
      return true;
    }
    if (holder) fail(holder, 'replaced', 1008);
    if (!known) {
      rooms.set(room.key, room);
      log({ type: 'room-created', room: room.key, pv: room.pv });
    }

    conn.ro = join.ro || verdict.readOnly === true;
    conn.user = join.user;
    conn.key = join.userKey;
    conn.from = from;
    conn.resetFrom = null;
    const peers = [];
    for (const other of room.conns) {
      const peer = { id: other.id, user: other.user, ro: other.ro };
      if (other.hasPresence) peer.presence = other.presence;
      peers.push(peer);
    }
    clearTimeout(room.idleTimer);
    room.idleTimer = null;
    room.conns.add(conn);
    conn.room = room;
    log({ type: 'join', id: conn.id, room: room.key, from, rev: room.rev, ro: conn.ro });
    // The log goes out in one piece; its bytes are not held against the client
    // until the socket has been empty once.
    conn.socket.cork();
    send(conn, { t: 'hello', pv: room.pv, ep: room.ep, rev: room.rev, from, you: conn.id, ro: conn.ro, v: matched, peers });
    for (let index = from - room.floor; index < room.log.length && !conn.dead; index++) {
      conn.allowance += room.log[index].bytes;
      write(conn, room.log[index].frame);
    }
    conn.socket.uncork();
    if (conn.dead) return true;
    broadcast(room, { t: 'peer', id: conn.id, user: conn.user, ro: conn.ro }, conn);
    return true;
  }

  function onOp(conn, message, text) {
    const { base, id, ops } = message;
    if (!isCount(base)) return false;
    if (typeof id !== 'string' || id.length < 1 || id.length > MAX_OP_ID) return false;
    if (!Array.isArray(ops) || ops.length < 1 || ops.length > MAX_OPS_PER_MESSAGE) return false;
    const raw = rawOps(text);
    if (raw && raw.depth > MAX_OPS_DEPTH) return false;
    const room = conn.room;
    if (!room) return true;
    const nack = (reason) => {
      send(conn, reason ? { t: 'nack', id, rev: room.rev, reason } : { t: 'nack', id, rev: room.rev });
      return true;
    };
    if (conn.ro) return nack('read-only');
    // Every attempt to write is counted, whatever becomes of it.
    if (opt.opsPerSecond > 0 && !takeToken(conn.opTokens, opt.opsPerSecond, opsBurst)) return nack('rate');
    if (base !== room.rev) return nack();
    // The one thing that is looked at: an operation that opens with a table of
    // formats, `{ "t": "st", "f": [...] }`, adds that many formats to the file
    // of everyone, and a connection adds only so many a minute. Counted when
    // the operation would be placed: one that is sent again costs once.
    const table = ops[0];
    if (opt.formatsPerMinute > 0 && isObject(table) && table.t === 'st' && Array.isArray(table.f)) {
      if (!takeToken(conn.formatTokens, opt.formatsPerMinute / 60, formatsBurst, table.f.length)) return nack('rate');
    }
    // The operations go out in the very text they came in (every number, every
    // name in its place); only when that text is not found are they written
    // back as they were read.
    const head = `{"t":"op","rev":${room.rev + 1},"id":${JSON.stringify(id)},"from":${JSON.stringify(conn.id)},"ops":`;
    const frame = encodeFrame(OP_TEXT, Buffer.from(`${head}${raw ? raw.text : JSON.stringify(ops)}}`, 'utf8'));
    room.rev++;
    room.log.push({ rev: room.rev, id, from: conn.id, frame, bytes: frame.length, sequence: ++logSequence });
    room.logBytes += frame.length;
    totalLogBytes += frame.length;
    trimLog(room);
    for (const other of [...room.conns]) write(other, frame);
    return true;
  }

  function onPresence(conn, message) {
    const room = conn.room;
    if (!room) return true;
    // What comes too fast is dropped whole: it is neither passed on nor kept.
    if (opt.presencePerSecond > 0 && !takeToken(conn.presenceTokens, opt.presencePerSecond, presenceBurst)) return true;
    const data = message.data === undefined ? null : message.data;
    const encoded = Buffer.from(JSON.stringify(data), 'utf8');
    if (encoded.length > MAX_PRESENCE_BYTES) return true;
    const frame = encodeFrame(
      OP_TEXT,
      Buffer.concat([Buffer.from(`{"t":"presence","from":${JSON.stringify(conn.id)},"data":`), encoded, Buffer.from('}')]),
    );
    conn.presence = data;
    conn.hasPresence = true;
    for (const other of [...room.conns]) {
      if (other !== conn) write(other, frame);
    }
    return true;
  }

  /**
   * A connection has only so many saves a minute written down, and so has the
   * room, whoever says them: as many connections as one likes do not fill the
   * table faster. A save that one of the two refuses takes nothing from the other.
   */
  function maySave(conn, room) {
    const ofRoom = opt.roomSavedPerMinute > 0;
    if (ofRoom && !takeToken(room.savedTokens, opt.roomSavedPerMinute / 60, roomSavedBurst)) return false;
    if (!(opt.savedPerMinute > 0) || takeToken(conn.savedTokens, opt.savedPerMinute / 60, savedBurst)) return true;
    if (ofRoom) room.savedTokens.tokens += 1;
    return false;
  }

  function onSaved(conn, message) {
    const versions = versionList(message.v, 1);
    if (!versions) return false;
    if (message.rev !== undefined && typeof message.rev !== 'number') return false;
    const room = conn.room;
    if (!room || conn.ro) return true;
    const out = { t: 'saved', from: conn.id, v: versions };
    const rev = message.rev;
    // A revision is believed only when its sender has seen it (it joined at or
    // before it) and when the table does not say otherwise of one of the
    // versions: one that has another revision, or one that left the table.
    // What is not believed is passed on without the revision.
    const seen = isCount(rev) && rev >= room.floor && rev >= conn.from && rev <= room.rev;
    if (seen && versions.every((version) => !room.evicted.has(version) && (revisionOf(room, version) ?? rev) === rev)) {
      const unknown = versions.some((version) => revisionOf(room, version) === undefined);
      if (!unknown) {
        // It says what the table says: there is nothing to write.
        out.rev = rev;
      } else if (!sealed(room) && maySave(conn, room)) {
        room.lastSaved = new Set(versions);
        for (const version of versions) setVersion(room, version, rev);
        out.rev = rev;
      }
    }
    broadcast(room, out, conn);
    return true;
  }

  // --- The outside -----------------------------------------------------------

  const api = {
    /** Serves the WebSocket upgrades of a server that exists. */
    attach(server) {
      if (closed) throw new Error('collab server: it has been closed');
      if (attached.has(server)) return api;
      warnIfOpen();
      const listener = (request, socket, head) => {
        try {
          onUpgrade(server, request, socket, head);
        } catch (error) {
          log({ type: 'error', message: String((error && error.name) || 'Error') });
          socket.destroy();
        }
      };
      attached.set(server, listener);
      server.on('upgrade', listener);
      return api;
    },

    /** Opens a server of its own. Plain HTTP requests are answered with 426. */
    listen(port = 0, host) {
      if (closed) return Promise.reject(new Error('collab server: it has been closed'));
      if (ownServer) return Promise.reject(new Error('collab server: it is listening already'));
      const server = http.createServer((request, response) => {
        request.resume();
        response.writeHead(426, { Upgrade: 'websocket', Connection: 'close', 'Content-Type': 'text/plain' });
        response.end('This is a WebSocket endpoint.\n');
      });
      ownServer = server;
      api.attach(server);
      return new Promise((resolve, reject) => {
        const failed = (error) => {
          ownServer = null;
          server.removeListener('upgrade', attached.get(server));
          attached.delete(server);
          reject(error);
        };
        server.once('error', failed);
        server.listen(port, host, () => {
          server.removeListener('error', failed);
          server.on('error', (error) => log({ type: 'error', message: String(error && error.code) }));
          const address = server.address();
          resolve({ port: address.port, host: address.address });
        });
      });
    },

    /** Cuts every connection, forgets every room and stops every timer. */
    close() {
      closed = true;
      clearInterval(pingTimer);
      pingTimer = null;
      for (const [server, listener] of attached) server.removeListener('upgrade', listener);
      attached.clear();
      for (const conn of [...connections]) drop(conn, 'server-closed');
      for (const room of rooms.values()) clearTimeout(room.idleTimer);
      rooms.clear();
      totalLogBytes = 0;
      const server = ownServer;
      ownServer = null;
      if (!server) return Promise.resolve();
      return new Promise((resolve) => {
        server.close(() => resolve());
        if (typeof server.closeAllConnections === 'function') server.closeAllConnections();
      });
    },

    /** What is in memory now, for diagnosis. */
    rooms() {
      return [...rooms.values()].map((room) => ({
        key: room.key,
        pv: room.pv,
        rev: room.rev,
        floor: room.floor,
        connections: room.conns.size,
        logEntries: room.log.length,
        logBytes: room.logBytes,
      }));
    },
  };
  return api;
}

// --- Run as a program ----------------------------------------------------------

const USAGE = 'usage: node collab-server.mjs [--port 8787] [--host 127.0.0.1] [--path /prefix] [--origin https://app.example]...';

async function main(args) {
  const settings = { port: 8787, host: '127.0.0.1', path: undefined, origins: [] };
  for (let i = 0; i < args.length; i++) {
    const name = args[i];
    if (name === '--help' || name === '-h') {
      console.log(USAGE);
      return;
    }
    const value = args[++i];
    if (name === '--port' && /^\d{1,5}$/.test(value || '') && Number(value) <= 65535) settings.port = Number(value);
    else if (name === '--host' && value) settings.host = value;
    else if (name === '--path' && value) settings.path = value;
    else if (name === '--origin' && value) settings.origins.push(value);
    else {
      console.error(`collab server: cannot read ${name}${value === undefined ? '' : ' ' + value}\n${USAGE}`);
      process.exitCode = 2;
      return;
    }
  }
  const server = createCollabServer({
    path: settings.path,
    origins: settings.origins.length ? settings.origins : undefined,
  });
  const { port, host } = await server.listen(settings.port, settings.host);
  const stop = () => {
    server.close().then(() => {
      process.exitCode = 0;
    });
  };
  // Before the line that says it listens: whoever reads that line may stop it.
  process.once('SIGINT', stop);
  process.once('SIGTERM', stop);
  console.log(`collab relay listening on ws://${host.includes(':') ? `[${host}]` : host}:${port}${settings.path || ''}`);
}

function runAsProgram() {
  try {
    return Boolean(process.argv[1]) && realpathSync(process.argv[1]) === realpathSync(fileURLToPath(import.meta.url));
  } catch {
    return false;
  }
}

if (runAsProgram()) {
  main(process.argv.slice(2)).catch((error) => {
    console.error(`collab server: ${(error && error.message) || error}`);
    process.exitCode = 1;
  });
}
