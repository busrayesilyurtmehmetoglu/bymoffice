// The rules of the relay, in the page: for tests, for a demo, and for tabs of one browser that
// edit together without a server (one tab holds the room, see serve()). Nothing is kept beyond
// the page. server/collab-server.mjs is the same rules behind a WebSocket.
//
// The rules: a room has a number (rev) and a log. A step that names the rev the room is at is
// taken: the rev goes up by one and everybody gets the step, its author too; any other step is
// answered with "nack" to its author alone. Whoever joins gets the log from where its file
// (or its last step) is. Selections and "saved" are passed on and not kept in the log.

import { TAB_ALIVE_MS, TAB_SILENT_MS } from './transport';
import { MAX_MESSAGE_CHARS, MAX_PRESENCE_BYTES, parseClientMessage, utf8Length, type CollabUser, type ServerMessage } from './protocol';
import type { CollabTransport } from './session';

export interface MemoryServerOptions {
  /** Most steps and bytes the log of a room keeps; the oldest go first. Defaults 20000 and 32 MiB. */
  maxLogEntries?: number;
  maxLogBytes?: number;
  /** Most connections a room takes. Default 64. */
  maxConnections?: number;
  /** How often in a minute at most those in a room are asked to save because someone cannot join. Default 2. */
  syncPerMinute?: number;
  /** The clock. Default: Date.now. */
  now?(): number;
  /**
   * Runs what delivers a message; a test orders the deliveries with it. `lane` names the
   * connection and the direction: what is delivered on one lane has to keep its order.
   * Default: a microtask later.
   */
  deliver?(run: () => void, lane: string): void;
}

interface Entry {
  rev: number;
  id: string;
  from: string;
  ops: unknown[];
  bytes: number;
}

interface Connection {
  /** When it sent its last joins. */
  joins?: number[];
  id: string;
  push(text: string): void;
  room: Room | undefined;
  user: CollabUser;
  ro: boolean;
  presence: unknown;
  gone: boolean;
  /** The secret this connection joined with: nobody else is told it (see ClientMessage "join"). */
  key: string | undefined;
  /** The step it joined at: it has seen the room from there on. */
  from: number;
  /** The room it was told to load the file of again; it may join that one and no other. */
  refused: string | undefined;
}

interface Room {
  /** The name of the room as it was started: another one every time it is started anew. */
  ep: string;
  /** What the file the room was started with is known by. */
  founding: Set<string>;
  key: string;
  rev: number;
  floor: number;
  log: Entry[];
  bytes: number;
  /** The step a version of the file holds. Written once: what a version is known to hold is never changed. */
  versions: Map<string, number>;
  members: Set<Connection>;
  /** The version of the contract its first participant spoke. */
  pv: number | undefined;
  /** The versions the last "saved" that was taken named: never let go of. */
  latest: string[];
  /** Versions the table let go of: never written again. */
  gone: Set<string>;
  /** When those in the room were last asked to save. */
  syncs: number[];
}

/** What a channel between tabs has to be able to do (a BroadcastChannel). */
export interface ChannelLike {
  postMessage(message: unknown): void;
  onmessage: ((event: { data: unknown }) => void) | null;
  close(): void;
}

/** A transport into a relay of the same page. */
export interface MemoryTransport extends CollabTransport {
  /** Cuts the connection, as a network that went away would; resume() connects again. */
  suspend(): void;
  resume(): void;
}

/**
 * Remembers a version that left the table, among the latest 4096 that did: the oldest of them is forgotten
 * when there are more. (The table lets nobody into a room: a version that is written again long after it
 * left does no harm, while a room that stopped writing versions down would keep out whoever comes back
 * with a file it saved.)
 */
function forget(gone: Set<string>, version: string): void {
  gone.delete(version);
  if (gone.size >= 4096) {
    const oldest = gone.values().next();
    if (!oldest.done) gone.delete(oldest.value);
  }
  gone.add(version);
}

/** Most joins of one connection in a minute. */
export const MAX_JOINS_PER_MINUTE = 10;

let epochs = 0;
/** A name for a room that is started: no two rooms of a page, or of two pages, are likely to share one. */
function epochName(): string {
  return `${Date.now().toString(36)}${Math.floor(Math.random() * 0x7fffffff).toString(36)}${(++epochs).toString(36)}`;
}

export class MemoryServer {
  private readonly options: MemoryServerOptions;
  private readonly roomsByKey = new Map<string, Room>();
  private next = 0;
  private lanes = 0;
  private closed = false;
  private readonly served: (() => void)[] = [];

  constructor(options: MemoryServerOptions = {}) {
    this.options = options;
  }

  private deliver(run: () => void, lane: string): void {
    if (this.options.deliver) this.options.deliver(run, lane);
    else void Promise.resolve().then(run);
  }

  /** What the rooms hold, for whoever wants to look. */
  rooms(): { key: string; rev: number; floor: number; connections: number; logEntries: number; logBytes: number }[] {
    return Array.from(this.roomsByKey.values()).map((room) => ({ key: room.key, rev: room.rev, floor: room.floor, connections: room.members.size, logEntries: room.log.length, logBytes: room.bytes }));
  }

  /** A connection as the relay sees it: `push` is given what the relay sends. */
  private accept(push: (text: string) => void): Connection {
    return { id: `c${(++this.next).toString(36)}`, push, room: undefined, user: {}, ro: false, presence: undefined, gone: false, key: undefined, from: 0, refused: undefined };
  }

  private say(connection: Connection, message: ServerMessage): void {
    if (!connection.gone) connection.push(JSON.stringify(message));
  }

  private leaveRoom(connection: Connection): void {
    const room = connection.room;
    if (!room) return;
    connection.room = undefined;
    room.members.delete(connection);
    room.members.forEach((other) => this.say(other, { t: 'left', id: connection.id }));
  }

  private drop(connection: Connection): void {
    if (connection.gone) return;
    this.leaveRoom(connection);
    connection.gone = true;
  }

  private receive(connection: Connection, text: string): void {
    if (connection.gone || this.closed) return;
    let message;
    try {
      if (text.length > MAX_MESSAGE_CHARS) throw new Error('too large');
      message = parseClientMessage(text);
    } catch {
      this.say(connection, { t: 'error', code: 'bad-message' });
      this.drop(connection);
      return;
    }
    if (message.t === 'ping') return this.say(connection, { t: 'pong' });
    if (message.t === 'leave') {
      // Out of its room, and free to join that room again (and no other) over this connection.
      const left = connection.room;
      if (!left) return;
      this.leaveRoom(connection);
      connection.refused = left.key;
      return;
    }
    if (message.t === 'join') {
      // Only so many joins a minute from one connection: leaving and joining over and over tells everybody of it every time.
      const at = this.options.now ? this.options.now() : Date.now();
      connection.joins = (connection.joins || []).filter((time) => at - time < 60000);
      if (connection.joins.length >= MAX_JOINS_PER_MINUTE) {
        this.say(connection, { t: 'error', code: 'join-rate' });
        this.drop(connection);
        return;
      }
      connection.joins.push(at);
      const key = message.room !== undefined ? `#${message.room}` : '';
      // One room per connection: who is in a room stays there, and who was told to load its file again joins that room.
      if (connection.room || (connection.refused !== undefined && connection.refused !== key)) {
        this.say(connection, { t: 'error', code: 'bad-message' });
        this.drop(connection);
        return;
      }
      let room = this.roomsByKey.get(key);
      const made = !room;
      if (!room) room = { key, ep: epochName(), founding: new Set(), rev: 0, floor: 0, log: [], bytes: 0, versions: new Map(), members: new Set(), pv: undefined, latest: [], gone: new Set(), syncs: [] };
      const pv = message.pv === undefined ? 1 : message.pv;
      if (room.pv !== undefined && room.pv !== pv) {
        this.say(connection, { t: 'error', code: 'version' });
        this.drop(connection);
        return;
      }
      // The key is a secret between a participant and the relay: who knows it is that participant, coming back over a
      // new connection, and takes the place of the old one (which the relay may not have seen go yet).
      room.members.forEach((other) => {
        if (message.key === undefined || other.key !== message.key) return;
        this.say(other, { t: 'error', code: 'replaced' });
        this.drop(other);
      });
      if (room.members.size >= (this.options.maxConnections === undefined ? 64 : this.options.maxConnections)) {
        this.say(connection, { t: 'error', code: 'room-full' });
        this.drop(connection);
        return;
      }
      const fresh = room.rev === 0 && room.members.size === 0;
      let from: number | undefined;
      let named: string[] = [];
      // A step is taken on the word of who joins only for this room as it was started: a room that was started anew
      // counts its steps anew, and "step 3" of the one before it says nothing about it.
      if (message.since !== undefined && message.ep === room.ep && message.since >= room.floor && message.since <= room.rev) {
        from = message.since;
        named = message.v.filter((version) => room!.versions.has(version));
      } else if (fresh) {
        // Started (anew) with the file of who comes first. With the file it was started with before, and nothing
        // done since, it is the room it was: whoever was in it comes back to it.
        if (!message.v.some((version) => room!.founding.has(version))) room.ep = epochName();
        room.versions.clear();
        room.founding.clear();
        for (const version of message.v) {
          room.versions.set(version, 0);
          room.founding.add(version);
        }
        from = 0;
        named = message.v.slice();
      } else if (room.floor === 0 && message.v.some((version) => room!.founding.has(version))) {
        // Otherwise only with the file the room was started with. What a file was said to hold later is no ground to
        // start from: anybody may have said it. A file that was saved in the room says its step itself (see `since`).
        from = 0;
        named = message.v.filter((version) => room!.founding.has(version));
      }
      if (from === undefined) {
        // The versions of the latest step any is known for.
        let latest = -1;
        room.versions.forEach((rev) => {
          if (rev > latest) latest = rev;
        });
        const known: string[] = [];
        room.versions.forEach((rev, version) => {
          if (rev === latest) known.push(version);
        });
        this.say(connection, { t: 'reset', rev: room.rev, v: known.slice(0, 8) });
        connection.refused = key;
        // Nobody can come in while no file holds the steps of the room: those who are in it are asked to save.
        // Only for someone who may write (a reader has no say in when a file is saved), only those who can save are asked, and not often.
        const now = this.options.now ? this.options.now() : Date.now();
        room.syncs = room.syncs.filter((at) => now - at < 60000);
        if (message.ro !== true && room.members.size > 0 && room.syncs.length < (this.options.syncPerMinute === undefined ? 2 : this.options.syncPerMinute)) {
          room.syncs.push(now);
          room.members.forEach((other) => {
            if (!other.ro) this.say(other, { t: 'sync' });
          });
        }
        return;
      }
      if (made) this.roomsByKey.set(key, room);
      room.pv = pv;
      connection.refused = undefined;
      connection.key = message.key;
      connection.from = from;
      connection.user = message.user || {};
      connection.ro = message.ro === true;
      connection.presence = undefined;
      const peers: Extract<ServerMessage, { t: 'hello' }>['peers'] = [];
      room.members.forEach((other) => {
        peers.push(other.presence !== undefined && other.presence !== null ? { id: other.id, user: other.user, ro: other.ro, presence: other.presence } : { id: other.id, user: other.user, ro: other.ro });
        this.say(other, { t: 'peer', id: connection.id, user: connection.user, ro: connection.ro });
      });
      room.members.add(connection);
      connection.room = room;
      this.say(connection, { t: 'hello', rev: room.rev, from, ep: room.ep, user: connection.user, you: connection.id, ro: connection.ro, pv, v: Array.from(new Set(named)), peers });
      for (const entry of room.log) if (entry.rev > from) this.say(connection, { t: 'op', rev: entry.rev, id: entry.id, from: entry.from, ops: entry.ops });
      return;
    }
    const room = connection.room;
    // Not in a room (not yet, or told to load its file again): nothing is taken.
    if (!room) return;
    if (message.t === 'op') {
      if (connection.ro) return this.say(connection, { t: 'nack', id: message.id, rev: room.rev, reason: 'read-only' });
      if (message.base !== room.rev) return this.say(connection, { t: 'nack', id: message.id, rev: room.rev });
      room.rev++;
      const entry: Entry = { rev: room.rev, id: message.id, from: connection.id, ops: message.ops, bytes: utf8Length(text) };
      room.log.push(entry);
      room.bytes += entry.bytes;
      const maxEntries = this.options.maxLogEntries === undefined ? 20000 : this.options.maxLogEntries;
      const maxBytes = this.options.maxLogBytes === undefined ? 32 * 1024 * 1024 : this.options.maxLogBytes;
      while (room.log.length > 0 && (room.log.length > maxEntries || room.bytes > maxBytes)) {
        const oldest = room.log.shift() as Entry;
        room.bytes -= oldest.bytes;
        room.floor = oldest.rev;
      }
      room.versions.forEach((rev, version) => {
        if (rev >= room.floor) return;
        room.versions.delete(version);
        forget(room.gone, version);
      });
      room.members.forEach((other) => this.say(other, { t: 'op', rev: entry.rev, id: entry.id, from: entry.from, ops: entry.ops }));
      return;
    }
    if (message.t === 'presence') {
      if (utf8Length(JSON.stringify(message.data)) > MAX_PRESENCE_BYTES) return;
      connection.presence = message.data;
      room.members.forEach((other) => {
        if (other !== connection) this.say(other, { t: 'presence', from: connection.id, data: message.data });
      });
      return;
    }
    if (message.t === 'saved') {
      if (connection.ro) return;
      // Only a step the sender has seen the room at; and what a version is known to hold stays what it is.
      const claimed = message.rev !== undefined && message.rev >= room.floor && message.rev >= connection.from && message.rev <= room.rev ? message.rev : undefined;
      // What a version is known to hold stays what it is, and a version the table let go of is never written
      // again: a "saved" that says otherwise is passed on without its step.
      let rev = claimed;
      for (const version of message.v) {
        if (room.gone.has(version) || (room.versions.has(version) && room.versions.get(version) !== claimed)) rev = undefined;
      }
      if (rev !== undefined) {
        for (const version of message.v) if (!room.versions.has(version)) room.versions.set(version, rev);
        room.latest = message.v.slice();
        // The oldest go when there are too many, but never what the room started with.
        while (room.versions.size > 16) {
          let oldest: string | undefined;
          room.versions.forEach((at, version) => {
            if (oldest === undefined && at > 0 && room.latest.indexOf(version) < 0) oldest = version;
          });
          if (oldest === undefined) break;
          room.versions.delete(oldest);
          forget(room.gone, oldest);
        }
      }
      room.members.forEach((other) => {
        if (other === connection) return;
        this.say(other, rev !== undefined ? { t: 'saved', from: connection.id, v: message.v, rev } : { t: 'saved', from: connection.id, v: message.v });
      });
    }
  }

  /** A transport for a session in this page. */
  connect(): MemoryTransport {
    let listener: ((text: string) => void) | undefined;
    let state: ((state: 'open' | 'closed') => void) | undefined;
    let connection: Connection | undefined;
    let over = false;
    const lane = ++this.lanes;
    const up = `${lane}:up`;
    const down = `${lane}:down`;
    const open = (): void => {
      if (over || connection || this.closed) return;
      const mine = this.accept((text) =>
        this.deliver(() => {
          if (connection === mine && listener) listener(text);
        }, down),
      );
      connection = mine;
      this.deliver(() => {
        if (connection === mine && state) state('open');
      }, down);
    };
    const cut = (): void => {
      const mine = connection;
      if (!mine) return;
      connection = undefined;
      this.deliver(() => this.drop(mine), up);
      if (state) state('closed');
    };
    return {
      send: (text) => {
        const mine = connection;
        if (mine) this.deliver(() => this.receive(mine, text), up);
      },
      onMessage: (callback) => {
        listener = callback;
      },
      onState: (callback) => {
        state = callback;
        open();
      },
      close: () => {
        over = true;
        const mine = connection;
        connection = undefined;
        if (mine) this.deliver(() => this.drop(mine), up);
      },
      suspend: cut,
      resume: open,
    };
  }

  /**
   * Answers the transports that other tabs of this browser open with broadcastTransport() under
   * the same name. Returns what stops it. The room lives in this page and goes with it.
   */
  serve(name: string, channel?: ChannelLike): () => void {
    const port = channel || (new BroadcastChannel(channelName(name)) as unknown as ChannelLike);
    const connections = new Map<string, Connection>();
    /** When each tab was last heard of: one that was closed without a word is let go of, so that nobody stays listed who is gone. */
    const heard = new Map<string, number>();
    const clock = (): number => (this.options.now ? this.options.now() : Date.now());
    const sweep = setInterval(() => {
      const now = clock();
      connections.forEach((connection, from) => {
        if (now - (heard.get(from) || 0) <= TAB_SILENT_MS) return;
        this.drop(connection);
        connections.delete(from);
        heard.delete(from);
        port.postMessage({ to: from, kind: 'bye' });
      });
    }, TAB_ALIVE_MS);
    const timer = sweep as unknown as { unref?(): void };
    if (typeof timer.unref === 'function') timer.unref();
    port.onmessage = (event) => {
      const data = event.data as { to?: unknown; from?: unknown; kind?: unknown; text?: unknown } | null;
      if (!data || data.to !== 'relay' || typeof data.from !== 'string' || data.from.length > 64) return;
      const from = data.from;
      let connection = connections.get(from);
      if (connection) heard.set(from, clock());
      if (data.kind === 'alive') {
        // A tab the relay let go of (it slept too long) is told, and opens anew.
        if (!connection || connection.gone) port.postMessage({ to: from, kind: 'bye' });
        return;
      }
      if (data.kind === 'open') {
        if (connection) this.drop(connection);
        connection = this.accept((text) => port.postMessage({ to: from, kind: 'msg', text }));
        connections.set(from, connection);
        heard.set(from, clock());
        port.postMessage({ to: from, kind: 'accept' });
      } else if (data.kind === 'close') {
        if (connection) this.drop(connection);
        connections.delete(from);
      } else if (data.kind === 'msg' && typeof data.text === 'string') {
        // A tab the relay does not know (the page that held the room was loaded again): it has to open anew.
        if (!connection || connection.gone) port.postMessage({ to: from, kind: 'bye' });
        else this.receive(connection, data.text);
      }
    };
    const stop = (): void => {
      clearInterval(sweep);
      connections.forEach((connection, from) => {
        this.drop(connection);
        port.postMessage({ to: from, kind: 'bye' });
      });
      connections.clear();
      port.onmessage = null;
      port.close();
    };
    this.served.push(stop);
    return stop;
  }

  /** Ends every room. */
  close(): void {
    this.closed = true;
    for (const stop of this.served.splice(0)) stop();
    this.roomsByKey.forEach((room) => room.members.forEach((connection) => (connection.gone = true)));
    this.roomsByKey.clear();
  }
}

/** The name of the channel the tabs of a room talk over. */
export function channelName(name: string): string {
  return `bymoffice-collab:${name}`;
}
