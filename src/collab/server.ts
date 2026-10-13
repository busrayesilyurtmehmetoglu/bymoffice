// The rules of the relay, in the page: for tests, for a demo, and for tabs of one browser that
// edit together without a server (one tab holds the room, see serve()). Nothing is kept beyond
// the page. server/collab-server.mjs is the same rules behind a WebSocket.
//
// The rules: a room has a number (rev) and a log. A step that names the rev the room is at is
// taken: the rev goes up by one and everybody gets the step, its author too; any other step is
// answered with "nack" to its author alone. Whoever joins gets the log from where its file
// (or its last step) is. Selections and "saved" are passed on and not kept in the log.

import { MAX_MESSAGE_CHARS, MAX_PRESENCE_BYTES, parseClientMessage, utf8Length, type CollabUser, type ServerMessage } from './protocol';
import type { CollabTransport } from './session';

export interface MemoryServerOptions {
  /** Most steps and bytes the log of a room keeps; the oldest go first. Defaults 20000 and 32 MiB. */
  maxLogEntries?: number;
  maxLogBytes?: number;
  /** Most connections a room takes. Default 64. */
  maxConnections?: number;
  /** How often at most those in a room are asked to save because someone cannot join. Default 5000. */
  syncIntervalMs?: number;
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
  id: string;
  push(text: string): void;
  room: Room | undefined;
  user: CollabUser;
  ro: boolean;
  presence: unknown;
  gone: boolean;
  /** What every step of this connection is named with (see ClientMessage "join"). */
  key: string | undefined;
  /** The step it joined at: it has seen the room from there on. */
  from: number;
  /** The room it was told to load the file of again; it may join that one and no other. */
  refused: string | undefined;
}

interface Room {
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
  syncAt: number;
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
    if (message.t === 'join') {
      const key = message.room !== undefined ? `#${message.room}` : '';
      // One room per connection: who is in a room stays there, and who was told to load its file again joins that room.
      if (connection.room || (connection.refused !== undefined && connection.refused !== key)) {
        this.say(connection, { t: 'error', code: 'bad-message' });
        this.drop(connection);
        return;
      }
      let room = this.roomsByKey.get(key);
      const made = !room;
      if (!room) room = { key, rev: 0, floor: 0, log: [], bytes: 0, versions: new Map(), members: new Set(), pv: undefined, syncAt: 0 };
      const pv = message.pv === undefined ? 1 : message.pv;
      if (room.pv !== undefined && room.pv !== pv) {
        this.say(connection, { t: 'error', code: 'version' });
        this.drop(connection);
        return;
      }
      let taken = false;
      room.members.forEach((other) => {
        if (message.key !== undefined && other.key === message.key) taken = true;
      });
      if (taken) {
        this.say(connection, { t: 'error', code: 'key-taken' });
        this.drop(connection);
        return;
      }
      if (room.members.size >= (this.options.maxConnections === undefined ? 64 : this.options.maxConnections)) {
        this.say(connection, { t: 'error', code: 'room-full' });
        this.drop(connection);
        return;
      }
      const fresh = room.rev === 0 && room.members.size === 0;
      let from: number | undefined;
      let named: string[] = [];
      if (message.since !== undefined && message.since >= room.floor && message.since <= room.rev && !(fresh && message.since === 0)) {
        from = message.since;
        named = message.v.filter((version) => room!.versions.has(version));
      } else {
        for (const version of message.v) {
          const rev = room.versions.get(version);
          if (rev !== undefined && rev >= room.floor && (from === undefined || rev > from)) from = rev;
        }
        if (from !== undefined) named = message.v.filter((version) => room!.versions.get(version) === from);
        else if (fresh) {
          room.versions.clear();
          for (const version of message.v) room.versions.set(version, 0);
          from = 0;
          named = message.v.slice();
        }
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
        const now = Date.now();
        if (room.members.size > 0 && now - room.syncAt >= (this.options.syncIntervalMs === undefined ? 5000 : this.options.syncIntervalMs)) {
          room.syncAt = now;
          room.members.forEach((other) => this.say(other, { t: 'sync' }));
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
      this.say(connection, { t: 'hello', rev: room.rev, from, you: connection.id, ro: connection.ro, pv, v: Array.from(new Set(named)), peers });
      for (const entry of room.log) if (entry.rev > from) this.say(connection, { t: 'op', rev: entry.rev, id: entry.id, from: entry.from, ops: entry.ops });
      return;
    }
    const room = connection.room;
    // Not in a room (not yet, or told to load its file again): nothing is taken.
    if (!room) return;
    if (message.t === 'op') {
      if (connection.ro) return this.say(connection, { t: 'nack', id: message.id, rev: room.rev, reason: 'read-only' });
      // A step is named with what its connection joined under: nobody names a step as someone else would.
      if (connection.key === undefined || message.id.slice(0, connection.key.length + 1) !== `${connection.key}:`) return this.say(connection, { t: 'nack', id: message.id, rev: room.rev, reason: 'id' });
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
        if (rev < room.floor) room.versions.delete(version);
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
      const rev = message.rev !== undefined && message.rev >= room.floor && message.rev >= connection.from && message.rev <= room.rev ? message.rev : undefined;
      if (rev !== undefined) {
        for (const version of message.v) if (!room.versions.has(version)) room.versions.set(version, rev);
        // The oldest go when there are too many, but never what the room started with.
        while (room.versions.size > 16) {
          let oldest: string | undefined;
          room.versions.forEach((at, version) => {
            if (oldest === undefined && at > 0) oldest = version;
          });
          if (oldest === undefined) break;
          room.versions.delete(oldest);
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
    port.onmessage = (event) => {
      const data = event.data as { to?: unknown; from?: unknown; kind?: unknown; text?: unknown } | null;
      if (!data || data.to !== 'relay' || typeof data.from !== 'string' || data.from.length > 64) return;
      const from = data.from;
      let connection = connections.get(from);
      if (data.kind === 'open') {
        if (connection) this.drop(connection);
        connection = this.accept((text) => port.postMessage({ to: from, kind: 'msg', text }));
        connections.set(from, connection);
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
