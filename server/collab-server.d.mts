import type { IncomingMessage, Server } from 'node:http';

/** What `authorize` is asked. */
export interface CollabAuthorizeRequest {
  /**
   * The key of the room: the path the request resolves to, with every segment percent-decoded (`/p/%61` is
   * `/p/a`), then `#` and the `room` of the join if it named one. A decoded path never holds a `#`.
   */
  room: string;
  /** The `token` of the join message, or else the `token` of the query. */
  token: string | undefined;
  origin: string | undefined;
  /** The request that was upgraded. */
  request: IncomingMessage;
}

/**
 * `true` lets the connection write, `{ readOnly: true }` lets it read, `false` turns it away. `user`: who
 * the participant is, as the application found out (from the token, a session): its `name` and `color`
 * are told to the room IN PLACE of what the client says of itself, and to the client in its `hello`.
 * Without it the names in a room are what the participants say, and prove nothing.
 */
export type CollabVerdict = boolean | { readOnly?: boolean; user?: { name?: string; color?: string } };

/**
 * A line of diagnosis. It never carries a token or the `key` of a join. `{ type: 'warning', message }` is the relay saying that it
 * was started without `authorize` and without `origins`.
 */
export interface CollabLogEvent {
  type: string;
  id?: string;
  room?: string;
  [detail: string]: unknown;
}

/**
 * Whoever gets into a room reads the file that is edited there and changes it. Without `authorize` that is
 * everyone who can reach the relay: in production `authorize` is not optional. A relay that serves without
 * `authorize` and without `origins` says so once (to `log`, or else to `console.warn`).
 */
export interface CollabServerOptions {
  /**
   * Only paths that start with this are upgraded (the path as it resolves and decoded: `/a/../b` is `/b`,
   * `/%62` is `/b`). Default: every path.
   */
  path?: string;
  /** The origins that are let in. Given, an upgrade from another origin or without one is answered with 403. */
  origins?: string[];
  /** Default 1 MiB. */
  maxMessageBytes?: number;
  /** Default 64. */
  maxConnectionsPerRoom?: number;
  /** Default 1000. */
  maxRooms?: number;
  /** Default 20000. */
  maxLogEntries?: number;
  /** Default 32 MiB. */
  maxLogBytes?: number;
  /**
   * How long an empty room is kept. Default 5 minutes. A room that was let go of and is joined again is another
   * room: it has a new name (`hello.ep`), and a `since` that names the old one is not looked at.
   */
  idleRoomMs?: number;
  /**
   * How long a connection may take to send its `join`, and after a `reset` to join again. Default 10 seconds;
   * 0 waits for ever.
   */
  joinTimeoutMs?: number;
  /** Default 30 seconds; a connection that leaves two pings unanswered is cut. 0 sends none. */
  pingIntervalMs?: number;
  /** A connection with more than this waiting to be written is cut. Default 8 MiB. */
  maxBufferedBytes?: number;
  /** Connections of all rooms together; the next one is told `server-full`. Default 10000. */
  maxConnections?: number;
  /** Connections of one remote address; the next one is told `server-full`. Default 0: as many as it likes. */
  maxConnectionsPerIp?: number;
  /** The logs of all rooms together; over it the oldest operations of the relay are cut. Default 512 MiB. */
  maxTotalLogBytes?: number;
  /** Operations a connection may send a second; over it they are refused with `nack`, reason `rate`. Default 50; 0 is no limit. */
  opsPerSecond?: number;
  /** How many operations a connection may send at once before the rate counts. Default 200. */
  opsBurst?: number;
  /** Presence messages a connection may send a second; over it they are dropped. Default 20; 0 is no limit. */
  presencePerSecond?: number;
  /**
   * How often a minute the writers of a room are sent `sync` (a bucket of as many tokens). Default 2; 0 is no
   * limit.
   */
  syncPerMinute?: number;
  /**
   * Saves a minute of one connection that are written to the table of versions (a bucket of as many tokens);
   * over it a `saved` is passed on without its `rev`, and a client that comes back with that file is not told
   * that the room has heard of it. Default 40, above the thirty a minute a client saves by itself; 0 is no limit.
   */
  savedPerMinute?: number;
  /**
   * Saves a minute of a whole room, whoever says them, that are written to the table of versions (a bucket of
   * as many tokens, beside that of every connection); over it a `saved` is passed on without its `rev`.
   * Default 120; 0 is no limit.
   */
  roomSavedPerMinute?: number;
  /** How many times one connection may join in a minute (after `reset` or `leave`); over it the connection is closed with `error join-rate`. Default 10; 0: any number. */
  joinsPerMinute?: number;
  /**
   * Formats a minute one connection may add: the length of `f` of an operation that opens with
   * `{ "t": "st", "f": [...] }` (a bucket of as many tokens); over it the operation is refused with `nack`,
   * reason `rate`. Default 2000; 0 is no limit.
   */
  formatsPerMinute?: number;
  /**
   * How many versions that left its table a room remembers, so that they are not written again: the latest
   * that left. When another leaves after that, the one that left longest ago is forgotten and may be written
   * like a version the room never had; the room goes on writing new versions. Default 4096; with 0 no `saved`
   * is written to the table.
   */
  maxEvictedVersions?: number;
  /**
   * The clock of the rates: milliseconds that only go forward. Default `performance.now()`. For tests.
   */
  now?(): number;
  /** Without it everyone who can reach the relay may read and write every room. */
  authorize?(request: CollabAuthorizeRequest): CollabVerdict | Promise<CollabVerdict>;
  log?(event: CollabLogEvent): void;
}

export interface CollabRoomInfo {
  key: string;
  /** The version of the protocol of the first to join it. */
  pv: number;
  rev: number;
  /** The revision before the oldest operation that is still in the log. */
  floor: number;
  connections: number;
  logEntries: number;
  logBytes: number;
}

export interface CollabServer {
  /** Serves the WebSocket upgrades of a server that exists. */
  attach(server: Server): CollabServer;
  /** Opens a server of its own; plain HTTP requests are answered with 426. */
  listen(port?: number, host?: string): Promise<{ port: number; host: string }>;
  /** Cuts every connection, forgets every room and stops every timer. */
  close(): Promise<void>;
  rooms(): CollabRoomInfo[];
}

export function createCollabServer(options?: CollabServerOptions): CollabServer;
