import type { IncomingMessage, Server } from 'node:http';

/** What `authorize` is asked. */
export interface CollabAuthorizeRequest {
  /** The key of the room: the path the request resolves to, then `#` and the `room` of the join if it named one. */
  room: string;
  /** The `token` of the join message, or else the `token` of the query. */
  token: string | undefined;
  origin: string | undefined;
  /** The request that was upgraded. */
  request: IncomingMessage;
}

/** `true` lets the connection write, `{ readOnly: true }` lets it read, `false` turns it away. */
export type CollabVerdict = boolean | { readOnly?: boolean };

/**
 * A line of diagnosis. It never carries a token. `{ type: 'warning', message }` is the relay saying that it
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
  /** Only paths that start with this are upgraded (the path as it resolves: `/a/../b` is `/b`). Default: every path. */
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
  /** How long an empty room is kept. Default 5 minutes. */
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
  /** The least time between two `sync` messages to a room. Default 5 seconds. */
  syncIntervalMs?: number;
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
