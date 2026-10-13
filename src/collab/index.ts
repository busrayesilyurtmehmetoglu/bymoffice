// Collaborative editing of a workbook: what mount() uses, and what an application that brings
// its own transport or runs the rules of the relay itself needs.

import { contentVersion } from './protocol';
import { MemoryServer, type MemoryServerOptions } from './server';
import { broadcastTransport, socketTransport } from './transport';

export { contentVersion, parseClientMessage, parseServerMessage, type ClientMessage, type CollabUser, type ServerMessage } from './protocol';
export { MemoryServer, type MemoryServerOptions, type MemoryTransport } from './server';
export { CollabSession, PEER_COLORS, type CollabNotice, type CollabNoticeCode, type CollabPeer, type CollabPresence, type CollabState, type CollabTransport, type SessionChange, type SessionHost, type SessionOptions } from './session';
export { broadcastTransport, isSocketUrl, socketTransport, type BroadcastTransportOptions, type PausableTransport, type SocketTransportOptions } from './transport';
export { CollabUnsupportedError, WireError, type WireOp } from './types';

/** `BymOffice.collab`: the relay of the page and the transports, for tests, demos and tabs that edit together without a server. */
export const collab = {
  /** A relay in this page: `connect()` gives a transport for an editor of the page, `serve(name)` answers other tabs. */
  memoryServer: (options?: MemoryServerOptions): MemoryServer => new MemoryServer(options),
  broadcastTransport,
  socketTransport,
  contentVersion,
};
