// Collaborative editing of a workbook: what mount() uses, and what an application that brings
// its own transport or runs the rules of the relay itself needs.

import { readCollabStamp, writeSharedFile } from '../formats/xlsx/collab-stamp';
import { contentVersion } from './protocol';
import { MemoryServer, type MemoryServerOptions } from './server';
import { broadcastTransport, socketTransport } from './transport';

export { contentVersion, parseClientMessage, parseServerMessage, type ClientMessage, type CollabUser, type ServerMessage } from './protocol';
export { MemoryServer, type MemoryServerOptions, type MemoryTransport } from './server';
export { MAX_PEOPLE, listPeople, type CollabPerson, type People } from './people';
export { CollabSession, PEER_COLORS, type CollabNotice, type CollabNoticeCode, type CollabPeer, type CollabPresence, type CollabState, type CollabTransport, type SessionChange, type SessionHost, type SessionOptions } from './session';
export { broadcastTransport, isSocketUrl, socketTransport, type BroadcastTransportOptions, type PausableTransport, type SocketTransportOptions } from './transport';
export { CollabUnsupportedError, WireError, type WireOp } from './types';

/**
 * `BymOffice.collab`: the relay of the page and the transports, for tests, demos and tabs that edit together
 * without a server; and what a program that runs a CollabSession itself (in Node, or with a view of its own)
 * needs beside the session to read and write files that say which step of the room they hold.
 */
export const collab = {
  /** A relay in this page: `connect()` gives a transport for an editor of the page, `serve(name)` answers other tabs. */
  memoryServer: (options?: MemoryServerOptions): MemoryServer => new MemoryServer(options),
  broadcastTransport,
  socketTransport,
  contentVersion,
  /** The room and the step a file was saved at, for the `stamp` of a session that is made without mount(). */
  readCollabStamp,
  /** Writes the file of a session with the room and the step in it, as mount() saves. */
  writeSharedFile,
};
