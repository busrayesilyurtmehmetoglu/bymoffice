// The messages a participant and the relay exchange (see docs/collab-sozlesmesi.md). The relay
// orders the operations and knows nothing of what they say; everything else is here: what is
// sent, and how what arrives is read. Nothing that arrives is trusted.

import { WireError } from './types';

/** Longest message, in characters of its JSON text (the relay counts bytes of UTF-8: at most three times as many). */
export const MAX_MESSAGE_CHARS = 1024 * 1024;
/** Most bytes a step of operations may take on the wire; the rest of the megabyte is for what stands around it. */
export const MAX_STEP_BYTES = 900 * 1024;
export const MAX_PRESENCE_BYTES = 16 * 1024;
const MAX_VERSIONS = 8;
const MAX_VERSION_CHARS = 200;
const MAX_PEERS = 1024;

/** How a participant is shown to the others. */
export interface CollabUser {
  name?: string;
  /** "#RRGGBB". */
  color?: string;
}

export interface PeerMessage {
  id: string;
  user: CollabUser;
  ro: boolean;
  presence?: unknown;
}

export type ClientMessage =
  | { t: 'join'; v: string[]; since?: number; ep?: string; user?: CollabUser; token?: string; ro?: boolean; room?: string; pv?: number; key?: string }
  | { t: 'op'; base: number; id: string; ops: unknown[] }
  | { t: 'presence'; data: unknown }
  | { t: 'saved'; v: string[]; rev?: number }
  | { t: 'ping' }
  /** Out of the room, over a connection that stays: the next message of it is a `join` of that room again. */
  | { t: 'leave' };

export type ServerMessage =
  | { t: 'hello'; rev: number; from: number; ep: string; user?: CollabUser; you: string; ro: boolean; pv?: number; v: string[]; peers: PeerMessage[] }
  | { t: 'op'; rev: number; id: string; from: string; ops: unknown[] }
  | { t: 'nack'; id: string; rev: number; reason?: string }
  | { t: 'reset'; rev: number; v: string[] }
  | { t: 'presence'; from: string; data: unknown }
  | { t: 'saved'; from: string; v: string[]; rev?: number }
  | { t: 'peer'; id: string; user: CollabUser; ro: boolean }
  | { t: 'left'; id: string }
  | { t: 'pong' }
  | { t: 'sync' }
  | { t: 'error'; code: string };

const fail = (what: string): never => {
  throw new WireError(`Invalid ${what}`);
};

const isRev = (value: unknown): value is number => typeof value === 'number' && Math.floor(value) === value && value >= 0 && value <= Number.MAX_SAFE_INTEGER;

function text(value: unknown, max: number, what: string): string {
  if (typeof value !== 'string' || value.length === 0 || value.length > max) return fail(what);
  return value;
}

function versions(value: unknown, min: number): string[] {
  if (!Array.isArray(value) || value.length < min || value.length > MAX_VERSIONS) return fail('versions');
  return value.map((item) => text(item, MAX_VERSION_CHARS, 'version'));
}

/** A name and a colour as they may be shown: plain text of one line, a colour of six hex digits. */
/**
 * Characters that take no room or turn the text around: the soft hyphen, the combining grapheme joiner, the
 * Arabic letter mark, Hangul fillers, the Mongolian vowel separator, spaces and joiners of no width, marks,
 * embeddings, overrides and isolates of direction, deprecated format characters, the byte order mark, the
 * interlinear annotation marks, the blank of Braille, variation selectors, and the block of tags.
 */
const INVISIBLE = /[\u00AD\u034F\u061C\u115F\u1160\u180E\u200B-\u200F\u202A-\u202E\u2060-\u206F\u2800\u3164\uFE00-\uFE0F\uFEFF\uFFA0\uFFF9-\uFFFB]|\uDB40[\uDC00-\uDFFF]/g;

/** Most code points of a name that is shown. */
export const MAX_NAME_LENGTH = 64;

/**
 * A name as it is shown: text of someone else. Control characters become spaces; characters that
 * turn the direction of the text around, or hide in it (marks and overrides of direction, isolates,
 * joiners of no width, the byte order mark), are taken out, so that a name cannot rearrange or
 * imitate what stands beside it; runs of space are one space; and it is cut at 64 code points
 * (never through one).
 */
export function cleanName(value: string): string {
  // In this order, so that cleaning what was cleaned changes nothing: halves of characters go first, then what
  // cannot be seen, then space is made one, the ends are trimmed, and the length is cut (and trimmed again).
  const whole = Array.from(value.slice(0, 4096))
    .filter((point) => !(point.length === 1 && point.charCodeAt(0) >= 0xd800 && point.charCodeAt(0) <= 0xdfff))
    .join('');
  const plain = whole
    .replace(/[\u0000-\u001F\u007F-\u009F\u2028\u2029]+/g, ' ')
    .replace(INVISIBLE, '')
    .replace(/\s+/g, ' ')
    .trim();
  return Array.from(plain).slice(0, MAX_NAME_LENGTH).join('').trim();
}

export function cleanUser(value: unknown): CollabUser {
  const user: CollabUser = {};
  if (!value || typeof value !== 'object') return user;
  const from = value as { name?: unknown; color?: unknown };
  if (typeof from.name === 'string') {
    const name = cleanName(from.name);
    if (name !== '') user.name = name;
  }
  if (typeof from.color === 'string' && /^#[0-9A-Fa-f]{6}$/.test(from.color)) user.color = from.color.toUpperCase();
  return user;
}

function parse(textOf: string): Record<string, unknown> {
  if (typeof textOf !== 'string' || textOf.length > MAX_MESSAGE_CHARS) return fail('message');
  let value: unknown;
  try {
    value = JSON.parse(textOf);
  } catch {
    return fail('message');
  }
  if (!value || typeof value !== 'object' || Array.isArray(value)) return fail('message');
  return value as Record<string, unknown>;
}

function peer(value: unknown): PeerMessage {
  if (!value || typeof value !== 'object') return fail('peer');
  const from = value as Record<string, unknown>;
  const out: PeerMessage = { id: text(from.id, 100, 'peer'), user: cleanUser(from.user), ro: from.ro === true };
  if (from.presence !== undefined && from.presence !== null) out.presence = from.presence;
  return out;
}

/** Reads a message of the relay. Throws WireError for anything else. */
export function parseServerMessage(textOf: string): ServerMessage {
  const m = parse(textOf);
  switch (m.t) {
    case 'hello': {
      if (!isRev(m.rev) || !isRev(m.from) || m.from > m.rev) return fail('hello');
      if (!Array.isArray(m.peers) || m.peers.length > MAX_PEERS) return fail('hello');
      if (typeof m.ep !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(m.ep)) return fail('hello');
      const hello: ServerMessage = { t: 'hello', rev: m.rev, from: m.from, ep: m.ep, you: text(m.you, 100, 'hello'), ro: m.ro === true, v: m.v === undefined ? [] : versions(m.v, 0), peers: m.peers.map(peer) };
      // The participant as the room has it, when the relay says (it may have given the name itself).
      if (m.user !== undefined) (hello as { user?: CollabUser }).user = cleanUser(m.user);
      // The version of the contract the room speaks, when the relay says: the session compares it with its own.
      // A relay of another language may write it as a text of digits ("3"); nothing else is a version.
      if (m.pv !== undefined) {
        const pv = typeof m.pv === 'string' && /^[1-9][0-9]{0,3}$/.test(m.pv) ? Number(m.pv) : m.pv;
        if (!isRev(pv) || pv < 1 || pv > 1000) return fail('hello');
        (hello as { pv?: number }).pv = pv;
      }
      return hello;
    }
    case 'op':
      if (!isRev(m.rev) || !Array.isArray(m.ops)) return fail('op');
      return { t: 'op', rev: m.rev, id: text(m.id, 100, 'op'), from: text(m.from, 100, 'op'), ops: m.ops };
    case 'nack': {
      if (!isRev(m.rev)) return fail('nack');
      const out: ServerMessage = { t: 'nack', id: text(m.id, 100, 'nack'), rev: m.rev };
      if (typeof m.reason === 'string') out.reason = m.reason.slice(0, 64);
      return out;
    }
    case 'reset':
      return { t: 'reset', rev: isRev(m.rev) ? m.rev : 0, v: m.v === undefined ? [] : versions(m.v, 0) };
    case 'presence':
      return { t: 'presence', from: text(m.from, 100, 'presence'), data: m.data };
    case 'saved': {
      const out: ServerMessage = { t: 'saved', from: text(m.from, 100, 'saved'), v: versions(m.v, 1) };
      if (isRev(m.rev)) out.rev = m.rev;
      return out;
    }
    case 'peer':
      return { t: 'peer', id: text(m.id, 100, 'peer'), user: cleanUser(m.user), ro: m.ro === true };
    case 'left':
      return { t: 'left', id: text(m.id, 100, 'left') };
    case 'pong':
      return { t: 'pong' };
    case 'sync':
      return { t: 'sync' };
    case 'error':
      return { t: 'error', code: typeof m.code === 'string' ? m.code.slice(0, 64) : 'error' };
    default:
      return fail('message');
  }
}

/** Reads a message of a participant, as the relay does. Throws WireError for anything else. */
export function parseClientMessage(textOf: string): ClientMessage {
  const m = parse(textOf);
  switch (m.t) {
    case 'join': {
      const out: ClientMessage = { t: 'join', v: versions(m.v, 0) };
      if (m.since !== undefined) {
        if (!isRev(m.since)) return fail('join');
        out.since = m.since;
      }
      if (m.ep !== undefined) {
        if (typeof m.ep !== 'string' || !/^[A-Za-z0-9_-]{1,64}$/.test(m.ep)) return fail('join');
        out.ep = m.ep;
      }
      if (m.user !== undefined) out.user = cleanUser(m.user);
      if (m.token !== undefined) out.token = text(m.token, 4096, 'join');
      if (m.ro === true) out.ro = true;
      if (m.room !== undefined) out.room = text(m.room, 200, 'join');
      if (m.pv !== undefined) {
        if (!isRev(m.pv) || m.pv < 1 || m.pv > 1000) return fail('join');
        out.pv = m.pv;
      }
      if (m.key !== undefined) {
        out.key = text(m.key, 40, 'join');
        if (!/^[A-Za-z0-9_-]+$/.test(out.key)) return fail('join');
      }
      if (out.room !== undefined && out.room.indexOf('#') >= 0) return fail('join');
      return out;
    }
    case 'op':
      if (!isRev(m.base) || !Array.isArray(m.ops) || m.ops.length < 1 || m.ops.length > 10001) return fail('op');
      return { t: 'op', base: m.base, id: text(m.id, 100, 'op'), ops: m.ops };
    case 'presence':
      return { t: 'presence', data: m.data === undefined ? null : m.data };
    case 'saved': {
      const out: ClientMessage = { t: 'saved', v: versions(m.v, 1) };
      if (isRev(m.rev)) out.rev = m.rev;
      return out;
    }
    case 'ping':
      return { t: 'ping' };
    case 'leave':
      return { t: 'leave' };
    default:
      return fail('message');
  }
}

/** How many bytes a text takes as UTF-8. */
export function utf8Length(value: string): number {
  let bytes = 0;
  for (let i = 0; i < value.length; i++) {
    const c = value.charCodeAt(i);
    if (c < 0x80) bytes += 1;
    else if (c < 0x800) bytes += 2;
    else if (c >= 0xd800 && c <= 0xdbff && i + 1 < value.length) {
      bytes += 4;
      i++;
    } else bytes += 3;
  }
  return bytes;
}

/**
 * A short text that stands for the bytes of a file: two hashes of 32 bits and the length. Not
 * a secret and no proof: it tells one version of a file from another when the server names none.
 */
export function contentVersion(bytes: Uint8Array): string {
  let a = 0x811c9dc5;
  let b = 0x01000193;
  for (let i = 0; i < bytes.length; i++) {
    a = Math.imul(a ^ bytes[i], 0x01000193);
    b = (Math.imul(b, 31) + bytes[i] + (b >>> 27)) | 0;
  }
  const hex = (value: number): string => `0000000${(value >>> 0).toString(16)}`.slice(-8);
  return `h:${bytes.length}:${hex(a)}${hex(b)}`;
}
