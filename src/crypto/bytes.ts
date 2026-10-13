// Byte helpers of the cryptographic code: nothing here knows a file format.

/** Zeroes buffers that held a key, a password or anything derived from them. */
export function wipe(...buffers: (Uint8Array | Uint32Array | Int32Array | undefined | null)[]): void {
  for (let i = 0; i < buffers.length; i++) {
    const buffer = buffers[i];
    if (buffer) buffer.fill(0);
  }
}

/**
 * Whether two byte strings are the same, in a time that depends on their
 * lengths only: where the first difference is cannot be told from how long
 * the comparison took.
 */
export function constantTimeEqual(a: Uint8Array, b: Uint8Array): boolean {
  let diff = a.length ^ b.length;
  const length = Math.min(a.length, b.length);
  for (let i = 0; i < length; i++) diff |= a[i] ^ b[i];
  return diff === 0;
}

/** The UTF-16 code units of a text, little-endian: how Office hashes a password. */
export function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >>> 8;
  }
  return out;
}

export function concatBytes(...parts: Uint8Array[]): Uint8Array {
  let size = 0;
  for (let i = 0; i < parts.length; i++) size += parts[i].length;
  const out = new Uint8Array(size);
  let at = 0;
  for (let i = 0; i < parts.length; i++) {
    out.set(parts[i], at);
    at += parts[i].length;
  }
  return out;
}

export function le32(value: number): Uint8Array {
  return new Uint8Array([value & 0xff, (value >>> 8) & 0xff, (value >>> 16) & 0xff, (value >>> 24) & 0xff]);
}

const HEX = '0123456789abcdef';

export function toHex(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i++) out += HEX[bytes[i] >>> 4] + HEX[bytes[i] & 15];
  return out;
}

export function fromHex(text: string): Uint8Array {
  if (text.length % 2 !== 0 || /[^0-9a-fA-F]/.test(text)) throw new Error('Not a hexadecimal string');
  const out = new Uint8Array(text.length / 2);
  for (let i = 0; i < out.length; i++) out[i] = parseInt(text.substr(i * 2, 2), 16);
  return out;
}

const BASE64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';

export function toBase64(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i];
    const b = i + 1 < bytes.length ? bytes[i + 1] : 0;
    const c = i + 2 < bytes.length ? bytes[i + 2] : 0;
    out += BASE64[a >> 2] + BASE64[((a & 3) << 4) | (b >> 4)];
    out += i + 1 < bytes.length ? BASE64[((b & 15) << 2) | (c >> 6)] : '=';
    out += i + 2 < bytes.length ? BASE64[c & 63] : '=';
  }
  return out;
}

/** The bytes of a base64 text, white space aside; undefined for anything that is not base64. */
export function fromBase64(text: string): Uint8Array | undefined {
  const clean = text.replace(/[ \t\r\n]/g, '');
  if (clean.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(clean)) return undefined;
  const padding = clean.endsWith('==') ? 2 : clean.endsWith('=') ? 1 : 0;
  const out = new Uint8Array((clean.length / 4) * 3 - padding);
  let at = 0;
  for (let i = 0; i < clean.length; i += 4) {
    let word = 0;
    for (let k = 0; k < 4; k++) {
      const code = clean.charCodeAt(i + k);
      // A-Z, a-z, 0-9, + and /; "=" counts as zero.
      const value = code >= 97 ? code - 71 : code >= 65 ? code - 65 : code >= 48 && code <= 57 ? code + 4 : code === 43 ? 62 : code === 47 ? 63 : 0;
      word = (word << 6) | value;
    }
    if (at < out.length) out[at++] = word >>> 16;
    if (at < out.length) out[at++] = (word >>> 8) & 0xff;
    if (at < out.length) out[at++] = word & 0xff;
  }
  return out;
}

/** What randomBytes() needs of its surroundings: the `crypto` of a browser, a worker or Node. */
export interface RandomScope {
  crypto?: { getRandomValues?: (array: Uint8Array) => unknown };
}

// `globalThis` does not exist in the old WebViews the engine still runs in.
const scope: RandomScope = (typeof globalThis !== 'undefined' ? globalThis : typeof self !== 'undefined' ? self : {}) as RandomScope;

/** The most bytes getRandomValues() fills in one call. */
const RANDOM_QUOTA = 65536;

/**
 * `count` bytes from the random number generator of the system. Throws where
 * there is none: salts and keys are never made from Math.random(), so that a
 * file is not encrypted with something that can be guessed.
 */
export function randomBytes(count: number, from: RandomScope = scope): Uint8Array {
  const source = from.crypto;
  if (!source || typeof source.getRandomValues !== 'function') throw new Error('Cannot encrypt: this environment has no secure source of random numbers (crypto.getRandomValues)');
  const out = new Uint8Array(count);
  for (let at = 0; at < count; at += RANDOM_QUOTA) source.getRandomValues(out.subarray(at, Math.min(count, at + RANDOM_QUOTA)));
  // A stand-in that fills nothing in would give an all-zero salt and key. (Eight zero bytes come by chance once in 2^64 times.)
  if (count >= 8) {
    let any = 0;
    for (let i = 0; i < count; i++) any |= out[i];
    if (any === 0) throw new Error('Cannot encrypt: this environment has no secure source of random numbers (crypto.getRandomValues)');
  }
  return out;
}
