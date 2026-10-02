// Byte access for the Word binary reader.
//
// Every offset in a .doc file comes from the file itself, so nothing here
// trusts one: the readers return 0 for anything outside the buffer, and the
// structure readers check their ranges before they loop. A file that cannot
// be read at all is reported with an Error whose message starts with
// "Invalid doc:".

/** The file is a Word document this reader does not read (Word 6/95, encrypted). */
export class UnsupportedDocError extends Error {
  // The same name as the error open() throws for a format it does not know, so hosts need one check.
  override name = 'UnsupportedFileError';
}

export const invalid = (message: string): Error => new Error(`Invalid doc: ${message}`);

export const u8 = (b: Uint8Array, o: number): number => (o >= 0 && o < b.length ? b[o] : 0);

export const u16 = (b: Uint8Array, o: number): number => (o >= 0 && o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);

export const i16 = (b: Uint8Array, o: number): number => {
  const value = u16(b, o);
  return value >= 0x8000 ? value - 0x10000 : value;
};

export const u32 = (b: Uint8Array, o: number): number =>
  o >= 0 && o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16)) + b[o + 3] * 0x1000000 : 0;

export const i32 = (b: Uint8Array, o: number): number => (o >= 0 && o + 4 <= b.length ? b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24) : 0);

/** UTF-16LE text of `count` characters at `o`; cut short at the end of the buffer. */
export function utf16(b: Uint8Array, o: number, count: number): string {
  if (o < 0 || count <= 0) return '';
  const available = Math.floor((b.length - o) / 2);
  if (count > available) count = available;
  let out = '';
  // In chunks: String.fromCharCode takes its characters as arguments.
  const codes: number[] = [];
  for (let i = 0; i < count; i++) {
    codes.push(b[o + i * 2] | (b[o + i * 2 + 1] << 8));
    if (codes.length === 4096) {
      out += String.fromCharCode.apply(null, codes);
      codes.length = 0;
    }
  }
  if (codes.length) out += String.fromCharCode.apply(null, codes);
  return out;
}

/** The characters 0x80..0x9F of a compressed piece that are not the Unicode character of the same value (MS-DOC 2.9.73). */
const COMPRESSED: Record<number, number> = {
  0x82: 0x201a,
  0x83: 0x0192,
  0x84: 0x201e,
  0x85: 0x2026,
  0x86: 0x2020,
  0x87: 0x2021,
  0x88: 0x02c6,
  0x89: 0x2030,
  0x8a: 0x0160,
  0x8b: 0x2039,
  0x8c: 0x0152,
  0x91: 0x2018,
  0x92: 0x2019,
  0x93: 0x201c,
  0x94: 0x201d,
  0x95: 0x2022,
  0x96: 0x2013,
  0x97: 0x2014,
  0x98: 0x02dc,
  0x99: 0x2122,
  0x9a: 0x0161,
  0x9b: 0x203a,
  0x9c: 0x0153,
  0x9f: 0x0178,
};

const COMPRESSED_TABLE = (() => {
  const table = new Uint16Array(256);
  for (let i = 0; i < 256; i++) table[i] = COMPRESSED[i] === undefined ? i : COMPRESSED[i];
  return table;
})();

/** Text of a compressed piece: one byte a character. */
export function compressedText(b: Uint8Array, o: number, count: number): string {
  if (o < 0 || count <= 0) return '';
  if (count > b.length - o) count = b.length - o;
  let out = '';
  const codes: number[] = [];
  for (let i = 0; i < count; i++) {
    codes.push(COMPRESSED_TABLE[b[o + i]]);
    if (codes.length === 4096) {
      out += String.fromCharCode.apply(null, codes);
      codes.length = 0;
    }
  }
  if (codes.length) out += String.fromCharCode.apply(null, codes);
  return out;
}

/** Text of 8-bit characters (Windows-1252, as far as Latin-1 goes) up to `count` bytes or the first zero byte. */
export function ansiText(b: Uint8Array, o: number, count: number): string {
  let end = o;
  const limit = Math.min(b.length, o + Math.max(0, count));
  while (end < limit && b[end] !== 0) end++;
  return compressedText(b, o, end - o);
}

/** The bytes as a string of characters 0..255: a key for caches of what was read from them. */
export function bytesKey(b: Uint8Array, start: number, end: number): string {
  if (end > b.length) end = b.length;
  if (start >= end) return '';
  if (end - start <= 1024) return String.fromCharCode.apply(null, b.subarray(start, end) as unknown as number[]);
  let out = '';
  for (let i = start; i < end; i += 1024) out += String.fromCharCode.apply(null, b.subarray(i, Math.min(end, i + 1024)) as unknown as number[]);
  return out;
}

/** A PLC: n+1 character positions followed by n data elements of `size` bytes (MS-DOC 2.8.1). */
export interface Plc {
  count: number;
  /** n+1 positions. */
  cps: number[];
  /** Offset of data element i in the buffer the PLC was read from. */
  data(i: number): number;
}

/**
 * Reads a PLC at [fc, fc + lcb) of the table stream; an empty or broken one
 * has count 0. With `ordered`, a position that is lower than the one before
 * it is raised to that: the stretches between the positions never overlap, so
 * going through all of them never takes more than going through the text once.
 */
export function readPlc(b: Uint8Array, fc: number, lcb: number, size: number, ordered = false): Plc {
  const empty: Plc = { count: 0, cps: [], data: () => 0 };
  if (lcb < 4 || fc < 0 || fc + lcb > b.length) return empty;
  const count = Math.floor((lcb - 4) / (4 + size));
  if (count < 0) return empty;
  const cps: number[] = [];
  for (let i = 0; i <= count; i++) {
    let cp = u32(b, fc + i * 4);
    if (ordered && i > 0 && cp < cps[i - 1]) cp = cps[i - 1];
    cps.push(cp);
  }
  const base = fc + (count + 1) * 4;
  return { count, cps, data: (i) => base + i * size };
}

/**
 * Reads an STTB (MS-DOC 2.2.4): a table of strings, each with `cbExtra`
 * bytes of other data. Strings are 16-bit when the table starts with 0xFFFF,
 * 8-bit otherwise. `max` bounds how many entries are read.
 */
export function readSttb(b: Uint8Array, fc: number, lcb: number, max = 100000): { strings: string[]; extra: number[]; cbExtra: number } {
  const strings: string[] = [];
  const extra: number[] = [];
  const end = fc + lcb;
  if (lcb < 4 || fc < 0 || end > b.length) return { strings, extra, cbExtra: 0 };
  let p = fc;
  const wide = u16(b, p) === 0xffff;
  if (wide) p += 2;
  let count = u16(b, p);
  p += 2;
  const cbExtra = u16(b, p);
  p += 2;
  if (count > max) count = max;
  for (let i = 0; i < count && p < end; i++) {
    if (wide) {
      const cch = u16(b, p);
      p += 2;
      if (p + cch * 2 > end) break;
      strings.push(utf16(b, p, cch));
      p += cch * 2;
    } else {
      const cch = u8(b, p);
      p += 1;
      if (p + cch > end) break;
      strings.push(compressedText(b, p, cch));
      p += cch;
    }
    extra.push(p);
    p += cbExtra;
  }
  return { strings, extra, cbExtra };
}
