// The piece table (Clx, MS-DOC 2.9.38): where the characters of the document
// are in the WordDocument stream.
//
// The text of a document is a sequence of character positions (CPs). A piece
// maps a run of CPs to a place in the stream, either as 16-bit characters or
// "compressed" to one byte each. A file that was saved incrementally (fast
// save) has many pieces, in any order in the stream, and each may carry a
// property modifier (Prm) that applies to everything in it.
//
// The whole text is decoded once into one string; every other structure of
// the file refers to it by CP.

import { compressedText, invalid, u16, u32, u8, utf16 } from './bytes';

export interface Piece {
  cpStart: number;
  cpEnd: number;
  /** Offset of the first character in the WordDocument stream. */
  fc: number;
  /** One byte a character instead of two. */
  compressed: boolean;
  /** Property modifier as read; see PieceTable.modifier. */
  prm: number;
}

export interface PieceTable {
  pieces: Piece[];
  /** The text of all pieces: text.charCodeAt(cp) is the character at CP `cp`. */
  text: string;
  /** Index of the piece that holds the CP; -1 when no piece does. */
  pieceAt(cp: number): number;
  /** Offset in the WordDocument stream of the character at the CP; -1 when no piece holds it. */
  fcOf(cp: number): number;
  /** The properties a piece adds to its text as a list of sprms; undefined when it adds none. */
  modifier(piece: Piece): Uint8Array | undefined;
}

/** More pieces than this is not a genuine file. */
const MAX_PIECES = 2000000;

/** Sprm for each "isprm" of a Prm0 (MS-DOC 2.9.213); all take a one-byte operand. */
const PRM0: Record<number, number> = {
  0x00: 0x2879,
  0x04: 0x2602,
  0x05: 0x2461,
  0x07: 0x2405,
  0x08: 0x2406,
  0x09: 0x2407,
  0x0c: 0x260a,
  0x0d: 0x2470,
  0x0e: 0x240c,
  0x0f: 0x2471,
  0x18: 0x2416,
  0x19: 0x2417,
  0x1d: 0x261b,
  0x25: 0x2423,
  0x2c: 0x242a,
  0x32: 0x2430,
  0x33: 0x2431,
  0x35: 0x2433,
  0x36: 0x2434,
  0x37: 0x2435,
  0x38: 0x2436,
  0x39: 0x2437,
  0x3a: 0x2438,
  0x41: 0x0800,
  0x42: 0x0801,
  0x43: 0x0802,
  0x47: 0x0806,
  0x4b: 0x080a,
  0x4d: 0x2a0c,
  0x4e: 0x0858,
  0x4f: 0x2859,
  0x50: 0x0811,
  0x51: 0x0818,
  0x53: 0x2a33,
  0x55: 0x0835,
  0x56: 0x0836,
  0x57: 0x0837,
  0x58: 0x0838,
  0x59: 0x0839,
  0x5a: 0x083a,
  0x5b: 0x083b,
  0x5c: 0x083c,
  0x5e: 0x2a3e,
  0x62: 0x2a42,
  0x68: 0x2a44,
  0x6a: 0x2a46,
  0x6c: 0x2a48,
  0x73: 0x2a53,
  0x74: 0x0854,
  0x75: 0x0855,
  0x76: 0x0856,
  0x78: 0x2640,
  0x7b: 0x2a90,
  0x7c: 0x2a86,
  0x7e: 0x2443,
};

/**
 * Reads the Clx at [fc, fc + lcb) of the table stream and decodes the text
 * from the WordDocument stream. Throws when the table is not usable: without
 * it there is no text.
 */
export function readPieceTable(table: Uint8Array, fc: number, lcb: number, word: Uint8Array): PieceTable {
  if (lcb < 5 || fc < 0 || fc + lcb > table.length) throw invalid('the piece table lies outside the table stream');
  const end = fc + lcb;
  /** Ranges of the property lists (Prc) in the table stream, by index. */
  const lists: number[] = [];
  let p = fc;
  let plc = -1;
  let plcLength = 0;
  while (p < end) {
    const type = u8(table, p);
    if (type === 1) {
      const cb = u16(table, p + 1);
      if (p + 3 + cb > end) throw invalid('a property list of the piece table is cut short');
      lists.push(p + 3, p + 3 + cb);
      p += 3 + cb;
    } else if (type === 2) {
      plcLength = u32(table, p + 1);
      plc = p + 5;
      break;
    } else throw invalid('the piece table is corrupt');
  }
  if (plc < 0 || plcLength < 4 || plc + plcLength > end) throw invalid('the piece table has no pieces');
  const count = Math.floor((plcLength - 4) / 12);
  if (count > MAX_PIECES) throw invalid('the piece table has too many pieces');

  const pieces: Piece[] = [];
  const parts: string[] = [];
  let total = 0;
  let previous = u32(table, plc);
  if (previous !== 0) throw invalid('the piece table does not start at the first character');
  const data = plc + (count + 1) * 4;
  // Each character takes at least a byte of the stream; text that would be longer than twice the stream is not text of this file.
  const limit = word.length * 2 + 16;
  for (let i = 0; i < count; i++) {
    const cpEnd = u32(table, plc + (i + 1) * 4);
    if (cpEnd < previous) throw invalid('the pieces of the text are out of order');
    const raw = u32(table, data + i * 8 + 2);
    const compressed = (raw & 0x40000000) !== 0;
    const offset = compressed ? (raw & 0x3fffffff) >>> 1 : raw & 0x3fffffff;
    const chars = cpEnd - previous;
    if (chars > 0) {
      const bytes = compressed ? chars : chars * 2;
      if (offset + bytes > word.length) throw invalid('text lies outside the WordDocument stream');
      total += chars;
      if (total > limit) throw invalid('the pieces describe more text than the file holds');
      pieces.push({ cpStart: previous, cpEnd, fc: offset, compressed, prm: u16(table, data + i * 8 + 6) });
      parts.push(compressed ? compressedText(word, offset, chars) : utf16(word, offset, chars));
    }
    previous = cpEnd;
  }
  const text = parts.join('');

  const pieceAt = (cp: number): number => {
    let lo = 0;
    let hi = pieces.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const piece = pieces[mid];
      if (cp < piece.cpStart) hi = mid - 1;
      else if (cp >= piece.cpEnd) lo = mid + 1;
      else return mid;
    }
    return -1;
  };

  const modifiers = new Map<number, Uint8Array | undefined>();
  const modifier = (piece: Piece): Uint8Array | undefined => {
    const prm = piece.prm;
    if (prm === 0) return undefined;
    if (modifiers.has(prm)) return modifiers.get(prm);
    let out: Uint8Array | undefined;
    if (prm & 1) {
      const index = prm >> 1;
      if (index * 2 + 1 < lists.length) out = table.subarray(lists[index * 2], lists[index * 2 + 1]);
    } else {
      const sprm = PRM0[(prm >> 1) & 0x7f];
      if (sprm !== undefined) out = new Uint8Array([sprm & 0xff, sprm >> 8, prm >> 8]);
    }
    modifiers.set(prm, out);
    return out;
  };

  return {
    pieces,
    text,
    pieceAt,
    fcOf(cp: number): number {
      const index = pieceAt(cp);
      if (index < 0) return -1;
      const piece = pieces[index];
      return piece.fc + (cp - piece.cpStart) * (piece.compressed ? 1 : 2);
    },
    modifier,
  };
}
