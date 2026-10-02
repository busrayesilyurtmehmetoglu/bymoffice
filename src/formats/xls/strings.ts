// Strings of BIFF records and the shared string table (SST).
//
// BIFF8 strings are UTF-16, stored with one byte per character when every
// character is below U+0100. BIFF5 strings are bytes in the workbook's code page.

import type { Font, RichRun, SharedString } from '../../model/sheet/types';
import { decodeBytes } from './codepage';
import type { RecordReader } from './records';

/** A string as the file has it: formatting runs are (first character, font index) pairs. */
export interface RawString {
  text: string;
  /** ich0, ifnt0, ich1, ifnt1, ... */
  runs?: number[];
}

const FLAG_WIDE = 0x01;
const FLAG_EXT = 0x04;
const FLAG_RICH = 0x08;

/**
 * The part of a BIFF8 string after its character count: option flags, the characters, and the
 * formatting runs and phonetic data that the flags announce. Phonetic data is skipped.
 */
export function readStringBody(r: RecordReader, count: number): RawString {
  const flags = r.u8();
  const runCount = flags & FLAG_RICH ? r.u16() : 0;
  const extSize = flags & FLAG_EXT ? r.u32() : 0;
  const result: RawString = { text: r.chars(count, (flags & FLAG_WIDE) !== 0) };
  if (runCount > 0) {
    const runs: number[] = [];
    for (let i = 0; i < runCount; i++) runs.push(r.u16(), r.u16());
    result.runs = runs;
  }
  // Some writers state more phonetic data than they wrote; what is there is skipped.
  if (extSize > 0) r.skipUpTo(extSize);
  return result;
}

/** A BIFF8 string with a 16-bit character count (XLUnicodeString and its rich variant). */
export function readString16(r: RecordReader): RawString {
  return readStringBody(r, r.u16());
}

/** A BIFF8 string with an 8-bit character count (ShortXLUnicodeString). */
export function readString8(r: RecordReader): string {
  return readStringBody(r, r.u8()).text;
}

/** A BIFF5 byte string of `count` bytes. */
export function readByteString(r: RecordReader, count: number, codepage: number): string {
  const bytes = r.take(count);
  return decodeBytes(bytes, 0, bytes.length, codepage);
}

/** A string with a 16-bit length, in the form the BIFF version uses. */
export function readText16(r: RecordReader, biff8: boolean, codepage: number): RawString {
  if (biff8) return readString16(r);
  return { text: readByteString(r, r.u16(), codepage) };
}

/** A string with an 8-bit length, in the form the BIFF version uses. */
export function readText8(r: RecordReader, biff8: boolean, codepage: number): string {
  if (biff8) return readString8(r);
  return readByteString(r, r.u8(), codepage);
}

/**
 * The strings of an SST record. The count in the record is not trusted: reading stops where
 * the data ends, and nothing is allocated from the count.
 */
export function readSst(r: RecordReader): RawString[] {
  r.skip(4);
  const unique = r.u32();
  const strings: RawString[] = [];
  for (let i = 0; i < unique && !r.done; i++) strings.push(readString16(r));
  return strings;
}

/**
 * Formatting runs as the model has them. The text before the first run has the font of the
 * cell; `fontOf` answers the font for a font index of the file.
 */
export function toSharedString(raw: RawString, fontOf: (index: number) => Font | undefined): SharedString {
  const text = raw.text;
  const pairs = raw.runs;
  if (!pairs || pairs.length === 0 || text.length === 0) return { text };
  const runs: RichRun[] = [];
  let at = 0;
  let font: Font | undefined;
  for (let i = 0; i < pairs.length; i += 2) {
    // Runs are in ascending order; one that is not, or starts past the text, changes nothing before it.
    const start = Math.min(Math.max(pairs[i], at), text.length);
    if (start > at) runs.push(font ? { text: text.slice(at, start), font } : { text: text.slice(at, start) });
    at = start;
    font = fontOf(pairs[i + 1]);
  }
  if (at < text.length) runs.push(font ? { text: text.slice(at), font } : { text: text.slice(at) });
  // One run without a font is plain text.
  if (runs.length === 0 || (runs.length === 1 && !runs[0].font)) return { text };
  return { text, runs };
}
