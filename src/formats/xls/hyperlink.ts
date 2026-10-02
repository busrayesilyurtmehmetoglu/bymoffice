// HLINK record -> hyperlink of the model.
//
// The record wraps a "hyperlink object": optional display text, a target (a URL
// or a file path, each in its own moniker format) and a location inside the target.

import type { Hyperlink } from '../../model/sheet/types';
import { decodeBytes } from './codepage';
import type { RecordReader } from './records';

const HAS_MONIKER = 0x001;
const HAS_LOCATION = 0x008;
const HAS_DISPLAY_NAME = 0x010;
const HAS_FRAME_NAME = 0x080;
const MONIKER_AS_STRING = 0x100;

/** Longest text taken from a link; a URL or a path is far shorter. */
const MAX_LINK_CHARS = 32768;

/** UTF-16 text up to the first NUL. */
function utf16(bytes: Uint8Array): string {
  let out = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) {
    const code = bytes[i] | (bytes[i + 1] << 8);
    if (code === 0) break;
    out += String.fromCharCode(code);
  }
  return out;
}

/** A string of the hyperlink object: a character count, then UTF-16 with a closing NUL. */
function linkString(r: RecordReader): string {
  const count = r.u32();
  if (count > MAX_LINK_CHARS) throw new Error('Invalid xls: a hyperlink string is too long');
  return utf16(r.take(count * 2));
}

/** Reads the target out of a moniker; undefined for a kind this reader does not know (the rest of the record is then unreadable). */
function readMoniker(r: RecordReader): string | undefined {
  const classId = r.take(16);
  if (classId[0] === 0xe0 && classId[1] === 0xc9 && classId[2] === 0xea && classId[3] === 0x79) {
    // URL moniker: the text may be followed by a GUID and flags inside the same length.
    const size = r.u32();
    if (size > MAX_LINK_CHARS * 2) throw new Error('Invalid xls: a hyperlink target is too long');
    return utf16(r.take(size));
  }
  if (classId[0] === 0x03 && classId[1] === 0x03 && classId[2] === 0x00 && classId[3] === 0x00) {
    // File moniker: a path in the ANSI code page, and often the same path in UTF-16.
    const up = r.u16();
    const ansiSize = r.u32();
    if (ansiSize > MAX_LINK_CHARS) throw new Error('Invalid xls: a hyperlink target is too long');
    const ansi = r.take(ansiSize);
    let end = 0;
    while (end < ansi.length && ansi[end] !== 0) end++;
    let path = decodeBytes(ansi, 0, end, 1252);
    r.skip(24);
    const unicodeSize = r.u32();
    if (unicodeSize > 0) {
      const size = r.u32();
      r.skip(2);
      if (size > MAX_LINK_CHARS * 2) throw new Error('Invalid xls: a hyperlink target is too long');
      path = utf16(r.take(size));
    }
    let prefix = '';
    for (let i = 0; i < up && i < 64; i++) prefix += '../';
    return prefix + path.replace(/\\/g, '/');
  }
  return undefined;
}

/** Reads an HLINK record. A link whose range is outside the sheet is left out (undefined). */
export function readHyperlink(r: RecordReader): Hyperlink | undefined {
  const top = r.u16();
  const bottom = r.u16();
  const left = r.u16();
  const right = r.u16();
  r.skip(16);
  r.skip(4);
  const flags = r.u32();
  if (top > bottom || left > right || right > 255) return undefined;
  const link: Hyperlink = { range: { top, left, bottom, right } };
  if (flags & HAS_DISPLAY_NAME) {
    const display = linkString(r);
    if (display !== '') link.display = display;
  }
  if (flags & HAS_FRAME_NAME) linkString(r);
  if (flags & HAS_MONIKER) {
    const target = flags & MONIKER_AS_STRING ? linkString(r) : readMoniker(r);
    // After a moniker of an unknown kind there is no telling where the location starts.
    if (target === undefined) return link;
    if (target !== '') link.target = target;
  }
  if (flags & HAS_LOCATION) {
    const location = linkString(r);
    if (location !== '') link.location = location;
  }
  return link;
}

/** Reads an HLINKTOOLTIP record: the range of the link it belongs to, and the text. */
export function readHyperlinkTooltip(r: RecordReader): { top: number; left: number; bottom: number; right: number; text: string } {
  r.skip(2);
  const top = r.u16();
  const bottom = r.u16();
  const left = r.u16();
  const right = r.u16();
  return { top, left, bottom, right, text: utf16(r.take(Math.min(r.remaining, MAX_LINK_CHARS * 2))) };
}
