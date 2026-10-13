// A TrueType font cut down to the glyphs a file uses. The glyphs keep their
// numbers: the ones that are not used are still counted, with no outline, so
// that the codes of the text (which are glyph numbers) mean the same in the
// subset as in the font, and a page can be written before it is known which
// glyphs the pages after it use.
//
// What stays: the outlines of the used glyphs as they were, byte for byte
// (their instructions too), the glyphs they are made of, the advances, and
// the programs the instructions call (cvt, fpgm, prep). What a PDF does not
// read of an embedded font (the character map, names, kerning, layout
// tables) is left out.

import { FontError, fontTable, type PdfFontFace } from './sfnt';

/** How many glyphs deep a glyph may be made of other glyphs. */
const MAX_PARTS_PER_GLYPH = 4096;
/** Bytes one glyph may have; a larger one is none that is written. */
export const MAX_GLYPH_BYTES = 256 * 1024;

const fail = (message: string): never => {
  throw new FontError(message);
};

interface GlyphPlaces {
  /** Start and end of a glyph in the glyf table. */
  range(glyph: number): [number, number];
  /** The glyphs a glyph is made of; none for a glyph with an outline of its own. */
  parts(glyph: number): number[];
  bytes(glyph: number): Uint8Array;
}

function glyphPlaces(face: PdfFontFace): GlyphPlaces {
  if (face.kind !== 'truetype') fail('Only a TrueType font can be cut down to its glyphs');
  const loca = fontTable(face, 'loca');
  const glyf = fontTable(face, 'glyf');
  const range = (glyph: number): [number, number] => {
    const from = face.longLoca ? loca.u32(glyph * 4) : loca.u16(glyph * 2) * 2;
    const to = face.longLoca ? loca.u32(glyph * 4 + 4) : loca.u16(glyph * 2 + 2) * 2;
    if (from > to || to > glyf.length) fail(`Glyph ${glyph} of the font lies outside its glyf table`);
    if (to - from > MAX_GLYPH_BYTES) fail(`Glyph ${glyph} of the font is larger than a glyph may be`);
    return [from, to];
  };
  /** Reads the outline of a glyph with contours of its own to its end: throws when the glyph has no room for what it says it holds. */
  const outline = (glyph: number, from: number, to: number): void => {
    const short = (): never => fail(`The outline of glyph ${glyph} of the font does not fit the glyph`);
    const contours = glyf.i16(from);
    // A glyph that says it has no contours has nothing more to say.
    if (contours === 0) return;
    let at = from + 10 + contours * 2;
    if (at + 2 > to) short();
    let points = 0;
    for (let i = 0; i < contours; i++) {
      const end = glyf.u16(from + 10 + i * 2);
      if (end + 1 <= points && i > 0) short();
      points = end + 1;
    }
    at += 2 + glyf.u16(at);
    // The flags (one may stand for many points), then the coordinates they say there are.
    let coordinates = 0;
    for (let seen = 0; seen < points; ) {
      if (at >= to) short();
      const flag = glyf.u8(at++);
      let times = 1;
      if ((flag & 0x08) !== 0) {
        if (at >= to) short();
        times += glyf.u8(at++);
      }
      if (seen + times > points) short();
      seen += times;
      coordinates += times * (((flag & 0x02) !== 0 ? 1 : (flag & 0x10) !== 0 ? 0 : 2) + ((flag & 0x04) !== 0 ? 1 : (flag & 0x20) !== 0 ? 0 : 2));
    }
    if (at + coordinates > to) short();
  };
  return {
    range,
    parts: (glyph) => {
      const [from, to] = range(glyph);
      // An empty glyph, or one with contours of its own.
      if (to === from) return [];
      if (to - from < 10) fail(`Glyph ${glyph} of the font is shorter than its header`);
      if (glyf.i16(from) >= 0) {
        outline(glyph, from, to);
        return [];
      }
      const out: number[] = [];
      let at = from + 10;
      for (let i = 0; ; i++) {
        if (i >= MAX_PARTS_PER_GLYPH) fail(`Glyph ${glyph} of the font is made of too many parts`);
        if (at + 4 > to) fail(`A part of glyph ${glyph} of the font lies outside the glyph`);
        const flags = glyf.u16(at);
        const part = glyf.u16(at + 2);
        if (part >= face.glyphCount) fail(`Glyph ${glyph} of the font names a part the font does not have`);
        out.push(part);
        // The arguments (bytes or words), then a scale, two scales or a matrix.
        at += 4 + ((flags & 0x0001) !== 0 ? 4 : 2) + ((flags & 0x0008) !== 0 ? 2 : (flags & 0x0040) !== 0 ? 4 : (flags & 0x0080) !== 0 ? 8 : 0);
        if (at > to) fail(`A part of glyph ${glyph} of the font lies outside the glyph`);
        if ((flags & 0x0020) === 0) break;
      }
      return out;
    },
    bytes: (glyph) => {
      const [from, to] = range(glyph);
      return face.bytes.subarray(glyf.start + from, glyf.start + to);
    },
  };
}

function closure(face: PdfFontFace, places: GlyphPlaces, used: Iterable<number>): number[] {
  const kept = new Set<number>([0]);
  const waiting: number[] = [];
  const add = (glyph: number): void => {
    if (typeof glyph !== 'number' || !(glyph >= 0) || Math.floor(glyph) !== glyph || glyph >= face.glyphCount) fail(`The font has no glyph ${String(glyph)}`);
    if (kept.has(glyph)) return;
    kept.add(glyph);
    waiting.push(glyph);
  };
  waiting.push(0);
  const all = Array.from(used);
  for (let i = 0; i < all.length; i++) add(all[i]);
  // Every glyph is looked at once: glyphs that are made of each other in a circle end here.
  while (waiting.length > 0) {
    const parts = places.parts(waiting.pop() as number);
    for (let i = 0; i < parts.length; i++) add(parts[i]);
  }
  return Array.from(kept).sort((a, b) => a - b);
}

/**
 * The glyphs a subset holds for the glyphs that are used: those, the glyphs
 * they are made of, and glyph 0 (what stands for a missing character), in
 * rising order. Throws a FontError for a number that is no glyph of the font
 * and for a glyph that does not lie in the font as it must.
 */
export function glyphClosure(face: PdfFontFace, used: Iterable<number>): number[] {
  return closure(face, glyphPlaces(face), used);
}

function checksum(bytes: Uint8Array, from: number, length: number): number {
  let sum = 0;
  const end = from + length;
  for (let i = from; i < end; i += 4) {
    sum = (sum + bytes[i] * 16777216 + ((i + 1 < end ? bytes[i + 1] << 16 : 0) | (i + 2 < end ? bytes[i + 2] << 8 : 0) | (i + 3 < end ? bytes[i + 3] : 0))) >>> 0;
  }
  return sum;
}

function put32(bytes: Uint8Array, at: number, value: number): void {
  bytes[at] = (value >>> 24) & 255;
  bytes[at + 1] = (value >>> 16) & 255;
  bytes[at + 2] = (value >>> 8) & 255;
  bytes[at + 3] = value & 255;
}

/**
 * The font with the outlines of the used glyphs only (see glyphClosure): a
 * TrueType file of the tables a PDF reads. Throws a FontError for a font that
 * has no TrueType outlines and for glyphs that are not as they must be.
 */
export function subsetTrueType(face: PdfFontFace, used: Iterable<number>): Uint8Array {
  const places = glyphPlaces(face);
  const kept = closure(face, places, used);
  const count = face.glyphCount;

  // The glyphs one after the other, each at a multiple of four; loca in its long form.
  let size = 0;
  const data: Uint8Array[] = [];
  for (let i = 0; i < kept.length; i++) {
    const bytes = places.bytes(kept[i]);
    data.push(bytes);
    size += (bytes.length + 3) & ~3;
  }
  // Glyphs that follow each other in their font are in it once: a subset that is larger than the glyphs of its
  // font (and the padding of each) is made of a font that lies about where its glyphs are.
  if (size > fontTable(face, 'glyf').length + kept.length * 4) fail('The glyphs of the subset are larger than the glyphs of the font');
  const glyf = new Uint8Array(size);
  const loca = new Uint8Array((count + 1) * 4);
  let at = 0;
  let next = 0;
  for (let glyph = 0; glyph <= count; glyph++) {
    put32(loca, glyph * 4, at);
    if (next < kept.length && kept[next] === glyph) {
      glyf.set(data[next], at);
      at += (data[next].length + 3) & ~3;
      next++;
    }
  }

  const copy = (tag: string, length?: number, pad = false): Uint8Array => {
    const table = fontTable(face, tag);
    if (length === undefined) return face.bytes.slice(table.start, table.end);
    if (table.length < length && !pad) fail(`The ${tag} table of the font is cut short`);
    // What a table lacks at its end is nothing: zeros.
    const out = new Uint8Array(length);
    out.set(face.bytes.subarray(table.start, table.start + Math.min(length, table.length)));
    return out;
  };
  const head = copy('head', 54);
  put32(head, 8, 0);
  // indexToLocFormat: long.
  head[50] = 0;
  head[51] = 1;
  const hhea = copy('hhea', 36);
  const metrics = (hhea[34] << 8) | hhea[35];
  const tables: { [tag: string]: Uint8Array } = { glyf, head, hhea, hmtx: copy('hmtx', metrics * 4 + (count - metrics) * 2, true), loca, maxp: copy('maxp', 32) };
  for (const tag of ['cvt ', 'fpgm', 'prep']) if (face.tables.has(tag)) tables[tag] = copy(tag);

  const tags = Object.keys(tables).sort();
  let power = 1;
  let selector = 0;
  while (power * 2 <= tags.length) {
    power *= 2;
    selector++;
  }
  let total = 12 + tags.length * 16;
  for (let i = 0; i < tags.length; i++) total += (tables[tags[i]].length + 3) & ~3;
  const out = new Uint8Array(total);
  put32(out, 0, 0x00010000);
  out[4] = tags.length >> 8;
  out[5] = tags.length & 255;
  out[6] = (power * 16) >> 8;
  out[7] = (power * 16) & 255;
  out[8] = selector >> 8;
  out[9] = selector & 255;
  out[10] = (tags.length * 16 - power * 16) >> 8;
  out[11] = (tags.length * 16 - power * 16) & 255;
  let offset = 12 + tags.length * 16;
  let headAt = 0;
  for (let i = 0; i < tags.length; i++) {
    const tag = tags[i];
    const table = tables[tag];
    const entry = 12 + i * 16;
    for (let k = 0; k < 4; k++) out[entry + k] = tag.charCodeAt(k);
    out.set(table, offset);
    put32(out, entry + 4, checksum(out, offset, table.length));
    put32(out, entry + 8, offset);
    put32(out, entry + 12, table.length);
    if (tag === 'head') headAt = offset;
    offset += (table.length + 3) & ~3;
  }
  // The sum of the whole file comes to a fixed number through this field.
  put32(out, headAt + 8, (0xb1b0afba - checksum(out, 0, out.length)) >>> 0);
  return out;
}

const usable = new WeakMap<PdfFontFace, Map<number, boolean>>();

/**
 * Whether a glyph of a font can be put into a file: the font has it, its
 * outline lies inside the font and is read to its end, and so is every glyph
 * it is made of. Asked before the glyph is written on a page, so that a
 * broken glyph leaves that page a picture instead of being found when the
 * file is ended. The outlines of a font with CFF outlines are not read: the
 * font goes in whole, and a glyph of it is one the font has.
 */
export function glyphUsable(face: PdfFontFace, glyph: number): boolean {
  if (typeof glyph !== 'number' || !(glyph >= 0) || Math.floor(glyph) !== glyph || glyph >= face.glyphCount) return false;
  if (face.kind !== 'truetype') return true;
  let known = usable.get(face);
  if (!known) {
    known = new Map();
    usable.set(face, known);
  }
  let sound = known.get(glyph);
  if (sound === undefined) {
    try {
      closure(face, glyphPlaces(face), [glyph]);
      sound = true;
    } catch (error) {
      if (!(error instanceof FontError)) throw error;
      sound = false;
    }
    known.set(glyph, sound);
  }
  return sound;
}
