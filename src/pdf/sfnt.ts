// Reading a font file (TrueType with glyf outlines, or OpenType with a CFF
// table) as far as a PDF needs it: which glyph a character is, how wide the
// glyphs are, the numbers of the font descriptor, and whether the font may be
// embedded at all.
//
// A font file comes from an application or from the network and is not
// trusted: every offset and every length is checked against the table it
// lies in before anything is read, every loop is bounded by the size of what
// it walks, and nothing is copied or expanded by a count the file states. A
// file that is not as it must be throws a FontError; the page that needed the
// font is then written as a picture.
//
// There is no shaping here: no kerning, no joining of letters. Of the rules
// by which a font replaces glyphs only the ligatures of Latin text are read
// (fontLigatures), so that text can be written with the ligatures a browser
// draws; a font with rules that are not read says so.

/** Why a font file cannot be used. */
export class FontError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'FontError';
  }
}

/** A font file larger than this is not read. */
export const MAX_FONT_BYTES = 32 * 1024 * 1024;
const MAX_TABLES = 256;
/** Groups of a format 12 character map that are looked at; a font with more is refused. */
const MAX_CMAP_GROUPS = 200000;

export interface FontTable {
  offset: number;
  length: number;
}

const fail = (message: string): never => {
  throw new FontError(message);
};

/** Big endian numbers of a table, every read checked against the end of the table. */
export class FontReader {
  readonly end: number;
  constructor(
    readonly bytes: Uint8Array,
    readonly start: number,
    length: number,
    readonly what: string,
  ) {
    if (!(start >= 0) || !(length >= 0) || start + length > bytes.length) fail(`The ${what} table lies outside the font file`);
    this.end = start + length;
  }
  get length(): number {
    return this.end - this.start;
  }
  private at(offset: number, size: number): number {
    if (!(offset >= 0) || Math.floor(offset) !== offset || this.start + offset + size > this.end) fail(`The ${this.what} table is read beyond its end`);
    return this.start + offset;
  }
  u8(offset: number): number {
    return this.bytes[this.at(offset, 1)];
  }
  u16(offset: number): number {
    const at = this.at(offset, 2);
    return (this.bytes[at] << 8) | this.bytes[at + 1];
  }
  i16(offset: number): number {
    const value = this.u16(offset);
    return value >= 0x8000 ? value - 0x10000 : value;
  }
  u32(offset: number): number {
    const at = this.at(offset, 4);
    return this.bytes[at] * 16777216 + ((this.bytes[at + 1] << 16) | (this.bytes[at + 2] << 8) | this.bytes[at + 3]);
  }
  i32(offset: number): number {
    const value = this.u32(offset);
    return value >= 0x80000000 ? value - 0x100000000 : value;
  }
}

/** The character map of a font: segments (format 4) or groups (format 12), searched, never expanded. */
interface CharacterMap {
  glyph(codePoint: number): number;
}

function format4(table: FontReader, at: number): CharacterMap {
  const length = table.u16(at + 2);
  const segments = table.u16(at + 6) >> 1;
  if (segments < 1 || 16 + segments * 8 > length || at + length > table.length) fail('The character map of the font does not fit its table');
  const ends = at + 14;
  const starts = ends + segments * 2 + 2;
  const deltas = starts + segments * 2;
  const offsets = deltas + segments * 2;
  const limit = at + length;
  return {
    glyph: (codePoint) => {
      if (codePoint > 0xffff) return 0;
      // The first segment that ends at or after the character.
      let low = 0;
      let high = segments - 1;
      while (low < high) {
        const middle = (low + high) >> 1;
        if (table.u16(ends + middle * 2) < codePoint) low = middle + 1;
        else high = middle;
      }
      if (table.u16(ends + low * 2) < codePoint) return 0;
      const start = table.u16(starts + low * 2);
      if (start > codePoint) return 0;
      const delta = table.u16(deltas + low * 2);
      const rangeOffset = table.u16(offsets + low * 2);
      if (rangeOffset === 0) return (codePoint + delta) & 0xffff;
      const place = offsets + low * 2 + rangeOffset + (codePoint - start) * 2;
      if (place + 2 > limit) return 0;
      const glyph = table.u16(place);
      return glyph === 0 ? 0 : (glyph + delta) & 0xffff;
    },
  };
}

function format12(table: FontReader, at: number): CharacterMap {
  const length = table.u32(at + 4);
  const groups = table.u32(at + 12);
  if (groups > MAX_CMAP_GROUPS || 16 + groups * 12 > length || at + length > table.length) fail('The character map of the font does not fit its table');
  const first = at + 16;
  return {
    glyph: (codePoint) => {
      let low = 0;
      let high = groups - 1;
      while (low <= high) {
        const middle = (low + high) >> 1;
        const start = table.u32(first + middle * 12);
        const end = table.u32(first + middle * 12 + 4);
        if (codePoint < start) high = middle - 1;
        else if (codePoint > end) low = middle + 1;
        else {
          const glyph = table.u32(first + middle * 12 + 8) + (codePoint - start);
          return glyph > 0xffff ? 0 : glyph;
        }
      }
      return 0;
    },
  };
}

function characterMap(table: FontReader): CharacterMap {
  const count = table.u16(2);
  if (4 + count * 8 > table.length) fail('The character map of the font names more tables than it holds');
  // The best table: all of Unicode before the BMP alone; a map of symbols (3, 0) is none of characters.
  let best = -1;
  let bestRank = 0;
  for (let i = 0; i < count; i++) {
    const platform = table.u16(4 + i * 8);
    const encoding = table.u16(6 + i * 8);
    const offset = table.u32(8 + i * 8);
    if (offset + 4 > table.length) continue;
    const format = table.u16(offset);
    let rank = 0;
    if (format === 12 && ((platform === 3 && encoding === 10) || (platform === 0 && (encoding === 4 || encoding === 6)))) rank = platform === 3 ? 4 : 3;
    else if (format === 4 && ((platform === 3 && encoding === 1) || platform === 0)) rank = platform === 3 ? 2 : 1;
    if (rank > bestRank) {
      bestRank = rank;
      best = offset;
    }
  }
  if (best < 0) return fail('The font has no character map for Unicode');
  return bestRank >= 3 ? format12(table, best) : format4(table, best);
}

/** A name of the font as a PostScript name may be: printable ASCII without what means something in a PDF, 63 characters at most. */
function cleanName(text: string): string {
  let out = '';
  for (let i = 0; i < text.length && out.length < 63; i++) {
    const code = text.charCodeAt(i);
    if (code > 0x20 && code < 0x7f && '[](){}<>/%#'.indexOf(text.charAt(i)) < 0) out += text.charAt(i);
  }
  return out;
}

function postScriptName(table: FontReader | undefined): string {
  if (!table || table.length < 6) return '';
  const count = table.u16(2);
  const strings = table.u16(4);
  if (6 + count * 12 > table.length) return '';
  let found = '';
  for (let i = 0; i < count && i < 4096; i++) {
    const at = 6 + i * 12;
    if (table.u16(at + 6) !== 6) continue;
    const platform = table.u16(at);
    const length = table.u16(at + 8);
    const offset = strings + table.u16(at + 10);
    if (offset + length > table.length || length > 512) continue;
    let text = '';
    if (platform === 3 || platform === 0) for (let k = 0; k + 1 < length; k += 2) text += String.fromCharCode(table.u16(offset + k));
    else if (platform === 1) for (let k = 0; k < length; k++) text += String.fromCharCode(table.u8(offset + k));
    else continue;
    const clean = cleanName(text);
    if (clean !== '') {
      found = clean;
      // The Windows name is the one programs agree on.
      if (platform === 3) break;
    }
  }
  return found;
}

/**
 * Looks at the CFF table as far as needed to know that its glyphs are what
 * the rest of the file says: version 1, one font, not keyed by CIDs (the
 * codes of the text are then glyph numbers), with as many charstrings as the
 * font has glyphs.
 */
function checkCff(table: FontReader, glyphs: number): void {
  if (table.length < 4 || table.u8(0) !== 1) fail('The CFF table of the font is not of version 1');
  const header = table.u8(2);
  // An INDEX: count, offset size, count + 1 offsets (from 1), the data.
  const index = (at: number): { count: number; data: number; offset(i: number): number; next: number } => {
    const count = table.u16(at);
    if (count === 0) return { count, data: at + 2, offset: () => 0, next: at + 2 };
    const size = table.u8(at + 2);
    if (size < 1 || size > 4) fail('An index of the CFF table of the font has no offset size');
    const offsets = at + 3;
    const data = offsets + (count + 1) * size - 1;
    const offset = (i: number): number => {
      let value = 0;
      for (let b = 0; b < size; b++) value = value * 256 + table.u8(offsets + i * size + b);
      return value;
    };
    const last = offset(count);
    if (last < 1 || data + last > table.length) fail('An index of the CFF table of the font does not fit');
    return { count, data, offset, next: data + last };
  };
  const names = index(header);
  if (names.count !== 1) fail('The CFF table of the font holds more than one font');
  const top = index(names.next);
  if (top.count !== 1) fail('The CFF table of the font has no top dictionary');
  const from = top.data + top.offset(0);
  const to = top.data + top.offset(1);
  if (!(from <= to) || to > table.length) fail('The top dictionary of the CFF table does not fit');
  // The operators of the dictionary, each after its operands.
  let charStrings = -1;
  let last = 0;
  for (let at = from, steps = 0; at < to; steps++) {
    if (steps > 4096) fail('The top dictionary of the CFF table is too long');
    const b0 = table.u8(at);
    if (b0 >= 32 && b0 <= 246) {
      last = b0 - 139;
      at += 1;
    } else if (b0 >= 247 && b0 <= 254) {
      const b1 = table.u8(at + 1);
      last = b0 <= 250 ? (b0 - 247) * 256 + b1 + 108 : -(b0 - 251) * 256 - b1 - 108;
      at += 2;
    } else if (b0 === 28) {
      last = table.i16(at + 1);
      at += 3;
    } else if (b0 === 29) {
      last = table.i32(at + 1);
      at += 5;
    } else if (b0 === 30) {
      // A real number: nibbles up to the one that ends it.
      at += 1;
      for (let k = 0; ; k++) {
        if (k > 64) fail('A number of the CFF table does not end');
        const pair = table.u8(at++);
        if ((pair & 15) === 15 || pair >> 4 === 15) break;
      }
    } else if (b0 === 12) {
      // ROS: the font is keyed by CIDs.
      if (table.u8(at + 1) === 30) fail('The CFF font is keyed by CIDs, which is not read here');
      at += 2;
    } else if (b0 <= 21) {
      if (b0 === 17) charStrings = last;
      at += 1;
    } else fail('The top dictionary of the CFF table holds a byte that cannot be');
  }
  if (!(charStrings > 0) || charStrings >= table.length) fail('The CFF table of the font has no glyphs');
  if (index(charStrings).count !== glyphs) fail('The CFF table of the font has another number of glyphs than the font says');
}

/** A ligature: `glyph` stands for the first glyph of its set followed by `rest`. */
export interface LigatureRule {
  glyph: number;
  rest: number[];
}

/** One lookup of ligatures: by their first glyph, the ligatures in the order the font tries them. */
export type LigatureLookup = Map<number, LigatureRule[]>;

/** What a font says of the ligatures of Latin text. */
export interface FontLigatures {
  /**
   * 'none': the font replaces no glyphs of Latin text by default. 'known':
   * `variants` are the sets of ligature lookups it has, one per language it
   * tells apart (which of them a browser uses depends on the language of the
   * text). 'unknown': it replaces glyphs by rules that are not read here.
   */
  state: 'none' | 'known' | 'unknown';
  variants: LigatureLookup[][];
  /** The font has rules for other scripts than Latin, which are not read. */
  otherScripts: boolean;
}

const MAX_GSUB_ITEMS = 512;
const MAX_LIGATURES = 20000;
const MAX_LIGATURE_GLYPHS = 16;

/** The glyphs a coverage table lists, in its order. */
function coverageGlyphs(table: FontReader, at: number, glyphCount: number): number[] {
  const format = table.u16(at);
  const count = table.u16(at + 2);
  const out: number[] = [];
  if (format === 1) {
    for (let i = 0; i < count; i++) out.push(table.u16(at + 4 + i * 2));
  } else if (format === 2) {
    for (let i = 0; i < count; i++) {
      const start = table.u16(at + 4 + i * 6);
      const end = table.u16(at + 6 + i * 6);
      if (end < start || out.length + (end - start + 1) > glyphCount) fail('A coverage table of the font lists more glyphs than the font has');
      for (let glyph = start; glyph <= end; glyph++) out.push(glyph);
    }
  } else fail('A coverage table of the font is of an unknown format');
  return out;
}

function readLigatures(face: PdfFontFace): FontLigatures {
  const place = face.tables.get('GSUB');
  if (!place) return { state: 'none', variants: [], otherScripts: false };
  const table = new FontReader(face.bytes, place.offset, place.length, 'GSUB');
  if (table.u16(0) !== 1) fail('The GSUB table of the font is of an unknown version');
  const scriptList = table.u16(4);
  const featureList = table.u16(6);
  const lookupList = table.u16(8);
  const featureCount = table.u16(featureList);
  const lookupCount = table.u16(lookupList);
  /** The lookups of a language system that are applied to plain text; undefined when it has rules that are not read. */
  const lookupsOf = (langSys: number): number[] | undefined => {
    const required = table.u16(langSys + 2);
    const count = table.u16(langSys + 4);
    if (count > MAX_GSUB_ITEMS) fail('A language system of the font names too many features');
    const features: { index: number; required: boolean }[] = [];
    if (required !== 0xffff) features.push({ index: required, required: true });
    for (let i = 0; i < count; i++) features.push({ index: table.u16(langSys + 6 + i * 2), required: false });
    const lookups: number[] = [];
    for (let i = 0; i < features.length; i++) {
      const index = features[i].index;
      if (index >= featureCount) fail('A language system of the font names a feature the font does not have');
      const record = featureList + 2 + index * 6;
      const tag = String.fromCharCode(table.u8(record), table.u8(record + 1), table.u8(record + 2), table.u8(record + 3));
      const feature = featureList + table.u16(record + 4);
      const named = table.u16(feature + 2);
      if (named > MAX_GSUB_ITEMS) fail('A feature of the font names too many lookups');
      const ligatures = tag === 'liga' || tag === 'clig' || tag === 'rlig';
      // Alternates by context are on by default as ligatures are, and a feature a language requires is always on.
      if (!ligatures && (tag === 'calt' || features[i].required) && named > 0) return undefined;
      if (!ligatures) continue;
      for (let k = 0; k < named; k++) {
        const lookup = table.u16(feature + 4 + k * 2);
        if (lookup >= lookupCount) fail('A feature of the font names a lookup the font does not have');
        if (lookups.indexOf(lookup) < 0) lookups.push(lookup);
      }
    }
    // Lookups are applied in the order of their list, whatever order the features name them in.
    return lookups.sort((a, b) => a - b);
  };
  const scripts = table.u16(scriptList);
  if (scripts > MAX_GSUB_ITEMS) fail('The font has too many scripts');
  let otherScripts = false;
  const sets: number[][] = [];
  for (let i = 0; i < scripts; i++) {
    const record = scriptList + 2 + i * 6;
    const tag = String.fromCharCode(table.u8(record), table.u8(record + 1), table.u8(record + 2), table.u8(record + 3));
    if (tag !== 'DFLT' && tag !== 'latn') {
      otherScripts = true;
      continue;
    }
    const script = scriptList + table.u16(record + 4);
    const languages = table.u16(script + 2);
    if (languages > MAX_GSUB_ITEMS) fail('A script of the font has too many languages');
    const systems: number[] = [];
    if (table.u16(script) !== 0) systems.push(script + table.u16(script));
    for (let k = 0; k < languages; k++) systems.push(script + table.u16(script + 4 + k * 6 + 4));
    for (let k = 0; k < systems.length; k++) {
      const lookups = lookupsOf(systems[k]);
      if (!lookups) return { state: 'unknown', variants: [], otherScripts };
      if (!sets.some((set) => set.join() === lookups.join())) sets.push(lookups);
    }
  }
  // Every lookup that is named, read once.
  const read = new Map<number, LigatureLookup>();
  let total = 0;
  for (let i = 0; i < sets.length; i++) {
    for (let k = 0; k < sets[i].length; k++) {
      const index = sets[i][k];
      if (read.has(index)) continue;
      const lookup = lookupList + table.u16(lookupList + 2 + index * 2);
      let type = table.u16(lookup);
      const subtables = table.u16(lookup + 4);
      if (subtables > MAX_GSUB_ITEMS) fail('A lookup of the font has too many subtables');
      const rules: LigatureLookup = new Map();
      for (let s = 0; s < subtables; s++) {
        let subtable = lookup + table.u16(lookup + 6 + s * 2);
        if (type === 7 || table.u16(lookup) === 7) {
          // An extension: the type it stands for, and where its subtable is.
          if (table.u16(subtable) !== 1) fail('An extension lookup of the font is of an unknown format');
          type = table.u16(subtable + 2);
          subtable += table.u32(subtable + 4);
        }
        // Anything but ligatures under a feature of ligatures: rules that are not read.
        if (type !== 4) return { state: 'unknown', variants: [], otherScripts };
        if (table.u16(subtable) !== 1) fail('A ligature lookup of the font is of an unknown format');
        const firsts = coverageGlyphs(table, subtable + table.u16(subtable + 2), face.glyphCount);
        const setCount = table.u16(subtable + 4);
        if (setCount > firsts.length) fail('A ligature lookup of the font has more sets than first glyphs');
        for (let f = 0; f < setCount; f++) {
          // The first subtable that knows a glyph is the one that is asked about it.
          if (rules.has(firsts[f])) continue;
          const set = subtable + table.u16(subtable + 6 + f * 2);
          const count = table.u16(set);
          const list: LigatureRule[] = [];
          for (let l = 0; l < count; l++) {
            if (++total > MAX_LIGATURES) fail('The font has too many ligatures');
            const ligature = set + table.u16(set + 2 + l * 2);
            const glyph = table.u16(ligature);
            const parts = table.u16(ligature + 2);
            if (parts < 1 || parts > MAX_LIGATURE_GLYPHS || glyph >= face.glyphCount) fail('A ligature of the font is not one');
            const rest: number[] = [];
            for (let c = 0; c + 1 < parts; c++) rest.push(table.u16(ligature + 4 + c * 2));
            list.push({ glyph, rest });
          }
          rules.set(firsts[f], list);
        }
      }
      read.set(index, rules);
    }
  }
  const variants = sets.map((set) => set.map((index) => read.get(index) as LigatureLookup).filter((lookup) => lookup.size > 0));
  if (!variants.some((variant) => variant.length > 0)) return { state: 'none', variants: [], otherScripts };
  return { state: 'known', variants, otherScripts };
}

const ligatures = new WeakMap<PdfFontFace, FontLigatures>();

/**
 * The ligatures of Latin text a font has (the features liga, clig and rlig
 * of the scripts DFLT and latn, in every language system), read once. A font
 * whose GSUB table cannot be read to its end, or that replaces glyphs by
 * rules of another kind (alternates by context, a feature a language
 * requires), is 'unknown': which glyphs a browser draws of it cannot be said.
 */
export function fontLigatures(face: PdfFontFace): FontLigatures {
  let known = ligatures.get(face);
  if (!known) {
    try {
      known = readLigatures(face);
    } catch (error) {
      if (!(error instanceof FontError)) throw error;
      known = { state: 'unknown', variants: [], otherScripts: true };
    }
    ligatures.set(face, known);
  }
  return known;
}

/**
 * The glyphs of a text after the ligature lookups of one language system
 * were applied, as OpenType applies them: lookup after lookup, each from the
 * left, the first ligature of a set that fits. Every glyph of the result
 * says how many glyphs of the input it stands for.
 */
export function applyLigatures(variant: readonly LigatureLookup[], glyphs: readonly number[]): { glyph: number; count: number }[] {
  let line = glyphs.map((glyph) => ({ glyph, count: 1 }));
  for (let l = 0; l < variant.length; l++) {
    const lookup = variant[l];
    const out: { glyph: number; count: number }[] = [];
    for (let i = 0; i < line.length; ) {
      const rules = lookup.get(line[i].glyph);
      let taken = 0;
      if (rules) {
        for (let r = 0; r < rules.length && taken === 0; r++) {
          const rest = rules[r].rest;
          if (i + rest.length >= line.length) continue;
          let fits = true;
          for (let c = 0; c < rest.length && fits; c++) fits = line[i + 1 + c].glyph === rest[c];
          if (!fits) continue;
          let count = 0;
          for (let c = 0; c <= rest.length; c++) count += line[i + c].count;
          out.push({ glyph: rules[r].glyph, count });
          taken = rest.length + 1;
        }
      }
      if (taken === 0) {
        out.push(line[i]);
        taken = 1;
      }
      i += taken;
    }
    line = out;
  }
  return line;
}

/** A font file, read. */
export interface PdfFontFace {
  /** The file, as it was given. */
  readonly bytes: Uint8Array;
  readonly kind: 'truetype' | 'cff';
  readonly tables: ReadonlyMap<string, FontTable>;
  readonly unitsPerEm: number;
  readonly glyphCount: number;
  /** The PostScript name of the font; "Font" when it states none that can be used. */
  readonly name: string;
  /** In units of the font. */
  readonly bbox: readonly [number, number, number, number];
  readonly ascent: number;
  readonly descent: number;
  readonly capHeight: number;
  /** In degrees, counter-clockwise from the vertical: negative for a font that leans to the right. */
  readonly italicAngle: number;
  readonly fixedPitch: boolean;
  readonly serif: boolean;
  /** usWeightClass of the OS/2 table. */
  readonly weight: number;
  /** fsType of the OS/2 table: what the licence of the font allows; undefined for a font without that table. */
  readonly fsType: number | undefined;
  /** 0 or 1 for a TrueType font: how its loca table is stored. */
  readonly longLoca: boolean;
  /** The glyph of a character; 0 when the font has none. */
  glyph(codePoint: number): number;
  /** The advance of a glyph in units of the font. */
  advance(glyph: number): number;
}

/**
 * Reads a font file. Throws a FontError for what is no TrueType or OpenType
 * font (a collection, a WOFF or WOFF2 file, a font of bitmaps), for a font
 * with variations or with glyphs in colour, for a file beyond MAX_FONT_BYTES
 * and for one whose tables are not as they must be.
 */
export function readFont(bytes: Uint8Array): PdfFontFace {
  if (!(bytes instanceof Uint8Array)) return fail('A font file must be bytes');
  if (bytes.length > MAX_FONT_BYTES) return fail(`A font file of more than ${MAX_FONT_BYTES} bytes is not read`);
  if (bytes.length < 12) return fail('The font file is shorter than its header');
  const file = new FontReader(bytes, 0, bytes.length, 'directory');
  const version = file.u32(0);
  if (version === 0x774f4646) return fail('The font is a WOFF file: only TrueType and OpenType files can be put into a PDF');
  if (version === 0x774f4632) return fail('The font is a WOFF2 file: only TrueType and OpenType files can be put into a PDF');
  if (version === 0x74746366) return fail('The font file is a collection of fonts, which is not read here');
  if (version !== 0x00010000 && version !== 0x4f54544f && version !== 0x74727565) return fail('The file is no TrueType or OpenType font');
  const count = file.u16(4);
  if (count < 1 || count > MAX_TABLES || 12 + count * 16 > bytes.length) return fail('The table directory of the font does not fit');
  const tables = new Map<string, FontTable>();
  for (let i = 0; i < count; i++) {
    const at = 12 + i * 16;
    const tag = String.fromCharCode(bytes[at], bytes[at + 1], bytes[at + 2], bytes[at + 3]);
    const offset = file.u32(at + 8);
    const length = file.u32(at + 12);
    if (offset + length > bytes.length) return fail(`The ${tag} table lies beyond the end of the font file`);
    // The first of two tables with one tag is the one that counts.
    if (!tables.has(tag)) tables.set(tag, { offset, length });
  }
  // What a browser draws of such a font is not what its plain outlines say: an instance of its variations, or
  // glyphs in colours or from pictures.
  if (tables.has('fvar')) return fail('The font has variations, which are not read here');
  if (tables.has('COLR') || tables.has('SVG ') || tables.has('sbix') || tables.has('CBDT')) return fail('The font has glyphs in colour, which are not read here');
  const reader = (tag: string): FontReader | undefined => {
    const table = tables.get(tag);
    return table ? new FontReader(bytes, table.offset, table.length, tag) : undefined;
  };
  const need = (tag: string): FontReader => reader(tag) || fail(`The font has no ${tag} table`);

  const head = need('head');
  if (head.length < 54 || head.u32(12) !== 0x5f0f3cf5) return fail('The head table of the font is not one');
  const unitsPerEm = head.u16(18);
  if (unitsPerEm < 16 || unitsPerEm > 16384) return fail('The font has no units per em that can be');
  const bbox: [number, number, number, number] = [head.i16(36), head.i16(38), head.i16(40), head.i16(42)];
  const locaFormat = head.i16(50);

  const maxp = need('maxp');
  if (maxp.length < 6) return fail('The maxp table of the font is cut short');
  const glyphCount = maxp.u16(4);
  if (glyphCount < 1) return fail('The font has no glyphs');

  const hhea = need('hhea');
  if (hhea.length < 36) return fail('The hhea table of the font is cut short');
  const metrics = hhea.u16(34);
  if (metrics < 1 || metrics > glyphCount) return fail('The font states a number of advances that does not go with its glyphs');
  const hmtx = need('hmtx');
  if (hmtx.length < metrics * 4) return fail('The hmtx table of the font is shorter than its advances');

  const map = characterMap(need('cmap'));

  let kind: 'truetype' | 'cff';
  if (tables.has('glyf')) {
    kind = 'truetype';
    if (locaFormat !== 0 && locaFormat !== 1) return fail('The font does not say how its glyphs are indexed');
    const loca = need('loca');
    if (loca.length < (glyphCount + 1) * (locaFormat === 1 ? 4 : 2)) return fail('The loca table of the font is too short for its glyphs');
    // Where the glyphs are: one after the other, inside the glyf table. A font that says otherwise (glyphs that
    // overlap, or lie outside) is none whose glyphs can be told apart.
    const glyfLength = need('glyf').length;
    let before = 0;
    for (let glyph = 0; glyph <= glyphCount; glyph++) {
      const at = locaFormat === 1 ? loca.u32(glyph * 4) : loca.u16(glyph * 2) * 2;
      if (at < before || at > glyfLength) return fail(`The loca table of the font puts glyph ${glyph} outside its glyphs or before the glyph in front of it`);
      before = at;
    }
    if (maxp.length < 32 || maxp.u32(0) !== 0x00010000) return fail('The maxp table of the font is not that of a TrueType font');
  } else if (tables.has('CFF ')) {
    kind = 'cff';
    checkCff(need('CFF '), glyphCount);
  } else return fail('The font has neither TrueType nor CFF outlines');

  const os2 = reader('OS/2');
  let fsType: number | undefined;
  let weight = 400;
  let capHeight = 0;
  let serif = false;
  if (os2 && os2.length >= 10) {
    weight = Math.min(1000, Math.max(1, os2.u16(4)));
    fsType = os2.u16(8);
    if (os2.length >= 32) {
      // sFamilyClass: classes 1 to 5 and 7 are the ones with serifs.
      const family = os2.i16(30) >> 8;
      serif = (family >= 1 && family <= 5) || family === 7;
    }
    if (os2.length >= 90 && os2.u16(0) >= 2) capHeight = os2.i16(88);
  }
  const ascent = hhea.i16(4);
  const descent = hhea.i16(6);
  // A font whose box is none gets the box of its line: a descriptor of a PDF needs one.
  if (!(bbox[2] > bbox[0]) || !(bbox[3] > bbox[1])) {
    bbox[0] = 0;
    bbox[1] = Math.min(descent, 0);
    bbox[2] = unitsPerEm;
    bbox[3] = Math.max(ascent, bbox[1] + unitsPerEm);
  }
  if (!(capHeight > 0)) capHeight = Math.round(Math.max(ascent, 0) * 0.7);
  const post = reader('post');
  let italicAngle = 0;
  let fixedPitch = false;
  if (post && post.length >= 16) {
    italicAngle = post.i32(4) / 65536;
    if (!(italicAngle > -90 && italicAngle < 90)) italicAngle = 0;
    fixedPitch = post.u32(12) !== 0;
  }
  const lastAdvance = hmtx.u16((metrics - 1) * 4);
  return {
    bytes,
    kind,
    tables,
    unitsPerEm,
    glyphCount,
    name: postScriptName(reader('name')) || 'Font',
    bbox,
    ascent,
    descent,
    capHeight,
    italicAngle,
    fixedPitch,
    serif,
    weight,
    fsType,
    longLoca: locaFormat === 1,
    glyph: (codePoint) => {
      if (!(codePoint >= 0) || codePoint > 0x10ffff) return 0;
      const glyph = map.glyph(codePoint);
      return glyph > 0 && glyph < glyphCount ? glyph : 0;
    },
    advance: (glyph) => (glyph >= 0 && glyph < metrics ? hmtx.u16(glyph * 4) : lastAdvance),
  };
}

/** A table of a font, to be read. Throws a FontError when the font has none of that tag. */
export function fontTable(face: PdfFontFace, tag: string): FontReader {
  const table = face.tables.get(tag);
  if (!table) return fail(`The font has no ${tag} table`);
  return new FontReader(face.bytes, table.offset, table.length, tag);
}

export interface FontLicence {
  /** Whether the font may be put into a file at all. */
  embed: boolean;
  /** Whether a part of it may be (false: the whole font or none). */
  subset: boolean;
  /** Why not, when it may not. */
  reason?: string;
}

/**
 * What the licence bits of a font (OS/2.fsType) allow. "Restricted licence
 * embedding" (bit 1) and "bitmap embedding only" (bit 9) keep the font out of
 * a file, whatever the other bits say; "no subsetting" (bit 8) lets it in
 * only as a whole. A font that does not state its licence is not embedded.
 */
export function fontLicence(face: Pick<PdfFontFace, 'fsType'>): FontLicence {
  const type = face.fsType;
  if (type === undefined) return { embed: false, subset: false, reason: 'the font does not say whether it may be embedded' };
  if ((type & 0x0002) !== 0) return { embed: false, subset: false, reason: 'the licence of the font does not allow embedding' };
  if ((type & 0x0200) !== 0) return { embed: false, subset: false, reason: 'the licence of the font allows only its bitmaps to be embedded' };
  return { embed: true, subset: (type & 0x0100) === 0 };
}
